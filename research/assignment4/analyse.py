"""Build the Assignment 4 evidence from archived observations, without API calls."""
from __future__ import annotations

import csv
import hashlib
import importlib.metadata
import json
import platform
import shutil
import statistics as st
import sys
import tempfile
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from analysis.export_public_data import export
from analysis.google_followup_summary import summarise

OUT = ROOT / "assignment4"
DATA = OUT / "data"
TABLES = OUT / "tables"
FIGURES = OUT / "figures"


def read(path):
    with path.open(encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def write(name, rows):
    with (TABLES / name).open("w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=list(rows[0]), lineterminator="\n")
        writer.writeheader()
        writer.writerows(rows)


def sha(path):
    payload = path.read_bytes()
    if path.suffix in {".csv", ".json", ".md", ".py", ".txt"}:
        payload = payload.replace(b"\r\n", b"\n")
    return hashlib.sha256(payload).hexdigest()


def build():
    for folder in (DATA, TABLES, FIGURES):
        folder.mkdir(parents=True, exist_ok=True)
    temp_root = (ROOT.parent.parent if ROOT.name == "research" else ROOT.parent) / "tmp"
    temp_root.mkdir(exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="rmt-a4-", dir=temp_root) as temp:
        exported = export(ROOT, Path(temp))
        for p in Path(temp).iterdir():
            shutil.copyfile(p, DATA / p.name)
    rows = read(DATA / "ors_processed.csv")
    raw = read(DATA / "ors_observations.csv")
    batches = read(DATA / "google_ors_summary.csv")
    followup = [r for r in batches if r["phase"] == "followup"]
    protocol = json.loads((ROOT / "docs/google_followup_protocol.json").read_text())
    period = summarise(followup, protocol)
    (TABLES / "period_summary.json").write_text(json.dumps(period, indent=2), encoding="utf-8")
    write("period_summary.csv", period["groups"])
    write("batch_results.csv", batches)

    coverage = Counter((r["date_astana"], r["phase"], r["period"], r["usable"]) for r in rows)
    write("coverage.csv", [dict(date_astana=k[0], phase=k[1], period=k[2], usable=k[3], n=v)
                           for k, v in sorted(coverage.items())])
    routes = {r["route_id"]: r for r in read(DATA / "routes.csv")}
    per_route = []
    for rid in sorted(routes):
        group = [r for r in rows if r["route_id"] == rid]
        per_route.append(dict(route_id=rid, name=routes[rid]["name"], n=len(group),
                              min_duration_s=min(float(r["duration_s"]) for r in group),
                              median_duration_s=st.median(float(r["duration_s"]) for r in group),
                              max_duration_s=max(float(r["duration_s"]) for r in group),
                              median_distance_m=st.median(float(r["distance_m"]) for r in group),
                              distinct_duration_distance_pairs=len({(r["duration_s"], r["distance_m"]) for r in group})))
    write("route_summary.csv", per_route)

    ratio_fields = {"duration_s": "seconds", "duration_no_traffic_s": "seconds", "distance_m": "metres",
                    "latency_s": "seconds", "speed_kmh": "km/h"}
    missingness = []
    for field in rows[0]:
        values = [r[field] for r in rows if r[field] != ""]
        numeric = [float(v) for v in values] if field in ratio_fields else []
        level = "ratio" if field in ratio_fields else "interval" if field.startswith("timestamp") or field == "date_astana" else "nominal"
        missingness.append(dict(variable=field, level=level, unit=ratio_fields.get(field, ""),
                                valid=len(values), missing=len(rows)-len(values),
                                missing_pct=100*(len(rows)-len(values))/len(rows),
                                minimum=min(numeric) if numeric else "",
                                median=st.median(numeric) if numeric else "",
                                maximum=max(numeric) if numeric else "",
                                categories="; ".join(sorted(set(values))) if level == "nominal" and len(set(values)) <= 8 else ""))
    write("variables.csv", missingness)

    required = "timestamp_utc timestamp_astana cycle_id route_id provider duration_s distance_m status http_code request_id latency_s phase".split()
    main = [r for r in rows if r["phase"] in ("pilot", "scheduled")]
    scheduled = [r for r in rows if r["phase"] == "scheduled"]
    cloud_runs = json.loads((ROOT / "docs/cloud_runs.json").read_text())
    manual_cycles = {"gh-" + str(r["databaseId"]) for r in cloud_runs["runs"] if r["event"] == "workflow_dispatch"}
    manual_verification = sum(r["cycle_id"] in manual_cycles for r in scheduled)
    inventory = dict(exported)
    inventory.update(main_ors=len(main), pilot_ors=sum(r["phase"] == "pilot" for r in rows),
                     scheduled_ors=len(scheduled), exploratory_ors=sum(r["phase"] == "exploratory" for r in rows),
                     manual_verification_ors=manual_verification, automatic_ors=len(scheduled)-manual_verification,
                     automatic_pct_original_plan=100*(len(scheduled)-manual_verification)/768,
                     followup_ors=sum(r["phase"] == "followup" for r in rows),
                     planned_followup_pairs=80, actual_followup_pairs=sum(int(r["n_valid_pairs"]) for r in followup),
                     planned_scheduled_ors=768, scheduled_pct_original_plan=100*len(scheduled)/768,
                     pilot_routes=len({r["route_id"] for r in rows if r["phase"] == "pilot"}),
                     failed_ors=sum(r["usable"] != "true" for r in rows),
                     main_latency_le_10=sum(float(r["latency_s"]) <= 10 for r in main),
                     all_latency_le_10=sum(float(r["latency_s"]) <= 10 for r in rows),
                     raw_fields=len(raw[0]), raw_cells=len(raw)*len(raw[0]),
                     raw_blank_cells=sum(v == "" for r in raw for v in r.values()),
                     required_cells=len(rows)*len(required),
                     required_blank_cells=sum(r[k] == "" for r in rows for k in required),
                     main_periods=dict(Counter(r["period"] for r in main)),
                     all_periods=dict(Counter(r["period"] for r in rows)),
                     source_start_utc=min(r["timestamp_utc"] for r in rows),
                     source_end_utc=max(r["timestamp_utc"] for r in rows),
                     weekdays={date:datetime.fromisoformat(date).strftime("%A") for date in sorted({r["date_astana"] for r in rows})},
                     unchanged_ors_routes=sum(r["distinct_duration_distance_pairs"] == 1 for r in per_route))
    inventory["raw_blank_pct"] = 100*inventory["raw_blank_cells"]/inventory["raw_cells"]
    (TABLES / "inventory.json").write_text(json.dumps(inventory, indent=2), encoding="utf-8")

    cleaning = [
        dict(operation="Import main series", rule="data/raw/*.csv; retain pilot and scheduled labels", input_rows=len(main), affected=0, output_rows=len(main)),
        dict(operation="Add exploratory ORS", rule="separate exploratory cycle; never count as follow-up", input_rows=inventory["exploratory_ors"], affected=0, output_rows=len(main)+inventory["exploratory_ors"]),
        dict(operation="Add follow-up ORS", rule="four registered comparison windows", input_rows=inventory["followup_ors"], affected=0, output_rows=len(rows)),
        dict(operation="Screen measurements", rule="HTTP 200 and ok; finite positive time and distance; consistent timestamps; speed 2-200 km/h", input_rows=len(rows), affected=inventory["failed_ors"], output_rows=sum(r["usable"]=="true" for r in rows)),
        dict(operation="Screen duplicate identities", rule="exact row, phase/cycle/route/provider conflict or reused request ID", input_rows=len(rows), affected=exported["ors_duplicates"], output_rows=sum(r["usable"]=="true" for r in rows)),
        dict(operation="Separate outside-period main rows", rule="other is excluded only from peak/off-peak comparison", input_rows=len(main), affected=inventory["main_periods"]["other"], output_rows=len(main)-inventory["main_periods"]["other"]),
        dict(operation="Select paired-window summaries", rule="phase=followup; registered cycle; at least five valid pairs", input_rows=len(batches), affected=len(batches)-len(followup), output_rows=len(followup)),
    ]
    write("cleaning.csv", cleaning)
    sensitivity = []
    for label, low, high, phases in [("Recorded rule",2,200,None), ("Allow speeds below 2",0,200,None),
                                     ("Upper speed limit 160",2,160,None), ("Exclude pilot",2,200,{"scheduled","followup","exploratory"})]:
        selected = [r for r in rows if low <= float(r["speed_kmh"]) <= high and (phases is None or r["phase"] in phases)]
        sensitivity.append(dict(rule=label, ors_kept=len(selected), ors_excluded=len(rows)-len(selected),
                                valid_followup_pairs=period["n_valid_pairs"],
                                peak_minus_offpeak_s=period["peak_minus_offpeak_window_median_s"],
                                peak_offpeak_ratio=period["peak_to_offpeak_window_median_ratio"]))
    write("sensitivity.csv", sensitivity)
    # Window deletion checks assess dependence on the few sampled windows, not cleaning.
    window_checks = []
    for omitted in followup:
        s = summarise([r for r in followup if r is not omitted], protocol)
        window_checks.append(dict(omitted_window=omitted["cycle_id"], windows_remaining=s["valid_windows"],
                                  peak_minus_offpeak_s=s["peak_minus_offpeak_window_median_s"],
                                  peak_offpeak_ratio=s["peak_to_offpeak_window_median_ratio"]))
    write("window_sensitivity.csv", window_checks)

    plt.rcParams.update({"font.family":"DejaVu Sans", "font.size":10, "axes.spines.top":False,
                         "axes.spines.right":False, "axes.titleweight":"bold", "figure.dpi":160,
                         "savefig.dpi":200, "axes.axisbelow":True})
    manifest = []
    def save(fig, name, sources, description):
        fig.tight_layout(pad=1.2)
        fig.savefig(FIGURES/name, facecolor="white")
        plt.close(fig)
        manifest.append(dict(file="assignment4/figures/"+name, sources=sources, description=description, sha256=sha(FIGURES/name)))
    labels = ["24 Oct\n17:42 peak", "24 Oct\n22:00 off-peak", "25 Oct\n10:00 off-peak", "25 Oct\n17:30 peak"]
    colors = ["#B3533A" if r["period"] == "peak" else "#267083" for r in followup]
    fig, ax = plt.subplots(figsize=(8.4,4))
    vals = [float(r["median_absolute_difference_s"])/60 for r in followup]
    bars=ax.bar(labels, vals, color=colors, width=.6)
    ax.bar_label(bars, labels=[f"{v:.2f}" for v in vals], padding=4)
    ax.set(ylabel="Median absolute difference (minutes)", title="Google and ORS disagreement in each recorded window", ylim=(0,15))
    ax.text(.99,.96,"20 valid pairs in every window",ha="right",va="top",transform=ax.transAxes,fontsize=9)
    save(fig,"01_window_disagreement.png",["assignment4/tables/batch_results.csv"],"Four recorded window medians; no pooled route-level median or uncertainty interval.")

    fig,ax=plt.subplots(figsize=(8.4,3.7))
    groups=period["groups"]; vals=[g["median_of_window_median_absolute_differences_s"]/60 for g in groups]
    bars=ax.barh(["Peak\n2 windows, 40 pairs","Off-peak\n2 windows, 40 pairs"], vals, color=["#B3533A","#267083"],height=.52)
    ax.bar_label(bars,labels=[f"{v:.2f} min" for v in vals],padding=5)
    ax.set(xlabel="Median of window medians (minutes)",xlim=(0,14),title="Recorded peak and off-peak summaries")
    ax.invert_yaxis()
    save(fig,"02_period_comparison.png",["assignment4/tables/period_summary.csv"],"Each period has two equally weighted windows; descriptive comparison only.")

    fig,ax=plt.subplots(figsize=(8.4,4))
    phases=["pilot","scheduled","exploratory","followup"]; names=["Pilot","Scheduled","Exploratory","Follow-up"]
    bottom=np.zeros(4)
    for period_name,color in [("peak","#B3533A"),("offpeak","#267083"),("other","#B7B7B7")]:
        counts=[sum(r["phase"]==phase and r["period"]==period_name for r in rows) for phase in phases]
        bars=ax.bar(names,counts,bottom=bottom,color=color,label=period_name.replace("offpeak","off-peak"))
        for j,n in enumerate(counts):
            if n:ax.text(j,bottom[j]+n/2,str(n),ha="center",va="center",fontsize=9)
        bottom+=counts
    ax.legend(frameon=False,ncol=3,loc="upper left")
    ax.set(ylabel="Archived ORS observations",title="Coverage by collection phase and time category",ylim=(0,470))
    save(fig,"03_coverage.png",["assignment4/tables/coverage.csv"],"Every ORS row is counted once; phases are kept separate.")

    fig,ax=plt.subplots(figsize=(8.4,4))
    x=np.arange(len(followup)); width=.23
    for j,(field,label,color) in enumerate([("median_absolute_difference_s","Median absolute gap","#267083"),("mae_s","Mean absolute gap","#707070"),("rmse_s","Root mean square gap","#B3533A")]):
        ax.bar(x+(j-1)*width,[float(r[field])/60 for r in followup],width,label=label,color=color)
    ax.set_xticks(x,labels);ax.set(ylabel="Provider disagreement (minutes)",title="Typical gaps and larger gaps across windows",ylim=(0,17))
    ax.legend(frameon=False,ncol=3,fontsize=8,loc="upper left")
    save(fig,"04_disagreement_metrics.png",["assignment4/tables/batch_results.csv"],"Three distinct summaries of the same paired differences; no claim about observed journey accuracy.")

    fig,ax=plt.subplots(figsize=(8.4,3.8))
    vals=[float(r["length_comparable_pct"]) for r in followup]
    bars=ax.bar(labels,vals,color=colors,width=.6)
    ax.bar_label(bars,labels=[f"{r['n_length_comparable']}/20 ({v:.0f}%)" for r,v in zip(followup,vals)],padding=4)
    ax.axhline(70,color="#555555",linestyle="--",linewidth=1)
    ax.set(ylabel="Pairs with similar route lengths (%)",title="Route lengths within five percent of the pair mean",ylim=(0,105))
    ax.text(.99,.95,"Dashed line: original 70% reference",ha="right",va="top",transform=ax.transAxes,fontsize=8)
    save(fig,"05_length_comparability.png",["assignment4/tables/batch_results.csv"],"The original reference concerns TomTom-2GIS and is not retested here; equal length does not establish equal path.")

    fig,ax=plt.subplots(figsize=(8.4,4))
    ax.scatter([r["route_id"] for r in per_route],[r["median_duration_s"]/60 for r in per_route],color="#267083",s=33)
    ax.set(ylabel="ORS duration (minutes)",xlabel="Directed route",title="ORS estimates stayed unchanged within every route")
    ax.tick_params(axis="x",labelsize=8)
    save(fig,"06_ors_routes.png",["assignment4/tables/route_summary.csv"],"All recorded durations and distances are identical within each route; each point is one route, not an independent request.")
    sources = sorted(set(list((ROOT/"data/raw").glob("*.csv"))+[
        ROOT/"data/exploratory/ors/observations.csv",ROOT/"data/exploratory/google_request_status.csv",
        ROOT/"data/followup/ors/observations.csv",ROOT/"data/followup/google_request_status.csv",
        ROOT/"results/google_control_benchmark.csv",ROOT/"results/google_followup_benchmark.csv",
        ROOT/"data/routes.csv",ROOT/"docs/cloud_runs.json",ROOT/"docs/expected_results.md",ROOT/"docs/expected_results.json",
        ROOT/"docs/google_followup_protocol.json",ROOT/"docs/google_control_protocol.json",
        ROOT/"data/sources/astana_overpass_query.txt",ROOT/"data/sources/astana_osm_manifest.json"] ))
    versions = {p:importlib.metadata.version(p) for p in ("matplotlib","numpy","tzdata")}
    evidence = dict(generated_at_utc=datetime.now(timezone.utc).isoformat(), source_snapshot_utc=exported["snapshot_utc"],
                    hash_rule="SHA-256 after CRLF-to-LF normalization for CSV, JSON, Markdown, Python and plain text; exact bytes for binary figures.",
                    python=platform.python_version(),packages=versions,
                    inputs={p.relative_to(ROOT).as_posix():sha(p) for p in sources},
                    scripts={"assignment4/analyse.py":sha(Path(__file__)),"analysis/export_public_data.py":sha(ROOT/"analysis/export_public_data.py")},
                    outputs={p.relative_to(ROOT).as_posix():sha(p) for folder in (DATA,TABLES) for p in sorted(folder.iterdir())},
                    figures=manifest)
    (OUT/"evidence_manifest.json").write_text(json.dumps(evidence,indent=2),encoding="utf-8")
    print(json.dumps(dict(ors=len(rows),followup_pairs=period["n_valid_pairs"],figures=len(manifest),period_contrast_s=period["peak_minus_offpeak_window_median_s"])))


if __name__ == "__main__":
    build()
