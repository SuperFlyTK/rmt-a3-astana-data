"""Build the team review draft from the published Assignment 4 evidence.

Run with Python and python-pptx, for example: python build.py
The editable deck and rehearsal notes are written to output/.
"""

import csv
import json
import re
from datetime import datetime, timedelta, timezone
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt


HERE = Path(__file__).resolve().parent
RESEARCH = HERE.parent
TABLES = RESEARCH / "assignment4" / "tables"
SOURCE = (
    "https://github.com/SuperFlyTK/rmt-a3-astana-data/blob/"
    "d1d8df1c89aaa5146839749586e073ab8d657043/research/"
)
DOI = "https://doi.org/"
PAPER, INK, NAVY = "F0F0EE", "222840", "1C2D6C"
MUTED, RULE, SOFT = "454F6A", "747E99", "C2CBE3"
OFFPEAK = "7282AF"
W, H = 13.333333, 7.5


def load_json(path):
    return json.loads(path.read_text(encoding="utf-8"))


def load_csv(path):
    with path.open(encoding="utf-8", newline="") as source:
        return list(csv.DictReader(source))


inventory = load_json(TABLES / "inventory.json")
metrics = load_json(TABLES / "reported_metrics.json")["rounded_values"]
protocol = load_json(RESEARCH / "docs" / "google_followup_protocol.json")
windows = [
    row for row in load_csv(TABLES / "batch_results.csv")
    if row["phase"] == "followup"
]
sensitivity = load_csv(TABLES / "window_sensitivity.csv")
peak_gap = metrics["peak_s"]
offpeak_gap = metrics["offpeak_s"]
contrast = metrics["contrast_s"]
ratio = metrics["ratio"]
range_s = [float(row["peak_minus_offpeak_s"]) for row in sensitivity]
low, high = f"{min(range_s):.2f}", f"{max(range_s):.2f}"
observations = inventory["ors_observations"]
usable = inventory["ors_usable"]
valid_pairs = sum(int(row["n_valid_pairs"]) for row in windows)
automatic = inventory["automatic_ors"]
planned_automatic = inventory["planned_scheduled_ors"]
route_count = inventory["routes"]
route_max = next(
    row["maximum"] for row in load_csv(TABLES / "variables.csv")
    if row["variable"] == "latency_s"
)
assert valid_pairs == inventory["actual_followup_pairs"]
assert len(windows) == len(protocol["windows"]) == 4
assert all(int(row["n_valid_pairs"]) == route_count for row in windows)
assert observations == usable == inventory["all_latency_le_10"]

deck = Presentation()
deck.slide_width, deck.slide_height = Inches(W), Inches(H)
deck.core_properties.title = "Travel-time estimates in Astana"
deck.core_properties.subject = "RMT final defence • SE-2424 • team review draft"
deck.core_properties.author = "SE-2424 research team"
audit = []
notes = []


def color(hex_code):
    return RGBColor.from_string(hex_code)


def shape(slide, kind, x, y, w, h, *, fill=None, stroke=None):
    item = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    style = item.element.find(qn("p:style"))
    item.element.remove(style)
    if fill:
        item.fill.solid()
        item.fill.fore_color.rgb = color(fill)
    else:
        item.fill.background()
    if stroke:
        item.line.color.rgb = color(stroke)
        item.line.width = Pt(0.85)
    else:
        item.line.fill.background()
    return item


def rule(slide, x, y, w, *, fill=RULE):
    shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, x, y, w, 0.012, fill=fill)


def tx(slide, value, x, y, w, h, size=21, *, face="Arial", bold=False,
       ink=INK, align=PP_ALIGN.LEFT, link=None):
    assert x >= 0 and y >= 0 and x + w <= W + 0.001 and y + h <= H + 0.001
    assert size >= 18, (value, size)
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    if link:
        box.click_action.hyperlink.address = link
    frame = box.text_frame
    frame.clear()
    frame.word_wrap = True
    frame.margin_left = frame.margin_right = frame.margin_top = frame.margin_bottom = 0
    frame.vertical_anchor = MSO_ANCHOR.TOP
    for i, line in enumerate(value.split("\n")):
        p = frame.paragraphs[0] if i == 0 else frame.add_paragraph()
        p.alignment = align
        p.line_spacing = 1.05
        p.space_after = Pt(0)
        run = p.add_run()
        run.text = line
        run.font.name = face
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = color(ink)
    audit[-1].append(value)


def page(title=None, *, dark=False):
    slide = deck.slides.add_slide(deck.slide_layouts[6])
    audit.append([])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = color(NAVY if dark else PAPER)
    border = shape(
        slide, MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE,
        0.28, 0.26, W - 0.56, H - 0.52, stroke=SOFT if dark else RULE,
    )
    border.adjustments[0] = 0.06
    shape(
        slide, MSO_AUTO_SHAPE_TYPE.OVAL, W / 2 - 0.065, 0.195, 0.13, 0.13,
        fill=SOFT if dark else NAVY,
    )
    if title:
        tx(slide, title, 0.72, 0.62, 11.85, 0.83, 38, face="Georgia", bold=True,
           ink=PAPER if dark else NAVY)
    return slide


def footer(slide, label, source, *, dark=False):
    ink = PAPER if dark else MUTED
    rule(slide, 0.73, 6.54, 11.85, fill=SOFT if dark else RULE)
    tx(slide, label, 0.73, 6.68, 11.85, 0.39, 18, ink=ink, link=SOURCE + source)


def cue(slide, speaker, exam, rehearsal, core, extra="", sources=()):
    label = next((line for line in audit[-1]), "Slide")
    notes.append((len(audit), label, speaker, exam, rehearsal, core, extra, sources))
    source_text = "\n".join(sources)
    slide.notes_slide.notes_text_frame.text = (
        f"{speaker} · Exam {exam}s · Three-minute-per-person rehearsal {rehearsal}s\n\n"
        f"Exam cues\n{core}\n\n"
        + (f"Rehearsal extension\n{extra}\n\n" if extra else "")
        + f"Sources\n{source_text}"
    )


def row(slide, y, label, detail, *, label_w=3.25, size=21):
    rule(slide, 0.73, y - 0.12, 11.85)
    tx(slide, label, 0.75, y, label_w, 0.58, size, bold=True, ink=NAVY)
    tx(slide, detail, 0.85 + label_w, y, 11.65 - label_w, 0.68, size)


s = page()
tx(s, "Travel-time\nestimates\nin Astana", 0.75, 0.96, 7.8, 3.95, 78,
   face="Georgia", bold=True, ink=NAVY)
rule(s, 8.98, 2.47, 3.36)
tx(s, "Google Routes and\nopenrouteservice", 8.98, 2.67, 3.47, 1.04, 25, bold=True, ink=NAVY)
tx(s, "Final defence\nSE-2424", 8.98, 4.03, 3.36, 0.82, 21, ink=MUTED)
rule(s, 0.75, 5.57, 11.83)
tx(s, "Yerdaulet Akhanov     Syrym Shadiyarbek     Akylbek Sabyrzhan",
   0.75, 5.88, 11.83, 0.46, 22, bold=True)
cue(s, "Yerdaulet", 20, 20,
    "Introduce all three teammates and SE-2424. We wanted to see how much two navigation estimates can differ for the same trip in Astana.",
    sources=("Assignment 4 title page and abstract",))

s = page("Two services, one trip")
tx(s, "The same endpoints can produce different driving estimates.", 0.75, 1.65, 11.4, 0.6, 24)
rule(s, 0.75, 2.57, 11.82)
tx(s, protocol["current_rq"], 0.75, 2.93, 11.13, 2.4, 37,
   face="Georgia", bold=True, ink=NAVY)
footer(s, "Research question · A3 amendment and A4 §1.4", "docs/google_followup_protocol.json")
cue(s, "Yerdaulet", 40, 47,
    "Give an example of planning a drive between the same endpoints with two estimates. State the question as written. Agreement is how close the estimates are. Accuracy requires an independently timed journey. Access and publication permissions led to the final Google–ORS comparison.",
    "Point to the historical questions in the backup if asked. A1 named Yandex and A2 named 2GIS.",
    ("Assignment 3, dated 24 October amendment", "Assignment 4 §1.1 and §1.4"))

s = page("Aim and original objectives")
tx(s, "Describe the size, direction and spread of Google–ORS gaps.", 0.75, 1.55, 11.75, 0.53, 24, bold=True)
for y, name, detail in [
    (2.4, "O1", "Select 20 representative routes and fix periods."),
    (3.17, "O2", "Pair Google and Yandex for 14 days."),
    (3.94, "O3", "Calculate gaps, spread and confidence intervals."),
    (4.71, "O4", "Test period differences and estimate effect size."),
    (5.48, "O5", "Check route, time and period patterns."),
]:
    row(s, y, name, detail, label_w=1.05)
footer(s, "A1 objectives · A4 Table 5", "assignment4/tables/objectives.csv")
cue(s, "Yerdaulet", 30, 36,
    "The final aim is descriptive. The five objectives came from A1, including the original provider pair, the 14-day target and the plan for statistical tests. We report what happened to each objective after the design changed.",
    "Magnitude asks how many seconds apart. Direction asks which provider was higher. Spread asks how much the gaps vary.",
    ("Assignment 1 T3", "Assignment 4 §1.4–1.5 and Table 5"))

s = page("What the studies disagreed on")
tx(s, "Agreement across references", 0.75, 1.69, 5.72, 0.85, 25,
   face="Georgia", bold=True, ink=NAVY)
tx(s, "Elnaggar et al. (2025)", 0.75, 2.83, 5.6, 0.4, 21, bold=True,
   link=DOI + "10.3390/su17188194")
tx(s, "Under 5% against Google navigation.", 0.75, 3.28, 5.6, 0.65)
tx(s, "Bastarianto et al. (2025)", 0.75, 4.16, 5.6, 0.4, 21, bold=True,
   link=DOI + "10.1016/j.cstp.2025.101508")
tx(s, "Google estimates diverged from reported journeys.", 0.75, 4.61, 5.6, 0.92)
tx(s, "Google as a benchmark", 6.94, 1.69, 5.62, 0.85, 25,
   face="Georgia", bold=True, ink=NAVY)
tx(s, "Boeing and Zhou (2026)", 6.94, 2.83, 5.6, 0.4, 21, bold=True,
   link=DOI + "10.1080/13658816.2026.2628193")
tx(s, "Evaluated their model against Google.", 6.94, 3.28, 5.6, 0.92)
tx(s, "Valero et al. (2025)", 6.94, 4.16, 5.6, 0.4, 21, bold=True,
   link=DOI + "10.32866/001c.143380")
tx(s, "Google commute gaps varied by travel mode.", 6.94, 4.61, 5.6, 0.92)
tx(s, "Astana gap · Fixed-route agreement by period", 0.75, 6.02, 11.8, 0.41, 20, bold=True)
tx(s, "A2 §6.3", 0.75, 6.68, 11.8, 0.4, 18, ink=MUTED)
cue(s, "Yerdaulet", 40, 48,
    "A2 found two disagreements. Elnaggar reported less than five percent discrepancy using Google navigation as the reference. Bastarianto found consequential mismatches against reported journeys. Boeing and Zhou used Google as a benchmark for their open-data model. Valero found that Google and reported times differed by commute mode. Our gap was the local comparison across fixed periods.",
    "Those papers measure different outcomes. Reports can involve rounding and recall, and agreement with a reference model does not validate travel on the road.",
    ("Assignment 2 §6.3 and §7.2", "The four article DOIs are linked on the slide."))

s = page("How we collected the estimates")
for y, title, detail in [
    (1.83, "Endpoints", "Public OpenStreetMap nodes on 20 directed routes."),
    (2.76, "Collector", "Python on GitHub Actions called Google Routes and ORS."),
    (3.69, "Pairing", "Same endpoints, request starts within 60 seconds."),
    (4.62, "Screening", "Check status, time, speed and duplicates. Retain Google aggregates."),
]:
    row(s, y, title, detail, label_w=3.62)
tx(s, "Traffic-aware Google. ORS control without live traffic.",
   0.75, 5.63, 11.8, 0.40, 21, bold=True)
tx(s, "Astana time   Peak 17:30–19:30   Off-peak 10:00–12:00, 22:00–24:00",
   0.75, 6.09, 11.8, 0.38, 18)
footer(s, "A4 §2 · Protocol", "docs/google_followup_protocol.json")
cue(s, "Akylbek", 40, 56,
    "Explain the instrument: the fixed route CSV, Python collector and dated protocol. We chose public road nodes on selected corridors, tested three pilot cycles and scheduled ORS requests. For the four follow-up windows we queried Google traffic-aware optimal and ORS for each route. The starts had to be within 60 seconds. We logged actual times and status, then checked speed, timestamps and duplicates.",
    "One prospective batch ran per window. Late or missing observations stayed out. Individual Google durations, distances and geometry were processed in memory. We kept Google request metadata and window aggregates alongside permitted ORS observations.",
    ("Assignment 4 §2.1–2.5 and §3.3", SOURCE + "docs/google_followup_protocol.json"))

s = page("What entered the analysis")
tx(s, f"{usable}/{observations}", 0.76, 1.65, 5.3, 1.05, 57,
   face="Georgia", bold=True, ink=NAVY)
tx(s, "ORS records analyzed/collected\nOne record is one route request", 0.79, 2.77, 5.5, 1.03, 22)
tx(s, f"{len(windows)}/{len(protocol['windows'])} windows", 6.92, 1.65, 5.62, 1.05, 52,
   face="Georgia", bold=True, ink=NAVY)
tx(s, f"Comparison unit = one window summary\n{valid_pairs}/{inventory['planned_followup_pairs']} route pairs (24–25 Oct 2026)",
   6.95, 2.77, 5.55, 1.03, 21)
rule(s, 0.75, 4.16, 11.83)
tx(s, "Population · Public OD pairs on selected Astana corridors\nSample · 20 fixed directed routes across two weekend dates",
   0.78, 4.47, 11.8, 1.13, 23)
tx(s, "Variables · Period, duration (s). Route-level correlation unavailable.",
   0.78, 5.76, 11.8, 0.49, 21)
footer(s, "A4 Tables 1–4", "assignment4/tables/inventory.json")
cue(s, "Akylbek", 30, 45,
    "Keep the denominators separate: 572 ORS rows passed the checks. The comparison kept four window summaries from 80 valid route pairs. The exploratory 20 pairs served as an instrument check. Our purposive sample covers public endpoints on selected arterial corridors, with no weekday observations.",
    "All times use Astana time, UTC+5. Peak was 17:30–19:30, and off-peak was 10:00–12:00 or 22:00–24:00. Requests outside those periods were classified as other. The protocol stopped at 17:40 on October 25.",
    ("Assignment 3 T6", "Assignment 4 Tables 1–4", SOURCE + "assignment4/tables/inventory.json"))

s = page("The peak gaps were larger")
tx(s, "Window median absolute gap (s)", 0.75, 1.46, 7.1, 0.4, 21, bold=True)
zone = timezone(timedelta(hours=5))
for index, window in enumerate(windows):
    timestamp = datetime.fromisoformat(window["start_utc"]).astimezone(zone)
    label = timestamp.strftime("%d Oct %H:%M")
    value = float(window["median_absolute_difference_s"])
    peak = window["period"] == "peak"
    y = 2.15 + index * 0.76
    tx(s, label, 0.75, y - 0.01, 2.34, 0.43, 18)
    width = value / 800 * 3.94
    shape(s, MSO_AUTO_SHAPE_TYPE.RECTANGLE, 3.15, y, width, 0.33,
          fill=NAVY if peak else OFFPEAK)
    tx(s, f"{value:.2f}", 3.26 + width, y - 0.01, 1.28, 0.43, 18, bold=True)
tx(s, "Period median of window medians", 8.5, 1.46, 4.15, 0.71, 20, bold=True)
tx(s, f"Peak  {peak_gap} s", 8.5, 2.17, 4.0, 0.48, 24, bold=True, ink=NAVY)
tx(s, f"Off-peak  {offpeak_gap} s", 8.5, 2.82, 4.02, 0.48, 24, bold=True, ink=MUTED)
tx(s, f"+{contrast} s", 8.5, 3.65, 4.02, 0.73, 38, face="Georgia", bold=True, ink=NAVY)
tx(s, f"{ratio}× peak/off-peak", 8.51, 4.52, 4.0, 0.52, 20)
tx(s, "Dark bars mark peak.", 0.75, 5.16, 7.5, 0.41, 18, ink=MUTED)
tx(s, "Amended expectation supported. Falsified if peak ≤ off-peak.",
   0.75, 5.66, 11.83, 0.6, 22, bold=True, ink=NAVY)
footer(s, "A4 Figures 2–3 · Table 4", "assignment4/tables/batch_results.csv")
cue(s, "Syrym", 40, 55,
    f"Each bar is the median absolute difference across 20 routes in one window. Dark blue marks peak. The period figures take the median of two window medians. Peak was {peak_gap} seconds and off-peak was {offpeak_gap}. The difference was {contrast} seconds, or about four and a half minutes. The amended expectation would fail if the peak summary were equal to or below the off-peak summary.",
    "The first peak window was the largest gap at 748.85 seconds. The morning off-peak window was the smallest at 379.40. The 1.664 ratio compares the two period summaries and applies to these four windows.",
    ("Assignment 4 Table 4 and Figures 2–3", SOURCE + "assignment4/tables/reported_metrics.json"))

s = page("Larger gaps, narrower spread")
tx(s, "Median signed gap was positive in all four windows.",
   0.75, 1.55, 11.8, 0.55, 24, bold=True, ink=NAVY)
rule(s, 0.75, 2.31, 11.8)
tx(s, "Signed spread (IQR)", 0.75, 2.64, 5.66, 0.48, 24, bold=True)
tx(s, "Peak  218.30–230.23 s\nOff-peak  261.30–279.48 s",
   0.75, 3.36, 5.83, 1.2, 25, bold=True, ink=NAVY)
tx(s, "Peak minus off-peak,\none window dropped", 6.74, 2.57, 5.81, 0.86, 23, bold=True)
tx(s, f"{low}–{high} s", 6.74, 3.47, 5.8, 0.74, 38, face="Georgia", bold=True, ink=NAVY)
tx(s, "The gap stayed positive.", 6.76, 4.39, 5.7, 0.54, 22)
tx(s, f"ORS duration stayed constant on {route_count}/{route_count} routes.",
   0.75, 5.68, 11.8, 0.58, 22, bold=True)
footer(s, "A4 Tables 4, 8 and 10", "assignment4/tables/window_sensitivity.csv")
cue(s, "Syrym", 30, 45,
    f"Positive signed medians mean Google typically returned a longer estimate. IQR is the spread of the middle half of the signed gaps. Both peak IQRs were smaller than both off-peak IQRs. Drop any single window and the peak contrast stays positive, between {low} and {high} seconds. ORS duration stayed the same within every route.",
    "The leave-one-window-out ratio ranges from 1.495 to 1.833. Similar route lengths covered 65% to 75% of pairs, so route choice and provider speed assumptions can contribute to the duration gap.",
    ("Assignment 4 Tables 4, 8 and 10", SOURCE + "assignment4/tables/window_sensitivity.csv"))

s = page("Four expectations stayed untested")
tx(s, "0 original TomTom–2GIS pairs collected", 0.75, 1.52, 11.8, 0.57, 24, bold=True)
tx(s, "Falsification rule", 6.1, 2.12, 6.3, 0.36, 18, bold=True)
for y, name, detail in [
    (2.75, "ER1 · Median absolute gap ratio ≥1.5", "Peak/off-peak ratio <1.5"),
    (3.55, "ER2 · ≥70% comparable lengths", "<70% within 5% of pair mean"),
    (4.35, "ER3 · Positive signed median", "TomTom minus 2GIS median ≤0"),
    (5.15, "ER4 · Greater peak spread", "Peak signed IQR ≤ off-peak IQR"),
]:
    row(s, y, name, detail, label_w=5.25)
footer(s, "A3 original rules · A4 Table 6", "docs/expected_results.md")
cue(s, "Yerdaulet", 25, 29,
    "All four rules belonged to TomTom–2GIS. We obtained zero eligible pairs, so all four stayed untested. Briefly name the expected ratio, length share, sign and spread. The observed Google–ORS ratio cannot serve as a test of ER1.",
    "ER4 also concerns signed spread, which is different from the size of the absolute gap.",
    ("Assignment 3 T9", "Assignment 4 Table 6", SOURCE + "docs/expected_results.md"))

s = page("The ORS instrument passed its checks")
for x, title, stat, observed, rule_text in [
    (0.75, "ER5 · Valid routes", f"{inventory['failed_ors']}/{observations} failed",
     "Supported in the archive", "Expected <2% failed\nFalsified at ≥2%"),
    (6.93, "ER6 · Response time", f"{observations}/{observations} within 10 s",
     f"Supported · Slowest {route_max} s", "Expected ≥95% within 10 s\nFalsified below 95%"),
]:
    tx(s, title, x, 1.69, 5.66, 0.54, 24, bold=True)
    tx(s, stat, x, 2.56, 5.67, 0.78, 34, face="Georgia", bold=True, ink=NAVY)
    tx(s, observed, x, 3.57, 5.67, 0.54, 21)
    rule(s, x, 4.47, 5.57)
    tx(s, rule_text, x, 4.8, 5.58, 1.03, 23)
footer(s, "A4 Tables 1–3 and 6", "assignment4/tables/expectations.csv")
cue(s, "Akylbek", 15, 21,
    f"ER5 and ER6 passed for the archive: zero failed route screens, and all {observations} responses finished within ten seconds. The slowest took {route_max} seconds.",
    "These tests measure instrument reliability for the archived requests. The thresholds were set before the pilot.",
    ("Assignment 4 Tables 1–3 and 6", SOURCE + "assignment4/tables/expectations.csv"))

s = page("The scheduled target was missed")
for x, title, stat, observed, rule_text in [
    (0.75, "ER7 · Pilot coverage", f"{route_count}/{route_count} routes",
     "Supported · Valid in all three cycles", "Expected all routes to return a valid pilot result\nFalsified below 20 routes"),
    (6.93, "ER8 · Automatic coverage", f"{automatic}/{planned_automatic} attempts",
     f"Missed · {inventory['automatic_pct_original_plan']:.3f}% in 48-hour plan", "Expected ≥90%\nFalsified below 692 attempts"),
]:
    tx(s, title, x, 1.69, 5.65, 0.54, 24, bold=True)
    tx(s, stat, x, 2.56, 5.65, 0.78, 38, face="Georgia", bold=True, ink=NAVY)
    tx(s, observed, x, 3.57, 5.65, 0.56, 21)
    rule(s, x, 4.47, 5.57)
    tx(s, rule_text, x, 4.78, 5.58, 1.25, 21)
footer(s, "A4 Tables 1 and 6 · Figure 1", "assignment4/tables/expectations.csv")
cue(s, "Akylbek", 20, 26,
    f"All {route_count} routes worked in all three pilot cycles, supporting ER7. ER8 required at least 692 scheduled attempts. We recorded {automatic} automatic attempts out of {planned_automatic}, or 53.125 percent. Four manual checks are counted separately.",
    "The four completed Google–ORS windows met a separate amended target. A complete comparison batch does not fill a gap in scheduled ORS coverage.",
    ("Assignment 4 Tables 1 and 6 and Figure 1", SOURCE + "assignment4/tables/inventory.json"))

s = page("What each objective delivered")
for y, title, detail in [
    (1.76, "O1 · Partial", "20 fixed routes and periods. Representativeness untested."),
    (2.68, "O2 · Missed", "80 Google–ORS pairs over two days. Yandex absent."),
    (3.60, "O3 · Partial", "Gap and spread metrics. Confidence intervals unavailable."),
    (4.52, "O4 · Partial", f"{contrast} s descriptive contrast. Inference unavailable."),
    (5.44, "O5 · Partial", "Windows differed. Recurring route patterns unresolved."),
]:
    row(s, y, title, detail, size=21)
footer(s, "A4 Table 5 · Figure 3", "assignment4/tables/objectives.csv")
cue(s, "Syrym", 30, 37,
    "O1 fixed routes and periods, with purposive sampling. O2 missed the original providers and dates. O3 retained medians, MAE, RMSE and IQR, with no route-level confidence intervals. O4 gives a descriptive period contrast. O5 shows differences across windows, with recurring patterns unresolved.",
    "MAE is the average absolute gap. RMSE gives larger gaps more weight. Correlation would describe co-movement between estimates, and needs paired route-level values here.",
    ("Assignment 4 Table 5 and §3.5", SOURCE + "assignment4/tables/objectives.csv"))

s = page("Answer and limits", dark=True)
tx(s, f"On the sampled routes and dates, Figure 3 shows {contrast} seconds more median absolute disagreement between Google and ORS at peak than off-peak.",
   0.75, 1.73, 11.6, 2.14, 33, face="Georgia", bold=True, ink=PAPER)
rule(s, 0.75, 4.19, 11.82, fill=SOFT)
tx(s, "Aggregate retention prevents route-level testing.\nTwo weekend dates limit generalization.\nAccuracy needs timed journeys. Provider and traffic effects overlap.",
   0.75, 4.53, 11.6, 1.52, 23, ink=PAPER)
footer(s, "A4 Figure 3 · §3.11", "assignment4/tables/reported_metrics.json", dark=True)
cue(s, "Syrym", 35, 43,
    "Give the one-sentence answer. We can describe the sampled estimates. We cannot reconstruct individual route differences or test a weekday pattern from two weekend dates. Google and ORS differ in provider and traffic inputs, so the design cannot isolate a traffic effect. Independent timed journeys would be needed to rank accuracy.",
    "A future comparison could seek permission to retain both sets of estimates and geometries, sample weekdays and weekends, and add independently timed journeys.",
    ("Assignment 4 Figure 3, §3.11 and conclusion", SOURCE + "assignment4/tables/reported_metrics.json"))

s = page("Division of labour")
tx(s, "10 weeks · Shared authorship across A1–A4", 0.75, 1.60, 11.8, 0.58, 25, bold=True)
tx(s, "Question and literature, collection protocol, analysis and reporting.",
   0.75, 2.35, 11.8, 0.71, 23)
rule(s, 0.75, 3.29, 11.8)
tx(s, "Defence responsibilities", 0.75, 3.58, 11.8, 0.58, 24, face="Georgia", bold=True, ink=NAVY)
for x, name, detail in [
    (0.75, "Yerdaulet Akhanov", "Framing, literature,\noriginal comparison"),
    (4.84, "Syrym Shadiyarbek", "Metrics, window results,\nconclusion"),
    (8.94, "Akylbek Sabyrzhan", "Route sampling, collection,\ninstrument checks"),
]:
    tx(s, name, x, 4.55, 3.58, 0.49, 21, bold=True, ink=NAVY)
    tx(s, detail, x, 5.16, 3.61, 0.94, 20)
tx(s, "Individual research tasks remain unverified.", 0.75, 6.68, 11.8, 0.4, 18, ink=MUTED)
cue(s, "Akylbek", 25, 32,
    "The submitted reports name all three team members. The record covers the question and literature review, instrument setup, and the final analysis and report. Name who prepared and presented each defence section.",
    "The planned responsibilities in the contribution file give each member a lead area to rehearse. Each should be ready to explain a concrete project decision they actually remember.",
    ("A1–A4 title pages", "CONTRIBUTION-DISTRIBUTION.txt, proposed A3/A4 responsibilities"))

s = page("A1 research question and hypothesis")
tx(s, "How does evening peak traffic affect the magnitude and variability of travel-time discrepancies between Google Maps and Yandex Maps across selected arterial routes in Astana?",
   0.75, 1.72, 11.7, 2.16, 33, face="Georgia", bold=True, ink=NAVY)
tx(s, "H1 expected significantly greater peak gaps. Falsified if the test lacked evidence at the predefined significance level. Untested after the provider change.",
   0.75, 4.48, 11.68, 1.62, 24)
cue(s, "Backup", 0, 0,
    "Historical A1 wording for Q&A. H0 stated that the magnitude of discrepancy does not differ by period. The original synchronized provider pair and 14-day plan were not executed.",
    sources=("Assignment 1 T1 and T3",))

s = page("A2 research question")
tx(s, "How do the magnitude and variability of discrepancies between Google Maps and 2GIS traffic-aware car travel-time estimates differ between evening peak and off-peak periods for fixed origin-destination pairs on Astana’s arterial corridors, after accounting for differences in the returned routes?",
   0.75, 1.72, 11.7, 3.42, 32, face="Georgia", bold=True, ink=NAVY)
cue(s, "Backup", 0, 0,
    "Historical A2 wording for Q&A. The amended A3 and A4 question is on slide 2. No quantitative 2GIS comparison was performed.",
    sources=("Assignment 2 §7.1",))


def write_notes(output):
    lines = [
        "# RMT final defence rehearsal", "",
        "Use the 7:00 cut for the exam. It assigns 2:40 to Results and Discussion.",
        "The 9:00 rehearsal assigns three minutes to each teammate and 3:33 to Results and Discussion.",
        "The guide’s per-person minimum and total cap cannot both be met with three presenters.",
        "",
        "| Slide | Speaker | Exam | Rehearsal |",
        "|---|---|---:|---:|",
    ]
    for i, label, who, exam, rehearsal, *_ in notes:
        if exam:
            lines.append(f"| {i}. {label.replace(chr(10), ' ')} | {who} | {exam}s | {rehearsal}s |")
    lines.extend(("", "## Speaking cues", ""))
    for i, label, who, exam, rehearsal, core, extra, sources in notes:
        lines.extend((f"### {i}. {label.replace(chr(10), ' ')}", "", core, ""))
        if extra:
            lines.extend(("**Rehearsal extension**", "", extra, ""))
        if sources:
            lines.extend(("Evidence", "", *(f"- {source}" for source in sources), ""))
    lines.extend((
        "## Q&A practice", "",
        f"- **Why are there two sample counts?** Each of the {observations} ORS rows is a request. "
        f"The period analysis uses {len(windows)} window summaries, based on {valid_pairs} valid pairs.",
        "- **How did you reduce bias?** Fixed endpoints, predefined windows and a 60-second pairing rule "
        "limited timing and endpoint differences. Purposive sampling and different provider models remain limitations.",
        f"- **Which expectation failed?** ER8 required at least 692 automatic attempts. "
        f"The archive contains {automatic}, or 53.125% of {planned_automatic}. The four manual checks are separate.",
        "- **Why no correlation or Wilcoxon test?** Individual paired Google values were processed transiently. "
        "Window aggregates cannot recover route-level co-movement or paired route-day differences.",
        f"- **What did the sensitivity check show?** Dropping any one window left a peak contrast of "
        f"{low}–{high} seconds. The direction held across that check.",
        "- **What would establish accuracy?** Independently timed drives on the same routes and at matching times.",
        "",
    ))
    (output / "RMT_Rehearsal_Notes.md").write_text("\n".join(lines), encoding="utf-8")


def main():
    output = HERE / "output"
    output.mkdir(exist_ok=True)
    counts = [len(re.findall(r"\S+", " ".join(lines))) for lines in audit]
    assert max(counts) <= 60, counts
    assert sum(note[3] for note in notes) == 420
    assert sum(note[3] for note in notes[6:12]) == 160
    assert sum(note[4] for note in notes[6:12]) == 213
    for member in ("Yerdaulet", "Akylbek", "Syrym"):
        assert sum(note[4] for note in notes if note[2] == member) == 180
    deck.save(output / "RMT_Final_Defence_Team_Draft.pptx")
    write_notes(output)
    print("Slide word counts:", counts)


if __name__ == "__main__":
    main()
