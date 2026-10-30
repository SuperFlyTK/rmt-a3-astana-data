# Assignment 4 evidence

The completed manuscript is [the team Word report](RMT_A4_Team_AkhanovYerdaulet_ShadiyarbekSyrym_SabyrzhanAkylbek.docx).

This folder analyses the archived collection through 25 October 2026 at 17:30:53 Astana time. It contains 572 ORS observations, 100 Google request-status records and five aggregate comparisons. Four comparisons belong to the planned follow-up and contain 80 valid pairs; the exploratory comparison is separate.

## Browse the files

- [Data snapshot](data): the final export used in Assignment 4, with all collection phases labelled.
- [Tables](tables): inventory, row accounting, variable levels and missingness, coverage, every comparison batch, period summaries, original objectives, unchanged expectations and sensitivity checks.
- [Figures](figures): six plots made from the archived measurements and comparison summaries.
- [Evidence manifest](evidence_manifest.json): input, script, table and figure hashes.
- [Submission checks](submission_checks.json): body-only word counts and report counts.
- [Verification](verification.json): checks on source integrity, expectations, report contents and numeric coverage.
- [Link checks](link_checks.json): recorded public-link responses, with new links marked pending until publication. Run `python assignment4/check_links.py` after publishing to check every link.

The source repository collects observations and exports them to `SuperFlyTK/rmt-a3-astana-data`. The data repository is a publication copy. The earlier copy was missing 64 ORS records and the final planned paired window. The Assignment 4 snapshot includes them. The Assignment 3 report, original frozen samples and preregistered expectations remain historical records.

## Reproduce

From the repository root, with Python 3.11 or later:

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe assignment4/analyse.py
.\.venv\Scripts\python.exe assignment4/build_report.py
.\.venv\Scripts\python.exe assignment4/verify.py
```

Exact analysis package versions are listed in [requirements-used.txt](requirements-used.txt). Text-file hashes normalize CRLF to LF so Git's Windows newline conversion does not invalidate the manifest; figures use exact byte hashes. The DOCX has its own byte hash, and rebuilding may change its ZIP timestamps. Rebuilding uses the archived collection and makes no routing requests. The scripts stop for review if the report's stated final sample changes.

## Interpretation limits

The period comparison uses the median of window median absolute differences. It does not pool or reconstruct individual Google estimates. No confidence interval, p-value, accuracy ranking or causal traffic effect is reported. The original TomTom-2GIS expectations remain untested; the original automatic collection target was missed. A four-window sensitivity check is descriptive and is not an uncertainty interval.

ORS attribution: © openrouteservice by HeiGIT | Data from OpenStreetMap. Endpoint source: © OpenStreetMap contributors. Provider terms and the source licences continue to apply.
