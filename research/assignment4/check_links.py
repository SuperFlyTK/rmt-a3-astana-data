"""Check every external report link without credentials or routing API calls."""
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timezone
from pathlib import Path
import json
import urllib.error
import urllib.request

HERE = Path(__file__).resolve().parent


def check(url):
    request_url = url
    # Raw content is a direct existence check for the exact file linked by GitHub.
    if url.startswith("https://github.com/") and "/blob/" in url:
        request_url = url.replace("https://github.com/", "https://raw.githubusercontent.com/", 1).replace("/blob/", "/", 1)
    try:
        request = urllib.request.Request(request_url, headers={"User-Agent":"RMT-study-link-check/1.0"})
        with urllib.request.urlopen(request, timeout=30) as response:
            first = response.read(512)
            pdf = request_url.split("?")[0].endswith(".pdf")
            good = response.status == 200 and (not pdf or first.startswith(b"%PDF"))
            return dict(url=url,request_url=request_url,status=response.status,ok=good,
                        final_url=response.url,expected_pdf=pdf)
    except (urllib.error.URLError, OSError) as error:
        return dict(url=url,request_url=request_url,ok=False,error=str(error))


def main():
    urls = json.loads((HERE/"report_links.json").read_text())
    with ThreadPoolExecutor(max_workers=5) as pool:
        results = list(pool.map(check, urls))
    report = dict(checked_at_utc=datetime.now(timezone.utc).isoformat(),
                  check_method="Unauthenticated GET; GitHub blob links checked using their exact raw-content URL; PDF links require a PDF header.",
                  total=len(results),passed=sum(r["ok"] for r in results),results=results)
    (HERE/"link_checks.json").write_text(json.dumps(report,indent=2),encoding="utf-8")
    print(json.dumps({k:report[k] for k in ("total","passed")}))
    for result in results:
        if not result["ok"]:
            print(json.dumps(result))
    return 0 if report["total"] == report["passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
