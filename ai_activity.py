"""Day-by-day activity of every Copilot / AI surface from the Purview audit log (all CopilotInteraction records).

Streams the records of one Graph audit search (no raw copy: a month is ~0.5M records) and writes
  <usage-dir>/ai_activity/daily_<from>_<to>.csv   date x solution x app x agent x user x model provider
Merge several periods by keeping each file; ai_cost_report.py reads them all.

    python ai_activity.py "<Usage dir>" --from 2026-05-01 --to 2026-10-01    # create search, wait, aggregate
    python ai_activity.py "<Usage dir>" --query-id <id>                        # search created earlier (portal or API)
"""
import argparse, collections, csv, datetime as dt, pathlib, re, sys, time
import requests
import pull_credits as pc
from cowork_audit import Q, session, wait

SOLUTIONS = [  # first match wins; AppIdentity examples: Copilot.M365Copilot.CoworkChat, Copilot.Studio.<id>
    (re.compile(r"Cowork", re.I), "Copilot Cowork"),
    (re.compile(r"\bCode\b|CopilotCode|\.Code", re.I), "Copilot Code"),
    (re.compile(r"^Copilot\.Studio\b", re.I), "Copilot Studio agents"),
    (re.compile(r"^MicrosoftAgent\.", re.I), "Microsoft prebuilt agents"),
    (re.compile(r"Bizchat|WebChat", re.I), "Copilot Chat"),
    (re.compile(r"M365Copilot\.Apps", re.I), "Copilot in M365 apps"),
    (re.compile(r"GitHub", re.I), "GitHub Copilot"),
    (re.compile(r"Fabric|PowerBI", re.I), "Copilot in Fabric / Power BI"),
    (re.compile(r"^Notebook_", re.I), "Copilot Notebooks"),
]
COLS = ["date", "solution", "app", "agent", "user", "provider", "interactions", "threads", "license"]


def solution(app: str) -> str:
    return next((name for rx, name in SOLUTIONS if rx.search(app)), "Other")


def app_family(app: str) -> str:
    """Copilot.Studio.Default-<tenant>-<agent> -> Copilot.Studio.Default; keeps agent detail in the agent column."""
    if app.startswith("Notebook_"):
        return "Notebook"
    return re.sub(r"\.(T_|Default-)?[0-9a-f-]{20,}.*$", "", app, flags=re.I) or app


def create(s, start, end) -> str:
    body = {"displayName": f"ai_activity {start:%Y-%m-%d}..{end:%Y-%m-%d}", "recordTypeFilters": ["copilotInteraction"],
            "filterStartDateTime": f"{start:%Y-%m-%dT%H:%M:%SZ}", "filterEndDateTime": f"{end:%Y-%m-%dT%H:%M:%SZ}"}
    r = s.post(Q, json=body, timeout=60)
    if not r.ok:
        sys.exit(f"{r.status_code} creating audit search: {r.text[:300]}")
    print(f"search {r.json()['id']} created", flush=True)
    return r.json()["id"]


def stream(s, url):
    """pc.pages with a token refresh: a 0.5M-record read outlives the 1 h access token."""
    while url:
        for _ in range(6):
            r = s.get(url, timeout=120)
            if r.status_code == 401:
                s.headers["Authorization"] = f"Bearer {pc.token(False, ['https://graph.microsoft.com/AuditLogsQuery.Read.All'])}"
            elif r.status_code in (429, 503, 504):
                time.sleep(int(r.headers.get("Retry-After", 10)))
            else:
                break
        if not r.ok:
            sys.exit(f"{r.status_code} reading records: {r.text[:300]}")
        body = r.json()
        yield from body.get("value", [])
        url = body.get("@odata.nextLink")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("usage_dir")
    ap.add_argument("--from", dest="start")
    ap.add_argument("--to", dest="end")
    ap.add_argument("--query-id")
    ap.add_argument("--max-wait-min", type=int, default=240)
    ap.add_argument("--unattended", action="store_true")
    a = ap.parse_args()

    s = session(not a.unattended)
    if a.query_id:
        qid = a.query_id
    else:
        parse = lambda x: dt.datetime.fromisoformat(x).replace(tzinfo=dt.timezone.utc)
        qid = create(s, parse(a.start), parse(a.end))
    q = wait(s, qid, a.max_wait_min)
    start, end = (q[k][:10].replace("-", "") for k in ("filterStartDateTime", "filterEndDateTime"))

    n, agg, threads = 0, collections.Counter(), collections.defaultdict(set)
    lic = {}
    for rec in stream(s, f"{Q}/{qid}/records?$top=1000"):
        ad = rec.get("auditData") or {}
        ev = ad.get("CopilotEventData") or {}
        app = str(ad.get("AppIdentity") or "unknown")
        prov = ((ev.get("ModelTransparencyDetails") or [{}])[0]).get("ModelProviderName", "")
        key = (str(rec.get("createdDateTime", ""))[:10], solution(app), app_family(app), ad.get("AgentName", ""),
               (rec.get("userPrincipalName") or ad.get("UserId") or "").lower(), prov)
        agg[key] += 1
        threads[key].add(ev.get("ThreadId") or ev.get("ConversationId"))
        lic[key] = ev.get("LicenseType", "")
        n += 1
        if n % 50000 == 0:
            print(f"{n:,} records", flush=True)

    out = pathlib.Path(a.usage_dir) / "ai_activity"
    out.mkdir(parents=True, exist_ok=True)
    f = out / f"daily_{start}_{end}.csv"
    with open(f, "w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(COLS)
        for k, v in sorted(agg.items()):
            w.writerow([*k, v, len(threads[k] - {None}), lic[k]])
    print(f"{n:,} records -> {len(agg):,} rows in {f}")


if __name__ == "__main__":
    main()
