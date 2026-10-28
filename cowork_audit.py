"""Cowork activity from the Purview unified audit log, via the Graph Audit Log Query API.

Creates one audit search (record type copilotInteraction, keyword CoworkChat), waits for it, reads every record,
keeps only Cowork (AppIdentity ...CoworkChat) and writes:
  <usage-dir>/cowork_audit/raw/cowork_<from>_<to>.jsonl   full records (named per-user data, keep restricted)
  <usage-dir>/cowork_audit/cowork_events.csv               one row per interaction, merged across runs

Sign-in reuses pull_credits' delegated cache (needs Purview Audit Reader + AuditLogsQuery.Read.All).
Prompt/response text is not in the audit log; that needs eDiscovery.

    python cowork_audit.py "<Usage dir>" --days 7
    python cowork_audit.py "<Usage dir>" --from 2026-10-01 --to 2026-10-28
    python cowork_audit.py "<Usage dir>" --query-id <id>     # resume a search created earlier
"""
import argparse, csv, datetime as dt, json, pathlib, sys, time
import requests
import pull_credits as pc

G = "https://graph.microsoft.com"
Q = f"{G}/v1.0/security/auditLog/queries"
COLS = ["time_utc", "user", "client_ip", "thread_id", "model", "provider", "prompts", "tools", "items_read", "item_types",
        "labels", "approvals_requested", "approvals_granted", "record_id"]


def session(interactive: bool) -> requests.Session:
    s = requests.Session()
    s.headers["Authorization"] = f"Bearer {pc.token(interactive, [f'{G}/AuditLogsQuery.Read.All'])}"
    return s


def create(s, start: dt.datetime, end: dt.datetime) -> str:
    body = {"displayName": f"cowork_audit {start:%Y-%m-%d}..{end:%Y-%m-%d}", "recordTypeFilters": ["copilotInteraction"],
            "keywordFilter": "CoworkChat",
            "filterStartDateTime": f"{start:%Y-%m-%dT%H:%M:%SZ}", "filterEndDateTime": f"{end:%Y-%m-%dT%H:%M:%SZ}"}
    r = s.post(Q, json=body, timeout=60)
    if not r.ok:
        sys.exit(f"{r.status_code} creating audit search: {r.text[:300]}")
    print(f"search {r.json()['id']} created", flush=True)
    return r.json()["id"]


def wait(s, qid: str, max_min: int):
    for _ in range(max_min * 2):
        try:
            r = s.get(f"{Q}/{qid}", timeout=120)
            status = r.json().get("status") if r.ok else f"http {r.status_code}"
        except requests.exceptions.RequestException as e:  # slow Graph response: keep polling
            status = type(e).__name__
        if status == "succeeded":
            return r.json()
        if status in ("failed", "cancelled"):
            sys.exit(f"audit search {qid} {status}")
        time.sleep(30)
    sys.exit(f"audit search {qid} still running after {max_min} min; rerun with --query-id {qid}")


def flatten(rec: dict) -> dict:
    ad = rec.get("auditData") or {}
    ev = ad.get("CopilotEventData") or {}
    res = ev.get("AccessedResources") or []
    reads = [x for x in res if not str(x.get("Type", "")).startswith("approval")]
    mt = (ev.get("ModelTransparencyDetails") or [{}])[0]
    names = lambda t: ";".join(sorted({x.get("Name", "") for x in res if x.get("Type") == t}))
    return {"time_utc": rec.get("createdDateTime"), "user": rec.get("userPrincipalName") or ad.get("UserId"),
            "client_ip": rec.get("clientIp") or ad.get("ClientIP"), "thread_id": ev.get("ThreadId") or ev.get("ConversationId"),
            "model": mt.get("ModelName", ""), "provider": mt.get("ModelProviderName", ""),
            "prompts": sum(1 for m in ev.get("Messages") or [] if m.get("isPrompt")),
            "tools": ";".join(sorted({p.get("Name", "") for p in ev.get("AISystemPlugin") or []})),
            "items_read": len(reads), "item_types": ";".join(sorted({str(x.get("Type")) for x in reads})),
            "labels": ";".join(sorted({x["SensitivityLabelId"] for x in reads if x.get("SensitivityLabelId")})),
            "approvals_requested": names("approvalRequested"), "approvals_granted": names("approvalGrantedOnce"),
            "record_id": rec.get("id")}


def is_cowork(rec: dict) -> bool:
    ad = rec.get("auditData") or {}
    return "CoworkChat" in str(ad.get("AppIdentity", "")) or (ad.get("CopilotEventData") or {}).get("AppHost") == "cowork"


def merge_csv(path: pathlib.Path, rows: list):
    old = list(csv.DictReader(open(path, encoding="utf-8"))) if path.exists() else []
    by_id = {r["record_id"]: r for r in old + rows}
    out = sorted(by_id.values(), key=lambda r: r["time_utc"] or "")
    with open(path, "w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=COLS)
        w.writeheader()
        w.writerows(out)
    return len(out)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("usage_dir")
    ap.add_argument("--days", type=int, default=7)
    ap.add_argument("--from", dest="start")
    ap.add_argument("--to", dest="end")
    ap.add_argument("--query-id")
    ap.add_argument("--max-wait-min", type=int, default=90)
    ap.add_argument("--unattended", action="store_true")
    a = ap.parse_args()

    now = dt.datetime.now(dt.timezone.utc).replace(microsecond=0)
    end = dt.datetime.fromisoformat(a.end).replace(tzinfo=dt.timezone.utc) if a.end else now
    start = dt.datetime.fromisoformat(a.start).replace(tzinfo=dt.timezone.utc) if a.start else end - dt.timedelta(days=a.days)
    out = pathlib.Path(a.usage_dir) / "cowork_audit"
    (out / "raw").mkdir(parents=True, exist_ok=True)

    s = session(not a.unattended)
    qid = a.query_id or create(s, start, end)
    q = wait(s, qid, a.max_wait_min)
    start, end = (dt.datetime.fromisoformat(q[k].replace("Z", "+00:00")) for k in ("filterStartDateTime", "filterEndDateTime"))
    recs = [r for r in pc.pages(s, f"{Q}/{qid}/records?$top=1000") if is_cowork(r)]
    raw = out / "raw" / f"cowork_{start:%Y%m%d}_{end:%Y%m%d}.jsonl"
    raw.write_text("".join(json.dumps(r, ensure_ascii=False) + "\n" for r in recs), encoding="utf-8")
    total = merge_csv(out / "cowork_events.csv", [flatten(r) for r in recs])
    print(f"{len(recs)} Cowork records {start:%Y-%m-%d}..{end:%Y-%m-%d}; cowork_events.csv now {total} rows")


if __name__ == "__main__":
    main()
