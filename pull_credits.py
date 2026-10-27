"""Daily snapshot of Copilot credit consumption per user, via the Graph beta Cost Management API.

Sign-in is delegated (the API has no app-only route for user data): the first run prints a
device code; later runs reuse the refresh token in a local MSAL cache until Entra expires it.
The token is never printed. Output: <usage-dir>/snapshots/credits_<YYYY-MM-DD>.csv

    python pull_credits.py "<path to Teams channel>/Usage"
"""
import csv, datetime as dt, json, os, pathlib, sys, time
import msal, requests

CLIENT = "14d82eec-204b-4c2f-b7e8-296a70dab67e"  # Microsoft Graph Command Line Tools
SCOPES = [f"https://graph.microsoft.com/{s}" for s in (
    "CopilotCostManagement.Read.All", "CopilotCostManagement-Policy.Read.All",
    "CopilotCostManagement-Assignment.Read.All", "CopilotCostManagement-UserData.Read.All",
    "GroupMember.Read.All", "Sites.Read.All", "Reports.Read.All")]
# SharePoint list with user -> CR, e.g. CR_SITE="contoso.sharepoint.com:/sites/MyTeam". Unset = skip the list.
CR_SITE, CR_LIST = os.environ.get("CR_SITE"), os.environ.get("CR_LIST", "CR Mapping")
G = "https://graph.microsoft.com"
CM = f"{G}/beta/copilot/costManagement"
CACHE = pathlib.Path.home() / ".config/copilot-credits/msal_cache.json"


def token(interactive: bool, scopes: list = SCOPES) -> str:
    cache = msal.SerializableTokenCache()
    if CACHE.exists():
        cache.deserialize(CACHE.read_text())
    app = msal.PublicClientApplication(CLIENT, token_cache=cache,
                                       authority="https://login.microsoftonline.com/organizations", timeout=30)
    accts = app.get_accounts()
    res = app.acquire_token_silent(scopes, account=accts[0]) if accts else None
    if not res:
        if not interactive:
            sys.exit("Sign-in expired: run once by hand to get a new device code.")
        flow = app.initiate_device_flow(scopes=scopes)
        print(flow.get("message", flow), flush=True)
        while time.time() < flow["expires_at"]:
            try:
                res = app.acquire_token_by_device_flow(flow)
                break
            except requests.exceptions.RequestException as e:  # network blip while polling: keep the same code
                print(f"network error while waiting for sign-in, retrying: {type(e).__name__}", flush=True)
        else:
            sys.exit("Device code expired before sign-in finished.")
    if "access_token" not in res:
        sys.exit(f"Sign-in failed: {res.get('error')}: {res.get('error_description')}")
    if cache.has_state_changed:
        CACHE.parent.mkdir(parents=True, exist_ok=True)
        CACHE.write_text(cache.serialize())
        os.chmod(CACHE, 0o600)
    return res["access_token"]


def pages(s: requests.Session, url: str):
    while url:
        for attempt in range(5):
            r = s.get(url, timeout=60)
            if r.status_code != 429:
                break
            time.sleep(int(r.headers.get("Retry-After", 5)))
        if not r.ok:
            raise SystemExit(f"{r.status_code} on {url}: {r.text[:300]}")
        body = r.json()
        yield from body.get("value", [])
        url = body.get("@odata.nextLink")


def export_cr_list(s: requests.Session, out_dir: pathlib.Path):
    """Copy the SharePoint 'CR Mapping' list (UserEmail, CR_code, DateFrom, DateTo) to cr_mapping.csv."""
    if not CR_SITE:
        return
    r = s.get(f"{G}/v1.0/sites/{CR_SITE}", timeout=60)
    site = r.json() if r.ok else None
    lst = site and next(pages(s, f"{G}/v1.0/sites/{site['id']}/lists?$filter=displayName eq '{CR_LIST}'"), None)
    if not lst:
        print(f"CR list '{CR_LIST}' not found on {CR_SITE} ({r.status_code}), keeping previous cr_mapping.csv")
        return
    cols = ["UserEmail", "CR_code", "DateFrom", "DateTo"]
    with open(out_dir / "cr_mapping.csv", "w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=cols, extrasaction="ignore")
        w.writeheader()
        for it in pages(s, f"{G}/v1.0/sites/{site['id']}/lists/{lst['id']}/items?expand=fields&$top=999"):
            f = it["fields"]
            w.writerow({**f, "UserEmail": f.get("UserEmail") or f.get("Title")})  # first column is often a renamed Title
    print(f"CR list exported from {site.get('webUrl')}")


def export_copilot_usage(s: requests.Session, out: pathlib.Path, today: str):
    """Licensed M365 Copilot activity (last 28 days): per-user CSV (prompts, active days, last use per app),
    plus per-app summary and daily trend. No Cowork fields: Cowork tasks come only from the admin-center export."""
    r = s.get(f"{G}/v1.0/copilot/reports/getMicrosoft365CopilotUsageUserDetail(period='D28')", timeout=120)
    if not r.ok:
        raise SystemExit(f"{r.status_code} usage detail: {r.text[:300]}")
    (out / f"copilot_usage_{today}.csv").write_bytes(r.content)
    for name in ("UserCountSummary", "UserCountTrend"):
        rows = list(pages(s, f"{G}/beta/reports/getMicrosoft365Copilot{name}(period='D28')?$format=application/json"))
        (out / f"copilot_{name[9:].lower()}_{today}.json").write_text(json.dumps(rows, indent=1))
    print(f"Copilot usage: {r.text.count(chr(10)) - 1} licensed-user rows, summary + trend saved")


def main(usage_dir: str, interactive: bool = True):
    out = pathlib.Path(usage_dir) / "snapshots"
    out.mkdir(exist_ok=True)
    s = requests.Session()
    s.headers["Authorization"] = f"Bearer {token(interactive)}"
    today = dt.date.today().isoformat()

    export_cr_list(s, out.parent)
    policies = list(pages(s, f"{CM}/spendingPolicies"))
    (out / f"policies_{today}.json").write_text(json.dumps(policies, indent=1))

    users = {}
    for p in policies:
        for g in pages(s, f"{CM}/spendingPolicies/{p['id']}/microsoft.graph.selectedGroupsSpendingPolicy/assignedGroups"):
            for u in pages(s, f"{G}/v1.0/groups/{g['id']}/transitiveMembers/microsoft.graph.user"
                              "?$select=id,userPrincipalName,displayName&$top=999"):
                users[u["id"]] = u
    # ponytail: all-users policies have no groups; enumerate /users if one is ever enabled.

    rows = []
    for uid, u in users.items():
        for b in pages(s, f"{CM}/userBalances/{uid}/serviceBalances"):
            if b.get("consumedQuantity"):
                rows.append({"snapshotDate": today, "upn": u["userPrincipalName"].lower(),
                             "displayName": u["displayName"], "userId": uid,
                             "spendingPolicyId": b.get("spendingPolicyId"), "serviceId": b["serviceId"],
                             "creditsMTD": round(b["consumedQuantity"], 2),
                             "asOf": (b.get("dataStatus") or {}).get("lastModifiedDateTime")})
    f = out / f"credits_{today}.csv"
    with open(f, "w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=list(rows[0]) if rows else ["snapshotDate"])
        w.writeheader(); w.writerows(rows)
    try:
        export_copilot_usage(s, out, today)
    except SystemExit as e:  # usage is secondary: keep the credit snapshot even if reports fail
        print(f"Copilot usage report skipped: {e}")
    print(f"{len(users)} users in policy groups, {len(rows)} rows with credits -> {f}")


if __name__ == "__main__":
    main(sys.argv[1], interactive="--unattended" not in sys.argv)
