"""Read-only probe: which cost/usage sources can the signed-in admin reach?

Checks, one line each (OK / FAIL + HTTP status and reason):
  1. Graph beta Copilot Cost Management  - tenant credit balance, spending policies (reuses pull_credits sign-in)
  2. Azure billing accounts              - agreement type (EA or MCA)
  3. Azure Cost Management               - month-to-date cost per subscription, Copilot / Power Platform meters
  4. Power Platform licensing API        - Copilot Studio credits per agent (yesterday)
  5. Graph subscribedSkus                - licences bought vs assigned

Sign-in is delegated with the Azure CLI public client (one device code covers Azure, Power Platform and Graph);
tokens stay in ~/.config/copilot-credits/ and are never printed. Nothing is written to the tenant.

    python probe_sources.py
"""
import datetime as dt, pathlib, re, sys, time
import msal, requests
import pull_credits

AZ_CLIENT = "04b07795-8ddb-461a-bbee-02f9e1bf7b46"  # Azure CLI
AZ_CACHE = pathlib.Path.home() / ".config/copilot-credits/msal_cache_azcli.json"
ARM, PPAPI, G = "https://management.azure.com", "https://api.powerplatform.com", "https://graph.microsoft.com"
AI_METERS = re.compile(r"copilot|power platform|agent|credit|dataverse", re.I)


class AzSignIn:
    """One device-code sign-in, then silent tokens for any resource the Azure CLI client may call."""

    def __init__(self):
        self.cache = msal.SerializableTokenCache()
        if AZ_CACHE.exists():
            self.cache.deserialize(AZ_CACHE.read_text())
        self.app = msal.PublicClientApplication(AZ_CLIENT, token_cache=self.cache,
                                                authority="https://login.microsoftonline.com/organizations")

    def token(self, resource: str) -> str:
        scopes = [f"{resource}/.default"]
        accts = self.app.get_accounts()
        res = self.app.acquire_token_silent(scopes, account=accts[0]) if accts else None
        if not res:
            flow = self.app.initiate_device_flow(scopes=scopes)
            print(flow.get("message", flow), flush=True)
            res = self.app.acquire_token_by_device_flow(flow)
        if self.cache.has_state_changed:
            AZ_CACHE.parent.mkdir(parents=True, exist_ok=True)
            AZ_CACHE.write_text(self.cache.serialize())
            AZ_CACHE.chmod(0o600)
        if "access_token" not in res:
            raise RuntimeError(f"sign-in for {resource} failed: {res.get('error')}: "
                               f"{(res.get('error_description') or '')[:200]}")
        return res["access_token"]


def call(method: str, url: str, tok: str, **kw) -> requests.Response:
    for _ in range(5):
        r = requests.request(method, url, headers={"Authorization": f"Bearer {tok}"}, timeout=90, **kw)
        if r.status_code != 429:
            return r
        time.sleep(int(r.headers.get("Retry-After", 10)))
    return r


def fail(r: requests.Response) -> str:
    try:
        err = r.json().get("error", {})
        msg = err.get("message", err) if isinstance(err, dict) else err
    except ValueError:
        msg = r.text
    return f"FAIL {r.status_code}: {str(msg)[:220]}"


def probe_graph_cost():
    tok = pull_credits.token(interactive=True)
    r = call("GET", f"{pull_credits.CM}/getTenantCreditBalance()", tok)
    if not r.ok:
        return fail(r)
    bal = {k: v for k, v in r.json().items() if not k.startswith("@")}
    r2 = call("GET", f"{pull_credits.CM}/spendingPolicies", tok)
    pol = len(r2.json().get("value", [])) if r2.ok else fail(r2)
    return f"OK  tenant balance {bal} | spending policies: {pol}"


def probe_billing(az: AzSignIn):
    r = call("GET", f"{ARM}/providers/Microsoft.Billing/billingAccounts?api-version=2024-04-01", az.token(ARM))
    if not r.ok:
        return fail(r)
    accts = [f"{a['properties'].get('displayName')} = {a['properties'].get('agreementType')}" for a in r.json()["value"]]
    return f"OK  {accts or 'no billing account visible (needs a billing role)'}"


def cost_by_meter(tok: str, sub_id: str):
    body = {"type": "ActualCost", "timeframe": "MonthToDate",
            "dataset": {"granularity": "None", "aggregation": {"cost": {"name": "Cost", "function": "Sum"}},
                        "grouping": [{"type": "Dimension", "name": "MeterCategory"},
                                     {"type": "Dimension", "name": "Meter"}]}}
    r = call("POST", f"{ARM}/subscriptions/{sub_id}/providers/Microsoft.CostManagement/query?api-version=2023-11-01",
             tok, json=body)
    if not r.ok:
        return None, fail(r)
    p = r.json()["properties"]
    cols = [c["name"] for c in p["columns"]]
    return [dict(zip(cols, row)) for row in p["rows"]], None


def probe_cost_mgmt(az: AzSignIn):
    tok = az.token(ARM)
    r = call("GET", f"{ARM}/subscriptions?api-version=2022-12-01", tok)
    if not r.ok:
        return fail(r)
    subs = r.json()["value"]
    lines, denied = [f"OK  {len(subs)} subscriptions visible"], 0
    for s in subs:
        rows, err = cost_by_meter(tok, s["subscriptionId"])
        if err:
            denied += 1
            continue
        total = sum(x["Cost"] for x in rows)
        cur = rows[0].get("Currency", "") if rows else ""
        ai = [x for x in rows if AI_METERS.search(f"{x['MeterCategory']} {x['Meter']}")]
        lines.append(f"    {s['displayName']}: MTD {total:,.2f} {cur}")
        lines += [f"      AI meter  {x['MeterCategory']} / {x['Meter']}: {x['Cost']:,.2f}" for x in ai]
    if denied:
        lines.append(f"    no cost read on {denied} subscription(s) (needs Cost Management Reader)")
    return "\n".join(lines)


def probe_power_platform(az: AzSignIn):
    tok = az.token(PPAPI)
    r = call("GET", f"{PPAPI}/licensing/entitlements/MCSMessages?api-version=2024-10-01", tok)
    if not r.ok:
        return fail(r)
    cap = {k: v for k, v in r.json().items() if not isinstance(v, (dict, list))}
    day = (dt.date.today() - dt.timedelta(days=2)).isoformat()
    r2 = call("GET", f"{PPAPI}/licensing/entitlements/MCSMessages/resources?fromDate={day}&toDate={day}"
                     f"&includeFields=users%2Ctags%2CasOfDate&pageSize=5000&api-version=2024-10-01", tok)
    if not r2.ok:
        return f"OK  capacity {cap}\n    per-agent: {fail(r2)}"
    rows = r2.json().get("value", [])
    agents = {x.get("resourceId") or x.get("resourceName") for x in rows}
    sample = sorted(rows[0].keys()) if rows else []
    return f"OK  capacity {cap}\n    per-agent {day}: {len(rows)} rows, {len(agents)} agents; fields {sample}"


def probe_skus(az: AzSignIn):
    r = call("GET", f"{G}/v1.0/subscribedSkus", az.token(G))
    if not r.ok:
        return fail(r)
    rows = sorted(r.json()["value"], key=lambda s: -s["prepaidUnits"]["enabled"])
    return "OK  " + f"{len(rows)} SKUs\n" + "\n".join(
        f"    {s['skuPartNumber']}: {s['consumedUnits']} assigned / {s['prepaidUnits']['enabled']} bought"
        for s in rows if s["prepaidUnits"]["enabled"])


def main():
    az = AzSignIn()
    checks = [("1 Graph Copilot Cost Mgmt", probe_graph_cost), ("2 Azure billing accounts", lambda: probe_billing(az)),
              ("3 Azure Cost Management", lambda: probe_cost_mgmt(az)),
              ("4 Power Platform agents", lambda: probe_power_platform(az)), ("5 Graph licences", lambda: probe_skus(az))]
    for name, fn in checks:
        try:
            out = fn()
        except (RuntimeError, requests.RequestException, KeyError, SystemExit) as e:
            out = f"FAIL {type(e).__name__}: {str(e)[:220]}"
        print(f"\n== {name}\n{out}", flush=True)


if __name__ == "__main__":
    sys.exit(main())
