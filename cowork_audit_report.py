"""HTML report of Cowork activity (Purview audit, from cowork_audit.py) joined with credits and CR.

    python cowork_audit_report.py "<Usage dir>" [--logo logos.png]

Reads <usage-dir>/cowork_audit/cowork_events.csv, the newest M365AIConsumptionByUser*.csv (credits month to date)
and the CR sources build_report.py uses. Writes <usage-dir>/cowork_audit/Cowork_Activity_Report.html.
"""
import argparse, base64, collections, datetime as dt, html, pathlib
import pandas as pd
from build_report import assign_cr, load_cr

USD_PER_CREDIT = 0.01
LABELS = {"533616b6-00a5-4cd1-b577-93208fa93eb1": "Internal Use"}  # others: ask Security for the label list

CSS = """:root{--navy:#002957;--teal:#00a884;--ink:#1d2433;--mute:#5b6577;--line:#dde3ec;--zebra:#f5f8fb}
body{margin:0;background:#fff;color:var(--ink);font:16px/1.55 -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif}
.wrap{max-width:1000px;margin:0 auto;padding:24px 18px 48px}
.label{display:inline-block;font-size:12px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:var(--navy);border:1.5px solid var(--navy);border-radius:4px;padding:1px 8px}
header img{max-width:260px;width:55%;display:block;margin:12px 0 4px}
h1{color:var(--navy);font-size:26px;margin:6px 0 2px} .sub{color:var(--mute);margin:0 0 8px}
h2{color:var(--navy);font-size:20px;margin:32px 0 10px;padding-left:10px;border-left:5px solid var(--teal)}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px;margin:14px 0}
.kpi{border:1px solid var(--line);border-radius:10px;padding:10px 12px}.kpi b{display:block;font-size:24px;color:var(--navy)}.kpi span{color:var(--mute);font-size:13px}
table{border-collapse:collapse;width:100%;font-size:14px;margin:8px 0}th,td{text-align:left;padding:6px 8px;border-bottom:1px solid var(--line);vertical-align:top}
th{background:var(--navy);color:#fff;font-weight:600}tr:nth-child(even) td{background:var(--zebra)}td.n,th.n{text-align:right;white-space:nowrap}
.bar{height:10px;background:var(--teal);border-radius:3px;display:inline-block;vertical-align:middle}
.box{border-left:5px solid #c77700;background:#fff4e0;padding:8px 14px;border-radius:0 8px 8px 0;margin:10px 0}
.note{color:var(--mute);font-size:13.5px} code{background:#eef2f7;padding:1px 5px;border-radius:4px;font-size:13px}
footer{margin-top:34px;padding-top:10px;border-top:1px solid var(--line);color:var(--mute);font-size:13px}
@media print{.wrap{max-width:none;padding:0}table,.kpi,.box{page-break-inside:avoid}h2{page-break-after:avoid}}"""


def esc(x):
    return html.escape(str(x))


def table(df: pd.DataFrame, num=(), bar=None) -> str:
    top = df[bar].max() if bar and len(df) else 0
    head = "".join(f'<th class="{"n" if c in num else ""}">{esc(c)}</th>' for c in df.columns)
    rows = []
    for _, r in df.iterrows():
        cells = []
        for c in df.columns:
            v = r[c]
            txt = f"{v:,.0f}" if c in num and isinstance(v, (int, float)) and pd.notna(v) else ("" if pd.isna(v) else esc(v))
            if c == bar and top:
                txt += f' <span class="bar" style="width:{max(2, 120 * v / top):.0f}px"></span>'
            cells.append(f'<td class="{"n" if c in num else ""}">{txt}</td>')
        rows.append("<tr>" + "".join(cells) + "</tr>")
    return f"<table><tr>{head}</tr>{''.join(rows)}</table>"


def counts(series: pd.Series, n=15) -> collections.Counter:
    return collections.Counter(x for v in series.dropna() for x in str(v).split(";") if x).most_common(n)


def credits(d: pathlib.Path) -> tuple[pd.DataFrame, str]:
    files = [f for f in d.glob("M365AIConsumptionByUser*.csv") if "history" not in f.name]
    if not files:
        return pd.DataFrame(columns=["upn", "credits"]), "none"
    f = max(files, key=lambda p: p.stat().st_mtime)
    x = pd.read_csv(f)
    return pd.DataFrame({"upn": x["User Principal Name"].str.lower(), "credits": x["Monthly credits used"]}), f.name


def build(d: pathlib.Path, logo: str) -> str:
    e = pd.read_csv(d / "cowork_audit" / "cowork_events.csv")
    e["user"] = e["user"].str.lower()
    e["t"] = pd.to_datetime(e["time_utc"], utc=True)
    e["day"], e["week"] = e["t"].dt.date, (e["t"].dt.tz_localize(None).dt.to_period("W-SUN").dt.start_time).dt.date
    month = e["t"].max().strftime("%Y-%m")

    cred, cred_src = credits(d)
    users = e.groupby("user").agg(interactions=("record_id", "size"), tasks=("thread_id", "nunique"),
                                  active_days=("day", "nunique"), items_read=("items_read", "sum"),
                                  claude=("provider", lambda s: (s == "Anthropic").sum()),
                                  last=("t", "max")).reset_index().rename(columns={"user": "upn"})
    users = users.merge(cred, on="upn", how="left")
    users["month"] = month
    users["CR"] = assign_cr(users, load_cr(d)).replace("UNMAPPED", "No CR")
    users["credits_per_task"] = users["credits"] / users["tasks"]
    users["usd"] = users["credits"] * USD_PER_CREDIT

    known = e[e["provider"].fillna("") != ""]
    by_model = (e.assign(model=e["model"].fillna("not recorded").replace("", "not recorded"),
                         provider=e["provider"].fillna("").replace("", "—"))
                .groupby(["provider", "model"]).size().reset_index(name="interactions")
                .sort_values("interactions", ascending=False))
    claude_share = (known["provider"] == "Anthropic").mean() * 100 if len(known) else 0

    weekly = e.groupby("week").agg(interactions=("record_id", "size"), users=("user", "nunique"),
                                   tasks=("thread_id", "nunique")).reset_index()
    weekly["week"] = weekly["week"].astype(str)

    cr = users.groupby("CR").agg(users=("upn", "size"), tasks=("tasks", "sum"), interactions=("interactions", "sum"),
                                 credits=("credits", "sum"), usd=("usd", "sum")).reset_index().sort_values("credits", ascending=False)

    ut = users.sort_values("credits", ascending=False)[
        ["upn", "CR", "tasks", "interactions", "active_days", "items_read", "claude", "credits", "credits_per_task", "usd"]]
    ut.columns = ["User", "CR", "Tasks", "Interactions", "Active days", "Items read", "Claude turns", "Credits MTD",
                  "Credits / task", "USD"]

    tools = pd.DataFrame(counts(e["tools"], 20), columns=["Tool", "Interactions"])
    items = pd.DataFrame(counts(e["item_types"], 12), columns=["Item type", "Interactions"])
    labels = pd.DataFrame([(LABELS.get(k, k), v) for k, v in counts(e["labels"], 12)], columns=["Sensitivity label", "Interactions"])
    req, grant = dict(counts(e["approvals_requested"], 200)), dict(counts(e["approvals_granted"], 200))
    appr = pd.DataFrame([(k, v, grant.get(k, 0)) for k, v in sorted(req.items(), key=lambda kv: -kv[1])[:15]],
                        columns=["Action needing approval", "Requested", "Granted once"])
    labelled = (e["labels"].fillna("") != "").sum()

    k = lambda v, t: f'<div class="kpi"><b>{v}</b><span>{t}</span></div>'
    kpis = "".join([k(f"{len(e):,}", "interactions"), k(e["user"].nunique(), "users"), k(f"{e['thread_id'].nunique():,}", "tasks (threads)"),
                    k(f"{users['credits'].sum():,.0f}", f"credits MTD ({users['credits'].notna().sum()} users matched)"),
                    k(f"${users['usd'].sum():,.0f}", "PAYG list price"), k(f"{claude_share:.0f}%", "of model-tagged turns on Claude")])
    period = f"{e['t'].min():%d %b} – {e['t'].max():%d %b %Y}"
    img = f'<img src="data:image/png;base64,{logo}" alt="Digital Workplace and rAIse">' if logo else ""

    return f"""<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Cowork Activity Report</title><style>{CSS}</style></head><body><div class="wrap">
<header><span class="label">Almirall · Internal use · Restricted: named per-user data</span>{img}
<h1>Copilot Cowork — activity and cost</h1>
<p class="sub">Purview audit log, {period} (UTC) · credits from {esc(cred_src)} · generated {dt.date.today():%d %b %Y}</p></header>
<div class="kpis">{kpis}</div>

<h2>1. Which models ran Cowork</h2>
<p>Each audit record carries the model that served the turn. {len(e) - len(known):,} records have no model recorded.</p>
{table(by_model.rename(columns={"provider": "Provider", "model": "Model", "interactions": "Interactions"}), num=("Interactions",), bar="Interactions")}
<div class="box"><b>Claude (Anthropic)</b> is processed outside the EU Data Boundary. Automatic model selection does route Cowork work to Claude: {claude_share:.0f}% of turns with a model recorded.</div>

<h2>2. Weekly activity</h2>
{table(weekly.rename(columns={"week": "Week starting", "interactions": "Interactions", "users": "Users", "tasks": "Tasks"}), num=("Interactions", "Users", "Tasks"), bar="Interactions")}

<h2>3. By CR</h2>
{table(cr.rename(columns={"users": "Users", "tasks": "Tasks", "interactions": "Interactions", "credits": "Credits MTD", "usd": "USD"}), num=("Users", "Tasks", "Interactions", "Credits MTD", "USD"), bar="Credits MTD")}

<h2>4. By user</h2>
<p class="note">Credits are month to date for all Copilot credit services (mostly Cowork). Credits / task = credits ÷ distinct Cowork threads: an estimate, the audit log has no per-task cost.</p>
{table(ut, num=("Tasks", "Interactions", "Active days", "Items read", "Claude turns", "Credits MTD", "Credits / task", "USD"), bar="Credits MTD")}

<h2>5. What Cowork did</h2>
<h3>Tools used</h3>{table(tools, num=("Interactions",), bar="Interactions")}
<h3>Content it read</h3>{table(items, num=("Interactions",), bar="Interactions")}
<h3>Sensitivity labels on content read</h3>
<p>{labelled:,} of {len(e):,} interactions touched labelled content.</p>{table(labels, num=("Interactions",), bar="Interactions")}
<h3>Actions that asked the user for approval</h3>{table(appr, num=("Requested", "Granted once"))}

<h2>6. Sources and limits</h2>
<ul>
<li><b>Activity:</b> Purview unified audit log, record type <code>CopilotInteraction</code>, <code>AppIdentity = Copilot.M365Copilot.CoworkChat</code>, read through the Graph Audit Log Query API.</li>
<li><b>Not in the audit log:</b> prompt and response text (needs eDiscovery), credits per task.</li>
<li><b>Credits:</b> Microsoft 365 admin center consumption export (month to date, per user). Daily per-user × service balances come from the Graph Cost Management API in the same toolset.</li>
<li><b>CR:</b> CR Mapping list, overrides, then CoworkUserDetails.</li>
</ul>
<footer>Almirall · Internal use · Digital Workplace (IT) · Generated by cowork_audit_report.py (copilot-credits-report).</footer>
</div></body></html>"""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("usage_dir")
    ap.add_argument("--logo")
    a = ap.parse_args()
    d = pathlib.Path(a.usage_dir)
    logo = base64.b64encode(pathlib.Path(a.logo).read_bytes()).decode() if a.logo else ""
    out = d / "cowork_audit" / "Cowork_Activity_Report.html"
    out.write_text(build(d, logo), encoding="utf-8")
    print(out)


if __name__ == "__main__":
    main()
