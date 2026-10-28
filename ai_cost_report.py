"""Cost and activity over time for every Copilot / AI solution: monthly, weekly and daily tables.

    python ai_cost_report.py "<Usage dir>" [--logo logos.png]

Cost: snapshots/credits_<date>.csv (Graph Cost Management, credits month to date per user x service). The API has no
history, so each snapshot pins a real total; the credits between two snapshots are spread over the days in between in
proportion to that user's audit-log activity on the matching solution (evenly if none). Daily snapshots = exact days.
Activity: ai_activity/daily_*.csv (ai_activity.py). Writes ai_activity/AI_Cost_Activity_Report.html and
AI_daily_cost_activity.csv.
"""
import argparse, base64, datetime as dt, html, pathlib
import pandas as pd
from ai_activity import solution

USD_PER_CREDIT = 0.01
SERVICE = {"cowork": "Copilot Cowork", "copilotCode": "Copilot Code", "copilotStudio": "Copilot Studio agents",
           "workIQ": "Work IQ API", "gitHubCopilotApp": "GitHub Copilot", "autopilot": "Copilot Autopilot"}
CSS = """:root{--navy:#002957;--teal:#00a884;--ink:#1d2433;--mute:#5b6577;--line:#dde3ec;--zebra:#f5f8fb}
body{margin:0;background:#fff;color:var(--ink);font:16px/1.55 -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif}
.wrap{max-width:1040px;margin:0 auto;padding:24px 18px 48px}
.label{display:inline-block;font-size:12px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:var(--navy);border:1.5px solid var(--navy);border-radius:4px;padding:1px 8px}
header img{max-width:260px;width:55%;display:block;margin:12px 0 4px}
h1{color:var(--navy);font-size:26px;margin:6px 0 2px}.sub{color:var(--mute);margin:0 0 8px}
h2{color:var(--navy);font-size:20px;margin:32px 0 10px;padding-left:10px;border-left:5px solid var(--teal)}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:10px;margin:14px 0}
.kpi{border:1px solid var(--line);border-radius:10px;padding:10px 12px}.kpi b{display:block;font-size:23px;color:var(--navy)}.kpi span{color:var(--mute);font-size:13px}
.scroll{overflow-x:auto}table{border-collapse:collapse;width:100%;font-size:13.5px;margin:8px 0}
th,td{text-align:left;padding:5px 8px;border-bottom:1px solid var(--line);white-space:nowrap}
th{background:var(--navy);color:#fff;font-weight:600}tr:nth-child(even) td{background:var(--zebra)}td.n,th.n{text-align:right}
.box{border-left:5px solid #c77700;background:#fff4e0;padding:8px 14px;border-radius:0 8px 8px 0;margin:10px 0}
.note{color:var(--mute);font-size:13.5px}
footer{margin-top:34px;padding-top:10px;border-top:1px solid var(--line);color:var(--mute);font-size:13px}
@media print{.wrap{max-width:none;padding:0}table,.kpi,.box{page-break-inside:avoid}h2{page-break-after:avoid}.scroll{overflow:visible}}"""


def table(df: pd.DataFrame) -> str:
    head = "".join(f'<th class="{"n" if df[c].dtype.kind in "if" else ""}">{html.escape(str(c))}</th>' for c in df.columns)
    body = "".join("<tr>" + "".join(
        f'<td class="n">{v:,.0f}</td>' if df[c].dtype.kind in "if" else f"<td>{html.escape(str(v))}</td>"
        for c, v in zip(df.columns, r)) + "</tr>" for r in df.fillna(0).itertuples(index=False))
    return f'<div class="scroll"><table><tr>{head}</tr>{body}</table></div>'


def load_activity(d: pathlib.Path) -> pd.DataFrame:
    files = sorted((d / "ai_activity").glob("daily_*.csv"))
    a = pd.concat([pd.read_csv(f) for f in files], ignore_index=True)
    a["solution"] = a["app"].astype(str).map(solution)  # reclassify with the current rules
    a["date"] = pd.to_datetime(a["date"])
    return a


def daily_cost(d: pathlib.Path, act: pd.DataFrame) -> pd.DataFrame:
    """date, service, user, credits: snapshot deltas spread over activity days."""
    snaps = []
    for f in sorted((d / "snapshots").glob("credits_*.csv")):
        s = pd.read_csv(f)
        snaps.append(pd.DataFrame({"snap": pd.to_datetime(s["snapshotDate"]), "user": s["upn"].str.lower(),
                                   "service": s["serviceId"].map(SERVICE).fillna(s["serviceId"]), "mtd": s["creditsMTD"]}))
    if not snaps:
        return pd.DataFrame(columns=["date", "service", "user", "credits"])
    s = pd.concat(snaps).sort_values("snap")
    s["month"] = s["snap"].dt.to_period("M")
    out = []
    for (user, svc, month), g in s.groupby(["user", "service", "month"]):
        prev_mtd, prev_day = 0.0, month.start_time - pd.Timedelta(days=1)
        for snap, mtd in zip(g["snap"], g["mtd"]):
            delta = mtd - prev_mtd
            days = pd.date_range(prev_day + pd.Timedelta(days=1), snap)
            w = (act[(act["user"] == user) & (act["solution"] == svc) & act["date"].between(days[0], days[-1])]
                 .groupby("date")["interactions"].sum().reindex(days, fill_value=0))
            w = w / w.sum() if w.sum() else pd.Series(1 / len(days), index=days)
            out += [(day, svc, user, delta * share) for day, share in w.items() if share]
            prev_mtd, prev_day = mtd, snap
    return pd.DataFrame(out, columns=["date", "service", "user", "credits"])


def pivot(df, idx, col, val, fmt=None):
    p = df.pivot_table(index=idx, columns=col, values=val, aggfunc="sum", fill_value=0)
    p["Total"] = p.sum(axis=1)
    p = p.reset_index()
    if fmt:
        p[idx] = p[idx].map(fmt)
    return p


def build(d: pathlib.Path, logo: str) -> tuple[str, pd.DataFrame]:
    act = load_activity(d)
    cost = daily_cost(d, act)
    act["month"], cost["month"] = act["date"].dt.to_period("M").astype(str), cost["date"].dt.to_period("M").astype(str)
    wk = lambda s: s.dt.to_period("W-SUN").dt.start_time
    act["week"], cost["week"] = wk(act["date"]), wk(cost["date"])

    m_cost = pivot(cost, "month", "service", "credits")
    m_cost["USD"] = m_cost["Total"] * USD_PER_CREDIT
    w_cost = pivot(cost, "week", "service", "credits", lambda x: f"{x:%d %b %Y}")
    w_cost["USD"] = w_cost["Total"] * USD_PER_CREDIT
    m_act = pivot(act, "month", "solution", "interactions")
    m_users = act.groupby(["month", "solution"])["user"].nunique().unstack(fill_value=0).reset_index()
    w_act = pivot(act, "week", "solution", "interactions", lambda x: f"{x:%d %b %Y}")

    day_cost = cost.groupby(["date", "service"])["credits"].sum().unstack(fill_value=0).add_suffix(" credits")
    day_act = act.groupby(["date", "solution"])["interactions"].sum().unstack(fill_value=0).add_suffix(" interactions")
    daily = day_act.join(day_cost, how="outer").fillna(0).sort_index()
    daily["USD"] = daily.filter(like=" credits").sum(axis=1) * USD_PER_CREDIT
    daily = daily.reset_index().rename(columns={"index": "date"})
    recent = daily[daily["date"] >= daily["date"].max() - pd.Timedelta(days=34)].copy()
    recent["date"] = recent["date"].dt.strftime("%a %d %b")

    snaps = sorted(p.stem[8:] for p in (d / "snapshots").glob("credits_*.csv"))
    total = cost["credits"].sum()
    k = lambda v, t: f'<div class="kpi"><b>{v}</b><span>{t}</span></div>'
    kpis = "".join([k(f"{total:,.0f}", "credits billed (all snapshots)"), k(f"${total * USD_PER_CREDIT:,.0f}", "PAYG list price"),
                    k(f"{act['interactions'].sum():,}", "AI interactions in the audit log"),
                    k(f"{act['user'].nunique():,}", "people using any Copilot / AI"),
                    k(f"{act['date'].min():%d %b} – {act['date'].max():%d %b}", "activity history covered")])
    img = f'<img src="data:image/png;base64,{logo}" alt="Digital Workplace and rAIse">' if logo else ""

    page = f"""<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>AI Cost and Activity</title><style>{CSS}</style></head><body><div class="wrap">
<header><span class="label">Almirall · Internal use</span>{img}
<h1>Copilot and AI — cost and activity over time</h1>
<p class="sub">Cost snapshots: {", ".join(snaps)} · generated {dt.date.today():%d %b %Y}</p></header>
<div class="kpis">{kpis}</div>
<div class="box">Cost per day and per week is <b>estimated</b> between snapshots: each snapshot is an exact total, the credits in
between are spread over the days the person actually used that solution. A daily snapshot makes every day exact.</div>

<h2>1. Cost per month, by service (credits)</h2>{table(m_cost)}
<h2>2. Cost per week, by service (credits, estimated split)</h2>{table(w_cost)}
<h2>3. Activity per month, by AI solution (interactions)</h2>{table(m_act)}
<h3>Active people per month</h3>{table(m_users)}
<h2>4. Activity per week, by AI solution (interactions)</h2>{table(w_act)}
<h2>5. Last 5 weeks, day by day</h2>
<p class="note">Full history in AI_daily_cost_activity.csv next to this file.</p>{table(recent)}

<h2>6. What each solution costs and where its cost comes from</h2>
<table><tr><th>Solution</th><th>How it is paid</th><th>In this report</th></tr>
<tr><td>Copilot Cowork, Copilot Code</td><td>Copilot credits, pay as you go</td><td>Cost and activity</td></tr>
<tr><td>Copilot Studio agents</td><td>Copilot Studio credits, billed through Power Platform (0 in Microsoft 365 credits)</td><td>Activity only; cost needs Power Platform admin or the Azure subscription</td></tr>
<tr><td>Copilot Chat, Copilot in M365 apps, Microsoft prebuilt agents</td><td>Microsoft 365 Copilot licence per seat (Chat is free without one)</td><td>Activity only; seat cost is fixed per licence</td></tr>
<tr><td>Copilot in Fabric / Power BI, Notebooks</td><td>Fabric capacity / licence</td><td>Activity only</td></tr>
<tr><td>GitHub Copilot</td><td>GitHub seats and premium requests</td><td>Not yet: needs a GitHub enterprise token</td></tr>
<tr><td>Azure AI Foundry / OpenAI</td><td>Azure consumption</td><td>Not yet: needs Cost Management Reader</td></tr></table>
<footer>Almirall · Internal use · Digital Workplace (IT) · ai_cost_report.py (copilot-credits-report). Activity: Purview audit log
(CopilotInteraction, 180-day retention). Cost: Graph Copilot Cost Management API at 0.01 USD per credit.</footer>
</div></body></html>"""
    return page, daily


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("usage_dir")
    ap.add_argument("--logo")
    a = ap.parse_args()
    d = pathlib.Path(a.usage_dir)
    logo = base64.b64encode(pathlib.Path(a.logo).read_bytes()).decode() if a.logo else ""
    page, daily = build(d, logo)
    (d / "ai_activity" / "AI_Cost_Activity_Report.html").write_text(page, encoding="utf-8")
    daily.to_csv(d / "ai_activity" / "AI_daily_cost_activity.csv", index=False)
    print(d / "ai_activity" / "AI_Cost_Activity_Report.html")


if __name__ == "__main__":
    main()
