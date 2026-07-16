#!/usr/bin/env python3
"""Generate index.html for the job tracker GitHub Pages site from Job_Application_Tracker.xlsx."""
import datetime
import html
import json
import os

import openpyxl

XLSX = os.path.join(os.path.dirname(__file__), "..", "Job_Application_Tracker.xlsx")
OUT = os.path.join(os.path.dirname(__file__), "index.html")

def sheet_rows(ws):
    rows = []
    headers = [c.value for c in ws[1]]
    for row in ws.iter_rows(min_row=2):
        if all(c.value in (None, "") for c in row):
            continue
        rec = {}
        for h, c in zip(headers, row):
            val = c.value
            if isinstance(val, (datetime.date, datetime.datetime)):
                val = val.strftime("%Y-%m-%d")
            rec[h] = val
            if h in ("Job URL", "Job Link") and c.hyperlink:
                rec["_url"] = c.hyperlink.target
        # skip the placeholder example row
        if "example row" in str(rec.get("Role", "")).lower():
            continue
        rows.append(rec)
    return rows

wb = openpyxl.load_workbook(XLSX)
apps = sheet_rows(wb["Applications"])
new_jobs = sheet_rows(wb["New Jobs"]) if "New Jobs" in wb.sheetnames else []

data = {
    "generated": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
    "applications": apps,
    "new_jobs": new_jobs,
}

page = """<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Jestin Roy — Job Tracker</title>
<style>
  :root {
    --bg: #f6f7f9; --card: #ffffff; --text: #1a1d21; --muted: #667085;
    --border: #e4e7ec; --accent: #1f4e78; --link: #0563c1;
  }
  @media (prefers-color-scheme: dark) {
    :root {
      --bg: #101418; --card: #1a2027; --text: #e6e9ec; --muted: #98a2b3;
      --border: #2c343d; --accent: #6ea8dc; --link: #7ab8f5;
    }
  }
  * { box-sizing: border-box; margin: 0; }
  body { background: var(--bg); color: var(--text); font: 15px/1.55 -apple-system, "Segoe UI", Roboto, Arial, sans-serif; padding: 2rem 1rem 4rem; }
  .wrap { max-width: 1100px; margin: 0 auto; }
  h1 { font-size: 1.6rem; margin-bottom: .25rem; }
  .sub { color: var(--muted); margin-bottom: 1.5rem; font-size: .9rem; }
  .counts { display: flex; gap: .6rem; flex-wrap: wrap; margin-bottom: 2rem; }
  .counts span { background: var(--card); border: 1px solid var(--border); border-radius: 999px; padding: .3rem .85rem; font-size: .85rem; color: var(--muted); }
  .counts b { color: var(--text); }
  h2 { font-size: 1.15rem; margin: 2rem 0 .75rem; }
  .tablewrap { overflow-x: auto; background: var(--card); border: 1px solid var(--border); border-radius: 12px; }
  table { border-collapse: collapse; width: 100%; min-width: 760px; }
  th { text-align: left; font-size: .75rem; text-transform: uppercase; letter-spacing: .05em; color: var(--muted); padding: .7rem .9rem; border-bottom: 1px solid var(--border); }
  td { padding: .7rem .9rem; border-bottom: 1px solid var(--border); vertical-align: top; }
  tr:last-child td { border-bottom: none; }
  a { color: var(--link); text-decoration: none; }
  a:hover { text-decoration: underline; }
  .badge { display: inline-block; padding: .15rem .6rem; border-radius: 999px; font-size: .78rem; font-weight: 600; white-space: nowrap; }
  .s-applied, .s-submitted { background: #ddebf7; color: #1f4e78; }
  .s-in-progress { background: #fce4d6; color: #8a4b12; }
  .s-interview { background: #fff2cc; color: #7a6000; }
  .s-offer { background: #c6efce; color: #1d6b30; }
  .s-rejected, .s-not-proceeding { background: #e2e4e8; color: #555c66; }
  .s-saved { background: #e8e0f7; color: #4b3a80; }
  @media (prefers-color-scheme: dark) {
    .s-applied, .s-submitted { background: #1d3a57; color: #a8cdef; }
    .s-in-progress { background: #4a2f1a; color: #f0b98a; }
    .s-interview { background: #4a4213; color: #e8d67a; }
    .s-offer { background: #1d4426; color: #93d9a4; }
    .s-rejected, .s-not-proceeding { background: #333940; color: #9aa3ad; }
    .s-saved { background: #322a4d; color: #c3b3f0; }
  }
  .muted { color: var(--muted); }
  .notes { font-size: .85rem; color: var(--muted); max-width: 340px; }
  footer { margin-top: 2.5rem; color: var(--muted); font-size: .8rem; }
</style>
</head>
<body>
<div class="wrap">
  <h1>Job Application Tracker</h1>
  <div class="sub">Jestin Roy · auto-updated from Job_Application_Tracker.xlsx · last build __GENERATED__</div>
  <div class="counts" id="counts"></div>
  <h2>Applications</h2>
  <div class="tablewrap"><table id="apps"></table></div>
  <h2>New jobs found (last scans)</h2>
  <div class="tablewrap"><table id="newjobs"></table></div>
  <footer>Built with a scheduled Claude task · twice daily scans (12am / 12pm AEST)</footer>
</div>
<script>
const DATA = __DATA__;

const badge = s => {
  if (!s) return "";
  const cls = "s-" + s.toLowerCase().replace(/\\s+/g, "-");
  return `<span class="badge ${cls}">${s}</span>`;
};
const linkCell = r => {
  const url = r._url || r["Job URL"] || null;
  return url ? `<a href="${url}" target="_blank" rel="noopener">${new URL(url).hostname.replace("www.","")} ↗</a>` : "";
};
const esc = v => (v == null ? "" : String(v).replace(/&/g,"&amp;").replace(/</g,"&lt;"));

function render(id, rows, cols) {
  const t = document.getElementById(id);
  if (!rows.length) { t.outerHTML = '<p class="muted" style="padding:1rem">Nothing here yet.</p>'; return; }
  t.innerHTML = "<thead><tr>" + cols.map(c => `<th>${c.label}</th>`).join("") + "</tr></thead>" +
    "<tbody>" + rows.map(r => "<tr>" + cols.map(c => `<td>${c.fn(r)}</td>`).join("") + "</tr>").join("") + "</tbody>";
}

render("apps", DATA.applications, [
  {label:"Company", fn:r=>`<b>${esc(r.Company)}</b>`},
  {label:"Role", fn:r=>esc(r.Role)},
  {label:"Link", fn:linkCell},
  {label:"Applied", fn:r=>esc(r["Date Applied"])},
  {label:"Location", fn:r=>esc(r.Location)},
  {label:"Status", fn:r=>badge(r.Status)},
  {label:"Notes", fn:r=>`<span class="notes">${esc(r.Notes)}</span>`},
]);

render("newjobs", DATA.new_jobs, [
  {label:"Found", fn:r=>esc(r["Date Found"])},
  {label:"Company", fn:r=>`<b>${esc(r.Company)}</b>`},
  {label:"Role", fn:r=>esc(r.Role)},
  {label:"Location", fn:r=>esc(r.Location)},
  {label:"Link", fn:linkCell},
  {label:"Posted", fn:r=>esc(r.Posted)},
  {label:"Fit", fn:r=>`<span class="notes">${esc(r["Fit Notes"])}</span>`},
  {label:"Status", fn:r=>badge(r.Status)},
]);

const apps = DATA.applications;
const active = apps.filter(r => !["Rejected","Not proceeding"].includes(r.Status)).length;
document.getElementById("counts").innerHTML = [
  `<span><b>${apps.length}</b> applications</span>`,
  `<span><b>${active}</b> active</span>`,
  `<span><b>${apps.filter(r=>r.Status==="Interview").length}</b> interviews</span>`,
  `<span><b>${DATA.new_jobs.length}</b> new jobs queued</span>`,
].join("");
</script>
</body>
</html>
"""

page = page.replace("__DATA__", json.dumps(data)).replace("__GENERATED__", data["generated"])
with open(OUT, "w") as f:
    f.write(page)
print("wrote", OUT)
