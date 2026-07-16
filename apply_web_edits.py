#!/usr/bin/env python3
"""Pull the repo and apply web edits (rows flagged _edited in data.json) back into Job_Application_Tracker.xlsx.

Web-editable fields: Status + Notes (Applications), Status + Fit Notes (New Jobs).
Rows are matched by hyperlink URL first, then by Company+Role. Run before rebuilding the site.
"""
import json
import os
import subprocess

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
XLSX = os.path.join(HERE, "..", "Job_Application_Tracker.xlsx")

subprocess.run(["git", "-C", HERE, "pull", "--ff-only"], check=False)

with open(os.path.join(HERE, "data.json")) as f:
    data = json.load(f)

SHEETS = {
    # sheet name: (json key, link col header, editable {json field: column header})
    "Applications": ("applications", "Job Link", {"Status": "Status", "Notes": "Notes"}),
    "New Jobs": ("new_jobs", "Job Link", {"Status": "Status", "Fit Notes": "Fit Notes"}),
}

wb = openpyxl.load_workbook(XLSX)
applied = 0

for sheet_name, (key, link_col, editable) in SHEETS.items():
    if sheet_name not in wb.sheetnames:
        continue
    ws = wb[sheet_name]
    headers = {c.value: i for i, c in enumerate(ws[1], start=1)}
    edited = [r for r in data.get(key, []) if r.get("_edited")]
    if not edited:
        continue
    for rec in edited:
        for row in ws.iter_rows(min_row=2):
            link_cell = row[headers[link_col] - 1] if link_col in headers else None
            url = link_cell.hyperlink.target if (link_cell and link_cell.hyperlink) else None
            same_url = rec.get("_url") and url == rec["_url"]
            same_key = (
                str(row[headers["Company"] - 1].value) == str(rec.get("Company"))
                and str(row[headers["Role"] - 1].value) == str(rec.get("Role"))
            )
            if same_url or same_key:
                for field, col in editable.items():
                    if col in headers and rec.get(field) is not None:
                        ws.cell(row=row[0].row, column=headers[col], value=str(rec[field]).strip())
                applied += 1
                break

if applied:
    wb.save(XLSX)
print(f"applied {applied} web-edited row(s) to xlsx")
