"""Builds the input template sample_project.xlsx (synthetic G+3 block, BAC Rs 1 crore, 6 months, data date Month 4)."""
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

M = [f"M{i}" for i in range(1, 7)]
acts = pd.DataFrame({
    "Activity ID": ["A1", "A2", "A3", "A4", "A5"],
    "Activity": ["Site work and excavation", "Foundations", "RCC frame (G+3)", "Masonry", "Finishes and MEP"],
    "Budget": [800000, 1700000, 3500000, 1500000, 2500000],
})
plan = pd.DataFrame([
    ["A1", 100, None, None, None, None, None],
    ["A2", 40, 100, None, None, None, None],
    ["A3", 0, 20, 55, 100, None, None],
    ["A4", 0, 0, 20, 60, 100, None],
    ["A5", 0, 0, 0, 10, 50, 100],
], columns=["Activity ID"] + M)
apct = pd.DataFrame([
    ["A1", 100, 100, 100, 100, None, None],
    ["A2", 30, 90, 100, 100, None, None],
    ["A3", 0, 10, 40, 80, None, None],
    ["A4", 0, 0, 10, 45, None, None],
    ["A5", 0, 0, 0, 5, None, None],
], columns=["Activity ID"] + M)
acost = pd.DataFrame([
    ["A1", 860000, 0, 0, 0, None, None],
    ["A2", 540000, 1090000, 210000, 0, None, None],
    ["A3", 0, 370000, 1120000, 1460000, None, None],
    ["A4", 0, 0, 160000, 540000, None, None],
    ["A5", 0, 0, 0, 140000, None, None],
], columns=["Activity ID"] + M)
info = pd.DataFrame([
    ["Project name", "Sample G+3 Residential Block - Structure and Finishes Package"],
    ["Currency", "INR (Rs.)"],
    ["Planned duration (months)", 6],
    ["Data source", "Synthetic data prepared by the team for CE725 (no confidential project data)"],
    ["How to use", "Planned_% and Actual_% are CUMULATIVE % complete at month end. Actual_Cost is cost INCURRED IN THAT MONTH. Leave future months blank."],
])

out = "sample_project.xlsx"
with pd.ExcelWriter(out, engine="openpyxl") as w:
    info.to_excel(w, "Project_Info", index=False, header=False)
    acts.to_excel(w, "Activities", index=False)
    plan.to_excel(w, "Planned_%", index=False)
    apct.to_excel(w, "Actual_%", index=False)
    acost.to_excel(w, "Actual_Cost", index=False)

wb = load_workbook(out)
blue = Font(color="0000FF")
for ws in wb.worksheets:
    ws.column_dimensions["A"].width = 26 if ws.title == "Project_Info" else 12
    ws.column_dimensions["B"].width = 90 if ws.title == "Project_Info" else 28
    if ws.title != "Project_Info":
        for c in ws[1]:
            c.font = Font(bold=True)
            c.fill = PatternFill("solid", fgColor="D9D9D9")
        for row in ws.iter_rows(min_row=2, min_col=2):
            for c in row:
                if isinstance(c.value, (int, float)):
                    c.font = blue
                    if ws.title in ("Activities", "Actual_Cost"):
                        c.number_format = "#,##,##0"
    else:
        for r in ws.iter_rows():
            r[0].font = Font(bold=True)
            r[1].alignment = Alignment(wrap_text=True)
wb.save(out)
print("written", out)
