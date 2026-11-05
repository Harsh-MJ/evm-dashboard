"""Builds Benchmark_Workbook.xlsx: plain-Excel formula benchmark for Month 2 and Month 4 (no macros, no Python)."""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

import evm_engine as E

p = E.load_excel("sample_project.xlsx")
wb = Workbook()
B = Font(bold=True); BLUE = Font(color="0000FF"); HDR = PatternFill("solid", fgColor="D9D9D9")
INR = "#,##,##0"; IDX = "0.000"
thin = Side(style="thin", color="999999"); box = Border(top=thin, bottom=thin, left=thin, right=thin)

# ---------------- Inputs
ws = wb.active; ws.title = "Inputs"
ws["A1"] = "Benchmark inputs - copied from sample_project.xlsx (blue = input)"; ws["A1"].font = B
hdr = ["ID", "Activity", "Budget", "Plan% M1", "M2", "M3", "M4", "Act% M1", "M2", "M3", "M4", "Cost M1", "M2", "M3", "M4"]
for j, h in enumerate(hdr, 1):
    c = ws.cell(3, j, h); c.font = B; c.fill = HDR
plan = p.planned.ffill(axis=1).fillna(0)
apct = p.actual_pct.ffill(axis=1).fillna(0)
cost = p.actual_cost.fillna(0)
for i, aid in enumerate(p.activities.index, 4):
    ws.cell(i, 1, aid); ws.cell(i, 2, p.activities.loc[aid, "Activity"])
    vals = [p.activities.loc[aid, "Budget"]] + list(plan.loc[aid][:4]) + list(apct.loc[aid][:4]) + list(cost.loc[aid][:4])
    for j, v in enumerate(vals, 3):
        c = ws.cell(i, j, float(v)); c.font = BLUE
        c.number_format = INR if j == 3 or j >= 12 else "0"
ws.cell(9, 2, "BAC (total)").font = B
ws["C9"] = "=SUM(C4:C8)"; ws["C9"].number_format = INR; ws["C9"].font = B
ws["A11"] = "Planned cumulative PV by month (for Earned Schedule)"; ws["A11"].font = B
for m in range(0, 7):
    ws.cell(12, 2 + m, f"M{m}").font = B
# Planned % for Months 5 and 6 (needed only for the full PV curve) typed in columns Q and R,
# so PV5 and PV6 are formula-derived like PV1-PV4 (no value is copied from the tool).
for j, h in [(17, "Plan% M5"), (18, "Plan% M6")]:
    c = ws.cell(3, j, h); c.font = B; c.fill = HDR
for i, aid in enumerate(p.activities.index, 4):
    for j, mcol in [(17, 4), (18, 5)]:
        c = ws.cell(i, j, float(plan.loc[aid].iloc[mcol])); c.font = BLUE; c.number_format = "0"
for m in range(0, 7):
    if m == 0:
        ws.cell(13, 2, 0).font = BLUE
    else:
        col = chr(ord("D") + m - 1) if m <= 4 else {5: "Q", 6: "R"}[m]
        ws.cell(13, 2 + m, f"=SUMPRODUCT($C$4:$C$8,{col}4:{col}8)/100")
    ws.cell(13, 2 + m).number_format = INR
ws["A13"] = "PV"
ws["A14"] = "All PV values are formulas: SUMPRODUCT(budget, planned cumulative %)/100. Months 5-6 use the planned % in columns Q-R."
for col, w in zip("ABCDEFGHIJKLMNOPQR", [6, 26, 12] + [9] * 8 + [11] * 4 + [3, 9, 9]):
    ws.column_dimensions[col].width = w


def period_sheet(month: int):
    s = wb.create_sheet(f"Month {month}")
    s["A1"] = f"Hand benchmark - status at end of Month {month}"; s["A1"].font = Font(bold=True, size=13)
    pc = chr(ord("D") + month - 1)        # planned % col on Inputs
    ac_ = chr(ord("H") + month - 1)       # actual % col
    lastc = chr(ord("L") + month - 1)     # last cost col
    h = ["ID", "Activity", "Budget", "PV = Budget x Plan%", "EV = Budget x Act%", "AC = sum of costs", "CV = EV-AC", "SV = EV-PV", "CPI = EV/AC", "SPI = EV/PV"]
    for j, t in enumerate(h, 1):
        c = s.cell(3, j, t); c.font = B; c.fill = HDR; c.alignment = Alignment(wrap_text=True)
    for i in range(5):
        r, ir = 4 + i, 4 + i
        s.cell(r, 1, f"=Inputs!A{ir}"); s.cell(r, 2, f"=Inputs!B{ir}"); s.cell(r, 3, f"=Inputs!C{ir}")
        s.cell(r, 4, f"=Inputs!C{ir}*Inputs!{pc}{ir}/100")
        s.cell(r, 5, f"=Inputs!C{ir}*Inputs!{ac_}{ir}/100")
        s.cell(r, 6, f"=SUM(Inputs!L{ir}:{lastc}{ir})")
        s.cell(r, 7, f"=E{r}-F{r}"); s.cell(r, 8, f"=E{r}-D{r}")
        s.cell(r, 9, f'=IF(F{r}=0,"-",E{r}/F{r})'); s.cell(r, 10, f'=IF(D{r}=0,"-",E{r}/D{r})')
    s.cell(9, 2, "Project total").font = B
    for j, col in enumerate("CDEFGH", 3):
        s.cell(9, j, f"=SUM({col}4:{col}8)").font = B
    s["I9"] = "=E9/F9"; s["J9"] = "=E9/D9"
    for r in range(4, 10):
        for j in range(1, 11):
            s.cell(r, j).border = box
            s.cell(r, j).number_format = IDX if j >= 9 else INR
    # project-level measures
    s["A12"] = "Project-level measures"; s["A12"].font = B
    rows = [
        ("BAC", "=C9", "sum of activity budgets"),
        ("PV", "=D9", ""), ("EV", "=E9", ""), ("AC", "=F9", ""),
        ("CV = EV - AC", "=C15-C16", ""), ("SV = EV - PV", "=C15-C14", ""),
        ("CPI = EV / AC", "=C15/C16", ""), ("SPI = EV / PV", "=C15/C14", ""),
        ("EAC1 = BAC / CPI", "=C13/C19", "typical variance"),
        ("EAC2 = AC + (BAC - EV)", "=C16+(C13-C15)", "atypical variance"),
        ("EAC3 = AC + (BAC - EV)/(CPI x SPI)", "=C16+(C13-C15)/(C19*C20)", "cost and schedule"),
        ("ETC = EAC1 - AC", "=C21-C16", ""), ("VAC = BAC - EAC1", "=C13-C21", ""),
        ("TCPI = (BAC - EV)/(BAC - AC)", "=(C13-C15)/(C13-C16)", "CPI needed on remaining work"),
        ("TCPI - CPI", "=C26-C19", "> 0.10 => budget unlikely to be met"),
        ("Earned Schedule C (last whole month PV<=EV)", f"=MATCH(C15,Inputs!B13:H13,1)-1", "months"),
        ("ES = C + (EV-PV_C)/(PV_C+1 - PV_C)",
         "=C28+(C15-INDEX(Inputs!B13:H13,C28+1))/(INDEX(Inputs!B13:H13,C28+2)-INDEX(Inputs!B13:H13,C28+1))", "months"),
        ("SV(t) = ES - AT", f"=C29-{month}", "months"),
        ("SPI(t) = ES / AT", f"=C29/{month}", ""),
        ("IEAC(t) = PD / SPI(t)", "=6/C31", "months (PD = 6)"),
    ]
    for k, (lab, f, note) in enumerate(rows, 13):
        s.cell(k, 1, lab); s.cell(k, 3, f); s.cell(k, 4, note).font = Font(italic=True, color="666666")
        s.cell(k, 3).number_format = INR if k in (13, 14, 15, 16, 17, 18, 21, 22, 23, 24, 25) else IDX
        s.cell(k, 3).border = box
    s["A35"] = "Hand check (fill in pen-and-paper / calculator values; Diff must be within Rs 1 or 0.001)"; s["A35"].font = B
    s["A36"] = "Measure"; s["C36"] = "Workbook"; s["D36"] = "Hand value"; s["E36"] = "Diff"; s["F36"] = "Checked by / date"
    for j in range(1, 7):
        s.cell(36, j).font = B; s.cell(36, j).fill = HDR
    for n, k in enumerate([14, 15, 16, 19, 20, 21, 26, 29]):
        r = 37 + n
        s.cell(r, 1, f"=A{k}"); s.cell(r, 3, f"=C{k}"); s.cell(r, 5, f'=IF(D{r}="","",D{r}-C{r})')
        s.cell(r, 3).number_format = s.cell(k, 3).number_format
    s.column_dimensions["A"].width = 44; s.column_dimensions["B"].width = 24
    for col in "CDEFGHIJ":
        s.column_dimensions[col].width = 15
    s.row_dimensions[3].height = 32


period_sheet(2); period_sheet(4)
wb.save("Benchmark_Workbook.xlsx")
print("ok")
