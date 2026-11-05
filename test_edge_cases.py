"""Edge-case tests for the EVM engine (Validation Report, Table V-4). Run: python test_edge_cases.py
E11-E14 were added after the second review (Validation Report D7-D10)."""
import math
import evm_engine as E

def fresh():
    return E.load_excel("sample_project.xlsx")

FAILED = []

def case(name, fn):
    try:
        res = fn(); print(f"PASS  {name}: {res}")
    except Exception as e:
        FAILED.append(name)
        print(f"FAIL  {name}: {type(e).__name__}: {e}")

def t1():  # Month 1: activities not started -> NaN indices, not crash
    a = E.activity_table(fresh(), 1)
    assert math.isnan(a.loc["A3", "CPI"]) and math.isnan(a.loc["A5", "SPI"])
    return "unstarted activities give '-' (NaN) not divide-by-zero"

def t2():  # project fully complete on budget -> CPI=SPI=1, EAC=BAC
    p = fresh()
    for m in ["M5", "M6"]:
        p.actual_pct[m] = p.actual_pct["M4"]
    p.actual_pct.loc[:, "M6"] = 100
    tot = p.activities["Budget"]
    p.actual_cost.loc[:, :] = 0; p.actual_cost["M6"] = tot
    m = E.project_metrics(p, 6)
    assert abs(m["CPI"] - 1) < 1e-12 and abs(m["SPI"] - 1) < 1e-12 and abs(m["EAC (BAC/CPI)"] - m["BAC"]) < 1e-6
    assert abs(m["ES (months)"] - 6) < 1e-12
    return f"CPI={m['CPI']}, SPI={m['SPI']}, EAC=BAC, ES=6"

def t3():  # percent > 100 flagged
    p = fresh(); p.actual_pct.loc["A3", "M4"] = 120
    msgs = E.validate(p); assert any("outside 0-100" in x for x in msgs); return msgs[0]

def t4():  # decreasing cumulative % flagged
    p = fresh(); p.actual_pct.loc["A3", "M4"] = 30
    msgs = E.validate(p); assert any("decreases" in x for x in msgs); return msgs[0]

def t5():  # missing activity in a sheet flagged
    p = fresh(); p.actual_cost = p.actual_cost.drop("A5")
    msgs = E.validate(p); assert any("missing" in x for x in msgs); return msgs[0]

def t6():  # zero / negative budget flagged
    p = fresh(); p.activities.loc["A1", "Budget"] = 0
    msgs = E.validate(p); assert any("positive" in x for x in msgs); return msgs[0]

def t7():  # under budget but behind schedule -> commentary says under budget & behind
    p = fresh(); p.actual_cost = p.actual_cost * 0.8
    c = " ".join(E.commentary(p, 4)); assert "under budget" in c and "behind plan" in c
    return "commentary direction correct for CPI>1, SPI<1"

def t8():  # EV beyond total PV (ahead of whole plan) -> ES capped at planned duration
    p = fresh(); p.actual_pct.loc[:, "M4"] = 100
    m = E.project_metrics(p, 4); assert abs(m["ES (months)"] - 6) < 1e-12
    return f"ES capped at {m['ES (months)']} months"

def t9():  # no actuals at all -> error, not crash
    p = fresh(); p.actual_pct.loc[:, :] = float("nan")
    msgs = E.validate(p); assert any("no actual progress" in x for x in msgs); return msgs[-1]

def t10():  # 3-dp rounding of SPI = 0.8375 exactly
    s = f"{E.project_metrics(fresh(), 2)['SPI']:.3f}"
    assert s == "0.838", s
    return f"0.8375 displayed as {s}"

def t11():  # no cost booked yet (AC = 0) -> commentary must not crash, CPI undefined
    p = fresh(); p.actual_cost.loc[:, :] = 0
    m = E.project_metrics(p, 1); c = " ".join(E.commentary(p, 1))
    assert math.isnan(m["CPI"]) and "no actual cost has been booked" in c
    return "CPI not defined; commentary asks for cost entry (no crash)"

def t12():  # budget exhausted: AC >= BAC with work remaining -> TCPI undefined and flagged
    p = fresh(); p.actual_cost["M4"] = p.actual_cost["M4"] * 20
    m = E.project_metrics(p, 4); c = " ".join(E.commentary(p, 4))
    assert m["Budget exhausted"] and math.isnan(m["TCPI (to BAC)"]) and "can no longer be met" in c
    return "TCPI not defined; commentary says budget cannot be met (was -0.105 before fix)"

def t13():  # gap in reporting: Month 3 blank, Month 4 filled -> data date is still Month 4
    p = fresh(); p.actual_pct["M3"] = float("nan")
    assert p.data_date == 4, p.data_date
    return f"data date = Month {p.data_date} (was 3 before fix)"

def t14():  # only Month 1 reported -> metrics and commentary work for a one-month project file
    p = fresh()
    for c in ["M2", "M3", "M4"]:
        p.actual_pct[c] = float("nan"); p.actual_cost[c] = float("nan")
    assert p.data_date == 1 and len(E.commentary(p, 1)) >= 3
    return "data date = Month 1; app shows a fixed month instead of a slider (slider needs min < max)"

for i, f in enumerate([t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13, t14], 1):
    case(f"E{i}", f)

if FAILED:
    raise SystemExit(f"{len(FAILED)} edge case(s) failed: {', '.join(FAILED)}")
