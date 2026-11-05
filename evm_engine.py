"""
evm_engine.py - Earned Value Management calculation engine
CE725 Construction Economics and Finance, Topic 7: EVM Dashboard

All formulas follow the standard definitions in:
  - PMI (2019) The Standard for Earned Value Management
  - GAO (2020) Cost Estimating and Assessment Guide, GAO-20-195G
  - Lipke, W. (2003) "Schedule is Different", The Measurable News (Earned Schedule)

Input tables (from the Excel template):
  activities : Activity ID, Activity, Budget (BAC_i)
  planned    : Activity ID + one column per month = planned CUMULATIVE % complete
  actual_pct : Activity ID + one column per month = actual CUMULATIVE % complete
  actual_cost: Activity ID + one column per month = actual cost INCURRED IN THAT MONTH (Rs)

Months for which actuals are not yet available are left blank.
"""
from __future__ import annotations

import math
from dataclasses import dataclass

import pandas as pd

MONTH_PREFIX = "M"


# --------------------------------------------------------------------------
# Loading and validation
# --------------------------------------------------------------------------
@dataclass
class ProjectData:
    name: str
    activities: pd.DataFrame      # index = Activity ID; columns Activity, Budget
    planned: pd.DataFrame         # index = Activity ID; columns M1..Mn (cum %)
    actual_pct: pd.DataFrame      # index = Activity ID; columns M1..Mk (cum %)
    actual_cost: pd.DataFrame     # index = Activity ID; columns M1..Mk (monthly Rs)

    @property
    def months(self) -> list[str]:
        return list(self.planned.columns)

    @property
    def bac(self) -> float:
        return float(self.activities["Budget"].sum())

    @property
    def data_date(self) -> int:
        """Last month for which actual % complete has been reported (a blank month in between is carried forward)."""
        filled = [i for i, c in enumerate(self.actual_pct.columns, 1) if self.actual_pct[c].notna().any()]
        return max(filled) if filled else 0


def _month_cols(df: pd.DataFrame) -> list[str]:
    return [c for c in df.columns if str(c).upper().startswith(MONTH_PREFIX) and str(c)[1:].isdigit()]


def _tidy(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    df["Activity ID"] = df["Activity ID"].astype(str).str.strip()
    return df.set_index("Activity ID")


def load_excel(path_or_buffer, name: str = "Project") -> ProjectData:
    xl = pd.ExcelFile(path_or_buffer)
    acts = _tidy(pd.read_excel(xl, "Activities"))
    plan = _tidy(pd.read_excel(xl, "Planned_%"))
    apct = _tidy(pd.read_excel(xl, "Actual_%"))
    acost = _tidy(pd.read_excel(xl, "Actual_Cost"))
    try:
        info = pd.read_excel(xl, "Project_Info", header=None)
        name = str(info.iloc[0, 1])
    except Exception:
        pass
    months = _month_cols(plan)
    plan = plan[months].astype(float)
    apct = apct[[m for m in months if m in apct.columns]].astype(float)
    acost = acost[[m for m in months if m in acost.columns]].astype(float)
    return ProjectData(name, acts[["Activity", "Budget"]], plan, apct, acost)


def validate(p: ProjectData) -> list[str]:
    """Return a list of human-readable input errors / warnings (empty = clean)."""
    msgs = []
    ids = set(p.activities.index)
    for label, df in [("Planned_%", p.planned), ("Actual_%", p.actual_pct), ("Actual_Cost", p.actual_cost)]:
        missing = ids - set(df.index)
        if missing:
            msgs.append(f"ERROR: sheet {label} is missing activities {sorted(missing)}")
    if (p.activities["Budget"] <= 0).any():
        msgs.append("ERROR: every activity budget must be positive")
    for label, df in [("Planned_%", p.planned), ("Actual_%", p.actual_pct)]:
        vals = df.ffill(axis=1).fillna(0)
        if ((vals < 0) | (vals > 100)).any().any():
            msgs.append(f"ERROR: {label} has values outside 0-100 %")
        if (vals.diff(axis=1).fillna(0) < -1e-9).any().any():
            msgs.append(f"WARNING: {label} decreases in some month - % complete should be cumulative")
    if (p.actual_cost.fillna(0) < 0).any().any():
        msgs.append("WARNING: negative actual cost found (credit note?) - check entry")
    final_plan = p.planned.ffill(axis=1).iloc[:, -1]
    short = final_plan[final_plan < 100 - 1e-9]
    if len(short):
        msgs.append(f"WARNING: planned % does not reach 100 for {list(short.index)}")
    if p.data_date == 0:
        msgs.append("ERROR: no actual progress has been entered yet")
    return msgs


# --------------------------------------------------------------------------
# Core calculations
# --------------------------------------------------------------------------
def _cum_pct(df: pd.DataFrame, upto: int) -> pd.Series:
    """Cumulative % at month `upto` (1-based). Blank cells carry the last value forward."""
    if upto == 0:
        return pd.Series(0.0, index=df.index)
    filled = df.iloc[:, :upto].ffill(axis=1).fillna(0)
    return filled.iloc[:, -1]


def activity_table(p: ProjectData, month: int) -> pd.DataFrame:
    """Activity-level PV, EV, AC, CV, SV, CPI, SPI at end of `month`."""
    b = p.activities["Budget"]
    pv = b * _cum_pct(p.planned, month) / 100.0
    ev = b * _cum_pct(p.actual_pct, month) / 100.0
    ac = p.actual_cost.iloc[:, :month].fillna(0).sum(axis=1)
    t = pd.DataFrame({
        "Activity": p.activities["Activity"],
        "Budget (BAC)": b,
        "Planned %": _cum_pct(p.planned, month),
        "Actual %": _cum_pct(p.actual_pct, month),
        "PV": pv, "EV": ev, "AC": ac,
    })
    t["CV"] = t["EV"] - t["AC"]
    t["SV"] = t["EV"] - t["PV"]
    t["CPI"] = t.apply(lambda r: r.EV / r.AC if r.AC else math.nan, axis=1)
    t["SPI"] = t.apply(lambda r: r.EV / r.PV if r.PV else math.nan, axis=1)
    return t


def pv_curve(p: ProjectData) -> list[float]:
    """Cumulative project PV at end of month 0..n."""
    return [0.0] + [float(activity_table(p, m)["PV"].sum()) for m in range(1, len(p.months) + 1)]


def earned_schedule(p: ProjectData, ev: float) -> float:
    """Earned Schedule (Lipke 2003): time at which PV equalled current EV.
    ES = C + (EV - PV_C) / (PV_{C+1} - PV_C), with C = last whole month where PV_C <= EV."""
    pv = pv_curve(p)
    if ev >= pv[-1]:
        return float(len(pv) - 1)
    c = max(i for i, v in enumerate(pv) if v <= ev + 1e-9)
    nxt = pv[c + 1]
    return c + (ev - pv[c]) / (nxt - pv[c]) if nxt > pv[c] else float(c)


def project_metrics(p: ProjectData, month: int) -> dict:
    """All project-level EVM measures at end of `month`."""
    t = activity_table(p, month)
    bac = p.bac
    pv, ev, ac = float(t.PV.sum()), float(t.EV.sum()), float(t.AC.sum())
    cpi = ev / ac if ac else math.nan
    spi = ev / pv if pv else math.nan
    pd_months = len(p.months)                       # planned duration
    eac1 = bac / cpi if cpi else math.nan           # future work at current CPI
    eac2 = ac + (bac - ev)                          # future work at budget rate
    eac3 = ac + (bac - ev) / (cpi * spi) if cpi and spi else math.nan  # CPI x SPI
    # TCPI is only meaningful while budget remains; once AC >= BAC the budget is exhausted
    # and BAC cannot be met at any efficiency, so TCPI is reported as undefined (flagged in commentary).
    tcpi_bac = (bac - ev) / (bac - ac) if bac - ac > 0 else math.nan
    es = earned_schedule(p, ev)
    spi_t = es / month if month else math.nan
    ieac_t = pd_months / spi_t if spi_t else math.nan
    return {
        "Month": month, "BAC": bac, "PV": pv, "EV": ev, "AC": ac,
        "CV": ev - ac, "SV": ev - pv, "CV%": (ev - ac) / ev * 100 if ev else math.nan,
        "SV%": (ev - pv) / pv * 100 if pv else math.nan,
        "CPI": cpi, "SPI": spi,
        "EAC (BAC/CPI)": eac1, "EAC (AC+BAC-EV)": eac2, "EAC (CPIxSPI)": eac3,
        "ETC (BAC/CPI)": eac1 - ac, "VAC (BAC/CPI)": bac - eac1,
        "TCPI (to BAC)": tcpi_bac, "Budget exhausted": ac >= bac,
        "% Complete": ev / bac * 100, "% Spent": ac / bac * 100,
        "ES (months)": es, "SV(t) (months)": es - month, "SPI(t)": spi_t,
        "Planned duration": pd_months, "IEAC(t) (months)": ieac_t,
    }


def trend(p: ProjectData) -> pd.DataFrame:
    """Project metrics for every month up to the data date."""
    return pd.DataFrame([project_metrics(p, m) for m in range(1, p.data_date + 1)]).set_index("Month")


# --------------------------------------------------------------------------
# Rule-based plain-language commentary (numbers come only from metrics)
# --------------------------------------------------------------------------
def _rs(x: float) -> str:
    """Indian-style rupee formatting, e.g. 1,00,00,000. Undefined values (NaN) print as 'not defined'."""
    if x is None or (isinstance(x, float) and math.isnan(x)):
        return "not defined"
    neg = x < 0
    s = f"{abs(round(x)):d}"
    if len(s) > 3:
        head, tail = s[:-3], s[-3:]
        head = ",".join([head[max(0, i - 2):i] for i in range(len(head), 0, -2)][::-1])
        s = head + "," + tail
    return ("-" if neg else "") + "Rs. " + s


def band(index: float) -> str:
    if math.isnan(index):
        return "n/a"
    if index >= 1.0:
        return "green"
    if index >= 0.9:
        return "amber"
    return "red"


def commentary(p: ProjectData, month: int) -> list[str]:
    m = project_metrics(p, month)
    t = activity_table(p, month)
    out = []
    cost_word = "under budget" if m["CV"] >= 0 else "over budget"
    if m["AC"] == 0:
        out.append(
            f"At the end of Month {month}, work worth {_rs(m['EV'])} has been earned against "
            f"{_rs(m['PV'])} planned, but no actual cost has been booked yet, so CPI and the cost "
            f"forecasts are not defined. Check that costs for this month have been entered."
        )
    else:
        out.append(
            f"At the end of Month {month}, work worth {_rs(m['EV'])} has been earned against "
            f"{_rs(m['PV'])} planned and {_rs(m['AC'])} actually spent. The project is "
            f"{_rs(abs(m['CV']))} {cost_word} (CPI = {m['CPI']:.3f})."
        )
    if m["Budget exhausted"]:
        out.append(
            f"Actual cost ({_rs(m['AC'])}) has reached or passed the budget ({_rs(m['BAC'])}) with "
            f"{_rs(m['BAC'] - m['EV'])} of work still to do. The original budget can no longer be met at any "
            f"efficiency (TCPI not defined); a revised budget (EAC) must be agreed."
        )
    if m["SV"] < 0:
        out.append(
            f"Progress is behind plan by {_rs(abs(m['SV']))} of work (SPI = {m['SPI']:.3f}); in time terms the "
            f"project has earned {m['ES (months)']:.2f} months of schedule in {month} months, i.e. about "
            f"{abs(m['SV(t) (months)']):.2f} months late."
        )
    elif m["PV"] == 0:
        out.append("No work was planned by this month, so SPI is not defined.")
    else:
        out.append(f"Progress is ahead of plan by {_rs(m['SV'])} of work (SPI = {m['SPI']:.3f}).")
    if m["AC"] > 0 and not math.isnan(m["EAC (BAC/CPI)"]):
        out.append(
            f"If the current cost efficiency continues, the project will cost {_rs(m['EAC (BAC/CPI)'])} at completion "
            f"({_rs(abs(m['VAC (BAC/CPI)']))} {'over' if m['VAC (BAC/CPI)'] < 0 else 'under'} the budget of "
            f"{_rs(m['BAC'])}), with a forecast duration of about {m['IEAC(t) (months)']:.1f} months against "
            f"{m['Planned duration']} planned."
        )
    gap = m["TCPI (to BAC)"] - m["CPI"]
    if gap > 0.10:
        out.append(
            f"To still finish within budget the remaining work must be done at a CPI of {m['TCPI (to BAC)']:.3f}, "
            f"which is {gap:.3f} above the CPI achieved so far ({m['CPI']:.3f}). A TCPI more than 0.10 above the "
            f"cumulative CPI is generally regarded as unachievable, so a revised budget (EAC) should be agreed."
        )
    elif gap > 0:
        out.append(
            f"Remaining work needs a CPI of {m['TCPI (to BAC)']:.3f} to meet the budget (current CPI "
            f"{m['CPI']:.3f}) - tight but within the 0.10 recovery threshold."
        )
    started = t[t["Actual %"] > 0]
    worst_cost = started.sort_values("CV").head(1)
    if len(worst_cost) and worst_cost.iloc[0].CV < 0:
        r = worst_cost.iloc[0]
        out.append(f"Largest cost overrun: {r.Activity} ({_rs(abs(r.CV))}, CPI {r.CPI:.3f}).")
    worst_sched = t.sort_values("SV").head(1)
    if len(worst_sched) and worst_sched.iloc[0].SV < 0:
        r = worst_sched.iloc[0]
        out.append(f"Largest schedule slippage: {r.Activity} ({_rs(abs(r.SV))} of planned work not yet earned).")
    return out
