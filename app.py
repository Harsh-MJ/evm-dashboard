"""
EVM Dashboard - CE725 Construction Economics and Finance, Topic 7
Run:  streamlit run app.py
"""
import io
import math

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

import evm_engine as E

st.set_page_config(page_title="EVM Dashboard | CE725", layout="wide")

INK, PV_C, EV_C, AC_C, GRID = "#1f2933", "#7b8794", "#1b7f3b", "#c0392b", "#e4e7eb"
RAG = {"green": "#1b7f3b", "amber": "#d68910", "red": "#c0392b", "n/a": "#7b8794"}


def idx(x):
    return "–" if x is None or (isinstance(x, float) and math.isnan(x)) else f"{x:.3f}"


def lakh(x):
    return "–" if x is None or (isinstance(x, float) and math.isnan(x)) else f"{x/1e5:,.2f} L"


# ---------------------------------------------------------------- data load
st.sidebar.title("EVM Dashboard")
st.sidebar.caption("CE725 · Topic 7 · Earned Value Management")
up = st.sidebar.file_uploader("Upload project template (.xlsx)", type="xlsx")
if up is None:
    st.sidebar.info("Using built-in sample: G+3 residential block, BAC Rs 1 crore. "
                    "Download the template below, fill it in, and upload it.")
    src = "sample_project.xlsx"
else:
    src = up

base = E.load_excel(src)
with open("sample_project.xlsx", "rb") as f:
    st.sidebar.download_button("⬇ Download blank/sample template", f.read(), "evm_template.xlsx")

tabs = st.tabs(["📊 Dashboard", "🧱 Activities", "✏️ Edit inputs", "📐 Formulas", "ℹ️ About"])

# editable copies (Edit inputs tab) - stored in session so the dashboard reflects edits
key = f"data::{getattr(up, 'name', 'sample')}"
if key not in st.session_state:
    st.session_state[key] = {
        "activities": base.activities.copy(), "planned": base.planned.copy(),
        "actual_pct": base.actual_pct.copy(), "actual_cost": base.actual_cost.copy(),
    }
d = st.session_state[key]

with tabs[2]:
    st.subheader("Edit inputs (what-if)")
    st.caption("Change any cell; the dashboard updates immediately. Planned % and Actual % are cumulative; "
               "Actual cost is the amount spent in that month. Leave future months blank.")
    d["activities"] = st.data_editor(d["activities"], key=key + "a", num_rows="fixed")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**Planned cumulative % complete**")
        d["planned"] = st.data_editor(d["planned"], key=key + "p")
        st.markdown("**Actual cumulative % complete**")
        d["actual_pct"] = st.data_editor(d["actual_pct"], key=key + "x")
    with c2:
        st.markdown("**Actual cost incurred in month (Rs)**")
        d["actual_cost"] = st.data_editor(d["actual_cost"], key=key + "c")
    if st.button("Reset to uploaded file"):
        del st.session_state[key]
        st.rerun()

p = E.ProjectData(base.name, d["activities"], d["planned"], d["actual_pct"], d["actual_cost"])
msgs = E.validate(p)
for m in msgs:
    (st.sidebar.error if m.startswith("ERROR") else st.sidebar.warning)(m)
if any(m.startswith("ERROR") for m in msgs):
    st.error("Fix the input errors shown in the sidebar to see results.")
    st.stop()

if p.data_date > 1:
    month = st.sidebar.slider("Reporting month (status date)", 1, p.data_date, p.data_date)
else:  # a slider needs min < max, so with only Month 1 reported there is nothing to choose
    month = 1
    st.sidebar.markdown("**Reporting month (status date):** Month 1 (only month reported so far)")
eac_method = st.sidebar.radio("EAC method used for headline forecast",
                              ["EAC (BAC/CPI)", "EAC (AC+BAC-EV)", "EAC (CPIxSPI)"])
M = E.project_metrics(p, month)
T = E.trend(p).loc[:month]
A = E.activity_table(p, month)

# ---------------------------------------------------------------- dashboard
with tabs[0]:
    st.title(p.name)
    st.caption(f"Status at end of Month {month} of {M['Planned duration']} · BAC {E._rs(M['BAC'])} · amounts in lakh (L)")

    k = st.columns(6)
    k[0].metric("Planned Value (PV)", lakh(M["PV"]))
    k[1].metric("Earned Value (EV)", lakh(M["EV"]), f"{M['% Complete']:.1f}% complete", delta_color="off")
    k[2].metric("Actual Cost (AC)", lakh(M["AC"]), f"{M['% Spent']:.1f}% of BAC spent", delta_color="off")
    k[3].metric("CPI", idx(M["CPI"]), f"{M['CV']/1e5:+,.2f} L CV")
    k[4].metric("SPI", idx(M["SPI"]), f"{M['SV']/1e5:+,.2f} L SV")
    k[5].metric(eac_method.replace("EAC ", "EAC "), lakh(M[eac_method]),
                f"{(M['BAC'] - M[eac_method])/1e5:+,.2f} L VAC")

    left, right = st.columns([3, 2])
    with left:
        months = list(range(0, len(p.months) + 1))
        pv = E.pv_curve(p)
        ev = [0] + list(T["EV"]); ac = [0] + list(T["AC"])
        fig = go.Figure()
        fig.add_scatter(x=months, y=[v/1e5 for v in pv], name="PV (baseline)", line=dict(color=PV_C, dash="dot"))
        fig.add_scatter(x=months[:month+1], y=[v/1e5 for v in ev], name="EV", line=dict(color=EV_C, width=3))
        fig.add_scatter(x=months[:month+1], y=[v/1e5 for v in ac], name="AC", line=dict(color=AC_C, width=3))
        fig.add_scatter(x=[month, M["IEAC(t) (months)"]], y=[M["AC"]/1e5, M[eac_method]/1e5],
                        name="Forecast to EAC", line=dict(color=AC_C, dash="dash"))
        fig.add_hline(y=M["BAC"]/1e5, line=dict(color=INK, width=1), annotation_text="BAC")
        fig.add_vline(x=month, line=dict(color=GRID), annotation_text="status date")
        fig.update_layout(title="S-curve: PV, EV, AC (Rs lakh, cumulative)", height=420,
                          xaxis_title="Month", yaxis_title="Rs lakh", plot_bgcolor="white",
                          legend=dict(orientation="h", y=-0.2), margin=dict(t=50, l=10, r=10))
        fig.update_xaxes(gridcolor=GRID); fig.update_yaxes(gridcolor=GRID)
        st.plotly_chart(fig, width="stretch")
    with right:
        fig2 = go.Figure()
        fig2.add_scatter(x=T.index, y=T["CPI"], name="CPI", mode="lines+markers", line=dict(color=AC_C))
        fig2.add_scatter(x=T.index, y=T["SPI"], name="SPI", mode="lines+markers", line=dict(color=EV_C))
        fig2.add_scatter(x=T.index, y=T["SPI(t)"], name="SPI(t)", mode="lines+markers", line=dict(color=PV_C, dash="dot"))
        fig2.add_hrect(y0=0.9, y1=1.0, fillcolor="#fdebd0", opacity=0.5, line_width=0)
        fig2.add_hline(y=1.0, line=dict(color=INK, width=1))
        fig2.update_layout(title="Performance index trend", height=420, plot_bgcolor="white",
                           xaxis_title="Month", legend=dict(orientation="h", y=-0.2), margin=dict(t=50, l=10, r=10),
                           yaxis=dict(range=[min(0.5, T[['CPI','SPI']].min().min()-0.05), 1.15]))
        fig2.update_xaxes(gridcolor=GRID, dtick=1); fig2.update_yaxes(gridcolor=GRID)
        st.plotly_chart(fig2, width="stretch")

    st.subheader("Progress commentary")
    for line in E.commentary(p, month):
        st.markdown(f"- {line}")
    st.caption("Commentary is generated by fixed rules from the computed indices (no AI model at run time), "
               "so every number quoted matches the results table.")

    c1, c2 = st.columns([3, 2])
    with c1:
        st.subheader("Forecasts")
        fc = pd.DataFrame({
            "Method": ["BAC / CPI (current cost efficiency continues)", "AC + (BAC − EV) (rest at budget rate)",
                       "AC + (BAC − EV)/(CPI×SPI) (cost and schedule both)"],
            "EAC (Rs)": [M["EAC (BAC/CPI)"], M["EAC (AC+BAC-EV)"], M["EAC (CPIxSPI)"]],
        })
        fc["ETC (Rs)"] = fc["EAC (Rs)"] - M["AC"]
        fc["VAC (Rs)"] = M["BAC"] - fc["EAC (Rs)"]
        st.dataframe(fc.style.format({c: "{:,.0f}" for c in fc.columns[1:]}), hide_index=True, width="stretch")
        if M["Budget exhausted"]:
            st.markdown("**TCPI** to meet BAC: **not defined**: actual cost has reached the budget, "
                        "so BAC cannot be met at any efficiency.")
        else:
            st.markdown(f"**TCPI** to meet BAC: **{M['TCPI (to BAC)']:.3f}** (current CPI {M['CPI']:.3f}; "
                        f"a gap above 0.10 means the budget is unlikely to be met)")
    with c2:
        st.subheader("Schedule (Earned Schedule extension)")
        st.markdown(
            f"- Earned Schedule ES = **{M['ES (months)']:.3f} months** at actual time {month}\n"
            f"- SV(t) = ES − AT = **{M['SV(t) (months)']:.3f} months**\n"
            f"- SPI(t) = ES / AT = **{M['SPI(t)']:.3f}**\n"
            f"- Forecast duration IEAC(t) = PD / SPI(t) = **{M['IEAC(t) (months)']:.2f} months** (planned {M['Planned duration']})")
        st.caption("Unlike SPI, SPI(t) does not drift back to 1.0 as the project approaches completion (Lipke, 2003).")

    # export
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        pd.Series(M).to_frame("Value").to_excel(w, sheet_name="Project_metrics")
        A.to_excel(w, sheet_name="Activity_metrics")
        E.trend(p).to_excel(w, sheet_name="Trend")
        pd.DataFrame({"Commentary": E.commentary(p, month)}).to_excel(w, sheet_name="Commentary", index=False)
    st.download_button("⬇ Download results (Excel)", buf.getvalue(), f"evm_results_month{month}.xlsx")

# ---------------------------------------------------------------- activities
with tabs[1]:
    st.subheader(f"Activity drill-down at end of Month {month}")
    show = A.copy()
    show["Cost"] = show["CPI"].map(E.band).str.upper()
    show["Schedule"] = show["SPI"].map(E.band).str.upper()

    def colour(v):
        return f"color: {RAG[E.band(v)]}; font-weight: 600" if isinstance(v, float) else ""
    st.dataframe(
        show.style.format({"Budget (BAC)": "{:,.0f}", "PV": "{:,.0f}", "EV": "{:,.0f}", "AC": "{:,.0f}",
                           "CV": "{:,.0f}", "SV": "{:,.0f}", "CPI": "{:.3f}", "SPI": "{:.3f}",
                           "Planned %": "{:.0f}", "Actual %": "{:.0f}"}, na_rep="–")
            .map(colour, subset=["CPI", "SPI"]),
        width="stretch")
    st.caption("Status bands: GREEN ≥ 1.00, AMBER 0.90–0.99, RED < 0.90 (cost / schedule). "
               "Activities not yet started show – for indices.")
    fig3 = go.Figure()
    fig3.add_bar(x=show["Activity"], y=show["CV"]/1e5, name="CV", marker_color=AC_C)
    fig3.add_bar(x=show["Activity"], y=show["SV"]/1e5, name="SV", marker_color=PV_C)
    fig3.update_layout(title="Cost and schedule variance by activity (Rs lakh)", barmode="group",
                       plot_bgcolor="white", height=380)
    st.plotly_chart(fig3, width="stretch")
    st.subheader("Monthly trend (project level)")
    st.dataframe(E.trend(p).drop(columns="Budget exhausted").T.style.format("{:,.3f}", na_rep="–"),
                 width="stretch", height=600)

# ---------------------------------------------------------------- formulas
with tabs[3]:
    st.subheader("Formulas used")
    st.markdown(r"""
| Measure | Formula | Meaning |
|---|---|---|
| PV | $\sum_i BAC_i \times \text{planned cum.\,\%}_i$ | Budgeted cost of work scheduled |
| EV | $\sum_i BAC_i \times \text{actual cum.\,\%}_i$ | Budgeted cost of work performed |
| AC | $\sum_i \sum_{t\le T} \text{cost}_{i,t}$ | Actual cost of work performed |
| CV, SV | $EV-AC$, $EV-PV$ | Negative = over budget / behind schedule |
| CPI, SPI | $EV/AC$, $EV/PV$ | < 1 = inefficient / slow |
| EAC₁ | $BAC/CPI$ | Typical-variance forecast |
| EAC₂ | $AC + (BAC-EV)$ | Atypical (one-off) variance |
| EAC₃ | $AC + \dfrac{BAC-EV}{CPI\times SPI}$ | Cost and schedule both drive cost |
| ETC, VAC | $EAC-AC$, $BAC-EAC$ | Cost still to come, final over/under-run |
| TCPI | $\dfrac{BAC-EV}{BAC-AC}$ | CPI needed on remaining work to hit BAC |
| ES | $C + \dfrac{EV-PV_C}{PV_{C+1}-PV_C}$ | Time at which planned value equalled EV |
| SPI(t), IEAC(t) | $ES/AT$, $PD/SPI(t)$ | Time-based schedule index and duration forecast |
""")
    st.caption("Sources: PMI (2019) Standard for Earned Value Management; GAO-20-195G (2020); Lipke (2003). "
               "See the final report for the full reference list.")

with tabs[4]:
    st.markdown("""
**What this tool does.** Converts monthly planned progress, actual progress and actual cost for a construction
package into Earned Value Management indices, three EAC forecasts, TCPI, an Earned Schedule duration forecast,
activity-level RAG status and a plain-language progress commentary.

**Data privacy.** The app runs locally; uploaded files are held in memory only and are not sent to any AI service.

**Validation.** Every project-level figure was checked against a hand calculation at Month 2 and Month 4
(see `test_benchmark.py` and the Validation Report).

**Limitations.** Assumes % complete is reported honestly and on the same basis as the budget; linear interpolation
for Earned Schedule; no resource loading or critical-path analysis; INR single currency.
""")
