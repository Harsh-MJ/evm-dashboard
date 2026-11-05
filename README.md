# EVM Dashboard (CE725 Topic 7)

Run locally:
    pip install -r requirements.txt
    streamlit run app.py

Verify against the hand benchmark:
    python test_benchmark.py      # 30 checks, Months 2 and 4
    python test_edge_cases.py     # 14 edge cases (E11-E14 added after the second review)

Files
- app.py              Streamlit dashboard (5 tabs)
- evm_engine.py       Calculation engine (all formulas)
- sample_project.xlsx Input template + sample G+3 project (BAC Rs 1 crore)
- Benchmark_Workbook.xlsx  Plain-Excel formula benchmark with hand-check columns
- make_sample.py, make_benchmark.py  Scripts that generated the two workbooks

Changes in v1.1 (5 Nov 2026, second review): no crash when no cost is booked (AC = 0); TCPI reported as
'not defined' with a warning once actual cost reaches BAC (was negative); data date = last reported month even
if a month is blank; no slider crash when only Month 1 is reported; BAC caption in Indian grouping; redundant
'TCPI to EAC' (always equal to CPI) removed; benchmark PV for Months 5-6 now formula-derived.

Deploy for TAs (free): push this folder to a public GitHub repo -> share.streamlit.io -> New app -> pick repo, main file app.py -> Deploy. Put the resulting link in the submission folder.
