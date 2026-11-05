"""
Benchmark test: tool output vs. hand-calculated values (Validation Report, Tables V-2 and V-3).
Tolerances from the project charter: Rs 1 for amounts, 0.001 for indices.
Run:  python test_benchmark.py      (or: pytest test_benchmark.py)
"""
import evm_engine as E

HAND = {
    2: {"PV": 3200000, "EV": 2680000, "AC": 2860000, "CV": -180000, "SV": -520000,
        "CPI": 0.937, "SPI": 0.838, "EAC (BAC/CPI)": 10671642, "EAC (AC+BAC-EV)": 10180000,
        "EAC (CPIxSPI)": 12187333, "ETC (BAC/CPI)": 7811642, "VAC (BAC/CPI)": -671642,
        "TCPI (to BAC)": 1.025, "ES (months)": 1.698, "SPI(t)": 0.849},
    4: {"PV": 7150000, "EV": 6100000, "AC": 6490000, "CV": -390000, "SV": -1050000,
        "CPI": 0.940, "SPI": 0.853, "EAC (BAC/CPI)": 10639344, "EAC (AC+BAC-EV)": 10390000,
        "EAC (CPIxSPI)": 11353576, "ETC (BAC/CPI)": 4149344, "VAC (BAC/CPI)": -639344,
        "TCPI (to BAC)": 1.111, "ES (months)": 3.567, "SPI(t)": 0.892},
}
INDEX_KEYS = {"CPI", "SPI", "TCPI (to BAC)", "ES (months)", "SPI(t)"}


def compare():
    p = E.load_excel("sample_project.xlsx")
    rows, ok = [], True
    for month, exp in HAND.items():
        got = E.project_metrics(p, month)
        for k, hand in exp.items():
            tol = 0.001 if k in INDEX_KEYS else 1.0   # charter tolerances: 0.001 for indices, Rs 1 for amounts
            diff = got[k] - hand
            passed = abs(diff) <= tol
            ok &= passed
            rows.append((month, k, hand, got[k], diff, "PASS" if passed else "FAIL"))
    return rows, ok


def test_benchmark():
    rows, ok = compare()
    assert ok, [r for r in rows if r[-1] == "FAIL"]


if __name__ == "__main__":
    rows, ok = compare()
    print(f"{'Mo':>2} {'Measure':<18}{'Hand':>16}{'Tool':>18}{'Diff':>12}  Result")
    for r in rows:
        print(f"{r[0]:>2} {r[1]:<18}{r[2]:>16,.3f}{r[3]:>18,.4f}{r[4]:>12.4f}  {r[5]}")
    print("ALL PASS" if ok else "SOME FAILED")
