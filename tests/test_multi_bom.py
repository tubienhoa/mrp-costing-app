"""Kiểm thử nạp nhiều file BOM (mỗi file 1 FG). Chạy: python make_samples.py && python test_multi_bom.py"""
import os, sys, datetime as dt, warnings, logging
warnings.filterwarnings("ignore"); logging.disable(logging.CRITICAL)
sys.stdout.reconfigure(encoding="utf-8")
D = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(D))
import app


def rd(n):
    return open(os.path.join(D, n), "rb").read(), n


P = dict(project="T", er=25400.0, start_date=dt.date(2026, 11, 11), lead_days=21, sga=5.0, tax=20.0,
         default_scrap=0.0, kg_price=22000.0, unit_price=0.0, rm_prefix="R-", pallet_qty=50)
so_df, _, _ = app.parse_so(*rd("sample_so.xlsx"))


def run(files):
    bom_df, sc, ctx, infos, _ = app.parse_bom_multi([rd(f) for f in files])
    return app.run_engine(bom_df, sc, ctx, so_df, None, P), sc


def mrp(res):
    return {r["Mã NVL"]: round(r["Nhu cầu MRP"], 6) for r in res["mrp_df"].to_dict("records")}


# 1. Gộp 1 file == tách 2 file (cùng mã RM, cùng nhu cầu)
r1, sc1 = run(["sample_bom.xlsx"])
r2, sc2 = run(["sample_bom_table.xlsx", "sample_bom_chair.xlsx"])
assert mrp(r1) == mrp(r2), (mrp(r1), mrp(r2))
m = mrp(r2)
assert abs(m["R-O-S01-00-032-001"] - 862.50) < 1e-6 and abs(m["R-P-S01-052-00-002"] - 12540.00) < 1e-6, m
assert len(r1["edges"]) == len(r2["edges"]) and set(sc1) == set(sc2)
assert all(" › " in s for s in r2["fg_source"]["FG-CHAIR-02"])
print("1. OK – tách file = gộp file:", len(m), "mã RM, 862.50 / 12,540.00")

# 2. Một FG nằm ở 2 file → dừng báo lỗi
try:
    run(["sample_bom.xlsx", "sample_bom_chair.xlsx"])
    raise AssertionError("phải báo trùng FG")
except app.DuplicateFGError as e:
    assert "FG-CHAIR-02" in e.dup and len(e.dup["FG-CHAIR-02"]) == 2, e.dup
    print("2. OK – trùng FG:", e.dup)

# 3. Mã dùng chung có định mức khác nhau giữa 2 file → cảnh báo, dùng file đầu
r3, _ = run(["sample_bom_table.xlsx", "sample_bom_chair_diff.xlsx"])
w3 = [w for w in r3["warnings"] if "BAS-01" in w and "khác nhau" in w]
assert len(w3) == 1, r3["warnings"]
print("3. OK –", w3[0])
r3b, _ = run(["sample_bom_table.xlsx", "sample_bom_chair.xlsx"])
assert not [w for w in r3b["warnings"] if "khác nhau" in w], "không được cảnh báo khi cấu trúc giống nhau"

# 4. File BOM thừa (FG không có trong SO) → cảnh báo
r4, _ = run(["sample_bom_table.xlsx", "sample_bom_chair.xlsx", "sample_bom_extra.xlsx"])
assert any("sample_bom_extra.xlsx" in w and "file thừa" in w for w in r4["warnings"]), r4["warnings"]
assert mrp(r4) == mrp(r2)
print("4. OK – cảnh báo file thừa")

# 5. TK 622/627 theo công đoạn riêng từng FG
fs = r2["fg_stages"]
print("   công đoạn:", fs)
assert "Sơn" not in fs["FG-CHAIR-02"] and "Sơn" in fs["FG-TABLE-01"]
rates = [{"stage": s, "labor": 1000.0 * (k + 1), "oh": 100.0 * (k + 1)} for k, s in enumerate(r2["used_stages"])]
rate_of = {x["stage"]: x for x in rates}
fin = app.compute_financials(r2, so_df, rates, P)
exp622 = sum(l["Số Lượng"] * sum(rate_of[s]["labor"] for s in fs[l["Mã Sản Phẩm ERP"]]) for l in r2["so_lines"])
assert abs(fin["tk622"] - exp622) < 1e-6, (fin["tk622"], exp622)
assert abs(sum(rate_of[s]["labor"] * q for s, q in fin["stage_qty"].items()) - fin["tk622"]) < 1e-6
print(f"5. OK – TK 622 theo FG = {fin['tk622']:,.0f}")

# 6. Excel: sheet 22 cột D & sheet 20 cột F là công thức theo FG
x = app.build_master_excel(r2, fin, so_df, rates, P)
import io, openpyxl
wb = openpyxl.load_workbook(io.BytesIO(x))
ws22, ws20 = wb[app.SH[22]], wb[app.SH[20]]
for r in range(9, 9 + len(rates)):
    assert str(ws22.cell(r, 4).value).startswith("="), ws22.cell(r, 4).value
for r in range(9, 9 + len(r2["so_lines"])):
    assert str(ws20.cell(r, 6).value).startswith("="), ws20.cell(r, 6).value
print("6. OK – Excel 22/20 công thức sống:", ws22["D9"].value, "|", ws20["F10"].value)
print("ALL MULTI-BOM TESTS PASSED")
