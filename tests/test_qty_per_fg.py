"""Quy ước SỐ LƯỢNG của kỹ thuật: SL mỗi dòng = TỔNG số chiếc trên 1 thành phẩm (đã nhân sẵn SL cụm).
Nhu cầu / 1 FG của mỗi chi tiết = SL dòng × KL (= cột TỔNG KL), bất kể cụm cha có SL bao nhiêu.
Chạy: python test_qty_per_fg.py"""
import io, os, sys, datetime as dt, warnings, logging
warnings.filterwarnings("ignore"); logging.disable(logging.CRITICAL)
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import openpyxl, pandas as pd
import app

HDR = ["MỤC", "MÃ CHI TIẾT", "MÔ TẢ", "VẬT LIỆU", "DÀY", "RỘNG", "DÀI", "SỐ LƯỢNG", "KHỐI LƯỢNG NGUYÊN LIỆU (kg)", "ĐVT"]
ROWS = [  # MỤC, mã, mô tả, vật liệu, dày, rộng, dài, SL (tổng theo FG), KL/1 chi tiết, ĐVT
    ["1", "FG-A", "Thành phẩm A", None, None, None, None, 1, None, "Bộ"],
    ["1.1", "CUM-X", "Cụm hàn X", None, None, None, None, 2, None, "Bộ"],
    ["1.1.1", "P1", "Thép tấm", "SS400", 12, 80, 170, 2, 1.28, "Kg"],      # 2 cụm × 1 tấm
    ["1.1.2", "P2", "Thép tấm", "SS400", 12, 65, 165, 2, 1.01, "Kg"],
    ["1.1.3", "P3", "Thép tròn đặc", "SS400", None, 24, 42, 2, 0.15, "Kg"],
    ["1.2", "P4", "Thép tấm", "SS400", 3, 100, 200, 3, 0.5, "Kg"],          # ngay dưới FG
    ["2", "FG-B", "Thành phẩm B", None, None, None, None, 1, None, "Bộ"],
    ["2.1", "CUM-Z", "Cụm Z", None, None, None, None, 2, None, "Bộ"],
    ["2.1.1", "CUM-W", "Cụm W", None, None, None, None, 4, None, "Bộ"],     # 2 Z × 2 W
    ["2.1.1.1", "P5", "Thép tấm", "SS400", 3, 50, 50, 8, 1.0, "Kg"],        # 4 W × 2 tấm
    ["2.2", "CUM-Y", "Cụm Y", None, None, None, None, 1, None, "Bộ"],
    ["2.2.1", "P6", "Thép tấm", "SS400", 3, 50, 60, 4, 1.0, "Kg"],          # cụm SL 1 → không đổi
    ["2.2.2", "BL", "Bulong M10", None, None, None, None, 6, None, "Cái"],  # vật tư mua ngoài trong cụm
]
EXPECT = {"FG-A": 2 * 1.28 + 2 * 1.01 + 2 * 0.15 + 3 * 0.5, "FG-B": 8 * 1.0 + 4 * 1.0}

wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Form"
for c, h in enumerate(HDR, 1):
    ws.cell(1, c, h)
for r, row in enumerate(ROWS, 2):
    for c, v in enumerate(row, 1):
        if v is not None:
            ws.cell(r, c, v)
b = io.BytesIO(); wb.save(b)
bom_df, sc, ctx, infos, _ = app.parse_bom(b.getvalue(), "qty.xlsx")
so = pd.DataFrame([{"STT": k + 1, "Mã Sản Phẩm KH": c, "Mã Sản Phẩm ERP": c, "Tên Thành Phẩm": c, "Quy Cách": "",
                    "ĐVT": "Bộ", "Số Lượng": q, "Giá Bán FOB ($USD)": 0.0} for k, (c, q) in enumerate([("FG-A", 1.0), ("FG-B", 10.0)])])
P = dict(project="T", er=25400.0, start_date=dt.date(2026, 11, 13), lead_days=21, sga=5.0, tax=20.0,
         default_scrap=0.0, kg_price=22000.0, unit_price=0.0, rm_prefix="R-", pallet_qty=50)
res = app.run_engine(bom_df, sc, ctx, so, None, P)
nodes, cum = res["nodes"], res["fg_cum"]

for fg, exp in EXPECT.items():
    kg = sum(v for c, v in cum[fg].items() if nodes[c]["type"] == "RM" and nodes[c]["uom"] == "Kg")
    print(f"{fg}: app {kg:.4f} kg/FG – kỹ thuật SUM(SL×KL) {exp:.4f}")
    assert abs(kg - exp) < 1e-9, (fg, kg, exp)
# Từng chi tiết: R_cum = SL tổng theo FG
for code, fg, n in (("P1", "FG-A", 2), ("P2", "FG-A", 2), ("P3", "FG-A", 2), ("P4", "FG-A", 3),
                    ("P5", "FG-B", 8), ("P6", "FG-B", 4), ("RM-BL", "FG-B", 6)):
    assert abs(cum[fg][code] - n) < 1e-9, (code, cum[fg][code], n)
# MRP = SL SO × nhu cầu / FG
m = {r["Mã NVL"]: r["Nhu cầu MRP"] for r in res["mrp_df"].to_dict("records")}
assert abs(m["RM-BL"] - 60) < 1e-9, m
print("Dòng 58-60 kiểu GSS: P1 = 2.56, P2 = 2.02, P3 = 0.30 kg / FG →",
      round(cum["FG-A"]["P1"] * 1.28, 4), round(cum["FG-A"]["P2"] * 1.01, 4), round(cum["FG-A"]["P3"] * 0.15, 4))
print("QTY-PER-FG TESTS PASSED")
