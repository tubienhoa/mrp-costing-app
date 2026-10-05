import sys, os, io, datetime as dt, warnings, logging
warnings.filterwarnings("ignore")
logging.disable(logging.CRITICAL)
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import app
import openpyxl

D = os.path.dirname(os.path.abspath(__file__))
rd = lambda n: open(os.path.join(D, n), "rb").read()
so_df, so_w, so_i = app.parse_so(rd("sample_so.xlsx"), "sample_so.xlsx")
print(so_df.to_string()); print(so_w, so_i)
bom_df, stage_cols, ctx, infos, bw = app.parse_bom(rd("sample_bom.xlsx"), "sample_bom.xlsx")
for i in infos: print(i)
print(bom_df[["Sheet", "MỤC", "Cấp", "Mã chi tiết", "Mô tả", "SL/Cha", "KL NVL (kg)", "Mã NVL", "Hao hụt (%)"] + stage_cols].to_string())
use_po = "--nopo" not in sys.argv
po_df = app.parse_po(rd("sample_po.xlsx"), "sample_po.xlsx")[0] if use_po else None
P = dict(project="TEST", er=25400.0, start_date=dt.date(2026, 10, 2), lead_days=21, sga=5.0, tax=20.0,
         default_scrap=0.0, kg_price=22000.0, unit_price=0.0, rm_prefix="R-", pallet_qty=50)
res = app.run_engine(bom_df, stage_cols, ctx, so_df, po_df, P)
print(res["mrp_df"].to_string())
for w in res["warnings"]: print("W:", w)
for (p, c), e in res["edges"].items(): print(f"  {p} -> {c}: {e['rate']:g} eff {e['eff']:g}")
rates = [{"stage": s, "labor": 10000, "oh": 5000} for s in res["used_stages"]]
fin = app.compute_financials(res, so_df, rates, P)
print(fin["pl"].to_string()); print(fin["sku"].to_string())
x = app.build_master_excel(res, fin, so_df, rates, P)
open(os.path.join(D, "master.xlsx"), "wb").write(x)
names, stats = app.count_formulas(x)
print(len(names), names)
print(stats.to_string())

# ---- Kiểm tra quy tắc định giá: chỉ lấy giá từ PO, còn lại = 0 ----
m = res["mrp_df"]
if use_po:
    for _, r in m.iterrows():
        if r["Nguồn giá"].startswith("PO:") or r["Nguồn giá"] == "BOM / Live Editor":
            assert r["Đơn giá (VND)"] > 0, r["Mã NVL"]
        else:
            assert r["Nguồn giá"] == "Chưa có giá PO (0đ)" and r["Đơn giá (VND)"] == 0, (r["Mã NVL"], r["Nguồn giá"])
    print("PRICING OK (có PO):", m["Nguồn giá"].value_counts().to_dict())
else:
    own = m["Nguồn giá"] == "BOM / Live Editor"   # giá do người dùng nhập tại BOM/Editor (mẫu: R-B-M10)
    assert (m.loc[~own, "Đơn giá (VND)"] == 0).all() and (m.loc[~own, "Nguồn giá"] == "Chưa nạp PO (0đ)").all()
    assert (m.loc[own, "Đơn giá (VND)"] > 0).all()
    assert abs(fin["tk621"] - m.loc[own, "Thành tiền (VND)"].sum()) < 1e-6
    print("PRICING OK (không PO): đơn giá = 0 trừ dòng nhập tay;", int(own.sum()), "dòng có giá BOM; TK 621 =", fin["tk621"])
