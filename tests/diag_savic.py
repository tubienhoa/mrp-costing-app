import os
import sys, datetime as dt, warnings, logging
warnings.filterwarnings("ignore"); logging.disable(logging.CRITICAL)
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import pandas as pd
pd.set_option("display.width", 250); pd.set_option("display.max_columns", 40); pd.set_option("display.max_colwidth", 28)
import app
B = r"C:\Users\MyPC\Desktop\ERP_App\Costing phien ban ky thuat\Savic"
so_df, w, si = app.parse_so(open(B + r"\template_don_hang_so.xlsx", "rb").read(), "so.xlsx")
print(so_df.to_string()); print(w, si["mapping"])
bom_df, sc, ctx, infos, bw = app.parse_bom(open(B + r"\BOM - Savic -dev.xlsx", "rb").read(), "bom.xlsx")
for i in infos: print(i)
print(bom_df.drop(columns=["Vật liệu"]).to_string())
P = dict(project="T", er=25400.0, start_date=dt.date(2026, 10, 2), lead_days=21, sga=5.0, tax=20.0, default_scrap=0.0,
         kg_price=22000.0, unit_price=0.0, rm_prefix="R-", pallet_qty=50)
res = app.run_engine(bom_df, sc, ctx, so_df, None, P)
print(res["mrp_df"][["Mã NVL", "Tên NVL", "ĐVT", "Nhu cầu MRP"]].to_string())
for w in res["warnings"]: print("W:", w)
if "-e" in sys.argv:
    for (p, c), e in res["edges"].items(): print(f"  {p} -> {c}: {e['rate']:g} eff {e['eff']:g}")
    rates = [{"stage": s, "labor": 10000, "oh": 5000} for s in res["used_stages"]]
    fin = app.compute_financials(res, so_df, rates, P)
    print(fin["sku"].to_string()); print(fin["pl"].to_string())

# ---- Quy tắc giá mới: không có PO → đơn giá 0, TK 621 = 0 ----
m = res["mrp_df"]
assert (m["Đơn giá (VND)"] == 0).all() and (m["Nguồn giá"] == "Chưa nạp PO (0đ)").all()
fin0 = app.compute_financials(res, so_df, [{"stage": s, "labor": 10000, "oh": 5000} for s in res["used_stages"]], P)
assert fin0["tk621"] == 0, fin0["tk621"]
assert abs(m["Nhu cầu MRP"].sum() - 30896.59) < 0.01
# Áp giá tạm Sidebar (chủ động) → tái lập mốc cũ 679,725,036
res2 = app.run_engine(bom_df, sc, ctx, so_df, None, dict(P, apply_sidebar=True))
assert round(res2["mrp_df"]["Thành tiền (VND)"].sum()) == 679725036
print("PRICING OK: không PO TK621 = 0; áp giá tạm Sidebar TK621 = 679,725,036")
