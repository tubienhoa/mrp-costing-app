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
