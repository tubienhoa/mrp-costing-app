import sys, os, warnings, logging, datetime as dt
warnings.filterwarnings("ignore"); logging.disable(logging.CRITICAL)
sys.stdout.reconfigure(encoding="utf-8")
D = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__)))); sys.path.insert(0, os.path.join(D, "pylib"))
import app, formulas
B = r"C:\Users\MyPC\Desktop\ERP_App\Costing phien ban ky thuat\Savic"
so_df = app.parse_so(open(B + r"\template_don_hang_so.xlsx", "rb").read(), "so.xlsx")[0]
bom_df, sc, ctx, _, _ = app.parse_bom(open(B + r"\BOM - Savic -dev.xlsx", "rb").read(), "bom.xlsx")
P = dict(project="SAVIC", er=25400.0, start_date=dt.date(2026, 10, 2), lead_days=21, sga=5.0, tax=20.0, default_scrap=0.0,
         kg_price=22000.0, unit_price=0.0, rm_prefix="R-", pallet_qty=50)
res = app.run_engine(bom_df, sc, ctx, so_df, None, P)
rates = [{"stage": s, "labor": 10000, "oh": 5000} for s in res["used_stages"]]
fin = app.compute_financials(res, so_df, rates, P)
p = os.path.join(D, "master_savic.xlsx"); open(p, "wb").write(app.build_master_excel(res, fin, so_df, rates, P))
sol = formulas.ExcelModel().loads(p).finish().calculate()
def get(sh, c):
    for k, v in sol.items():
        if k.upper().endswith(f"]{sh.upper()}'!{c}"): return v.value[0][0]
xl_mrp = sum(get("04. Hoach dinh MRP", f"E{r}") for r in range(9, 9 + len(res["rm_used"])))
print("MRP tổng  Python:", round(res["mrp_df"]["Nhu cầu MRP"].sum(), 4), " Excel:", round(xl_mrp, 4))
print("TK621     Python:", round(fin["tk621"]), " Excel:", round(get("21. TK 621 CP NVL TT", f"G{9 + len(res['rm_used'])}")))
print("NPAT      Python:", round(fin["npat"]), " Excel:", round(get("24. Bao cao P&L (KQKD)", "D18")))
for j in range(2): print("Giá thành/SP", get("20. Tong hop Gia thanh", f"B{9+j}"), round(get("20. Tong hop Gia thanh", f"H{9+j}")), "| Py", round(fin["sku"].iloc[j]["Giá thành/SP"]))
