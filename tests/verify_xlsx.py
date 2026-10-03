import sys, os, warnings
warnings.filterwarnings("ignore")
D = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(D, "pylib"))
import formulas
xl = formulas.ExcelModel().loads(os.path.join(D, "master.xlsx")).finish()
sol = xl.calculate()

def get(sheet, cell):
    for k, v in sol.items():
        if k.upper().endswith(f"[MASTER.XLSX]{sheet.upper()}'!{cell}"):
            return v.value[0][0]
    return "NOT FOUND"

checks = [
    ("04. Hoach dinh MRP", "E9"), ("04. Hoach dinh MRP", "E10"), ("04. Hoach dinh MRP", "E13"),
    ("02. Bao gia & Don hang", "I11"), ("21. TK 621 CP NVL TT", "G16"), ("22. TK 622 CP NC TT", "E14"),
    ("23. TK 627 CP SXC", "E14"), ("20. Tong hop Gia thanh", "E9"), ("20. Tong hop Gia thanh", "E10"),
    ("20. Tong hop Gia thanh", "I11"),
]
for sh, c in checks:
    print(f"{sh}!{c} = {get(sh, c)}")
for r in range(9, 19):
    print("P&L", r, get("24. Bao cao P&L (KQKD)", f"B{r}"), get("24. Bao cao P&L (KQKD)", f"D{r}"))
print("Trace L9", get("19. Pha he Truy vet Nguoc", "L9"), "Pallet", get("18. Danh sach Pallet ID", "G10"))
print("WIP E9", get("12. Theo doi WIP", "B9"), get("12. Theo doi WIP", "E9"))
print("NCC", get("07. So sanh Gia NCC", "J9"), "IQC", get("09. Kiem tra NVL IQC", "J9"))
errs = [(k, v.value[0][0]) for k, v in sol.items() if hasattr(v, "value") and isinstance(v.value[0][0], formulas.tokens.operand.XlError)]
print("Formula errors:", len(errs), errs[:10])
