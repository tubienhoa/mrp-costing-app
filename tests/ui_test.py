import os, sys
from streamlit.testing.v1 import AppTest

D = os.path.dirname(os.path.abspath(__file__))
WRAP = os.path.join(D, "ui_wrapper.py")
open(WRAP, "w", encoding="utf-8").write(r'''
import os, sys
import streamlit as st
sys.path.insert(0, r"D:\CLAUDE CODE APP\MRP_APP")
import app
D = os.path.dirname(os.path.abspath(__file__))
MODE = os.environ.get("UI_MODE", "3")
class F:
    def __init__(self, n): self.name = n; self._b = open(os.path.join(D, n), "rb").read()
    def getvalue(self): return self._b
files = {"up_so": "sample_so.xlsx", "up_bom": "sample_bom.xlsx", "up_po": "sample_po.xlsx"}
if MODE == "S": files = {"up_so": r"C:\Users\MyPC\Desktop\ERP_App\Costing phien ban ky thuat\Savic\template_don_hang_so.xlsx", "up_bom": r"C:\Users\MyPC\Desktop\ERP_App\Costing phien ban ky thuat\Savic\BOM - Savic -dev.xlsx"}
if MODE == "0": files = {}
if MODE == "2": files.pop("up_po")
if MODE == "M": files["up_bom"] = ["sample_bom_table.xlsx", "sample_bom_chair_diff.xlsx", "sample_bom_extra.xlsx"]
if MODE == "D": files["up_bom"] = ["sample_bom.xlsx", "sample_bom_chair.xlsx"]
_orig = st.file_uploader
def fake(label, *a, key=None, **k):
    _orig(label, *a, key=key, **k)
    kk = key.rsplit("_", 1)[0]
    if kk not in files:
        return [] if k.get("accept_multiple_files") else None
    v = files[kk]
    if k.get("accept_multiple_files"):
        return [F(x) for x in (v if isinstance(v, list) else [v])]
    return F(v)
st.file_uploader = fake
app.st.file_uploader = fake
app.main()
''')
for mode in ("0", "2", "3", "S", "M", "D"):
    os.environ["UI_MODE"] = mode
    at = AppTest.from_file(WRAP, default_timeout=120).run()
    if mode != "0":
        assert not at.metric, "chưa nhấn nút mà đã tính"
        [b for b in at.button if "TÍNH TOÁN" in b.label][0].click().run()
    print(f"=== MODE {mode}: exceptions={len(at.exception)} errors={[e.value for e in at.error]}")
    assert not at.exception, f"MODE {mode} có exception"
    if mode == "M":
        ws = [w.value for w in at.warning]
        assert any("BAS-01" in w for w in ws) and any("file thừa" in w for w in ws), ws
        assert not at.error, "MODE M không được có lỗi"
        print(" warnings:", [w[:90] for w in ws])
    if mode == "D":
        assert any("Lỗi trùng BOM" in e.value and "FG-CHAIR-02" in e.value for e in at.error), "MODE D phải báo trùng FG"
    for e in at.exception:
        print(e.value, "\n", "\n".join(e.stack_trace[-8:]))
    print(" success:", [s.value[:60] for s in at.success])
    print(" info:", [s.value[:60] for s in at.info][:2])
    print(" metrics:", [(m.label, m.value) for m in at.metric][:12])
    print(" mrp rows:", [len(d.value) for d in at.dataframe][:8]); print(" tabs:", len(at.tabs), " dataframes:", len(at.dataframe))
