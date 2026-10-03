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
_orig = st.file_uploader
def fake(label, *a, key=None, **k):
    _orig(label, *a, key=key, **k)
    kk = key.rsplit("_", 1)[0]
    return F(files[kk]) if kk in files else None
st.file_uploader = fake
app.st.file_uploader = fake
app.main()
''')
for mode in ("0", "2", "3", "S"):
    os.environ["UI_MODE"] = mode
    at = AppTest.from_file(WRAP, default_timeout=120).run()
    if mode != "0":
        assert not at.metric, "chưa nhấn nút mà đã tính"
        [b for b in at.button if "TÍNH TOÁN" in b.label][0].click().run()
    print(f"=== MODE {mode}: exceptions={len(at.exception)} errors={[e.value for e in at.error]}")
    for e in at.exception:
        print(e.value, "\n", "\n".join(e.stack_trace[-8:]))
    print(" success:", [s.value[:60] for s in at.success])
    print(" info:", [s.value[:60] for s in at.info][:2])
    print(" metrics:", [(m.label, m.value) for m in at.metric][:12])
    print(" mrp rows:", [len(d.value) for d in at.dataframe][:8]); print(" tabs:", len(at.tabs), " dataframes:", len(at.dataframe))
