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
if MODE == "W": files["up_so"] = ["sample_so.xlsx", "sample_so.xlsx"]; files.pop("up_po")
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
for mode in ("0", "2", "3", "S", "M", "D", "W"):
    os.environ["UI_MODE"] = mode
    at = AppTest.from_file(WRAP, default_timeout=120).run()
    if mode != "0":
        assert not at.metric, "chưa nhấn nút mà đã tính"
        [b for b in at.button if "TÍNH TOÁN" in b.label][0].click().run()
    print(f"=== MODE {mode}: exceptions={len(at.exception)} errors={[e.value for e in at.error]}")
    assert not at.exception, f"MODE {mode} có exception"
    if mode == "M":
        ws = [w.value for w in at.warning]
        assert any("file thừa" in w for w in ws) and not any("BAS-01" in w for w in ws), ws
        assert any("BAS-01" in i.value and "định mức khác nhau" in i.value for i in at.info), [i.value for i in at.info]
        assert all("Chưa xuất được phiếu PR" in e.value for e in at.error), "MODE M chỉ được có lỗi chặn PR"
        print(" warnings:", [w[:90] for w in ws])
    if mode in ("2", "S", "W"):  # Tab Đề xuất mua hàng (PR): điền nhanh → bảng tự cập nhật
        [b for b in at.button if "Điền nhanh" in b.label][0].click().run()
        assert not at.exception, [e.value for e in at.exception]
        dls = [d for d in at.get("download_button") if "PR" in str(d.proto.label)]
        if mode in ("2", "W"):  # mẫu: chân ghế thép hộp thiếu Rộng B → không ước lượng, chặn tải PR
            assert not dls and any("Chưa xuất được phiếu PR" in e.value and "Rộng B" in e.value for e in at.error),                 [e.value for e in at.error]
            assert any("Cấp phát" in m.value for m in at.markdown), "thiếu bảng cấp phát theo lệnh SX"
            print(" PR: OK – chặn tải vì thiếu kích thước, vẫn có bảng cấp phát cho dòng tính được")
        else:
            assert dls, "không có nút tải Phiếu PR"
            print(" PR: OK – có nút tải", dls[0].proto.label)
        if mode == "W":  # 2 file SO → 4 dòng SO, MRP gấp đôi
            so_tbl = [d.value for d in at.dataframe if "File SO" in d.value.columns]
            assert any(len(t) == 4 for t in so_tbl), ("thiếu cột File SO / sai số dòng", [len(t) for t in so_tbl])
            print(" Nhiều SO: OK – 4 dòng SO có cột File SO")
        if mode == "S":  # Tối ưu cắt: đổi phương án thép tấm 8.0 + mua thêm 2 → bảng PR cập nhật, không lỗi
            sbs = [s for s in at.selectbox if s.label == "Phương án"]
            assert len(sbs) == 6, len(sbs)
            sb = sbs[0]
            other = [k for k in ("1250x2500", "1500x3000", "1500x6000") if k != sb.value][0]
            sb.select(other).run()
            [n for n in at.number_input if n.label == "Mua thêm"][0].set_value(2).run()
            assert not at.exception, [e.value for e in at.exception]
            pr = [d.value for d in at.dataframe if "SỐ LƯỢNG CẦN MUA (4)" in d.value.columns][0]
            r0 = pr.iloc[0]
            assert r0["Mua thêm"] == 2 and r0["SỐ LƯỢNG CẦN MUA (4)"] == r0["Theo tối ưu (1)"] + r0["Contingency (3)"] + 2, r0
            assert r0["Phương án"].startswith("Khổ " + other.replace("x", "×")), (r0["Phương án"], other)
            assert [d for d in at.get("download_button") if "PR" in str(d.proto.label)]
            print(" PR tối ưu cắt: OK – đổi phương án", other, "+ mua thêm 2 →", r0["SỐ LƯỢNG CẦN MUA (4)"], r0["ĐVT"])
    if mode == "D":
        assert any("Lỗi trùng BOM" in e.value and "FG-CHAIR-02" in e.value for e in at.error), "MODE D phải báo trùng FG"
    for e in at.exception:
        print(e.value, "\n", "\n".join(e.stack_trace[-8:]))
    print(" success:", [s.value[:60] for s in at.success])
    print(" info:", [s.value[:60] for s in at.info][:2])
    print(" metrics:", [(m.label, m.value) for m in at.metric][:12])
    print(" mrp rows:", [len(d.value) for d in at.dataframe][:8]); print(" tabs:", len(at.tabs), " dataframes:", len(at.dataframe))
