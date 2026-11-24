"""Tối ưu cắt trong Phiếu PR. Chạy: python test_pr_cut.py (từ trong tests/; cần sample_*.xlsx từ make_samples.py).
- Dòng tự dựng: 12 đoạn 2,1 m → 6 cây 6 m; contingency áp vào số chi tiết; chi tiết thiếu kích thước → cờ ước lượng.
- Sơ đồ cắt chứa ĐỦ mọi chi tiết (đã cộng contingency), không chồng lấn, không vượt phôi.
- Excel PR: tính lại bằng thư viện `formulas` khớp pr_compute; 5 sheet.
- Savic (nếu có file): MRP không đổi 30,896.59 kg; in bảng so sánh cách cũ / tối ưu theo từng loại thép."""
import datetime as dt
import io
import logging
import os
import sys
import warnings
from collections import Counter

warnings.filterwarnings("ignore")
logging.disable(logging.CRITICAL)
sys.stdout.reconfigure(encoding="utf-8")
D = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(D))
sys.path.insert(0, os.path.join(D, "pylib"))
import openpyxl  # noqa: E402

import app  # noqa: E402

P = dict(project="T", er=25400.0, start_date=dt.date(2026, 11, 24), lead_days=21, sga=5.0, tax=20.0,
         default_scrap=0.0, kg_price=22000.0, unit_price=0.0, rm_prefix="R-", pallet_qty=50)
CFG = app.pr_cut_cfg(None)


def pp(t=0, w=0, l=0, w2=0, kg=1.0, note=""):
    return {"rm": "RM-X", "kg": kg, "note": note, "desc": "ct", "t": t, "w": w, "w2": w2, "l": l, "klr": 7.85}


def row(group, std, parts, kg_net=100.0):
    """parts = [(mã, SL chi tiết, pp)] – giống dòng build_pr_base (chưa có ghi chú định mức → cách cũ quy KL)."""
    return {"Nhóm": group, "_std": std, "_parsed": [(c, q, p, "kg", None, None) for c, q, p in parts],
            "Cách cũ": kg_net, "ĐVT cách cũ": "Kg", "KL tịnh theo BOM (kg)": kg_net, "Tên mặt hàng và quy cách": std}


def plan_pieces(o):
    got = Counter()
    if o["kind"] == "bar":
        for b in o["plan"]:
            assert sum(b["cuts"]) + CFG["kerf"] * (len(b["cuts"]) - 1) + CFG["trim"] <= b["stock_mm"]
            for lab in b["labels"]:
                got[lab] += 1
    else:
        for sh in o["plan"]:
            for i, (x, y, a, b, lab) in enumerate(sh):
                assert x + a <= o["Lg"] and y + b <= o["W"]
                for (x2, y2, a2, b2, _) in sh[i + 1:]:
                    assert x + a <= x2 or x2 + a2 <= x or y + b <= y2 or y2 + b2 <= y
                got[lab] += 1
    return got


def test_bar_row():
    r = row(app.PR_GROUPS[1], "Thép tròn đặc phi 32", [("A", 12, pp(w=32, l=2100, kg=13.25))], kg_net=12 * 13.25)
    co = app.pr_cut_options(r, 0, CFG)
    o = {x["key"]: x for x in co["options"]}["6000"]
    assert o["u0"] == 6 and o["u1"] == 6, o          # chia tổng mét (25,2 / 6) chỉ ra 5 cây → thiếu
    assert o["old"] == 5                              # cách cũ quy KL: 159 kg / 37,88 kg/cây → 5 (thiếu 1 cây)
    co = app.pr_cut_options(r, 0.10, CFG)             # contingency 10% áp vào số chi tiết: 12 + ROUNDUP(1,2) = 14 đoạn
    o = co["options"][0]
    assert o["u0"] == 6 and o["u1"] == 7 and plan_pieces(o) == Counter({"A": 14})
    # 12 m bật thêm → có 3 phương án: chỉ 6 m / chỉ 12 m / trộn
    cfg = app.pr_cut_cfg({"bars": [(6000, True), (12000, True)]})
    labels = [x["label"] for x in app.pr_cut_options(r, 0, cfg)["options"]]
    assert labels == ["Chỉ 6 m", "Chỉ 12 m", "Trộn 6 m + 12 m"], labels


def test_kerf_row():
    r = row(app.PR_GROUPS[1], "Thép tròn đặc phi 20", [("A", 2, pp(w=20, l=2500)), ("B", 1, pp(w=20, l=1000))])
    cfg = app.pr_cut_cfg({"trim": 0})
    o = app.pr_cut_options(r, 0, cfg)["options"][0]
    assert o["u0"] == 2                               # 2500+2500+1000 + 2 mạch cưa > 6000


def test_missing_dims_flagged():
    r = row(app.PR_GROUPS[1], "Thép hộp 40x40x2",
            [("OK", 10, pp(t=2, w=40, w2=40, l=1000, kg=2.4)), ("NOLEN", 5, pp(t=2, w=40, w2=40, l=0, kg=2.4)),
             ("NOW2", 3, pp(t=2, w=40, l=900, kg=2.1))])
    co = app.pr_cut_options(r, 0, CFG)
    o = co["options"][0]
    assert set(co["missing"]) == {"NOLEN", "NOW2"} and o["est0"] >= 1 and set(o["est_codes"]) == {"NOLEN", "NOW2"}
    assert plan_pieces(o) == Counter({"OK": 10})
    eff = app.pr_effective([r], [co])[0]
    assert eff["flag"] and "ước lượng" in eff["Phương án"]
    # Không có chi tiết nào đủ kích thước → không tối ưu, giữ cách cũ, gắn cờ
    r2 = row(app.PR_GROUPS[0], "Thép tấm 3.0", [("X", 4, pp(t=3, w=100, l=0))])
    r2.update({"ĐVT": "Kg", "Kích thước": "", "KL / ĐVT (kg)": None, "Số lượng theo BOM (1)": 100.0})
    assert app.pr_cut_options(r2, 0, CFG) is None
    e2 = app.pr_effective([r2], [None])[0]
    assert e2["flag"] and e2["opt"] is None and e2["ĐVT"] == "Kg"


def test_sheet_row_and_too_big():
    r = row(app.PR_GROUPS[0], "Thép tấm 4.0", [("BIG", 3, pp(t=4, w=240, l=2734, kg=20.5)),
                                               ("S", 50, pp(t=4, w=300, l=400, kg=3.77))])
    co = app.pr_cut_options(r, 0.03, CFG)
    by = {o["key"]: o for o in co["options"]}
    assert by["1250x2500"]["too_long"] == ["BIG"] and co["rec"] != "1250x2500"   # quá khổ → không đề xuất
    for o in co["options"]:
        got = plan_pieces(o)
        want = Counter({"S": 50 + 2, "BIG": 3 + 1})
        for lab in o["too_long"]:
            want.pop(lab)
        assert got == want, (o["key"], got, want)


def test_random_rows_every_piece_cut():
    import random
    rnd = random.Random(3)
    for _ in range(25):
        parts = [(f"P{i}", rnd.randint(1, 60), pp(w=25, l=rnd.randint(100, 5900), kg=1.0)) for i in range(rnd.randint(1, 6))]
        rate = rnd.choice([0, 0.03, 0.1])
        r = row(app.PR_GROUPS[1], "Thép tròn đặc phi 25", parts)
        for o in app.pr_cut_options(r, rate, app.pr_cut_cfg({"bars": [(6000, True), (12000, True)]}))["options"]:
            want = Counter()
            for c, q, _ in parts:
                want[c] += q + app._ceil(q * rate)
            assert plan_pieces(o) == want and o["u1"] >= o["u0"]
            Lmin = min(o["lens"])
            assert o["bought_mm"] >= sum((p["l"]) * q for _, q, p in parts) and o["u1"] * Lmin <= o["bought_mm"]


def excel_check(res, label):
    base = app.build_pr_base(res)
    rows = base["rows"]
    inputs = [{"Số lượng tồn (2)": (k % 3), "Tỷ lệ contingency (%)": 3, "SL PR Kỹ thuật (5)": (None if k % 2 else 5)}
              for k in range(len(rows))]
    cut = app.pr_cut_all(rows, inputs, CFG)
    eff = app.pr_effective(rows, cut, extras={0: 1})
    calc_rows = app.pr_compute(rows, inputs, eff)
    xb = app.build_pr_excel(base, inputs, res, P, {"bom_names": ["b"], "so_name": "s"}, eff, CFG)
    wb = openpyxl.load_workbook(io.BytesIO(xb))
    assert wb.sheetnames == ["PR_de_xuat", "Chi_tiet_theo_model", "Dinh_muc_cat", "Ghi_chu", "So_do_cat"], wb.sheetnames
    try:
        import formulas
    except ImportError:
        print(f"  [{label}] (bỏ qua tính lại Excel: chưa cài formulas)")
        return base, cut, eff
    path = os.path.join(D, f"_pr_{label}.xlsx")
    open(path, "wb").write(xb)
    sol = formulas.ExcelModel().loads(path).finish().calculate()
    val = {}
    for k, v in sol.items():
        if "PR_DE_XUAT'!" in k.upper() and hasattr(v, "value"):
            val[k.upper().split("'!")[1]] = v.value[0][0]
    errs = [k for k, v in sol.items() if hasattr(v, "value") and isinstance(v.value[0][0], formulas.tokens.operand.XlError)]
    assert not errs, errs[:5]
    ws = wb["PR_de_xuat"]
    rr = [c.row for c in ws["C"] if c.row >= 9 and c.value and ws.cell(c.row, 1).value is not None
          and isinstance(ws.cell(c.row, 1).value, int)]
    n = 0
    for r_, c in zip(rr, calc_rows):
        for col, key in (("P", "Số lượng mua (4)"), ("S", "Số lượng đặt mua"), ("U", "KL theo SL mua (kg)"),
                         ("N", "Contingency (3)")):
            x = val.get(f"{col}{r_}")
            y = c[key]
            x = None if x in ("", None) else float(x)
            assert (x is None and y is None) or abs(x - y) < 1e-6, (label, r_, col, x, y)
            n += 1
    os.remove(path)
    print(f"  [{label}] Excel PR tính lại khớp pr_compute: {len(rr)} dòng × 4 cột, 0 lỗi công thức")
    return base, cut, eff


def test_sample_excel():
    rd = lambda n: (open(os.path.join(D, n), "rb").read(), n)  # noqa: E731
    so_df, _, _ = app.parse_so(*rd("sample_so.xlsx"))
    bom_df, sc, ctx, _, _ = app.parse_bom(*rd("sample_bom.xlsx"))
    res = app.run_engine(bom_df, sc, ctx, so_df, None, P)
    m = {r["Mã NVL"]: r["Nhu cầu MRP"] for r in res["mrp_df"].to_dict("records")}
    assert abs(m["R-O-S01-00-032-001"] - 862.50) < 0.01 and abs(m["R-P-S01-052-00-002"] - 12540.00) < 0.01
    excel_check(res, "mau")


def test_savic():
    B = r"C:\Users\MyPC\Desktop\ERP_App\Costing phien ban ky thuat\Savic"
    if not os.path.exists(B):
        print("  (bỏ qua Savic: không có file)")
        return
    so_df, _, _ = app.parse_so(open(B + r"\template_don_hang_so.xlsx", "rb").read(), "so.xlsx")
    bom_df, sc, ctx, _, _ = app.parse_bom(open(B + r"\BOM - Savic -dev.xlsx", "rb").read(), "bom.xlsx")
    res = app.run_engine(bom_df, sc, ctx, so_df, None, P)
    assert abs(res["mrp_df"]["Nhu cầu MRP"].sum() - 30896.59) < 0.01
    base, cut, eff = excel_check(res, "savic")
    print("\n  SAVIC (SO 200 + 200) – contingency 3% vào số chi tiết, phương án đề xuất, cây 6 m, khổ 1250×2500 / 1500×3000 / 1500×6000")
    print(f"  {'Vật tư':<24}{'KL tịnh kg':>11}  {'Phương án':<16}{'Cách cũ':>8}{'Tối ưu (1)':>11}{'+cont':>7}{'% phế':>7}  KL mua kg")
    for r, co, e in zip(base["rows"], cut, eff):
        if not co:
            continue
        o = e["opt"]
        print(f"  {r['Tên mặt hàng và quy cách']:<24}{r['KL tịnh theo BOM (kg)']:>11,.1f}  {o['label']:<16}{o['old']:>8}"
              f"{o['u0']:>11}{o['u1']:>7}{o['waste_pct']:>7.2f}  {o['kg1']:,.1f}")
        want = Counter()
        for c, q, p, *_ in r["_parsed"]:
            want[app.show_code(c)] += app._ceil(q) + app._ceil(app._ceil(q) * 0.03)
        assert plan_pieces(o) == want and not o["est_codes"]


if __name__ == "__main__":
    n = 0
    for name, fn in list(globals().items()):
        if name.startswith("test_") and callable(fn):
            fn()
            n += 1
            print("OK", name)
    print(f"ALL {n} PR CUT TESTS PASSED")
