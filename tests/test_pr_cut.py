"""Tối ưu cắt + cấp phát theo lệnh SX trong Phiếu PR. Chạy: python test_pr_cut.py (từ trong tests/; cần sample_*.xlsx
từ make_samples.py).
- Chỉ thuật toán: 12 đoạn 2,1 m → 6 cây 6 m; contingency áp vào số chi tiết; GHI CHÚ không ảnh hưởng kết quả.
- Thiếu kích thước / quá khổ / quá dài → lỗi, không ra số, không ước lượng; chặn xuất PR.
- Sơ đồ cắt chứa ĐỦ mọi chi tiết (đã cộng contingency), không chồng lấn, không vượt phôi.
- Cấp phát: Σ cấp phát mọi lệnh = KL mua theo sơ đồ; số chi tiết từng lệnh = SL BOM của lệnh; dự phòng riêng.
- Nhiều file SO: ghép dòng, cột File SO, MRP cộng chung.
- Excel PR: 7 sheet, tính lại bằng `formulas` khớp pr_compute, ô kiểm tra cấp phát = 0.
- Savic (nếu có file): MRP 30,896.59 kg; tấm 8.0 = 44 tấm 1500×6000, hộp 80x80x3 = 86 cây 6 m (3%)."""
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
SO2 = [{"Mã Sản Phẩm ERP": "FG-A", "Tên Thành Phẩm": "A"}, {"Mã Sản Phẩm ERP": "FG-B", "Tên Thành Phẩm": "B"}]


def pp(t=0, w=0, l=0, w2=0, kg=1.0, note=""):
    return {"rm": "RM-X", "kg": kg, "note": note, "desc": "ct", "t": t, "w": w, "w2": w2, "l": l, "klr": 7.85}


def row(group, std, parts, kg_net=100.0):
    """parts = [(mã, {dòng SO: SL chi tiết}, pp)] – giống dòng build_pr_base."""
    ps = [{"code": c, "q": sum(ln.values()), "pp": p, "lines": dict(ln)} for c, ln, p in parts]
    return {"Nhóm": group, "_std": std, "_parts": ps, "KL tịnh theo BOM (kg)": kg_net, "Tên mặt hàng và quy cách": std,
            "ĐVT": "Kg", "Kích thước": "", "KL / ĐVT (kg)": None, "Số lượng theo BOM (1)": kg_net, "Mã NVL app": "RM-X",
            "_mrp_lines": [kg_net / 2, kg_net / 2]}


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


def check_alloc(r, o, rate):
    """Σ cấp phát = KL mua; số chi tiết từng lệnh = SL BOM; dự phòng = phần contingency."""
    al = app.pr_allocate(r, o, rate)
    tot = sum(a["kg_alloc"] for a in al["lines"].values())
    assert abs(tot - o["kg1"]) < 1e-6, (tot, o["kg1"])
    want = Counter()
    for it in r["_parts"]:
        for j, n in app._split_int(it["lines"], app._ceil(it["q"])).items():
            want[j] += n
        want[app.PR_DP] += app._pr_qty(it["q"], rate, 1) - app._ceil(it["q"])
    got = Counter({j: a["pieces"] for j, a in al["lines"].items()})
    assert +got == +want, (got, want)
    for s in al["stocks"]:
        assert abs(sum(s["owners"].values()) - 1) < 1e-9
    return al


def test_bar_row():
    r = row(app.PR_GROUPS[1], "Thép tròn đặc phi 32", [("A", {0: 12}, pp(w=32, l=2100, kg=13.25))], kg_net=12 * 13.25)
    co = app.pr_cut_options(r, 0, CFG)
    o = {x["key"]: x for x in co["options"]}["6000"]
    assert o["u0"] == 6 and o["u1"] == 6, o          # chia tổng mét (25,2 / 6) chỉ ra 5 cây → thiếu
    co = app.pr_cut_options(r, 0.10, CFG)             # contingency 10% áp vào số chi tiết: 12 + ROUNDUP(1,2) = 14 đoạn
    o = co["options"][0]
    assert o["u0"] == 6 and o["u1"] == 7 and plan_pieces(o) == Counter({"A": 14})
    check_alloc(r, o, 0.10)
    cfg = app.pr_cut_cfg({"bars": [(6000, True), (12000, True)]})
    labels = [x["label"] for x in app.pr_cut_options(r, 0, cfg)["options"]]
    assert labels == ["Chỉ 6 m", "Chỉ 12 m", "Trộn 6 m + 12 m"], labels


def test_notes_never_used():
    a = row(app.PR_GROUPS[0], "Thép tấm 4.0", [("S", {0: 50}, pp(t=4, w=300, l=400, kg=3.77))])
    b = row(app.PR_GROUPS[0], "Thép tấm 4.0", [("S", {0: 50}, pp(t=4, w=300, l=400, kg=3.77,
                                                               note="1250x2500 được 2 pcs/1 tấm"))])
    ka = [(o["key"], o["u0"], o["u1"]) for o in app.pr_cut_options(a, 0.03, CFG)["options"]]
    kb = [(o["key"], o["u0"], o["u1"]) for o in app.pr_cut_options(b, 0.03, CFG)["options"]]
    assert ka == kb
    assert not hasattr(app, "_PR_SHEET_RE") and not hasattr(app, "pr_cut_defaults")


def test_kerf_row():
    r = row(app.PR_GROUPS[1], "Thép tròn đặc phi 20", [("A", {0: 2}, pp(w=20, l=2500)), ("B", {0: 1}, pp(w=20, l=1000))])
    o = app.pr_cut_options(r, 0, app.pr_cut_cfg({"trim": 0}))["options"][0]
    assert o["u0"] == 2                               # 2500+2500+1000 + 2 mạch cưa > 6000


def test_missing_dims_is_error():
    r = row(app.PR_GROUPS[1], "Thép hộp 40x40x2",
            [("OK", {0: 10}, pp(t=2, w=40, w2=40, l=1000, kg=2.4)), ("NOLEN", {0: 5}, pp(t=2, w=40, w2=40, l=0)),
             ("NOW2", {1: 3}, pp(t=2, w=40, l=900))])
    co = app.pr_cut_options(r, 0, CFG)
    assert co["options"] == [] and "NOLEN (thiếu Dài)" in co["error"] and "NOW2 (thiếu Rộng B)" in co["error"]
    e = app.pr_effective([r], [co])[0]
    assert e["opt"] is None and e["error"] and e["ĐVT"] == "Cây"
    assert app.pr_compute([r], [{}], [e])[0]["Số lượng mua (4)"] is None
    assert len(app.pr_blockers([r], [e])) == 1
    r0 = row(app.PR_GROUPS[0], "Thép tấm 3.0", [])
    assert "không có chi tiết" in app.pr_cut_options(r0, 0, CFG)["error"]


def test_too_long_or_too_big():
    r = row(app.PR_GROUPS[1], "Thép tròn đặc phi 20", [("LONG", {0: 2}, pp(w=20, l=7000))])
    co = app.pr_cut_options(r, 0, CFG)                 # chỉ bật cây 6 m → không cắt được → lỗi
    assert co["error"] and "LONG" in co["error"] and not co["options"]
    co = app.pr_cut_options(r, 0, app.pr_cut_cfg({"bars": [(6000, True), (12000, True)]}))
    assert [o["label"] for o in co["options"]] == ["Chỉ 12 m", "Trộn 6 m + 12 m"] and co["invalid"]
    s = row(app.PR_GROUPS[0], "Thép tấm 4.0", [("BIG", {0: 3}, pp(t=4, w=240, l=2734, kg=20.5)),
                                               ("S", {0: 50, 1: 20}, pp(t=4, w=300, l=400, kg=3.77))])
    co = app.pr_cut_options(s, 0.03, CFG)
    assert "1250x2500" not in {o["key"] for o in co["options"]} and co["invalid"][0][0] == "Khổ 1250×2500"
    for o in co["options"]:
        assert plan_pieces(o) == Counter({"S": 70 + 3, "BIG": 3 + 1}), o["key"]
        check_alloc(s, o, 0.03)


def test_allocation_shares():
    # 1 tấm 1250x2500: lệnh 0 chiếm 2 chi tiết 1000x1250 (gần hết) – lệnh 1 không có → toàn bộ tấm cho lệnh 0
    r = row(app.PR_GROUPS[0], "Thép tấm 10.0", [("P", {0: 2}, pp(t=10, w=600, l=1000, kg=47.1))])
    o = app.pr_cut_options(r, 0, CFG)["options"][0]
    al = check_alloc(r, o, 0)
    assert set(al["lines"]) == {0} and abs(al["lines"][0]["kg_alloc"] - o["kg1"]) < 1e-9
    # 2 lệnh chung 1 cây: chia theo chiều dài chi tiết
    b = row(app.PR_GROUPS[1], "Thép tròn đặc phi 20", [("X", {0: 1}, pp(w=20, l=3000)), ("Y", {1: 1}, pp(w=20, l=1500))])
    o = app.pr_cut_options(b, 0, CFG)["options"][0]
    al = check_alloc(b, o, 0)
    assert o["u1"] == 1 and abs(al["lines"][0]["kg_alloc"] / al["lines"][1]["kg_alloc"] - 2) < 1e-9
    rows = app.pr_allocation_rows([b], app.pr_effective([b], [app.pr_cut_options(b, 0, CFG)]), [{}], SO2)
    assert [x["Lệnh SX (dòng SO)"] for x in rows] == ["#1 FG-A", "#2 FG-B"]


def test_split_int():
    assert app._split_int({0: 2.4, 1: 1.6}, 4) == {0: 2, 1: 2}
    assert app._split_int({0: 3.0, 1: 2.0}, 5) == {0: 3, 1: 2}
    assert sum(app._split_int({0: 0.3, 1: 0.3, 2: 0.3}, 1).values()) == 1


def test_random_rows_every_piece_cut():
    import random
    rnd = random.Random(3)
    for _ in range(25):
        parts = [(f"P{i}", {j: rnd.randint(1, 30) for j in range(rnd.randint(1, 2))}, pp(w=25, l=rnd.randint(100, 5900)))
                 for i in range(rnd.randint(1, 6))]
        rate = rnd.choice([0, 0.03, 0.1])
        r = row(app.PR_GROUPS[1], "Thép tròn đặc phi 25", parts)
        for o in app.pr_cut_options(r, rate, app.pr_cut_cfg({"bars": [(6000, True), (12000, True)]}))["options"]:
            want = Counter()
            for c, ln, _ in parts:
                q = sum(ln.values())
                want[c] += q + app._ceil(q * rate)
            assert plan_pieces(o) == want and o["u1"] >= o["u0"]
            check_alloc(r, o, rate)
    for _ in range(15):
        parts = [(f"S{i}", {0: rnd.randint(1, 20), 1: rnd.randint(0, 20)},
                  pp(t=5, w=rnd.randint(50, 1200), l=rnd.randint(50, 2400))) for i in range(rnd.randint(1, 5))]
        r = row(app.PR_GROUPS[0], "Thép tấm 5.0", parts)
        for o in app.pr_cut_options(r, 0.05, CFG)["options"]:
            check_alloc(r, o, 0.05)


def excel_check(res, label, expect_block=False):
    base = app.build_pr_base(res)
    rows = base["rows"]
    inputs = [{"Số lượng tồn (2)": (k % 3), "Tỷ lệ contingency (%)": 3} for k in range(len(rows))]
    cut = app.pr_cut_all(rows, inputs, CFG)
    eff = app.pr_effective(rows, cut, extras={0: 1})
    assert bool(app.pr_blockers(rows, eff)) == expect_block, app.pr_blockers(rows, eff)
    calc_rows = app.pr_compute(rows, inputs, eff)
    xb = app.build_pr_excel(base, inputs, res, P, {"bom_names": ["b"], "so_name": "s"}, eff, CFG)
    wb = openpyxl.load_workbook(io.BytesIO(xb))
    assert wb.sheetnames == app.PR_SHEETS, wb.sheetnames
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
        if "'!" in k and hasattr(v, "value"):
            sh, cell = k.upper().split("'!")
            val[(sh.split("]")[-1], cell)] = v.value[0][0]
    errs = [k for k, v in sol.items() if hasattr(v, "value") and isinstance(v.value[0][0], formulas.tokens.operand.XlError)]
    assert not errs, errs[:5]
    ws = wb["PR_de_xuat"]
    rr = [c.row for c in ws["C"] if c.row >= 9 and c.value and isinstance(ws.cell(c.row, 1).value, int)]
    for r_, c in zip(rr, calc_rows):
        for col, key in (("O", "Số lượng mua (4)"), ("Q", "KL theo SL mua (kg)"), ("M", "Contingency (3)")):
            x = val.get(("PR_DE_XUAT", f"{col}{r_}"))
            y = c[key]
            x = None if x in ("", None) or str(x) == "empty" else float(x)
            assert (x is None and y is None) or abs(x - y) < 1e-6, (label, r_, col, x, y)
    w7 = wb["Cap_phat_LSX"]
    checks = [c.row for c in w7["H"] if c.value == "Kiểm tra (= 0)"]
    for r_ in checks:
        assert abs(float(val[("CAP_PHAT_LSX", f"I{r_}")])) < 1e-3, (label, r_)
    os.remove(path)
    print(f"  [{label}] Excel PR tính lại khớp pr_compute: {len(rr)} dòng × 3 cột; {len(checks)} ô kiểm tra cấp phát = 0")
    return base, cut, eff


def rd(n):
    return open(os.path.join(D, n), "rb").read(), n


def sample_res(so_files=("sample_so.xlsx",)):
    so_df, _, _ = app.parse_so_multi([rd(n) for n in so_files])
    bom_df, sc, ctx, _, _ = app.parse_bom(*rd("sample_bom.xlsx"))
    return app.run_engine(bom_df, sc, ctx, so_df, None, P)


def test_sample_excel():
    res = sample_res()
    m = {r["Mã NVL"]: r["Nhu cầu MRP"] for r in res["mrp_df"].to_dict("records")}
    assert abs(m["R-O-S01-00-032-001"] - 862.50) < 0.01 and abs(m["R-P-S01-052-00-002"] - 12540.00) < 0.01
    excel_check(res, "mau", expect_block=True)       # chân ghế thép hộp thiếu Rộng B → chặn


def test_multi_so():
    one, two = sample_res(), sample_res(("sample_so.xlsx", "sample_so.xlsx"))
    so2 = pd_records = two["so_lines"]
    assert len(pd_records) == 2 * len(one["so_lines"]) and all(l["File SO"] == "sample_so.xlsx" for l in so2)
    assert [l["STT"] for l in so2] == list(range(1, len(so2) + 1))
    a = {r["Mã NVL"]: r["Nhu cầu MRP"] for r in one["mrp_df"].to_dict("records")}
    b = {r["Mã NVL"]: r["Nhu cầu MRP"] for r in two["mrp_df"].to_dict("records")}
    assert a.keys() == b.keys() and all(abs(b[k] - 2 * a[k]) < 1e-6 for k in a)
    base = app.build_pr_base(two)
    eff = app.pr_effective(base["rows"], app.pr_cut_all(base["rows"], [{}] * len(base["rows"]), CFG))
    al = app.pr_allocation_rows(base["rows"], eff, [{}] * len(base["rows"]), so2)
    assert {x["Lệnh SX (dòng SO)"] for x in al} >= {"#1 FG-TABLE-01 – sample_so.xlsx", "#3 FG-TABLE-01 – sample_so.xlsx"}
    assert abs(app.parse_so_multi([rd("sample_so.xlsx")])[0]["Số Lượng"].sum()
               - app.parse_so(*rd("sample_so.xlsx"))[0]["Số Lượng"].sum()) < 1e-9


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
    print("\n  SAVIC (SO 200 + 200) – contingency 3% vào số chi tiết, phương án đề xuất")
    print(f"  {'Vật tư':<24}{'KL tịnh kg':>11}  {'Phương án':<16}{'(1)':>6}{'+cont':>7}{'% phế':>7}{'KL mua kg':>11}{'Σ cấp phát':>12}")
    got = {}
    for r, co, e in zip(base["rows"], cut, eff):
        if not co:
            continue
        o = e["opt"]
        al = check_alloc(r, o, 0.03)
        got[r["Tên mặt hàng và quy cách"]] = (o["key"], o["u0"])
        print(f"  {r['Tên mặt hàng và quy cách']:<24}{r['KL tịnh theo BOM (kg)']:>11,.1f}  {o['label']:<16}{o['u0']:>6}"
              f"{o['u1']:>7}{o['waste_pct']:>7.2f}{o['kg1']:>11,.1f}{sum(a['kg_alloc'] for a in al['lines'].values()):>12,.1f}")
    assert got["Thép tấm 8.0"] == ("1500x6000", 44) and got["Thép hộp 80x80x3"] == ("6000", 86), got


if __name__ == "__main__":
    n = 0
    for name, fn in list(globals().items()):
        if name.startswith("test_") and callable(fn):
            fn()
            n += 1
            print("OK", name)
    print(f"ALL {n} PR CUT TESTS PASSED")
