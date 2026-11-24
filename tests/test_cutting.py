"""Bất biến tối ưu cắt (chép từ METALIC_SMART_QUOTATION/backend/tests/test_mrp.py + thử ngẫu nhiên cho tấm).
Chạy: python test_cutting.py (từ trong thư mục tests/)."""
import os
import random
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from cutting import cut_bars, pack_sheets, summarize_bars  # noqa: E402


def _count(res):
    return sum(len(b.cuts) for b in res["bars"])


def test_total_length_division_would_be_short_but_optimizer_is_not():
    # 12 đoạn × 2,1 m = 25,2 m: chia tổng ra 5 cây 6 m (THIẾU); thực tế mỗi cây chỉ cắt 2 đoạn ⇒ 6 cây
    res = cut_bars([(2100, 12, "A")], [6000], kerf=3, trim=10)
    assert len(res["bars"]) == 6 and _count(res) == 12
    for b in res["bars"]:
        assert b.used(3) + 10 <= b.stock_mm


def test_kerf_counts_two_halves_plus_one_do_not_fit_exactly():
    # 2500 + 2500 + 1000 = 6000 trên giấy, nhưng 2 mạch cưa làm vượt ⇒ không được xếp chung 1 cây
    res = cut_bars([(2500, 2, "A"), (1000, 1, "B")], [6000], kerf=3, trim=0)
    assert len(res["bars"]) == 2 and _count(res) == 3


def test_mixed_lengths_and_shrink_to_shorter_stock():
    res = cut_bars([(2100, 12, "A")], [6000, 12000], kerf=3, trim=10)
    assert _count(res) == 12 and sum(b.stock_mm for b in res["bars"]) == 30000      # 2×12 m (5 đoạn) + 1×6 m (2 đoạn)
    res = cut_bars([(2500, 10, "A")], [4000], kerf=3)
    assert len(res["bars"]) == 10                                                    # cây 4 m chỉ được 1 đoạn 2,5 m


def test_too_long_piece_is_reported_not_dropped():
    res = cut_bars([(6500, 1, "LONG"), (500, 2, "S")], [6000], kerf=3, trim=10)
    assert not res["ok"] and res["too_long"] == ["LONG"]


def test_every_piece_is_cut_on_random_instances():
    rnd = random.Random(7)
    for _ in range(30):
        pcs = [(rnd.randint(150, 5900), rnd.randint(1, 15), f"P{i}") for i in range(rnd.randint(1, 8))]
        stocks = rnd.choice([[6000], [6000, 12000], [3000, 6000]])
        pcs = [p for p in pcs if p[0] + 3 <= max(stocks) - 10 + 3]
        res = cut_bars(pcs, stocks, kerf=3, trim=10)
        assert res["ok"]
        want = {}
        for ln, q, lab in pcs:
            want[(ln, lab)] = want.get((ln, lab), 0) + q
        got = {}
        for b in res["bars"]:
            assert b.used(3) + 10 <= b.stock_mm and b.stock_mm in stocks
            for c, lab in zip(b.cuts, b.labels):
                got[(c, lab)] = got.get((c, lab), 0) + 1
        assert got == want
        assert res["bought_mm"] >= sum(ln * q for ln, q, _ in pcs)
        s = summarize_bars(res, 3, 500, 10)
        assert s["bars"] == len(res["bars"]) and s["waste_mm"] >= 0


def test_pack_sheets_places_all_without_overlap():
    pcs = [(800, 400, 5, "P"), (300, 200, 10, "Q"), (1200, 1200, 1, "R")]
    res = pack_sheets(pcs, 1250, 2500, gap=5)
    placed = [q for sh in res["sheets"] for q in sh]
    assert len(placed) == 16 and not res["too_big"]
    for sh in res["sheets"]:
        for i, (x, y, l, w, _) in enumerate(sh):
            assert x + l <= 2500 and y + w <= 1250
            for (x2, y2, l2, w2, _) in sh[i + 1:]:
                assert x + l <= x2 or x2 + l2 <= x or y + w <= y2 or y2 + w2 <= y
    assert pack_sheets([(3000, 100, 1, "BIG")], 1250, 2500)["too_big"] == ["BIG"]


def test_pack_sheets_random_every_piece_placed():
    rnd = random.Random(11)
    for _ in range(30):
        W, Lg = rnd.choice([(1250, 2500), (1500, 3000), (1500, 6000)])
        pcs = [(rnd.randint(50, Lg), rnd.randint(30, W), rnd.randint(1, 20), f"S{i}") for i in range(rnd.randint(1, 6))]
        res = pack_sheets(pcs, W, Lg, gap=5)
        placed = {}
        for sh in res["sheets"]:
            for i, (x, y, l, w, lab) in enumerate(sh):
                assert x >= 0 and y >= 0 and x + l <= Lg and y + w <= W
                for (x2, y2, l2, w2, _) in sh[i + 1:]:
                    assert x + l <= x2 or x2 + l2 <= x or y + w <= y2 or y2 + w2 <= y
                placed[lab] = placed.get(lab, 0) + 1
        for a, b, q, lab in pcs:
            if lab in res["too_big"]:
                assert max(a, b) > Lg or min(a, b) > W
            else:
                assert placed.get(lab, 0) == q


if __name__ == "__main__":
    n = 0
    for name, fn in list(globals().items()):
        if name.startswith("test_") and callable(fn):
            fn()
            n += 1
            print("OK", name)
    print(f"ALL {n} CUTTING TESTS PASSED")
