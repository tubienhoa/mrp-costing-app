# Chép nguyên văn từ D:\METALIC_SMART_QUOTATION\backend\app\mrp\cutting.py (2026-11-24) – chỉ dùng cho Phiếu PR.
# Không sửa thuật toán ở đây mà không chạy lại tests/test_cutting.py.
"""Tối ưu cắt (hàm thuần, không đụng CSDL).

1 chiều (cây: ống, hộp, V, U, I, láp, la…): bài toán cutting stock. Thuật toán "chọn mẫu cắt tốt nhất rồi lặp" (sequential
heuristic): mỗi vòng, với từng loại cây được phép, giải bài toán xếp ba lô có giới hạn (bounded knapsack) để tìm mẫu cắt
dùng nhiều chiều dài nhất; chọn mẫu ít phế nhất theo tỷ lệ, lặp mẫu đó nhiều lần nhất có thể. Kết quả luôn ĐỦ chi tiết
(mỗi đoạn được xếp vào một cây cụ thể); kèm cận dưới để biết còn cách tối ưu bao xa.

2 chiều (tấm): xếp chữ nhật theo kệ (FFDH: cao trước, cho xoay 90°), có khe cắt giữa các chi tiết. Chi tiết thiếu kích thước
phôi được ước lượng theo kg (gắn cờ), không bao giờ bỏ qua.

Đơn vị: mm, kg. Mạch cưa (kerf) cộng vào mỗi đoạn; phần bỏ đầu cây (trim) trừ khỏi chiều dài cây.
"""
from __future__ import annotations

import math
from dataclasses import dataclass, field

MAX_CELLS = 4000          # độ phân giải DP: chiều dài cây chia tối đa 4000 ô (≈1,5 mm/ô với cây 6 m) – làm tròn LÊN nên luôn cắt được


@dataclass
class Bar:
    stock_mm: int
    cuts: list[int] = field(default_factory=list)          # chiều dài từng đoạn (mm, chưa gồm mạch cưa)
    labels: list[str] = field(default_factory=list)        # mã chi tiết tương ứng từng đoạn

    def used(self, kerf: int) -> int:
        return sum(self.cuts) + kerf * max(0, len(self.cuts) - 1) if self.cuts else 0


def _knapsack(items: list[tuple[int, int, int]], cap_cells: int) -> dict[int, int]:
    """items = [(chỉ số loại đoạn, số ô, số lượng còn)] → {chỉ số: số đoạn} có tổng ô lớn nhất ≤ cap_cells.
    Ba lô có giới hạn bằng tách nhị phân (1, 2, 4, …) – mỗi gói chỉ dùng một lần."""
    packs: list[tuple[int, int, int]] = []               # (chỉ số, số ô của gói, số đoạn trong gói)
    for idx, w, cnt in items:
        if w <= 0 or cnt <= 0 or w > cap_cells:
            continue
        cnt = min(cnt, cap_cells // w)
        k = 1
        while cnt > 0:
            t = min(k, cnt)
            packs.append((idx, w * t, t))
            cnt -= t
            k *= 2
    # tập tổng đạt được giữ dạng bit của một số nguyên; ô c được đạt lần đầu ở gói who[c] (từ c − w của gói đó,
    # vốn đạt được bằng các gói TRƯỚC nó) ⇒ truy ngược được và mỗi gói dùng tối đa một lần.
    mask = (1 << (cap_cells + 1)) - 1
    reach = 1
    who = [-1] * (cap_cells + 1)
    for p, (_, w, _) in enumerate(packs):
        nw = (reach | (reach << w)) & mask
        newly = nw & ~reach
        while newly:
            low = newly & -newly
            who[low.bit_length() - 1] = p
            newly ^= low
        reach = nw
    out: dict[int, int] = {}
    c = reach.bit_length() - 1
    while c > 0:
        idx, w, t = packs[who[c]]
        out[idx] = out.get(idx, 0) + t
        c -= w
    return out


def cut_bars(pieces: list[tuple[int, int, str]], stocks: list[int], kerf: int = 3, trim: int = 0) -> dict:
    """pieces = [(dài mm, số đoạn, nhãn)], stocks = các chiều dài cây được phép mua (mm).
    Trả {bars: [Bar], ok, too_long: [nhãn], lower_bound_mm, bought_mm}."""
    stocks = sorted({int(s) for s in stocks if s and s > 0})
    need: dict[tuple[int, str], int] = {}
    for ln, q, lab in pieces:
        if ln > 0 and q > 0:
            need[(int(math.ceil(ln)), lab)] = need.get((int(math.ceil(ln)), lab), 0) + int(q)
    if not stocks:
        return {"bars": [], "ok": False, "too_long": [], "reason": "Chưa chọn loại cây nào", "lower_bound_mm": 0, "bought_mm": 0}
    cap_max = max(stocks) - trim + kerf
    too_long = sorted({lab for (ln, lab) in need if ln + kerf > cap_max})
    if too_long:
        return {"bars": [], "ok": False, "too_long": too_long,
                "reason": f"Có đoạn dài hơn cây dài nhất ({max(stocks)} mm trừ đầu cây): cần nối hoặc đặt cây dài hơn",
                "lower_bound_mm": 0, "bought_mm": 0}
    keys = list(need)
    remain = [need[k] for k in keys]
    bars: list[Bar] = []
    while any(remain):
        cand = None
        for s in stocks:
            cap = s - trim + kerf                          # n đoạn dùng n mạch cưa, cộng lại một mạch "ảo" ở cuối
            cell = max(1, math.ceil(cap / MAX_CELLS))
            items = [(i, math.ceil((keys[i][0] + kerf) / cell), remain[i]) for i in range(len(keys)) if remain[i]]
            pat = _knapsack(items, cap // cell)
            if not pat:
                continue
            used = sum(keys[i][0] * n for i, n in pat.items()) + kerf * (sum(pat.values()) - 1)
            waste = (s - used) / s
            # ít phế theo tỷ lệ là tốt nhất; bằng nhau thì ưu tiên cây dài hơn (ít cây, ít lần cắt)
            score = (round(waste, 6), -s)
            if cand is None or score < cand[0]:
                cand = (score, s, pat)
        if cand is None:                                    # không thể xảy ra vì đã kiểm too_long
            break
        _, s, pat = cand
        rep = min(remain[i] // n for i, n in pat.items())          # ≥ 1 vì mẫu không vượt số còn lại
        for _ in range(rep):
            b = Bar(stock_mm=s)
            for i, n in sorted(pat.items(), key=lambda x: -keys[x[0]][0]):
                take = min(n, remain[i])
                b.cuts += [keys[i][0]] * take
                b.labels += [keys[i][1]] * take
                remain[i] -= take
            bars.append(b)
    # thu nhỏ: cây nào thừa nhiều mà các đoạn của nó vừa một cây ngắn hơn thì đổi sang cây ngắn hơn
    for b in bars:
        for s in stocks:
            if s < b.stock_mm and b.used(kerf) + trim <= s:
                b.stock_mm = s
                break
    # cận dưới: tổng chiều dài đoạn + mạch cưa (không phương án nào mua ít hơn số này); 1 loại cây thì làm tròn lên số cây
    lb = sum((ln + kerf) * q for (ln, _), q in need.items())
    if len(stocks) == 1 and lb:
        lb = math.ceil(lb / (stocks[0] - trim + kerf)) * stocks[0]
    bought = sum(b.stock_mm for b in bars)
    return {"bars": bars, "ok": True, "too_long": [], "reason": None, "lower_bound_mm": lb, "bought_mm": bought}


def summarize_bars(res: dict, kerf: int, min_remnant: int, trim: int = 0) -> dict:
    """Gom kết quả cắt thành: số cây theo chiều dài, tổng mua, phế, đoạn dư dùng lại được, sơ đồ cắt."""
    bars: list[Bar] = res["bars"]
    by_len: dict[int, int] = {}
    waste = remnant = used = 0
    plan = []
    for n, b in enumerate(bars, 1):
        by_len[b.stock_mm] = by_len.get(b.stock_mm, 0) + 1
        u = b.used(kerf)
        left = b.stock_mm - u - trim
        used += sum(b.cuts)
        if left >= min_remnant:
            remnant += left
        else:
            waste += max(0, left)
        plan.append({"no": n, "stock_mm": b.stock_mm, "cuts": b.cuts, "labels": b.labels, "left_mm": left,
                     "reusable": left >= min_remnant})
    bought = sum(b.stock_mm for b in bars)
    return {"count_by_len": {str(k): v for k, v in sorted(by_len.items())}, "bars": len(bars), "bought_mm": bought,
            "used_mm": used, "waste_mm": bought - used - remnant, "remnant_mm": remnant,
            "waste_pct": round((bought - used - remnant) / bought * 100, 2) if bought else 0.0, "plan": plan,
            "lower_bound_mm": res.get("lower_bound_mm", 0)}


# ------------------------------------------------------------------ 2 chiều: tấm
def pack_sheets(pieces: list[tuple[int, int, int, str]], sheet_w: int, sheet_l: int, gap: int = 5) -> dict:
    """pieces = [(dài, rộng, số lượng, nhãn)] mm. Xếp theo kệ (FFDH), cho xoay. Trả {sheets: [[(x, y, l, w, nhãn)]], too_big: [...]}."""
    items: list[tuple[int, int, str]] = []
    too_big = []
    for a, b, q, lab in pieces:
        lo, sh = max(a, b), min(a, b)
        if lo > sheet_l or sh > sheet_w:
            too_big.append(lab)
            continue
        items += [(lo, sh, lab)] * int(q)
    # đặt cạnh dài theo chiều dài tấm, kệ chạy dọc chiều rộng: mỗi kệ có chiều cao = cạnh ngắn của chi tiết đầu kệ
    items.sort(key=lambda t: (-t[1], -t[0]))
    sheets: list[list] = []
    shelves: list[list] = []      # mỗi kệ: [chỉ số tấm, y, cao, x đang dùng]
    used_h: list[int] = []        # chiều cao đã dùng của từng tấm
    for lo, sh, lab in items:
        placed = False
        for s in shelves:
            si, y, h, x = s
            if sh <= h and x + lo <= sheet_l:
                sheets[si].append((x, y, lo, sh, lab))
                s[3] = x + lo + gap
                placed = True
                break
        if placed:
            continue
        for si in range(len(sheets)):                       # mở kệ mới trên tấm đã có
            if used_h[si] + sh <= sheet_w:
                y = used_h[si]
                sheets[si].append((0, y, lo, sh, lab))
                shelves.append([si, y, sh, lo + gap])
                used_h[si] = y + sh + gap
                placed = True
                break
        if not placed:
            sheets.append([(0, 0, lo, sh, lab)])
            used_h.append(sh + gap)
            shelves.append([len(sheets) - 1, 0, sh, lo + gap])
    return {"sheets": sheets, "too_big": sorted(set(too_big))}
