import openpyxl, os
from openpyxl.styles import Alignment
OUT = os.path.dirname(os.path.abspath(__file__))

# ---------------- Technical BOM ----------------
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "BOM BAN"
ws["A1"] = "BẢNG ĐỊNH MỨC KỸ THUẬT XƯỞNG"
ws.merge_cells("A1:P1")
ws["A2"] = "Sản phẩm: Bàn thép"
top = ["MỤC", "MÃ CHI TIẾT", "MÔ TẢ", "VẬT LIỆU", "KÍCH THƯỚC (mm)", None, None, "SỐ LƯỢNG",
       "KHỐI LƯỢNG NGUYÊN LIỆU (kg)", "MÃ NVL", "ĐVT", "HAO HỤT (%)", "CÔNG ĐOẠN", None, None, None, "ĐƠN GIÁ"]
sub = [None, None, None, None, "Dày", "Rộng", "Dài", None, None, None, None, None, "Cắt CNC", "Chấn", "Hàn", "Sơn", None]
for i, v in enumerate(top, 1):
    if v: ws.cell(4, i, v)
for i, v in enumerate(sub, 1):
    if v: ws.cell(5, i, v)
for i, v in enumerate(top, 1):
    if v and not sub[i - 1] and v not in ("KÍCH THƯỚC (mm)", "CÔNG ĐOẠN"):
        ws.merge_cells(start_row=4, start_column=i, end_row=5, end_column=i)
ws.merge_cells("E4:G4"); ws.merge_cells("M4:P4")
for i in range(1, 18):
    ws.cell(6, i, f"({i})")
rows = [
    ["1", "FG-TABLE-01", "Bàn thép hoàn chỉnh", None, None, None, None, 1, None, None, "Cái", None, None, None, None, "x", None],
    ["1.1", "KHUNG-01", "Khung bàn hàn", None, None, None, None, 1, None, None, "Bộ", None, None, None, "x", None, None],
    ["1.1.1", "CHAN-01", "Chân bàn thép ống", "SS400", 2, 32, 700, 4, 1.725, "R-O-S01-00-032-001", "Cái", 0, "x", None, None, None, None],
    ["1.1.2", "MAT-01", "Mặt bàn thép tấm", "SS400", 8, 600, 1200, 1, 45.2, None, "Cái", 5, "x", "x", None, None, None],
    ["1.1.3", "BAS-01", "Bas thép tấm", "SS400", 8, 50, 50, 4, 0.157, None, "Cái", None, "x", None, None, None, None],
    ["1.2", None, "Bulong M10x30", None, None, None, None, 8, None, "R-B-M10", "Bộ", None, None, None, None, None, 2500],
    ["1.3", "TRUC-01", "Trục thép tròn đặc", "S45C", None, 20, 300, 2, 0.74, None, "Cái", None, "x", None, None, None, None],
    ["1.4", "BAS-01", "Bas thép tấm (dùng chung)", "SS400", 8, 50, 50, 2, 0.157, None, "Cái", None, "x", None, None, None, None],
    [None, None, "Tổng cộng", None, None, None, None, None, None, None, None, None, None, None, None, None, None],
]
for r, row in enumerate(rows, 7):
    for c, v in enumerate(row, 1):
        if v is not None: ws.cell(r, c, v)

ws2 = wb.create_sheet("FG-CHAIR-02")
ws2["A1"] = "ĐỊNH MỨC GHẾ"
hdr = ["STT", "Mã chi tiết", "Tên chi tiết", "Vật liệu", "Dày", "Rộng", "Dài", "SL", "KL NVL (kg)", "Mã vật tư", "Công đoạn"]
for i, v in enumerate(hdr, 1): ws2.cell(3, i, v)
data2 = [
    [1, "FG-CHAIR-02", "Ghế thép", None, None, None, None, 1, None, None, None],
    ["1.1", "P-PLATE", "Tấm đế thép tấm", "SS400", 5, 400, 400, 1, 6.27, "R-P-S01-052-00-002", "Cắt CNC, Chấn"],
    ["1.2", "P-LEG", "Chân ghế thép hộp", "SS400", 1.2, 25, 25, 4, 0.42, None, "Cắt CNC - Hàn"],
    ["1.3", "P-PIPE", "Tay vịn ống", "SS400", 1.4, 21, 600, 2, 0, None, "Cắt CNC"],
]
for r, row in enumerate(data2, 4):
    for c, v in enumerate(row, 1):
        if v is not None: ws2.cell(r, c, v)
# indent-based sub child of P-LEG without dotted MỤC -> flat (all level 1) on this sheet
wb.save(os.path.join(OUT, "sample_bom.xlsx"))

# ---------------- SO ----------------
wb = openpyxl.Workbook(); ws = wb.active; ws.title = "SO"
ws["A1"] = "SALES ORDER"; ws.merge_cells("A1:I1")
h = ["STT", "Mã SP Khách hàng", "Mã SP ERP", "Tên thành phẩm", "Quy cách", "ĐVT", "Số lượng", "Giá FOB (USD)", "Thành tiền (USD)"]
for i, v in enumerate(h, 1): ws.cell(3, i, v)
for r, row in enumerate([[1, "CUST-T1", "FG-TABLE-01", "Bàn thép", "1200x600", "Cái", 125, 85.5, None],
                         [2, "CUST-C2", "FG-CHAIR-02", "Ghế thép", "400x400", "Cái", 2000, 22, None],
                         [None, None, None, "Tổng cộng", None, None, 2125, None, None]], 4):
    for c, v in enumerate(row, 1):
        if v is not None: ws.cell(r, c, v)
wb.save(os.path.join(OUT, "sample_so.xlsx"))

# ---------------- PO ----------------
wb = openpyxl.Workbook(); ws = wb.active; ws.title = "PO"
h = ["Số PO", "Nhà cung cấp", "Mã vật tư", "Tên hàng hóa", "ĐVT", "Số lượng", "Đơn giá (VND)", "Ngày hẹn giao"]
for i, v in enumerate(h, 1): ws.cell(1, i, v)
for r, row in enumerate([["PO-001", "Hòa Phát", "R-O-S01-00-032-001", "Thép ống phi 32x2", "Kg", 900, 21000, "15/10/2026"],
                         ["PO-002", "Pomina", "R-P-S01-052-00-002", "Thép tấm 5ly", "Kg", 13000, 19500, "12/10/2026"],
                         ["PO-003", "Hòa Phát", "", "Thép tấm 8.0", "Kg", 7000, 20500, "12/10/2026"]], 2):
    for c, v in enumerate(row, 1):
        ws.cell(r, c, v)
wb.save(os.path.join(OUT, "sample_po.xlsx"))
print("ok")
