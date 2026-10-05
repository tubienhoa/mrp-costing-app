BẢN ĐẶC TẢ KỸ THUẬT VÀ PHÂN TÍCH LOGIC PHẦN MỀM

DỰ ÁN: METALIC ERP COSTING \& MRP AUTOMATION PORTAL

Mục Tiêu Cốt Lõi:	Tự động hóa Hoạch định Nhu cầu Vật tư (Multi-Level Technical BOM MRP) \& Hạch toán Báo cáo Giá thành / P\&L Dự án Cơ khí Xuất khẩu với 100% công thức liên kết sống Excel.



1\. TỔNG QUAN KIẾN TRÚC VẬN HÀNH \& TỆP ĐẦU VÀO

Hệ thống được thiết kế theo cơ chế Clean Slate Startup (Màn hình sạch) và tiếp nhận tối đa 3 tệp thô đầu vào (định dạng .xlsx hoặc .csv):

Tệp Đầu Vào	Tính Bắt Buộc	Mục Đích \& Nội Dung

1\. Sales Order (SO)	BẮT BUỘC	Sản lượng đơn hàng \& Giá bán FOB ($USD) theo từng mã SKU Thành phẩm.

2\. Technical BOM	BẮT BUỘC	Cấu trúc Kỹ thuật xưởng đa cấp, tỷ lệ tiêu hao NVL \& công đoạn gia công.

3\. Purchase Order (PO)	TÙY CHỌN	Sổ tra cứu giá mua thực tế, số PO và nhà cung cấp vật tư.



Nguyên tắc Vận hành Độc lập (Operating Policy):

•	Khởi động ứng dụng: Khi chưa có file tải lên, ứng dụng hiển thị trạng thái chờ và thông báo yêu cầu nạp tối thiểu 2 tệp bắt buộc (SO + Technical BOM).

•	Kịch bản 2 Tệp (SO + Technical BOM): Hệ thống lập tức kích hoạt bộ máy rã BOM đa cấp, tự động bóc tách danh mục Nguyên vật liệu thô (NVL), tính toán Nhu cầu MRP. Do không có PO, đơn giá vật tư mặc định bằng 0 (nguồn giá hiển thị "Chưa nạp PO (0đ)") và TK 621 = 0; người dùng có thể nhập giá bổ sung qua Live BOM Editor.

•	Kịch bản 3 Tệp (SO + Technical BOM + PO): Hệ thống dùng danh mục vật tư rã từ Technical BOM làm gốc, đối chiếu sang file PO theo Mã vật tư ERP hoặc Tên vật tư chuẩn. Dòng khớp: lấy Đơn giá mua thực tế và Mã PO thực tế (nguồn giá "PO: <Số PO>"). Dòng không khớp: đơn giá = 0 (nguồn giá "Chưa có giá PO (0đ)").

2\. CẤU TRÚC BOM KỸ THUẬT \& THUẬT TOÁN CHUYỂN ĐỔI SANG BOM ERP

2.1. Nhận diện Cây Cấu trúc Phân cấp (Tree Hierarchy): Căn cứ vào cột MỤC (1, 1.1, 1.1.1 hoặc khoảng lùi đầu dòng) để xác định mối quan hệ Cha - Con đệ quy.

2.2. Quy trình Công đoạn \& Sinh Mã BTP Trung Gian: Khi kiểm tra các cột công đoạn (Cắt CNC, chấn, hàn, sơn...) tại dòng Mã Cha, qua mỗi công đoạn được đánh dấu, hệ thống tự động khởi tạo 1 Mã BTP trung gian quản lý dở dang (WIP) theo cú pháp: {Mã\_Cha}.BTP\_{Tên\_Công\_Đoạn}.

2.3. Tự Động Định Danh Nguyên Vật Liệu Thô: Trích xuất tên NVL thô chuẩn từ MÔ TẢ và kích thước: Thép tấm {Độ dày}, Thép tròn đặc phi {Chiều rộng}, Thép hộp {Dài}x{Rộng}x{Dày}, Thép ống phi {Rộng}x{Dày}x{Dài}.

2.4. Đơn Vị Tính \& Trọng Lượng NVL: Nếu cột KHỐI LƯỢNG NGUYÊN LIỆU (kg) > 0, hệ thống tự động chốt ĐVT là 'Kg' và dùng làm định mức tiêu hao kỹ thuật cho 1 chi tiết.

3\. CẤU TRÚC DỮ LIỆU SO \& PO (SALES ORDER \& PURCHASE ORDER SCHEMAS)

Mọi tệp SO và PO tải lên được quét qua Smart Header Parser để chuẩn hóa về cấu trúc chuẩn:



SO\_SCHEMA = {

&#x20;   "STT": "int",

&#x20;   "Mã Sản Phẩm KH": "str",

&#x20;   "Mã Sản Phẩm ERP": "str",        # Cột khóa chính (FG Code)

&#x20;   "Tên Thành Phẩm": "str",

&#x20;   "Quy Cách": "str",

&#x20;   "ĐVT": "str",

&#x20;   "Số Lượng": "int",               # Bắt buộc > 0

&#x20;   "Giá Bán FOB ($USD)": "float"

}



PO\_SCHEMA = {

&#x20;   "Số PO Mua": "str",

&#x20;   "Nhà Cung Cấp": "str",

&#x20;   "Mã Vật Tư ERP": "str",          # Cột khóa chính tra cứu (Material Code)

&#x20;   "Tên Hàng Hóa Phụ Kiện Mua": "str",

&#x20;   "ĐVT": "str",

&#x20;   "Số Lượng Mua Thực Tế": "float",

&#x20;   "Đơn Giá Mua (VND)": "float",

&#x20;   "Ngày Hẹn Giao": "str"

}

4\. THUẬT TOÁN HOẠCH ĐỊNH MRP ĐỆ QUY ĐA CẤP

Cấu trúc BOM được biểu diễn dưới dạng đồ thị có hướng (Directed Acyclic Graph - DAG) với Node gốc (Thành phẩm FG), Node trung gian (BTP SG) và Node lá (Raw Material RM).

Công thức Tỷ lệ tiêu hao tích lũy nhân dồn (Cumulative Usage Rate):



R\_cum(RM\_k, FG\_i) = SUM\_{P in Path} \[ PRODUCT\_{(Parent, Child) in P} Rate(Parent, Child) \* (1 + Scrap(Parent, Child) / 100) ]

Công thức Tổng nhu cầu MRP tuyệt đối (Gross MRP Requirement):



Q\_MRP(RM\_k) = SUM\_{FG\_i in SO} \[ Q\_SO(FG\_i) \* R\_cum(RM\_k, FG\_i) ]

5\. BỘ MÁY HẠCH TOÁN GIÁ THÀNH \& P\&L QUẢN TRỊ

Bộ máy tài chính thực hiện hạch toán chi tiết chi phí sản xuất theo 3 tài khoản kế toán chi phí (TK 621, TK 622, TK 627) và tính toán P\&L theo thời gian thực:

Chi Phí NVL Trực Tiếp (TK 621): TK 621 = SUM \[ Q\_MRP(RM\_k) \* UnitPrice(RM\_k) ], trong đó UnitPrice = Đơn giá PO nếu RM_k khớp PO; ngược lại = 0 (không có PO: = giá nhập tại Live BOM Editor nếu có, còn lại 0). Hệ thống không tự động dùng giá mặc định; giá tham khảo ở Sidebar chỉ được áp khi người dùng chủ động bật "Áp giá tạm từ Sidebar".

Chi Phí Nhân Công (TK 622): TK 622 = SUM\_{Stage} \[ Total SO Qty \* Labor Rate\_{Stage} ]

Chi Phí Sản Xuất Chung (TK 627): TK 627 = SUM\_{Stage} \[ Total SO Qty \* Overhead Rate\_{Stage} ]

Tổng Giá Vốn Hàng Bán (COGS): COGS = TK 621 + TK 622 + TK 627

Doanh Thu Thuần (VND): Revenue = Total SO Qty \* FOB USD \* Exchange Rate (B6)

Lợi Nhuận Gộp (Gross Profit): Gross Profit = Revenue - COGS

Chi Phí Quản Lý \& Bán Hàng (OPEX): OPEX = Revenue \* (% SG\&A)

Lợi Nhuận Thuần (NPAT): NPAT = MAX(0, EBIT \* (1 - % Thuế TNDN))

6\. KIẾN TRÚC BỘ KẾT XUẤT EXCEL MASTER 25 SHEET LIÊN THÔNG

Yêu Cầu Tối Thượng: 100% CÔNG THỨC LIÊN KẾT SỐNG (NO HARDCODED VALUES). Tuyệt đối không ghi số tĩnh vào các ô tính toán. Mọi ô kết quả phải chứa công thức Excel sống liên kết từ Tỷ giá (B6) và Sản lượng SO.

Xử lý MergedCell trong OpenPyXL bằng hàm bắt buộc safe\_set\_cell():



def safe\_set\_cell(ws, row\_or\_coord, col=None, value=None):

&#x20;   try:

&#x20;       cell = ws\[row\_or\_coord] if isinstance(row\_or\_coord, str) else ws.cell(row=row\_or\_coord, column=col)

&#x20;       if type(cell).\_\_name\_\_ == "MergedCell":

&#x20;           for rng in ws.merged\_cells.ranges:

&#x20;               if cell.coordinate in rng:

&#x20;                   top\_left = ws.cell(row=rng.min\_row, column=rng.min\_col)

&#x20;                   top\_left.value = value

&#x20;                   return top\_left

&#x20;           return None

&#x20;       cell.value = value

&#x20;       return cell

&#x20;   except Exception:

&#x20;       return None

7\. DANH SÁCH 25 SHEET TRONG FILE EXCEL MASTER

Tên Sheet Excel	Mô Tả Chức Năng \& Công Thức Liên Kết

01\. Thu muc \& Thuat ngu	Danh mục hệ thống \& Thuật ngữ ERP

02\. Bao gia \& Don hang	Bảng SO, ô B6 chứa tỷ giá, cột Thành tiền =G9\*H9\*$B$6

03\. KH Giao hang \& Lenh SX	Lệnh sản xuất (WO), công thức kế thừa từ Sheet 02

04\. Hoach dinh MRP	Bảng rã MRP, công thức nhu cầu ='02. SO'!$G$9 \* Rate

05\. Lich ra chi tiet DPP	Lịch phân rã công đoạn dở dang WIP

06\. De nghi Mua hang PR	Chứng từ PR kế thừa từ Sheet 04 MRP

07\. So sanh Gia NCC	So sánh báo giá nhà cung cấp

08\. Don dat mua hang PO	Bảng PO mua hàng, đơn giá tra cứu từ PO/BOM

09 - 18. Chứng từ Kho \& QA/QC	IQC, GRN, OUT, WIP, FQC, FGRN, DO, GDN, AR, Pallet ID

19\. Pha he Truy vet Nguoc	Ma trận truy vết ngược 360 độ từ Invoice về Lot NVL

Tong hop Gia thanh	Bảng tính giá thành sản xuất đích danh SKU

TK 621, 622, 627	Bảng hạch toán chi tiết các tài khoản chi phí

24\. Bao cao P\&L (KQKD)	Báo cáo kết quả hoạt động kinh doanh dự án

8\. DANH SÁCH KIỂM THỬ VÀ NGHIỆM THU (ACCEPTANCE CHECKLIST)

•	 \[X] Khởi động Nhanh: Ứng dụng khởi chạy mượt màn hình sạch khi chưa nạp file.

•	 \[X] Nạp 2 Tệp (SO + Technical BOM): Hệ thống mở khóa tính toán 100%, không đòi hỏi file PO.

•	 \[X] Chuyển đổi BOM Kỹ thuật: Đọc chính xác file BOM, tự động định danh chuỗi 'Thép tấm 8.0', 'Thép tròn đặc phi 20'...

•	 \[X] Tính toán MRP Chính xác: Kết quả MRP khớp đúng tuyệt đối với số liệu phân xưởng (R-O-S01-00-032-001 = 862.50 kg, R-P-S01-052-00-002 = 12,540.00 kg).

•	 \[X] Xuất Excel 25 Sheet Công Thức Sống: File .xlsx tải về giữ nguyên toàn bộ 25 Sheet, ô B6 chứa tỷ giá hối đoái, mọi ô tính toán đều chứa công thức liên kết Excel (=...).

•	 \[X] Không Bị Lỗi Ô Gộp: Sử dụng hàm safe\_set\_cell() ghi file thành công mà không bị crash ứng dụng.



