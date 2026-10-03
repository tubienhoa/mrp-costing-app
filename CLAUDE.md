# CLAUDE.md – METALIC ERP COSTING & MRP AUTOMATION PORTAL

> Cập nhật lần cuối: 2026-10-02. Đọc file này trước khi sửa `app.py`.
> Đặc tả gốc: `PRD_SPEC.md` (nguồn sự thật cho công thức và danh sách sheet).

## 1. Bối cảnh dự án

- Ứng dụng **Streamlit một file** (`app.py`) cho công ty cơ khí xuất khẩu **Metalic**.
- Chức năng: rã Technical BOM đa cấp → BOM ERP (FG / SG / BTP / RM) → MRP → hạch toán TK 621/622/627 → P&L → xuất **Excel Master 25 sheet, 100% công thức sống**.
- Đầu vào: **SO** (bắt buộc), **Technical BOM** (bắt buộc), **PO** (tùy chọn); định dạng `.xlsx/.xlsm/.csv`.
  - Kịch bản 2 tệp: giá NVL lấy từ BOM → Live BOM Editor → Sidebar.
  - Kịch bản 3 tệp: danh mục NVL lấy từ BOM làm gốc, đối chiếu sang PO để lấy đơn giá, số PO và NCC.
- Chưa nạp đủ SO + BOM → màn hình chờ (`clean_slate()`), gọi `st.stop()`.
- Môi trường: Windows 11, Python 3.14, Streamlit 1.63, pandas 3.0, openpyxl 3.1. Chạy app: `streamlit run app.py`.
- Dự án **chưa dùng git**.

### File dữ liệu thật của người dùng (ngoài repo)

| File | Đường dẫn | Trạng thái |
|---|---|---|
| SO Savic | `C:\Users\MyPC\Desktop\ERP_App\Costing phien ban ky thuat\Savic\template_don_hang_so.xlsx` | ✅ Đọc đúng |
| BOM Savic | `C:\Users\MyPC\Desktop\ERP_App\Costing phien ban ky thuat\Savic\BOM - Savic -dev.xlsx` | ✅ Đọc đúng (sau 5 bản sửa ở mục 4) |
| BOM Savic (bản ERP) | `C:\Users\MyPC\Downloads\zalo_file\BOM - Savic - ERP.xlsx` | ⏳ Chưa thử |
| Jack Pad (SO + BOM) | `C:\Users\MyPC\Desktop\ERP_App\Costing\Jack pad\` (`BOM_Jack Pad_New.xlsx`, `BOM - ATW - Jack Pad.xlsx`, `template_don_hang_so_ Jack Pad.xlsx`) | ⏳ Chưa thử |
| GP Orcana + PO | `C:\Users\MyPC\Desktop\ERP_App\Costing\` (`BOM - GP Orcana - M2608-Lot-14.xlsx`, `Mua hang PO 8201.xlsx`, `template_don_hang_so - Mẫu.xlsx`, bản `.csv` trong `CSV\`) | ⏳ Chưa thử |

## 2. Kiến trúc `app.py`

Các phần được đánh số trong comment. Số dòng thay đổi theo thời gian, nên tra theo **tên hàm**.

| # | Phần | Hàm / hằng chính |
|---|---|---|
| 1 | Tiện ích chuẩn hóa | `norm()` (bỏ dấu, đ→d, chữ thường), `to_float()` (xử lý `1.234,5` / `1,234.5`), `to_str()`, `ncode()` (bỏ khoảng trắng, viết hoa – dùng để so khớp mã), `slug()`, `fmt_dim()`, `fmt_plate()`, `truthy_mark()` |
| 2 | Ghi ô gộp | `safe_set_cell()` – **bắt buộc theo đặc tả**: mọi lệnh ghi Excel đều phải đi qua hàm này (`XLWriter.put` gọi nó) |
| 3 | Đọc file thô | `read_grids()` → mỗi sheet trả về `raw` (chỉ ô thật), `filled` (lan giá trị ô gộp), `indent`, `pct` (ô định dạng %). `_read_csv_grid()` dò dấu phân cách và encoding |
| 4 | Smart Header Parser | `SO_SPEC`, `PO_SPEC`, `BOM_SPEC` = danh sách `(field, [regex ưu tiên], regex loại trừ)` so trên tiêu đề đã `norm()`. `map_columns()`, `detect_header()` (tiêu đề 1 hoặc 2 tầng + `muc_cols`), `detect_stage_columns()` (`STAGE_RE` / `STAGE_EXCL`), `_row_is_index()` (bỏ hàng (1)(2)(3)), `FOOTER_RE` |
| 5 | Parser | `parse_so()`, `parse_po()`, `parse_bom()` – có `@st.cache_data`. `parse_bom` trả về `(df, stage_cols, sheet_ctx, infos, warns)`. Cột df: `BOM_BASE_COLS` + các cột công đoạn `"CĐ | <tên>"` (**kiểu float = số thứ tự công đoạn, 0 = không làm**) |
| 6 | Tên NVL chuẩn | `std_material_name(desc, material, t, w, l, w2)` |
| 7 | Bộ máy | `run_engine()` → dict `nodes, edges, children, fg_cum, so_lines, rm_used, R, mrp_df, wip, warnings, used_stages`. `compute_financials()` tính giá thành SKU và P&L |
| 8 | Excel 25 sheet | `SH` (tên sheet), `XLWriter`, `build_master_excel()`, `count_formulas()` |
| 9 | Giao diện | `main()`: sidebar (upload + tham số `P` + bảng đơn giá công đoạn); 7 tab: Tổng quan / SO & PO / Live BOM Editor / BOM ERP / MRP / Giá thành & P&L / Xuất Excel. Tạo tab trước → render Editor → chạy engine → điền các tab còn lại |

### Quy tắc chính của bộ máy (`run_engine`)

- **Cây cha–con:** dựng theo thứ tự dòng + cột `Cấp`, dùng stack (không dựa vào tiền tố chuỗi MỤC, vì Excel làm mất "1.10" → 1.1).
  - Cách tính `Cấp`: MỤC dạng chấm `1.1.1` → số phần; MỤC chữ / La Mã → 0 (nhóm); MỤC trống khi file có dạng chấm → cấp số gần nhất + 1. Không có MỤC dạng chấm → dùng khoảng lùi đầu dòng (`alignment.indent`, dấu cách).
- **Ghép FG giữa SO và BOM** (`match_fg`), theo thứ tự:
  1. trùng Mã ERP;
  2. trùng Mã KH;
  3. Mã ERP có dạng `<tiền tố> - <mã BOM>`;
  4. cột `Mã TP` trong BOM;
  5. tên sheet / vùng tiêu đề chứa mã SO;
  6. SO chỉ có 1 FG và BOM không khớp gì → gán cả BOM cho FG đó;
  7. nếu vẫn không khớp → `NO-FG-<sheet>` kèm cảnh báo.
- **Phân loại dòng lá:**
  - Mã bắt đầu bằng tiền tố NVL (mặc định `R-`, đổi ở Sidebar) → RM, định mức = SL × (KL kg nếu > 0, ngược lại 1).
  - Dòng có Mã NVL / KL > 0 / nhận diện được hình dạng thép → chi tiết SG + một con RM.
  - Còn lại → vật tư mua ngoài (RM).
- **Mã BTP:** mỗi công đoạn được đánh dấu tại dòng cha sinh `{Mã}.BTP_{SLUG}`, nối theo **số thứ tự công đoạn**. Ô ghi `x` thì xếp theo vị trí cột; nếu bằng nhau thì theo vị trí cột. Chuỗi: `X ← BTP_cuối ← … ← BTP_đầu ← (con + RM)`, mỗi cạnh có định mức 1.
- **Hao hụt** áp lên cạnh tiêu hao RM (nếu dòng có RM), ngược lại áp lên cạnh cha→dòng. Hiệu lực: `eff = rate × (1 + scrap/100)`.
- **DAG:** một mã xuất hiện nhiều lần chỉ được rã cấu trúc con ở lần đầu (`expanded[code] = occurrence`), nên không bị nhân đôi nhu cầu. Có phát hiện vòng lặp.
- **R_cum:** `cum(node)` đệ quy có memo, cộng qua mọi đường đi. `Q_MRP = Σ Q_SO × R_cum`.
- **Đơn giá:** PO (khớp mã, rồi khớp tên chuẩn / tên) > `Đơn giá` trong BOM / Editor > Sidebar (`kg_price` cho ĐVT Kg, `unit_price` cho ĐVT khác). Giá = 0 → có cảnh báo.
- **TK 622 / 627** = Tổng SL SO × Σ đơn giá các công đoạn **có xuất hiện trong BOM** (đúng công thức đặc tả).
- **NPAT** = `MAX(0, EBIT × (1 − %thuế))` – theo đúng đặc tả. Lỗ thì hiện 0 và có cảnh báo.

### Excel Master (`build_master_excel`)

- Thứ tự 25 sheet: 01 → 25. Sheet 25 là `25. BOM ERP Da cap`, do tôi thêm vì đặc tả chỉ liệt kê 24 sheet.
- Bố cục chuẩn: hàng 1–3 là tiêu đề gộp, **hàng 8 là header, dữ liệu bắt đầu từ hàng 9**. Dòng NVL thứ i nằm ở hàng `9+i` trong các sheet 04/06/07/08/09/10/11/21; dòng SO thứ j nằm ở hàng `9+j` trong các sheet 02/03/13–18/20.
- `'02. Bao gia & Don hang'!$B$6` = tỷ giá; `I = G*H*$B$6`; dòng tổng ở `9+nso`.
- Sheet 04: `E = Σ '02'!$G$(9+j) * <ô R_cum>`; ma trận R_cum (cột G trở đi) ghi giá trị số, tô xanh lá (là định mức kỹ thuật).
- Màu ô: vàng = dữ liệu nhập; xanh lá = định mức; còn lại = công thức. `wb.calculation.fullCalcOnLoad = True`.
- **Bất biến:** không ghi số tĩnh vào ô kết quả. Mọi ô kết quả phải là công thức `=...`.

## 3. Các file trong dự án

- `app.py` – toàn bộ ứng dụng.
- `PRD_SPEC.md` – đặc tả gốc.
- `requirements.txt` – streamlit, pandas, openpyxl.
- `tests/` – script kiểm thử, chạy bằng `python <file>` **từ trong thư mục `tests/`**:
  - `make_samples.py` – sinh `sample_so/bom/po.xlsx` (BOM tiêu đề 2 tầng, ô gộp, nút dùng chung, sheet ghép FG theo ngữ cảnh).
  - `test_engine.py` – chạy parser + engine + xuất `master.xlsx` trên dữ liệu mẫu (`--nopo` = kịch bản 2 tệp).
  - `test_indent.py` – SO dạng CSV + BOM phân cấp bằng khoảng lùi đầu dòng.
  - `diag_savic.py` – chẩn đoán trên file Savic thật (`-e` để in cạnh BOM ERP và P&L).
  - `verify_xlsx.py`, `verify_savic.py` – tính lại công thức Excel bằng thư viện `formulas` và so với Python. Cần chạy `pip install formulas` (hoặc `--target tests/pylib`). Lỗi `#NAME?` ở `HYPERLINK` (sheet 01) là do thư viện không hỗ trợ hàm này, Excel vẫn chạy bình thường.
  - `ui_test.py` – Streamlit `AppTest` cho các kịch bản 0 (clean slate) / 2 / 3 / S (Savic). Giả lập `st.file_uploader`.
  - `dump.py <file.xlsx> [maxrows]` – in giá trị, công thức, ô gộp, khoảng lùi của từng sheet. **Luôn chạy script này trước khi sửa parser cho một file mới.**

## 4. Quy tắc đọc file thật – 5 lỗi cấu trúc đã giải quyết (BOM Savic, sheet `Form`)

| # | Triệu chứng | Nguyên nhân | Cách xử lý trong code |
|---|---|---|---|
| 1 | Tiêu đề thành `"Bill of Materials \| CHIỀU DÀI"`; dòng `mm/inch` bị đọc như dữ liệu; CHIỀU DÀI bị gán thành "Vật liệu" | Ô gộp tiêu đề `H1:AD4` bị ghép với dòng tiêu đề thật ở dòng 5 | `detect_header`: chỉ ghép 2 tầng khi bản thân dòng trên khớp ≥ 2 cột; khi điểm bằng nhau trên cùng dòng thì **ưu tiên 2 tầng**. Các trường kích thước loại cột `inch`; `material` loại `bill of` |
| 2 | Cây phẳng, mọi chi tiết rơi vào `NO-FG-FORM` → **MRP trống** | MỤC gộp `A5:E6`: **cấp = vị trí cột** (A7=1 là FG, B8=1 là chi tiết cấp 2) | `detect_header` trả `muc_cols` (các cột liền kề cùng tiêu đề MỤC). `parse_bom` dựng chuỗi `1`, `1.1` … `2.10` từ cột đầu tiên có giá trị |
| 3 | FG không khớp | SO: `FG - SVC - SAV01829`; BOM: `SAV01829` (= Mã SP KH) | `match_fg`: khớp Mã KH và hậu tố `-<mã BOM>` của mã ERP (sau khi `ncode`) |
| 4 | Thứ tự BTP sai; "DIỆN TÍCH BỀ MẶT SƠN" bị coi là công đoạn Sơn | Ô công đoạn chứa **số thứ tự** (Cắt=1, Chấn=2; FG: Hàn 1 → Xi mạ 2 → Đóng gói 3), không phải `x` | Cột công đoạn lưu dạng float; `row_stages()` sắp theo (số thứ tự, vị trí cột). `STAGE_EXCL` loại `dien tich, m2, tong, hoan thien, be mat` |
| 5 | Cộng thừa 1% hao hụt | "TỈ LỆ HAO HỤT" = **hệ số** (1, 0.95…), đã nhân sẵn trong công thức KL nguyên liệu `=H*J*N*Q*10^-6*R` | `parse_bom`: tiêu đề không có `%`, mọi giá trị thuộc (0, 1] và ô không định dạng % → chế độ hệ số. Có cột KL → hao hụt = 0; không có cột KL → `(1/f − 1) × 100` |

Các quy tắc khác:
- Tiêu đề SO Savic ở dòng 4: `Mã Thành Phẩm ERP`, `Số Lượng (Cái)`, `Đơn Giá (USD)` (được nhận là FOB). Các dòng STT 3–15 không có mã sẽ bị bỏ qua.
- BOM Savic: `KHỐI LƯỢNG NGUYÊN LIỆU (kg)` (cột T) là **khối lượng / 1 chi tiết** → định mức = SL × T. Cột U (`TỔNG KL`) và các cột KL thành phẩm **không dùng**.
- Sheet `Data` (bảng tra tỷ trọng) tự bị bỏ qua vì không có tiêu đề BOM. Vùng tính phụ ở dòng 30–35 và dòng `TOTAL` cũng bị bỏ qua.
- Ô `x` trong cột "Phay tiện (MC)" được xếp theo vị trí cột, tức sau Cắt (=1).

### Mốc kiểm chứng (phải giữ đúng sau mọi lần sửa)

- **Savic (SO 200 + 200 bộ): tổng MRP = 30,896.59 kg** = ô `U28` (154.483 kg/bộ, cả 2 FG) × 200. Chi tiết:
  - Thép tấm 8.0 = 19,755.20
  - Thép tấm 4.0 = 5,653.86
  - Thép hộp 80x80x3 = 3,264.03
  - Thép tấm 6.0 = 1,783.11
  - Thép tấm 10.0 = 291.47
  - Thép tròn đặc phi 32 = 148.92
  - 6 mã RM, 33 mã BTP, 67 cạnh BOM ERP.
- Excel Savic tính lại khớp Python: TK 621 = 679,725,036; giá thành/SP 1,813,751 và 1,764,875 (với NC 10,000 / SXC 5,000 mỗi công đoạn, giá 22,000/kg).
- Dữ liệu mẫu (`tests/`): `R-O-S01-00-032-001 = 862.50 kg`, `R-P-S01-052-00-002 = 12,540.00 kg` (mốc nghiệm thu trong đặc tả, hằng `ACCEPTANCE`). ⚠️ Đây là dữ liệu **tự dựng** cho khớp mốc. File thật chứa 2 mã này **chưa được xác định** (có thể là Jack Pad hoặc GP Orcana).

## 5. Quy tắc đặt tên NVL và giá mặc định

- `std_material_name()` (đặc tả mục 2.3), nhận diện hình dạng từ `MÔ TẢ` + `Vật liệu` (đã `norm`), theo thứ tự:
  1. **hộp** → `Thép hộp {A}x{B}x{T}` nếu có cột `Rộng B`. ⚠️ Đây là **lệch có chủ ý** so với đặc tả `{Dài}x{Rộng}x{Dày}`: với Savic, Dài là chiều dài cắt 750, đặt theo đặc tả sẽ ra "Thép hộp 750x80x3". Không có `Rộng B` thì dùng đúng đặc tả. **Chờ người dùng xác nhận.**
  2. **ống** → `Thép ống phi {Rộng}x{Dày}x{Dài}`
  3. **tròn đặc / láp / trục** → `Thép tròn đặc phi {Rộng}`
  4. **tấm** → `Thép tấm {Dày:.1f}` (ví dụ `8.0`, `1.5`)
  - Thiếu kích thước thì thử đọc từ mô tả (`phi 20`, `t=8`, `40x40x2`).
- Mã RM khi BOM không có Mã NVL: lấy mã PO nếu tên PO trùng tên chuẩn; ngược lại sinh `RM-` + slug (ví dụ `RM-THEP-TAM-8.0`, `RM-THEP-HOP-80X80X3`).
- ĐVT: KL nguyên liệu > 0 → `Kg`; ngược lại lấy ĐVT của dòng hoặc `Cái`.
- **Giá mặc định ở Sidebar: 22,000 VND/kg** (ĐVT Kg), **0 VND** cho ĐVT khác. Tỷ giá 25,400 (ô B6). SG&A 5%, thuế TNDN 20%, lead time 21 ngày, 50 SP/pallet, tiền tố NVL `R-`.
- Đơn giá công đoạn mặc định (`DEFAULT_STAGE_RATES`, VND/SP, chỉ là số gợi ý): cắt 12k/6k, chấn 8k/4k, khoan 6k/3k, hàn 25k/12k, mài 6k/3k, sơn 18k/15k, mạ 15k/12k, lắp 15k/6k, đóng gói 8k/3k; còn lại 10k/5k (nhân công / SXC).
- ⚠️ Savic hiện **chưa có giá thật** (BOM không có cột đơn giá, chưa có PO) → cả 6 RM đang dùng giá 22,000/kg.

## 6. Việc cần làm tiếp theo

1. **Chạy thử các bộ file thật còn lại** (bảng ở mục 1): Jack Pad, GP Orcana (+ `Mua hang PO 8201` → kịch bản 3 tệp), `BOM - Savic - ERP.xlsx`, bản CSV. Với mỗi file: chạy `tests/dump.py` trước, rồi `diag_savic.py` (sửa đường dẫn).
2. **Tìm file thật chứa `R-O-S01-00-032-001` / `R-P-S01-052-00-002`** và xác nhận 862.50 / 12,540.00 kg trên dữ liệu thật. Nếu lệch, kiểm tra lại cách hiểu "KL nguyên liệu / 1 chi tiết" và cách áp hao hụt.
3. **Hỏi người dùng** xác nhận quy tắc tên thép hộp (A×B×T hay đúng chữ đặc tả Dài×Rộng×Dày).
4. Nhập giá NVL thật cho Savic (file PO hoặc Live Editor) và đơn giá công đoạn thật của nhà máy.
5. Mở `master.xlsx` bằng **Excel thật** để kiểm tra hiển thị, `HYPERLINK`, định dạng ngày (mới kiểm bằng thư viện `formulas`).
6. Tùy chọn:
   - cho phép đổi tên / ghép mã RM tự sinh (bảng ánh xạ RM);
   - tận dụng cột "HOÀN THIỆN BỀ MẶT" (ví dụ "Mạ kẽm nhúng nóng");
   - khi PO dùng tên khác tên chuẩn ("Thép tấm 5ly"), khớp mờ theo độ dày;
   - đưa thư mục vào git.

## 7. Lưu ý khi sửa code

- Sau mọi thay đổi parser/engine, chạy lại tối thiểu: `tests/test_engine.py` (862.50 / 12,540), `tests/diag_savic.py` (tổng 30,896.59) và `tests/ui_test.py` (0 exception).
- Khi thêm regex vào `*_SPEC`: thứ tự field = thứ tự ưu tiên. Mỗi cột chỉ được gán cho 1 field. Field có hậu tố `#` (ví dụ `"Mã Sản Phẩm ERP#"`) là phương án dự phòng chạy sau.
- `norm("Mã")` và `norm("Mạ")` đều thành `"ma"`. Đừng nới lỏng `STAGE_EXCL`.
- Sửa nội dung PowerShell bằng `.Replace()` với chuỗi `@'...'@`: `` `n `` **không** được mở rộng trong chuỗi nháy đơn (đã từng làm mất một dòng regex). Nên dùng công cụ Edit.
- Giao diện viết bằng tiếng Việt; tên sheet Excel không dấu, tối đa 31 ký tự.
