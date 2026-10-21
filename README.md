# SohoaTHCS - CÔNG CỤ SỐ HÓA VÀ CHUẨN HÓA DỮ LIỆU VĂN BẰNG CHỨNG CHỈ THCS

Ứng dụng đồ họa chuyên dụng (**Desktop GUI Application**) phục vụ công tác số hóa, làm sạch, chuẩn hóa và tổng hợp dữ liệu cấp bằng tốt nghiệp Trung học cơ sở (THCS) của các Phòng/Sở Giáo dục và Đào tạo theo đúng chuẩn quy định của Bộ Giáo dục & Đào tạo.

![Python](https://img.shields.io/badge/Python-3.9%2B-blue?logo=python)
![Platform](https://img.shields.io/badge/Platform-macOS%20%7C%20Windows-green)
![Build](https://img.shields.io/badge/Build-GitHub%20Actions-orange?logo=github-actions)
![License](https://img.shields.io/badge/License-MIT-purple)

---

## 🚀 TẢI VỀ ỨNG DỤNG CHẠY TRỰC TIẾP (KHÔNG CẦN CÀI PYTHON)

Hệ thống được đóng gói tự động qua **GitHub Actions** thành các tệp chạy độc lập duy nhất:

| Hệ điều hành | Tệp tải về | Hướng dẫn sử dụng |
| :--- | :--- | :--- |
| 🪟 **Windows (x64)** | [**`SohoaTHCS.exe`**](https://github.com/hungphungvan/so_hoa_vbcc_thcs/releases/latest) | Tải về và click đúp để chạy trực tiếp trên Windows 10/11. Tự động hỗ trợ màn hình độ phân giải cao High-DPI. |
| 🍏 **macOS (Khuyên dùng)** | [**`SohoaTHCS_macOS.dmg`**](https://github.com/hungphungvan/so_hoa_vbcc_thcs/releases/latest) | **Tệp đĩa cài đặt chuẩn macOS:** Mở tệp `.dmg`, kéo biểu tượng ứng dụng vào thư mục `Applications` là xong! |
| 🍏 **macOS (Bản nén Zip)** | [**`SohoaTHCS_macOS.zip`**](https://github.com/hungphungvan/so_hoa_vbcc_thcs/releases/latest) | Giải nén, click đúp file `SohoaTHCS.app` để chạy. |

> [!NOTE]
> **Hướng dẫn khởi chạy trên macOS:**
> 
> **Cách 1: Chạy từ thư mục mã nguồn (Nhanh nhất & Tự động cập nhật)**
> 👉 Click đúp vào tệp **`Chay_Chuong_Trinh.command`** trong thư mục dự án để khởi động ngay giao diện đồ họa. Script sẽ tự động cập nhật bản mới nhất từ GitHub trong 1 giây mỗi khi chạy!
> 
> **Cách 2: Chạy từ tệp cài đặt `.dmg` tải từ GitHub Releases**
> 👉 Tải file `SohoaTHCS_macOS.dmg`, mở tệp và kéo ứng dụng vào `Applications`.
> Nếu macOS hiển thị cảnh báo Gatekeeper lần đầu mở: **Click chuột phải** (hoặc `Control` + Click) vào ứng dụng $\rightarrow$ chọn **Open (Mở)** $\rightarrow$ chọn tiếp **Open**.

---

## 📁 CẤU TRÚC THƯ MỤC DỰ ÁN

```text
so_hoa_vbcc_thcs/
├── data/                                 # Dữ liệu danh mục tra cứu & tệp mẫu chuẩn
│   ├── Danh_Muc_Tra_Cuu.xlsx             # 5 sheet: DS_Don_Vi (kèm Xã & Mã trường), Gioi_Tinh, Dan_Toc, Quoc_Tich, Noi_Sinh
│   ├── Mau_DS_Cap_Bang_THCS.xlsx         # Mẫu xuất chuẩn 28 cột Mô-đun 1 (Sheet 'Data')
│   └── Mau_DS_Tai_Lieu_So_Hoa.xlsx       # Mẫu xuất chuẩn 11 cột Mô-đun 2 (Sheet 'Data')
│
├── src/                                  # Mã nguồn xử lý nghiệp vụ
│   ├── version.py                        # Thông tin phiên bản (APP_VERSION) và repo GitHub
│   ├── updater.py                        # Tự động kiểm tra bản cập nhật từ GitHub Releases ngầm & SSL fallback
│   ├── lookup_manager.py                 # Nạp danh mục linh hoạt (ưu tiên file ngoài), tra cứu mã đơn vị/xã/trường
│   ├── normalizer.py                     # Động cơ chuyển đổi TCVN3/VNI, Title Case, ngày sinh, niên khóa, auto-fill ngày cấp
│   ├── excel_processor.py                # Pipeline 4 giai đoạn toàn cục, quét đệ quy rglob, cache an toàn 2 lớp
│   ├── exporter.py                       # Động cơ xuất file: Gộp, Tách, DS_Loi (4 cột truy vết), Nhat_Ky_Thay_Doi (10 cột)
│   └── gui/
│       ├── theme.py                      # Hệ thống chủ đề High-Contrast, chữ to rõ >= 10pt, hỗ trợ toàn diện Dark Mode
│       ├── app_window.py                 # Cửa sổ chính Tkinter, điều phối đa luồng worker, tiến trình, in-app updater
│       └── dialogs.py                    # Hộp thoại tương tác đối chiếu (Trường, Niên khóa, Ngày sinh, Ngày cấp) & Data Inspector
│
├── tests/                                # Bộ kiểm thử tự động toàn diện
│   ├── test_global_pipeline.py           # Kiểm thử pipeline 4 giai đoạn, hủy ngang, cache, xuất tệp và bảng mã TCVN3/VNI
│   ├── test_theme.py                     # Kiểm thử hệ thống phông chữ >= 10pt và bảng màu High-Contrast
│   └── test_updater.py                   # Kiểm thử hàm so sánh phiên bản ngữ nghĩa (semantic versioning)
│
├── main.py                               # Điểm khởi chạy ứng dụng (High-DPI Awareness, UTF-8 stdout)
├── Chay_Chuong_Trinh.bat                 # Script click đúp khởi chạy & đồng bộ tự động trên Windows
├── Chay_Chuong_Trinh.command             # Script click đúp khởi chạy & đồng bộ tự động trên macOS
└── .github/workflows/release.yml         # CI/CD tự động đóng gói release cho Windows (.exe) và macOS (.dmg/.zip)
```

---

## 📖 CẤU TRÚC 3 MÔ-ĐUN NGHIỆP VỤ

Hệ thống cung cấp 3 mô-đun nghiệp vụ hoàn chỉnh:
1. **MÔ-ĐUN 1: CHUẨN HOÁ DỮ LIỆU CẤP BẰNG THCS** (Biểu mẫu 28 cột chuẩn theo `data/Mau_DS_Cap_Bang_THCS.xlsx`).
2. **MÔ-ĐUN 2: XUẤT DANH SÁCH TÀI LIỆU SỐ HÓA** (Biểu mẫu 11 cột chuẩn theo `data/Mau_DS_Tai_Lieu_So_Hoa.xlsx`).
3. **MÔ-ĐUN 3: TIỆN ÍCH NHẬP & TÁCH EXCEL THEO CỘT TÙY BIẾN** (Xử lý tệp hỗn hợp nhiều đơn vị, gộp/tách linh hoạt theo bất kỳ cột nào có cùng cấu trúc tiêu đề).

---

## 🛠️ QUY TRÌNH QUÉT & TRÍCH XUẤT DỮ LIỆU TỪ TỆP EXCEL NGUỒN

### 1. Khả năng tương thích 3 cấu trúc thư mục nguồn (Quét đệ quy `rglob`)
Phần mềm tự động quét đệ quy toàn bộ thư mục và nhận diện dữ liệu dù đơn vị tổ chức lưu trữ theo bất kỳ mô hình nào:
- **Cấu trúc 1:** Thư mục Đơn vị $\rightarrow$ Các tệp Excel là Trường $\rightarrow$ Các Sheet là Niên khóa.
- **Cấu trúc 2:** Thư mục Đơn vị $\rightarrow$ Các thư mục con là Trường $\rightarrow$ Các tệp Excel là Niên khóa.
- **Cấu trúc 3:** Thư mục Đơn vị $\rightarrow$ Các tệp Excel là Niên khóa $\rightarrow$ Các Sheet là Trường (hoặc chung 1 sheet).

### 2. Tự động nhận diện định dạng Sheet (Dạng AB 28 cột & Dạng AE 31 cột)
- Quét tìm dòng tiêu đề học sinh (chứa STT hoặc Họ và tên).
- Nếu phát hiện **Dạng AE (31 cột)**: Hệ thống tự động cắt bỏ 3 cột thừa Q, R, S (Điểm các môn thi, Hội đồng thi, Địa danh) để chuyển về đúng 28 cột chuẩn.
- Nếu phát hiện **Dạng AB (28 cột chuẩn)**: Tự động đọc dữ liệu và bù đủ 28 cột bằng `None`.

### 3. Tự động nhận diện và chuyển đổi bảng mã cũ (TCVN3 & VNI-Windows $\rightarrow$ Unicode)
- Rất nhiều hồ sơ THCS cũ từ các năm 199x – 201x được soạn thảo bằng các font chữ cũ như `.VnTime` (bảng mã **TCVN3/ABC**) hoặc `VNI-Times` (bảng mã **VNI-Windows**).
- Hệ thống tích hợp **Động cơ phân giải bảng mã thông minh 4 lớp (Precedence-based Legacy Encoding Engine)**:
  - **Lớp 1**: Bỏ qua chuỗi thuần ASCII ($O(1)$) để đạt tốc độ tối đa.
  - **Lớp 2**: Kiểm tra ký tự Unicode cấp cao (`ord >= 256` hoặc ký tự độc quyền Unicode tiếng Việt) để tuyệt đối không chuyển đổi nhầm chuỗi đã là Unicode chuẩn (như *"Hoàng Văn Thái"*, *"Toàn diện"*).
  - **Lớp 3**: Nhận diện và chuyển đổi bảng mã VNI-Windows (chuỗi ghép dấu đa ký tự như `öôù` $\rightarrow$ `ướ`, `aâù` $\rightarrow$ `ấ`, `Ñ` $\rightarrow$ `Đ`).
  - **Lớp 4**: Nhận diện và chuyển đổi bảng mã TCVN3 (ký tự đơn byte đặc thù như `µ` $\rightarrow$ `à`, `¸` $\rightarrow$ `á`, `©` $\rightarrow$ `đ`).
- **Phạm vi tự động 100%**: Áp dụng xuyên suốt trên Tên Sheet, Dòng tiêu đề cột, Toàn bộ ô dữ liệu học sinh, và Dòng chân trang/tổng kết mà không cần người dùng thao tác thủ công và không làm chậm tốc độ xử lý.

### 4. Cơ chế loại bỏ dòng rác và chân trang
- Tự động bỏ qua các dòng tổng kết, người lập bảng, người ký, thủ trưởng đơn vị dựa trên bộ từ khóa kiểm soát: *"danh sách này có", "tổng số", "tổng cộng", "số học sinh", "công nhận tốt nghiệp", "hiệu trưởng", "người ký", "thủ trưởng"...*
- Tự động dừng quét khi gặp liên tiếp 20 dòng trống.

### 5. Quy trình xử lý toàn cục 4 giai đoạn xuyên suốt tất cả File & Sheet (Global Multi-Phase Pipeline)
Khác với mô hình cũ xử lý lắt nhắt theo từng Sheet đơn lẻ, phần mềm áp dụng kiến trúc ETL toàn cục 4 giai đoạn kết hợp tối ưu hóa dòng dữ liệu (Value-only Stream):
1. **Giai đoạn 0 (Trích xuất thô & Tối ưu RAM):** Đọc dữ liệu dạng `iter_rows(values_only=True)`, không tạo Cell DOM trong bộ nhớ, tiết kiệm tới ~60% dung lượng RAM và tăng tốc độ đọc gấp 3-5 lần; tự động chuyển đổi bảng mã cũ sang Unicode và đóng file ngay.
2. **Giai đoạn 1/4 (Chuẩn hóa Ngày sinh Cột I):** Tự động chuẩn hóa format `dd/MM/yyyy` cho 100% học sinh; mở hộp thoại `DateFixDialog` (kèm ngữ cảnh 10 học sinh cùng lớp qua cửa sổ trượt đối chiếu) để xử lý dứt điểm toàn bộ ngày sinh lỗi.
3. **Giai đoạn 2/4 (Chuẩn hóa Tên trường Cột N):** Tự động so khớp danh mục trường trực thuộc của Đơn vị; mở hộp thoại `SchoolBlockDialog` cho các nhóm trường bị trống hoặc có tên lạ ngoài danh mục (kèm nút gợi ý theo thứ tự Thư mục $\rightarrow$ File $\rightarrow$ Sheet).
4. **Giai đoạn 3/4 (Chuẩn hóa Niên khóa Cột O):** Gom nhóm và tính toán **Năm sinh đa số** chính xác tuyệt đối (do Ngày sinh ở Giai đoạn 1 đã được làm sạch 100%). Tự động chấp nhận nếu $\text{Năm tốt nghiệp} == \text{Năm sinh đa số} + 15$; áp dụng cơ chế bảo vệ an toàn 2 lớp khi dùng cache.
5. **Giai đoạn 4/4 (Chuẩn hóa Ngày cấp bằng Cột S):** Khi Trường và Niên khóa đã sạch 100% trên toàn đơn vị, hệ thống tự động gom nhóm `(Trường, Niên khóa)` để điền ngày cấp bằng chuẩn có nhiều học sinh nhất cho các trường hợp chỉ có tháng/năm hoặc năm.
6. **Giai đoạn 5 (Xuất tệp kết quả & Báo cáo kiểm soát):** Tự động tính toán Cột Y (Trạng thái số hóa), Cột Z (Nguyên nhân lỗi), xuất các tệp Gộp, Tách, Danh sách lỗi tổng hợp, Danh sách tài liệu số hóa và tệp Nhật ký thay đổi chi tiết theo đúng tùy chọn được tích chọn.

---

## 📋 CHI TIẾT CÁCH LẤY DỮ LIỆU, CHUẨN HÓA, SỬA & BỔ SUNG TỪNG CỘT

### 🔷 MÔ-ĐUN 1: CHUẨN HOÁ DỮ LIỆU CẤP BẰNG (BIỂU MẪU 28 CỘT A $\rightarrow$ AB)

| Cột | Tên trường thông tin | Cách lấy dữ liệu gốc | Quy tắc chuẩn hóa (Format, Data Type) | Quy tắc sửa & Xử lý khi sai | Quy tắc bổ sung khi bị thiếu |
| :---: | :--- | :--- | :--- | :--- | :--- |
| **A** | **STT** | Cột 1 (index 0) từ Excel gốc | Chuyển thành số nguyên `int` liên tục ($1, 2, 3...$). | Bỏ qua các dòng tiêu đề trùng lặp, dòng tổng kết hoặc STT không phải là số hợp lệ. | Đánh lại số thứ tự liên tục $1..N$ tự động khi xuất file gộp, file tách và file lỗi. Không tô vàng. |
| **B** | **Số hiệu bằng** | Cột 2 (index 1) từ Excel gốc | Ép định dạng Text `@` (`cell.number_format = '@'`, `cell.data_type = 's'`), xóa khoảng trắng thừa. | Giữ nguyên chuỗi ký tự gốc (kể cả số 0 ở đầu). | Nếu thiếu: Giữ ô trống (**không tô vàng** cảnh báo). |
| **C** | **Số vào sổ gốc cấp văn bằng** | Cột 3 (index 2) từ Excel gốc | Ép định dạng Text `@`, xóa khoảng trắng thừa. | Giữ nguyên chuỗi ký tự gốc (kể cả dạng `123/QĐ`, `01/THCS`). | Nếu thiếu: Giữ ô trống (**không tô vàng** cảnh báo). |
| **D** | **Tên văn bằng** | Cố định theo quy chuẩn Bộ GD&ĐT | Cố định chuỗi: `"Bằng tốt nghiệp trung học cơ sở"`. | Ghi đè tự động, không phụ thuộc vào dữ liệu gõ sai trong file nguồn. | Tự động điền đầy đủ 100% cho mọi học sinh. Không tô vàng. |
| **E** | **Họ, chữ đệm và tên** | Cột 5 (index 4) từ Excel gốc | Chuyển sang **Title Case** chuẩn tiếng Việt (`vietnamese_title_case`). Xóa khoảng trắng thừa, chuẩn Unicode NFC. | Hỗ trợ chuẩn xác các tên có dấu gạch nối (dân tộc thiểu số như *Ê-đê, Ba-na, K'Ho, Bru-Vân Kiều*). | Bắt buộc phải có chữ cái tiếng Việt/Latinh (nếu trống sẽ không trích xuất dòng này). |
| **F** | **Mã người học** | Cột 6 (index 5) từ Excel gốc | Ép định dạng Text `@`, xóa khoảng trắng thừa. | Giữ nguyên mã số định danh học sinh từ phần mềm quản lý trường học. | Nếu thiếu: Giữ ô trống (**không tô vàng**). |
| **G** | **Số định danh cá nhân (CCCD)** | Cột 7 (index 6) từ Excel gốc | Ép định dạng Text `@`, xóa khoảng trắng thừa. | Giữ nguyên 12 chữ số CCCD (không bị Excel làm mất số 0 ở đầu hoặc chuyển thành số thực dạng `2.02E+11`). | Nếu thiếu: Giữ ô trống (**không tô vàng**). |
| **H** | **Hộ chiếu** | Cột 8 (index 7) từ Excel gốc | Ép định dạng Text `@`, viết hoa toàn bộ mã hộ chiếu. | Giữ nguyên mã số hộ chiếu của học sinh người nước ngoài. | Nếu thiếu: Giữ ô trống (**không tô vàng**). |
| **I** | **Ngày, tháng, năm sinh** | Cột 9 (index 8) từ Excel gốc | **Định dạng Text `@` tuyệt đối dạng `dd/MM/yyyy`**. Hỗ trợ cả số serial Excel, số liền 7-8 chữ số (`12241993`), chuỗi ISO (`2007-05-15 00:00:00`). | 1. **Tự động đảo MM/DD/YYYY $\rightarrow$ DD/MM/YYYY** khi Tháng $\le 12$ và Ngày $> 12$ (kể cả số liền `12241993` $\rightarrow$ `24/12/1993`).<br>2. **Mở hộp thoại `DateFixDialog`** khi sai logic (tháng > 12, ngày 00, năm nhuận...) kèm **cửa sổ trượt đối chiếu**: đầu DS hiển thị tối đa 10 HS sau; cuối hiển thị 10 HS trước; giữa hiển thị 5 trước - 5 sau.<br>3. Có nút bấm *"🔄 Đảo Ngày ↔ Tháng"* và tự tính gợi ý từ niên khóa. | Nếu thiếu hẳn ngày sinh: Ghi nhận trạng thái Cột Y là **"Chưa số hoá"**, Cột Z là **"Thiếu ngày sinh"** và **tô màu vàng**. |
| **J** | **Giới tính** | Cột 10 (index 9) từ Excel gốc | Chuẩn hóa Title Case: `"Nam"` hoặc `"Nữ"` hoặc `"Khác"`. | Tự động ánh xạ từ sheet `Gioi_Tinh` (`1` $\rightarrow$ Nam, `2` $\rightarrow$ Nữ, `nam/nữ/nu/khac` $\rightarrow$ Nam/Nữ/Khác). | Nếu thiếu: Giữ ô trống và **tô màu vàng cảnh báo**. |
| **K** | **Dân tộc** | Cột 11 (index 10) từ Excel gốc | Chuẩn hóa Title Case tiếng Việt. | Tự động ánh xạ từ mã số sang tên chữ dựa trên sheet `Dan_Toc` trong `Danh_Muc_Tra_Cuu.xlsx` (ví dụ: `01` $\rightarrow$ Kinh, `02` $\rightarrow$ Tày, `03` $\rightarrow$ Thái...). | Nếu thiếu: Giữ ô trống và **tô màu vàng cảnh báo**. |
| **L** | **Quốc tịch** | Cột 12 (index 11) từ Excel gốc | Chuẩn hóa Title Case tiếng Việt. | Tự động ánh xạ từ mã số dựa trên sheet `Quoc_Tich` (mặc định phổ biến: `Việt Nam`). | Nếu thiếu: Giữ ô trống và **tô màu vàng cảnh báo**. |
| **M** | **Nơi sinh** | Cột 13 (index 12) từ Excel gốc | Chuẩn hóa Title Case tiếng Việt. | Ánh xạ từ mã tỉnh/thành sheet `Noi_Sinh`. **Quy tắc đặc biệt**: Cụm từ phân loại hành chính *"Thành phố"*, *"Thị xã"*, *"Thị trấn"* chỉ viết hoa chữ cái đầu tiên (ví dụ: *Thành phố Hà Nội, Thị xã Sơn Tây, Tỉnh Nam Định*). | Nếu thiếu: Giữ ô trống và **tô màu vàng cảnh báo**. |
| **N** | **Tên trường đã học** | Cột 14 (index 13) từ Excel gốc | Chuyển về chữ thường, bỏ khoảng trắng thừa để đối chiếu với danh mục trường trực thuộc của Đơn vị (sheet `DS_Don_Vi`). | 1. **Khớp danh mục**: Tự động gán tên chuẩn gốc (Title Case) từ `DS_Don_Vi`, không mở dialog.<br>2. **Không khớp**: Mở `SchoolBlockDialog` kèm các gợi ý gần đúng nhất (Fuzzy match) để người dùng chọn trường chuẩn trực thuộc.<br>3. Kết quả ghi vào Excel bắt buộc là tên trường chuẩn từ danh mục `DS_Don_Vi`. | **Nếu bị trống**: Hiển thị dialog kèm thông tin `📁 Thư mục`, `📄 Tên File`, `📑 Tên Sheet` và danh sách nút gợi ý theo thứ tự **Thư mục $\rightarrow$ File $\rightarrow$ Sheet** để người dùng chọn. |
| **O** | **Niên khóa** | Cột 15 (index 14) từ Excel gốc | Chuẩn hóa về định dạng **`YYYY-YYYY`** (ví dụ: `2018-2022`). Hỗ trợ cả dạng viết tắt `18-22`, `18_22`. | 1. **Nếu hiệu số năm bằng 4**: Kiểm tra với năm sinh đa số trong nhóm: nếu $\text{Năm tốt nghiệp} == \text{Năm sinh} + 15$ $\rightarrow$ **Tự động chấp nhận**, không hỏi người dùng.<br>2. Nếu khác năm sinh hoặc hiệu khác 4 hoặc chỉ có 1 năm: Mở hộp thoại hiển thị tối đa 15 HS trong nhóm để đối chiếu. Người dùng chỉ cần nhập Năm tốt nghiệp, hệ thống tự động tính niên khóa 4 năm chuẩn.<br>3. **Cơ chế ghi nhớ an toàn 2 lớp**: Khi chọn *"Tự động ghi nhớ"*, ở các nhóm tiếp theo phần mềm chỉ tự động áp dụng nếu $\text{Năm tốt nghiệp đã nhớ} == \text{Năm sinh đa số} + 15$. Nếu phát hiện lệch tuổi, hệ thống cảnh báo và mở lại hộp thoại. | **Nếu bị trống**: Hiển thị tối đa 15 HS kèm tên File, Sheet; gợi ý mặc định $\text{Năm tốt nghiệp} = \text{Năm sinh đa số} + 15$ để người dùng xác nhận hoặc nhập lại. |
| **P** | **Điểm các môn thi** | Cột 16 (index 15) từ Excel gốc | Ép định dạng Text `@`. | Giữ nguyên điểm số nếu có trong hồ sơ cũ. | Nếu thiếu: Để trống (**không tô vàng**). |
| **Q** | **Hội đồng thi** | Cột 17 (index 16) từ Excel gốc | Chuẩn hóa Title Case, xóa khoảng trắng thừa. | Giữ nguyên tên Hội đồng thi (nếu có). | Nếu thiếu: Để trống và **tô màu vàng**. |
| **R** | **Địa danh nơi cơ quan cấp bằng đặt trụ sở** | Cột 18 (index 17) từ Excel gốc | Chuẩn hóa Title Case tiếng Việt. | Giữ nguyên địa danh hành chính (nếu có). | Nếu thiếu: Để trống và **tô màu vàng**. |
| **S** | **Ngày tháng năm cấp bằng** | Cột 19 (index 18) từ Excel gốc | **Định dạng Text `@` tuyệt đối dạng `dd/MM/yyyy`**. Hỗ trợ cả số serial Excel và cắt bỏ giờ ISO. | **Chạy tự động 2 bước theo nhóm (Trường, Niên khóa)**:<br>• *Bước 1*: Quét tìm tất cả các ngày cấp chuẩn `dd/MM/yyyy` theo từng cặp `(Trường, Niên khóa)`, xếp theo số lượng HS giảm dần. Ngày chuẩn $\rightarrow$ Giữ nguyên.<br>• *Bước 2*: Với ngày chưa chuẩn của đúng (Trường, Niên khóa) đó:<br>  - Dạng `MM/yyyy`: tìm ngày chuẩn cùng tháng & năm điền vào; không thấy $\rightarrow$ bỏ trống.<br>  - Dạng `yyyy`: tìm ngày chuẩn cùng năm có nhiều HS nhất điền vào; không thấy $\rightarrow$ bỏ trống.<br>  - Trống hoặc dạng khác $\rightarrow$ bỏ trống. | Nếu không tìm thấy ngày chuẩn trong nhóm: Để trống và **tô màu vàng**. |
| **T** | **Tên cơ quan cấp bằng** | Cột 20 (index 19) từ Excel gốc | Chuẩn hóa Title Case tiếng Việt. | Giữ nguyên tên cơ quan cấp bằng (nếu có). | Nếu thiếu: Để trống và **tô màu vàng**. |
| **U** | **Chức danh người ký bằng** | Cột 21 (index 20) từ Excel gốc | Chuẩn hóa Title Case (ví dụ: *Trưởng phòng, Phó Trưởng phòng, Hiệu trưởng*). | Giữ nguyên chức danh người ký (nếu có). | Nếu thiếu: Để trống và **tô màu vàng**. |
| **V** | **Họ, chữ đệm, tên người ký bằng** | Cột 22 (index 21) từ Excel gốc | Chuẩn hóa Title Case (`vietnamese_title_case`). | Giữ nguyên họ tên người ký bằng (nếu có). | Nếu thiếu: Để trống và **tô màu vàng**. |
| **W** | **Tình trạng văn bằng** | Cột 23 (index 22) từ Excel gốc | Chuẩn hóa Title Case (ví dụ: *Cấp lần đầu, Cấp lại, Đã thu hồi*). | Giữ nguyên tình trạng văn bằng (nếu có). | Nếu thiếu: Để trống (**không tô vàng**). |
| **X** | **Ghi chú** | Cột 24 (index 23) từ Excel gốc | Ép định dạng Text `@`. | Giữ nguyên vẹn ghi chú ban đầu từ file gốc (nếu có) và không bị chèn chuỗi `"None"`. | Nếu file gốc để trống thì để trống hoàn toàn (**không tô vàng**). |
| **Y** | **Trạng thái số hóa** | Tự động tính toán logic theo quy định | Ép kiểu Text `@`. Chỉ nhận 1 trong 3 giá trị chuẩn. | **Quy tắc ấn định tự động 100%**:<br>1. `"Chưa số hoá"`: Khi thiếu hẳn ngày sinh (Cột I rỗng cả ở file gốc và sau xử lý).<br>2. `"Số hoá chưa đầy đủ, chính xác"`: Khi có ngày sinh nhưng sai định dạng hoặc sai logic thực tế.<br>3. `"Đã số hóa đầy đủ, chính xác"`: Khi ngày sinh đầy đủ và hoàn toàn chính xác. | Luôn được điền tự động 100% cho mọi học sinh. Không tô vàng. |
| **Z** | **Thông tin chưa bảo đảm đầy đủ, chính xác và nguyên nhân** | Tự động tính toán logic theo quy định | Ép kiểu Text `@`. | **Quy tắc giải trình lỗi tự động**:<br>1. Nếu Cột Y là `"Chưa số hoá"` $\rightarrow$ Điền `"Thiếu ngày sinh"`.<br>2. Nếu Cột Y là `"Số hoá chưa đầy đủ, chính xác"` $\rightarrow$ Ghi rõ nguyên nhân cụ thể (ví dụ: *Ngày sinh không hợp lệ: Tháng 02 năm 2007 không thể có 30 ngày*).<br>3. Nếu Cột Y là `"Đã số hóa đầy đủ, chính xác"` $\rightarrow$ Bỏ trống (`None`). | Luôn được điền tự động 100% cho mọi học sinh. Không tô vàng. |
| **AA** | **Mã đơn vị** | Nạp từ danh mục tra cứu `DS_Don_Vi` | Ép định dạng Text `@` (ví dụ: `722`, `723`, `724`...). | Lấy chính xác mã số đơn vị của Phòng GD&ĐT được người dùng chọn trên combobox giao diện. | Tự động điền đầy đủ 100% cho mọi học sinh. Không tô vàng. |
| **AB** | **Tên đơn vị** | Nạp từ danh mục tra cứu `DS_Don_Vi` | Chuẩn hóa Title Case tiếng Việt. | Lấy chính xác tên đơn vị quản lý cấp bằng được chọn (ví dụ: *Phòng Giáo dục và Đào tạo Huyện Đông Anh*). | Tự động điền đầy đủ 100% cho mọi học sinh. Không tô vàng. |

---

### 🔷 MÔ-ĐUN 2: DANH SÁCH TÀI LIỆU SỐ HÓA (BIỂU MẪU 11 CỘT A $\rightarrow$ K)

Mô-đun 2 tự động tổng hợp hồ sơ số hóa từ kết quả của Mô-đun 1 theo biểu mẫu chuẩn `data/Mau_DS_Tai_Lieu_So_Hoa.xlsx` gồm 11 cột:

| Cột | Tên trường thông tin | Cách lấy & Tổng hợp dữ liệu từ Mô-đun 1 | Quy tắc tính toán, chuẩn hóa & Sinh mã tự động |
| :---: | :--- | :--- | :--- |
| **A** | **STT** | Đánh số thứ tự tăng dần liên tục ($1, 2, 3...$) | Tự động sinh cho từng dòng hồ sơ/sổ gốc tài liệu số hóa. |
| **B** | **Mã đơn vị** | Lấy từ Cột AA (Mã đơn vị) của Mô-đun 1 | Ép định dạng Text `@` tuyệt đối (ví dụ: `722`). |
| **C** | **Tên đơn vị** | Lấy từ Cột AB (Tên đơn vị) của Mô-đun 1 | Giữ nguyên tên đơn vị chuẩn (ví dụ: *Phòng Giáo dục và Đào tạo Quận Ba Đình*). |
| **D** | **Mã sổ gốc** | Sinh tự động theo quy tắc định danh mã hóa | Ép định dạng Text `@` theo 3 tùy chọn xuất:<br>• *Theo trường*: `{Mã_Đơn_Vị}_{Mã_Trường}` (ví dụ: `722_THCSND`).<br>• *Theo năm*: `{Mã_Đơn_Vị}_{Năm}` (ví dụ: `722_2022`).<br>• *Theo trường & năm*: `{Mã_Đơn_Vị}_{Mã_Trường}_{Năm}` (ví dụ: `722_THCSND_2022`).<br>*(Nếu trường chưa có mã trong danh mục, dùng tên trường không dấu viết liền).* |
| **E** | **Tên tài liệu** | Sinh tự động từ Tên trường, Tên đơn vị và Khung năm tốt nghiệp | • *Theo trường*: `Danh sách cấp bằng THCS trường {Tên_Trường} từ {Từ_Năm} đến {Đến_Năm}` (hoặc `năm {Năm}` nếu chỉ có 1 năm).<br>• *Theo năm*: `Danh sách cấp bằng THCS đơn vị {Tên_Đơn_Vị} năm {Năm}`.<br>• *Theo trường & năm*: `Danh sách cấp bằng THCS trường {Tên_Trường} năm {Năm}`. |
| **F** | **Từ năm** | Quét toàn bộ học sinh trong nhóm để tìm năm ra trường ($Y_{ra}$) nhỏ nhất | Trích xuất từ Niên khóa (Cột O) hoặc Ngày cấp bằng (Cột S) hoặc Năm sinh + 15. |
| **G** | **Đến năm** | Quét toàn bộ học sinh trong nhóm để tìm năm ra trường ($Y_{ra}$) lớn nhất | Nếu nhóm chỉ có 1 niên khóa duy nhất thì Đến năm = Từ năm. |
| **H** | **Ngày lập sổ** | Lấy ngày xuất tệp hiện tại trên hệ thống | Ép định dạng Text `@` dạng `dd/MM/yyyy` (ví dụ: `21/10/2026`). |
| **I** | **Số trang** | Mặc định để trống (`None`) | Dành cho đơn vị tự điền số trang thực tế sau khi in/đóng tập sổ gốc. |
| **J** | **Tình trạng tài liệu** | Mặc định cố định | Cố định giá trị: `"Nguyên vẹn"`. |
| **K** | **Thông tư** | Mặc định để trống (`None`) | Dành cho đơn vị ghi chú số hiệu thông tư/quy chế văn bằng tương ứng. |

---

### 🔷 MÔ-ĐUN 3: BỘ TIỆN ÍCH NHẬP & TÁCH EXCEL THEO CỘT TÙY BIẾN

Mô-đun 3 là công cụ tiền xử lý và hậu xử lý dữ liệu mạnh mẽ, giúp giải quyết triệt để các tình huống thực tế:
- **Tình huống 1:** Nhận được **file hỗn hợp của nhiều đơn vị/phòng GD&ĐT** $\rightarrow$ Tách nhanh thành các file độc lập theo Cột *Đơn vị* để nạp vào Mô-đun 1 & 2.
- **Tình huống 2:** Có nhiều file Excel lẻ từ các trường/năm cần gộp chung lại thành 1 file tổng $\rightarrow$ Tự động gộp dữ liệu và đánh lại STT.

#### Hai tính năng cốt lõi:
1. **📤 Tách tệp theo cột tùy chọn (Custom Splitter):**
   - Cho phép người dùng tự do chọn **bất kỳ cột nào** làm căn cứ tách (ví dụ: *Tên đơn vị, Tên trường, Năm tốt nghiệp, Xếp loại...*).
   - Có bảng **Xem trước (Preview)**: Thống kê số lượng nhóm và số dòng dữ liệu của từng file con dự kiến trước khi bấm tách.
   - **Giữ nguyên khối Header & Định dạng:** Toàn bộ banner, tiêu đề, font chữ và độ rộng cột được sao chép nguyên vẹn sang mọi file con.
   - Tự động đánh lại STT liên tục ($1, 2, 3...$) cho từng file con.

2. **📥 Nhập / Gộp nhiều tệp (Header-matching Merger):**
   - Hỗ trợ chọn cùng lúc nhiều file hoặc chọn cả thư mục.
   - **Cơ chế kiểm tra tính tương thích tiêu đề (Header Validation):** So sánh số cột và tên cột của các file với file mẫu để phát hiện kịp thời các file lệch cấu trúc.
   - Cắt bỏ banner và tiêu đề lặp lại của các file sau, nối tiếp phần thân dữ liệu mượt mà.
   - Tự động đánh lại STT toàn bảng từ $1 \rightarrow \text{Tổng số dòng}$.
   - Tùy chọn thêm cột ghi rõ *Tệp nguồn gốc* để dễ dàng truy vết sau này.

#### Thuật toán nhận diện dòng dữ liệu chuẩn xác:
- **Dòng dữ liệu hợp lệ:** Bắt buộc thỏa mãn đồng thời: **Có STT hợp lệ VÀ Có Họ và tên hợp lệ** (loại bỏ 100% dòng banner, dòng tiêu đề ngắt trang, dòng trắng và dòng chữ ký chân trang).
- **Cơ chế dừng sớm (Early Exit):** Tự động kết thúc quét ngay khi phát hiện **15 dòng liên tiếp có STT trống**, triệt tiêu tình trạng đơ/chậm do dòng trống ảo cuối file Excel.

---

## 📦 HỆ THỐNG XUẤT TỆP ĐA DẠNG & BÁO CÁO KIỂM SOÁT

Sau khi chuẩn hóa, toàn bộ tệp kết quả được lưu tại thư mục `Ket_Qua_So_Hoa/`. Giao diện phần mềm cung cấp **4 cột tùy chọn xuất khoa học** cùng **ràng buộc an toàn (Validation Guard)**: người dùng bắt buộc chọn ít nhất 01 tùy chọn trước khi bấm *Bắt đầu xử lý*.

### 1. Cột 1: Gộp dữ liệu
- **Gộp theo Trường**: Xuất tệp `{Ten_Don_Vi}_Đa_Gop_Theo_Truong.xlsx` (sắp xếp: Tên trường $\rightarrow$ Năm tốt nghiệp $\rightarrow$ STT gốc).
- **Gộp theo Năm**: Xuất tệp `{Ten_Don_Vi}_Đa_Gop_Theo_Nam.xlsx` (sắp xếp: Năm tốt nghiệp $\rightarrow$ Tên trường $\rightarrow$ STT gốc).

### 2. Cột 2: Tách dữ liệu
- **Tách theo Trường**: Thư mục `Da_Tach_Theo_Truong/{Ten_Truong}_Tu_{Min}_Den_{Max}.xlsx`.
- **Tách theo Năm**: Thư mục `Da_Tach_Theo_Nam/{Ten_Don_Vi}_{Nam}.xlsx`.
- **Tách theo Trường & Năm**: Thư mục `Tach_Theo_Truong_Va_Nam/{Ten_Truong}/{Ten_Truong}_{Nam}.xlsx`.

### 3. Cột 3: Danh mục hồ sơ số hóa (Mô-đun 2)
- **Tài liệu số hóa theo Trường**: `{Ten_Don_Vi}_DS_Tai_Lieu_So_Hoa_Theo_Truong.xlsx`.
- **Tài liệu số hóa theo Năm**: `{Ten_Don_Vi}_DS_Tai_Lieu_So_Hoa_Theo_Nam.xlsx`.
- **Tài liệu số hóa theo Trường & Năm**: `{Ten_Don_Vi}_DS_Tai_Lieu_So_Hoa_Theo_Truong_Va_Nam.xlsx`.

### 4. Cột 4: Báo cáo & Kiểm soát (Mặc định chọn sẵn cả 2)
- **Danh sách lỗi tổng hợp (`{Ten_Don_Vi}_DS_Loi.xlsx`)**:
  - Gom toàn bộ học sinh có trạng thái *"Chưa số hoá"* hoặc *"Số hoá chưa đầy đủ, chính xác"*.
  - Mở rộng thêm **4 cột chỉ đường truy vết nguồn gốc (Cột AC $\rightarrow$ AF)**:
    - **Cột AC (29)**: *Tệp nguồn gốc* (Tên file Excel ban đầu).
    - **Cột AD (30)**: *Sheet nguồn* (Tên Sheet chứa học sinh).
    - **Cột AE (31)**: *Dòng Excel gốc* (Tọa độ dòng thực tế trong file Excel gốc).
    - **Cột AF (32)**: *STT ban đầu* (Số thứ tự ban đầu của học sinh).
  - Giúp tra soát tận nơi dữ liệu gốc chỉ trong 1 giây mà không cần sinh các file lỗi vụn vặt theo từng trường hay từng năm (người dùng dễ dàng dùng tính năng Auto-Filter của Excel để lọc).
- **Nhật ký thay đổi (`{Ten_Don_Vi}_Nhat_Ky_Thay_Doi.xlsx`)**:
  - Ghi nhận chi tiết toàn bộ các thao tác đã can thiệp chuẩn hóa: Chuẩn hóa ngày sinh, chuẩn hóa tên trường, suy luận/sửa niên khóa, điền ngày cấp bằng theo nhóm...
  - Bảng nhật ký gồm **10 cột chuẩn trang trọng**:
    1. `STT`
    2. `Tệp nguồn gốc`
    3. `Sheet nguồn`
    4. `Dòng gốc`
    5. `STT học sinh`
    6. `Họ và tên`
    7. `Hạng mục chuẩn hóa`
    8. `Dữ liệu ban đầu (Trước)`
    9. `Dữ liệu chuẩn hóa (Sau)`
    10. `Căn cứ / Phương thức chuẩn hóa`
  - Có dòng Banner tiêu đề đơn vị trang trọng và tự động kích hoạt **Auto-Filter** cho bảng dữ liệu.
- **Cửa sổ Tra cứu & Đối soát song song (Data Inspector)**:
  - Tích hợp nút *"🔍 Tra cứu & Đối soát dữ liệu"* trực tiếp trên giao diện: tra cứu tức thì học sinh theo tên/STT và xem bảng so sánh trực quan đối ứng giữa **[Dữ liệu gốc]** và **[Dữ liệu chuẩn hóa]**. Người dùng không cần tạo thêm tệp dữ liệu gốc cồng kềnh mà vẫn kiểm soát 100% dữ liệu ban đầu.

---

## 🎨 GIAO DIỆN CÔNG THÁI HỌC, CHỮ TO RÕ & HỖ TRỢ TOÀN DIỆN DARK MODE

Giao diện ứng dụng được thiết kế tỉ mỉ, tối ưu hóa công thái học (**Ergonomic & High-Contrast Design**) cho người dùng làm việc liên tục nhiều giờ với bảng tính:

- **Hệ thống phông chữ to rõ**: Xóa bỏ hoàn toàn cỡ chữ 8pt, 9pt siêu nhỏ, nâng chuẩn toàn bộ nội dung lên $\ge 10\text{pt}$ (nhãn thông tin, danh sách trường, checkbox, console log), tiêu đề khối $11\text{pt}$ in đậm.
- **Bảng màu tương phản cao (High-Contrast Palette)**:
  - **☀️ Chế độ Sáng (Light Mode)**: Nền cửa sổ `#F0F2F5`, khối chức năng trắng tinh khôi `#FFFFFF` viền nổi rõ nét `#D1D5DB`, chữ đen than đậm `#111827`, chữ chú thích `#374151` (không còn xám mờ).
  - **🌙 Chế độ Tối (Dark Mode)**: Nền than chì `#18181B`, thẻ Card `#27272A` viền `#3F3F46`, chữ trắng sáng `#F9FAFB`, chữ phụ `#D1D5DB`, màu trạng thái dạ quang sắc nét và dịu mắt.
- **Nút chuyển đổi nhanh `☀️/🌙`**: Tích hợp nút chuyển đổi giao diện tức thì tại thanh công cụ dưới cùng, hoặc tự động kích hoạt theo cài đặt Dark Mode của hệ điều hành Windows / macOS.
- **Hỗ trợ Windows Per-Monitor High-DPI Awareness**: Tự động nhận diện và khử nhòe mờ do cơ chế bitmap scaling của Windows trên các màn hình Full HD, 2K, 4K (tỉ lệ phóng 125% - 150%).
- **Khắc phục lỗi hiển thị trên macOS**: Tự động kích hoạt cơ chế micro-resize bù trừ giúp Tkinter vẽ giao diện hoàn chỉnh ngay khi mở, khắc phục triệt để hiện tượng cửa sổ trắng/đen trên macOS.

---

## 🔄 TỰ ĐỘNG KIỂM TRA CẬP NHẬT TRONG ỨNG DỤNG (IN-APP AUTO UPDATER)

Phần mềm tích hợp mô-đun cập nhật thông minh ([`src/updater.py`](src/updater.py)):
- **Kiểm tra tự động ngầm**: Mỗi khi mở ứng dụng, một luồng ngầm (`daemon thread`) sẽ truy vấn GitHub Releases API để kiểm tra xem có bản phát hành mới hơn phiên bản hiện tại hay không mà hoàn toàn không làm đơ giật giao diện.
- **Nút trạng thái thông minh**:
  - Khi đang ở bản mới nhất: Hiển thị `🔄 Bản v2.0.0` (click vào để kiểm tra thủ công bất kỳ lúc nào).
  - Khi có bản phát hành mới: Tự động đổi thành `🔥 Có bản mới [vX.X]!` và thông báo nổi bật trên console nhật ký.
- **Cơ chế SSL Fallback thông minh**: Tự động xử lý chứng chỉ CA trên macOS (vượt qua hoàn toàn lỗi phổ biến `CERTIFICATE_VERIFY_FAILED` trên máy Mac chưa cài chứng chỉ Python hoặc môi trường PyInstaller).
- **1-Click Cập nhật**: Bấm vào nút cập nhật sẽ hiển thị Release Notes chi tiết và mở trực tiếp trang tải bản mới trên trình duyệt web mặc định.

---

## 📁 TÙY BIẾN DANH MỤC TRA CỨU & DANH SÁCH TRƯỜNG (`Danh_Muc_Tra_Cuu.xlsx`)

Phần mềm tích hợp cơ chế nhận diện động ưu tiên tải tài nguyên ([`src/lookup_manager.py`](src/lookup_manager.py)):
- **Mặc định**: Tệp danh mục chuẩn đã được đóng gói sẵn bên trong ứng dụng (`.exe` / `.app`).
- **Khi muốn thêm trường mới hoặc thay đổi danh sách trường/đơn vị**:
  - Người dùng chỉ cần đặt tệp `Danh_Muc_Tra_Cuu.xlsx` (hoặc thư mục `data/Danh_Muc_Tra_Cuu.xlsx`) **ngay bên cạnh tệp thực thi `SohoaTHCS.exe` hoặc ứng dụng `SohoaTHCS.app`**.
  - Phần mềm sẽ tự động phát hiện và ưu tiên nạp tệp bên ngoài này thay cho tệp mặc định nội bộ mà **không cần phải đóng gói hay build lại mã nguồn**.
  - Nguồn tệp danh mục đang nạp được hiển thị trực tiếp và minh bạch trên khung nhật ký log (`📖 Nguồn danh mục: ...`) ngay khi mở phần mềm.
- **Cấu trúc 5 Sheet trong tệp danh mục**:
  1. `DS_Don_Vi`: Gồm các cột *Mã trường, Tên trường đã học, Mã đơn vị, Tên đơn vị, Xã/phường*. Combobox trên giao diện hiển thị trực quan dạng `Tên Đơn vị - Xã/phường [Mã: XXX]` giúp chống nhầm lẫn tuyệt đối.
  2. `Gioi_Tinh`: Bảng ánh xạ mã giới tính (1 $\rightarrow$ Nam, 2 $\rightarrow$ Nữ...).
  3. `Dan_Toc`: Bảng ánh xạ mã dân tộc (01 $\rightarrow$ Kinh, 02 $\rightarrow$ Tày...).
  4. `Quoc_Tich`: Bảng ánh xạ quốc tịch (VN $\rightarrow$ Việt Nam...).
  5. `Noi_Sinh`: Bảng ánh xạ mã địa danh tỉnh/thành phố.

---

## 💻 HƯỚNG DẪN KHỞI CHẠY TỪ MÃ NGUỒN

Nếu bạn muốn chạy trực tiếp từ mã nguồn hoặc tham gia phát triển:

### Cách 1: Chạy nhanh bằng script click đúp (Khuyên dùng)
- **Trên Windows**: Click đúp tệp **`Chay_Chuong_Trinh.bat`**.
- **Trên macOS**: Click đúp tệp **`Chay_Chuong_Trinh.command`**.
*(Script sẽ tự động kiểm tra và đồng bộ mã nguồn mới nhất từ GitHub qua `git pull`, kiểm tra môi trường Python/UV và khởi chạy ngay).*

### Cách 2: Khởi chạy bằng công cụ siêu tốc `uv`
```bash
# Cài đặt uv (nếu chưa có):
# Windows: powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
# macOS/Linux: curl -LsSf https://astral.sh/uv/install.sh | sh

# Khởi chạy ứng dụng:
uv run --with openpyxl main.py
```

### Cách 3: Khởi chạy bằng Python truyền thống
```bash
# Cài đặt thư viện phụ thuộc:
pip install openpyxl certifi

# Khởi chạy:
python main.py
```

---

## 🧪 BỘ KIỂM THỬ TỰ ĐỘNG (UNIT TESTS)

Dự án được bảo vệ bởi bộ kiểm thử tự động toàn diện trong thư mục `tests/`:
- [`tests/test_global_pipeline.py`](tests/test_global_pipeline.py): Kiểm tra quy trình 4 giai đoạn toàn cục, cơ chế hủy ngang an toàn, bộ nhớ cache trường & niên khóa, tùy chọn xuất file, và kiểm tra chuyển đổi tự động 100% bảng mã TCVN3 và VNI-Windows.
- [`tests/test_theme.py`](tests/test_theme.py): Kiểm tra độ chuẩn hóa cỡ chữ $\ge 10\text{pt}$ và bảng màu High-Contrast của cả Light Mode và Dark Mode.
- [`tests/test_updater.py`](tests/test_updater.py): Kiểm tra thuật toán phân tích và so sánh số học chuỗi phiên bản ngữ nghĩa (Semantic Versioning).

**Chạy toàn bộ bài test:**
```bash
uv run --with openpyxl python -m unittest discover tests
```

---

## ⚙️ HƯỚNG DẪN ĐÓNG GÓI BẰNG GITHUB ACTIONS

Dự án đã được tích hợp sẵn quy trình CI/CD hoàn toàn tự động tại [`.github/workflows/release.yml`](.github/workflows/release.yml):

- **Kích hoạt tự động**: Mỗi khi có commit mới được `push` lên nhánh `main` (hoặc `master`).
- **Máy chủ ảo Windows (`windows-latest`)**: Tự động cài đặt Python 3.11, OpenPyXL, PyInstaller và biên dịch thành tệp thực thi duy nhất `SohoaTHCS.exe` (`--onefile --windowed`).
- **Máy chủ ảo macOS (`macos-latest`)**: Tự động biên dịch thành ứng dụng `SohoaTHCS.app` và đóng gói thành `SohoaTHCS_macOS.zip` cũng như đĩa cài đặt chuẩn `SohoaTHCS_macOS.dmg`.
- **Tự động xuất bản GitHub Release**: Tải cả 3 tệp nhị phân lên mục [Releases](https://github.com/hungphungvan/so_hoa_vbcc_thcs/releases/latest) với tag `latest` để người dùng tải về sử dụng ngay lập tức mà không cần cài đặt Python.
