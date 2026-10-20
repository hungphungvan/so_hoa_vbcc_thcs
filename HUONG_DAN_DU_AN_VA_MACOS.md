# CẨM NANG KIẾN TRÚC, CHỨC NĂNG & VẬN HÀNH DỰ ÁN SỐ HÓA VBCC THCS
*(Dành cho phiên làm việc trên macOS và môi trường đa nền tảng)*

---

## 1. TỔNG QUAN DỰ ÁN & MỤC TIÊU CỐT LÕI

Dự án là phần mềm đồ họa chuyên dụng (**GUI Desktop Application**) phục vụ công tác số hóa, làm sạch và chuẩn hóa dữ liệu cấp bằng tốt nghiệp THCS của các Phòng/Sở Giáo dục và Đào tạo theo đúng chuẩn biểu mẫu của Bộ GD&ĐT.

Toàn bộ hệ thống hiện tại được tối ưu hóa tập trung vào **2 Mô-đun duy nhất**:
1. **MÔ-ĐUN 1: CHUẨN HOÁ DỮ LIỆU CẤP BẰNG THCS** (Biểu mẫu 28 cột theo `data/Mau_DS_Cap_Bang_THCS.xlsx`).
2. **MÔ-ĐUN 2: XUẤT DANH SÁCH TÀI LIỆU SỐ HÓA** (Biểu mẫu 11 cột theo `data/Mau_DS_Tai_Lieu_So_Hoa.xlsx`).

---

## 2. CẤU TRÚC THƯ MỤC & CÁC TỆP NGUỒN CỐT LÕI

```text
so_hoa_vbcc_thcs/
├── data/                                 # Thư mục dữ liệu danh mục & template chuẩn
│   ├── Danh_Muc_Tra_Cuu.xlsx             # 5 sheet: DS_Don_Vi (378 đơn vị kèm Xã & Mã trường), Gioi_Tinh, Dan_Toc, Quoc_Tich, Noi_Sinh
│   ├── Mau_DS_Cap_Bang_THCS.xlsx         # Mẫu xuất 28 cột (Sheet 'Data')
│   └── Mau_DS_Tai_Lieu_So_Hoa.xlsx       # Mẫu xuất 11 cột (Sheet 'Data')
│
├── src/                                  # Mã nguồn Python xử lý nghiệp vụ
│   ├── lookup_manager.py                 # Nạp danh mục động, nạp Đơn vị kèm Xã/phường & Mã, ánh xạ Mã trường
│   ├── normalizer.py                     # Thuật toán chuẩn hóa chuỗi, Title Case, nơi sinh, đảo MM/DD, auto-fill ngày cấp
│   ├── exporter.py                       # Động cơ xuất file: Gộp, Tách, DS_Lỗi, Dữ liệu gốc, Tài liệu số hóa, ép TEXT (@)
│   ├── excel_processor.py                # Quét đệ quy 3 cấu trúc thư mục, nhận diện ngữ cảnh, xử lý tương tác khối liên tiếp
│   └── gui/
│       ├── app_window.py                 # Giao diện chính Tkinter, quản lý đa luồng (Worker Thread), progress bar, thống kê
│       └── dialogs.py                    # Hộp thoại đối chiếu khối (Trường, Niên khóa, Ngày cấp, Ngày sinh) & Data Inspector
│
├── main.py                               # Điểm khởi chạy chương trình (Entry Point)
├── Chay_Chuong_Trinh.bat                 # File click đúp chạy nhanh trên Windows
├── Chay_Chuong_Trinh.command             # File click đúp chạy nhanh trên macOS
├── .github/workflows/release.yml         # CI/CD tự động đóng gói ứng dụng cho Windows (.exe) và macOS (.app/zip/dmg)
```

---

## 3. CHI TIẾT NGHIỆP VỤ & QUY TẮC HOẠT ĐỘNG CỦA TỪNG MÔ-ĐUN

### 🔷 MÔ-ĐUN 1: CHUẨN HOÁ DỮ LIỆU CẤP BẰNG (28 CỘT)

#### 1. Khả năng tương thích 3 cấu trúc thư mục nguồn (Quét đệ quy `rglob`)
Chương trình tự động quét sâu toàn bộ thư mục và nhận diện chính xác dù dữ liệu được tổ chức theo bất kỳ cách nào:
- **Cách 1:** Thư mục Đơn vị $\rightarrow$ các file Excel là Trường $\rightarrow$ các sheet là Niên khóa.
- **Cách 2:** Thư mục Đơn vị $\rightarrow$ các thư mục con là Trường $\rightarrow$ các file Excel là Niên khóa.
- **Cách 3:** Thư mục Đơn vị $\rightarrow$ các file Excel là Niên khóa $\rightarrow$ các sheet là Trường (hoặc chung 1 sheet).

#### 2. Cơ chế Fallback ngữ cảnh thông minh
Nếu trong bảng dữ liệu gốc, ô Tên trường (Cột N) hoặc Niên khóa (Cột O) bị bỏ trống:
- **Tên trường:** Tự động dò tìm từ Tên thư mục cha $\rightarrow$ Tên tệp $\rightarrow$ Tên sheet dựa trên danh mục trường trực thuộc.
- **Niên khóa:** Tự động trích xuất chuỗi năm (dạng `2018-2022`, `18-22`, `2020`) từ Tên file hoặc Tên sheet.

#### 3. Quy chuẩn làm sạch & chuẩn hóa dữ liệu
- **Họ và tên (Cột E):** Viết hoa chữ cái đầu mỗi từ (`vietnamese_title_case`), loại bỏ khoảng trắng thừa.
- **Dân tộc (Cột K), Quốc tịch (Cột L):** Tra mã số sang chữ chuẩn và Title Case.
- **Nơi sinh (Cột M):** Tra mã và Title Case, riêng các từ hành chính `"Thành phố"`, `"Thị xã"`, `"Thị trấn"` chỉ viết hoa chữ cái đầu tiên (chữ thứ hai viết thường: *Thành phố Hà Nội, Thị xã Sơn Tây*).
- **Ngày sinh (Cột I) & Ngày cấp bằng (Cột S):**
  - **Tự động đảo ngày tháng ngầm:** Khi phát hiện định dạng $MM/DD/YYYY$ ($A \le 12, B > 12$, ví dụ `03/14/2004`), chương trình **tự động chuyển thành `14/03/2004`** mà không cần người dùng xác nhận.
  - **Hộp thoại tương tác thông minh:** Nút *"Đảo Ngày ↔ Tháng"* chỉ sáng lên khi cả 2 số đều $\le 12$ (trường hợp nhập nhằng như `05/06/2004`).
- **Xử lý theo khối liên tiếp:** Khi một khối học sinh liên tiếp bị thiếu Tên trường hoặc Niên khóa, giao diện mở hộp thoại kèm bảng đối chiếu $\pm 5$ học sinh trước và sau khối để người dùng xác nhận 1-Click cho toàn khối.
- **Suy luận Niên khóa:** Tự động suy đoán niên khóa từ năm sinh phổ biến của học sinh: $Y_{ra} = Y_{birth} + 15, Y_{vao} = Y_{ra} - 4$.
- **Tự động điền Ngày cấp bằng (Cột S):** Gom nhóm 2 chiều `(Tên trường, Niên khóa)` trên toàn đơn vị để bù chéo ngày cấp bằng đã có cho các học sinh bị thiếu trong cùng nhóm.
- **Trạng thái số hóa (Cột Y) & Nguyên nhân (Cột Z):**
  - `"Chưa số hoá"`: Khi thiếu hẳn ngày sinh $\rightarrow$ Cột Z: `"Thiếu ngày sinh"`.
  - `"Số hoá chưa đầy đủ, chính xác"`: Khi ngày sinh sai định dạng hoặc sai logic $\rightarrow$ Cột Z: chi tiết lỗi cụ thể.
  - `"Đã số hóa đầy đủ, chính xác"`: Khi dữ liệu hợp lệ $\rightarrow$ Cột Z để trống.
- **Đính kèm tọa độ truy vết nguồn gốc (Cột X - Ghi chú):**
  Tự động gắn tag: `[Nguồn: {Tên_File} | Sheet: {Tên_Sheet} | Dòng {Dòng_Gốc}]`.
- **Ép định dạng TEXT (`@`) tuyệt đối:**
  Toàn bộ Cột I (Ngày sinh), Cột S (Ngày cấp), và các cột mã số được cấu hình `cell.number_format = '@'` và `cell.data_type = 's'`, ngăn chặn 100% lỗi Excel tự parse hoặc tự đổi định dạng ngày tháng theo cài đặt vùng miền (Locale) của hệ điều hành.

#### 4. Hệ thống xuất file đa tùy chọn
- **File Gộp:**
  - `{Ten_Don_Vi}_Đa_Gop_Theo_Truong.xlsx` (sắp xếp: Trường $\rightarrow$ Năm $\rightarrow$ STT).
  - `{Ten_Don_Vi}_Đa_Gop_Theo_Nam.xlsx` (sắp xếp: Năm $\rightarrow$ Trường $\rightarrow$ STT).
- **File Tách:**
  - `Da_Tach_Theo_Truong/{Ten_Truong}_Tu_{Min}_Den_{Max}.xlsx`.
  - `Da_Tach_Theo_Nam/{Ten_Don_Vi}_{Nam}.xlsx`.
  - `Tach_Theo_Truong_Va_Nam/{Ten_Truong}/{Ten_Truong}_{Nam}.xlsx`.
- **File Danh sách lỗi duy nhất:** `{Ten_Don_Vi}_DS_Loi.xlsx` mở rộng thêm 4 cột chỉ đường (Cột AC..AF: *Tệp nguồn, Sheet nguồn, Dòng Excel gốc, STT ban đầu*).
- **File Dữ liệu gốc tổng hợp:** `{Ten_Don_Vi}_Du_Lieu_Goc.xlsx` lưu nguyên vẹn dữ liệu gốc ban đầu.
- **Màn hình Data Inspector:** Tra cứu nhanh học sinh và đối chiếu trực quan song song giữa **[Dữ liệu gốc]** và **[Dữ liệu chuẩn hóa]** trực tiếp trên phần mềm.

---

### 🔷 MÔ-ĐUN 2: XUẤT DANH SÁCH TÀI LIỆU SỐ HÓA (11 CỘT)

Sử dụng biểu mẫu chuẩn `data/Mau_DS_Tai_Lieu_So_Hoa.xlsx` gồm 11 cột:
* Cột A: STT (tự tăng 1, 2, 3...)
* Cột B: Mã đơn vị (lấy từ Cột AA)
* Cột C: Tên đơn vị (lấy từ Cột AB)
* Cột D: Mã sổ gốc (tự động sinh)
* Cột E: Tên tài liệu
* Cột F: Từ năm
* Cột G: Đến năm
* Cột H: Ngày lập sổ (ngày hiện tại `dd/MM/yyyy`, định dạng TEXT `@`)
* Cột I: Số trang (để trống)
* Cột J: Tình trạng tài liệu (`"Nguyên vẹn"`)
* Cột K: Thông tư (để trống)

#### 3 Tùy chọn xuất tài liệu số hóa:
1. **Theo Trường (`{Ten_Don_Vi}_DS_Tai_Lieu_So_Hoa_Theo_Truong.xlsx`):**
   - Mã sổ gốc: `[Mã đơn vị]_[Mã trường]` *(Mã trường tra từ sheet `DS_Don_Vi`)*.
   - Tên tài liệu: `"Danh sách cấp bằng THCS trường {Tên trường} từ {Năm min} đến {Năm max}"` *(hoặc `năm {Năm}` nếu chỉ có 1 năm)*.
2. **Theo Năm (`{Ten_Don_Vi}_DS_Tai_Lieu_So_Hoa_Theo_Nam.xlsx`):**
   - Mã sổ gốc: `[Mã đơn vị]_[Năm]`.
   - Tên tài liệu: `"Danh sách cấp bằng THCS đơn vị {Tên đơn vị} năm {Năm}"`.
   - Từ năm = Đến năm = `Năm`.
3. **Theo Trường và Năm (`{Ten_Don_Vi}_DS_Tai_Lieu_So_Hoa_Theo_Truong_Va_Nam.xlsx`):**
   - Mã sổ gốc: `[Mã đơn vị]_[Mã trường]_[Năm]`.
   - Tên tài liệu: `"Danh sách cấp bằng THCS trường {Tên trường} năm {Năm}"`.
   - Từ năm = Đến năm = `Năm`.

---

## 4. HƯỚNG DẪN THIẾT LẬP & VẬN HÀNH TRÊN MACOS

### Bước 1: Sao chép mã nguồn sang macOS
Bạn có thể chuyển thư mục dự án qua AirDrop/USB hoặc clone từ GitHub:
```bash
git clone https://github.com/hungphungvan/so_hoa_vbcc_thcs.git
cd so_hoa_vbcc_thcs
```

### Bước 2: Cài đặt môi trường Python trên macOS
macOS đã có sẵn môi trường Python hoặc bạn có thể cài đặt dễ dàng qua Homebrew:
```bash
# Cài Python và Tkinter (nếu chưa có)
brew install python python-tk
```
Hoặc sử dụng công cụ quản lý siêu tốc `uv`:
```bash
curl -LsSf https://astral.sh/uv/install.sh | sh
```

### Bước 3: Khởi chạy phần mềm
Có 3 cách chạy phần mềm trên macOS:

* **Cách 1 (Nhanh nhất - Click đúp):**
  Click đúp chuột vào file `Chay_Chuong_Trinh.command` trong thư mục dự án.
  *(Nếu macOS báo file chưa có quyền thực thi, mở Terminal chạy 1 lần: `chmod +x Chay_Chuong_Trinh.command`)*.

* **Cách 2 (Chạy qua công cụ UV):**
  ```bash
  uv run --with openpyxl python main.py
  ```

* **Cách 3 (Chạy qua Python thông thường):**
  ```bash
  pip install openpyxl
  python3 main.py
  ```

---

## 5. MỘT SỐ LƯU Ý ĐẶC THÙ KHI CHẠY TRÊN MACOS

1. **Hiển thị giao diện Tkinter trên macOS:**
   - Tkinter trên macOS sử dụng native theme **Aqua** của Apple, các nút bấm, combobox và thanh cuộn sẽ hiển thị bo tròn theo phong cách macOS.
   - Các phím tắt sao chép/dán trong ô tìm kiếm: Trên macOS sử dụng phím `Cmd + C`, `Cmd + V` hoặc `Ctrl + C`, `Ctrl + V`.

2. **Xử lý tệp ẩn hệ thống của macOS (`.DS_Store`, `._*`):**
   - macOS thường tự sinh các file ẩn như `.DS_Store` hoặc file bóng `._filename.xlsx` khi duyệt thư mục.
   - Hệ thống quét đệ quy của chương trình (`excel_processor.py`) **đã được tích hợp bộ lọc tự động bỏ qua toàn bộ các tệp bắt đầu bằng `._` hoặc `.DS_Store`**, đảm bảo không bao giờ bị lỗi đọc nhầm file rác của macOS.

3. **Mở thư mục kết quả trên macOS:**
   - Nút *"📂 Mở thư mục Kết quả số hóa"* trên giao diện đã được lập trình nhận diện hệ điều hành: Trên macOS sẽ tự động kích hoạt lệnh `open {output_dir}` để mở trực tiếp trong ứng dụng **Finder**.

4. **Bản đóng gói nhị phân độc lập (.dmg, .app hoặc .zip):**
   - Kịch bản GitHub Actions (`.github/workflows/release.yml`) đã thiết lập sẵn luồng build cho macOS runner (`macos-latest`). Khi workflow chạy xong trên GitHub, bạn chỉ cần tải file `SohoaTHCS_macOS.dmg` hoặc `SohoaTHCS_macOS.zip` từ mục **Releases**, mở file `.dmg` và kéo vào `Applications` để chạy trực tiếp mà không cần cài đặt bất kỳ công cụ dòng lệnh nào.
   - *Mẹo Gatekeeper của Apple:* Nếu macOS hiển thị thông báo *"Ứng dụng chưa được xác minh nhà phát triển"*, bạn chỉ cần giữ phím `Control` + Click chuột vào ứng dụng $\rightarrow$ chọn **Open** $\rightarrow$ chọn tiếp **Open** là ứng dụng sẽ khởi chạy bình thường.

---
*Tài liệu được cập nhật tự động đồng bộ với phiên bản mã nguồn mới nhất của dự án.*
