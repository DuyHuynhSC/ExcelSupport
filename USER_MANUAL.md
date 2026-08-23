# 📖 SÁCH HƯỚNG DẪN SỬ DỤNG TOÀN DIỆN (USER MANUAL)
## 📊 ExcelSupport — AI Sheet Navigator & Copilot for Microsoft Excel

---

## 📑 MỤC LỤC
1. [Giới Thiệu Chung & Cài Đặt](#1-giới-thiệu-chung--cài-đặt)
2. [Giao Diện Điều Hướng (Sheet Navigator & Workbook Explorer)](#2-giao-diện-điều-hướng-sheet-navigator--workbook-explorer)
3. [So Sánh Hai Bảng Tính & Đối Chiếu Dữ Liệu (Diff & Compare Tool)](#3-so-sánh-hai-bảng-tính--đối-chiếu-dữ-liệu-diff--compare-tool)
4. [Lọc Nâng Cao Đa Tiêu Chí (Advanced Filter Pro)](#4-lọc-nâng-cao-đa-tiêu-chí-advanced-filter-pro)
5. [Dọn Dẹp & Chuẩn Hóa Dữ Liệu (Data Cleaning Wizard)](#5-dọn-dẹp--chuẩn-hóa-dữ-liệu-data-cleaning-wizard)
6. [Xử Lý Trùng Lặp & Gom Nhóm Thông Minh (Duplicate Finder)](#6-xử-lý-trùng-lặp--gom-nhóm-thông-minh-duplicate-finder)
7. [Xóa Dòng Trống & Gộp Dữ Liệu An Toàn (Batch Cleaner & Safe Merge)](#7-xóa-dòng-trống--gộp-dữ-liệu-an-toàn-batch-cleaner--safe-merge)
8. [Tìm & Thay Thế Hàng Loạt Theo Bảng Tra Cứu (Batch Find & Replace)](#8-tìm--thay-thế-hàng-loạt-theo-bảng-tra-cứu-batch-find--replace)
9. [Trộn & Ghép Nối Dữ Liệu Trực Quan (Visual XLOOKUP / Merge Table Wizard)](#9-trộn--ghép-nối-dữ-liệu-trực-quan-visual-xlookup--merge-table-wizard)
10. [Phát Hiện Dữ Liệu Bất Thường & Trùng Lặp Ảo (Fuzzy Duplicate Cleaner)](#10-phát-hiện-dữ-liệu-bất-thường--trùng-lặp-ảo-fuzzy-duplicate-cleaner)
11. [Bộ Quản Trị & Chuyển Đổi File Excel Hàng Loạt (Batch File Converter)](#11-bộ-quản-trị--chuyển-đổi-file-excel-hàng-loạt-batch-file-converter)
12. [Thước Ngắm Thông Minh & Bảng Thống Kê Nổi (Ruler Plus & Dynamic HUD)](#12-thước-ngắm-thông-minh--bảng-thống-kê-nổi-ruler-plus--dynamic-hud)
13. [Trợ Lý AI, Dịch Thuật Nhật - Việt & Công Thức Tự Động (AI Copilot)](#13-trợ-lý-ai-dịch-thuật-nhật---việt--công-thức-tự-động-ai-copilot)
14. [Quản Lý Liên Kết Ngoài & Rà Soát Tiếng Việt (Audit Tools)](#14-quản-lý-liên-kết-ngoài--rà-soát-tiếng-việt-audit-tools)
15. [Tùy Chỉnh Giao Diện Sáng / Tối (Dark & Light Theme)](#15-tùy-chỉnh-giao-diện-sáng--tối-dark--light-theme)

---

## 1. Giới Thiệu Chung & Cài Đặt

### 1.1 Giới Thiệu
**ExcelSupport** là Add-in mở rộng cao cấp dành cho Microsoft Excel, tối ưu hóa hiệu suất làm việc với bảng tính lớn, dữ liệu nhiều dòng/cột và đa Workbook.

### 1.2 Yêu Cầu Cài Đặt
* **Hệ điều hành:** Windows 10, Windows 11 (32-bit hoặc 64-bit).
* **Phiên bản Microsoft Excel:** Microsoft 365, Excel 2021, Excel 2019, Excel 2016.
* **Môi trường:** .NET Framework 4.8.

### 1.3 Cách Nạp File Add-in
1. **Mở Nhanh:** Nhấp đúp trực tiếp vào file `ExcelSupport-AddIn64-packed.xll` (cho Excel 64-bit) hoặc `ExcelSupport-AddIn-packed.xll` (cho Excel 32-bit).
2. **Cài Đặt Cố Định:**
   - Trong Excel, chọn menu **File** $\rightarrow$ **Options** $\rightarrow$ **Add-ins**.
   - Tại mục **Manage** ở dưới cùng, chọn **Excel Add-ins** rồi bấm **Go...**.
   - Bấm nút **Browse...** $\rightarrow$ Trỏ tới file `.xll` $\rightarrow$ Nhấn **OK**.

---

## 2. Giao Diện Điều Hướng (Sheet Navigator & Workbook Explorer)

### 2.1 Bảng Điều Hướng (Task Pane Bên Trái)
* Mặc định hiển thị cố định ở bên trái màn hình Excel giúp bạn quan sát song song với vùng nhập liệu.
* **Vùng Trên:** Danh sách các Workbook đang mở.
* **Vùng Dưới:** Danh sách các Sheet của Workbook đang chọn.

### 2.2 Các Thao Tác Thường Dùng
* **Tìm kiếm Sheet tức thì:** Nhập từ khóa vào ô tìm kiếm, danh sách Sheet sẽ lọc theo thời gian thực.
* **Sắp xếp Sheet A-Z / Z-A:** Sắp xếp theo thứ tự chữ cái tự nhiên (`Sheet1`, `Sheet2`, `Sheet10`).
* **Đổi màu Tab:** Chuột phải vào tên Sheet $\rightarrow$ Chọn màu sắc để đánh dấu phân loại.
* **Hiện tất cả Sheet ẩn:** Chuột phải vào Workbook $\rightarrow$ Chọn **"Hiện tất cả Sheet ẩn"** để mở lại toàn bộ các sheet `Hidden` hoặc `Very Hidden`.
* **Tạo Bảng Mục Lục:** Chuột phải vào Workbook $\rightarrow$ Chọn **"Tạo Bảng Mục Lục"** để tự động tạo Sheet Index liên kết Hyperlink đến từng sheet.

---

## 3. So Sánh Hai Bảng Tính & Đối Chiếu Dữ Liệu (Diff & Compare Tool)

### 3.1 Vị Trí Mở
Bấm nút **`📊 So Sánh Workbooks`** trên thanh Ribbon (Tab NAVIGATOR).

### 3.2 Các Chế Độ So Sánh
1. **LCS theo Dòng (Row LCS - Khuyên dùng):** Tự động nhận diện các dòng được chèn mới hoặc bị xóa mà không làm xô lệch các dòng bên dưới.
2. **LCS theo Cột (Column LCS):** Nhận diện cột mới chèn hoặc xóa.
3. **Theo Cột Khóa (Key Column):** So sánh giá trị của các dòng có cùng Mã định danh (Mã NV, Mã SP...).
4. **Theo Tọa Độ Ô (Cell-by-Cell):** So sánh từng ô tại cùng tọa độ `A1`, `B2`...

### 3.3 Tính Năng Nổi Bật
* **Click đúp nhảy tới ô:** Nhấp đúp vào dòng sai khác trong bảng kết quả, Excel sẽ tự động chuyển đến file, mở sheet và chọn đúng ô bị lệch.
* **Tô màu khác biệt:** Bấm nút **Tô màu trực quan** để đánh dấu các ô thay đổi (Vàng: Sửa, Xanh: Thêm, Đỏ: Xóa).
* **Xuất Báo Cáo:** Xuất kết quả ra Sheet mới kèm Hyperlink hoặc file CSV.

---

## 4. Lọc Nâng Cao Đa Tiêu Chí (Advanced Filter Pro)

### 4.1 Vị Trí Mở
Bấm nút **`🌪️ Lọc Nâng Cao`** trên thanh Ribbon.

### 4.2 Tính Năng
* **Lọc Theo Danh Sách (Whitelist / Blacklist):** Nhập hoặc dán danh sách hàng trăm mã cần giữ lại hoặc cần loại bỏ. Giao diện tự động ghi nhớ giá trị lọc trước đó sau khi đóng/mở lại.
* **Lọc Theo Loại Dữ Liệu:**
  - Lọc ô chứa công thức / Chỉ lọc ô giá trị tĩnh.
  - Lọc ô bị lỗi (`#N/A`, `#VALUE!`, `#REF!`).
  - Lọc theo màu nền ô (Color Filter).
* **Chế Độ Khớp:** Khớp chính xác (Exact), Khớp chứa (Contains), Khớp ký tự đại diện (`*`, `?`), hoặc Biểu thức chính quy (Regex).

---

## 5. Dọn Dẹp & Chuẩn Hóa Dữ Liệu (Data Cleaning Wizard)

### 5.1 Vị Trí Mở
Bấm nút **`🧹 Dọn Dẹp Dữ Liệu`** trên Ribbon.

### 5.2 Các Tác Vụ Chuẩn Hóa 1-Click
* **Xử lý khoảng trắng:** Cắt khoảng trắng thừa đầu/cuối (`Trim`), xóa khoảng trắng không ngắt dòng `&nbsp;` / `\u00A0`, xóa ký tự xuống dòng `\n`.
* **Chuyển đổi kiểu chữ:** `IN HOA`, `in thường`, `Viết Hoa Từng Từ` (Title Case: "nguyễn văn an" $\rightarrow$ "Nguyễn Văn An").
* **Chuyển tên Việt sang Katakana Nhật:** "Nguyễn Văn Ánh" $\rightarrow$ "グエン・ヴァン・アイン".
* **Xóa dấu tiếng Việt:** "Trần Thị Mai" $\rightarrow$ "Tran Thi Mai".
* **Sửa lỗi số lưu dạng Text:** Chuyển đổi hàng loạt chuỗi số dạng text thành số thực để tính toán hàm `SUM`, `VLOOKUP`.
* **Điền ô trống:** Tự động điền giá trị từ ô phía trên xuống (`Fill Down`).

---

## 6. Xử Lý Trùng Lặp & Gom Nhóm Thông Minh (Duplicate Finder)

### 6.1 Vị Trí Mở
Bấm nút **`🔍 Tìm Trùng Lặp`** trên Ribbon.

### 6.2 Tính Năng
* Chọn 1 hoặc nhiều cột khóa kết hợp.
* Tùy chọn so khớp chính xác hoặc so khớp mờ (Fuzzy).
* Gom các dòng trùng thành từng cụm (**Nhóm 1, Nhóm 2...**) với huy hiệu màu sắc.
* Hành động: Tô màu nhóm trên sheet, Xóa dòng trùng thừa, hoặc Tách toàn bộ dòng trùng sang Sheet mới để rà soát.

---

## 7. Xóa Dòng Trống & Gộp Dữ Liệu An Toàn (Batch Cleaner & Safe Merge)

### 7.1 Xóa Dòng & Cột Trống Hàng Loạt
* Bấm nút **`🗑️ Xóa Dòng Trống`** trên Ribbon.
* Tùy chọn: Xóa dòng hoàn toàn trống hoặc xóa dòng trống theo cột khóa.
* Áp dụng đồng thời cho **Sheet hiện tại**, **Nhiều Sheet được chọn** hoặc **Toàn bộ Workbook**.

### 7.2 Gộp Ô & Gộp Nhiều Sheet Bảo Toàn Dữ Liệu
* Bấm nút **`🔗 Gộp Ô & Sheet`** trên Ribbon.
* **Gộp Ô Bảo Toàn Dữ Liệu:** Gộp nhiều ô liền kề mà không làm mất nội dung các ô phụ (tùy chỉnh ký tự phân cách như dấu phẩy, dấu chấm phẩy, xuống dòng).
* **Gộp Nhiều Sheet:** Gom toàn bộ dữ liệu từ danh sách sheet vào 1 Sheet Tổng Hợp duy nhất, tự động bỏ qua dòng tiêu đề từ sheet thứ 2.

---

## 8. Tìm & Thay Thế Hàng Loạt Theo Bảng Tra Cứu (Batch Find & Replace)

### 8.1 Vị Trí Mở
Bấm nút **`🔄 Tìm & Thay Thế`** trên thanh Ribbon.

### 8.2 Hướng Dẫn Sử Dụng
1. Nhập các cặp `Cần tìm` $\rightarrow$ `Thay thế thành` vào bảng.
2. Hoặc bấm nút **"📥 Nạp Từ Vùng Chọn Excel"** để tự động import 2 cột tra cứu từ bảng tính.
3. Chọn phạm vi thay thế: Vùng đang chọn / Sheet hiện tại / Tất cả Sheet trong file / Tất cả các file đang mở.
4. Bấm **"⚡ Thực Thi Thay Thế Hàng Loạt"**. Kết quả sẽ báo cáo số lượng ô đã được thay thế thành công kèm tùy chọn tô màu đánh dấu.

---

## 9. Trộn & Ghép Nối Dữ Liệu Trực Quan (Visual XLOOKUP / Merge Table Wizard)

### 9.1 Vị Trí Mở
Bấm nút **`🔍 Ghép Bảng (Join)`** trên thanh Ribbon.

### 9.2 Hướng Dẫn Sử Dụng
1. **Chọn Bảng 1 (Bảng nguồn chính):** Chọn File, Sheet và Cột Khóa đối chiếu.
2. **Chọn Bảng 2 (Bảng tra cứu):** Chọn File, Sheet, Cột Khóa và đánh dấu các cột dữ liệu cần ghép sang.
3. **Chọn Kiểu Ghép:**
   - **Left Join:** Giữ nguyên toàn bộ dòng của Bảng 1 và bổ sung thêm thông tin từ Bảng 2.
   - **Inner Join:** Chỉ lấy những dòng có mã khóa xuất hiện trên cả 2 bảng.
   - **Full Outer Join:** Lấy toàn bộ dòng của cả 2 bảng.
4. **Nơi Xuất:** Trích xuất kết quả ra Sheet mới hoặc chèn trực tiếp vào các cột bên cạnh Bảng 1.
5. Bấm **"⚡ Thực Thi Trộn & Ghép Bảng"**.

---

## 10. Phát Hiện Dữ Liệu Bất Thường & Trùng Lặp Ảo (Fuzzy Duplicate Cleaner)

### 10.1 Vị Trí Mở
Bấm nút **`🧹 Trùng Lặp Ảo`** trên thanh Ribbon.

### 10.2 Hướng Dẫn Sử Dụng
1. Chọn Sheet và Cột dữ liệu cần quét.
2. Kéo thanh trượt **Độ tương đồng (Fuzzy Threshold)** (khuyến nghị 80% - 90%).
3. Bật tùy chọn: Xóa khoảng trắng ẩn NBSP `\u00A0`, Bỏ qua dấu tiếng Việt, Không phân biệt hoa/thường.
4. Bấm **"🔍 Bắt Đầu Quét Trùng Lặp Ảo"**. Hệ thống sẽ gom các giá trị gần giống nhau thành từng nhóm kèm % giống nhau.
5. Hành động 1-Click:
   - **⚡ Chuẩn Hóa Về Giá Trị Chuẩn:** Tự động sửa các biến thể gõ sai về giá trị chuẩn phổ biến nhất.
   - **🎨 Tô Màu Rà Soát:** Đổi màu các ô tương đồng trên sheet để kiểm tra thủ công.

---

## 11. Bộ Quản Trị & Chuyển Đổi File Excel Hàng Loạt (Batch File Converter)

### 11.1 Vị Trí Mở
Bấm nút **`📂 Chuyển Đổi File`** trên thanh Ribbon.

### 11.2 Các Chế Độ
1. **🔄 Chuyển đổi định dạng hàng loạt:** Chọn danh sách file $\rightarrow$ Chọn định dạng đích (`.xlsx`, `.xls`, `.xlsb`, `.csv`, `.pdf`) $\rightarrow$ Bấm thực thi.
2. **✂️ Tách từng Sheet thành file riêng:** Tự động xuất mỗi Sheet trong từng file Excel thành một file `.xlsx` độc lập.
3. **📚 Gộp nhiều file Excel thành 1:** Gom toàn bộ các file Excel trong danh sách thành 1 file duy nhất (mỗi file tương ứng với một Sheet).

---

## 12. Thước Ngắm Thông Minh & Bảng Thống Kê Nổi (Ruler Plus & Dynamic HUD)

### 12.1 Thước Ngắm Chữ Thập (Grid Ruler)
* Bấm nút **`🎯 Thước Ngắm Dòng/Cột`** trên Ribbon để bật/tắt.
* Khi di chuyển con trỏ ô, một dải màu bán trong suốt sẽ làm nổi bật toàn bộ dòng và cột của ô đang chọn.
* **Đổi màu sắc:** Menu *Tùy Chỉnh Thước* $\rightarrow$ Chọn 1 trong 7 màu dịu mắt (Vàng dịu, Xanh biển lơ, Xanh ngọc lục, Cam đào, Tím lavender, Hồng phấn, Xám thanh lịch).
* **Chế độ:** Cả Dòng & Cột / Chỉ Dòng / Chỉ Cột.

### 12.2 Bảng Thống Kê Nổi (Floating HUD)
* Menu *Tùy Chỉnh Thước* $\rightarrow$ Bấm **"Bảng Thống Kê Nổi"**.
* Bảng HUD nổi trên màn hình, tính toán thời gian thực:
  - **Dòng hiện tại:** Tổng, Trung bình, Số lượng ô, Giá trị Max, Min.
  - **Cột hiện tại:** Tổng, Trung bình, Số lượng ô, Giá trị Max, Min (tự động loại trừ các dòng bị ẩn/bị lọc bởi AutoFilter).
* **Chỉnh cỡ chữ động:** Bấm nút **`[A-]`** hoặc **`[A+]`** trên góc HUD để phóng to/thu nhỏ cỡ chữ tùy ý (từ 10.5pt đến 26pt).

---

## 13. Trợ Lý AI, Dịch Thuật Nhật - Việt & Công Thức Tự Động (AI Copilot)

### 13.1 Cấu Hình Kết Nối AI
* Vào Tab **`⚙️ Cài Đặt`** trên Task Pane:
  - API Base URL: `https://api.openai.com/v1` (hoặc URL Local AI như `http://localhost:11434/v1`).
  - Model Name: `gpt-4o`, `gpt-4o-mini`, `qwen-3.6`, `deepseek-r1`.
  - API Key: Nhập Key của bạn $\rightarrow$ Bấm **Kiểm Tra Kết Nối** $\rightarrow$ Bấm **Lưu Cấu Hình**.

### 13.2 Dịch Thuật Ô Nhật ⇋ Việt (Kèm Từ Điển Glossary)
* Chọn vùng ô cần dịch $\rightarrow$ Vào tab **AI Assistant** $\rightarrow$ Bấm **Dịch Thuật**.
* AI sẽ tự động tra cứu từ điển chuyên ngành trong **Glossary** để dịch chính xác theo chuẩn của doanh nghiệp.

### 13.3 AI Tự Động Viết & Sửa Lỗi Công Thức 1-Click
* **Viết công thức từ tiếng Việt:** Nhập yêu cầu tự nhiên (e.g. *"Lấy giá bán từ bảng B nếu mã hàng khớp"*), AI sẽ sinh công thức `XLOOKUP`/`INDEX MATCH` chuẩn xác kèm nút **"⚡ Chèn Vào Ô"**.
* **Sửa lỗi công thức 1-Click:** Chọn ô bị lỗi (`#N/A`, `#VALUE!`), bấm nút **`💡 AI Công Thức`** trên Ribbon, AI tự động phân tích ngữ cảnh sheet và hiển thị nút **"⚡ Áp Dụng Công Thức Đã Sửa"**.

---

## 14. Quản Lý Liên Kết Ngoài & Rà Soát Tiếng Việt (Audit Tools)

### 14.1 Quản Lý Liên Kết Ngoài (External Links Manager)
* Bấm nút **`🔗 Liên Kết Ngoài`** trên Ribbon.
* Liệt kê toàn bộ các công thức tham chiếu sang file Excel khác.
* Cung cấp các thao tác: Mở file nguồn, Cắt đứt liên kết (chuyển công thức thành giá trị tĩnh), Đổi đường dẫn nguồn mới.

### 14.2 Rà Soát Tiếng Việt (Vietnamese Auditor)
* Bấm nút **`🇻🇳 Kiểm Tra Tiếng Việt`** trên Ribbon.
* Quét tìm tất cả các vị trí chứa tiếng Việt có dấu trong nội dung ô, tên sheet và ghi chú (comments).
* Click đúp vào kết quả để nhảy tới ô hoặc xuất báo cáo ra Sheet mới / file CSV.

---

## 15. Tùy Chỉnh Giao Diện Sáng / Tối (Dark & Light Theme)

* Bấm nút toggle **`🌙 / ☀️`** ở góc trên thanh điều hướng để chuyển đổi giữa 2 chế độ:
  - **Light Theme:** Giao diện nền sáng trang nhã.
  - **Dark Theme:** Giao diện nền Slate tối (`#0F172A`), độ tương phản cao, tối ưu chống mỏi mắt khi làm việc ban đêm.
* Mọi cửa sổ, hộp thoại, DataGrid, ComboBox đều tự động đồng bộ theo Theme đã chọn.

---

## 16. Sao Chép & Dán Vùng Lọc (Copy & Paste Visible Cells Only)

Khi một bảng tính đang được áp dụng Bộ lọc (Filter) hoặc có các dòng bị ẩn (Hidden Rows):

### 16.1 Thao Tác 1-Click Nhanh Trên Ribbon
* **📋 Copy Ô Hiển Thị (Copy Visible Only):**
  1. Quét chọn vùng dữ liệu đang bật Filter.
  2. Bấm nút mũi tên dưới **`Copy & Dán Vùng Lọc`** $\rightarrow$ Chọn **"📋 Copy Ô Hiển Thị"**.
  3. Hệ thống chỉ sao chép các ô thực sự đang hiển thị trên màn hình vào Clipboard (loại bỏ hoàn toàn các dòng bị ẩn).
* **⚡ Dán Vào Ô Hiển Thị (Paste to Visible Cells):**
  1. Copy một danh sách dữ liệu (từ Excel hoặc nơi khác).
  2. Chọn ô đầu tiên trên cột đích đang bật Filter.
  3. Bấm nút mũi tên $\rightarrow$ Chọn **"⚡ Dán Vào Ô Hiển Thị"**.
  4. Hệ thống sẽ tự động nhảy cóc qua các dòng bị ẩn/bị lọc và chỉ dán dữ liệu vào các dòng đang hiển thị. Dữ liệu ở các dòng ẩn được bảo vệ an toàn 100%.

### 16.2 Hộp Thoại Hướng Dẫn Trực Quan (Filtered Copy & Paste Wizard)
* Bấm trực tiếp vào nút **`Copy & Dán Vùng Lọc`** trên Ribbon.
* Chọn **Vùng Nguồn (Source)** và **Vùng Đích (Destination)** trực tiếp bằng nút **"📍 Lấy Vùng Chọn"**.
* Lựa chọn kiểu dán linh hoạt:
  - **🔢 Chỉ Dán Giá Trị (Values Only):** Mặc định, giữ nguyên định dạng tại đích.
  - **📐 Dán Công Thức (Formulas):** Dán công thức kèm điều chỉnh tham chiếu.
  - **🎨 Chỉ Dán Định Dạng (Formats Only):** Dán màu nền, viền và font chữ.
  - **📑 Dán Toàn Bộ (All):** Dán cả giá trị và định dạng.
* Tùy chọn nâng cao:
  - **🔄 Lặp lại nguồn nếu đích nhiều dòng hơn:** Tự động lặp lại chuỗi giá trị nguồn vào toàn bộ các dòng lọc tại đích.
  - **⚪ Bỏ qua ô nguồn bị trống (Skip Blanks):** Không ghi đè nếu ô nguồn rỗng.
* Bấm **"🚀 Thực Thi Sao Chép & Dán Ngay"** để hoàn tất tức thì.

---

## 17. Chuyển Đổi Ngôn Ngữ Thanh Ribbon (Language Settings)

ExcelSupport Add-In hỗ trợ 3 ngôn ngữ giao diện chuẩn hóa:

* 🇻🇳 **Tiếng Việt** (Ngôn ngữ mặc định, thân thiện, dễ hiểu).
* 🇬🇧 **English** (Tiếng Anh chuyên ngành chuẩn Microsoft Office quốc tế).
* 🇯🇵 **日本語** (Tiếng Nhật chuẩn hóa thuật ngữ Microsoft Excel doanh nghiệp).

### Cách chuyển đổi ngôn ngữ:
1. Trên thanh Ribbon Tab **`NAVIGATOR`**, tìm nhóm **`Ngôn Ngữ & Cài Đặt (Language & Settings)`** ở góc phải.
2. Bấm vào nút menu **`🌐 Ngôn Ngữ (Language)`**.
3. Chọn một trong ba ngôn ngữ:
   - **🇻🇳 Tiếng Việt**
   - **🇬🇧 English**
   - **🇯🇵 日本語**
4. Toàn bộ thanh Ribbon (tên nhóm, tên nút, Screentip tóm tắt và Supertip giải thích chi tiết) sẽ được **chuyển đổi ngôn ngữ ngay lập tức mà không cần khởi động lại Excel**.
5. Cài đặt ngôn ngữ được tự động lưu và duy trì cho mọi phiên làm việc tiếp theo.

---

*Tài liệu được cập nhật liên tục cùng các phiên bản phát hành mới của **ExcelSupport Add-In**.*
