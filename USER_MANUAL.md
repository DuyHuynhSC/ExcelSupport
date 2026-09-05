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
16. [Sao Chép & Dán Vùng Lọc (Copy & Paste Visible Cells Only)](#16-sao-chép--dán-vùng-lọc-copy--paste-visible-cells-only)
17. [Chuyển Đổi Ngôn Ngữ Thanh Ribbon (Language Settings)](#17-chuyển-đổi-ngôn-ngữ-thanh-ribbon-language-settings)
18. [Thống Kê & Đếm Trang Thiết Kế 2.0 (Design Page Counter 2.0)](#18-thống-kê--đếm-trang-thiết-kế-20-design-page-counter-20)
19. [Bộ Công Cụ Cơ Sở Dữ Liệu Oracle (Oracle Database Tools)](#19-bộ-công-cụ-cơ-sở-dữ-liệu-oracle-oracle-database-tools)
20. [Bác Sĩ Công Thức & Sửa Lỗi Tự Động (AI Formula Doctor)](#20-bác-sĩ-công-thức--sửa-lỗi-tự-động-ai-formula-doctor)
21. [Sao Lưu & Khôi Phục Dữ Liệu Tức Thì (Sheet Snapshot & Instant Undo)](#21-sao-lưu--khôi-phục-dữ-liệu-tức-thì-sheet-snapshot--instant-undo)
22. [Bộ Tiện Ích Chuyên Sâu IT / Khách Hàng Nhật Bản (Japan & IT Tools)](#22-bộ-tiện-ích-chuyên-sâu-it--khách-hàng-nhật-bản-japan--it-tools)

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
* **Phím Tắt Bật/Tắt Nhanh:** Nhấn **`Ctrl + Shift + W`** bất cứ lúc nào trong Excel để đóng hoặc mở nhanh thanh Task Pane này.
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

## 18. Thống Kê & Đếm Trang Thiết Kế 2.0 (Design Page Counter 2.0)

Công cụ chuyên dụng dành cho các dự án Offshore, Outsourcing và Quản trị dự án phần mềm để nghiệm thu khối lượng tài liệu thiết kế (Basic Design, Detail Design, Test Plan, Database Schema...).

### 18.1 Thuật Toán Định Mức Ký Tự & Tô Màu Đối Chiếu (Evidence)
* **Bảo Vệ File Gốc 100%:** Khi bắt đầu đếm, hệ thống tự động tạo một file bản sao an toàn tại thư mục tạm (`%TEMP%\ExcelSupport_DesignPages\Evidence_...xlsx`). File gốc của bạn không bị sửa đổi bất kỳ nội dung nào.
* **So Sánh Siêu Tốc Trên Bộ Nhớ RAM:** Đọc toàn bộ bảng tính thiết kế và bảng tính Template vào mảng 2D trong RAM, quét so khớp từng ô (`cell-by-cell`) để xác định:
  - Các ô có nội dung mới hoặc bị thay đổi so với Template.
  - Tổng số lượng ký tự thực tế tại các ô thay đổi (`Total Changed Characters`).
  - Số lượng hình vẽ, sơ đồ kiến trúc, UI layout mới được chèn thêm (`Added Shapes / Diagrams`).
* **Tô Màu Highlight Trực Quan (Evidence File):** Tự động tô màu highlight các ô thay đổi trên bản sao với các gam màu dịu mắt (Vàng, Xanh Pastel, Xanh lá, Cam nhạt).
* **Công Thức Quy Đổi Trang Tiêu Chuẩn:**
  $$\text{Số trang quy đổi} = \frac{\text{Tổng ký tự ô thay đổi}}{\text{Định mức ký tự / trang}} + (\text{Số hình vẽ mới} \times \text{Hệ số trang/sơ đồ})$$

### 18.2 Các Chế Độ Đếm & Thao Tác Siêu Tốc:
* **Chế độ 1: Đếm theo màu ô tự tô (Khuyên dùng - Manual User-Highlighted Cells):**
  - **Quy trình làm việc siêu tốc:**
    1. Bấm nút **`📝 Mở Bản Sao Mới Để Tô Màu`**: Hệ thống tự động nhân bản file thiết kế thành một file New và mở trực tiếp trên Excel.
    2. Bôi đen các vùng ô bạn đã thiết kế và nhấn phím tắt **`Ctrl + Shift + H`** để tô màu đánh dấu tức thì theo màu đang chọn trong ComboBox (hoặc dùng công cụ Fill Color trên Excel).
    3. **Xóa màu nhanh:** Nếu tô nhầm, quét chọn ô và nhấn phím tắt **`Ctrl + Shift + Alt + H`** (hoặc bấm nút **`🧹 Xóa màu ô đã chọn`**) để xóa sạch màu nền và khôi phục viền ô.
    4. Quay lại hộp thoại và bấm **`🔍 Phân Tích & Đếm Trang`**: Hệ thống sẽ quét các ô có màu chỉ định, đếm tổng ký tự và quy đổi ra số trang thiết kế tương ứng.
* **Chế độ 2: Tự động so sánh với Template gốc (Auto-Diff with Template):**
  - Tự động so khớp từng ô giữa Target Workbook và Template Workbook của khách hàng.
  - Tự động đếm các ô thay đổi, tự động tạo file bản sao **Evidence** đã tô màu trực quan.
* **Chế độ 3: Đếm theo ngắt trang in Excel (Print Breaks Grid):**
  - Phân tích theo lưới ngắt trang in của Excel.

### 18.3 Quản Lý Cấu Hình Mẫu Dự Án (Project Profile Presets) & Biểu Đồ Dashboard:
* **Cấu Hình Mẫu Dự Án Đa Dạng (Project Profile Presets):**
  - **Chuẩn Tiếng Nhật Tiêu Chuẩn (JIS Standard):** 600 ký tự/trang, 0.5 trang/sơ đồ (chuẩn nghiệm thu khách hàng Nhật).
  - **Dự Án Offshore (VN/EN Text Heavy):** 1.200 ký tự/trang, 0.5 trang/sơ đồ.
  - **Tài Liệu Backend / DB Spec (Code/Query/DDL):** 800 ký tự/trang, 0.25 trang/sơ đồ.
  - **Giao Diện Web / Mobile UI (Diagram Heavy):** 500 ký tự/trang, 1.0 trang/sơ đồ.
  - **Tùy chỉnh riêng:** Cho phép tự nhập định mức ký tự và hệ số quy đổi hình ảnh bất kỳ.
* **Báo Cáo Kèm Biểu Đồ Trực Quan Trong Excel (Charts Dashboard):**
  - Khi bấm **`📊 Xuất Báo Cáo Ra Excel`**, hệ thống tự động vẽ **2 Biểu đồ chuyên nghiệp**:
    1. **Doughnut Chart:** Tỷ trọng khối lượng thiết kế mới vs Khung mẫu Template ban đầu.
    2. **Clustered Column / Bar Chart:** Phân bổ khối lượng trang quy đổi chi tiết theo từng Sheet.

---

## 19. Bộ Công Cụ Cơ Sở Dữ Liệu Oracle (Oracle Database Tools)

Hỗ trợ các kỹ sư phát triển phần mềm, DBA và chuyên viên phân tích dữ liệu tương tác trực tiếp với cơ sở dữ liệu Oracle từ Excel.

### 19.1 Quick SQL Query (Phím Tắt: `Ctrl + Shift + Q`)
* **Thực Thi & Xem Trước An Toàn (Preview First):**
  1. Nhập cấu hình kết nối (Host, Port, SID / Service Name, User, Password).
  2. Viết câu lệnh SQL Query.
  3. Bấm **`▶ Thực Thi (Execute)`** $\rightarrow$ Dữ liệu kết quả được tải và hiển thị ngay trên lưới DataGrid xem trước để kiểm tra tính chính xác.
  4. Sau khi xem trước OK, bấm **`📥 Chèn Vào Sheet (Insert to Excel)`** để ghi kết quả vào Sheet hiện tại.
* **Tùy Chỉnh Định Dạng Bảng:**
  - Tự động tạo Auto-Filter, kẻ viền (Borders), tự căn chỉnh độ rộng cột (`AutoFit Columns`).
  - Tùy chỉnh màu nền Header bảng kết quả với gam màu Pastel Cyan mặc định trang nhã.

### 19.2 Oracle Table Compare (So Sánh Cấu Trúc & Dữ Liệu)
* So sánh định nghĩa DDL của bảng (Cột, Kiểu dữ liệu, Độ dài, Nullable, Primary Key).
* So sánh dữ liệu giữa 2 Database hoặc giữa 2 Schema (Môi trường Development vs UAT/Production).
* Làm nổi bật các dòng dữ liệu bị lệch (Mismatch), dòng mới (Added) hoặc dòng bị xóa (Deleted).

---

## 20. Bác Sĩ Công Thức & Sửa Lỗi Tự Động (AI Formula Doctor)

### 20.1 Chẩn Đoán & Sửa Lỗi Công Thức
* **Tự Động Quét Lỗi:** Quét toàn bộ Sheet hoặc vùng chọn để phát hiện các ô bị lỗi công thức: `#N/A`, `#VALUE!`, `#REF!`, `#DIV/0!`, `#NAME?`, `#NUM!`, `#NULL!`, `#CALC!`.
* **AI Giải Thích & Sinh Công Thức Mới:** AI phân tích cấu trúc dữ liệu xung quanh và giải thích nguyên nhân gây lỗi, đồng thời tự động đề xuất công thức mới chính xác.
* **Áp Dụng Linh Hoạt:**
  - Bấm **`Sửa Ô Này`**: Sửa ngay lập tức ô đang chọn.
  - Bấm **`Sửa Hàng Loạt Cột Này (Batch Apply Fix)`**: Tự động nhân bản công thức đã sửa xuống toàn bộ các ô lỗi khác trong cùng một cột.

### 20.2 Hiện Đại Hóa Công Thức (Modernize Formula)
* Chuyển đổi các công thức lồng ghép phức tạp thế hệ cũ sang các hàm mảng động hiện đại:
  - Thay thế chuỗi `IF(ISNA(VLOOKUP(...)))` bằng `XLOOKUP(..., "Không tìm thấy")`.
  - Sử dụng hàm `LET` để đặt biến trung gian, tăng tốc độ tính toán gấp nhiều lần cho bảng tính dung lượng lớn.
  - Tự động hóa trích xuất danh sách duy nhất bằng hàm `UNIQUE` và `FILTER`.

---

## 21. Sao Lưu & Khôi Phục Dữ Liệu Tức Thì (Sheet Snapshot & Instant Undo)

Cơ chế an toàn tối thượng giúp bạn yên tâm thực hiện các tác vụ chỉnh sửa dữ liệu quy mô lớn mà không sợ mất dữ liệu.

### 21.1 Chụp Ảnh Snapshot Bảng Tính
* **Tự Động Sao Lưu:** Tự động tạo một bản snapshot lưu trên RAM trước khi thực thi các tác vụ hàng loạt (Xóa trùng lặp, Thay thế hàng loạt, Dọn dẹp dữ liệu, Gộp ô...).
* **Sao Lưu Thủ Công:** Bấm nút **`📸 Snapshot Bảng Tính`** trên Ribbon bất kỳ lúc nào để lưu lại trạng thái làm việc hiện tại.

### 21.2 Khôi Phục Dữ Liệu (Instant Restore)
* Mở hộp thoại Quản lý Snapshot $\rightarrow$ Chọn mốc thời gian cần khôi phục.
* Chọn:
  - **Khôi phục đè lên Sheet hiện tại:** Phục hồi nguyên vẹn dữ liệu, công thức và độ rộng cột trong tích tắc.
  - **Khôi phục sang Sheet Mới (`Restored_...`):** Giữ nguyên sheet hiện tại và tạo thêm một sheet mới chứa toàn bộ dữ liệu snapshot để so sánh đối chiếu.

---

## 22. Bộ Tiện Ích Chuyên Sâu IT / Khách Hàng Nhật Bản (Japan & IT Tools)

Nhóm công cụ đặc thù trên thanh Ribbon (`Tiện Ích Nhật & IT` / `Japan & IT Tools`) giải quyết triệt để các vấn đề chuẩn hóa dữ liệu, tài liệu thiết kế và trao đổi kỹ thuật cho các dự án phần mềm làm việc với khách hàng Nhật Bản.

### 22.1 Chuyển Đổi Toàn Giác ⇋ Bán Giác (Zenkaku ⇋ Hankaku Converter)
* **Vị trí:** Bấm nút **`🅰 全/半 Chuyển Đổi Zenkaku/Hankaku`** trên Ribbon.
* **Tính năng:**
  - **Chuyển sang Bán Giác (To Hankaku):**
    - Chữ số toàn giác $\rightarrow$ Chữ số chuẩn (`０-９` $\rightarrow$ `0-9`).
    - Chữ cái alphabet toàn giác $\rightarrow$ Bán giác (`Ａ-Ｚ`, `ａ-ｚ` $\rightarrow$ `A-Z`, `a-z`).
    - Dấu cách tiếng Nhật toàn giác $\rightarrow$ Dấu cách thường (`\u3000` $\rightarrow$ ` `).
    - Dấu câu, ký hiệu đặc biệt (`！＠＃` $\rightarrow$ `!@#`).
    - Chuyển Katakana toàn giác sang bán giác (tùy chọn).
  - **Chuyển sang Toàn Giác (To Zenkaku):**
    - Chữ số, alphabet, dấu cách sang toàn giác (`0-9` $\rightarrow$ `０-９`, `A-Z` $\rightarrow$ `Ａ-Ｚ`).
    - Ghép chuẩn âm đục/bán đục Katakana bán giác sang toàn giác (`ｶﾞ` $\rightarrow$ `ガ`, `ﾊﾟ` $\rightarrow$ `パ`).
* **Xem trước trực quan:** Tích hợp ô Live Interactive Preview để xem kết quả chuyển đổi mẫu ngay tức thì trước khi áp dụng vào bảng tính.

### 22.2 Rà Soát Chuẩn Từ Vựng Katakana & Trường Âm (Katakana Spell & Chouon Validator)
* **Vị trí:** Bấm nút **`🈁 Rà Soát Katakana`** trên Ribbon.
* **Vấn đề giải quyết:** Trong tài liệu kỹ thuật tiếng Nhật, sự không nhất quán về trường âm (Chouon `ー`) thường gây lỗi nghiệm thu (Ví dụ: chỗ viết `サーバー` chỗ lại viết `サーバ`, `ユーザー` vs `ユーザ`, `コンピューター` vs `コンピュータ`).
* **Tính năng:**
  - **Quét & Gom nhóm thông minh:** Quét toàn bộ vùng chọn hoặc Sheet, tự động gom các biến thể cùng gốc từ Katakana thành từng nhóm.
  - **Chuẩn hóa 1-Click theo JIS hoặc Custom:**
    - **Chuẩn JIS Tiêu chuẩn:** Tự động chuẩn hóa về từ có/không có trường âm theo quy tắc JIS.
    - **Chuẩn hóa tùy chọn:** Chọn biến thể chuẩn mong muốn và bấm **`⚡ Chuẩn Hóa Cụm Này`** hoặc **`⚡ Chuẩn Hóa Tất Cả`** để thay thế đồng loạt trong toàn bộ bảng tính.
  - **Định vị nhanh:** Nhấp đúp vào dòng kết quả để nhảy ngay tới ô chứa từ sai lệch trên sheet.

### 22.3 Trích Xuất Bảng Sang Markdown & HTML (Table to Markdown & HTML Exporter)
* **Phím Tắt:** **`Ctrl + Shift + M`** (hoặc bấm nút **`📑 Xuất Markdown/HTML`** trên Ribbon).
* **Tính năng:**
  - Chuyển đổi vùng ô đang chọn trong Excel thành mã **Markdown Table** (dùng cho GitHub, GitLab, Jira, Confluence, Backlog, Slack) hoặc **HTML Table** (`<table>...</table>`).
  - **Căn lề thông minh (Smart Alignment):** Tự động nhận diện cột số và căn phải (`---:`), cột chữ căn trái (`:---`).
  - **Xử lý xuống dòng an toàn:** Tự động chuyển đổi các ký tự xuống dòng `\n` trong ô thành thẻ `<br>` để bảo toàn cấu trúc bảng Markdown.
  - **Hộp thoại Tabbed Preview:** Xem trước giao diện bảng Markdown và HTML trực tiếp, kèm nút **`📋 Copy Markdown`** và **`📋 Copy HTML`** 1-Click sao chép vào Clipboard.

---

*Tài liệu được cập nhật liên tục cùng các phiên bản phát hành mới của **ExcelSupport Add-In**.*
