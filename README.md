# 📊 ExcelSupport — AI Sheet Navigator & Copilot for Microsoft Excel

[![.NET](https://img.shields.io/badge/.NET%20Framework-4.8-512BD4?style=flat&logo=dotnet)](https://dotnet.microsoft.com/)
[![Excel-DNA](https://img.shields.io/badge/Excel--DNA-1.8.0-107C41?style=flat&logo=microsoftexcel)](https://excel-dna.net/)
[![WPF](https://img.shields.io/badge/UI-WPF%20MVVM-0078D4?style=flat)](https://learn.microsoft.com/en-us/dotnet/desktop/wpf/)
[![Oracle](https://img.shields.io/badge/Database-Oracle%20Managed-F80000?style=flat&logo=oracle)](https://www.oracle.com/database/)
[![AI-Ready](https://img.shields.io/badge/AI-OpenAI%20%7C%20Qwen%203.6%20%7C%20DeepSeek-FF6F00?style=flat)](https://github.com/QwenLM)
[![License](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![User Manual](https://img.shields.io/badge/Documentation-USER__MANUAL.md-2563EB?style=flat)](USER_MANUAL.md)

**ExcelSupport** là một Add-in mở rộng chuyên nghiệp, hiệu năng cao dành cho **Microsoft Excel**, được xây dựng trên nền tảng **Excel-DNA**, **WPF MVVM** và tích hợp **AI Copilot (OpenAI / Qwen / DeepSeek)** cùng bộ công cụ quản trị dữ liệu **Oracle Database**. Add-in cung cấp giải pháp toàn diện giúp tối ưu hóa xử lý bảng tính lớn, nhiều Sheet, nhiều Workbook, hỗ trợ đếm trang thiết kế chính xác và nâng cao năng suất lập trình/nghiệp vụ.

> 📖 **Xem tài liệu hướng dẫn sử dụng chi tiết từ A-Z tại: [USER_MANUAL.md](USER_MANUAL.md)**

---

## ✨ Danh Sách Tính Năng Nổi Bật

### 📑 1. Điều Hướng & Quản Lý Bảng Tính (Sheet & Workbook Navigation)
* **Dock Task Pane Cố Định (Phím Tắt: `Ctrl + Shift + W`):** Thanh điều hướng tích hợp bên trái màn hình Excel giúp theo dõi song song danh sách Workbook và Worksheet. Bật/tắt nhanh bằng phím tắt `Ctrl + Shift + W`.
* **Tìm Kiếm & Sắp Xếp Tức Thì:** Tìm kiếm Sheet thời gian thực với bộ lọc đa từ khóa (`|`), hỗ trợ sắp xếp theo thứ tự **Original**, **A-Z ↑**, **Z-A ↓**.
* **Quản Trị Sheet Nhanh:** Đổi màu Tab, hiển thị toàn bộ Sheet ẩn (`Hidden` / `Very Hidden`), tự động tạo Sheet Mục Lục (Table of Contents) với Hyperlink trực tiếp.
* **Tối Ưu Chống Giật (Anti-Flicker & Zero Lag):** Cơ chế đồng bộ 1 chiều thông minh loại bỏ hoàn toàn hiện tượng nhấp nháy màn hình hoặc đơ Excel khi chuyển Sheet / Workbook.

### 📐 2. Thống Kê & Đếm Trang Thiết Kế (Design Page Counter 2.0)
* **Quy Trình & Chế Độ Đếm Linh Hoạt:**
  * **Chế độ 1: Đếm theo màu ô tự tô (Khuyên dùng):** Bấm nút **`📝 Mở Bản Sao Mới Để Tô Màu`** $\rightarrow$ Excel mở một file bản sao New độc lập $\rightarrow$ Quét chọn ô và tô màu nhanh bằng phím tắt **`Ctrl + Shift + H`** (hoặc xóa màu bằng **`Ctrl + Shift + Alt + H`**) $\rightarrow$ Add-in tự động tính ra số trang thiết kế chuẩn xác.
  * **Chế độ 2: Tự động so sánh với Template gốc:** Tự động so khớp cell-by-cell với file Template của khách hàng, đếm ký tự thay đổi và tự động xuất file Evidence đã tô màu trực quan.
  * **Chế độ 3: Đếm theo ngắt trang in Excel:** Phân tích theo lưới ngắt trang in (Print Breaks Grid).
* **Quản Lý Cấu Hình Dự Án (Project Profiles & Presets):**
  * Hỗ trợ chọn nhanh hoặc lưu các cấu hình định mức chuyên biệt:
    * **Tiêu chuẩn Nhật Bản (JIS Standard):** 600 ký tự / trang, Hình vẽ 0.5 trang.
    * **Tài liệu Tiếng Việt / Anh:** 1.200 ký tự / trang, Hình vẽ 0.5 trang.
    * **Tài liệu Kỹ Thuật / Backend Data:** 800 ký tự / trang, Hình vẽ 0.3 trang.
    * **Thiết kế Giao diện Web / UI:** 500 ký tự / trang, Hình vẽ 0.6 trang.
* **Báo Cáo Nghiệm Thu Trực Quan Đính Kèm Biểu Đồ (Charts Dashboard):**
  * Nút **`🎨 Mở File Đã Tô Màu (Evidence)`**: Mở trực tiếp file copy đã highlight để kiểm tra và dùng làm bằng chứng nghiệm thu (Proof of Work).
  * Nút **`📊 Xuất Báo Cáo Ra Excel`**: Tạo sheet tổng hợp đầy đủ KPI, kèm **Biểu đồ tròn (Pie Chart)** phân bổ tỷ lệ khối lượng Thiết Kế vs Template và **Biểu đồ cột (Bar Chart)** số trang giữa các Sheet.

### 🗾 3. Bộ Tiện Ích Chuyên Sâu IT / Khách Hàng Nhật Bản (Japan & IT Tools)
* **Bộ Chuyển Đổi Toàn Giác ⇋ Bán Giác (Zenkaku ⇋ Hankaku Converter):**
  * Chuyển đổi hai chiều giữa chữ số (`０-９` $\leftrightarrow$ `0-9`), chữ cái (`Ａ-Ｚ`, `ａ-ｚ` $\leftrightarrow$ `A-Z`, `a-z`), khoảng trắng Nhật (`\u3000` $\leftrightarrow$ ` `), Katakana (`アイウ` $\leftrightarrow$ `ｱｲｳ`, có xử lý ghép âm đục `ｶﾞ` $\leftrightarrow$ `ガ`, âm bán đục `ﾊﾟ` $\leftrightarrow$ `パ`), dấu câu và ký hiệu.
  * Áp dụng trên: Vùng ô chọn (Selection), Sheet hiện tại, hoặc Toàn bộ Workbook.
* **Rà Soát Chuẩn Từ Vựng Katakana (Katakana Spell & Chouon Validator):**
  * Tự động quét toàn bộ bảng tính để phát hiện các từ Katakana viết lệch chuẩn (đặc biệt là quy tắc trường âm `ー`: `サーバー` vs `サーバ`, `ユーザー` vs `ユーザ`, `コンピューター` vs `コンピュータ`, `フォルダー` vs `フォルダ`, v.v.).
  * Hộp thoại trực quan hiển thị vị trí ô, xem trước và 1-click chuẩn hóa đồng loạt toàn bộ tài liệu theo chuẩn JIS.
* **Trích Xuất Bảng Sang Markdown & HTML Table (Phím Tắt: `Ctrl + Shift + M`):**
  * Chuyển đổi vùng chọn Excel sang **Markdown Table** (căn lề số sang phải, xử lý ngắt dòng `<br>`, chống vỡ bảng) để paste tức thì vào Jira, Confluence, GitHub PR, Redmine, Notion.
  * Xuất sang **HTML Table** kèm CSS inline viền và màu sắc hiện đại.

### 🗄️ 3. Bộ Công Cụ Cơ Sở Dữ Liệu Oracle (Oracle Database Tools)
* **Quick SQL Query:**
  * Thực thi câu lệnh SQL trực tiếp từ Excel bằng phím tắt `Ctrl + Shift + Q`.
  * **Tách riêng Quy trình Thực thi & Chèn dữ liệu:** Chạy query $\rightarrow$ Xem trước dữ liệu trên DataGrid (Preview) $\rightarrow$ Xác nhận rồi mới chèn vào Sheet.
  * Tùy chọn định dạng bảng thông minh, tùy chỉnh màu nền Header (Header Background Color với mã màu Pastel Cyan dịu mắt).
* **Oracle Table Compare:**
  * So sánh cấu trúc bảng (DDL, cột, kiểu dữ liệu, khóa chính, độ dài) và nội dung dữ liệu giữa 2 Database Oracle hoặc 2 Schema khác nhau.
  * Hiển thị chi tiết các dòng lệch dữ liệu, dòng thêm mới hoặc bị thiếu.

### 🩺 4. Bác Sĩ Công Thức & Trợ Lý AI (AI Formula Doctor & Copilot)
* **AI Formula Doctor:**
  * Quét và phát hiện tự động toàn bộ ô bị lỗi công thức (`#N/A`, `#VALUE!`, `#REF!`, `#DIV/0!`, `#NAME?`, `#NUM!`, `#NULL!`, `#CALC!`).
  * AI chẩn đoán nguyên nhân gốc rễ và tự động sinh công thức sửa lỗi chính xác.
  * Áp dụng sửa lỗi 1-Click cho từng ô hoặc sửa hàng loạt tự động theo cột (Batch Apply Fix).
  * **Hiện Đại Hóa Công Thức (Modernize Formula):** Nâng cấp các công thức lồng ghép cũ (`VLOOKUP` + `IFERROR` + `INDEX/MATCH`) sang các hàm hiện đại (`XLOOKUP`, `LET`, `FILTER`, `UNIQUE`).
* **Trợ Lý Dịch Thuật & Soạn Thảo:**
  * Dịch thuật dữ liệu Nhật ⇋ Việt / Anh theo ngữ cảnh bảng tính.
  * Quản lý Từ điển Thuật ngữ dự án (Glossary) với cơ chế Import/Export CSV UTF-8 BOM & JSON.
  * Tự động viết công thức Excel từ mô tả tiếng Việt tự nhiên.

### 📸 5. Sao Lưu Tức Thì & Chống Mất Dữ Liệu (Sheet Snapshot & Instant Undo)
* Tự động hoặc thủ công chụp lại toàn bộ trạng thái Sheet (dữ liệu, công thức, định dạng số, độ rộng cột) trước khi chạy các tác vụ nặng.
* Khôi phục (Restore) lại nguyên trạng dữ liệu chỉ với 1 click, bảo vệ an toàn 100% dữ liệu.

### 🛠️ 6. Bộ Tiện Ích Xử Lý Bảng Tính Chuyên Sâu
* **Diff & Compare Tool:** So sánh 2 Workbook/Sheet với thuật toán Row LCS thông minh, phát hiện chèn/xóa dòng, click đúp nhảy tới ô sai lệch.
* **Advanced Filter Pro:** Lọc danh sách Whitelist/Blacklist hàng trăm dòng, lọc theo công thức, ô lỗi, màu sắc, Regex/Wildcard.
* **Data Cleaning Wizard:** Cắt khoảng trắng thừa, xóa NBSP `\u00A0`, chuẩn hóa chữ hoa/thường, xóa dấu tiếng Việt, chuyển tên Việt sang Katakana.
* **Duplicate Finder & Fuzzy Duplicate:** Gom nhóm dòng trùng theo đa cột khóa, nhận diện sai lệch chính tả/dấu cách mờ bằng thuật toán Levenshtein & Jaro-Winkler.
* **Batch Cleaner & Safe Merge:** Xóa dòng/cột trống đa Sheet, gộp ô bảo toàn dữ liệu, gộp nhiều Sheet vào 1 Sheet Tổng Hợp.
* **Batch Find & Replace:** Tìm kiếm và thay thế đồng thời hàng trăm từ khóa theo bảng tra cứu đối chiếu.
* **Visual Table Merge (XLOOKUP Wizard):** Ghép 2 bảng dữ liệu trực quan bằng Left Join, Inner Join, Full Outer Join mà không cần viết công thức phức tạp.
* **Batch File Converter:** Chuyển đổi định dạng hàng loạt (`.xlsx`, `.xls`, `.xlsb`, `.csv`, `.pdf`), tách/gộp file nền siêu tốc.
* **Filtered Copy & Paste:** Copy và dán an toàn chỉ vào các ô hiển thị (`Visible Cells Only`), tự động bỏ qua dòng bị ẩn/lọc.
* **Ruler Plus & Dynamic HUD:** Thước ngắm chữ thập 7 màu dịu mắt, bảng HUD nổi thống kê nhanh (Tổng, Trung bình, Đếm) cho dòng/cột hiển thị.

### 🌓 7. Giao Diện & Đa Ngôn Ngữ
* **Hỗ Trợ 100% Dark / Light Theme:** Chuyển đổi giao diện Sáng / Tối Slate (`#0F172A`) đồng bộ trên tất cả các cửa sổ, hộp thoại và Task Pane.
* **Đa Ngôn Ngữ (Multi-Language Ribbon & UI):** Hỗ trợ chuyển đổi mượt mà 3 ngôn ngữ: **Tiếng Việt 🇻🇳**, **English 🇬🇧**, **日本語 🇯🇵** (940+ localization keys).

---

## 🏗️ Cấu Trúc Mã Nguồn

```
ExcelSupport/
├── ExcelSupport.csproj               # Project SDK .NET Framework 4.8 (WPF + WinForms + Excel-DNA)
├── ExcelSupport-AddIn.dna            # Manifest khai báo Add-in & nhúng DLL phụ thuộc cho Excel-DNA
├── AddInEvents.cs                    # Lớp nạp chính (IExcelAddIn), hook sự kiện Excel COM & quản lý bộ nhớ
├── README.md                         # Tài liệu tổng quan dự án
├── USER_MANUAL.md                    # Sách hướng dẫn sử dụng chi tiết từ A-Z
│
├── Models/                           # Tầng Model dữ liệu
│   ├── TableMergeModels.cs           # Model ghép nối bảng trực quan (Join Types, Options)
│   ├── FuzzyDuplicateModels.cs       # Model nhận diện trùng lặp ảo & gom cụm (Fuzzy Clusters)
│   ├── BatchFileConverterModels.cs   # Model chuyển đổi và gộp/tách file hàng loạt
│   ├── BatchFindReplaceModels.cs     # Model tìm & thay thế hàng loạt
│   ├── OracleConnectionConfig.cs     # Cấu hình kết nối Oracle Database
│   ├── SheetSnapshotItem.cs          # Model dữ liệu snapshot bảng tính
│   ├── GlossaryItem.cs               # Model thuật ngữ từ điển Nhật ⇋ Việt
│   └── VietnameseLocationItem.cs     # Model vị trí phát hiện tiếng Việt
│
├── Services/                         # Tầng dịch vụ & Xử lý nghiệp vụ
│   ├── DesignPageCounterService.cs   # Thuật toán đếm trang thiết kế, so khớp ký tự & tô màu Evidence
│   ├── OracleDataCompareService.cs   # Dịch vụ kết nối, query & so sánh dữ liệu Oracle Database
│   ├── AiFormulaDoctorService.cs     # Dịch vụ chẩn đoán, sửa lỗi & tối ưu công thức Excel bằng AI
│   ├── SheetSnapshotService.cs       # Dịch vụ chụp ảnh snapshot & hoàn tác dữ liệu
│   ├── TableMergeService.cs          # Dịch vụ ghép bảng bằng mảng 2D siêu tốc
│   ├── FuzzyDuplicateService.cs      # Thuật toán Levenshtein & Jaro-Winkler nhận diện trùng lặp ảo
│   ├── BatchFileConverterService.cs  # Chuyển đổi định dạng, tách/gộp file nền
│   ├── BatchFindReplaceService.cs    # Dịch vụ tìm và thay thế theo bảng tra cứu
│   ├── GridRulerService.cs           # Quản lý thước kẻ chữ thập & tính toán thống kê dòng/cột hiển thị
│   ├── GlossaryService.cs            # Dịch vụ Import/Export Glossary (CSV UTF-8 BOM & JSON)
│   ├── LocalizationService.cs        # Quản lý đa ngôn ngữ (Tiếng Việt, Tiếng Anh, Tiếng Nhật)
│   └── OpenAiClientService.cs        # HTTP Client kết nối OpenAI / Local LLM
│
├── ViewModels/                       # Tầng ViewModel (Mô hình MVVM)
│   ├── TaskPaneViewModel.cs          # ViewModel chính điều khiển Task Pane & Theme toggle
│   ├── DesignPageCounterViewModel.cs # ViewModel đếm số trang thiết kế & quản lý Evidence
│   ├── AiAssistantViewModel.cs       # Quản lý logic Dịch thuật, Glossary, Sinh & Sửa công thức
│   └── AiSettingsViewModel.cs        # Quản lý kiểm tra kết nối & Lưu cấu hình AI
│
├── Views/                            # Giao diện người dùng WPF (XAML) - Hỗ trợ 100% Dark Theme
│   ├── DesignPageCounterDialog.xaml  # Hộp thoại Thống Kê & Đếm Trang Thiết Kế
│   ├── OracleQuickQueryDialog.xaml   # Hộp thoại Thực Thi & Xem Trước SQL Query Oracle
│   ├── OracleTableCompareDialog.xaml # Hộp thoại So Sánh Bảng & Dữ Liệu Oracle Database
│   ├── AiFormulaDoctorDialog.xaml    # Hộp thoại Bác Sĩ Công Thức & Sửa Lỗi Tự Động
│   ├── SheetSnapshotDialog.xaml      # Hộp thoại Quản Lý Snapshot & Khôi Phục Dữ Liệu
│   ├── VisualTableMergeDialog.xaml   # Hộp thoại Trộn & Ghép Nối Bảng Trực Quan
│   ├── FuzzyDuplicateDialog.xaml     # Hộp thoại Phát Hiện Dữ Liệu Bất Thường & Trùng Lặp Ảo
│   ├── BatchFileConverterDialog.xaml # Hộp thoại Quản Trị & Chuyển Đổi File Hàng Loạt
│   ├── BatchFindReplaceDialog.xaml   # Hộp thoại Tìm & Thay Thế Hàng Loạt
│   ├── RulerHudWindow.xaml           # Bảng thống kê nổi (HUD) phóng to/thu nhỏ cỡ chữ động
│   ├── WorkbookCompareDialog.xaml    # Hộp thoại So Sánh 2 Workbooks / Sheets
│   ├── AdvancedFilterDialog.xaml     # Hộp thoại Lọc Nâng Cao Đa Tiêu Chí
│   ├── DataCleaningWizardDialog.xaml # Hộp thoại Dọn Dẹp & Chuẩn Hóa Dữ Liệu
│   └── DuplicateFinderDialog.xaml    # Hộp thoại Tìm & Xử Lý Trùng Lặp
│
└── Ribbon/                           # Tùy biến thanh công cụ Excel Ribbon
    └── RibbonController.cs           # Tab NAVIGATOR trên Ribbon với đầy đủ các nút công cụ
```

---

## 🚀 Hướng Dẫn Biên Dịch & Đóng Gói (Build & Publish)

### Yêu cầu hệ thống:
* **Hệ điều hành:** Windows 10 / Windows 11 (64-bit hoặc 32-bit).
* **Microsoft Excel:** Office 2016, 2019, 2021, Microsoft 365.
* **Môi trường:** .NET SDK (hỗ trợ target `net48`).

### 1. Biên dịch dự án:
```bash
dotnet build
```

### 2. Đóng gói bản phát hành Release (Self-Contained XLL):
```bash
dotnet publish -c release
```

File Add-in độc lập hoàn chỉnh (đã nhúng sẵn toàn bộ `Newtonsoft.Json.dll` và `Oracle.ManagedDataAccess.dll` nén LZMA) nằm tại thư mục:
* Bản **Excel 64-bit:** `bin\release\net48\publish\ExcelSupport-AddIn64-packed.xll`
* Bản **Excel 32-bit:** `bin\release\net48\publish\ExcelSupport-AddIn-packed.xll`

> 💡 **Cài Đặt:** Bạn chỉ cần copy duy nhất file `.xll` này sang máy tính của người dùng để mở hoặc nạp vào Excel mà không cần cài đặt thêm DLL phụ thuộc.

---

## 📄 Bản Quyền (License)
Dự án được phân phối dưới giấy phép [MIT License](LICENSE).
