# 📊 ExcelSupport — AI Sheet Navigator & Copilot for Microsoft Excel

[![.NET](https://img.shields.io/badge/.NET%20Framework-4.8-512BD4?style=flat&logo=dotnet)](https://dotnet.microsoft.com/)
[![Excel-DNA](https://img.shields.io/badge/Excel--DNA-1.8.0-107C41?style=flat&logo=microsoftexcel)](https://excel-dna.net/)
[![WPF](https://img.shields.io/badge/UI-WPF%20MVVM-0078D4?style=flat)](https://learn.microsoft.com/en-us/dotnet/desktop/wpf/)
[![AI-Ready](https://img.shields.io/badge/AI-OpenAI%20%7C%20Qwen%203.6%20%7C%20DeepSeek-FF6F00?style=flat)](https://github.com/QwenLM)
[![License](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![User Manual](https://img.shields.io/badge/Documentation-USER__MANUAL.md-2563EB?style=flat)](USER_MANUAL.md)

**ExcelSupport** là một Add-in chuyên nghiệp, hiệu năng cao dành cho Microsoft Excel, được xây dựng trên nền tảng **Excel-DNA**, **WPF MVVM** và tích hợp **AI Copilot (OpenAI / Qwen 3.6 / DeepSeek)**. Add-in cung cấp bộ công cụ toàn diện giúp tối ưu hóa xử lý bảng tính lớn, nhiều Sheet và nhiều Workbook trong doanh nghiệp.

> 📖 **Xem sách hướng dẫn sử dụng toàn diện từ A-Z tại: [USER_MANUAL.md](USER_MANUAL.md)**

---

## ✨ Danh Sách Tính Năng Chính

| Nhóm Tính Năng | Công Cụ | Mô Tả Tóm Tắt |
| :--- | :--- | :--- |
| **📁 Điều Hướng & Quản Lý** | **Sheet Navigator** | Dock Task Pane bên trái, tìm kiếm Sheet tức thì, sắp xếp A-Z, đổi màu Tab, hiện tất cả Sheet ẩn, tạo Mục Lục tự động. |
| **📊 So Sánh & Đối Chiếu** | **Diff & Compare Tool** | So sánh 2 Workbook/Sheet với thuật toán Row LCS thông minh, phát hiện chèn/xóa dòng, click đúp nhảy tới ô, tô màu sai lệch. |
| **🌪️ Lọc Dữ Liệu** | **Advanced Filter Pro** | Lọc Whitelist/Blacklist hàng trăm dòng, lọc theo công thức, ô lỗi, màu sắc, Regex/Wildcard, tự ghi nhớ lịch sử lọc. |
| **🧹 Chuẩn Hóa & Dọn Dẹp** | **Data Cleaning Wizard** | Cắt khoảng trắng thừa, xóa NBSP `\u00A0`, chuẩn hóa chữ hoa/thường, xóa dấu tiếng Việt, chuyển tên Việt sang Katakana. |
| **🔍 Trùng Lặp & Gom Nhóm** | **Duplicate Finder** | Gom nhóm các dòng trùng lặp theo đa cột khóa, so khớp chính xác/mờ, tô màu nhóm, xóa dòng trùng hoặc tách sheet. |
| **🗑️ Dòng Trống & Gộp Ô** | **Batch Cleaner & Safe Merge** | Xóa dòng/cột trống đa Sheet, gộp các ô bảo toàn dữ liệu kèm ký tự phân cách, gộp nhiều Sheet thành 1 Sheet Tổng Hợp. |
| **🔄 Thay Thế Hàng Loạt** | **Batch Find & Replace** | Thay thế đồng thời hàng trăm từ khóa/mã sản phẩm theo bảng đối chiếu tra cứu, nạp nhanh 1-Click từ vùng chọn Excel. |
| **🔍 Trộn & Ghép Nối Bảng** | **Visual XLOOKUP Wizard** | Ghép 2 bảng dữ liệu trực quan bằng Left Join, Inner Join, Full Outer Join mà không cần viết công thức phức tạp. |
| **🧹 Trùng Lặp Ảo & Lỗi Gõ** | **Fuzzy Duplicate Cleaner** | Nhận diện lỗi chính tả, khác biệt dấu tiếng Việt, khoảng trắng vô hình bằng Levenshtein/Jaro-Winkler, chuẩn hóa 1-Click. |
| **📂 Quản Trị & Đổi File** | **Batch File Converter** | Chuyển đổi định dạng hàng loạt (`.xlsx`, `.xls`, `.xlsb`, `.csv`, `.pdf`), tách sheet thành file riêng, gộp nhiều file vào 1. |
| **🎯 Thước Kẻ & Bảng Thống Kê** | **Ruler Plus & Dynamic HUD** | Thước ngắm chữ thập 7 màu dịu mắt, bảng HUD nổi hiển thị Tổng, Trung bình, Số lượng ô (chỉ tính dòng hiển thị) kèm chỉnh cỡ chữ động `[A-]` `[A+]`. |
| **🤖 Trợ Lý AI & Dịch Thuật** | **AI Copilot & Formula Fixer** | Dịch thuật ô Nhật ⇋ Việt kèm từ điển Glossary bắt buộc, tự động viết công thức từ tiếng Việt, chẩn đoán & sửa lỗi công thức 1-Click. |
| **🔗 Kiểm Soát & Rà Soát** | **External Links & VN Auditor** | Quản lý và cắt đứt liên kết ngoài (External Links), rà soát toàn bộ tiếng Việt có dấu trong nội dung ô, tên sheet và comment. |
| **🌓 Giao Diện Hiện Đại** | **Dark / Light Theme** | Chuyển đổi linh hoạt giữa giao diện Sáng và Tối Slate (`#0F172A`), tương thích 100% trên toàn bộ các cửa sổ và hộp thoại. |

---

## 🏗️ Cấu Trúc Mã Nguồn

```
ExcelSupport/
├── ExcelSupport.csproj               # Project SDK .NET Framework 4.8 (WPF + WinForms + Excel-DNA)
├── ExcelSupport-AddIn.dna            # Manifest khai báo Add-in cho Excel-DNA
├── AddInEvents.cs                    # Lớp nạp chính (IExcelAddIn), hook sự kiện Excel COM & quản lý bộ nhớ
├── README.md                         # Tài liệu tổng quan dự án
├── USER_MANUAL.md                    # Sách hướng dẫn sử dụng chi tiết từ A-Z
│
├── Models/                           # Tầng Model dữ liệu
│   ├── TableMergeModels.cs           # Model ghép nối bảng trực quan (Join Types, Options)
│   ├── FuzzyDuplicateModels.cs       # Model nhận diện trùng lặp ảo & gom cụm (Fuzzy Clusters)
│   ├── BatchFileConverterModels.cs   # Model chuyển đổi và gộp/tách file hàng loạt
│   ├── BatchFindReplaceModels.cs     # Model tìm & thay thế hàng loạt
│   ├── GlossaryItem.cs               # Model thuật ngữ từ điển Nhật ⇋ Việt
│   └── VietnameseLocationItem.cs     # Model vị trí phát hiện tiếng Việt
│
├── Services/                         # Tầng dịch vụ & Xử lý nghiệp vụ
│   ├── TableMergeService.cs          # Dịch vụ ghép bảng bằng mảng 2D siêu tốc
│   ├── FuzzyDuplicateService.cs      # Thuật toán Levenshtein & Jaro-Winkler nhận diện trùng lặp ảo
│   ├── BatchFileConverterService.cs  # Chuyển đổi định dạng, tách/gộp file nền
│   ├── BatchFindReplaceService.cs    # Dịch vụ tìm và thay thế theo bảng tra cứu
│   ├── GridRulerService.cs           # Quản lý thước kẻ chữ thập & tính toán thống kê dòng/cột hiển thị
│   ├── GlossaryService.cs            # Dịch vụ Import/Export Glossary (CSV UTF-8 BOM & JSON)
│   └── OpenAiClientService.cs        # HTTP Client kết nối OpenAI / Local LLM
│
├── ViewModels/                       # Tầng ViewModel (Mô hình MVVM)
│   ├── TaskPaneViewModel.cs          # ViewModel chính điều khiển Task Pane & Theme toggle
│   ├── AiAssistantViewModel.cs       # Quản lý logic Dịch thuật, Glossary, Sinh & Sửa công thức
│   └── AiSettingsViewModel.cs        # Quản lý kiểm tra kết nối & Lưu cấu hình AI
│
├── Views/                            # Giao diện người dùng WPF (XAML) - Hỗ trợ 100% Dark Theme
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

### 2. Đóng gói bản phát hành Release (Tự Động Đóng Gói .xll):
```bash
dotnet publish -c release
```

File Add-in độc lập (standalone `.xll`) sẽ nằm tại thư mục:
* Bản **Excel 64-bit:** `bin\release\net48\publish\ExcelSupport-AddIn64-packed.xll`
* Bản **Excel 32-bit:** `bin\release\net48\publish\ExcelSupport-AddIn-packed.xll`

---

## 📄 Bản Quyền (License)
Dự án được phân phối dưới giấy phép [MIT License](LICENSE).
