# 📊 ExcelSupport — AI Sheet Navigator & Copilot for Microsoft Excel

[![.NET](https://img.shields.io/badge/.NET%20Framework-4.8-512BD4?style=flat&logo=dotnet)](https://dotnet.microsoft.com/)
[![Excel-DNA](https://img.shields.io/badge/Excel--DNA-1.8.0-107C41?style=flat&logo=microsoftexcel)](https://excel-dna.net/)
[![WPF](https://img.shields.io/badge/UI-WPF%20MVVM-0078D4?style=flat)](https://learn.microsoft.com/en-us/dotnet/desktop/wpf/)
[![AI-Ready](https://img.shields.io/badge/AI-OpenAI%20%7C%20Qwen%203.6-FF6F00?style=flat)](https://github.com/QwenLM)
[![License](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

**ExcelSupport** là một Add-in chuyên nghiệp dành cho Microsoft Excel, được phát triển trên nền tảng **Excel-DNA**, **WPF MVVM** và tích hợp **AI Copilot (OpenAI / Qwen 3.6)**. Add-in cung cấp giải pháp toàn diện giúp tăng tốc độ làm việc với các file bảng tính lớn, nhiều Sheet và nhiều Workbook.

---

## ✨ Tính Năng Nổi Bật

### 1. 📁 Điều Hướng & Quản Lý Sheet Nâng Cao (Sheet Navigator)
* **Giao diện 2 vùng thông minh (Split View):**
  * **Vùng trên:** Danh sách toàn bộ các file Excel (Workbooks) đang mở kèm số lượng sheet và badge trạng thái đang kích hoạt.
  * **Vùng dưới:** Danh sách toàn bộ các Sheets của file đang chọn.
* **Tìm kiếm thời gian thực (Real-time Filter):** Lọc nhanh Workbook và Sheet theo từ khóa.
* **Sắp xếp linh hoạt (Sort):** Sắp xếp danh sách Workbook và Sheet theo thứ tự A-Z hoặc Z-A.
* **Quản lý trạng thái Ẩn / Hiện:**
  * Hỗ trợ 3 cấp độ: `Hiển thị (Visible)`, `Bị ẩn (Hidden)`, và `Ẩn sâu (Very Hidden)`.
  * Tính năng **"Hiện tất cả Sheet ẩn"** chỉ với 1 click chuột phải.
* **Đổi màu Tab Sheet trực quan:** Chọn bảng màu hoặc mở hộp thoại Color Picker để đánh dấu phân loại dữ liệu.
* **Thao tác nhanh:** Đóng file Excel trực tiếp từ Task Pane.

---

### 2. 🤖 Trợ Lý AI & Dịch Thuật Ô Nhật ⇋ Việt (AI Copilot)
* **Tương thích toàn diện:** Kết nối chuẩn OpenAI API (`api.openai.com`) và các mô hình AI mã nguồn mở triển khai nội bộ (Local LLM: `Qwen 3.6`, `DeepSeek`, `Llama`, `vLLM`, `Ollama`).
* 🇯🇵 **Dịch thuật ô thông minh (Nhật ⇋ Việt):**
  * Quét chọn vùng ô bất kỳ trên Excel $\rightarrow$ AI tự động dịch thuật hàng loạt bảo toàn ngữ cảnh bảng tính.
  * Tùy chọn **Ghi đè ô gốc** hoặc **Ghi sang cột phụ bên phải**.
* 💡 **Tạo công thức Excel bằng tiếng Việt (Text-to-Formula):**
  * Nhập yêu cầu bằng ngôn ngữ tự nhiên (Ví dụ: *"Tính tổng cột C nếu cột A là HN và cột B > 100"*).
  * AI trích xuất công thức chuẩn (`SUMIFS`, `XLOOKUP`, `INDEX/MATCH`, `REGEX...`), giải thích từng tham số và cung cấp nút **⚡ Chèn thẳng vào ô Excel đang chọn**.
* 🔍 **Gỡ lỗi ô tính (Cell Inspector & Debugger):**
  * Đọc thông tin ô tính hiện tại (Địa chỉ, Công thức, Giá trị, Mã lỗi `#N/A`, `#VALUE!`, `#REF!`, `#NAME?`).
  * AI tự động phân tích nguyên nhân lỗi và đề xuất công thức sửa đổi chính xác.
* 💬 **Hỏi đáp AI tự do:** Trợ lý giải đáp mọi thắc mắc về nghiệp vụ Excel, phân tích số liệu, VBA macro.

---

### 3. 📋 Tạo Bảng Mục Lục Tự Động (Auto-Generate Table of Contents)
* Tự động quét toàn bộ các sheet trong Workbook hiện tại và tạo một Sheet `Mục Lục` ở vị trí đầu tiên.
* **Gắn liên kết Hyperlink:** Bấm chuột vào tên sheet trên bảng mục lục là chuyển ngay đến sheet tương ứng.
* Bảng thống kê chi tiết: STT, Tên Sheet, Trạng thái (Hiện / Ẩn / Ẩn sâu), Màu sắc Tab và Cột Ghi chú.

---

### 4. ✂️ Đổi Tên, Tách Sheet & Gộp Sheet Nâng Cao
* ✏️ **Đổi tên Sheet (Single & Batch Rename):**
  * Đổi tên nhanh trực tiếp qua Dialog (tự động kiểm tra độ dài $\le 31$ ký tự và loại bỏ ký tự cấm `\ / ? * [ ] :`).
  * Đổi tên hàng loạt: Hỗ trợ thêm Tiền tố (Prefix), Hậu tố (Suffix), Tìm và Thay thế chuỗi ký tự trong tên toàn bộ sheet.
* 📤 **Tách Sheet (Split Sheets to .xlsx):**
  * Tách từng Sheet hoặc toàn bộ Sheet trong Workbook thành các file Excel riêng biệt (`.xlsx`).
  * Tùy chọn thư mục lưu trữ với hộp thoại chọn Folder trực quan.
* 📥 **Gộp Sheet (Merge Sheets & Import Files):**
  * **Gộp Dữ Liệu:** Tự động nối tiếp toàn bộ vùng dữ liệu từ nhiều sheet được chọn vào một sheet tổng hợp `Tong_Hop`, hỗ trợ bỏ qua dòng tiêu đề từ sheet thứ 2 trở đi.
  * **Nhập Sheet từ File Ngoài:** Chọn hàng loạt file `.xlsx` / `.xls` / `.csv` từ máy tính để gom toàn bộ sheet vào Workbook hiện tại.

---

### 5. 🌓 Chế Độ Dark / Light Theme
* Tích hợp nút toggle **`🌙 / ☀️`** ngay trên thanh điều hướng.
* **Dark Mode:** Tông màu Slate tối hiện đại (`#0F172A`), chống mỏi mắt khi làm việc buổi tối hoặc phân tích dữ liệu lâu.
* **Tự động lưu trạng thái (Persistence):** Tự ghi nhớ lựa chọn giao diện vào cấu hình người dùng.

---

## 🏗️ Cấu Trúc Mã Nguồn

```
ExcelSupport/
├── ExcelSupport.csproj               # Project SDK .NET Framework 4.8 (WPF + WinForms + Excel-DNA)
├── ExcelSupport-AddIn.dna            # Manifest khai báo Add-in cho Excel-DNA
├── AddInEvents.cs                    # Lớp nạp chính (IExcelAddIn), hook sự kiện Excel COM & quản lý bộ nhớ
├── .gitignore                        # File loại trừ git chuẩn cho Visual Studio / .NET / Excel-DNA
│
├── Services/                         # Tầng dịch vụ & Cấu hình AI
│   ├── AiConfig.cs                   # Model lưu thông tin cấu hình AI & Theme
│   ├── AiConfigManager.cs            # Quản lý đọc/ghi JSON cấu hình tại %APPDATA%\ExcelSupport
│   └── OpenAiClientService.cs        # HTTP Client tương thích OpenAI API, Auto-fallback max_tokens/temperature
│
├── ViewModels/                       # Tầng ViewModel (Mô hình MVVM)
│   ├── ViewModelBase.cs              # Base ViewModel hỗ trợ INotifyPropertyChanged & IsDarkTheme
│   ├── RelayCommand.cs               # Cung cấp ICommand cho WPF Binding
│   ├── WorksheetNodeViewModel.cs     # Đại diện cho một Worksheet Node
│   ├── WorkbookNodeViewModel.cs      # Đại diện cho một Workbook Node
│   ├── TaskPaneViewModel.cs          # ViewModel chính điều khiển Task Pane & Theme toggle
│   ├── AiAssistantViewModel.cs       # Quản lý logic Dịch thuật, Sinh công thức, Gỡ lỗi
│   └── AiSettingsViewModel.cs        # Quản lý kiểm tra kết nối & Lưu cấu hình AI
│
├── Views/                            # Giao diện người dùng WPF (XAML)
│   ├── WorkbookTreeViewControl.xaml     # Điều hướng Workbooks/Worksheets & Navigation Bar
│   ├── AiAssistantControl.xaml          # Giao diện AI Trợ lý (3 Sub-tabs: Dịch, Công thức, Gỡ lỗi)
│   └── AiSettingsControl.xaml           # Giao diện Cài đặt kết nối AI
│
├── Host/                             # Cầu nối tích hợp Excel Task Pane
│   ├── TaskPaneHostControl.cs        # WinForms ElementHost nhúng WPF View
│   └── TaskPaneRegistry.cs           # Quản lý vòng đời CustomTaskPane của Excel-DNA
│
└── Ribbon/                           # Tùy biến thanh công cụ Excel Ribbon
    └── RibbonController.cs           # Tab NAVIGATOR trên Ribbon với các nút công cụ nhanh
```

---

## 🚀 Hướng Dẫn Cài Đặt & Chạy Thử

### Yêu cầu hệ thống:
* **Hệ điều hành:** Windows 10 / Windows 11 (64-bit hoặc 32-bit).
* **Microsoft Excel:** Office 2016, Office 2019, Office 2021, Microsoft 365.
* **Môi trường build:** .NET SDK (hỗ trợ target `net48`).

### 1. Build dự án từ mã nguồn:
Mở terminal tại thư mục gốc của dự án:
```bash
dotnet build
```

Sau khi build thành công, file Add-in đóng gói độc lập (standalone `.xll`) sẽ nằm tại:
* Bản **Excel 64-bit:** `bin\Debug\net48\publish\ExcelSupport-AddIn64-packed.xll`
* Bản **Excel 32-bit:** `bin\Debug\net48\publish\ExcelSupport-AddIn-packed.xll`

### 2. Kích hoạt trong Microsoft Excel:
* **Cách 1 (Mở nhanh):** Nhấp đúp chuột vào file `.xll` hoặc kéo thả trực tiếp file `.xll` vào cửa sổ Excel đang mở.
* **Cách 2 (Cài đặt vĩnh viễn):**
  1. Trong Excel, vào **File** $\rightarrow$ **Options** $\rightarrow$ **Add-ins**.
  2. Tại mục **Manage**, chọn **Excel Add-ins** rồi bấm **Go...**.
  3. Bấm **Browse...** và chọn đường dẫn tới file `ExcelSupport-AddIn64-packed.xll`.

---

## ⚙️ Hướng Dẫn Cấu Hình AI

Vào tab **`⚙️ Cài Đặt`** trên Task Pane của Add-in:
1. **API Base URL:**
   * Dùng với **OpenAI chính thức:** `https://api.openai.com/v1`
   * Dùng với **Server AI nội bộ / Ollama / vLLM:** `http://192.168.1.100:8000/v1` hoặc `http://localhost:11434/v1`
2. **Model Name:**
   * Server nội bộ: `qwen-3.6`, `qwen2.5-coder`, `deepseek-r1`
   * OpenAI: `gpt-4o`, `gpt-4o-mini`, `o3-mini`
3. **API Key:** Nhập API Key (để trống nếu server nội bộ không yêu cầu xác thực).
4. Bấm **🔌 Kiểm Tra Kết Nối Ngay** $\rightarrow$ Khi hiện **✅ KẾT NỐI THÀNH CÔNG**, bấm **💾 Lưu Cấu Hình**.

---

## 🛡️ Tối Ưu Hiệu Năng & Quản Lý Bộ Nhớ
* Toàn bộ các đối tượng Excel COM Interop (`Workbook`, `Worksheet`, `Range`, `Hyperlink`) đều được kiểm soát và giải phóng tường minh thông qua khối lệnh `try...finally` kết hợp `Marshal.ReleaseComObject`.
* Đảm bảo không xảy ra hiện tượng giữ khóa file hoặc để lại tiến trình `excel.exe` chạy ngầm sau khi đóng bảng tính.
* Tự động bật/tắt `ScreenUpdating` trong quá trình ghi dữ liệu dịch hoặc tạo mục lục để tối đa hóa tốc độ xử lý và chống giật màn hình.

---

## 📄 Bản Quyền (License)
Dự án được phân phối dưới giấy phép [MIT License](LICENSE).
