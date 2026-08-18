# 📊 ExcelSupport — AI Sheet Navigator & Copilot for Microsoft Excel

[![.NET](https://img.shields.io/badge/.NET%20Framework-4.8-512BD4?style=flat&logo=dotnet)](https://dotnet.microsoft.com/)
[![Excel-DNA](https://img.shields.io/badge/Excel--DNA-1.8.0-107C41?style=flat&logo=microsoftexcel)](https://excel-dna.net/)
[![WPF](https://img.shields.io/badge/UI-WPF%20MVVM-0078D4?style=flat)](https://learn.microsoft.com/en-us/dotnet/desktop/wpf/)
[![AI-Ready](https://img.shields.io/badge/AI-OpenAI%20%7C%20Qwen%203.6-FF6F00?style=flat)](https://github.com/QwenLM)
[![License](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

**ExcelSupport** là một Add-in chuyên nghiệp dành cho Microsoft Excel, được phát triển trên nền tảng **Excel-DNA**, **WPF MVVM** và tích hợp **AI Copilot (OpenAI / Qwen 3.6 / DeepSeek)**. Add-in cung cấp bộ giải pháp toàn diện giúp tăng tốc độ làm việc với các file bảng tính lớn, nhiều Sheet và nhiều Workbook trong môi trường doanh nghiệp.

---

## ✨ Tính Năng Nổi Bật

### 1. 📁 Điều Hướng & Quản Lý Sheet Nâng Cao (Sheet Navigator)
* **Giao diện Task Pane mặc định bên trái (Dock Left):** Tiện lợi khi thao tác song song với bảng tính Excel.
* **Cấu trúc 2 vùng thông minh (Split View):**
  * **Vùng trên:** Danh sách toàn bộ các file Excel (Workbooks) đang mở kèm số lượng sheet và badge trạng thái file đang kích hoạt.
  * **Vùng dưới:** Danh sách toàn bộ các Sheets của file đang chọn.
* **Tìm kiếm thời gian thực (Real-time Filter):** Lọc nhanh Workbook và Sheet theo từ khóa.
* **Sắp xếp tự nhiên (Natural String Sorting):** Sắp xếp danh sách Workbook và Sheet theo chuẩn Windows (`StrCmpLogicalW`), phân biệt chính xác thứ tự số (`Sheet1`, `Sheet2`, `Sheet10`).
* **Quản lý trạng thái Ẩn / Hiện:**
  * Hỗ trợ 3 cấp độ: `Hiển thị (Visible)`, `Bị ẩn (Hidden)`, và `Ẩn sâu (Very Hidden)`.
  * Tính năng **"Hiện tất cả Sheet ẩn"** chỉ với 1 click chuột phải.
* **Đổi màu Tab Sheet trực quan:** Chọn bảng màu nhanh hoặc mở hộp thoại Color Picker để đánh dấu phân loại dữ liệu.
* **Thao tác nhanh:** Đóng file Excel trực tiếp từ Task Pane.

---

### 2. 🤖 Trợ Lý AI & Dịch Thuật Ô Nhật ⇋ Việt (AI Copilot)
* **Tương thích toàn diện:** Kết nối chuẩn OpenAI API (`api.openai.com`) và các mô hình AI mã nguồn mở triển khai nội bộ (Local LLM: `Qwen 3.6`, `DeepSeek-R1`, `Llama`, `vLLM`, `Ollama`).
* 🇯🇵 **Dịch thuật ô thông minh (Nhật ⇋ Việt):**
  * Quét chọn vùng ô bất kỳ trên Excel $\rightarrow$ AI tự động dịch thuật hàng loạt bảo toàn ngữ cảnh bảng tính.
  * Tùy chọn **Ghi đè ô gốc** hoặc **Ghi sang cột phụ bên phải**.
* 📖 **Từ điển Thuật ngữ Chuyên ngành (Glossary Management):**
  * Giao diện quản lý từ điển trực quan (DataGrid) hỗ trợ Thêm, Sửa trực tiếp, Xóa từng dòng hoặc Xóa toàn bộ.
  * 🔍 Ô tìm kiếm lọc nhanh thuật ngữ theo thời gian thực.
  * 📥 **Import File:** Hỗ trợ nhập từ file **`.json`** hoặc **`.csv`** (hỏi tùy chọn Gộp thêm hoặc Thay thế toàn bộ).
  * 📤 **Export File:** Hỗ trợ xuất từ điển ra file **`.json`** hoặc **`.csv`** (chuẩn **UTF-8 with BOM** mở trên Excel không lỗi font).
  * 🧠 **Ép quy tắc AI:** Khi dịch, hệ thống tự động đưa các cặp từ vựng trong Glossary vào Prompt để AI bắt buộc phải dịch chính xác theo chuẩn doanh nghiệp.
* 💡 **Tạo công thức Excel bằng tiếng Việt (Text-to-Formula):**
  * Nhập yêu cầu bằng ngôn ngữ tự nhiên (Ví dụ: *"Tính tổng cột C nếu cột A là HN và cột B > 100"*).
  * AI trích xuất công thức chuẩn (`SUMIFS`, `XLOOKUP`, `INDEX/MATCH`, `REGEX...`), giải thích từng tham số và cung cấp nút **⚡ Chèn thẳng vào ô Excel đang chọn**.
* 🔍 **Gỡ lỗi ô tính (Cell Inspector & Debugger):**
  * Đọc thông tin ô tính hiện tại (Địa chỉ, Công thức, Giá trị, Mã lỗi `#N/A`, `#VALUE!`, `#REF!`, `#NAME?`).
  * AI tự động phân tích nguyên nhân lỗi và đề xuất công thức sửa đổi chính xác.
* 💬 **Hỏi đáp AI tự do:** Trợ lý giải đáp mọi thắc mắc về nghiệp vụ Excel, phân tích số liệu, VBA macro.

---

### 3. 📊 So Sánh 2 Workbooks / 2 Sheets (Diff & Compare Tool)
* 🔍 **Đối chiếu dữ liệu chuyên sâu:** So sánh sự khác biệt giữa 2 phiên bản file Excel (File A cũ $\leftrightarrow$ File B mới) hoặc 2 Sheet bất kỳ.
* ⚙️ **Bộ thuật toán so sánh đa dạng (bao gồm Thuật toán LCS thông minh):**
  * **LCS theo Dòng (Row LCS - Mặc định):** Tự động phát hiện chính xác các dòng chèn mới hoặc dòng bị xóa mà **không làm lệch các dòng bên dưới**, tìm ra đúng các ô bị sửa đổi trên từng dòng tương ứng.
  * **LCS theo Cột (Column LCS):** Tự động nhận diện các cột bị chèn mới hoặc xóa bỏ.
  * **LCS 2 Chiều (2D LCS Grid):** Căn chỉnh cả dòng và cột 2D trước khi so sánh ma trận dữ liệu.
  * **Theo Tọa Độ Ô Cố Định (Cell-by-Cell Grid):** So sánh giá trị từng ô cùng tọa độ `A1`, `B2`, `C10`...
  * **Theo Cột Khóa Chính (Key Column ID):** Ghép dòng dựa trên mã định danh (Mã NV, Mã SP, Số Hóa đơn...).
* 🎯 **Điều hướng tức thì (Jump to Cell):** **Click đúp vào bất kỳ dòng sai khác nào** (hoặc bấm `🎯 Đi tới ô`): Excel tự động kích hoạt file $\rightarrow$ mở đúng sheet $\rightarrow$ cuộn bôi chọn ô tính đó (chạy mượt mà, không khóa ứng dụng).
* 🎨 **Tô Màu Trực Quan (Highlight Changes):** Tự động tô màu nổi bật các ô khác biệt trên Sheet (Màu Vàng: Ô thay đổi, Màu Xanh: Thêm mới, Màu Đỏ: Bị xóa).
* 📋 **Tạo Sheet Báo Cáo & Xuất CSV:** Tự động tạo Sheet Báo Cáo sai khác (`Diff_Report_...`) kèm Hyperlink bấm là chuyển đến ô tính, hoặc xuất file `.csv` chuẩn UTF-8 with BOM.
* 🌓 **Hỗ trợ toàn diện Dark / Light Theme:** Giao diện ComboBox, RadioButton, CheckBox và bảng dữ liệu hiển thị sắc nét, tương phản tối ưu.
* 🚀 **Vị trí gọi:** Nút lớn với Icon chuyên nghiệp **`📊 So Sánh Workbooks`** trên thanh Ribbon (Tab NAVIGATOR) và Menu chuột phải Workbook trên Task Pane.

---

### 4. 🇻🇳 Kiểm Tra & Định Vị Tiếng Việt Trong Workbooks (Vietnamese Text Auditor)
* 🔍 **Rà soát toàn diện:** Quét toàn bộ **Nội dung ô** (`UsedRange`), **Tên Worksheet**, và **Ghi chú (Comments)** để phát hiện mọi vị trí còn tồn tại tiếng Việt có dấu.
* 🎯 **Điều hướng tức thì (Jump to Cell):** **Click đúp vào dòng kết quả** (hoặc bấm nút `🎯 Đi tới ô`): Excel tự động kích hoạt Workbook $\rightarrow$ mở Sheet $\rightarrow$ cuộn và chọn đúng ô tính.
* 🌐 **Phạm vi quét linh hoạt:** Hỗ trợ quét theo **Sheet hiện tại**, **Workbook hiện tại** hoặc **Tất cả các Workbook đang mở**.
* 📊 **Xuất Báo Cáo:**
  * 📋 **Tạo Sheet Báo Cáo:** Tự động tạo một sheet mới trên Excel (`VN_Check_yyyyMMdd_HHmm`) có định dạng bảng và gắn **Hyperlink** bấm là nhảy tới ô tương ứng.
  * 📤 **Xuất File CSV:** Lưu danh sách ra file CSV chuẩn UTF-8 with BOM.
* 🚀 **Vị trí mở:** Nút **`🇻🇳 Kiểm Tra Tiếng Việt`** trên Ribbon (Tab NAVIGATOR) và trong **Menu chuột phải** của Workbook / Sheet trên Task Pane.

---

### 5. 🧹 Trình Dọn Dẹp & Chuẩn Hóa Dữ Liệu (Data Cleaning & Normalization Wizard)
* ✂️ **Xử lý khoảng trắng & Ký tự ẩn:** Xóa khoảng trắng thừa đầu/cuối (`Trim`), thu gọn nhiều dấu cách liên tiếp, xóa khoảng trắng không ngắt (`&nbsp;` / `\u00A0`), xóa dấu xuống dòng trong ô (`\r`, `\n`) và ký tự điều khiển (ASCII 0-31).
* 🔤 **Chuẩn hóa Chữ HOA / thường (Text Case):** `IN HOA TOÀN BỘ`, `in thường toàn bộ`, `Viết Hoa Đầu Mỗi Từ` (Proper/Title Case: "nguyễn văn an" $\rightarrow$ "Nguyễn Văn An"), `Viết hoa đầu câu`.
* 🇯🇵 🇻🇳 **Ngôn ngữ chuyên sâu:**
  * **Chuyển tên tiếng Việt sang Katakana Nhật Bản:** "Nguyễn Văn Ánh" $\rightarrow$ "グエン・ヴァン・アイン" (Hỗ trợ tùy chọn dấu chấm giữa `・` hoặc dấu cách).
  * **Xóa dấu tiếng Việt:** "Nguyễn Văn Ánh" $\rightarrow$ "Nguyen Van Anh".
  * **Chuyển đổi Nhật Bản:** Hankaku (Nửa chiều rộng `ｱｲｳｴｵ`) $\leftrightarrow$ Zenkaku (Toàn chiều rộng `アイウエオ`).
  * Xóa chữ số (chỉ giữ chữ), Xóa chữ cái (chỉ giữ số), Xóa ký tự đặc biệt.
* 🔢 **Sửa số lưu dạng Text & Ngày tháng:** Tự động chuyển chuỗi số dạng text thành Số thực để tính toán hàm `SUM`, `VLOOKUP`, và chuẩn hóa định dạng ngày tháng (`yyyy-MM-dd`, `dd/MM/yyyy`...).
* ⚠️ **Xử lý ô trống & Lỗi:** Điền ô trống bằng giá trị tùy biến hoặc sao chép giá trị từ trên xuống (`Fill Down`), thay thế mã lỗi `#N/A`, `#VALUE!`, `#REF!`.
* 👁️ **Khung Xem Thử Trực Quan (Live Preview):** Hiển thị ngay kết quả mẫu trước khi áp dụng vào Excel.
* 🚀 **Vị trí mở:** Nút lớn với Icon chuyên dụng **`🧹 Dọn Dẹp Dữ Liệu`** trên Ribbon (Nhóm Xử Lý Dữ Liệu).

---

### 6. 🔍 Tìm & Xử Lý Dữ Liệu Trùng Lặp Nâng Cao (Smart Duplicate Finder & Grouping)
* 📌 **Lựa chọn Cột Khóa Linh Hoạt:** Chọn 1 hoặc nhiều cột kết hợp để làm tiêu chí xác định trùng lặp (Composite Key), hỗ trợ nhận diện tự động dòng Header.
* ⚙️ **2 Chế độ so khớp thông minh:**
  * **Chính xác 100% (Exact Match):** Đối chiếu tuyệt đối các cột khóa.
  * **So khớp mờ (Fuzzy Match):** Nhận diện các dòng trùng do sai lệch chính tả, thừa dấu câu, khoảng trắng với thanh trượt tỷ lệ tương đồng ($70\% - 100\%$).
* 📊 **Gom Nhóm & Trực Quan Hóa (DataGrid):** Gom các dòng trùng vào từng cụm (**Nhóm 1, Nhóm 2...**) với huy hiệu màu sắc, phân biệt rõ **Dòng gốc (Master)** và **Dòng trùng (Duplicate)**. Click đúp để nhảy ngay tới dòng trên Excel.
* 🛠️ **Các hành động xử lý trực tiếp:**
  * 🎨 **Tô Màu Nhóm Trên Sheet:** Đổi màu các nhóm trùng trên bảng tính Excel với bảng màu pastel dễ nhìn.
  * 🗑️ **Xóa Dòng Trùng:** Tự động xóa sạch các dòng thừa, giữ lại dòng đầu hoặc dòng cuối.
  * 📋 **Tách Ra Sheet Mới:** Tự động copy các dòng trùng sang Sheet `Duplicates_yyyyMMdd_HHmm` có định dạng bảng chi tiết để đối soát.
* 🚀 **Vị trí mở:** Nút lớn với Icon chuyên dụng **`🔍 Tìm Trùng Lặp`** trên Ribbon (Nhóm Xử Lý Dữ Liệu).

---

### 7. 📋 Tạo Bảng Mục Lục Tự Động (Auto-Generate Table of Contents)
* Tự động quét toàn bộ các sheet trong Workbook hiện tại và tạo một Sheet `Mục Lục` ở vị trí đầu tiên.
* **Gắn liên kết Hyperlink:** Bấm chuột vào tên sheet trên bảng mục lục là chuyển ngay đến sheet tương ứng.
* Bảng thống kê chi tiết: STT, Tên Sheet, Trạng thái (Hiện / Ẩn / Ẩn sâu), Màu sắc Tab và Cột Ghi chú.

---

### 5. ✂️ Đổi Tên, Tách Sheet & Gộp Sheet Nâng Cao
* ✏️ **Đổi tên Sheet (Single & Batch Rename):**
  * Đổi tên nhanh trực tiếp qua Dialog (tự động kiểm tra độ dài $\le 31$ ký tự và loại bỏ ký tự cấm `\ / ? * [ ] :`).
  * Đổi tên hàng loạt: Hỗ trợ thêm Tiền tố (Prefix), Hậu tố (Suffix), Tìm và Thay thế chuỗi ký tự trong tên toàn bộ sheet.
* 📤 **Tách Sheet (Split Sheets to .xlsx):**
  * Tách từng Sheet hoặc toàn bộ Sheet trong Workbook thành các file Excel riêng biệt (`.xlsx`).
  * Checkbox **"Giữ nguyên các sheet của file hiện tại"**: Nếu bỏ chọn, Add-in sẽ tự động xóa các sheet đã xuất khỏi file gốc (tự động tạo sheet trắng an toàn nếu tất cả sheet bị xóa).
* 📥 **Gộp Sheet (Merge Sheets & Import Files):**
  * **Gộp Dữ Liệu:** Tự động nối tiếp toàn bộ vùng dữ liệu từ nhiều sheet được chọn vào một sheet tổng hợp `Tong_Hop` ở cuối Workbook. Hỗ trợ tùy chọn **Gộp toàn bộ dòng** hoặc **Bỏ qua dòng tiêu đề từ sheet thứ 2**.
  * **Nhập Sheet từ File Ngoài:** Chọn hàng loạt file `.xlsx` / `.xls` / `.csv` từ máy tính để gom toàn bộ sheet vào Workbook hiện tại.

---

### 6. 🌓 Chế Độ Dark / Light Theme
* Tích hợp nút toggle **`🌙 / ☀️`** ngay trên thanh điều hướng.
* **Dark Mode:** Phối màu Slate tối hiện đại (`#0F172A`), chống mỏi mắt khi làm việc ban đêm hoặc phân tích dữ liệu lớn.
* Toàn bộ các hộp thoại (Glossary, Kiểm tra tiếng Việt, Đổi tên, Tách/Gộp sheet) đều hỗ trợ hoàn hảo cả 2 chế độ sáng và tối.
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
├── Models/                           # Tầng Model dữ liệu
│   ├── GlossaryItem.cs               # Model thuật ngữ từ điển Nhật ⇋ Việt
│   └── VietnameseLocationItem.cs     # Model vị trí phát hiện tiếng Việt trong file Excel
│
├── Services/                         # Tầng dịch vụ & Cấu hình AI
│   ├── AiConfig.cs                   # Model cấu hình AI, Glossary & Theme
│   ├── AiConfigManager.cs            # Quản lý đọc/ghi JSON cấu hình tại %APPDATA%\ExcelSupport
│   ├── GlossaryService.cs            # Dịch vụ Import/Export Glossary (CSV UTF-8 BOM & JSON)
│   └── OpenAiClientService.cs        # HTTP Client tương thích OpenAI API, Auto-fallback max_tokens/temperature
│
├── ViewModels/                       # Tầng ViewModel (Mô hình MVVM)
│   ├── ViewModelBase.cs              # Base ViewModel hỗ trợ INotifyPropertyChanged & IsDarkTheme
│   ├── RelayCommand.cs               # Cung cấp ICommand cho WPF Binding
│   ├── WorksheetNodeViewModel.cs     # Đại diện cho một Worksheet Node (kèm context commands)
│   ├── WorkbookNodeViewModel.cs      # Đại diện cho một Workbook Node (kèm context commands)
│   ├── TaskPaneViewModel.cs          # ViewModel chính điều khiển Task Pane & Theme toggle
│   ├── AiAssistantViewModel.cs       # Quản lý logic Dịch thuật, Glossary, Sinh công thức, Gỡ lỗi
│   └── AiSettingsViewModel.cs        # Quản lý kiểm tra kết nối & Lưu cấu hình AI
│
├── Views/                            # Giao diện người dùng WPF (XAML)
│   ├── WorkbookTreeViewControl.xaml  # Điều hướng Workbooks/Worksheets & Navigation Bar
│   ├── AiAssistantControl.xaml       # Giao diện AI Trợ lý (3 Sub-tabs: Dịch, Công thức, Gỡ lỗi)
│   ├── AiSettingsControl.xaml        # Giao diện Cài đặt kết nối AI
│   ├── GlossaryDialog.xaml           # Hộp thoại Quản lý Thuật Ngữ (DataGrid, Import/Export)
│   ├── VietnameseCheckDialog.xaml    # Hộp thoại Kiểm Tra & Định Vị Tiếng Việt trong file
│   ├── RenameSheetDialog.xaml        # Hộp thoại Đổi tên sheet
│   └── SheetToolsDialog.xaml         # Hộp thoại Đổi tên hàng loạt, Tách & Gộp sheet
│
├── Host/                             # Cầu nối tích hợp Excel Task Pane
│   ├── TaskPaneHostControl.cs        # WinForms ElementHost nhúng WPF View
│   └── TaskPaneRegistry.cs           # Quản lý vòng đời CustomTaskPane của Excel-DNA (Mặc định Dock Left)
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

## 🛡️ Tối Ưu Hiệu Năng & An Toàn Bộ Nhớ
* Toàn bộ các đối tượng Excel COM Interop (`Workbook`, `Worksheet`, `Range`, `Hyperlink`) đều được kiểm soát và giải phóng tường minh thông qua khối lệnh `try...finally` kết hợp `Marshal.ReleaseComObject`.
* **Không treo Excel khi xử lý hàng loạt:** Sử dụng cờ kiểm soát `_isBatchProcessing` tạm ngắt sự kiện giao diện (`EnableEvents = false`, `ScreenUpdating = false`, `DisplayAlerts = false`) trong suốt quá trình Tách sheet, Gộp sheet, hoặc Quét tiếng Việt.
* Đảm bảo không xảy ra hiện tượng giữ khóa file hoặc để lại tiến trình `excel.exe` chạy ngầm sau khi đóng bảng tính.

---

## 📄 Bản Quyền (License)
Dự án được phân phối dưới giấy phép [MIT License](LICENSE).
