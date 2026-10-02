# Unified Project Profile & Database Settings Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Hợp nhất cấu hình Tài liệu Dự án (Project Profile) và Cấu hình Kết nối CSDL (Oracle Database Connection) thành một đối tượng quản lý duy nhất theo từng Dự án (Project Profile), kích hoạt Active đồng bộ toàn hệ thống, tích hợp vào Tab "Profile Settings" trong hộp thoại Settings (`RibbonCustomizeDialog`), và bổ sung phím tắt toàn cục `Ctrl + Shift + I` để mở nhanh Settings.

**Architecture:** 
- Mở rộng model `ProjectProfile` để chứa đối tượng `OracleConnectionProfile` (hoặc cấu hình DB tương đương) nhằm lưu giữ đầy đủ thông tin DB (Host, Port, Service/SID, Users credentials) gắn liền với từng dự án trong `project_profiles.json`.
- `ProjectProfileManager` trở thành Single Source of Truth cho cấu hình dự án & DB. `OracleConnectionManager` sẽ đồng bộ hoặc liên kết trực tiếp với danh sách profile của `ProjectProfileManager`.
- Tạo mới/Chuyển đổi UserControl `ProjectProfileSettingsControl` tích hợp cả 2 nhóm cấu hình (Tài liệu thiết kế & Database) với giao diện Master-Detail (danh sách dự án bên trái, chi tiết cấu hình bên phải chia tab/nhóm trực quan).
- Cập nhật `RibbonCustomizeDialog`: Tab 3 đổi tên thành "Profile Settings" và nhúng UserControl hợp nhất.
- Đăng ký ExcelCommand `Ctrl + Shift + I` (`^+I`) trong `AddInEvents.Commands.cs` và cập nhật Command Palette.

**Tech Stack:** C# .NET Framework 4.8, WPF (XAML), Excel-DNA, Newtonsoft.Json.

---

## Global Constraints

- Tuân thủ quy chuẩn `wpf-form-localization`: 100% chuỗi UI và thông báo phải hỗ trợ 3 ngôn ngữ (`vi.json`, `en.json`, `ja.json`).
- Hỗ trợ đầy đủ Dark Theme (`IsDarkTheme`) đồng bộ với Settings window.
- Đảm bảo an toàn tương thích ngược: Tự động migrate các profile kết nối cũ từ `oracle_connections.json` vào `project_profiles.json` nếu có.
- Tuyệt đối không làm gián đoạn các tính năng đang dùng `ProjectProfileManager` (Spec Launcher) và `OracleConnectionManager` (Table Compare, Quick Query).

---

## Proposed Changes

### 1. Data Models (`Models/Common/ProjectProfileModels.cs` & `Models/Oracle/OracleCompareModels.cs`)
- Bổ sung thuộc tính `DatabaseConnection` (`OracleConnectionProfile`) vào `ProjectProfile`.
- Cập nhật phương thức `Clone()` của `ProjectProfile` để sao chép sâu cả cấu hình `DatabaseConnection`.
- Khi khởi tạo `ProjectProfile` mới, tự động khởi tạo `DatabaseConnection` với các giá trị mặc định chuẩn (Host: localhost, Port: 1521, ServiceName: ORCL, kèm 1 user mẫu).

### 2. Synchronization & Migration (`Services/Common/ProjectProfileManager.cs` & `Services/Oracle/OracleConnectionManager.cs`)
- Trong `ProjectProfileManager.LoadConfig()`: Nếu `DatabaseConnection` trong các profile chưa có hoặc `project_profiles.json` mới tạo, kiểm tra nếu có file `oracle_connections.json` thì tự động ghép/nạp sang để người dùng không mất dữ liệu DB đã cấu hình trước đó.
- Cập nhật `OracleConnectionManager`: Cung cấp phương thức lấy danh sách profiles từ `ProjectProfileManager` hoặc đồng bộ hai chiều, để các dialog như `OracleTableCompareDialog` và `OracleQuickQueryDialog` hiển thị đúng danh sách dự án kèm DB của dự án đó.
- Khi một Profile được `SetActiveProfile(id)`: Profile đó trở thành dự án Active cho cả tính năng mở tài liệu (Spec Launcher) và công cụ Oracle (Quick Query / Table Compare chọn mặc định profile này).

### 3. Unified UI Control (`Views/Common/ProjectProfileSettingsControl.xaml`, `.xaml.cs`)
- Tạo UserControl `ProjectProfileSettingsControl` (kế thừa các điểm tối ưu của cả `ProjectProfileSettingsDialog` và `OracleConnectionSettingsControl`):
  - **Cột trái (Master)**:
    - Danh sách các dự án (ListBox) hiển thị Tên dự án, Thư mục gốc, Badge `⭐ ACTIVE`.
    - Bộ 3 nút quản lý: Thêm mới (Add), Nhân bản (Clone), Xóa (Delete).
  - **Cột phải (Detail)**:
    - Thanh trạng thái Active Banner: Hiển thị trạng thái dự án hiện tại có đang Active không, cùng nút `[⭐ Kích hoạt dự án này / Set as Active]`.
    - Tab/Nhóm 1: **Thông tin chung & Thư mục Tài liệu (Documents)**:
      - Tên dự án, Thư mục gốc (Root Folder kèm nút Browse).
      - Thư mục TKCT (Tiếng Việt & Tiếng Nhật kèm nút Chọn).
      - Thư mục TKCB (Tiếng Việt & Tiếng Nhật kèm nút Chọn).
      - Thư mục Chỉ Thị Test (Test Spec).
      - Các tùy chọn tìm kiếm (Không đệ quy thư mục con, Mở Read-Only mặc định, Định dạng file).
    - Tab/Nhóm 2: **Cơ sở dữ liệu (Database Connection)**:
      - Host, Port, Service Name / SID (Radio selector).
      - Bảng danh sách tài khoản đa người dùng (Multi-User Credentials List: Username, Password, Role, IsDefault).
      - Hàng nút hành động DB: Lưu cấu hình, Kiểm tra kết nối (Test Connection đơn lẻ), và **⚡ Test All Connection** (đã hoàn thiện ở bước trước).

### 4. Integration into Settings (`Views/Common/RibbonCustomizeDialog.xaml`, `.xaml.cs`)
- Đổi tên Tab 3 từ `{loc:Loc Settings_TabConnection}` thành `{loc:Loc Settings_TabProfile}` ("📁 Cấu Hình Dự Án (Profile)" / "📁 Profile Settings" / "📁 プロファイル設定").
- Thay thế `views:OracleConnectionSettingsControl` bằng `views:ProjectProfileSettingsControl`.
- Đảm bảo `ShowWindow(2, owner)` sẽ mở trực tiếp vào Tab 3 này.
- Cho phép `ProjectProfileSettingsDialog.ShowWindow` chuyển hướng (redirect) mở trực tiếp vào Tab 3 của `RibbonCustomizeDialog`.

### 5. Shortcut `Ctrl + Shift + I`
- Trong `Host/AddInEvents.Commands.cs`: Thêm class `SettingsCommands` với `[ExcelCommand(ShortCut = "^+I", Name = "OpenSettingsCommand")]` gọi `RibbonCustomizeDialog.ShowWindow(0, null)`.
- Trong `Services/Common/CommandPaletteService.cs`: Gán `ShortcutText = "Ctrl + Shift + I"` cho lệnh Cài Đặt.
- Cập nhật tooltip Ribbon (nếu cần).

### 6. Localization
- Cập nhật 3 file `Languages/vi.json`, `Languages/en.json`, `Languages/ja.json`:
  - Thêm / cập nhật `Settings_TabProfile`, `SpecProfile_TabGeneralDocs`, `SpecProfile_TabDatabase`, v.v.

---

## Tasks Breakdown

### Task 1: Update Models & Data Migration
**Files:**
- Modify: `Models/Common/ProjectProfileModels.cs`
- Modify: `Services/Common/ProjectProfileManager.cs`
- Modify: `Services/Oracle/OracleConnectionManager.cs`

- [ ] **Step 1.1**: Thêm thuộc tính `OracleConnectionProfile DatabaseConnection` vào class `ProjectProfile` kèm khởi tạo mặc định.
- [ ] **Step 1.2**: Cập nhật phương thức `Clone()` trong `ProjectProfile` để clone sâu cả `DatabaseConnection`.
- [ ] **Step 1.3**: Bổ sung logic di trú (migration) trong `ProjectProfileManager` để import cấu hình từ `oracle_connections.json` vào các project profile nếu chưa có.
- [ ] **Step 1.4**: Cập nhật `OracleConnectionManager` để lấy danh sách profile DB từ `ProjectProfileManager` và đồng bộ khi có thay đổi.

### Task 2: Build Unified `ProjectProfileSettingsControl`
**Files:**
- Create: `Views/Common/ProjectProfileSettingsControl.xaml`
- Create: `Views/Common/ProjectProfileSettingsControl.xaml.cs`

- [ ] **Step 2.1**: Thiết kế giao diện XAML cho UserControl gồm Master (Project List + Add/Clone/Delete) và Detail (Active Banner + SubTab Documents + SubTab Database Connection với Multi-user & Test All Connection).
- [ ] **Step 2.2**: Cài đặt Code-behind: Xử lý chọn profile, binding dữ liệu form, chọn thư mục, thêm/sửa/xóa user DB, nút Set Active, lưu cấu hình, kiểm tra kết nối đơn lẻ và Test All Connection song song.
- [ ] **Step 2.3**: Xử lý Dark Theme reactive (`IsDarkTheme`).

### Task 3: Integrate into Settings Window & Clean Up Legacy Dialogs
**Files:**
- Modify: `Views/Common/RibbonCustomizeDialog.xaml`
- Modify: `Views/Common/RibbonCustomizeDialog.xaml.cs`
- Modify: `Views/Common/ProjectProfileSettingsDialog.xaml.cs` (chuyển hướng sang RibbonCustomizeDialog)
- Modify: `Ribbon/RibbonController.cs` (nếu cần cập nhật action mở Profile Settings)

- [ ] **Step 3.1**: Đổi tên Header Tab 3 trong `RibbonCustomizeDialog.xaml` thành `Settings_TabProfile` và thay nội dung bằng `views:ProjectProfileSettingsControl`.
- [ ] **Step 3.2**: Trong `RibbonCustomizeDialog.xaml.cs`, xử lý đồng bộ theme và focus khi mở tab profile.
- [ ] **Step 3.3**: Cập nhật `ProjectProfileSettingsDialog.ShowWindow` để mở trực tiếp `RibbonCustomizeDialog.ShowWindow(2, owner)`.

### Task 4: Global Shortcut `Ctrl + Shift + I` & Command Palette
**Files:**
- Modify: `Host/AddInEvents.Commands.cs`
- Modify: `Services/Common/CommandPaletteService.cs`

- [ ] **Step 4.1**: Thêm `[ExcelCommand(ShortCut = "^+I", Name = "OpenSettingsCommand")]` gọi `RibbonCustomizeDialog.ShowWindow(0, null)`.
- [ ] **Step 4.2**: Cập nhật `CommandPaletteService` với `ShortcutText = "Ctrl + Shift + I"` cho lệnh Cài đặt.

### Task 5: Localization & Build Verification
**Files:**
- Modify: `Languages/vi.json`
- Modify: `Languages/en.json`
- Modify: `Languages/ja.json`

- [ ] **Step 5.1**: Bổ sung đầy đủ các localization keys cho tab Profile Settings, sub-tabs Documents và Database trên cả 3 thứ tiếng.
- [ ] **Step 5.2**: Thực thi `dotnet build` để kiểm tra toàn bộ solution biên dịch thành công 0 Warning, 0 Error.
- [ ] **Step 5.3**: Kiểm tra liên kết từ Oracle Quick Query / Table Compare và Spec Launchers với profile Active mới.
