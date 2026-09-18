# 📋 Kế Hoạch Triển Khai: Multi-User Connection Profiles & Table Structure Inspector

Tài liệu này ghi lại toàn bộ thiết kế kiến trúc, mô hình dữ liệu và các bước triển khai chi tiết cho 2 tính năng nâng cấp Cơ Sở Dữ Liệu Oracle.

---

## 🎯 1. Hai Tính Năng Nâng Cấp Trọng Tâm

### 1. Multi-User Connection Profiles
- **Tại "Connection Profile Details":**
  - Mở rộng mỗi Connection Profile hỗ trợ danh sách nhiều tài khoản User (`List<OracleUserCredential>`) thay vì chỉ 1 User đơn lẻ.
  - Cho phép thêm tay nhiều User với: Tên User, Mật khẩu, Vai trò / Ghi chú (`Role/Description`), Đánh dấu mặc định (`IsDefault`).
  - **Mặc định khi tạo mới một Profile:** Luôn có sẵn **2 User** (ví dụ: `SYSTEM / Admin` và `APP_USER / Read-Only`).
  - Tương thích ngược: Tự động chuyển đổi các profile cũ từ 1 user sang danh sách user.
- **Khi truy vấn dữ liệu (`Quick SQL Query`):**
  - Thêm ComboBox chọn User bên cạnh chọn Profile. Khi đổi profile thì nạp đúng danh sách user và chọn user mặc định.
  - Cho phép người dùng chuyển đổi user linh hoạt chỉ với 1 click để query với quyền tương ứng.

### 2. Quick SQL Query: Table Structure & Index Inspector
- **Nút bấm / Phím tắt `Ctrl + T`** trong Quick SQL Query để mở màn hình tra cứu cấu trúc bảng.
- Tự động nhận diện Tên Bảng từ câu lệnh `FROM [TABLE]` trong SQL Editor.
- Truy xuất thông tin từ Oracle Data Dictionary:
  - **Cột (Columns):** Tên cột, kiểu dữ liệu, nullable, Khóa chính (**PK ⭐**), giá trị mặc định, chú thích (comments).
  - **Index (Indexes):** Tên index, phân loại rõ ràng (**Index Chính / PK** vs **Index Phụ / Secondary Index**), tính duy nhất (`UNIQUE`), danh sách các cột tham gia index theo thứ tự (`Col1 ASC, Col2 DESC`), trạng thái index (`VALID`).
- Tác vụ 1-Click:
  - **`Chèn SELECT vào Editor`**: Sinh cú pháp `SELECT col1, col2... FROM table` đưa vào ô soạn thảo.
  - **`Xuất ra Sheet Excel`**: Xuất bảng đặc tả cấu trúc bảng sang một Sheet mới được kẻ viền, tô màu tiêu đề đẹp mắt.

---

## 📂 2. Danh Sách Các File Cần Chỉnh Sửa & Tạo Mới

1. **`Models/OracleCompareModels.cs`**:
   - Thêm model `OracleUserCredential`.
   - Thêm các model metadata `OracleTableColumnMetadata`, `OracleTableIndexMetadata`, `OracleTableStructureResult`.
   - Cập nhật `OracleConnectionProfile` (hỗ trợ `Users`, `SelectedUserId`, `GetEffectiveConfig`).

2. **`Services/OracleConnectionManager.cs`**:
   - Cập nhật hàm `Load()` migrate dữ liệu cũ sang cấu trúc `Users`.
   - Khởi tạo profile mặc định luôn có 2 User.

3. **`Services/OracleQuickQueryService.cs`**:
   - Thêm hàm `GetTableStructureAsync(...)` truy vấn columns, constraints và indexes từ Oracle.
   - Thêm hàm `ExportTableStructureToWorksheet(...)` xuất dữ liệu ra sheet Excel.

4. **`Views/OracleTableCompareDialog.xaml` & `.xaml.cs`**:
   - Tại tab Cài đặt Profile ("Connection Profile Details"): Thay đổi thành DataGrid danh sách Users kèm các nút Thêm, Sửa, Xóa, Đặt mặc định.

5. **`Views/OracleQuickQueryDialog.xaml` & `.xaml.cs`**:
   - Row 0: Thêm ComboBox chọn User song song với chọn Profile.
   - Row 1: Thêm nút bấm `🔍 Cấu Trúc Bảng` (phím tắt `Ctrl + T`).

6. **[NEW] `Views/OracleTableStructureDialog.xaml` & `.xaml.cs`**:
   - Hộp thoại giao diện Modern WPF Dark/Light Theme:
     - Tab 1: Danh sách Cột (Columns) kèm badge Khóa chính ⭐.
     - Tab 2: Danh sách Index (Index chính & Index phụ, Unique, Columns).
     - Các nút hành động: Chèn SELECT vào Editor, Xuất ra Sheet Excel.

7. **`Languages/LocalizationService.cs` & các file json ngôn ngữ**:
   - Bổ sung các nhãn đa ngôn ngữ (Tiếng Việt, Tiếng Anh, Tiếng Nhật).
