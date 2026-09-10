---
trigger: manual
---

# C# & WPF Development Rules

- **WPF Dark & Light Theming:** Mọi Window / UserControl mới phải hỗ trợ thuộc tính `IsDarkTheme`.
  - Bắt buộc dùng `ControlTemplate` tùy biến cho `ComboBox`, `DataGridCell` để tránh việc theme Win32 mặc định đè nền trắng hoặc xanh đậm làm chữ không đọc được.
  - `DataGrid` hiển thị văn bản nhiều dòng: Không đặt `RowHeight` cố định, sử dụng `MinRowHeight="32"`, `VerticalAlignment="Top"`, `Padding="8,6"`, `TextWrapping="Wrap"`.
- **Đa Ngôn Ngữ (Localization):**
  - Giao diện XAML: Sử dụng `{loc:Loc KeyName}`.
  - Code C#: Sử dụng `LocalizationService.Get("KeyName")`.
  - Luôn cập nhật đồng thời cả 3 tệp ngôn ngữ: `Languages/vi.json`, `Languages/en.json`, `Languages/ja.json`.
- **C# Hiện Đại & Ponytail Style:**
  - Tận dụng cú pháp C# ngắn gọn (pattern matching, switch expressions, null conditional `?.`, `var`).
  - Tránh tạo interface, abstract class hoặc factory một lần nếu không có nhu cầu mở rộng thực tế (YAGNI).
  - Tái sử dụng các service/helper có sẵn trong codebase (`Host/AddInEvents.*`, `Services/`) trước khi viết mới.
- **Tránh Khóa Luồng UI:**
  - Khi mở form có quét dữ liệu Excel, sử dụng `Dispatcher.BeginInvoke(..., DispatcherPriority.Background)` để UI vẽ xong hoàn chỉnh trước khi nạp dữ liệu.
  - Xử lý tác vụ nặng qua `Task.Run` hoặc `await` với `CancellationToken` để hỗ trợ hủy thao tác mượt mà.
