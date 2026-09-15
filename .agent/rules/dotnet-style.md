---
trigger: manual
---

# C# & WPF Development Rules

- **WPF Dark & Light Theming:** Mọi Window / UserControl mới phải hỗ trợ thuộc tính `IsDarkTheme`.
  - Bắt buộc dùng `ControlTemplate` tùy biến cho `ComboBox`, `DataGridCell` để tránh việc theme Win32 mặc định đè nền trắng hoặc xanh đậm làm chữ không đọc được.
  - `DataGrid` hiển thị văn bản nhiều dòng: Không đặt `RowHeight` cố định, sử dụng `MinRowHeight="32"`, `VerticalAlignment="Top"`, `Padding="8,6"`, `TextWrapping="Wrap"`.
- **Đa Ngôn Ngữ Bắt Buộc (Mandatory Localization):**
  - **TUYỆT ĐỐI KHÔNG** hardcode bất kỳ chuỗi hiển thị nào trong giao diện XAML hoặc code-behind C#.
  - **Giao diện XAML:** Mọi TextBlock, Button, Window Title, CheckBox, Label, ToolTip... PHẢI sử dụng markup `{loc:Loc KeyName}` (với `xmlns:loc="clr-namespace:ExcelSupport.Helpers"`).
  - **Code C# / Code-behind:** Mọi thông báo MessageBox, chuỗi trạng thái, tooltip, log hiển thị cho người dùng PHẢI dùng `LocalizationService.Get("KeyName", "FallbackText")`.
  - **Đồng bộ 3 tệp ngôn ngữ:** Khi thêm bất kỳ form hoặc tính năng mới nào, **BẮT BUỘC** cập nhật đầy đủ và đồng nhất cả 3 tệp:
    1. `Languages/vi.json` (Tiếng Việt)
    2. `Languages/en.json` (English)
    3. `Languages/ja.json` (日本語)
  - Khi tạo form mới, tham khảo skill hướng dẫn tại `.agent/skills/wpf-form-localization/SKILL.md`.
- **C# Hiện Đại & Ponytail Style:**
  - Tận dụng cú pháp C# ngắn gọn (pattern matching, switch expressions, null conditional `?.`, `var`).
  - Tránh tạo interface, abstract class hoặc factory một lần nếu không có nhu cầu mở rộng thực tế (YAGNI).
  - Tái sử dụng các service/helper có sẵn trong codebase (`Host/AddInEvents.*`, `Services/`) trước khi viết mới.
- **Tránh Khóa Luồng UI:**
  - Khi mở form có quét dữ liệu Excel, sử dụng `Dispatcher.BeginInvoke(..., DispatcherPriority.Background)` để UI vẽ xong hoàn chỉnh trước khi nạp dữ liệu.
  - Xử lý tác vụ nặng qua `Task.Run` hoặc `await` với `CancellationToken` để hỗ trợ hủy thao tác mượt mà.
