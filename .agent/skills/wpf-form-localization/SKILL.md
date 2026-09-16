---
name: wpf-form-localization
description: Quy chuẩn bắt buộc khi tạo mới hoặc chỉnh sửa bất kỳ Form / Window / UserControl / Dialog WPF trong ExcelSupport. Đảm bảo 100% phần tử UI và thông báo đều được áp dụng Localization (vi/en/ja), Dark Theme, và an toàn luồng Excel COM.
---

# Quy Chuẩn Bắt Buộc Khi Tạo Mới Form WPF (WPF Form Localization & Best Practices)

Bất kỳ khi nào tạo mới hoặc sửa đổi một form WPF (`Window`, `UserControl`, `Dialog`) trong dự án `ExcelSupport`, agent **BẮT BUỘC** phải tuân thủ nghiêm ngặt 5 bước dưới đây:

---

## 1. Khai báo Namespace Localization trong XAML

Ở thẻ gốc `<Window ...>` hoặc `<UserControl ...>`, luôn khai báo namespace:
```xml
xmlns:loc="clr-namespace:ExcelSupport.Helpers"
```

## 2. Áp dụng `{loc:Loc ...}` cho 100% chuỗi hiển thị trong XAML

**TUYỆT ĐỐI KHÔNG ĐƯỢC HARDCODE** tiếng Việt hay tiếng Anh trực tiếp vào thuộc tính của control. Tất cả đều phải dùng `{loc:Loc KeyName}`:

```xml
<!-- Tiêu đề cửa sổ -->
Title="{loc:Loc MyForm_WindowTitle}"

<!-- TextBlock, Label -->
<TextBlock Text="{loc:Loc MyForm_HeaderTitle}" FontSize="15" FontWeight="Bold"/>

<!-- Button -->
<Button Content="{loc:Loc MyForm_BtnSave}" Style="{StaticResource PrimaryActionBtn}" Click="OnSaveClick"/>

<!-- CheckBox -->
<CheckBox Content="{loc:Loc MyForm_EnableFeatureOption}"/>
```

*Quy ước đặt tên Key:*
- Đặt theo tiền tố tính năng / form: ví dụ `MyFeature_Title`, `MyFeature_BtnSave`, `MyFeature_ErrInvalidInput`.

---

## 3. Localization trong Code-Behind (C#)

Mọi thông báo hiển thị cho người dùng (`MessageBox`, `TextBlock.Text` động, tooltip, status bar) **BẮT BUỘC** sử dụng `LocalizationService`:

```csharp
using ExcelSupport.Services;
using WpfMessageBox = System.Windows.MessageBox;

// Lấy chuỗi không tham số (hoặc có fallback mặc định)
string msg = LocalizationService.Get("MyFeature_SuccessMsg", "Thao tác thành công!");

// Chuỗi có tham số định dạng ({0}, {1}...): truyền tham số trực tiếp vào LocalizationService.Get
// LƯU Ý: Tuyệt đối KHÔNG lồng string.Format(LocalizationService.Get("Key", "Fallback {0}"), count)
// vì chuỗi fallback sẽ bị hiểu nhầm là tham số args[0], gây lỗi FormatException!
string formatted = LocalizationService.Get("MyFeature_ItemCountFormat", count);

// Hiển thị MessageBox luôn kèm owner 'this' để modal đúng cửa sổ
WpfMessageBox.Show(
    this,
    LocalizationService.Get("MyFeature_ConfirmPrompt", "Bạn có chắc chắn muốn thực hiện?"),
    LocalizationService.Get("Common_Notice", "Thông Báo"),
    MessageBoxButton.YesNo,
    MessageBoxImage.Question);
```

---

## 4. Bắt Buộc Đồng Bộ Cả 3 Tệp Ngôn Ngữ

Sau khi định nghĩa các `KeyName`, agent **PHẢI** bổ sung ngay vào cả 3 tệp JSON:
1. `Languages/vi.json` (Tiếng Việt)
2. `Languages/en.json` (Tiếng Anh)
3. `Languages/ja.json` (日本語)

*Lưu ý:* Giữ đúng định dạng JSON hợp lệ, không thừa dấu phẩy ở cuối tệp.

---

## 5. Quy Tắc Ổn Định & An Toàn Luồng Excel COM

1. **Owner Window:** Khi mở dialog từ Excel hoặc từ form khác, luôn gắn owner (`WindowInteropHelper` hoặc `Owner = this`) để dialog không bị chìm xuống dưới Excel và người dùng không click nhầm vào sheet khi modal đang mở.
2. **Không Mutate Collection Trong Vòng Lặp:** Khi lưu danh sách (VD: Profile, Rule), không vừa `foreach` vừa sửa trực tiếp collection đó. Dùng phương thức atomic (như `SaveAllProfiles` với bản copy mới `new List<T>(items)`).
3. **Try-Catch UI Handlers:** Các sự kiện click quan trọng (`OnSaveClick`, `OnDeleteClick`, `OnExecuteClick`) phải luôn bọc trong `try / catch` để tránh crash cả tiến trình `EXCEL.EXE`.
