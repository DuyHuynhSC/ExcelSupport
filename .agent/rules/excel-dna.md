---
trigger: manual
---

# ExcelDNA & Excel Interop Optimization Rules

- **Đọc/Ghi Range dạng mảng 2D (Bulk Array):** Luôn đọc/ghi dữ liệu qua `Range.Value2` dạng mảng 2D `object[,]` trong 1 lệnh COM duy nhất. Tuyệt đối không duyệt `foreach (Range cell in area.Cells)` vì sẽ sinh hàng ngàn COM Object gây đơ/treo Excel.
- **Giới hạn vùng chọn với `UsedRange`:** Khi xử lý `_excelApp.Selection`, luôn giao cắt với `ws.UsedRange` (`_excelApp.Intersect(selection, usedRange)`) để tránh duyệt 1.048.576 dòng khi người dùng chọn nguyên cột/sheet.
- **Xử lý đặc thù `SpecialCells` trên 1 ô:** Khi vùng chọn chỉ có 1 ô (`Count == 1`), không gọi `SpecialCells(xlCellTypeVisible)` vì Excel sẽ tự động bung ra toàn bộ Sheet. Kiểm tra trực tiếp qua `EntireRow.Hidden` / `EntireColumn.Hidden`.
- **Mở Dialog WPF từ Macro/Shortcut:** Không dùng `ShowDialog()` chặn luồng trong `ExcelAsyncUtil.QueueAsMacro`. Sử dụng `window.Show()` với `WindowInteropHelper.Owner = ExcelApp.Hwnd` để không khóa luồng message loop của Excel.
- **Tối ưu hiệu năng hàng loạt (Batch Operations):** Tắt `ScreenUpdating = false`, `DisplayAlerts = false`, và chuyển `Calculation = xlCalculationManual` trước khi chạy batch, luôn khôi phục trong khối `finally`.
- **Giải phóng COM Object:** Giải phóng gọn gàng `Marshal.ReleaseComObject` trong khối `finally` cho các biến `Worksheet`, `Range`, `Workbook`.
- **Tận dụng cú pháp ExcelDNA:** Ưu tiên dùng `ExcelDnaUtil.Application`, `XlCall.Excel`, `ExcelAsyncUtil` có sẵn thay vì tự viết wrapper trung gian.
