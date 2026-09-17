# 📋 Kế Hoạch Triển Khai: Chuyển Đổi File Excel Sang Markdown (.md) Hàng Loạt

Tài liệu này ghi lại chi tiết thiết kế kỹ thuật, kiến trúc và các bước triển khai cụ thể để bắt đầu ngay trong phiên chat mới.

---

## 🎯 1. Mục Tiêu Nghiệp Vụ
Bổ sung định dạng đích **Markdown (`.md`)** vào tính năng **Quản Trị & Chuyển Đổi File Hàng Loạt (Batch File Converter)** sẵn có của ExcelSupport:
- Chuyển đổi hàng loạt các file Excel (`.xlsx`, `.xls`, `.xlsm`, `.xlsb`) sang file Markdown (`.md`) cực nhanh.
- Sinh bảng Markdown chuẩn **GitHub Flavored Markdown (GFM)** (hỗ trợ căn lề số sang phải, escape ký tự pipe `\|`, chuyển đổi xuống dòng thành `<br>` để chống vỡ bảng).
- Xử lý linh hoạt khi file có nhiều Sheet:
  - **Chế độ 1 (Mặc định):** Gộp toàn bộ Sheet vào 1 file `.md` duy nhất (có Mục lục liên kết ở đầu và tiêu đề `## Tên Sheet` cho từng bảng).
  - **Chế độ 2:** Tách mỗi Sheet thành 1 file `.md` riêng lẻ (`TênFile_TênSheet.md`).

---

## 📂 2. Các File Cần Chỉnh Sửa

1. **`c:\SourceCode\ExcelSupport\Models\BatchFileConverterModels.cs`**:
   - Thêm `ExcelOutputFormat.Markdown`.
   - Thêm enum `MarkdownSheetExportMode` (`SingleFileWithHeadings`, `SeparateFilePerSheet`).
   - Thêm các tùy chọn Markdown vào `BatchConvertOptions`: `MarkdownMode`, `IncludeMarkdownToc`, `ConvertLineBreaksToBr`.

2. **`c:\SourceCode\ExcelSupport\Services\BatchFileConverterService.cs`**:
   - Cập nhật `GetExtensionForFormat` trả về `.md`.
   - Bổ sung phương thức `ConvertWorkbookToMarkdown(app, wb, inputPath, outputDir, options)`.
   - Sử dụng mảng 2D COM `ws.UsedRange.Value2` để duyệt dữ liệu và sinh bảng Markdown chuẩn xác, hỗ trợ UTF-8 cho tiếng Việt và tiếng Nhật.

3. **`c:\SourceCode\ExcelSupport\Views\BatchFileConverterDialog.xaml`**:
   - Thêm Panel cấu hình tùy chọn Markdown (chọn Gộp 1 file hoặc Tách từng Sheet, checkbox Mục lục).

4. **`c:\SourceCode\ExcelSupport\Views\BatchFileConverterDialog.xaml.cs`**:
   - Xử lý ẩn/hiện Panel tùy chọn Markdown khi chọn định dạng trong `CboTargetFormat`.
   - Đọc các cấu hình tùy chọn truyền vào `BatchConvertOptions`.

5. **`c:\SourceCode\ExcelSupport\Languages\LocalizationService.cs`**:
   - Thêm nhãn và tooltip đa ngôn ngữ cho cả 3 thứ tiếng (Vi, En, Ja).
