# 📋 Kế Hoạch Triển Khai: Nâng Cấp Tính Năng So Sánh 2 File Excel (Workbook Compare Pro)

Tài liệu này ghi lại chi tiết thiết kế kỹ thuật, nguyên nhân lỗi và các bước triển khai cụ thể để bắt đầu ngay trong phiên chat mới.

---

## 🎯 1. Hai Vấn Đề Cần Giải Quyết

1. **Khắc phục triệt để lỗi treo con trỏ chuột khi click nhảy ô sai khác:**
   - **Nguyên nhân:** Hộp thoại `WorkbookCompareDialog` đang dùng `ShowDialog()` (Modal) và gán Owner là cửa sổ Excel (`Owner = ExcelApp.Hwnd`). Khi modal, Windows disable cửa sổ Excel. Khi DataGrid gọi `targetWb.Activate()` và `rng.Select()`, Excel cố nhận focus nhưng bị chặn bởi Modal WPF dialog, khiến trạng thái bắt chuột (Mouse Capture) của WPF DataGrid bị kẹt $\rightarrow$ Chuột bị treo quay vòng hoặc không click tiếp được.
   - **Giải pháp:** Chuyển sang **Modeless (`Show()`)**, bật `Topmost = true` (kèm nút ghim `📌`), giải phóng bắt chuột (`Mouse.Capture(null)` và `e.Handled = true`), điều hướng lệch pha qua `Dispatcher.BeginInvoke`.

2. **Nâng cấp trải nghiệm người dùng: Bật 2 file Song Song / Thẳng Đứng & Tô màu tức thời:**
   - **Xếp cửa sổ 1-Click:**
     - Nút **`Xem Song Song (Trái - Phải)`**: Gọi COM `app.Windows.Arrange(xlArrangeStyleVertical)` chia 50/50 chiều dọc.
     - Nút **`Xem Thẳng Đứng (Trên - Dưới)`**: Gọi COM `app.Windows.Arrange(xlArrangeStyleHorizontal)` chia 50/50 chiều ngang.
     - Checkbox **`Đồng bộ cuộn trang (Sync Scrolling)`**: `app.SyncScrollingSideBySide = true`.
   - **Điều hướng & Tô màu đồng thời (Dual Navigation & Instant Highlight):**
     - Click dòng sai khác hoặc bấm phím `Up`/`Down`: Nhảy đồng thời cả ô File A và ô File B.
     - Tô màu tức thời: File A tô màu đỏ/hồng nhạt (giá trị cũ), File B tô màu vàng/xanh lá (giá trị mới).
     - Tự động hoàn tác/dọn màu ô cũ khi chuyển sang điểm tiếp theo (để không làm bẩn file gốc).
   - **Phím tắt duyệt nhanh:** Phím `F7` (Điểm trước) và `F8` (Điểm tiếp theo).

---

## 📂 2. Các File Mã Nguồn Cần Chỉnh Sửa

1. **`c:\SourceCode\ExcelSupport\Host\AddInEvents.AuditorAndDiff.cs`**:
   - Thêm phương thức `ArrangeCompareWindows(string wb1Name, string wb2Name, bool isVertical, bool syncScroll)`.
   - Thêm phương thức `NavigateAndHighlightDualCells(...)` điều hướng cả 2 file và tô màu tức thời.
   - Thêm phương thức `ClearTransientHighlights()` để khôi phục màu nền cũ.

2. **`c:\SourceCode\ExcelSupport\Views\WorkbookCompareDialog.xaml.cs`**:
   - Chuyển `ShowWindow()` từ `ShowDialog()` sang `Show()` modeless, đặt `Topmost = true`.
   - Thêm nút / toggle ghim `📌 Luôn trên cùng`.
   - Xử lý `OnRowDoubleClick` và `OnGoToCellClick`: `e.Handled = true`, `Mouse.Capture(null)`.
   - Bắt sự kiện phím tắt `PreviewKeyDown` cho `F7`, `F8`, `Up`, `Down`.

3. **`c:\SourceCode\ExcelSupport\Views\WorkbookCompareDialog.xaml`**:
   - Thêm thanh công cụ Bố cục (Layout & View): Nút Song Song, Thẳng Đứng, Checkbox Đồng Bộ Cuộn, Checkbox Tô Màu Tự Động, Nút F7/F8.

4. **`c:\SourceCode\ExcelSupport\Languages\LocalizationService.cs`**:
   - Thêm localization keys cho 3 ngôn ngữ (Vi, En, Ja).
