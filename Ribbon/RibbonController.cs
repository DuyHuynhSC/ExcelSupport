using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using ExcelDna.Integration.CustomUI;
using ExcelSupport.Host;
using ExcelSupport.ViewModels;
using VmSortOrder = ExcelSupport.ViewModels.SortOrder;

namespace ExcelSupport.Ribbon
{
    [ComVisible(true)]
    [Guid("A7B5F8C1-6E3D-4A2B-9F1C-8D4E2A1B3C5D")]
    public class RibbonController : ExcelRibbon
    {
        public static RibbonController? Instance { get; private set; }

        private IRibbonUI? _ribbon;

        public override string GetCustomUI(string ribbonId)
        {
            Instance = this;
            return @"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnRibbonLoaded' loadImage='LoadImage'>
  <ribbon>
    <tabs>
      <tab id='tabWorkbookNav' label='NAVIGATOR' insertBeforeMso='TabHome'>
        
        <!-- Group 1: Bảng điều khiển chính -->
        <group id='grpNavExplorer' label='Điều Hướng'>
          <toggleButton id='btnToggleTaskPane' 
                        label='Sheet Navigator' 
                        size='large' 
                        image='navigator_icon' 
                        screentip='Bật / Tắt Sheet Navigator'
                        supertip='Mở Custom Task Pane hiển thị cây điều hướng toàn bộ Workbooks và Worksheets theo 2 vùng chuyên nghiệp.'
                        getPressed='GetTaskPanePressed' 
                        onAction='OnToggleTaskPane' />
          
          <button id='btnRefreshTree'
                  label='Làm Mới'
                  size='large'
                  imageMso='Refresh' 
                  screentip='Làm Mới Danh Sách'
                  supertip='Quét lại toàn bộ các file Excel và danh sách Sheet đang mở để cập nhật tức thì.'
                  onAction='OnRefreshTree' />
        </group>

        <!-- Group 2: Thao tác sắp xếp & tiện ích -->
        <group id='grpQuickTools' label='Thao Tác Nhanh'>
          <button id='btnCreateTOC'
                  label='Tạo Mục Lục'
                  size='normal'
                  imageMso='TableOfContentsInsert'
                  screentip='Tạo Mục Lục Sheet'
                  supertip='Tự động tạo Sheet Mục Lục chứa bảng danh sách tất cả các sheet và liên kết Hyperlink đến từng sheet.'
                  onAction='OnCreateTOC' />

          <button id='btnSplitSheets'
                  label='Tách Sheet'
                  size='normal'
                  imageMso='ExportExcelPath'
                  screentip='Tách Sheet thành file riêng'
                  supertip='Tách các sheet trong Workbook thành từng file .xlsx độc lập.'
                  onAction='OnSplitSheets' />

          <button id='btnMergeSheets'
                  label='Gộp Sheet'
                  size='normal'
                  imageMso='Consolidate'
                  screentip='Gộp Sheet'
                  supertip='Gộp dữ liệu từ nhiều sheet thành 1 sheet tổng hợp hoặc gom sheet từ nhiều file vào Workbook.'
                  onAction='OnMergeSheets' />

          <button id='btnQuickSortAZ'
                  label='Sắp xếp A-Z'
                  size='normal'
                  imageMso='SortAscendingExcel'
                  screentip='Sắp xếp danh sách A-Z'
                  supertip='Sắp xếp toàn bộ Workbook và Sheet theo thứ tự bảng chữ cái A-Z.'
                  onAction='OnSortAZ' />

          <button id='btnQuickSortZA'
                  label='Sắp xếp Z-A'
                  size='normal'
                  imageMso='SortDescendingExcel'
                  screentip='Sắp xếp danh sách Z-A'
                  supertip='Sắp xếp toàn bộ Workbook và Sheet theo thứ tự ngược Z-A.'
                  onAction='OnSortZA' />

          <button id='btnCloseCurrentWb'
                  label='Đóng File Này'
                  size='normal'
                  imageMso='FileClose'
                  screentip='Đóng Workbook Hiện Tại'
                  supertip='Đóng file Excel đang được kích hoạt.'
                  onAction='OnCloseCurrentWorkbook' />
        </group>

        <!-- Group 3: Kiểm Tra & Đối Soát -->
        <group id='grpAuditTools' label='Kiểm Tra &amp; Đối Soát'>
          <button id='btnCompareWorkbooks'
                  label='So Sánh Workbooks'
                  size='large'
                  image='compare_icon'
                  screentip='So Sánh 2 Workbooks / Sheets'
                  supertip='Đối chiếu sai khác dữ liệu giữa 2 file Excel hoặc 2 Sheet, tô màu trực quan và xuất báo cáo chi tiết.'
                  onAction='OnCompareWorkbooks' />

          <button id='btnCheckVietnamese'
                  label='Kiểm Tra Tiếng Việt'
                  size='large'
                  imageMso='Spelling'
                  screentip='Kiểm tra &amp; Định vị Tiếng Việt'
                  supertip='Quét toàn bộ ô, tên Sheet và ghi chú để tìm và nhảy tới các vị trí chứa tiếng Việt có dấu.'
                  onAction='OnCheckVietnamese' />

          <button id='btnExternalLinks'
                  label='Quản Lý Link Ngoài'
                  size='large'
                  image='link_icon'
                  screentip='Kiểm Tra &amp; Xử Lý Liên Kết Ngoài (Break Link)'
                  supertip='Quét và xử lý toàn bộ các công thức liên kết tới file ngoài không tồn tại, hỗ trợ bẻ gãy link, đóng băng giá trị, đổi file nguồn.'
                  onAction='OnExternalLinksManager' />
        </group>

        <!-- Group 4: Xử Lý Dữ Liệu -->
        <group id='grpDataTools' label='Xử Lý Dữ Liệu'>
          <button id='btnAdvancedFilter'
                  label='Bộ Lọc Nâng Cao'
                  size='large'
                  image='filter_icon'
                  screentip='Bộ Lọc Dữ Liệu Nâng Cao (Smart Advanced Filter)'
                  supertip='Lọc đa điều kiện (AND/OR), lọc danh sách paste từ clipboard, biểu thức số phức tạp, tô màu và trích xuất dữ liệu.'
                  onAction='OnAdvancedFilter' />

          <button id='btnDataCleaner'
                  label='Dọn Dẹp Dữ Liệu'
                  size='large'
                  image='cleaner_icon'
                  screentip='Trình Dọn Dẹp &amp; Chuẩn Hóa Dữ Liệu'
                  supertip='Xóa khoảng trắng thừa, chuẩn hóa chữ hoa/thường, xóa dấu tiếng Việt, chuyển sang Katakana, sửa số lưu dạng text, điền ô trống...'
                  onAction='OnDataCleaner' />

          <button id='btnDuplicateFinder'
                  label='Tìm Trùng Lặp'
                  size='large'
                  image='duplicate_icon'
                  screentip='Tìm &amp; Xử Lý Dữ Liệu Trùng Lặp Nâng Cao'
                  supertip='Tìm kiếm và gom nhóm các dòng trùng lặp theo 1 hoặc nhiều cột khóa, so khớp chính xác/mờ, tô màu và tách sheet.'
                  onAction='OnDuplicateFinder' />

          <button id='btnBatchBlankCleaner'
                  label='Xóa Dòng Trống'
                  size='large'
                  image='blank_cleaner_icon'
                  screentip='Xóa Dòng &amp; Cột Trống Hàng Loạt'
                  supertip='Quét và xóa/ẩn các dòng hoặc cột hoàn toàn trống (hoặc trống theo cột khóa) trên 1 hoặc nhiều Sheet/Workbook.'
                  onAction='OnBatchBlankCleaner' />

          <button id='btnBatchFindReplace'
                  label='Tìm &amp; Thay Thế'
                  size='large'
                  image='find_replace_icon'
                  screentip='Tìm &amp; Thay Thế Hàng Loạt Theo Bảng Tra Cứu'
                  supertip='Quét và thay thế đồng thời hàng trăm từ khóa/mã sản phẩm theo bảng đối chiếu tra cứu trên 1 hoặc nhiều Sheet/File.'
                  onAction='OnBatchFindReplace' />

          <button id='btnVisualTableMerge'
                  label='Ghép Bảng (Join)'
                  size='large'
                  image='table_merge_icon'
                  screentip='Trộn &amp; Ghép Nối Dữ Liệu Trực Quan (Visual XLOOKUP Wizard)'
                  supertip='Ghép 2 bảng dữ liệu dễ dàng theo Mã Khóa chung (Left Join, Inner Join, Full Outer Join) mà không cần viết công thức phức tạp.'
                  onAction='OnVisualTableMerge' />

          <button id='btnFuzzyDuplicate'
                  label='Trùng Lặp Ảo'
                  size='large'
                  image='fuzzy_duplicate_icon'
                  screentip='Phát Hiện Dữ Liệu Bất Thường &amp; Trùng Lặp Ảo'
                  supertip='Tìm kiếm các giá trị gần giống nhau (lỗi chính tả, khác biệt dấu tiếng Việt, khoảng trắng ẩn NBSP) và chuẩn hóa 1-Click.'
                  onAction='OnFuzzyDuplicate' />

          <button id='btnSafeMergeConsolidate'
                  label='Gộp Ô &amp; Sheet'
                  size='large'
                  image='merge_icon'
                  screentip='Gộp Ô &amp; Gộp Nhiều Sheet Bảo Toàn Dữ Liệu'
                  supertip='Gộp các ô không mất dữ liệu với dấu phân cách tùy chọn, gom dữ liệu từ nhiều Sheet thành 1 Sheet Tổng Hợp.'
                  onAction='OnSafeMergeConsolidate' />
        </group>

        <!-- Group 5: Quản Trị Tập Tin -->
        <group id='grpFileTools' label='Tập Tin Hàng Loạt'>
          <button id='btnBatchFileConverter'
                  label='Chuyển Đổi File'
                  size='large'
                  image='file_converter_icon'
                  screentip='Bộ Quản Trị &amp; Chuyển Đổi File Excel Hàng Loạt'
                  supertip='Chuyển đổi định dạng hàng loạt (.xlsx, .xls, .xlsb, .csv, .pdf), tách sheet thành file riêng hoặc gộp nhiều file vào một.'
                  onAction='OnBatchFileConverter' />
        </group>

        <!-- Group 6: Thước Ngắm & Hiển Thị -->
        <group id='grpViewTools' label='Thước Ngắm &amp; Hiển Thị'>
          <toggleButton id='btnToggleGridRuler'
                        label='Thước Ngắm Dòng/Cột'
                        size='large'
                        image='ruler_icon'
                        screentip='Thước Ngắm Giao Điểm Dòng &amp; Cột (Grid Ruler / Crosshair)'
                        supertip='Tự động tạo dải màu bán trong suốt làm nổi bật dòng và cột của ô đang chọn, giúp chống hoa mắt khi xem bảng tính lớn.'
                        onAction='OnToggleGridRuler'
                        getPressed='GetGridRulerPressed' />

          <menu id='mnuGridRulerOptions'
                label='Tùy Chỉnh Thước'
                size='large'
                image='ruler_settings_icon'
                screentip='Đổi Màu Sắc &amp; Chế Độ Thước Ngắm'
                supertip='Tùy chỉnh màu sắc nổi bật và chế độ hiển thị (Cả dòng &amp; cột, Chỉ dòng, Chỉ cột)'>
            
            <menuSeparator id='sepRulerColor' title='Màu Sắc Thước Ngắm' />
            <button id='btnColorYellow' label='Vàng Dịu' image='color_yellow_icon' onAction='OnSelectRulerColor' tag='Yellow' />
            <button id='btnColorSky' label='Xanh Biển Lơ' image='color_sky_icon' onAction='OnSelectRulerColor' tag='Sky' />
            <button id='btnColorEmerald' label='Xanh Ngọc Lục' image='color_emerald_icon' onAction='OnSelectRulerColor' tag='Emerald' />
            <button id='btnColorOrange' label='Cam Đào' image='color_orange_icon' onAction='OnSelectRulerColor' tag='Orange' />
            <button id='btnColorPurple' label='Tím Lavender' image='color_purple_icon' onAction='OnSelectRulerColor' tag='Purple' />
            <button id='btnColorPink' label='Hồng Phấn' image='color_pink_icon' onAction='OnSelectRulerColor' tag='Pink' />
            <button id='btnColorGray' label='Xám Thanh Lịch' image='color_gray_icon' onAction='OnSelectRulerColor' tag='Gray' />

            <menuSeparator id='sepRulerMode' title='Chế Độ Thước' />
            <button id='btnModeBoth' label='Cả Dòng &amp;&amp; Cột (Chữ Thập)' image='mode_both_icon' onAction='OnSelectRulerMode' tag='Both' />
            <button id='btnModeRow' label='Chỉ Dòng (Row Only)' image='mode_row_icon' onAction='OnSelectRulerMode' tag='Row' />
            <button id='btnModeCol' label='Chỉ Cột (Column Only)' image='mode_col_icon' onAction='OnSelectRulerMode' tag='Col' />

            <menuSeparator id='sepRulerHud' title='Bảng Thống Kê Nổi (HUD)' />
            <button id='btnToggleHud' label='Bảng Thống Kê Nổi (Chỉnh Cỡ Chữ Động)' image='hud_icon' onAction='OnToggleRulerHud' />
          </menu>
        </group>

        <!-- Group 7: Trợ Lý AI -->
        <group id='grpAiTools' label='Trợ Lý AI'>
          <button id='btnAiFormula'
                  label='AI Công Thức'
                  size='large'
                  image='ai_formula_icon'
                  screentip='AI Viết &amp; Sửa Lỗi Công Thức 1-Click'
                  supertip='Mở Trợ lý AI để tự động sinh công thức Excel chuẩn xác từ tiếng Việt hoặc chẩn đoán và sửa lỗi ô công thức đang chọn.'
                  onAction='OnAiFormula' />
        </group>

        <!-- Group 8: Hướng Dẫn & Trợ Giúp -->
        <group id='grpHelpTools' label='Hướng Dẫn'>
          <button id='btnUserManual'
                  label='Hướng Dẫn (Manual)'
                  size='large'
                  image='manual_icon'
                  screentip='Sách Hướng Dẫn Sử Dụng Toàn Diện'
                  supertip='Mở cẩm nang tra cứu và hướng dẫn chi tiết từ A-Z cho tất cả các tính năng của Add-in ExcelSupport.'
                  onAction='OnUserManual' />
        </group>

      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        public override object? LoadImage(string imageId)
        {
            if (imageId == "navigator_icon")
            {
                return CreateNavigatorBitmap();
            }
            if (imageId == "compare_icon")
            {
                return CreateCompareBitmap();
            }
            if (imageId == "cleaner_icon")
            {
                return CreateCleanerBitmap();
            }
            if (imageId == "duplicate_icon")
            {
                return CreateDuplicateBitmap();
            }
            if (imageId == "link_icon")
            {
                return CreateLinkBitmap();
            }
            if (imageId == "filter_icon")
            {
                return CreateFilterBitmap();
            }
            if (imageId == "blank_cleaner_icon")
            {
                return CreateBlankCleanerBitmap();
            }
            if (imageId == "find_replace_icon")
            {
                return CreateFindReplaceBitmap();
            }
            if (imageId == "ai_formula_icon")
            {
                return CreateAiFormulaBitmap();
            }
            if (imageId == "merge_icon")
            {
                return CreateMergeBitmap();
            }
            if (imageId == "table_merge_icon")
            {
                return CreateTableMergeBitmap();
            }
            if (imageId == "fuzzy_duplicate_icon")
            {
                return CreateFuzzyDuplicateBitmap();
            }
            if (imageId == "file_converter_icon")
            {
                return CreateFileConverterBitmap();
            }
            if (imageId == "manual_icon")
            {
                return CreateManualBitmap();
            }
            if (imageId == "ruler_icon")
            {
                return CreateRulerBitmap();
            }
            if (imageId == "ruler_settings_icon")
            {
                return CreateRulerSettingsBitmap();
            }
            if (imageId == "hud_icon")
            {
                return CreateHudBitmap();
            }
            if (imageId == "color_yellow_icon")
            {
                return CreateColorSwatchBitmap(Color.FromArgb(253, 224, 71), Color.FromArgb(234, 179, 8));
            }
            if (imageId == "color_sky_icon")
            {
                return CreateColorSwatchBitmap(Color.FromArgb(56, 189, 248), Color.FromArgb(2, 132, 199));
            }
            if (imageId == "color_emerald_icon")
            {
                return CreateColorSwatchBitmap(Color.FromArgb(74, 222, 128), Color.FromArgb(22, 163, 74));
            }
            if (imageId == "color_orange_icon")
            {
                return CreateColorSwatchBitmap(Color.FromArgb(251, 146, 60), Color.FromArgb(234, 88, 12));
            }
            if (imageId == "color_purple_icon")
            {
                return CreateColorSwatchBitmap(Color.FromArgb(192, 132, 252), Color.FromArgb(147, 51, 234));
            }
            if (imageId == "color_pink_icon")
            {
                return CreateColorSwatchBitmap(Color.FromArgb(244, 114, 182), Color.FromArgb(219, 39, 119));
            }
            if (imageId == "color_gray_icon")
            {
                return CreateColorSwatchBitmap(Color.FromArgb(148, 163, 184), Color.FromArgb(100, 116, 139));
            }
            if (imageId == "mode_both_icon")
            {
                return CreateModeIcon("Both");
            }
            if (imageId == "mode_row_icon")
            {
                return CreateModeIcon("Row");
            }
            if (imageId == "mode_col_icon")
            {
                return CreateModeIcon("Col");
            }
            return base.LoadImage(imageId);
        }

        private Bitmap CreateCompareBitmap()
        {
            var bmp = new Bitmap(32, 32);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.Clear(Color.Transparent);

                // 1. Sheet A (Phía sau bên trái - Màu Xanh Dương #2563EB)
                using (var brushA = new LinearGradientBrush(new Rectangle(2, 3, 17, 22),
                    Color.FromArgb(37, 99, 235), Color.FromArgb(29, 78, 216), 45f))
                {
                    FillRoundedRectangle(g, brushA, new Rectangle(2, 3, 17, 22), 3);
                }
                using (var penA = new Pen(Color.FromArgb(30, 64, 175), 1f))
                {
                    DrawRoundedRectangle(g, penA, new Rectangle(2, 3, 17, 22), 3);
                }

                // Header Sheet A
                using (var headBrush = new SolidBrush(Color.FromArgb(220, 255, 255, 255)))
                {
                    FillRoundedRectangle(g, headBrush, new Rectangle(4, 5, 13, 4), 1);
                }
                // Các dòng Sheet A
                using (var lineBrush = new SolidBrush(Color.FromArgb(160, 255, 255, 255)))
                {
                    g.FillRectangle(lineBrush, 4, 11, 13, 2);
                    g.FillRectangle(lineBrush, 4, 15, 13, 2);
                    g.FillRectangle(lineBrush, 4, 19, 13, 2);
                }

                // 2. Sheet B (Phía trước bên phải - Màu Xanh Lục Excel #107C41)
                using (var brushB = new LinearGradientBrush(new Rectangle(13, 7, 17, 22),
                    Color.FromArgb(16, 124, 65), Color.FromArgb(10, 85, 42), 45f))
                {
                    FillRoundedRectangle(g, brushB, new Rectangle(13, 7, 17, 22), 3);
                }
                using (var penB = new Pen(Color.FromArgb(6, 60, 28), 1f))
                {
                    DrawRoundedRectangle(g, penB, new Rectangle(13, 7, 17, 22), 3);
                }

                // Header Sheet B
                using (var headBrushB = new SolidBrush(Color.FromArgb(220, 255, 255, 255)))
                {
                    FillRoundedRectangle(g, headBrushB, new Rectangle(15, 9, 13, 4), 1);
                }
                // Các dòng Sheet B (với 1 ô màu Cam nổi bật thể hiện sự khác biệt!)
                using (var lineBrushB = new SolidBrush(Color.FromArgb(160, 255, 255, 255)))
                {
                    g.FillRectangle(lineBrushB, 15, 15, 13, 2);
                    // Dòng khác biệt màu cam
                    using (var diffBrush = new SolidBrush(Color.FromArgb(245, 158, 11)))
                    {
                        g.FillRectangle(diffBrush, 15, 19, 13, 2);
                    }
                    g.FillRectangle(lineBrushB, 15, 23, 13, 2);
                }

                // 3. Biểu tượng mũi tên so sánh 2 chiều ⇋ hình tròn ở góc dưới bên trái
                using (var circleBrush = new SolidBrush(Color.FromArgb(245, 158, 11))) // Amber Badge
                {
                    g.FillEllipse(circleBrush, 1, 17, 14, 14);
                }
                using (var circlePen = new Pen(Color.FromArgb(255, 255, 255), 1.2f))
                {
                    g.DrawEllipse(circlePen, 1, 17, 14, 14);
                }
                // Mũi tên ⇋ màu trắng bên trong hình tròn
                using (var arrowPen = new Pen(Color.White, 1.5f))
                {
                    // Mũi tên trên sang phải: ->
                    g.DrawLine(arrowPen, 3, 22, 11, 22);
                    g.DrawLine(arrowPen, 9, 20, 12, 22);
                    // Mũi tên dưới sang trái: <-
                    g.DrawLine(arrowPen, 4, 26, 12, 26);
                    g.DrawLine(arrowPen, 6, 28, 3, 26);
                }
            }
            return bmp;
        }

        private Bitmap CreateNavigatorBitmap()
        {
            // Tạo icon độ phân giải cao 32x32 với khử răng cưa chuẩn Office
            var bmp = new Bitmap(32, 32);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.Clear(Color.Transparent);

                // Khung nền ứng dụng Excel (Màu xanh thương hiệu Excel #107C41)
                using (var brush = new LinearGradientBrush(new Rectangle(1, 1, 30, 30),
                    Color.FromArgb(16, 124, 65), Color.FromArgb(10, 85, 42), 45f))
                {
                    FillRoundedRectangle(g, brush, new Rectangle(1, 1, 30, 30), 4);
                }

                // Viền khung
                using (var borderPen = new Pen(Color.FromArgb(6, 60, 28), 1f))
                {
                    DrawRoundedRectangle(g, borderPen, new Rectangle(1, 1, 30, 30), 4);
                }

                // Vùng 1: Header Workbook ở phía trên (Màu trắng mờ)
                using (var headerBrush = new SolidBrush(Color.FromArgb(230, 255, 255, 255)))
                {
                    FillRoundedRectangle(g, headerBrush, new Rectangle(4, 4, 24, 7), 2);
                }

                // Biểu tượng thư mục nhỏ ở Header
                using (var docBrush = new SolidBrush(Color.FromArgb(16, 124, 65)))
                {
                    g.FillRectangle(docBrush, 6, 6, 4, 3);
                    g.FillRectangle(docBrush, 12, 6, 13, 3);
                }

                // Vùng 2: Các Sheet nằm ở phía dưới (3 dòng thẻ Sheet trực quan)
                int[] yOffsets = { 13, 18, 23 };
                foreach (int y in yOffsets)
                {
                    // Nền từng Sheet
                    using (var sheetBrush = new SolidBrush(Color.FromArgb(240, 255, 255, 255)))
                    {
                        FillRoundedRectangle(g, sheetBrush, new Rectangle(4, y, 24, 4), 1);
                    }
                    // Dấu chấm đầu dòng của Sheet
                    using (var dotBrush = new SolidBrush(Color.FromArgb(16, 124, 65)))
                    {
                        g.FillRectangle(dotBrush, 6, y + 1, 2, 2);
                        g.FillRectangle(dotBrush, 9, y + 1, 16, 2);
                    }
                }
            }
            return bmp;
        }

        private static void FillRoundedRectangle(Graphics g, Brush brush, Rectangle bounds, int cornerRadius)
        {
            using (var path = GetRoundedPath(bounds, cornerRadius))
            {
                g.FillPath(brush, path);
            }
        }

        private static void DrawRoundedRectangle(Graphics g, Pen pen, Rectangle bounds, int cornerRadius)
        {
            using (var path = GetRoundedPath(bounds, cornerRadius))
            {
                g.DrawPath(pen, path);
            }
        }

        private static GraphicsPath GetRoundedPath(Rectangle bounds, int radius)
        {
            var path = new GraphicsPath();
            int diameter = radius * 2;
            var arc = new Rectangle(bounds.Location, new Size(diameter, diameter));

            path.AddArc(arc, 180, 90);
            arc.X = bounds.Right - diameter;
            path.AddArc(arc, 270, 90);
            arc.Y = bounds.Bottom - diameter;
            path.AddArc(arc, 0, 90);
            arc.X = bounds.Left;
            path.AddArc(arc, 90, 90);
            path.CloseFigure();
            return path;
        }

        public void OnRibbonLoaded(IRibbonUI ribbon)
        {
            _ribbon = ribbon;
            Instance = this;

            TaskPaneRegistry.VisibilityChanged += isVisible =>
            {
                _ribbon?.InvalidateControl("btnToggleTaskPane");
            };
        }

        public void InvalidateRibbon()
        {
            _ribbon?.InvalidateControl("btnToggleTaskPane");
        }

        public bool GetTaskPanePressed(IRibbonControl control)
        {
            return TaskPaneRegistry.IsTaskPaneVisible;
        }

        public void OnToggleTaskPane(IRibbonControl control, bool pressed)
        {
            if (AddInEvents.MainViewModel != null)
            {
                TaskPaneRegistry.ToggleTaskPane(AddInEvents.MainViewModel, pressed);
            }
        }

        public void OnRefreshTree(IRibbonControl control)
        {
            AddInEvents.Instance?.RefreshWorkbookTreePublic();
        }

        public void OnCreateTOC(IRibbonControl control)
        {
            ExcelDna.Integration.ExcelAsyncUtil.QueueAsMacro(() =>
            {
                AddInEvents.Instance?.CreateTableOfContents(null);
            });
        }

        public void OnSplitSheets(IRibbonControl control)
        {
            if (AddInEvents.MainViewModel?.SelectedWorkbook != null)
            {
                var dlg = new Views.SheetToolsDialog(AddInEvents.MainViewModel.SelectedWorkbook, 0, AddInEvents.MainViewModel.IsDarkTheme);
                dlg.ShowDialog();
            }
        }

        public void OnMergeSheets(IRibbonControl control)
        {
            if (AddInEvents.MainViewModel?.SelectedWorkbook != null)
            {
                var dlg = new Views.SheetToolsDialog(AddInEvents.MainViewModel.SelectedWorkbook, 1, AddInEvents.MainViewModel.IsDarkTheme);
                dlg.ShowDialog();
            }
        }

        public void OnSortAZ(IRibbonControl control)
        {
            if (AddInEvents.MainViewModel != null)
            {
                AddInEvents.MainViewModel.WorkbookSortOrder = VmSortOrder.Ascending;
                AddInEvents.MainViewModel.SheetSortOrder = VmSortOrder.Ascending;
            }
        }

        public void OnSortZA(IRibbonControl control)
        {
            if (AddInEvents.MainViewModel != null)
            {
                AddInEvents.MainViewModel.WorkbookSortOrder = VmSortOrder.Descending;
                AddInEvents.MainViewModel.SheetSortOrder = VmSortOrder.Descending;
            }
        }

        public void OnCloseCurrentWorkbook(IRibbonControl control)
        {
            if (AddInEvents.MainViewModel?.SelectedWorkbook != null)
            {
                AddInEvents.MainViewModel.CloseWorkbookCommand.Execute(AddInEvents.MainViewModel.SelectedWorkbook.WorkbookName);
            }
        }

        public void OnCompareWorkbooks(IRibbonControl control)
        {
            Views.WorkbookCompareDialog.ShowWindow(AddInEvents.MainViewModel?.SelectedWorkbook?.WorkbookName, AddInEvents.MainViewModel?.IsDarkTheme ?? false);
        }

        public void OnCheckVietnamese(IRibbonControl control)
        {
            Views.VietnameseCheckDialog.ShowWindow(AddInEvents.MainViewModel?.IsDarkTheme ?? false);
        }

        public void OnExternalLinksManager(IRibbonControl control)
        {
            Views.ExternalLinksManagerDialog.ShowWindow(AddInEvents.MainViewModel?.IsDarkTheme ?? false);
        }

        public void OnAdvancedFilter(IRibbonControl control)
        {
            Views.AdvancedFilterDialog.ShowWindow(AddInEvents.MainViewModel?.IsDarkTheme ?? false);
        }

        public void OnDataCleaner(IRibbonControl control)
        {
            Views.DataCleaningWizardDialog.ShowWindow(AddInEvents.MainViewModel?.IsDarkTheme ?? false);
        }

        public void OnDuplicateFinder(IRibbonControl control)
        {
            Views.DuplicateFinderDialog.ShowWindow(AddInEvents.MainViewModel?.IsDarkTheme ?? false);
        }

        public void OnBatchBlankCleaner(IRibbonControl control)
        {
            Views.BatchCleanerAndMergeDialog.ShowWindow(0, AddInEvents.MainViewModel?.IsDarkTheme ?? false);
        }

        public void OnBatchFindReplace(IRibbonControl control)
        {
            Views.BatchFindReplaceDialog.ShowWindow(AddInEvents.MainViewModel?.IsDarkTheme ?? false);
        }

        public void OnVisualTableMerge(IRibbonControl control)
        {
            Views.VisualTableMergeDialog.ShowWindow(AddInEvents.MainViewModel?.IsDarkTheme ?? false);
        }

        public void OnFuzzyDuplicate(IRibbonControl control)
        {
            Views.FuzzyDuplicateDialog.ShowWindow(AddInEvents.MainViewModel?.IsDarkTheme ?? false);
        }

        public void OnBatchFileConverter(IRibbonControl control)
        {
            Views.BatchFileConverterDialog.ShowWindow(AddInEvents.MainViewModel?.IsDarkTheme ?? false);
        }

        public void OnSafeMergeConsolidate(IRibbonControl control)
        {
            Views.BatchCleanerAndMergeDialog.ShowWindow(1, AddInEvents.MainViewModel?.IsDarkTheme ?? false);
        }

        public void OnAiFormula(IRibbonControl control)
        {
            if (AddInEvents.MainViewModel != null)
            {
                TaskPaneRegistry.ToggleTaskPane(AddInEvents.MainViewModel, true);
                AddInEvents.MainViewModel.SelectedTabIndex = 1; // Tab AI Assistant
                if (AddInEvents.MainViewModel.AiAssistant != null)
                {
                    AddInEvents.MainViewModel.AiAssistant.SelectedSubTab = 1; // Sub-Tab Sinh Công Thức
                }
            }
        }

        public void OnUserManual(IRibbonControl control)
        {
            Views.UserManualDialog.ShowWindow(AddInEvents.MainViewModel?.IsDarkTheme ?? false);
        }

        public void OnToggleGridRuler(IRibbonControl control, bool pressed)
        {
            var app = AddInEvents.Instance?.ExcelAppInstance;
            Services.GridRulerService.Toggle(app);
            _ribbon?.InvalidateControl("btnToggleGridRuler");
        }

        public bool GetGridRulerPressed(IRibbonControl control)
        {
            return Services.GridRulerService.IsEnabled;
        }

        public void OnSelectRulerColor(IRibbonControl control)
        {
            string colorKey = control.Tag ?? "Yellow";
            var app = AddInEvents.Instance?.ExcelAppInstance;
            Services.GridRulerService.SetColor(colorKey, app);
            if (!Services.GridRulerService.IsEnabled)
            {
                Services.GridRulerService.Toggle(app);
                _ribbon?.InvalidateControl("btnToggleGridRuler");
            }
        }

        public void OnSelectRulerMode(IRibbonControl control)
        {
            string modeTag = control.Tag ?? "Both";
            var mode = Services.GridRulerMode.BothRowAndCol;
            if (modeTag == "Row") mode = Services.GridRulerMode.RowOnly;
            else if (modeTag == "Col") mode = Services.GridRulerMode.ColOnly;

            var app = AddInEvents.Instance?.ExcelAppInstance;
            Services.GridRulerService.SetMode(mode, app);
            if (!Services.GridRulerService.IsEnabled)
            {
                Services.GridRulerService.Toggle(app);
                _ribbon?.InvalidateControl("btnToggleGridRuler");
            }
        }

        public void OnToggleRulerHud(IRibbonControl control)
        {
            var isDark = AddInEvents.MainViewModel?.IsDarkTheme ?? true;
            Views.RulerHudWindow.ForceOpenHud(isDark);
        }

        private Bitmap CreateCleanerBitmap()
        {
            var bmp = new Bitmap(32, 32);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.Clear(Color.Transparent);

                // Nền bảng tính xanh lục Excel
                using (var brush = new LinearGradientBrush(new Rectangle(2, 2, 28, 28),
                    Color.FromArgb(16, 124, 65), Color.FromArgb(10, 85, 42), 45f))
                {
                    FillRoundedRectangle(g, brush, new Rectangle(2, 2, 28, 28), 4);
                }

                // Viền
                using (var pen = new Pen(Color.FromArgb(6, 60, 28), 1f))
                {
                    DrawRoundedRectangle(g, pen, new Rectangle(2, 2, 28, 28), 4);
                }

                // Các ô lưới trắng mờ
                using (var gridBrush = new SolidBrush(Color.FromArgb(200, 255, 255, 255)))
                {
                    g.FillRectangle(gridBrush, 5, 5, 6, 4);
                    g.FillRectangle(gridBrush, 13, 5, 6, 4);
                    g.FillRectangle(gridBrush, 5, 11, 6, 4);
                    g.FillRectangle(gridBrush, 13, 11, 6, 4);
                }

                // Cây chổi thần / Tia sáng lấp lánh (Sparkle & Magic Wand)
                using (var wandBrush = new SolidBrush(Color.FromArgb(245, 158, 11)))
                using (var wandPen = new Pen(Color.FromArgb(254, 240, 138), 2f))
                {
                    // Thân đũa thần
                    g.DrawLine(wandPen, 10, 26, 24, 12);
                    // Đầu đũa ngôi sao / lấp lánh
                    g.FillEllipse(wandBrush, 22, 10, 6, 6);
                }

                // Tia lấp lánh nhỏ
                using (var starBrush = new SolidBrush(Color.White))
                {
                    g.FillEllipse(starBrush, 24, 6, 3, 3);
                    g.FillEllipse(starBrush, 27, 13, 3, 3);
                    g.FillEllipse(starBrush, 18, 11, 2, 2);
                }
            }
            return bmp;
        }

        private Bitmap CreateDuplicateBitmap()
        {
            var bmp = new Bitmap(32, 32);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.Clear(Color.Transparent);

                // Trang 1 phía sau (Màu tím Pastel)
                using (var card1 = new LinearGradientBrush(new Rectangle(8, 2, 20, 22),
                    Color.FromArgb(139, 92, 246), Color.FromArgb(109, 40, 217), 45f))
                {
                    FillRoundedRectangle(g, card1, new Rectangle(8, 2, 20, 22), 3);
                }

                // Trang 2 phía trước (Màu xanh dương đậm)
                using (var card2 = new LinearGradientBrush(new Rectangle(3, 8, 20, 22),
                    Color.FromArgb(37, 99, 235), Color.FromArgb(29, 78, 216), 45f))
                {
                    FillRoundedRectangle(g, card2, new Rectangle(3, 8, 20, 22), 3);
                }

                // Viền trắng cho trang trước
                using (var pen = new Pen(Color.White, 1f))
                {
                    DrawRoundedRectangle(g, pen, new Rectangle(3, 8, 20, 22), 3);
                }

                // Các dòng nội dung trên trang trước
                using (var lineBrush = new SolidBrush(Color.FromArgb(200, 255, 255, 255)))
                {
                    g.FillRectangle(lineBrush, 6, 12, 14, 2);
                    g.FillRectangle(lineBrush, 6, 16, 14, 2);
                    g.FillRectangle(lineBrush, 6, 20, 10, 2);
                }

                // Huy hiệu kính lúp & số 2 (Duplicate Badge)
                using (var badgeBrush = new SolidBrush(Color.FromArgb(245, 158, 11)))
                using (var badgePen = new Pen(Color.White, 1.5f))
                {
                    g.FillEllipse(badgeBrush, 17, 17, 13, 13);
                    g.DrawEllipse(badgePen, 17, 17, 13, 13);
                }

                // Dấu "=" màu trắng bên trong huy hiệu thể hiện trùng khớp
                using (var equalPen = new Pen(Color.White, 2f))
                {
                    g.DrawLine(equalPen, 21, 22, 26, 22);
                    g.DrawLine(equalPen, 21, 25, 26, 25);
                }
            }
            return bmp;
        }

        private Bitmap CreateLinkBitmap()
        {
            var bmp = new Bitmap(32, 32);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.Clear(Color.Transparent);

                // Nền gradient xanh dương đậm chuyên nghiệp
                using (var brush = new LinearGradientBrush(new Rectangle(2, 2, 28, 28),
                    Color.FromArgb(2, 132, 199), Color.FromArgb(3, 105, 161), 45f))
                {
                    FillRoundedRectangle(g, brush, new Rectangle(2, 2, 28, 28), 4);
                }

                // Viền ngoài
                using (var pen = new Pen(Color.FromArgb(2, 84, 130), 1f))
                {
                    DrawRoundedRectangle(g, pen, new Rectangle(2, 2, 28, 28), 4);
                }

                // Vẽ 2 mắt xích liên kết (Chain Links) xoay nghiêng 45 độ
                using (var linkPen = new Pen(Color.White, 2.5f))
                {
                    linkPen.StartCap = LineCap.Round;
                    linkPen.EndCap = LineCap.Round;

                    // Mắt xích 1: góc trên bên trái
                    g.DrawArc(linkPen, 6, 8, 10, 10, 135, 180);
                    g.DrawLine(linkPen, 9, 8, 14, 13);
                    g.DrawLine(linkPen, 6, 12, 11, 17);

                    // Mắt xích 2: góc dưới bên phải
                    g.DrawArc(linkPen, 15, 13, 10, 10, -45, 180);
                    g.DrawLine(linkPen, 20, 14, 15, 19);
                    g.DrawLine(linkPen, 24, 18, 19, 23);
                }

                // Huy hiệu tia chớp cam (Fix/Break Link Badge)
                using (var badgeBrush = new SolidBrush(Color.FromArgb(245, 158, 11)))
                using (var badgePen = new Pen(Color.White, 1.2f))
                {
                    g.FillEllipse(badgeBrush, 17, 17, 13, 13);
                    g.DrawEllipse(badgePen, 17, 17, 13, 13);
                }

                // Biểu tượng tia sét bên trong huy hiệu
                using (var boltBrush = new SolidBrush(Color.White))
                {
                    var pts = new PointF[]
                    {
                        new PointF(25f, 19f),
                        new PointF(21f, 24f),
                        new PointF(24f, 24f),
                        new PointF(22f, 28f),
                        new PointF(27f, 23f),
                        new PointF(24f, 23f)
                    };
                    g.FillPolygon(boltBrush, pts);
                }
            }
            return bmp;
        }

        private Bitmap CreateFilterBitmap()
        {
            var bmp = new Bitmap(32, 32);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.Clear(Color.Transparent);

                // Nền gradient tím - chàm hiện đại
                using (var brush = new LinearGradientBrush(new Rectangle(2, 2, 28, 28),
                    Color.FromArgb(99, 102, 241), Color.FromArgb(67, 56, 202), 45f))
                {
                    FillRoundedRectangle(g, brush, new Rectangle(2, 2, 28, 28), 4);
                }

                // Viền ngoài
                using (var pen = new Pen(Color.FromArgb(49, 46, 129), 1f))
                {
                    DrawRoundedRectangle(g, pen, new Rectangle(2, 2, 28, 28), 4);
                }

                // Vẽ Phễu lọc (Funnel) màu trắng
                using (var funnelBrush = new SolidBrush(Color.White))
                using (var funnelPen = new Pen(Color.White, 1.5f))
                {
                    funnelPen.LineJoin = LineJoin.Round;

                    var funnelPts = new PointF[]
                    {
                        new PointF(6f, 7f),
                        new PointF(26f, 7f),
                        new PointF(18f, 17f),
                        new PointF(18f, 24f),
                        new PointF(14f, 26f),
                        new PointF(14f, 17f)
                    };
                    g.FillPolygon(funnelBrush, funnelPts);
                }

                // Các vạch dòng dữ liệu bên trên phễu
                using (var linePen = new Pen(Color.FromArgb(200, 255, 255, 255), 1.5f))
                {
                    g.DrawLine(linePen, 9, 10, 23, 10);
                    g.DrawLine(linePen, 11, 13, 21, 13);
                }

                // Huy hiệu tia chớp cam / Smart Filter Badge
                using (var badgeBrush = new SolidBrush(Color.FromArgb(245, 158, 11)))
                using (var badgePen = new Pen(Color.White, 1.2f))
                {
                    g.FillEllipse(badgeBrush, 17, 17, 13, 13);
                    g.DrawEllipse(badgePen, 17, 17, 13, 13);
                }

                // Biểu tượng tia chớp
                using (var boltBrush = new SolidBrush(Color.White))
                {
                    var pts = new PointF[]
                    {
                        new PointF(25f, 19f),
                        new PointF(21f, 24f),
                        new PointF(24f, 24f),
                        new PointF(22f, 28f),
                        new PointF(27f, 23f),
                        new PointF(24f, 23f)
                    };
                    g.FillPolygon(boltBrush, pts);
                }
            }
            return bmp;
        }

        private Bitmap CreateRulerBitmap()
        {
            var bmp = new Bitmap(32, 32);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.Clear(Color.Transparent);

                // Nền gradient xanh ngọc (Teal Gradient)
                using (var brush = new LinearGradientBrush(new Rectangle(2, 2, 28, 28),
                    Color.FromArgb(13, 148, 136), Color.FromArgb(15, 118, 110), 45f))
                {
                    FillRoundedRectangle(g, brush, new Rectangle(2, 2, 28, 28), 4);
                }

                // Viền ngoài
                using (var pen = new Pen(Color.FromArgb(17, 94, 89), 1f))
                {
                    DrawRoundedRectangle(g, pen, new Rectangle(2, 2, 28, 28), 4);
                }

                // Vẽ lưới ô bảng tính mờ
                using (var gridPen = new Pen(Color.FromArgb(90, 255, 255, 255), 1f))
                {
                    g.DrawRectangle(gridPen, 5, 5, 22, 22);
                    g.DrawLine(gridPen, 16, 5, 16, 27);
                    g.DrawLine(gridPen, 5, 16, 27, 16);
                }

                // Dải thước ngang (Row Ruler) màu vàng bán trong suốt
                using (var hBrush = new SolidBrush(Color.FromArgb(200, 253, 224, 71)))
                {
                    g.FillRectangle(hBrush, 4, 13, 24, 6);
                }

                // Dải thước dọc (Column Ruler) màu vàng bán trong suốt
                using (var vBrush = new SolidBrush(Color.FromArgb(200, 253, 224, 71)))
                {
                    g.FillRectangle(vBrush, 13, 4, 6, 24);
                }

                // Ô giao điểm (Active Cell Focus) viền đỏ cam rực rỡ
                using (var focusPen = new Pen(Color.FromArgb(239, 68, 68), 2f))
                {
                    g.DrawRectangle(focusPen, 13, 13, 6, 6);
                }

                // Điểm tâm chữ thập màu trắng
                using (var centerBrush = new SolidBrush(Color.White))
                {
                    g.FillEllipse(centerBrush, 14, 14, 4, 4);
                }
            }
            return bmp;
        }

        private Bitmap CreateRulerSettingsBitmap()
        {
            var bmp = new Bitmap(32, 32);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.Clear(Color.Transparent);

                // Nền gradient tím - hồng cao cấp
                using (var brush = new LinearGradientBrush(new Rectangle(2, 2, 28, 28),
                    Color.FromArgb(168, 85, 247), Color.FromArgb(236, 72, 153), 45f))
                {
                    FillRoundedRectangle(g, brush, new Rectangle(2, 2, 28, 28), 4);
                }

                using (var pen = new Pen(Color.FromArgb(147, 51, 234), 1f))
                {
                    DrawRoundedRectangle(g, pen, new Rectangle(2, 2, 28, 28), 4);
                }

                // Bảng màu tròn (Color Palette Shape)
                using (var paletteBrush = new SolidBrush(Color.White))
                {
                    g.FillEllipse(paletteBrush, 5, 5, 22, 22);
                }

                // Các chấm màu trên bảng màu
                using (var yBrush = new SolidBrush(Color.FromArgb(234, 179, 8))) // Vàng
                using (var bBrush = new SolidBrush(Color.FromArgb(2, 132, 199))) // Xanh dương
                using (var gBrush = new SolidBrush(Color.FromArgb(22, 163, 74))) // Xanh lá
                using (var rBrush = new SolidBrush(Color.FromArgb(239, 68, 68))) // Đỏ
                {
                    g.FillEllipse(yBrush, 9, 8, 5, 5);
                    g.FillEllipse(bBrush, 17, 8, 5, 5);
                    g.FillEllipse(gBrush, 9, 16, 5, 5);
                    g.FillEllipse(rBrush, 17, 16, 5, 5);
                }
            }
            return bmp;
        }

        private Bitmap CreateColorSwatchBitmap(Color fill, Color border)
        {
            var bmp = new Bitmap(16, 16);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.Clear(Color.Transparent);

                using (var b = new SolidBrush(fill))
                {
                    g.FillEllipse(b, 1, 1, 13, 13);
                }
                using (var p = new Pen(border, 1.5f))
                {
                    g.DrawEllipse(p, 1, 1, 13, 13);
                }
            }
            return bmp;
        }

        private Bitmap CreateModeIcon(string mode)
        {
            var bmp = new Bitmap(16, 16);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.Clear(Color.Transparent);

                using (var bgBrush = new SolidBrush(Color.FromArgb(241, 245, 249)))
                using (var borderPen = new Pen(Color.FromArgb(203, 213, 225), 1f))
                {
                    g.FillRectangle(bgBrush, 0, 0, 15, 15);
                    g.DrawRectangle(borderPen, 0, 0, 15, 15);
                }

                using (var hBrush = new SolidBrush(Color.FromArgb(253, 224, 71)))
                using (var linePen = new Pen(Color.FromArgb(234, 179, 8), 1.5f))
                {
                    if (mode == "Both" || mode == "Row")
                    {
                        g.FillRectangle(hBrush, 1, 6, 14, 4);
                        g.DrawLine(linePen, 1, 8, 14, 8);
                    }
                    if (mode == "Both" || mode == "Col")
                    {
                        g.FillRectangle(hBrush, 6, 1, 4, 14);
                        g.DrawLine(linePen, 8, 1, 8, 14);
                    }
                }
            }
            return bmp;
        }

        private Bitmap CreateBlankCleanerBitmap()
        {
            var bmp = new Bitmap(32, 32);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.Clear(Color.Transparent);

                // Nền gradient đỏ cam (Crimson - Red)
                using (var brush = new LinearGradientBrush(new Rectangle(2, 2, 28, 28),
                    Color.FromArgb(239, 68, 68), Color.FromArgb(185, 28, 28), 45f))
                {
                    FillRoundedRectangle(g, brush, new Rectangle(2, 2, 28, 28), 4);
                }

                using (var pen = new Pen(Color.FromArgb(153, 27, 27), 1f))
                {
                    DrawRoundedRectangle(g, pen, new Rectangle(2, 2, 28, 28), 4);
                }

                // Dòng 1 (có dữ liệu - màu trắng)
                using (var rowBrush = new SolidBrush(Color.White))
                {
                    g.FillRectangle(rowBrush, 6, 7, 20, 3);
                }

                // Dòng 2 (dòng trống bị đứt đoạn / mờ)
                using (var dashPen = new Pen(Color.FromArgb(254, 202, 202), 1.5f))
                {
                    dashPen.DashPattern = new float[] { 2, 2 };
                    g.DrawLine(dashPen, 6, 14, 26, 14);
                }

                // Dòng 3 (có dữ liệu - màu trắng)
                using (var rowBrush = new SolidBrush(Color.White))
                {
                    g.FillRectangle(rowBrush, 6, 19, 20, 3);
                }

                // Huy hiệu dấu trừ tròn đỏ nổi bật
                using (var badgeBrush = new SolidBrush(Color.FromArgb(220, 38, 38)))
                using (var badgePen = new Pen(Color.White, 1.2f))
                {
                    g.FillEllipse(badgeBrush, 17, 17, 13, 13);
                    g.DrawEllipse(badgePen, 17, 17, 13, 13);
                }

                // Biểu tượng dấu trừ (-) màu trắng
                using (var whitePen = new Pen(Color.White, 2f))
                {
                    g.DrawLine(whitePen, 20, 23, 27, 23);
                }
            }
            return bmp;
        }

        private Bitmap CreateMergeBitmap()
        {
            var bmp = new Bitmap(32, 32);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.Clear(Color.Transparent);

                // Nền gradient tím chàm (Indigo Gradient)
                using (var brush = new LinearGradientBrush(new Rectangle(2, 2, 28, 28),
                    Color.FromArgb(99, 102, 241), Color.FromArgb(67, 56, 202), 45f))
                {
                    FillRoundedRectangle(g, brush, new Rectangle(2, 2, 28, 28), 4);
                }

                using (var pen = new Pen(Color.FromArgb(55, 48, 163), 1f))
                {
                    DrawRoundedRectangle(g, pen, new Rectangle(2, 2, 28, 28), 4);
                }

                // 2 ô ghép lại thành 1 ô lớn ở giữa
                using (var cellBrush = new SolidBrush(Color.FromArgb(224, 231, 255)))
                using (var cellPen = new Pen(Color.White, 1.2f))
                {
                    g.FillRectangle(cellBrush, 5, 8, 22, 16);
                    g.DrawRectangle(cellPen, 5, 8, 22, 16);
                }

                // Hai mũi tên hướng vào nhau (--> <--) màu tím đậm
                using (var arrowPen = new Pen(Color.FromArgb(79, 70, 229), 2f))
                {
                    // Mũi tên trái ->
                    g.DrawLine(arrowPen, 8, 16, 13, 16);
                    g.DrawLine(arrowPen, 11, 13, 14, 16);
                    g.DrawLine(arrowPen, 11, 19, 14, 16);

                    // Mũi tên phải <-
                    g.DrawLine(arrowPen, 24, 16, 19, 16);
                    g.DrawLine(arrowPen, 21, 13, 18, 16);
                    g.DrawLine(arrowPen, 21, 19, 18, 16);
                }
            }
            return bmp;
        }

        private Bitmap CreateFindReplaceBitmap()
        {
            var bmp = new Bitmap(32, 32);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.Clear(Color.Transparent);

                // Nền gradient xanh dương đậm (#0284C7 -> #0369A1)
                using (var brush = new LinearGradientBrush(new Rectangle(2, 2, 28, 28),
                    Color.FromArgb(2, 132, 199), Color.FromArgb(3, 105, 161), 45f))
                {
                    FillRoundedRectangle(g, brush, new Rectangle(2, 2, 28, 28), 4);
                }

                using (var pen = new Pen(Color.FromArgb(2, 100, 150), 1f))
                {
                    DrawRoundedRectangle(g, pen, new Rectangle(2, 2, 28, 28), 4);
                }

                // Vẽ kính lúp (Search)
                using (var glassPen = new Pen(Color.White, 2f))
                {
                    g.DrawEllipse(glassPen, 6, 6, 10, 10);
                    g.DrawLine(glassPen, 14, 14, 19, 19);
                }

                // Vẽ 2 mũi tên xoay tròn thay thế (Refresh / Replace arrows)
                using (var arrowBrush = new SolidBrush(Color.FromArgb(254, 240, 138)))
                using (var arrowPen = new Pen(Color.FromArgb(254, 240, 138), 1.8f))
                {
                    // Vòng tròn mũi tên
                    g.DrawArc(arrowPen, 16, 14, 11, 11, 30, 240);
                    // Đầu mũi tên
                    PointF[] arrowHead = new PointF[]
                    {
                        new PointF(26, 13),
                        new PointF(29, 19),
                        new PointF(23, 18)
                    };
                    g.FillPolygon(arrowBrush, arrowHead);
                }
            }
            return bmp;
        }

        private Bitmap CreateAiFormulaBitmap()
        {
            var bmp = new Bitmap(32, 32);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.Clear(Color.Transparent);

                // Nền gradient Tím vũ trụ & Xanh AI (#7C3AED -> #4F46E5)
                using (var brush = new LinearGradientBrush(new Rectangle(2, 2, 28, 28),
                    Color.FromArgb(124, 58, 237), Color.FromArgb(79, 70, 229), 45f))
                {
                    FillRoundedRectangle(g, brush, new Rectangle(2, 2, 28, 28), 4);
                }

                using (var pen = new Pen(Color.FromArgb(67, 56, 202), 1f))
                {
                    DrawRoundedRectangle(g, pen, new Rectangle(2, 2, 28, 28), 4);
                }

                // Chữ fx màu vàng sáng
                using (var font = new System.Drawing.Font("Georgia", 11, System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))
                using (var fxBrush = new SolidBrush(Color.FromArgb(254, 240, 138)))
                {
                    g.DrawString("fx", font, fxBrush, 4, 7);
                }

                // Ngôi sao lấp lánh AI (Sparkle Star)
                using (var starBrush = new SolidBrush(Color.White))
                {
                    PointF[] star = new PointF[]
                    {
                        new PointF(22, 6),
                        new PointF(23.5f, 10.5f),
                        new PointF(28, 12),
                        new PointF(23.5f, 13.5f),
                        new PointF(22, 18),
                        new PointF(20.5f, 13.5f),
                        new PointF(16, 12),
                        new PointF(20.5f, 10.5f)
                    };
                    g.FillPolygon(starBrush, star);
                }

                // Ngôi sao nhỏ phụ
                using (var starSmall = new SolidBrush(Color.FromArgb(224, 231, 255)))
                {
                    g.FillEllipse(starSmall, 24, 21, 4, 4);
                }
            }
            return bmp;
        }

        private Bitmap CreateHudBitmap()
        {
            var bmp = new Bitmap(32, 32);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.Clear(Color.Transparent);

                // Nền gradient Dark Slate (#0F172A -> #1E293B)
                using (var brush = new LinearGradientBrush(new Rectangle(2, 2, 28, 28),
                    Color.FromArgb(15, 23, 42), Color.FromArgb(30, 41, 59), 45f))
                {
                    FillRoundedRectangle(g, brush, new Rectangle(2, 2, 28, 28), 4);
                }

                using (var pen = new Pen(Color.FromArgb(56, 189, 248), 1.2f))
                {
                    DrawRoundedRectangle(g, pen, new Rectangle(2, 2, 28, 28), 4);
                }

                // Cột biểu đồ thống kê màu xanh ngọc + xanh biển + vàng
                using (var bar1 = new SolidBrush(Color.FromArgb(56, 189, 248)))
                using (var bar2 = new SolidBrush(Color.FromArgb(74, 222, 128)))
                using (var bar3 = new SolidBrush(Color.FromArgb(250, 204, 21)))
                {
                    g.FillRectangle(bar1, 6, 16, 4, 9);
                    g.FillRectangle(bar2, 13, 10, 4, 15);
                    g.FillRectangle(bar3, 20, 6, 4, 19);
                }

                // Đường chỉ số thống kê màu trắng nối các đỉnh cột
                using (var linePen = new Pen(Color.White, 1.5f))
                {
                    g.DrawLine(linePen, 8, 16, 15, 10);
                    g.DrawLine(linePen, 15, 10, 22, 6);
                }
            }
            return bmp;
        }

        private static Bitmap CreateTableMergeBitmap()
        {
            var bmp = new Bitmap(32, 32);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);

                // Bảng trái (Màu xanh biển)
                using (var b1 = new SolidBrush(Color.FromArgb(37, 99, 235)))
                {
                    g.FillRectangle(b1, 3, 5, 11, 22);
                }

                // Bảng phải (Màu xanh lá)
                using (var b2 = new SolidBrush(Color.FromArgb(16, 124, 65)))
                {
                    g.FillRectangle(b2, 18, 5, 11, 22);
                }

                // Mũi tên ghép nối ở giữa màu cam
                using (var arrowBrush = new SolidBrush(Color.FromArgb(249, 115, 22)))
                {
                    PointF[] points = {
                        new PointF(12, 12),
                        new PointF(20, 16),
                        new PointF(12, 20)
                    };
                    g.FillPolygon(arrowBrush, points);
                }
            }
            return bmp;
        }

        private static Bitmap CreateFuzzyDuplicateBitmap()
        {
            var bmp = new Bitmap(32, 32);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);

                // Nền vòng tròn tím
                using (var bg = new SolidBrush(Color.FromArgb(147, 51, 234)))
                {
                    g.FillEllipse(bg, 3, 3, 26, 26);
                }

                // Dấu sóng xấp xỉ ~ (Fuzzy) màu trắng
                using (var font = new System.Drawing.Font("Arial", 16, System.Drawing.FontStyle.Bold))
                using (var textBrush = new SolidBrush(Color.White))
                {
                    g.DrawString("≈", font, textBrush, 6, 2);
                }
            }
            return bmp;
        }

        private static Bitmap CreateFileConverterBitmap()
        {
            var bmp = new Bitmap(32, 32);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);

                // Folder màu cam/vàng
                using (var folderBrush = new SolidBrush(Color.FromArgb(245, 158, 11)))
                {
                    g.FillRectangle(folderBrush, 3, 6, 26, 20);
                }

                // Huy hiệu chuyển đổi xanh ngọc ở góc dưới
                using (var badge = new SolidBrush(Color.FromArgb(14, 165, 233)))
                {
                    g.FillEllipse(badge, 14, 14, 16, 16);
                }

                using (var p = new Pen(Color.White, 2f))
                {
                    g.DrawArc(p, 17, 17, 10, 10, 0, 270);
                }
            }
            return bmp;
        }

        private static Bitmap CreateManualBitmap()
        {
            var bmp = new Bitmap(32, 32);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);

                // Cuốn sách mở màu xanh dương đậm + xanh ngọc
                using (var bookLeft = new SolidBrush(Color.FromArgb(37, 99, 235)))
                using (var bookRight = new SolidBrush(Color.FromArgb(14, 165, 233)))
                {
                    g.FillRectangle(bookLeft, 4, 6, 11, 20);
                    g.FillRectangle(bookRight, 17, 6, 11, 20);
                }

                // Gáy sách ở giữa
                using (var spinePen = new Pen(Color.White, 2f))
                {
                    g.DrawLine(spinePen, 15, 5, 15, 27);
                }

                // Các dòng chữ mô phỏng màu trắng
                using (var linePen = new Pen(Color.White, 1.2f))
                {
                    g.DrawLine(linePen, 6, 10, 13, 10);
                    g.DrawLine(linePen, 6, 14, 13, 14);
                    g.DrawLine(linePen, 6, 18, 13, 18);

                    g.DrawLine(linePen, 19, 10, 26, 10);
                    g.DrawLine(linePen, 19, 14, 26, 14);
                    g.DrawLine(linePen, 19, 18, 26, 18);
                }
            }
            return bmp;
        }
    }
}
