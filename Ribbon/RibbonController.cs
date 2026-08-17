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

        <!-- Group 3: Kiểm tra & Đối Soát -->
        <group id='grpAuditTools' label='Kiểm Tra &amp; Đối Soát'>
          <button id='btnCompareWorkbooks'
                  label='So Sánh Workbooks'
                  size='large'
                  imageMso='CompareMaster'
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
            return base.LoadImage(imageId);
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
    }
}
