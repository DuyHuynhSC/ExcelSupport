using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using ExcelSupport.Host;
using ExcelSupport.Models;
using ExcelSupport.Services;
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using WpfMessageBox = System.Windows.MessageBox;

namespace ExcelSupport.Views
{
    public partial class ExternalLinksManagerDialog : System.Windows.Window
    {
        public static readonly DependencyProperty IsDarkThemeProperty =
            DependencyProperty.Register(nameof(IsDarkTheme), typeof(bool), typeof(ExternalLinksManagerDialog),
                new PropertyMetadata(false));

        public bool IsDarkTheme
        {
            get => (bool)GetValue(IsDarkThemeProperty);
            set => SetValue(IsDarkThemeProperty, value);
        }

        private readonly ExcelApp? _excelApp;
        private Workbook? _targetWb;
        private ExternalLinksScanResult? _scanResult;

        private ObservableCollection<ExternalSourceItem> _sourcesList = new ObservableCollection<ExternalSourceItem>();
        private ObservableCollection<BrokenFormulaCellItem> _formulaCellsList = new ObservableCollection<BrokenFormulaCellItem>();
        private ObservableCollection<ExternalNamedRangeItem> _namedRangesList = new ObservableCollection<ExternalNamedRangeItem>();

        private static ExternalLinksManagerDialog? _currentInstance;

        internal static void ShowWindow(bool isDarkTheme = false)
        {
            try
            {
                if (_currentInstance != null && _currentInstance.IsLoaded)
                {
                    _currentInstance.IsDarkTheme = isDarkTheme;
                    _currentInstance.Activate();
                    _currentInstance.RefreshData();
                    return;
                }

                var addIn = AddInEvents.Instance;
                var app = addIn?.ExcelAppInstance;

                _currentInstance = new ExternalLinksManagerDialog(app)
                {
                    IsDarkTheme = isDarkTheme
                };

                try
                {
                    if (app != null)
                    {
                        new System.Windows.Interop.WindowInteropHelper(_currentInstance).Owner = (IntPtr)app.Hwnd;
                    }
                }
                catch { }

                _currentInstance.ShowDialog();
                _currentInstance = null;
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Lỗi mở màn hình Quản lý liên kết ngoài:\n{ex.Message}\n\nChi tiết:\n{ex.StackTrace}",
                                   "ExcelSupport", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        public ExternalLinksManagerDialog(ExcelApp? app)
        {
            InitializeComponent();
            _excelApp = app;

            try
            {
                IsDarkTheme = AddInEvents.MainViewModel?.IsDarkTheme ?? false;
            }
            catch { }

            Loaded += OnDialogLoaded;
        }

        private void OnDialogLoaded(object sender, RoutedEventArgs e)
        {
            RefreshData();
        }

        private void RefreshData()
        {
            if (_excelApp == null) return;

            try
            {
                _targetWb = _excelApp.ActiveWorkbook;
                if (_targetWb == null)
                {
                    WpfMessageBox.Show("Không tìm thấy Workbook nào đang mở để quét liên kết ngoài.", "Thông Báo",
                                       MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                _scanResult = ExternalLinksService.ScanWorkbook(_targetWb);

                // Cập nhật thẻ thống kê
                TxtStatSources.Text = _scanResult.Sources.Count.ToString();
                TxtStatBroken.Text = _scanResult.TotalBrokenLinksCount.ToString();
                TxtStatCells.Text = _scanResult.FormulaCells.Count.ToString();
                TxtStatNames.Text = _scanResult.NamedRanges.Count.ToString();

                // Đổ dữ liệu vào Collections
                _sourcesList = new ObservableCollection<ExternalSourceItem>(_scanResult.Sources);
                GridSources.ItemsSource = _sourcesList;

                _formulaCellsList = new ObservableCollection<BrokenFormulaCellItem>(_scanResult.FormulaCells);
                GridFormulaCells.ItemsSource = _formulaCellsList;

                _namedRangesList = new ObservableCollection<ExternalNamedRangeItem>(_scanResult.NamedRanges);
                GridNamedRanges.ItemsSource = _namedRangesList;
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Lỗi quét liên kết ngoài:\n{ex.Message}", "Lỗi Quét",
                                   MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void OnRefreshClick(object sender, RoutedEventArgs e)
        {
            RefreshData();
        }

        #region Tab 1: External Sources Actions

        private void OnBreakSelectedSourceClick(object sender, RoutedEventArgs e)
        {
            if (_targetWb == null) return;

            var selected = GridSources.SelectedItem as ExternalSourceItem;
            if (selected == null)
            {
                WpfMessageBox.Show("Vui lòng chọn một file nguồn ngoài danh sách để bẻ gãy liên kết.", "Thông Báo",
                                   MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var confirm = WpfMessageBox.Show(
                $"Bạn có chắc chắn muốn bẻ gãy (Break Link) liên kết tới file:\n'{selected.FileName}'?\n\nToàn bộ công thức tham chiếu tới file này sẽ được chuyển thành giá trị tĩnh.",
                "Xác Nhận Break Link",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (confirm == MessageBoxResult.Yes)
            {
                bool ok = ExternalLinksService.BreakSpecificLink(_targetWb, selected.SourcePath);
                if (ok)
                {
                    WpfMessageBox.Show($"Đã bẻ gãy thành công liên kết tới '{selected.FileName}'!", "Thành Công",
                                       MessageBoxButton.OK, MessageBoxImage.Information);
                    RefreshData();
                }
                else
                {
                    WpfMessageBox.Show($"Không thể bẻ gãy liên kết tới '{selected.FileName}'.", "Thông Báo",
                                       MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
        }

        private void OnChangeLinkClick(object sender, RoutedEventArgs e)
        {
            if (_targetWb == null) return;

            var selected = GridSources.SelectedItem as ExternalSourceItem;
            if (selected == null)
            {
                WpfMessageBox.Show("Vui lòng chọn một file nguồn để đổi đường dẫn mới.", "Thông Báo",
                                   MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var openDlg = new Microsoft.Win32.OpenFileDialog
            {
                Title = $"Chọn file Excel mới thay thế cho '{selected.FileName}'",
                Filter = "Excel Files (*.xlsx;*.xls;*.xlsm;*.xlsb)|*.xlsx;*.xls;*.xlsm;*.xlsb|All Files (*.*)|*.*"
            };

            if (openDlg.ShowDialog() == true)
            {
                string newFilePath = openDlg.FileName;
                bool ok = ExternalLinksService.ChangeLinkSource(_targetWb, selected.SourcePath, newFilePath);
                if (ok)
                {
                    WpfMessageBox.Show($"Đã chuyển hướng liên kết thành công sang file mới:\n{Path.GetFileName(newFilePath)}", "Thành Công",
                                       MessageBoxButton.OK, MessageBoxImage.Information);
                    RefreshData();
                }
                else
                {
                    WpfMessageBox.Show("Không thể chuyển hướng liên kết. Vui lòng kiểm tra cấu trúc sheet của file mới.", "Thông Báo",
                                       MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
        }

        #endregion

        #region Tab 2: Formula Cells Actions

        private void OnCellSearchChanged(object sender, TextChangedEventArgs e)
        {
            if (_scanResult == null) return;

            string query = TxtSearchCells.Text.Trim();
            if (string.IsNullOrEmpty(query))
            {
                GridFormulaCells.ItemsSource = _formulaCellsList;
            }
            else
            {
                var filtered = _formulaCellsList.Where(c =>
                    c.SheetName.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0 ||
                    c.CellAddress.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0 ||
                    c.Formula.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0 ||
                    c.ExternalSource.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0 ||
                    c.CurrentValue.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0
                ).ToList();

                GridFormulaCells.ItemsSource = filtered;
            }
        }

        private void OnSelectAllCellsClick(object sender, RoutedEventArgs e)
        {
            var list = GridFormulaCells.ItemsSource as IEnumerable<BrokenFormulaCellItem>;
            if (list != null)
            {
                foreach (var item in list) item.IsSelected = true;
            }
        }

        private void OnUnselectAllCellsClick(object sender, RoutedEventArgs e)
        {
            var list = GridFormulaCells.ItemsSource as IEnumerable<BrokenFormulaCellItem>;
            if (list != null)
            {
                foreach (var item in list) item.IsSelected = false;
            }
        }

        private void OnGoToSelectedCellClick(object sender, RoutedEventArgs e)
        {
            var selected = GridFormulaCells.SelectedItem as BrokenFormulaCellItem;
            if (selected != null)
            {
                GoToCell(selected);
            }
        }

        private void OnGridCellDoubleClick(object sender, MouseButtonEventArgs e)
        {
            var selected = GridFormulaCells.SelectedItem as BrokenFormulaCellItem;
            if (selected != null)
            {
                GoToCell(selected);
            }
        }

        private void GoToCell(BrokenFormulaCellItem item)
        {
            if (_targetWb == null || _excelApp == null) return;

            try
            {
                _Worksheet? ws = null;
                Range? cell = null;
                try
                {
                    ws = _targetWb.Worksheets[item.SheetName] as _Worksheet;
                    if (ws != null)
                    {
                        ws.Activate();
                        cell = ws.Cells[item.Row, item.Column] as Range;
                        cell?.Select();
                    }
                }
                finally
                {
                    if (cell != null) Marshal.ReleaseComObject(cell);
                    if (ws != null) Marshal.ReleaseComObject(ws);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"GoToCell error: {ex.Message}");
            }
        }

        private void OnFreezeSelectedCellsClick(object sender, RoutedEventArgs e)
        {
            if (_targetWb == null) return;

            var selected = _formulaCellsList.Where(c => c.IsSelected).ToList();
            if (selected.Count == 0)
            {
                WpfMessageBox.Show("Vui lòng tích chọn ít nhất 1 ô công thức cần chuyển thành giá trị tĩnh.", "Thông Báo",
                                   MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var confirm = WpfMessageBox.Show(
                $"Bạn có chắc chắn muốn chuyển {selected.Count} ô công thức đã chọn thành giá trị tĩnh (Freeze Values)?",
                "Xác Nhận",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (confirm == MessageBoxResult.Yes)
            {
                int count = ExternalLinksService.ConvertCellsToValues(_targetWb, selected);
                WpfMessageBox.Show($"Đã chuyển thành công {count} ô thành giá trị tĩnh!", "Thành Công",
                                   MessageBoxButton.OK, MessageBoxImage.Information);
                RefreshData();
            }
        }

        #endregion

        #region Tab 3: Named Ranges Actions

        private void OnDeleteSelectedNamesClick(object sender, RoutedEventArgs e)
        {
            if (_targetWb == null) return;

            var selected = _namedRangesList.Where(n => n.IsSelected).ToList();
            if (selected.Count == 0)
            {
                WpfMessageBox.Show("Vui lòng tích chọn ít nhất 1 tên vùng (Named Range) cần xóa.", "Thông Báo",
                                   MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var confirm = WpfMessageBox.Show(
                $"Bạn có chắc chắn muốn xóa {selected.Count} tên vùng đã chọn?",
                "Xác Nhận Xóa",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (confirm == MessageBoxResult.Yes)
            {
                int count = ExternalLinksService.DeleteNamedRanges(_targetWb, selected);
                WpfMessageBox.Show($"Đã xóa thành công {count} tên vùng!", "Thành Công",
                                   MessageBoxButton.OK, MessageBoxImage.Information);
                RefreshData();
            }
        }

        #endregion

        #region Footer Actions

        private void OnHighlightCellsClick(object sender, RoutedEventArgs e)
        {
            if (_targetWb == null || _scanResult == null || _scanResult.FormulaCells.Count == 0)
            {
                WpfMessageBox.Show("Không có ô chứa liên kết ngoài nào để tô màu.", "Thông Báo",
                                   MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            int count = ExternalLinksService.HighlightCellsOnExcel(_targetWb, _scanResult.FormulaCells);
            WpfMessageBox.Show($"Đã tô màu đánh dấu {count} ô trên bảng tính Excel!\n(Màu vàng: Link ngoài | Màu đỏ: Link bị hỏng/file thiếu)", "Đã Tô Màu",
                               MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void OnExportReportClick(object sender, RoutedEventArgs e)
        {
            if (_targetWb == null || _scanResult == null) return;

            bool ok = ExternalLinksService.ExportReportToSheet(_targetWb, _scanResult);
            if (ok)
            {
                WpfMessageBox.Show("Đã tạo sheet báo cáo 'BaoCao_LinkNgoai' thành công!", "Xuất Báo Cáo",
                                   MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                WpfMessageBox.Show("Không thể tạo sheet báo cáo.", "Thông Báo",
                                   MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void OnBreakAllLinksClick(object sender, RoutedEventArgs e)
        {
            if (_targetWb == null || _scanResult == null) return;

            var confirm = WpfMessageBox.Show(
                "CẢNH BÁO: Thao tác này sẽ BẺ GÃY TOÀN BỘ liên kết ngoài trong toàn bộ Workbook và chuyển toàn bộ các ô công thức ngoài thành giá trị tĩnh.\n\nBạn có chắc chắn muốn thực hiện không?",
                "Xác Nhận Bẻ Gãy Toàn Bộ Link (Break All)",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (confirm == MessageBoxResult.Yes)
            {
                int brokenSources = ExternalLinksService.BreakAllWorkbookLinks(_targetWb);
                int convertedCells = ExternalLinksService.ConvertCellsToValues(_targetWb, _scanResult.FormulaCells);

                WpfMessageBox.Show($"Đã bẻ gãy {brokenSources} file liên kết ngoài và chuyển {convertedCells} ô công thức thành giá trị tĩnh thành công!", "Thành Công",
                                   MessageBoxButton.OK, MessageBoxImage.Information);
                RefreshData();
            }
        }

        private void OnCloseClick(object sender, RoutedEventArgs e)
        {
            Close();
        }

        #endregion
    }
}
