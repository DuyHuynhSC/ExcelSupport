using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using Microsoft.Office.Interop.Excel;
using ExcelSupport.Services;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using Window = System.Windows.Window;

namespace ExcelSupport.Views
{
    public partial class BatchCleanerAndMergeDialog : Window
    {
        private static BatchCleanerAndMergeDialog? _currentInstance;

        public static readonly DependencyProperty IsDarkThemeProperty =
            DependencyProperty.Register("IsDarkTheme", typeof(bool), typeof(BatchCleanerAndMergeDialog), new PropertyMetadata(false));

        public bool IsDarkTheme
        {
            get => (bool)GetValue(IsDarkThemeProperty);
            set => SetValue(IsDarkThemeProperty, value);
        }

        internal static void ShowWindow(int defaultTab = 0, bool isDarkTheme = false)
        {
            try
            {
                if (_currentInstance != null && _currentInstance.IsLoaded)
                {
                    _currentInstance.IsDarkTheme = isDarkTheme;
                    if (defaultTab >= 0 && defaultTab < _currentInstance.MainTabControl.Items.Count)
                    {
                        _currentInstance.MainTabControl.SelectedIndex = defaultTab;
                    }
                    _currentInstance.Activate();
                    return;
                }

                var addIn = AddInEvents.Instance;
                var app = addIn?.ExcelAppInstance;

                _currentInstance = new BatchCleanerAndMergeDialog(app, defaultTab, isDarkTheme);

                try
                {
                    if (app != null)
                    {
                        new WindowInteropHelper(_currentInstance).Owner = (IntPtr)app.Hwnd;
                    }
                }
                catch { }

                _currentInstance.ShowDialog();
                _currentInstance = null;
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Lỗi mở màn hình Tiện ích Xóa Trống & Gộp Ô/Sheet:\n{ex.Message}",
                                               "ExcelSupport", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private readonly ExcelApp? _excelApp;
        private readonly ObservableCollection<SheetItemInfo> _sheetItems = new ObservableCollection<SheetItemInfo>();

        public BatchCleanerAndMergeDialog(ExcelApp? app, int defaultTab = 0, bool isDarkTheme = false)
        {
            InitializeComponent();
            _excelApp = app;
            IsDarkTheme = isDarkTheme;

            ListSheets.ItemsSource = _sheetItems;

            Loaded += (s, e) =>
            {
                InitializeData();
                if (defaultTab >= 0 && defaultTab < MainTabControl.Items.Count)
                {
                    MainTabControl.SelectedIndex = defaultTab;
                }
            };
        }

        private void InitializeData()
        {
            if (_excelApp == null) return;

            try
            {
                var activeWs = _excelApp.ActiveSheet as _Worksheet;
                var activeWb = _excelApp.ActiveWorkbook;

                if (activeWs != null)
                {
                    TxtActiveSheetBadge.Text = $"Sheet: {activeWs.Name}";

                    // Nạp danh sách cột cho CboKeyColumn
                    CboKeyColumn.Items.Clear();
                    Range? usedRange = null;
                    try
                    {
                        usedRange = activeWs.UsedRange;
                        int totalCols = usedRange?.Columns.Count ?? 26;
                        int startCol = usedRange?.Column ?? 1;

                        for (int c = startCol; c < startCol + totalCols; c++)
                        {
                            string colLetter = GetColumnLetter(c);
                            string headerText = "";
                            try
                            {
                                var cell = activeWs.Cells[1, c] as Range;
                                if (cell?.Value2 != null)
                                {
                                    headerText = $" — {cell.Value2}";
                                }
                                if (cell != null) Marshal.ReleaseComObject(cell);
                            }
                            catch { }

                            CboKeyColumn.Items.Add($"Cột {colLetter}{headerText}");
                        }
                    }
                    catch { }
                    finally
                    {
                        if (usedRange != null) Marshal.ReleaseComObject(usedRange);
                    }

                    if (CboKeyColumn.Items.Count > 0) CboKeyColumn.SelectedIndex = 0;

                    // Thông tin vùng chọn Tab 2
                    Range? sel = _excelApp.Selection as Range;
                    if (sel != null)
                    {
                        TxtSelectionRangeInfo.Text = $"Vùng chọn: [{activeWs.Name}!{sel.Address[false, false]}] ({sel.Rows.Count:N0} dòng x {sel.Columns.Count:N0} cột)";
                        Marshal.ReleaseComObject(sel);
                    }
                    else
                    {
                        TxtSelectionRangeInfo.Text = "Chưa có vùng chọn hợp lệ. Vui lòng bôi đen các ô trên Excel.";
                    }

                    Marshal.ReleaseComObject(activeWs);
                }

                // Nạp danh sách Sheet cho Tab 3
                _sheetItems.Clear();
                if (activeWb != null)
                {
                    foreach (_Worksheet ws in activeWb.Worksheets)
                    {
                        int rCount = 0;
                        int cCount = 0;
                        Range? ur = null;
                        try
                        {
                            ur = ws.UsedRange;
                            if (ur != null)
                            {
                                rCount = ur.Rows.Count;
                                cCount = ur.Columns.Count;
                            }
                        }
                        catch { }
                        finally
                        {
                            if (ur != null) Marshal.ReleaseComObject(ur);
                        }

                        _sheetItems.Add(new SheetItemInfo
                        {
                            SheetName = ws.Name,
                            WorkbookName = activeWb.Name,
                            TotalRows = rCount,
                            TotalCols = cCount,
                            IsSelected = true
                        });

                        Marshal.ReleaseComObject(ws);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Init Data error: {ex.Message}");
            }
        }

        private static string GetColumnLetter(int colIndex)
        {
            int div = colIndex;
            string colLetter = string.Empty;
            while (div > 0)
            {
                int mod = (div - 1) % 26;
                colLetter = (char)(65 + mod) + colLetter;
                div = (div - mod) / 26;
            }
            return colLetter;
        }

        #region TAB 1: XÓA DÒNG / CỘT TRỐNG

        private void OnExecuteBlankCleanupClick(object sender, RoutedEventArgs e)
        {
            if (_excelApp == null) return;

            try
            {
                // Xác định Target
                BlankCleanupTarget target = BlankCleanupTarget.EntirelyBlankRows;
                int keyColIndex = 1;

                if (RbTargetBlankKeyCol.IsChecked == true)
                {
                    target = BlankCleanupTarget.BlankRowsInKeyColumn;
                    keyColIndex = Math.Max(1, CboKeyColumn.SelectedIndex + 1);
                }
                else if (RbTargetEntireBlankCols.IsChecked == true)
                {
                    target = BlankCleanupTarget.EntirelyBlankColumns;
                }

                // Xác định Action
                BlankCleanupAction action = BlankCleanupAction.Delete;
                if (RbActionHide.IsChecked == true) action = BlankCleanupAction.Hide;
                else if (RbActionHighlight.IsChecked == true) action = BlankCleanupAction.Highlight;

                // Xác định Scope
                var targetSheets = new List<_Worksheet>();

                if (RbScopeActiveSheet.IsChecked == true)
                {
                    var ws = _excelApp.ActiveSheet as _Worksheet;
                    if (ws != null) targetSheets.Add(ws);
                }
                else if (RbScopeAllSheets.IsChecked == true)
                {
                    var wb = _excelApp.ActiveWorkbook;
                    if (wb != null)
                    {
                        foreach (_Worksheet s in wb.Worksheets) targetSheets.Add(s);
                    }
                }
                else // AllWorkbooks
                {
                    foreach (Workbook wb in _excelApp.Workbooks)
                    {
                        foreach (_Worksheet s in wb.Worksheets) targetSheets.Add(s);
                    }
                }

                if (targetSheets.Count == 0)
                {
                    System.Windows.MessageBox.Show("Không tìm thấy Sheet nào để xử lý.", "ExcelSupport", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                string actionVerb = action == BlankCleanupAction.Delete ? "xóa vĩnh viễn" : (action == BlankCleanupAction.Hide ? "ẩn" : "tô màu");
                string targetNoun = target == BlankCleanupTarget.EntirelyBlankColumns ? "cột trống" : "dòng trống";

                var confirm = System.Windows.MessageBox.Show(
                    $"Bạn có chắc chắn muốn tiến hành {actionVerb} toàn bộ các {targetNoun} trên {targetSheets.Count} Sheet không?",
                    "Xác Nhận Xử Lý Trống", MessageBoxButton.YesNo, MessageBoxImage.Question);

                if (confirm != MessageBoxResult.Yes) return;

                int totalProcessed = 0;
                int sheetsAffected = 0;

                foreach (var ws in targetSheets)
                {
                    var (count, _) = BatchCleanerAndMergeService.ProcessBlankInSheet(ws, target, action, keyColIndex);
                    if (count > 0)
                    {
                        totalProcessed += count;
                        sheetsAffected++;
                    }
                    Marshal.ReleaseComObject(ws);
                }

                System.Windows.MessageBox.Show(
                    $"Hoàn tất!\n\nĐã thực hiện {actionVerb} tổng cộng {totalProcessed:N0} {targetNoun} trên {sheetsAffected} Sheet.",
                    "Kết Quả Xử Lý", MessageBoxButton.OK, MessageBoxImage.Information);

                InitializeData();
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Lỗi trong quá trình xử lý: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #endregion

        #region TAB 2: GỘP Ô BẢO TOÀN DỮ LIỆU

        private void OnExecuteSafeMergeClick(object sender, RoutedEventArgs e)
        {
            if (_excelApp == null) return;

            try
            {
                var activeWs = _excelApp.ActiveSheet as _Worksheet;
                Range? sel = _excelApp.Selection as Range;

                if (activeWs == null || sel == null)
                {
                    System.Windows.MessageBox.Show("Vui lòng chọn một vùng ô trên Excel trước khi bấm Gộp.", "ExcelSupport", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var options = new SafeMergeOptions
                {
                    IgnoreBlankCells = ChkMergeIgnoreBlank.IsChecked == true,
                    TrimSpaces = ChkMergeTrim.IsChecked == true
                };

                // Direction
                if (RbMergeAcross.IsChecked == true) options.Direction = SafeMergeDirection.AcrossRows;
                else if (RbMergeDown.IsChecked == true) options.Direction = SafeMergeDirection.DownColumns;
                else options.Direction = SafeMergeDirection.AllToOneCell;

                // Separator
                if (RbSepComma.IsChecked == true) options.SeparatorType = MergeSeparatorType.Comma;
                else if (RbSepSpace.IsChecked == true) options.SeparatorType = MergeSeparatorType.Space;
                else if (RbSepSemicolon.IsChecked == true) options.SeparatorType = MergeSeparatorType.Semicolon;
                else if (RbSepNewLine.IsChecked == true) options.SeparatorType = MergeSeparatorType.NewLine;
                else if (RbSepPipe.IsChecked == true) options.SeparatorType = MergeSeparatorType.Pipe;
                else
                {
                    options.SeparatorType = MergeSeparatorType.Custom;
                    options.CustomSeparator = TxtCustomSep.Text;
                }

                var (success, msg, _) = BatchCleanerAndMergeService.MergeSelectedCellsSafely(activeWs, sel, options);

                if (success)
                {
                    System.Windows.MessageBox.Show(msg, "Gộp Ô Thành Công", MessageBoxButton.OK, MessageBoxImage.Information);
                    Close();
                }
                else
                {
                    System.Windows.MessageBox.Show(msg, "Thông Báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                }

                Marshal.ReleaseComObject(sel);
                Marshal.ReleaseComObject(activeWs);
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Lỗi khi gộp ô: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #endregion

        #region TAB 3: GỘP NHIỀU SHEETS

        private void OnSelectAllSheetsClick(object sender, RoutedEventArgs e)
        {
            foreach (var item in _sheetItems) item.IsSelected = true;
            ListSheets.Items.Refresh();
        }

        private void OnDeselectAllSheetsClick(object sender, RoutedEventArgs e)
        {
            foreach (var item in _sheetItems) item.IsSelected = false;
            ListSheets.Items.Refresh();
        }

        private void OnExecuteCombineSheetsClick(object sender, RoutedEventArgs e)
        {
            if (_excelApp == null) return;

            try
            {
                var activeWb = _excelApp.ActiveWorkbook;
                if (activeWb == null)
                {
                    System.Windows.MessageBox.Show("Không có Workbook nào đang mở.", "ExcelSupport", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var selectedSheetNames = _sheetItems.Where(x => x.IsSelected).Select(x => x.SheetName).ToList();
                if (selectedSheetNames.Count == 0)
                {
                    System.Windows.MessageBox.Show("Vui lòng tick chọn ít nhất 1 Sheet trong danh sách để gộp.", "ExcelSupport", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var sourceSheets = new List<_Worksheet>();
                foreach (string name in selectedSheetNames)
                {
                    try
                    {
                        var s = activeWb.Worksheets[name] as _Worksheet;
                        if (s != null) sourceSheets.Add(s);
                    }
                    catch { }
                }

                var options = new CombineSheetsOptions
                {
                    HasHeaderRow = ChkHasHeader.IsChecked == true,
                    HeaderRowCount = 1,
                    AddSourceColumn = ChkAddSourceCol.IsChecked == true,
                    SourceColumnHeader = string.IsNullOrWhiteSpace(TxtSourceColHeader.Text) ? "Tên Sheet Nguồn" : TxtSourceColHeader.Text.Trim(),
                    SkipBlankRows = ChkCombineSkipBlank.IsChecked == true
                };

                var (success, msg, _) = BatchCleanerAndMergeService.CombineSheetsIntoOne(activeWb, sourceSheets, options);

                foreach (var s in sourceSheets) Marshal.ReleaseComObject(s);

                if (success)
                {
                    System.Windows.MessageBox.Show(msg, "Gộp Sheet Thành Công", MessageBoxButton.OK, MessageBoxImage.Information);
                    Close();
                }
                else
                {
                    System.Windows.MessageBox.Show(msg, "Thông Báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Lỗi khi gộp Sheet: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #endregion

        private void OnCloseClick(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
