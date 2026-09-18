using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using ExcelSupport.Host;
using ExcelSupport.Models;
using ExcelSupport.Services;
using Microsoft.Office.Interop.Excel;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using WpfMessageBox = System.Windows.MessageBox;

namespace ExcelSupport.Views
{
    public partial class VisualTableMergeDialog : System.Windows.Window
    {
        public static readonly DependencyProperty IsDarkThemeProperty =
            DependencyProperty.Register(nameof(IsDarkTheme), typeof(bool), typeof(VisualTableMergeDialog),
                new PropertyMetadata(false));

        public bool IsDarkTheme
        {
            get => (bool)GetValue(IsDarkThemeProperty);
            set => SetValue(IsDarkThemeProperty, value);
        }

        private static VisualTableMergeDialog? _currentInstance;

        internal static void ShowWindow(bool isDarkTheme = false)
        {
            try
            {
                if (_currentInstance != null && _currentInstance.IsLoaded)
                {
                    _currentInstance.IsDarkTheme = isDarkTheme;
                    _currentInstance.Activate();
                    return;
                }

                var addIn = AddInEvents.Instance;
                var app = addIn?.ExcelAppInstance;

                _currentInstance = new VisualTableMergeDialog(app)
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
                WpfMessageBox.Show($"Lỗi mở màn hình Trộn & Ghép Nối Bảng:\n{ex.Message}",
                                   "ExcelSupport", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private readonly ExcelApp? _excelApp;
        private List<MergeColumnItem> _table1Cols = new List<MergeColumnItem>();
        private List<MergeColumnItem> _table2Cols = new List<MergeColumnItem>();

        public VisualTableMergeDialog(ExcelApp? app)
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
            LoadWorkbooks();
        }

        private void LoadWorkbooks()
        {
            if (_excelApp == null) return;

            try
            {
                var wbNames = new List<string>();
                foreach (Workbook wb in _excelApp.Workbooks)
                {
                    wbNames.Add(wb.Name);
                }

                CboWb1.ItemsSource = wbNames.ToList();
                CboWb2.ItemsSource = wbNames.ToList();

                string activeWbName = _excelApp.ActiveWorkbook?.Name ?? string.Empty;
                if (!string.IsNullOrEmpty(activeWbName))
                {
                    CboWb1.SelectedItem = activeWbName;
                    CboWb2.SelectedItem = activeWbName;
                }
                else if (wbNames.Count > 0)
                {
                    CboWb1.SelectedIndex = 0;
                    CboWb2.SelectedIndex = 0;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"LoadWorkbooks error: {ex.Message}");
            }
        }

        private void OnWb1SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string? wbName = CboWb1.SelectedItem as string;
            if (string.IsNullOrEmpty(wbName) || _excelApp == null) return;

            try
            {
                foreach (Workbook wb in _excelApp.Workbooks)
                {
                    if (string.Equals(wb.Name, wbName, StringComparison.OrdinalIgnoreCase))
                    {
                        var sheetNames = new List<string>();
                        foreach (_Worksheet s in wb.Worksheets) sheetNames.Add(s.Name);
                        CboSheet1.ItemsSource = sheetNames;

                        string actSheet = (wb == _excelApp.ActiveWorkbook) ? (_excelApp.ActiveSheet as _Worksheet)?.Name ?? "" : "";
                        if (!string.IsNullOrEmpty(actSheet) && sheetNames.Contains(actSheet))
                            CboSheet1.SelectedItem = actSheet;
                        else if (sheetNames.Count > 0)
                            CboSheet1.SelectedIndex = 0;

                        break;
                    }
                }
            }
            catch { }
        }

        private void OnSheet1SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string? wbName = CboWb1.SelectedItem as string;
            string? sheetName = CboSheet1.SelectedItem as string;
            if (string.IsNullOrEmpty(wbName) || string.IsNullOrEmpty(sheetName)) return;

            _table1Cols = TableMergeService.GetSheetColumns(_excelApp, wbName!, sheetName!);
            CboKeyCol1.ItemsSource = _table1Cols;
            CboKeyCol1.DisplayMemberPath = nameof(MergeColumnItem.DisplayText);
            CboKeyCol1.SelectedValuePath = nameof(MergeColumnItem.ColumnIndex);

            if (_table1Cols.Count > 0)
            {
                CboKeyCol1.SelectedIndex = 0;
            }
        }

        private void OnWb2SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string? wbName = CboWb2.SelectedItem as string;
            if (string.IsNullOrEmpty(wbName) || _excelApp == null) return;

            try
            {
                foreach (Workbook wb in _excelApp.Workbooks)
                {
                    if (string.Equals(wb.Name, wbName, StringComparison.OrdinalIgnoreCase))
                    {
                        var sheetNames = new List<string>();
                        foreach (_Worksheet s in wb.Worksheets) sheetNames.Add(s.Name);
                        CboSheet2.ItemsSource = sheetNames;

                        if (sheetNames.Count > 1 && CboWb1.SelectedItem == CboWb2.SelectedItem)
                        {
                            // Mặc định chọn Sheet 2 nếu cùng workbook
                            CboSheet2.SelectedIndex = 1;
                        }
                        else if (sheetNames.Count > 0)
                        {
                            CboSheet2.SelectedIndex = 0;
                        }
                        break;
                    }
                }
            }
            catch { }
        }

        private void OnSheet2SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string? wbName = CboWb2.SelectedItem as string;
            string? sheetName = CboSheet2.SelectedItem as string;
            if (string.IsNullOrEmpty(wbName) || string.IsNullOrEmpty(sheetName)) return;

            _table2Cols = TableMergeService.GetSheetColumns(_excelApp, wbName!, sheetName!);
            CboKeyCol2.ItemsSource = _table2Cols;
            CboKeyCol2.DisplayMemberPath = nameof(MergeColumnItem.DisplayText);
            CboKeyCol2.SelectedValuePath = nameof(MergeColumnItem.ColumnIndex);

            if (_table2Cols.Count > 0)
            {
                CboKeyCol2.SelectedIndex = 0;
            }

            // Nạp các cột vào DataGrid (Mặc định chọn các cột không phải là cột khóa)
            GridColumns2.ItemsSource = null;
            GridColumns2.ItemsSource = _table2Cols;
        }

        private void OnSelectAllCols2Click(object sender, RoutedEventArgs e)
        {
            foreach (var col in _table2Cols) col.IsSelected = true;
            GridColumns2.Items.Refresh();
        }

        private void OnDeselectAllCols2Click(object sender, RoutedEventArgs e)
        {
            foreach (var col in _table2Cols) col.IsSelected = false;
            GridColumns2.Items.Refresh();
        }

        private void OnExecuteMergeClick(object sender, RoutedEventArgs e)
        {
            if (_excelApp == null) return;

            string? wb1 = CboWb1.SelectedItem as string;
            string? s1 = CboSheet1.SelectedItem as string;
            string? wb2 = CboWb2.SelectedItem as string;
            string? s2 = CboSheet2.SelectedItem as string;

            if (string.IsNullOrEmpty(wb1) || string.IsNullOrEmpty(s1))
            {
                WpfMessageBox.Show("Vui lòng chọn Bảng 1 (Bảng nguồn).", "Thông Báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrEmpty(wb2) || string.IsNullOrEmpty(s2))
            {
                WpfMessageBox.Show("Vui lòng chọn Bảng 2 (Bảng đối chiếu tra cứu).", "Thông Báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var key1Item = CboKeyCol1.SelectedItem as MergeColumnItem;
            var key2Item = CboKeyCol2.SelectedItem as MergeColumnItem;

            if (key1Item == null || key2Item == null)
            {
                WpfMessageBox.Show("Vui lòng chọn Cột Khóa trên cả 2 bảng để đối chiếu.", "Thông Báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var selectedCols2 = _table2Cols.Where(c => c.IsSelected).ToList();
            if (selectedCols2.Count == 0)
            {
                WpfMessageBox.Show("Vui lòng chọn ít nhất một cột từ Bảng 2 cần ghép sang.", "Thông Báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var joinType = TableJoinType.LeftJoin;
            if (RbJoinInner.IsChecked == true) joinType = TableJoinType.InnerJoin;
            else if (RbJoinOuter.IsChecked == true) joinType = TableJoinType.FullOuterJoin;

            var outputTarget = (RbOutputAdjacent.IsChecked == true)
                ? TableMergeOutputTarget.InsertAdjacentToTable1
                : TableMergeOutputTarget.CreateNewWorksheet;

            var options = new TableMergeOptions
            {
                Table1WorkbookName = wb1!,
                Table1SheetName = s1!,
                Table1KeyColIndex = key1Item.ColumnIndex,
                Table1HeaderRow = 1,

                Table2WorkbookName = wb2!,
                Table2SheetName = s2!,
                Table2KeyColIndex = key2Item.ColumnIndex,
                Table2HeaderRow = 1,

                SelectedColumnsFromTable2 = selectedCols2,
                JoinType = joinType,
                OutputTarget = outputTarget,

                TrimSpaces = (ChkTrim.IsChecked == true),
                IgnoreAccent = (ChkIgnoreAccent.IsChecked == true),
                MatchCase = (ChkMatchCase.IsChecked == true)
            };

            try
            {
                TxtFooterStatus.Text = "⏳ Đang xử lý trộn & ghép nối bảng dữ liệu...";

                var result = TableMergeService.ExecuteTableMerge(_excelApp, options);

                TxtFooterStatus.Text = $"✅ {result.Message}";

                if (result.Success)
                {
                    WpfMessageBox.Show(
                        $"{result.Message}\n\n• Tổng số dòng kết quả: {result.TotalRowsMerged:N0}\n• Số dòng khớp mã khóa: {result.MatchedRows:N0}\n• Số dòng không khớp: {result.UnmatchedRows:N0}",
                        "Ghép Bảng Thành Công",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                }
                else
                {
                    WpfMessageBox.Show(result.Message, "Thông Báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Lỗi thực thi ghép bảng:\n{ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                TxtFooterStatus.Text = "❌ Đã xảy ra lỗi trong quá trình ghép bảng.";
            }
        }

        private void OnCloseClick(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
