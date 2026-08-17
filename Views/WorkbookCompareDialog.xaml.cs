using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using ExcelSupport.Models;
using Microsoft.Win32;
using WpfMessageBox = System.Windows.MessageBox;
using WpfButton = System.Windows.Controls.Button;
using WpfSaveFileDialog = Microsoft.Win32.SaveFileDialog;

namespace ExcelSupport.Views
{
    public partial class WorkbookCompareDialog : Window, INotifyPropertyChanged
    {
        private bool _isDarkTheme;
        private readonly ObservableCollection<CompareDiffItem> _diffResults = new ObservableCollection<CompareDiffItem>();
        private ICollectionView? _diffView;
        private bool _isInitializing = true;

        public bool IsDarkTheme
        {
            get => _isDarkTheme;
            set
            {
                if (_isDarkTheme != value)
                {
                    _isDarkTheme = value;
                    OnPropertyChanged(nameof(IsDarkTheme));
                }
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private static WorkbookCompareDialog? _currentInstance;

        internal static void ShowWindow(string? defaultWb1Name = null, bool isDarkTheme = false)
        {
            if (_currentInstance != null && _currentInstance.IsLoaded)
            {
                _currentInstance.IsDarkTheme = isDarkTheme;
                if (!string.IsNullOrEmpty(defaultWb1Name) && _currentInstance.CboWorkbook1.Items.Contains(defaultWb1Name))
                {
                    _currentInstance.CboWorkbook1.SelectedItem = defaultWb1Name;
                }
                _currentInstance.Activate();
                return;
            }

            _currentInstance = new WorkbookCompareDialog(defaultWb1Name, isDarkTheme);
            _currentInstance.Closed += (s, e) => _currentInstance = null;

            try
            {
                var addIn = AddInEvents.Instance;
                if (addIn?.ExcelAppInstance != null)
                {
                    new System.Windows.Interop.WindowInteropHelper(_currentInstance).Owner = (IntPtr)addIn.ExcelAppInstance.Hwnd;
                }
            }
            catch { }

            _currentInstance.Show();
        }

        public WorkbookCompareDialog(string? defaultWb1Name = null, bool isDarkTheme = false)
        {
            InitializeComponent();
            IsDarkTheme = isDarkTheme;
            DataContext = this;

            _diffView = CollectionViewSource.GetDefaultView(_diffResults);
            _diffView.Filter = FilterDiffItem;
            GridDiffResults.ItemsSource = _diffView;

            InitKeyColumns();
            LoadWorkbooks(defaultWb1Name);

            _isInitializing = false;
        }

        private void InitKeyColumns()
        {
            CboKeyColumn.Items.Clear();
            for (int i = 1; i <= 26; i++)
            {
                string colLetter = ((char)('A' + (i - 1))).ToString();
                CboKeyColumn.Items.Add($"Cột {colLetter} ({i})");
            }
            CboKeyColumn.SelectedIndex = 0; // Default Column A
        }

        private void LoadWorkbooks(string? defaultWb1Name)
        {
            var addIn = AddInEvents.Instance;
            if (addIn == null) return;

            var wbList = addIn.GetOpenWorkbookNamesList();

            CboWorkbook1.Items.Clear();
            CboWorkbook2.Items.Clear();

            foreach (var wbName in wbList)
            {
                CboWorkbook1.Items.Add(wbName);
                CboWorkbook2.Items.Add(wbName);
            }

            if (CboWorkbook1.Items.Count > 0)
            {
                if (!string.IsNullOrEmpty(defaultWb1Name) && CboWorkbook1.Items.Contains(defaultWb1Name))
                {
                    CboWorkbook1.SelectedItem = defaultWb1Name;
                }
                else
                {
                    CboWorkbook1.SelectedIndex = 0;
                }
            }

            if (CboWorkbook2.Items.Count > 1)
            {
                CboWorkbook2.SelectedIndex = CboWorkbook1.SelectedIndex == 0 ? 1 : 0;
            }
            else if (CboWorkbook2.Items.Count > 0)
            {
                CboWorkbook2.SelectedIndex = 0;
            }

            RefreshSheets1();
            RefreshSheets2();
        }

        private void OnWorkbook1SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!_isInitializing)
            {
                RefreshSheets1();
            }
        }

        private void OnWorkbook2SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!_isInitializing)
            {
                RefreshSheets2();
            }
        }

        private void RefreshSheets1()
        {
            var addIn = AddInEvents.Instance;
            if (addIn == null) return;

            string wbName = CboWorkbook1.SelectedItem?.ToString() ?? string.Empty;
            CboSheet1.Items.Clear();
            CboSheet1.Items.Add("(Tất cả Sheet cùng tên)");

            if (!string.IsNullOrEmpty(wbName))
            {
                var sheets = addIn.GetWorksheetNamesList(wbName);
                foreach (var s in sheets)
                {
                    CboSheet1.Items.Add(s);
                }
            }

            CboSheet1.SelectedIndex = 0;
        }

        private void RefreshSheets2()
        {
            var addIn = AddInEvents.Instance;
            if (addIn == null) return;

            string wbName = CboWorkbook2.SelectedItem?.ToString() ?? string.Empty;
            CboSheet2.Items.Clear();
            CboSheet2.Items.Add("(Tất cả Sheet cùng tên)");

            if (!string.IsNullOrEmpty(wbName))
            {
                var sheets = addIn.GetWorksheetNamesList(wbName);
                foreach (var s in sheets)
                {
                    CboSheet2.Items.Add(s);
                }
            }

            CboSheet2.SelectedIndex = 0;
        }

        private void OnCompareClick(object sender, RoutedEventArgs e)
        {
            ExecuteCompare();
        }

        private void ExecuteCompare()
        {
            var addIn = AddInEvents.Instance;
            if (addIn == null) return;

            string wb1Name = CboWorkbook1.SelectedItem?.ToString() ?? string.Empty;
            string wb2Name = CboWorkbook2.SelectedItem?.ToString() ?? string.Empty;

            if (string.IsNullOrEmpty(wb1Name) || string.IsNullOrEmpty(wb2Name))
            {
                WpfMessageBox.Show("Vui lòng chọn đầy đủ File A và File B để so sánh.", "Chưa chọn file",
                                   MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (wb1Name == wb2Name && CboSheet1.SelectedIndex == 0 && CboSheet2.SelectedIndex == 0)
            {
                WpfMessageBox.Show("Vui lòng chọn 2 Workbook khác nhau, hoặc chọn cụ thể 2 Sheet khác nhau trong cùng 1 Workbook để so sánh.", "Thông báo",
                                   MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string? ws1Name = CboSheet1.SelectedIndex > 0 ? CboSheet1.SelectedItem?.ToString() : null;
            string? ws2Name = CboSheet2.SelectedIndex > 0 ? CboSheet2.SelectedItem?.ToString() : null;

            var options = new CompareOptions
            {
                Mode = RbKeyColumn.IsChecked == true ? CompareMode.KeyColumn : CompareMode.CellByCell,
                KeyColumnIndex = CboKeyColumn.SelectedIndex + 1,
                IgnoreWhitespace = ChkIgnoreWhitespace.IsChecked == true,
                CaseInsensitive = ChkCaseInsensitive.IsChecked == true,
                CompareFormulas = ChkCompareFormulas.IsChecked == true
            };

            TxtStatus.Text = "⏳ Đang so sánh dữ liệu...";
            _diffResults.Clear();

            var list = addIn.CompareWorkbooksOrSheets(wb1Name, ws1Name, wb2Name, ws2Name, options, msg =>
            {
                TxtStatus.Text = msg;
            });

            foreach (var item in list)
            {
                _diffResults.Add(item);
            }

            UpdateStatsBadges();

            if (list.Count == 0)
            {
                TxtStatus.Text = "✅ Hai file / sheet hoàn toàn trùng khớp! Không có bất kỳ sai khác nào.";
            }
            else
            {
                TxtStatus.Text = $"⚠️ Tìm thấy {list.Count} điểm sai khác. Click đúp vào dòng để chuyển tới ô tương ứng trên Excel.";
            }
        }

        private void UpdateStatsBadges()
        {
            int modifiedCount = 0;
            int addedCount = 0;
            int deletedCount = 0;

            foreach (var item in _diffResults)
            {
                switch (item.Type)
                {
                    case DiffType.Modified: modifiedCount++; break;
                    case DiffType.Added: addedCount++; break;
                    case DiffType.Deleted: deletedCount++; break;
                }
            }

            TxtCountModified.Text = $"{modifiedCount} thay đổi";
            TxtCountAdded.Text = $"{addedCount} thêm mới";
            TxtCountDeleted.Text = $"{deletedCount} đã xóa";
        }

        private bool FilterDiffItem(object obj)
        {
            if (obj is CompareDiffItem item)
            {
                // 1. Lọc theo loại
                if (RbFilterModified.IsChecked == true && item.Type != DiffType.Modified) return false;
                if (RbFilterAdded.IsChecked == true && item.Type != DiffType.Added) return false;
                if (RbFilterDeleted.IsChecked == true && item.Type != DiffType.Deleted) return false;

                // 2. Lọc theo từ khóa tìm kiếm
                string search = TxtSearch.Text?.Trim() ?? string.Empty;
                if (string.IsNullOrEmpty(search)) return true;

                return (!string.IsNullOrEmpty(item.SheetName) && item.SheetName.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0) ||
                       (!string.IsNullOrEmpty(item.CellAddress) && item.CellAddress.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0) ||
                       (!string.IsNullOrEmpty(item.KeyIdentifier) && item.KeyIdentifier.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0) ||
                       (!string.IsNullOrEmpty(item.OldValue) && item.OldValue.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0) ||
                       (!string.IsNullOrEmpty(item.NewValue) && item.NewValue.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0) ||
                       (!string.IsNullOrEmpty(item.TypeDescription) && item.TypeDescription.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0);
            }
            return false;
        }

        private void OnDiffFilterChanged(object sender, RoutedEventArgs e)
        {
            _diffView?.Refresh();
        }

        private void OnSearchTextChanged(object sender, TextChangedEventArgs e)
        {
            _diffView?.Refresh();
        }

        private void OnClearSearchClick(object sender, RoutedEventArgs e)
        {
            TxtSearch.Text = string.Empty;
        }

        private void OnRowDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (GridDiffResults.SelectedItem is CompareDiffItem item)
            {
                NavigateToItem(item);
            }
        }

        private void OnGoToCellClick(object sender, RoutedEventArgs e)
        {
            if (sender is WpfButton btn && btn.DataContext is CompareDiffItem item)
            {
                NavigateToItem(item);
            }
        }

        private void NavigateToItem(CompareDiffItem item)
        {
            var addIn = AddInEvents.Instance;
            if (addIn == null) return;

            string targetWb = !string.IsNullOrEmpty(item.Workbook2Name) ? item.Workbook2Name : item.Workbook1Name;
            string targetWs = item.SheetName;
            string cleanAddr = item.CellAddress.Split(' ')[0];

            if (!string.IsNullOrEmpty(cleanAddr) && !cleanAddr.StartsWith("Dòng"))
            {
                addIn.NavigateToCell(targetWb, targetWs, cleanAddr);
            }
            else
            {
                addIn.NavigateToCell(targetWb, targetWs, "A1");
            }
        }

        private void OnHighlightClick(object sender, RoutedEventArgs e)
        {
            if (_diffResults.Count == 0)
            {
                WpfMessageBox.Show("Không có kết quả sai khác nào để tô màu.", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var addIn = AddInEvents.Instance;
            if (addIn == null) return;

            string wb2Name = CboWorkbook2.SelectedItem?.ToString() ?? string.Empty;
            string sheetName = _diffResults[0].SheetName;

            var list = new List<CompareDiffItem>(_diffResults);
            addIn.HighlightDiffInWorksheet(list, wb2Name, sheetName);
        }

        private void OnCreateReportSheetClick(object sender, RoutedEventArgs e)
        {
            if (_diffResults.Count == 0)
            {
                WpfMessageBox.Show("Không có kết quả sai khác nào để tạo báo cáo.", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var addIn = AddInEvents.Instance;
            if (addIn == null) return;

            string targetWb = CboWorkbook2.SelectedItem?.ToString() ?? CboWorkbook1.SelectedItem?.ToString() ?? string.Empty;
            var list = new List<CompareDiffItem>(_diffResults);
            addIn.CreateDiffReportSheet(targetWb, list);
        }

        private void OnExportCsvClick(object sender, RoutedEventArgs e)
        {
            if (_diffResults.Count == 0)
            {
                WpfMessageBox.Show("Danh sách kết quả đang trống, không có dữ liệu để xuất file.", "Thông báo",
                                   MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var saveDialog = new WpfSaveFileDialog
            {
                Filter = "File CSV (*.csv)|*.csv",
                FileName = $"Diff_Report_{DateTime.Now:yyyyMMdd_HHmmss}.csv",
                Title = "Lưu Báo Cáo Sai Khác Ra File CSV"
            };

            if (saveDialog.ShowDialog() == true)
            {
                try
                {
                    var sb = new StringBuilder();
                    sb.AppendLine("STT,Tên Sheet,Vị Trí / Khóa,Loại Sai Khác,Giá Trị File A (Gốc),Giá Trị File B (Mới)");

                    foreach (var item in _diffResults)
                    {
                        string line = $"{item.Index}," +
                                      $"\"{EscapeCsv(item.SheetName)}\"," +
                                      $"\"{EscapeCsv(item.CellAddress)}\"," +
                                      $"\"{EscapeCsv(item.TypeDescription)}\"," +
                                      $"\"{EscapeCsv(item.OldValue)}\"," +
                                      $"\"{EscapeCsv(item.NewValue)}\"";
                        sb.AppendLine(line);
                    }

                    var utf8WithBom = new UTF8Encoding(true);
                    File.WriteAllText(saveDialog.FileName, sb.ToString(), utf8WithBom);

                    WpfMessageBox.Show($"✅ Đã xuất báo cáo thành công ra file:\n{saveDialog.FileName}",
                                       "Xuất CSV Thành Công", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    WpfMessageBox.Show($"Lỗi khi xuất file CSV:\n{ex.Message}", "Lỗi Xuất File",
                                       MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private static string EscapeCsv(string? val)
        {
            if (string.IsNullOrEmpty(val)) return string.Empty;
            return val!.Replace("\"", "\"\"");
        }

        private void OnCloseClick(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
