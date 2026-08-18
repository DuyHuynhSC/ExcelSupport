using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using ExcelSupport.Models;
using WpfMessageBox = System.Windows.MessageBox;

namespace ExcelSupport.Views
{
    public partial class DuplicateFinderDialog : Window, INotifyPropertyChanged
    {
        private bool _isDarkTheme;
        private bool _isInitialized = false;
        private static DuplicateFinderDialog? _currentInstance;

        public ObservableCollection<ColumnSelectionItem> ColumnsList { get; set; } = new ObservableCollection<ColumnSelectionItem>();
        public ObservableCollection<DuplicateGroupItem> AllDuplicates { get; set; } = new ObservableCollection<DuplicateGroupItem>();
        private ICollectionView _duplicatesView;

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

                _currentInstance = new DuplicateFinderDialog(isDarkTheme);
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
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Lỗi mở màn hình Tìm trùng lặp:\n{ex.Message}", "Lỗi Khởi Tạo",
                                   MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        public DuplicateFinderDialog(bool isDarkTheme = false)
        {
            InitializeComponent();
            _isInitialized = true;
            IsDarkTheme = isDarkTheme;
            DataContext = this;

            _duplicatesView = CollectionViewSource.GetDefaultView(AllDuplicates);
            _duplicatesView.Filter = FilterDuplicateRecord;
            GridDuplicates.ItemsSource = _duplicatesView;

            ItemsColumns.ItemsSource = ColumnsList;

            Loaded += (s, e) => ReloadColumns();
        }

        private void ReloadColumns()
        {
            if (!_isInitialized) return;
            ColumnsList.Clear();
            var addIn = AddInEvents.Instance;
            if (addIn == null) return;

            bool header = ChkFirstRowHeader?.IsChecked == true;
            var cols = addIn.GetActiveSheetColumnsInfo(header);
            foreach (var col in cols)
            {
                ColumnsList.Add(col);
            }
        }

        private void OnHeaderOptionChanged(object sender, RoutedEventArgs e)
        {
            if (!_isInitialized) return;
            ReloadColumns();
        }

        private void OnSelectAllColumnsClick(object sender, RoutedEventArgs e)
        {
            foreach (var col in ColumnsList) col.IsSelected = true;
        }

        private void OnDeselectAllColumnsClick(object sender, RoutedEventArgs e)
        {
            foreach (var col in ColumnsList) col.IsSelected = false;
        }

        private void OnFindDuplicatesClick(object sender, RoutedEventArgs e)
        {
            var addIn = AddInEvents.Instance;
            if (addIn == null) return;

            var options = new DuplicateFinderOptions
            {
                Mode = RbFuzzyMatch?.IsChecked == true ? DuplicateMatchMode.FuzzyMatch : DuplicateMatchMode.ExactMatch,
                FuzzySimilarityThreshold = SliderThreshold?.Value ?? 0.85,
                FirstRowIsHeader = ChkFirstRowHeader?.IsChecked == true,
                IgnoreWhitespace = ChkIgnoreSpaces?.IsChecked == true,
                CaseInsensitive = ChkIgnoreCase?.IsChecked == true,
                SelectedColumnIndices = ColumnsList.Where(c => c.IsSelected).Select(c => c.ColumnIndex).ToList()
            };

            if (options.SelectedColumnIndices.Count == 0)
            {
                WpfMessageBox.Show("Vui lòng tích chọn ít nhất 1 cột làm khóa so sánh.", "Thông Báo",
                                   MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            TxtStatus.Text = "⏳ Đang quét và phân tích dữ liệu...";
            AllDuplicates.Clear();

            var results = addIn.FindDuplicateGroups(options, msg => Dispatcher.Invoke(() => TxtStatus.Text = msg));

            foreach (var item in results)
            {
                AllDuplicates.Add(item);
            }

            int groupCount = results.Select(r => r.GroupId).Distinct().Count();
            TxtTotalDuplicates.Text = $"{results.Count} dòng trùng";
            TxtTotalGroups.Text = $"{groupCount} nhóm trùng";

            _duplicatesView.Refresh();
        }

        private bool FilterDuplicateRecord(object obj)
        {
            if (obj is not DuplicateGroupItem item) return false;

            // 1. Filter duplicate only
            if (RbFilterDuplicatesOnly?.IsChecked == true && item.IsMaster) return false;

            // 2. Filter search text
            string query = TxtSearch?.Text?.Trim() ?? string.Empty;
            if (!string.IsNullOrEmpty(query))
            {
                if (!item.RowValuesSummary.Contains(query) &&
                    !item.KeySummary.Contains(query) &&
                    !item.RowDisplay.Contains(query) &&
                    !item.GroupTitle.Contains(query))
                {
                    return false;
                }
            }

            return true;
        }

        private void OnFilterChanged(object sender, RoutedEventArgs e)
        {
            if (!_isInitialized) return;
            _duplicatesView?.Refresh();
        }

        private void OnSearchTextChanged(object sender, TextChangedEventArgs e)
        {
            if (!_isInitialized) return;
            _duplicatesView?.Refresh();
        }

        private void OnRowDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (GridDuplicates.SelectedItem is DuplicateGroupItem item)
            {
                NavigateToRow(item);
            }
        }

        private void OnGoToRowClick(object sender, RoutedEventArgs e)
        {
            if ((sender as FrameworkElement)?.DataContext is DuplicateGroupItem item)
            {
                NavigateToRow(item);
            }
        }

        private void NavigateToRow(DuplicateGroupItem item)
        {
            ExcelDna.Integration.ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    var addIn = AddInEvents.Instance;
                    if (addIn?.ExcelAppInstance == null) return;

                    dynamic app = addIn.ExcelAppInstance;
                    dynamic ws = app.ActiveSheet;
                    if (ws != null)
                    {
                        dynamic rowRange = ws.Rows[item.RowIndex];
                        rowRange.Select();
                        app.ActiveWindow.ScrollRow = Math.Max(1, item.RowIndex - 2);
                    }
                }
                catch { }
            });
        }

        private void OnHighlightClick(object sender, RoutedEventArgs e)
        {
            var addIn = AddInEvents.Instance;
            if (addIn == null || AllDuplicates.Count == 0) return;

            addIn.HighlightDuplicatesInWorksheet(AllDuplicates.ToList());
        }

        private void OnDeleteDuplicatesClick(object sender, RoutedEventArgs e)
        {
            var addIn = AddInEvents.Instance;
            if (addIn == null || AllDuplicates.Count == 0) return;

            bool success = addIn.DeleteDuplicateRowsInWorksheet(AllDuplicates.ToList(), keepFirst: true);
            if (success)
            {
                AllDuplicates.Clear();
                TxtTotalDuplicates.Text = "0 dòng trùng";
                TxtTotalGroups.Text = "0 nhóm trùng";
                _duplicatesView.Refresh();
            }
        }

        private void OnExtractToSheetClick(object sender, RoutedEventArgs e)
        {
            var addIn = AddInEvents.Instance;
            if (addIn == null || AllDuplicates.Count == 0) return;

            addIn.ExtractDuplicatesToNewSheet(AllDuplicates.ToList());
        }

        private void OnCloseClick(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
