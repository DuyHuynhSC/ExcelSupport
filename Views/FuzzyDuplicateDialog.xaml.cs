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
    public partial class FuzzyDuplicateDialog : System.Windows.Window
    {
        public static readonly DependencyProperty IsDarkThemeProperty =
            DependencyProperty.Register(nameof(IsDarkTheme), typeof(bool), typeof(FuzzyDuplicateDialog),
                new PropertyMetadata(false));

        public bool IsDarkTheme
        {
            get => (bool)GetValue(IsDarkThemeProperty);
            set => SetValue(IsDarkThemeProperty, value);
        }

        private static FuzzyDuplicateDialog? _currentInstance;

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

                _currentInstance = new FuzzyDuplicateDialog(app)
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
                WpfMessageBox.Show($"Lỗi mở màn hình Quét Trùng Lặp Ảo:\n{ex.Message}",
                                   "ExcelSupport", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private readonly ExcelApp? _excelApp;
        private List<MergeColumnItem> _columns = new List<MergeColumnItem>();
        private List<FuzzyClusterGroup> _currentClusters = new List<FuzzyClusterGroup>();
        private List<FuzzyRecordItem> _displayItems = new List<FuzzyRecordItem>();

        public FuzzyDuplicateDialog(ExcelApp? app)
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
            LoadSheets();
        }

        private void LoadSheets()
        {
            if (_excelApp == null) return;

            try
            {
                var wb = _excelApp.ActiveWorkbook;
                if (wb == null) return;

                var sheetNames = new List<string>();
                foreach (_Worksheet s in wb.Worksheets) sheetNames.Add(s.Name);

                CboSheets.ItemsSource = sheetNames;

                string actSheet = (_excelApp.ActiveSheet as _Worksheet)?.Name ?? "";
                if (!string.IsNullOrEmpty(actSheet) && sheetNames.Contains(actSheet))
                {
                    CboSheets.SelectedItem = actSheet;
                }
                else if (sheetNames.Count > 0)
                {
                    CboSheets.SelectedIndex = 0;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"LoadSheets error: {ex.Message}");
            }
        }

        private void OnSheetSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string? sheetName = CboSheets.SelectedItem as string;
            if (string.IsNullOrEmpty(sheetName) || _excelApp == null) return;

            string wbName = _excelApp.ActiveWorkbook?.Name ?? string.Empty;
            _columns = TableMergeService.GetSheetColumns(_excelApp, wbName, sheetName!);

            CboColumns.ItemsSource = _columns;
            CboColumns.DisplayMemberPath = nameof(MergeColumnItem.DisplayText);
            CboColumns.SelectedValuePath = nameof(MergeColumnItem.ColumnIndex);

            if (_columns.Count > 0)
            {
                CboColumns.SelectedIndex = 0;
            }
        }

        private void OnSliderThresholdChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (TxtThresholdValue != null)
            {
                TxtThresholdValue.Text = $"{(int)SliderThreshold.Value}%";
            }
        }

        private void OnStartScanClick(object sender, RoutedEventArgs e)
        {
            if (_excelApp == null) return;

            string? sheetName = CboSheets.SelectedItem as string;
            var selectedCol = CboColumns.SelectedItem as MergeColumnItem;

            if (string.IsNullOrEmpty(sheetName) || selectedCol == null)
            {
                WpfMessageBox.Show("Vui lòng chọn Sheet và Cột dữ liệu cần quét.", "Thông Báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            _Worksheet? ws = null;
            try
            {
                foreach (_Worksheet s in _excelApp.ActiveWorkbook.Worksheets)
                {
                    if (string.Equals(s.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        ws = s;
                        break;
                    }
                }

                if (ws == null)
                {
                    WpfMessageBox.Show("Không tìm thấy Sheet được chọn.", "Thông Báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                TxtFooterStatus.Text = "⏳ Đang quét và phân tích dữ liệu tương đồng...";

                var options = new FuzzyScanOptions
                {
                    TargetColumnIndex = selectedCol.ColumnIndex,
                    StartRow = 2,
                    SimilarityThreshold = SliderThreshold.Value,
                    Algorithm = FuzzyMatchAlgorithm.JaroWinkler,
                    CleanInvisibleSpaces = (ChkCleanSpaces.IsChecked == true),
                    IgnoreAccent = (ChkIgnoreAccent.IsChecked == true),
                    IgnoreCase = (ChkIgnoreCase.IsChecked == true)
                };

                _currentClusters = FuzzyDuplicateService.ScanFuzzyDuplicates(ws, options);
                TxtClusterCountBadge.Text = LocalizationService.Get("Fuzzy_ClusterBadge", _currentClusters.Count);

                _displayItems = _currentClusters.SelectMany(c => c.Items).ToList();
                GridClusters.ItemsSource = null;
                GridClusters.ItemsSource = _displayItems;

                if (_currentClusters.Count == 0)
                {
                    TxtFooterStatus.Text = "✅ Không phát hiện dữ liệu trùng lặp ảo nào ở ngưỡng tương đồng này.";
                    WpfMessageBox.Show("Không tìm thấy dữ liệu trùng lặp ảo nào trong cột được chọn.", "Kết Quả Quét", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    int totalVariants = _currentClusters.Sum(c => c.Count);
                    TxtFooterStatus.Text = $"✅ Đã phát hiện {_currentClusters.Count:N0} nhóm ({totalVariants:N0} ô) có dữ liệu trùng lặp ảo!";
                }
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Lỗi quét dữ liệu:\n{ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                TxtFooterStatus.Text = "❌ Đã xảy ra lỗi trong quá trình quét.";
            }
        }

        private void OnStandardizeClick(object sender, RoutedEventArgs e)
        {
            if (_excelApp == null || _currentClusters.Count == 0)
            {
                WpfMessageBox.Show("Không có danh sách trùng lặp ảo nào để chuẩn hóa. Vui lòng bấm Quét trước.", "Thông Báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var selectedCol = CboColumns.SelectedItem as MergeColumnItem;
            if (selectedCol == null) return;

            string? sheetName = CboSheets.SelectedItem as string;
            _Worksheet? ws = null;
            foreach (_Worksheet s in _excelApp.ActiveWorkbook.Worksheets)
            {
                if (string.Equals(s.Name, sheetName, StringComparison.OrdinalIgnoreCase)) { ws = s; break; }
            }
            if (ws == null) return;

            var confirm = WpfMessageBox.Show(
                "Bạn có chắc chắn muốn chuẩn hóa tất cả các biến thể đã chọn về giá trị chuẩn của từng nhóm không?",
                "Xác Nhận Chuẩn Hóa",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (confirm != MessageBoxResult.Yes) return;

            try
            {
                int count = FuzzyDuplicateService.StandardizeValues(ws, selectedCol.ColumnIndex, _currentClusters);
                TxtFooterStatus.Text = $"✅ Đã chuẩn hóa thành công {count:N0} ô dữ liệu!";
                WpfMessageBox.Show($"Đã chuẩn hóa thành công {count:N0} ô dữ liệu về giá trị chuẩn!", "Hoàn Tất", MessageBoxButton.OK, MessageBoxImage.Information);

                // Quét lại để cập nhật
                OnStartScanClick(sender, e);
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Lỗi chuẩn hóa:\n{ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OnHighlightClick(object sender, RoutedEventArgs e)
        {
            if (_excelApp == null || _currentClusters.Count == 0)
            {
                WpfMessageBox.Show("Không có danh sách trùng lặp nào để tô màu rà soát.", "Thông Báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var selectedCol = CboColumns.SelectedItem as MergeColumnItem;
            if (selectedCol == null) return;

            string? sheetName = CboSheets.SelectedItem as string;
            _Worksheet? ws = null;
            foreach (_Worksheet s in _excelApp.ActiveWorkbook.Worksheets)
            {
                if (string.Equals(s.Name, sheetName, StringComparison.OrdinalIgnoreCase)) { ws = s; break; }
            }
            if (ws == null) return;

            try
            {
                int count = FuzzyDuplicateService.HighlightClusters(ws, selectedCol.ColumnIndex, _currentClusters);
                TxtFooterStatus.Text = $"🎨 Đã tô màu phân biệt {count:N0} ô trùng lặp ảo trên Sheet!";
                WpfMessageBox.Show($"Đã tô màu đánh dấu {count:N0} ô thuộc {_currentClusters.Count:N0} nhóm trùng lặp trên Sheet!", "Tô Màu Thành Công", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Lỗi tô màu:\n{ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OnCloseClick(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
