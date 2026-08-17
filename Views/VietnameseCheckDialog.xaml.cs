using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using ExcelSupport.Models;
using MessageBox = System.Windows.MessageBox;

namespace ExcelSupport.Views
{
    public partial class VietnameseCheckDialog : Window, INotifyPropertyChanged
    {
        private bool _isDarkTheme;
        private readonly ObservableCollection<VietnameseLocationItem> _results = new ObservableCollection<VietnameseLocationItem>();
        private readonly ICollectionView _view;

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

        private static VietnameseCheckDialog? _currentInstance;

        internal static void ShowWindow(bool isDarkTheme = false)
        {
            if (_currentInstance != null && _currentInstance.IsLoaded)
            {
                _currentInstance.IsDarkTheme = isDarkTheme;
                _currentInstance.Activate();
                _currentInstance.ExecuteScan();
                return;
            }

            _currentInstance = new VietnameseCheckDialog(isDarkTheme);
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

        private bool _isInitialized = false;

        public VietnameseCheckDialog(bool isDarkTheme = false)
        {
            InitializeComponent();
            IsDarkTheme = isDarkTheme;
            DataContext = this;

            _view = CollectionViewSource.GetDefaultView(_results);
            _view.Filter = FilterResult;
            GridResults.ItemsSource = _view;

            _isInitialized = true;

            // Auto scan current workbook when opened
            Loaded += (s, e) => ExecuteScan();
        }

        private void OnScopeRadioChecked(object sender, RoutedEventArgs e)
        {
            if (_isInitialized)
            {
                ExecuteScan();
            }
        }

        private bool FilterResult(object obj)
        {
            if (obj is VietnameseLocationItem item)
            {
                string search = TxtSearch.Text?.Trim() ?? string.Empty;
                if (string.IsNullOrEmpty(search)) return true;

                return (!string.IsNullOrEmpty(item.WorkbookName) && item.WorkbookName.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0) ||
                       (!string.IsNullOrEmpty(item.SheetName) && item.SheetName.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0) ||
                       (!string.IsNullOrEmpty(item.CellAddress) && item.CellAddress.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0) ||
                       (!string.IsNullOrEmpty(item.TextContent) && item.TextContent.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0) ||
                       (!string.IsNullOrEmpty(item.TypeDescription) && item.TypeDescription.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0);
            }
            return true;
        }

        private void OnSearchTextChanged(object sender, TextChangedEventArgs e)
        {
            _view?.Refresh();
            UpdateCountBadge();
        }

        private void UpdateCountBadge()
        {
            int total = _results.Count;
            if (total == 0)
            {
                TxtTotalCount.Text = "0 vị trí (Không có tiếng Việt)";
                BadgeResult.Background = new SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(IsDarkTheme ? "#064E3B" : "#DCFCE7"));
                BadgeResult.BorderBrush = new SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(IsDarkTheme ? "#059669" : "#86EFAC"));
                TxtTotalCount.Foreground = new SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(IsDarkTheme ? "#6EE7B7" : "#166534"));
            }
            else
            {
                TxtTotalCount.Text = $"{total} vị trí phát hiện";
                BadgeResult.Background = new SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(IsDarkTheme ? "#78350F" : "#FEF3C7"));
                BadgeResult.BorderBrush = new SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(IsDarkTheme ? "#D97706" : "#FCD34D"));
                TxtTotalCount.Foreground = new SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(IsDarkTheme ? "#FDE68A" : "#92400E"));
            }
        }

        private void OnScanClick(object sender, RoutedEventArgs e)
        {
            ExecuteScan();
        }

        private void ExecuteScan()
        {
            var addIn = AddInEvents.Instance;
            if (addIn == null) return;

            AddInEvents.VietnameseScanScope scope = AddInEvents.VietnameseScanScope.ActiveWorkbook;
            if (RbActiveSheet.IsChecked == true) scope = AddInEvents.VietnameseScanScope.ActiveSheet;
            else if (RbAllWorkbooks.IsChecked == true) scope = AddInEvents.VietnameseScanScope.AllWorkbooks;

            TxtStatus.Text = "⏳ Đang quét kiểm tra tiếng Việt...";
            _results.Clear();

            var list = addIn.ScanVietnameseLocations(scope, msg =>
            {
                TxtStatus.Text = msg;
            });

            foreach (var item in list)
            {
                _results.Add(item);
            }

            UpdateCountBadge();

            if (list.Count == 0)
            {
                TxtStatus.Text = "✅ Tuyệt vời! Không phát hiện bất kỳ ký tự tiếng Việt có dấu nào trong phạm vi đã quét.";
            }
            else
            {
                TxtStatus.Text = $"⚠️ Tìm thấy {list.Count} vị trí có tiếng Việt. Click đúp vào dòng để chuyển đến ô tương ứng trong Excel.";
            }
        }

        private void OnRowDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (GridResults.SelectedItem is VietnameseLocationItem item)
            {
                NavigateToItem(item);
            }
        }

        private void OnGoToCellClick(object sender, RoutedEventArgs e)
        {
            if (sender is System.Windows.Controls.Button btn && btn.DataContext is VietnameseLocationItem item)
            {
                NavigateToItem(item);
            }
        }

        private void NavigateToItem(VietnameseLocationItem item)
        {
            var addIn = AddInEvents.Instance;
            if (addIn == null) return;

            bool ok = addIn.NavigateToCell(item.WorkbookName, item.SheetName, item.CellAddress);
            if (ok)
            {
                TxtStatus.Text = $"🎯 Đã chuyển đến [{item.WorkbookName}] ➔ [{item.SheetName}] ➔ Ô: {item.CellAddress}";
            }
            else
            {
                TxtStatus.Text = $"❌ Không thể chuyển đến ô [{item.CellAddress}] trong [{item.SheetName}].";
            }
        }

        private void OnCreateReportSheetClick(object sender, RoutedEventArgs e)
        {
            if (_results.Count == 0)
            {
                MessageBox.Show("Không có vị trí tiếng Việt nào để tạo báo cáo.", "Tạo Báo Cáo", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var addIn = AddInEvents.Instance;
            if (addIn == null) return;

            addIn.CreateVietnameseReportSheet(_results.ToList());
        }

        private void OnExportCsvClick(object sender, RoutedEventArgs e)
        {
            if (_results.Count == 0)
            {
                MessageBox.Show("Danh sách kết quả đang trống, không có dữ liệu để xuất.", "Xuất Báo Cáo", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            using (var dlg = new SaveFileDialog())
            {
                dlg.Title = "Lưu file kết quả kiểm tra tiếng Việt (CSV)";
                dlg.Filter = "File CSV (*.csv)|*.csv";
                dlg.FileName = $"Vietnamese_Check_Report_{DateTime.Now:yyyyMMdd_HHmm}";

                if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    try
                    {
                        var sb = new StringBuilder();
                        sb.AppendLine("STT,Workbook,Sheet,CellAddress,LocationType,TextContent");

                        foreach (var item in _results)
                        {
                            string wb = EscapeCsv(item.WorkbookName);
                            string ws = EscapeCsv(item.SheetName);
                            string addr = EscapeCsv(item.CellAddress);
                            string type = EscapeCsv(item.TypeDescription);
                            string content = EscapeCsv(item.TextContent);

                            sb.AppendLine($"{item.Index},{wb},{ws},{addr},{type},{content}");
                        }

                        var utf8WithBom = new UTF8Encoding(encoderShouldEmitUTF8Identifier: true);
                        File.WriteAllText(dlg.FileName, sb.ToString(), utf8WithBom);

                        TxtStatus.Text = $"✅ Đã xuất {_results.Count} vị trí ra file CSV thành công!";
                        MessageBox.Show($"✅ Đã xuất {_results.Count} vị trí ra file:\n{dlg.FileName}", "Xuất CSV Thành Công", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Lỗi khi xuất file:\n{ex.Message}", "Lỗi Export", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
        }

        private static string EscapeCsv(string field)
        {
            if (string.IsNullOrEmpty(field)) return string.Empty;
            if (field.Contains(",") || field.Contains("\"") || field.Contains("\n") || field.Contains("\r"))
            {
                return "\"" + field.Replace("\"", "\"\"") + "\"";
            }
            return field;
        }

        private void OnCloseClick(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
