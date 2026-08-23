using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;
using ExcelSupport.Helpers;
using ExcelSupport.Models;
using ExcelSupport.Services;
using WpfMessageBox = System.Windows.MessageBox;
using WpfMessageBoxButton = System.Windows.MessageBoxButton;
using WpfMessageBoxImage = System.Windows.MessageBoxImage;
using MediaColor = System.Windows.Media.Color;
using WpfBinding = System.Windows.Data.Binding;
using WpfClipboard = System.Windows.Clipboard;

namespace ExcelSupport.Views
{
    public partial class OracleTableCompareDialog : Window
    {
        private readonly ObservableCollection<OracleTableColumnInfo> _tableColumns = new ObservableCollection<OracleTableColumnInfo>();
        private OracleCompareResult? _lastResult;
        private List<OracleRowDiffItem> _allDiffItems = new List<OracleRowDiffItem>();
        private bool _isComparing = false;
        private readonly bool _isDarkTheme;

        public OracleTableCompareDialog(bool isDarkTheme = false)
        {
            InitializeComponent();
            _isDarkTheme = isDarkTheme;

            icKeyColumns.ItemsSource = _tableColumns;

            if (_isDarkTheme)
            {
                ApplyDarkTheme();
            }
        }

        private static OracleTableCompareDialog? _currentInstance;

        public static void ShowWindow(bool isDarkTheme = false)
        {
            try
            {
                if (_currentInstance != null && _currentInstance.IsLoaded)
                {
                    _currentInstance.Activate();
                    return;
                }

                _currentInstance = new OracleTableCompareDialog(isDarkTheme);

                try
                {
                    var addIn = AddInEvents.Instance;
                    if (addIn?.ExcelAppInstance != null)
                    {
                        var helper = new System.Windows.Interop.WindowInteropHelper(_currentInstance);
                        helper.Owner = new IntPtr(addIn.ExcelAppInstance.Hwnd);
                    }
                }
                catch { }

                _currentInstance.Show();
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Không thể mở hộp thoại Đối soát Oracle:\n{ex.Message}",
                                   "Lỗi giao diện", WpfMessageBoxButton.OK, WpfMessageBoxImage.Error);
            }
        }

        private void ApplyDarkTheme()
        {
            RootGrid.Background = new SolidColorBrush(MediaColor.FromRgb(15, 23, 42)); // Slate 900
        }

        #region Connection Testing & Metadata Loading

        private OracleConnectionConfig BuildConfigA()
        {
            int.TryParse(txtPortA.Text, out int port);
            if (port <= 0) port = 1521;

            return new OracleConnectionConfig
            {
                Host = txtHostA.Text.Trim(),
                Port = port,
                ServiceNameOrSid = txtServiceA.Text.Trim(),
                ServiceType = (rbSidA.IsChecked == true) ? OracleServiceNameType.SID : OracleServiceNameType.ServiceName,
                Username = txtUserA.Text.Trim(),
                Password = txtPassA.Password
            };
        }

        private OracleConnectionConfig BuildConfigB()
        {
            int.TryParse(txtPortB.Text, out int port);
            if (port <= 0) port = 1521;

            return new OracleConnectionConfig
            {
                Host = txtHostB.Text.Trim(),
                Port = port,
                ServiceNameOrSid = txtServiceB.Text.Trim(),
                ServiceType = (rbSidB.IsChecked == true) ? OracleServiceNameType.SID : OracleServiceNameType.ServiceName,
                Username = txtUserB.Text.Trim(),
                Password = txtPassB.Password
            };
        }

        private async void BtnTestA_Click(object sender, RoutedEventArgs e)
        {
            btnTestA.IsEnabled = false;
            txtStatusBadgeA.Text = "⏳ Đang kết nối...";
            txtStatusBadgeA.Foreground = new SolidColorBrush(MediaColor.FromRgb(217, 119, 6)); // Amber

            var config = BuildConfigA();
            var (success, msg, version) = await OracleDataCompareService.TestConnectionAsync(config);

            btnTestA.IsEnabled = true;
            if (success)
            {
                txtStatusBadgeA.Text = $"🟢 {version}";
                txtStatusBadgeA.Foreground = new SolidColorBrush(MediaColor.FromRgb(22, 163, 74)); // Green
                txtStatus.Text = $"DB A kết nối thành công: {version}";

                await LoadSchemasAAsync(config);
            }
            else
            {
                txtStatusBadgeA.Text = "🔴 Thất bại";
                txtStatusBadgeA.Foreground = new SolidColorBrush(MediaColor.FromRgb(220, 38, 38)); // Red
                txtStatus.Text = msg;
                WpfMessageBox.Show(msg, "Lỗi kết nối Database A", WpfMessageBoxButton.OK, WpfMessageBoxImage.Error);
            }
        }

        private async void BtnTestB_Click(object sender, RoutedEventArgs e)
        {
            btnTestB.IsEnabled = false;
            txtStatusBadgeB.Text = "⏳ Đang kết nối...";
            txtStatusBadgeB.Foreground = new SolidColorBrush(MediaColor.FromRgb(217, 119, 6)); // Amber

            var config = BuildConfigB();
            var (success, msg, version) = await OracleDataCompareService.TestConnectionAsync(config);

            btnTestB.IsEnabled = true;
            if (success)
            {
                txtStatusBadgeB.Text = $"🟢 {version}";
                txtStatusBadgeB.Foreground = new SolidColorBrush(MediaColor.FromRgb(22, 163, 74)); // Green
                txtStatus.Text = $"DB B kết nối thành công: {version}";

                await LoadSchemasBAsync(config);
            }
            else
            {
                txtStatusBadgeB.Text = "🔴 Thất bại";
                txtStatusBadgeB.Foreground = new SolidColorBrush(MediaColor.FromRgb(220, 38, 38)); // Red
                txtStatus.Text = msg;
                WpfMessageBox.Show(msg, "Lỗi kết nối Database B", WpfMessageBoxButton.OK, WpfMessageBoxImage.Error);
            }
        }

        private async Task LoadSchemasAAsync(OracleConnectionConfig config)
        {
            var schemas = await OracleDataCompareService.GetSchemasAsync(config);
            cboSchemaA.ItemsSource = schemas;
            if (schemas.Contains(config.Username.ToUpperInvariant()))
            {
                cboSchemaA.SelectedItem = config.Username.ToUpperInvariant();
            }
            else if (schemas.Count > 0)
            {
                cboSchemaA.SelectedIndex = 0;
            }
        }

        private async Task LoadSchemasBAsync(OracleConnectionConfig config)
        {
            var schemas = await OracleDataCompareService.GetSchemasAsync(config);
            cboSchemaB.ItemsSource = schemas;
            if (schemas.Contains(config.Username.ToUpperInvariant()))
            {
                cboSchemaB.SelectedItem = config.Username.ToUpperInvariant();
            }
            else if (schemas.Count > 0)
            {
                cboSchemaB.SelectedIndex = 0;
            }
        }

        private async void CboSchemaA_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string? schema = cboSchemaA.SelectedItem as string ?? cboSchemaA.Text;
            if (string.IsNullOrWhiteSpace(schema)) return;

            var config = BuildConfigA();
            var tables = await OracleDataCompareService.GetTablesAndViewsAsync(config, schema);
            cboTableA.ItemsSource = tables;
            if (tables.Count > 0) cboTableA.SelectedIndex = 0;
        }

        private async void CboSchemaB_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string? schema = cboSchemaB.SelectedItem as string ?? cboSchemaB.Text;
            if (string.IsNullOrWhiteSpace(schema)) return;

            var config = BuildConfigB();
            var tables = await OracleDataCompareService.GetTablesAndViewsAsync(config, schema);
            cboTableB.ItemsSource = tables;
            if (tables.Count > 0) cboTableB.SelectedIndex = 0;
        }

        private async void CboTableA_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string? schema = cboSchemaA.SelectedItem as string ?? cboSchemaA.Text;
            string? table = cboTableA.SelectedItem as string ?? cboTableA.Text;
            if (string.IsNullOrWhiteSpace(table)) return;

            // Đồng bộ tên bảng sang DB B nếu DB B chưa chọn
            if (string.IsNullOrWhiteSpace(cboTableB.Text) && cboTableB.Items.Contains(table))
            {
                cboTableB.SelectedItem = table;
            }

            var config = BuildConfigA();
            var cols = await OracleDataCompareService.GetTableColumnsAsync(config, schema ?? "", table);
            _tableColumns.Clear();
            foreach (var col in cols)
            {
                _tableColumns.Add(col);
            }
        }

        private void RbMode_Checked(object sender, RoutedEventArgs e)
        {
            if (!IsLoaded || icKeyColumns == null || rbModeKey == null) return;
            icKeyColumns.Visibility = (rbModeKey.IsChecked == true) ? Visibility.Visible : Visibility.Collapsed;
        }

        #endregion

        #region Comparison Execution

        private async void BtnStartCompare_Click(object sender, RoutedEventArgs e)
        {
            if (_isComparing) return;

            string schemaA = cboSchemaA.Text.Trim();
            string tableA = cboTableA.Text.Trim();
            string schemaB = cboSchemaB.Text.Trim();
            string tableB = cboTableB.Text.Trim();

            if (string.IsNullOrWhiteSpace(tableA) || string.IsNullOrWhiteSpace(tableB))
            {
                WpfMessageBox.Show("Vui lòng chọn bảng dữ liệu cho cả Database A và Database B trước khi so sánh.",
                                   "Thiếu thông tin bảng", WpfMessageBoxButton.OK, WpfMessageBoxImage.Warning);
                return;
            }

            var configA = BuildConfigA();
            var configB = BuildConfigB();

            if (string.IsNullOrWhiteSpace(configA.Username) || string.IsNullOrWhiteSpace(configB.Username))
            {
                WpfMessageBox.Show("Vui lòng nhập đầy đủ thông tin đăng nhập cho cả 2 Database.",
                                   "Thiếu thông tin kết nối", WpfMessageBoxButton.OK, WpfMessageBoxImage.Warning);
                return;
            }

            var options = new OracleCompareOptions
            {
                Mode = (rbModeKey.IsChecked == true) ? OracleCompareMode.ByKeyColumns : OracleCompareMode.Sequential,
                IgnoreWhitespace = chkIgnoreSpaces.IsChecked == true,
                IgnoreCase = chkIgnoreCase.IsChecked == true,
                SelectedKeyColumns = _tableColumns.Where(c => c.IsSelectedKey).Select(c => c.ColumnName).ToList(),
                SelectedCompareColumns = _tableColumns.Where(c => c.IsSelectedCompare).Select(c => c.ColumnName).ToList()
            };

            if (int.TryParse(txtMaxRows.Text, out int maxR) && maxR > 0)
            {
                options.MaxRows = maxR;
            }

            if (options.Mode == OracleCompareMode.ByKeyColumns && options.SelectedKeyColumns.Count == 0)
            {
                WpfMessageBox.Show("Chế độ so sánh theo Khóa yêu cầu ít nhất 1 cột khóa được chọn. Vui lòng tích chọn cột làm khóa ở danh sách bên dưới.",
                                   "Chưa chọn Cột Khóa", WpfMessageBoxButton.OK, WpfMessageBoxImage.Warning);
                return;
            }

            _isComparing = true;
            btnStartCompare.IsEnabled = false;
            pbProgress.Visibility = Visibility.Visible;
            pbProgress.Value = 0;

            var progressReporter = new Progress<(string StatusText, double ProgressPercent)>(report =>
            {
                txtStatus.Text = report.StatusText;
                pbProgress.Value = report.ProgressPercent;
            });

            try
            {
                var result = await OracleDataCompareService.CompareTablesAsync(configA, configB, schemaA, tableA, schemaB, tableB, options, progressReporter);
                _lastResult = result;
                _allDiffItems = result.DiffItems;

                UpdateStatisticsAndGrid(result);
                txtStatus.Text = $"Hoàn tất đối soát trong {result.ExecutionTime.TotalSeconds:F2}s. Khớp: {result.MatchCount:N0} dòng, Sai lệch: {result.ModifiedCount:N0} dòng, Chỉ có ở A: {result.MissingInBCount:N0}, Chỉ có ở B: {result.MissingInACount:N0}.";
            }
            catch (Exception ex)
            {
                txtStatus.Text = $"Lỗi: {ex.Message}";
                WpfMessageBox.Show(ex.Message, "Lỗi đối soát dữ liệu", WpfMessageBoxButton.OK, WpfMessageBoxImage.Error);
            }
            finally
            {
                _isComparing = false;
                btnStartCompare.IsEnabled = true;
                pbProgress.Visibility = Visibility.Collapsed;
            }
        }

        private void UpdateStatisticsAndGrid(OracleCompareResult result)
        {
            txtStatTotal.Text = $"Tổng: {result.DiffItems.Count:N0} dòng (A: {result.TotalRowsA:N0} | B: {result.TotalRowsB:N0})";
            txtStatDiff.Text = $"Sai lệch: {result.ModifiedCount:N0} (A: -{result.MissingInBCount:N0}, B: +{result.MissingInACount:N0})";
            txtStatMatch.Text = $"Khớp: {result.MatchCount:N0}";

            // Rebuild dynamic columns in DataGrid
            while (dgDiffResults.Columns.Count > 4)
            {
                dgDiffResults.Columns.RemoveAt(dgDiffResults.Columns.Count - 1);
            }

            foreach (var col in result.Columns)
            {
                // Col in DB A
                var colA = new DataGridTextColumn
                {
                    Header = $"{col} (A)",
                    Binding = new WpfBinding($"RowValuesA[{col}]"),
                    Width = new DataGridLength(1, DataGridLengthUnitType.Auto)
                };

                // Col in DB B
                var colB = new DataGridTextColumn
                {
                    Header = $"{col} (B)",
                    Binding = new WpfBinding($"RowValuesB[{col}]"),
                    Width = new DataGridLength(1, DataGridLengthUnitType.Auto)
                };

                dgDiffResults.Columns.Add(colA);
                dgDiffResults.Columns.Add(colB);
            }

            ApplyFilter();
        }

        #endregion

        #region Filtering & Search

        private void Filter_Checked(object sender, RoutedEventArgs e)
        {
            if (!IsLoaded) return;
            ApplyFilter();
        }

        private void TxtSearchFilter_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!IsLoaded) return;
            ApplyFilter();
        }

        private void ApplyFilter()
        {
            if (dgDiffResults == null) return;

            if (_allDiffItems == null || _allDiffItems.Count == 0)
            {
                dgDiffResults.ItemsSource = null;
                return;
            }

            IEnumerable<OracleRowDiffItem> query = _allDiffItems;

            if (rbFilterDiff?.IsChecked == true)
            {
                query = query.Where(r => r.Status == OracleRowDiffStatus.Modified);
            }
            else if (rbFilterMissingB?.IsChecked == true)
            {
                query = query.Where(r => r.Status == OracleRowDiffStatus.MissingInB);
            }
            else if (rbFilterMissingA?.IsChecked == true)
            {
                query = query.Where(r => r.Status == OracleRowDiffStatus.MissingInA);
            }
            else if (rbFilterMatch?.IsChecked == true)
            {
                query = query.Where(r => r.Status == OracleRowDiffStatus.Identical);
            }

            string filterText = txtSearchFilter?.Text?.Trim() ?? string.Empty;
            if (!string.IsNullOrEmpty(filterText))
            {
                query = query.Where(r =>
                    r.KeyDisplay.IndexOf(filterText, StringComparison.OrdinalIgnoreCase) >= 0 ||
                    r.DifferingColumnsSummary.IndexOf(filterText, StringComparison.OrdinalIgnoreCase) >= 0 ||
                    r.RowValuesA.Values.Any(v => v != null && v.ToString()!.IndexOf(filterText, StringComparison.OrdinalIgnoreCase) >= 0) ||
                    r.RowValuesB.Values.Any(v => v != null && v.ToString()!.IndexOf(filterText, StringComparison.OrdinalIgnoreCase) >= 0));
            }

            dgDiffResults.ItemsSource = query.ToList();
        }

        #endregion

        #region Actions & Excel Output

        private void BtnInsertActiveSelection_Click(object sender, RoutedEventArgs e)
        {
            if (_lastResult == null || _lastResult.DiffItems.Count == 0)
            {
                WpfMessageBox.Show("Vui lòng tiến hành so sánh dữ liệu trước khi chèn vào Excel.",
                                   "Chưa có kết quả", WpfMessageBoxButton.OK, WpfMessageBoxImage.Information);
                return;
            }

            try
            {
                var app = AddInEvents.Instance?.ExcelAppInstance;
                if (app == null)
                {
                    WpfMessageBox.Show("Không thể kết nối với tiến trình Excel.", "Lỗi Excel", WpfMessageBoxButton.OK, WpfMessageBoxImage.Error);
                    return;
                }

                OracleDataCompareService.InsertDiffToActiveSelection(_lastResult, app, highlightOnlyDiffs: false);
                txtStatus.Text = "Đã chèn dữ liệu so sánh và tô màu trực tiếp vào vị trí đang chọn trong Excel!";
                WpfMessageBox.Show("Đã chèn bảng kết quả so sánh và tô màu nổi bật các ô sai lệch vào vị trí đang chọn thành công!",
                                   "Thành công", WpfMessageBoxButton.OK, WpfMessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show(ex.Message, "Lỗi chèn vào Excel", WpfMessageBoxButton.OK, WpfMessageBoxImage.Error);
            }
        }

        private void BtnExportExcel_Click(object sender, RoutedEventArgs e)
        {
            if (_lastResult == null || _lastResult.DiffItems.Count == 0)
            {
                WpfMessageBox.Show("Vui lòng tiến hành so sánh dữ liệu trước khi xuất báo cáo Excel.",
                                   "Chưa có kết quả", WpfMessageBoxButton.OK, WpfMessageBoxImage.Information);
                return;
            }

            try
            {
                var app = AddInEvents.Instance?.ExcelAppInstance;
                if (app == null)
                {
                    WpfMessageBox.Show("Không thể kết nối với tiến trình Excel.", "Lỗi Excel", WpfMessageBoxButton.OK, WpfMessageBoxImage.Error);
                    return;
                }

                OracleDataCompareService.ExportDiffReportToExcel(_lastResult, app, highlightOnlyDiffs: false);
                txtStatus.Text = "Đã xuất báo cáo đối soát ra Sheet mới trong Excel thành công!";
                WpfMessageBox.Show("Đã tạo Sheet báo cáo đối soát dữ liệu Oracle và tô màu nổi bật các ô sai lệch thành công!",
                                   "Thành công", WpfMessageBoxButton.OK, WpfMessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show(ex.Message, "Lỗi xuất báo cáo Excel", WpfMessageBoxButton.OK, WpfMessageBoxImage.Error);
            }
        }

        private void BtnCopySummary_Click(object sender, RoutedEventArgs e)
        {
            if (_lastResult == null)
            {
                WpfMessageBox.Show("Chưa có kết quả để sao chép.", "Thông báo", WpfMessageBoxButton.OK, WpfMessageBoxImage.Information);
                return;
            }

            var sb = new StringBuilder();
            sb.AppendLine("=== BÁO CÁO ĐỐI SOÁT BẢNG ORACLE DATABASE ===");
            sb.AppendLine($"Database A (Gốc): {_lastResult.SchemaA}.{_lastResult.TableA} ({_lastResult.TotalRowsA:N0} dòng)");
            sb.AppendLine($"Database B (Đối chiếu): {_lastResult.SchemaB}.{_lastResult.TableB} ({_lastResult.TotalRowsB:N0} dòng)");
            sb.AppendLine($"Thời gian xử lý: {_lastResult.ExecutionTime.TotalSeconds:F2} giây");
            sb.AppendLine("-------------------------------------------------");
            sb.AppendLine($"✅ Trùng khớp: {_lastResult.MatchCount:N0} dòng");
            sb.AppendLine($"⚠️ Sai lệch giá trị: {_lastResult.ModifiedCount:N0} dòng");
            sb.AppendLine($"➖ Chỉ có ở DB A (Thiếu ở B): {_lastResult.MissingInBCount:N0} dòng");
            sb.AppendLine($"➕ Chỉ có ở DB B (Thiếu ở A): {_lastResult.MissingInACount:N0} dòng");
            sb.AppendLine("=================================================");

            WpfClipboard.SetText(sb.ToString());
            txtStatus.Text = "Đã sao chép tóm tắt kết quả vào Clipboard!";
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        #endregion
    }
}
