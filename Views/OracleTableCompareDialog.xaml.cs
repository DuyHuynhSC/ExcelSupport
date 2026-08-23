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
using WpfMessageBoxResult = System.Windows.MessageBoxResult;
using MediaColor = System.Windows.Media.Color;
using WpfBinding = System.Windows.Data.Binding;
using WpfClipboard = System.Windows.Clipboard;

namespace ExcelSupport.Views
{
    public partial class OracleTableCompareDialog : Window
    {
        private readonly ObservableCollection<OracleTableColumnInfo> _tableColumns = new ObservableCollection<OracleTableColumnInfo>();
        private readonly ObservableCollection<OracleConnectionProfile> _profiles = new ObservableCollection<OracleConnectionProfile>();
        private OracleCompareResult? _lastResult;
        private List<OracleRowDiffItem> _allDiffItems = new List<OracleRowDiffItem>();
        private bool _isComparing = false;
        private readonly bool _isDarkTheme;

        // Cached active configs
        private OracleConnectionConfig? _activeConfigA;
        private OracleConnectionConfig? _activeConfigB;

        public OracleTableCompareDialog(bool isDarkTheme = false)
        {
            InitializeComponent();
            _isDarkTheme = isDarkTheme;

            icKeyColumns.ItemsSource = _tableColumns;

            // Load saved connection profiles
            RefreshProfilesList();

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

        #region Profile Management

        private void RefreshProfilesList()
        {
            var list = OracleConnectionManager.GetProfiles();
            _profiles.Clear();
            foreach (var p in list)
            {
                _profiles.Add(p);
            }

            cboProfileA.ItemsSource = null;
            cboProfileA.ItemsSource = _profiles;

            cboProfileB.ItemsSource = null;
            cboProfileB.ItemsSource = _profiles;

            lstProfiles.ItemsSource = null;
            lstProfiles.ItemsSource = _profiles;

            if (_profiles.Count > 0)
            {
                if (cboProfileA.SelectedIndex < 0) cboProfileA.SelectedIndex = 0;
                if (cboProfileB.SelectedIndex < 0) cboProfileB.SelectedIndex = _profiles.Count > 1 ? 1 : 0;
                if (lstProfiles.SelectedIndex < 0) lstProfiles.SelectedIndex = 0;
            }
        }

        private void BtnManageProfiles_Click(object sender, RoutedEventArgs e)
        {
            mainTabControl.SelectedIndex = 1; // Switch to Tab Settings
        }

        private void CboProfileA_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!IsLoaded) return;
            if (cboProfileA.SelectedItem is OracleConnectionProfile p)
            {
                _activeConfigA = p.ToConnectionConfig();
                txtStatusBadgeA.Text = "⚪ Sẵn sàng kết nối";
                txtStatusBadgeA.Foreground = new SolidColorBrush(MediaColor.FromRgb(100, 116, 139));
                if (!string.IsNullOrWhiteSpace(p.DefaultWhereClause) && string.IsNullOrWhiteSpace(txtWhereA.Text))
                {
                    txtWhereA.Text = p.DefaultWhereClause;
                }
            }
        }

        private void CboProfileB_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!IsLoaded) return;
            if (cboProfileB.SelectedItem is OracleConnectionProfile p)
            {
                _activeConfigB = p.ToConnectionConfig();
                txtStatusBadgeB.Text = "⚪ Sẵn sàng kết nối";
                txtStatusBadgeB.Foreground = new SolidColorBrush(MediaColor.FromRgb(100, 116, 139));
                if (!string.IsNullOrWhiteSpace(p.DefaultWhereClause) && string.IsNullOrWhiteSpace(txtWhereB.Text))
                {
                    txtWhereB.Text = p.DefaultWhereClause;
                }
            }
        }

        private void LstProfiles_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (lstProfiles.SelectedItem is OracleConnectionProfile p)
            {
                txtSettingProfileName.Text = p.Name;
                txtSettingHost.Text = p.Host;
                txtSettingPort.Text = p.Port.ToString();
                txtSettingService.Text = p.ServiceNameOrSid;
                rbSettingService.IsChecked = p.ServiceType == OracleServiceNameType.ServiceName;
                rbSettingSid.IsChecked = p.ServiceType == OracleServiceNameType.SID;
                txtSettingUser.Text = p.Username;
                txtSettingPass.Password = p.Password;
                txtSettingDefaultSchema.Text = p.DefaultSchema;
                txtSettingDefaultTable.Text = p.DefaultTable;
                txtSettingDefaultWhere.Text = p.DefaultWhereClause;
                txtSettingStatus.Text = string.Empty;
            }
        }

        private void BtnNewProfile_Click(object sender, RoutedEventArgs e)
        {
            var newProfile = new OracleConnectionProfile
            {
                Name = $"Connection {_profiles.Count + 1}",
                Host = "localhost",
                Port = 1521,
                ServiceNameOrSid = "ORCL",
                ServiceType = OracleServiceNameType.ServiceName,
                Username = "",
                Password = ""
            };

            OracleConnectionManager.AddOrUpdateProfile(newProfile);
            RefreshProfilesList();
            lstProfiles.SelectedItem = _profiles.FirstOrDefault(p => p.Id == newProfile.Id);
        }

        private void BtnCloneProfile_Click(object sender, RoutedEventArgs e)
        {
            if (lstProfiles.SelectedItem is OracleConnectionProfile p)
            {
                var clone = p.Clone();
                OracleConnectionManager.AddOrUpdateProfile(clone);
                RefreshProfilesList();
                lstProfiles.SelectedItem = _profiles.FirstOrDefault(x => x.Id == clone.Id);
            }
        }

        private void BtnDeleteProfile_Click(object sender, RoutedEventArgs e)
        {
            if (lstProfiles.SelectedItem is OracleConnectionProfile p)
            {
                var confirm = WpfMessageBox.Show(
                    $"Bạn có chắc chắn muốn xóa cấu hình kết nối '{p.Name}' không?",
                    "Xác nhận xóa", WpfMessageBoxButton.YesNo, WpfMessageBoxImage.Question);

                if (confirm == WpfMessageBoxResult.Yes)
                {
                    OracleConnectionManager.DeleteProfile(p.Id);
                    RefreshProfilesList();
                }
            }
        }

        private void BtnSettingSave_Click(object sender, RoutedEventArgs e)
        {
            if (lstProfiles.SelectedItem is OracleConnectionProfile p)
            {
                int.TryParse(txtSettingPort.Text, out int port);
                if (port <= 0) port = 1521;

                p.Name = string.IsNullOrWhiteSpace(txtSettingProfileName.Text) ? "Connection" : txtSettingProfileName.Text.Trim();
                p.Host = txtSettingHost.Text.Trim();
                p.Port = port;
                p.ServiceNameOrSid = txtSettingService.Text.Trim();
                p.ServiceType = (rbSettingSid.IsChecked == true) ? OracleServiceNameType.SID : OracleServiceNameType.ServiceName;
                p.Username = txtSettingUser.Text.Trim();
                p.Password = txtSettingPass.Password;
                p.DefaultSchema = txtSettingDefaultSchema.Text.Trim();
                p.DefaultTable = txtSettingDefaultTable.Text.Trim();
                p.DefaultWhereClause = txtSettingDefaultWhere.Text.Trim();

                OracleConnectionManager.AddOrUpdateProfile(p);
                RefreshProfilesList();

                txtSettingStatus.Text = "✅ Đã lưu cấu hình thành công!";
                txtSettingStatus.Foreground = new SolidColorBrush(MediaColor.FromRgb(22, 163, 74));
            }
        }

        private async void BtnSettingTest_Click(object sender, RoutedEventArgs e)
        {
            btnSettingTest.IsEnabled = false;
            txtSettingStatus.Text = "⏳ Đang kiểm tra kết nối...";
            txtSettingStatus.Foreground = new SolidColorBrush(MediaColor.FromRgb(217, 119, 6));

            int.TryParse(txtSettingPort.Text, out int port);
            if (port <= 0) port = 1521;

            var config = new OracleConnectionConfig
            {
                Host = txtSettingHost.Text.Trim(),
                Port = port,
                ServiceNameOrSid = txtSettingService.Text.Trim(),
                ServiceType = (rbSettingSid.IsChecked == true) ? OracleServiceNameType.SID : OracleServiceNameType.ServiceName,
                Username = txtSettingUser.Text.Trim(),
                Password = txtSettingPass.Password
            };

            var (success, msg, version) = await OracleDataCompareService.TestConnectionAsync(config);
            btnSettingTest.IsEnabled = true;

            if (success)
            {
                txtSettingStatus.Text = $"🟢 Kết nối thành công! Phiên bản: {version}";
                txtSettingStatus.Foreground = new SolidColorBrush(MediaColor.FromRgb(22, 163, 74));
            }
            else
            {
                txtSettingStatus.Text = $"🔴 {msg}";
                txtSettingStatus.Foreground = new SolidColorBrush(MediaColor.FromRgb(220, 38, 38));
            }
        }

        #endregion

        #region Connection Testing & Metadata Loading

        private async void BtnConnectA_Click(object sender, RoutedEventArgs e)
        {
            if (cboProfileA.SelectedItem is not OracleConnectionProfile p)
            {
                WpfMessageBox.Show("Vui lòng chọn một cấu hình kết nối cho Database A.", "Thông báo", WpfMessageBoxButton.OK, WpfMessageBoxImage.Warning);
                return;
            }

            btnConnectA.IsEnabled = false;
            txtStatusBadgeA.Text = "⏳ Đang kết nối...";
            txtStatusBadgeA.Foreground = new SolidColorBrush(MediaColor.FromRgb(217, 119, 6));

            _activeConfigA = p.ToConnectionConfig();
            var (success, msg, version) = await OracleDataCompareService.TestConnectionAsync(_activeConfigA);

            btnConnectA.IsEnabled = true;
            if (success)
            {
                txtStatusBadgeA.Text = $"🟢 {version}";
                txtStatusBadgeA.Foreground = new SolidColorBrush(MediaColor.FromRgb(22, 163, 74));
                txtStatus.Text = $"DB A kết nối thành công: {version}";

                await LoadSchemasAAsync(_activeConfigA, p.DefaultSchema, p.DefaultTable);
            }
            else
            {
                txtStatusBadgeA.Text = "🔴 Thất bại";
                txtStatusBadgeA.Foreground = new SolidColorBrush(MediaColor.FromRgb(220, 38, 38));
                txtStatus.Text = msg;
                WpfMessageBox.Show(msg, "Lỗi kết nối Database A", WpfMessageBoxButton.OK, WpfMessageBoxImage.Error);
            }
        }

        private async void BtnConnectB_Click(object sender, RoutedEventArgs e)
        {
            if (cboProfileB.SelectedItem is not OracleConnectionProfile p)
            {
                WpfMessageBox.Show("Vui lòng chọn một cấu hình kết nối cho Database B.", "Thông báo", WpfMessageBoxButton.OK, WpfMessageBoxImage.Warning);
                return;
            }

            btnConnectB.IsEnabled = false;
            txtStatusBadgeB.Text = "⏳ Đang kết nối...";
            txtStatusBadgeB.Foreground = new SolidColorBrush(MediaColor.FromRgb(217, 119, 6));

            _activeConfigB = p.ToConnectionConfig();
            var (success, msg, version) = await OracleDataCompareService.TestConnectionAsync(_activeConfigB);

            btnConnectB.IsEnabled = true;
            if (success)
            {
                txtStatusBadgeB.Text = $"🟢 {version}";
                txtStatusBadgeB.Foreground = new SolidColorBrush(MediaColor.FromRgb(22, 163, 74));
                txtStatus.Text = $"DB B kết nối thành công: {version}";

                await LoadSchemasBAsync(_activeConfigB, p.DefaultSchema, p.DefaultTable);
            }
            else
            {
                txtStatusBadgeB.Text = "🔴 Thất bại";
                txtStatusBadgeB.Foreground = new SolidColorBrush(MediaColor.FromRgb(220, 38, 38));
                txtStatus.Text = msg;
                WpfMessageBox.Show(msg, "Lỗi kết nối Database B", WpfMessageBoxButton.OK, WpfMessageBoxImage.Error);
            }
        }

        private async Task LoadSchemasAAsync(OracleConnectionConfig config, string? defaultSchema = null, string? defaultTable = null)
        {
            try
            {
                txtStatus.Text = "Đang tải danh sách Schema từ Database A...";
                var schemas = await OracleDataCompareService.GetSchemasAsync(config);
                cboSchemaA.ItemsSource = schemas;

                string userUpper = !string.IsNullOrWhiteSpace(defaultSchema) ? defaultSchema.ToUpperInvariant() : config.Username.ToUpperInvariant();
                var matched = schemas.FirstOrDefault(s => s.Equals(userUpper, StringComparison.OrdinalIgnoreCase));
                if (matched != null)
                {
                    cboSchemaA.SelectedItem = matched;
                }
                else if (schemas.Count > 0)
                {
                    cboSchemaA.SelectedIndex = 0;
                }

                if (!string.IsNullOrWhiteSpace(defaultTable))
                {
                    cboTableA.Text = defaultTable;
                }
            }
            catch (Exception ex)
            {
                txtStatus.Text = $"Lỗi tải schema: {ex.Message}";
            }
        }

        private async Task LoadSchemasBAsync(OracleConnectionConfig config, string? defaultSchema = null, string? defaultTable = null)
        {
            try
            {
                txtStatus.Text = "Đang tải danh sách Schema từ Database B...";
                var schemas = await OracleDataCompareService.GetSchemasAsync(config);
                cboSchemaB.ItemsSource = schemas;

                string userUpper = !string.IsNullOrWhiteSpace(defaultSchema) ? defaultSchema.ToUpperInvariant() : config.Username.ToUpperInvariant();
                var matched = schemas.FirstOrDefault(s => s.Equals(userUpper, StringComparison.OrdinalIgnoreCase));
                if (matched != null)
                {
                    cboSchemaB.SelectedItem = matched;
                }
                else if (schemas.Count > 0)
                {
                    cboSchemaB.SelectedIndex = 0;
                }

                if (!string.IsNullOrWhiteSpace(defaultTable))
                {
                    cboTableB.Text = defaultTable;
                }
            }
            catch (Exception ex)
            {
                txtStatus.Text = $"Lỗi tải schema: {ex.Message}";
            }
        }

        private async void CboSchemaA_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!IsLoaded) return;
            string schema = cboSchemaA.SelectedItem?.ToString() ?? cboSchemaA.Text;
            if (string.IsNullOrWhiteSpace(schema) || _activeConfigA == null) return;

            try
            {
                txtStatus.Text = $"Đang tải danh sách bảng từ Schema [{schema}] (DB A)...";
                var tables = await OracleDataCompareService.GetTablesAndViewsAsync(_activeConfigA, schema);
                cboTableA.ItemsSource = tables;
                if (tables.Count > 0)
                {
                    cboTableA.SelectedIndex = 0;
                }
                txtStatus.Text = $"Đã tải {tables.Count} bảng/view từ DB A.";
            }
            catch (Exception ex)
            {
                txtStatus.Text = $"Lỗi tải bảng DB A: {ex.Message}";
            }
        }

        private async void CboSchemaB_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!IsLoaded) return;
            string schema = cboSchemaB.SelectedItem?.ToString() ?? cboSchemaB.Text;
            if (string.IsNullOrWhiteSpace(schema) || _activeConfigB == null) return;

            try
            {
                txtStatus.Text = $"Đang tải danh sách bảng từ Schema [{schema}] (DB B)...";
                var tables = await OracleDataCompareService.GetTablesAndViewsAsync(_activeConfigB, schema);
                cboTableB.ItemsSource = tables;

                // Sync with table A selection if matched
                string tableA = cboTableA.SelectedItem?.ToString() ?? cboTableA.Text;
                var matchedTable = tables.FirstOrDefault(t => t.Equals(tableA, StringComparison.OrdinalIgnoreCase));
                if (matchedTable != null)
                {
                    cboTableB.SelectedItem = matchedTable;
                }
                else if (tables.Count > 0)
                {
                    cboTableB.SelectedIndex = 0;
                }
                txtStatus.Text = $"Đã tải {tables.Count} bảng/view từ DB B.";
            }
            catch (Exception ex)
            {
                txtStatus.Text = $"Lỗi tải bảng DB B: {ex.Message}";
            }
        }

        private async void CboTableA_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!IsLoaded) return;
            string table = cboTableA.SelectedItem?.ToString() ?? cboTableA.Text;
            string schema = cboSchemaA.SelectedItem?.ToString() ?? cboSchemaA.Text;
            if (string.IsNullOrWhiteSpace(table) || _activeConfigA == null) return;

            // Auto-select same table in DB B if available
            if (cboTableB.ItemsSource is List<string> bTables)
            {
                var matched = bTables.FirstOrDefault(t => t.Equals(table, StringComparison.OrdinalIgnoreCase));
                if (matched != null) cboTableB.SelectedItem = matched;
            }

            try
            {
                txtStatus.Text = $"Đang lấy thông tin cột của bảng [{table}]...";
                var cols = await OracleDataCompareService.GetTableColumnsAsync(_activeConfigA, schema, table);
                _tableColumns.Clear();
                foreach (var c in cols)
                {
                    _tableColumns.Add(c);
                }
                txtStatus.Text = $"Bảng [{table}] có {cols.Count} cột. (Đã tự động chọn Primary Key).";
            }
            catch (Exception ex)
            {
                txtStatus.Text = $"Lỗi lấy cấu trúc cột: {ex.Message}";
            }
        }

        #endregion

        #region Start Comparison

        private async void BtnStartCompare_Click(object sender, RoutedEventArgs e)
        {
            if (_isComparing) return;

            // Switch to Compare tab if on settings tab
            mainTabControl.SelectedIndex = 0;

            if (_activeConfigA == null)
            {
                if (cboProfileA.SelectedItem is OracleConnectionProfile pA) _activeConfigA = pA.ToConnectionConfig();
            }
            if (_activeConfigB == null)
            {
                if (cboProfileB.SelectedItem is OracleConnectionProfile pB) _activeConfigB = pB.ToConnectionConfig();
            }

            if (_activeConfigA == null || _activeConfigB == null)
            {
                WpfMessageBox.Show("Vui lòng chọn cấu hình và kết nối thành công tới cả 2 Database.", "Cảnh báo", WpfMessageBoxButton.OK, WpfMessageBoxImage.Warning);
                return;
            }

            string tableA = cboTableA.SelectedItem?.ToString() ?? cboTableA.Text.Trim();
            string tableB = cboTableB.SelectedItem?.ToString() ?? cboTableB.Text.Trim();

            if (string.IsNullOrWhiteSpace(tableA) || string.IsNullOrWhiteSpace(tableB))
            {
                WpfMessageBox.Show("Vui lòng chọn hoặc nhập tên Bảng cần so sánh cho cả 2 Database.", "Cảnh báo", WpfMessageBoxButton.OK, WpfMessageBoxImage.Warning);
                return;
            }

            string schemaA = cboSchemaA.SelectedItem?.ToString() ?? cboSchemaA.Text.Trim();
            string schemaB = cboSchemaB.SelectedItem?.ToString() ?? cboSchemaB.Text.Trim();

            int.TryParse(txtMaxRows.Text, out int maxRows);

            var options = new OracleCompareOptions
            {
                Mode = (rbModeSeq.IsChecked == true) ? OracleCompareMode.Sequential : OracleCompareMode.ByKeyColumns,
                IgnoreWhitespace = chkIgnoreSpaces.IsChecked == true,
                IgnoreCase = chkIgnoreCase.IsChecked == true,
                MaxRows = maxRows,
                WhereClauseA = txtWhereA.Text.Trim(),
                WhereClauseB = txtWhereB.Text.Trim()
            };

            if (options.Mode == OracleCompareMode.ByKeyColumns)
            {
                options.SelectedKeyColumns = _tableColumns.Where(c => c.IsSelectedKey).Select(c => c.ColumnName).ToList();
                if (options.SelectedKeyColumns.Count == 0)
                {
                    var confirm = WpfMessageBox.Show(
                        "Chưa có Cột Khóa chính nào được chọn để so khớp bản ghi!\n\nBạn có muốn chuyển sang so sánh theo Thứ Tự Dòng (Sequential) không?",
                        "Chưa chọn Khóa", WpfMessageBoxButton.YesNo, WpfMessageBoxImage.Question);

                    if (confirm == WpfMessageBoxResult.Yes)
                    {
                        rbModeSeq.IsChecked = true;
                        options.Mode = OracleCompareMode.Sequential;
                    }
                    else
                    {
                        return;
                    }
                }
            }

            _isComparing = true;
            btnStartCompare.IsEnabled = false;
            pbProgress.Visibility = Visibility.Visible;
            pbProgress.IsIndeterminate = false;
            pbProgress.Value = 0;
            txtStatus.Text = "Đang trích xuất dữ liệu từ cả 2 Database và tiến hành so sánh...";

            var progress = new Progress<(string StatusText, double ProgressPercent)>(p =>
            {
                txtStatus.Text = p.StatusText;
                if (p.ProgressPercent > 0)
                {
                    pbProgress.Value = p.ProgressPercent;
                }
            });

            try
            {
                var result = await OracleDataCompareService.CompareTablesAsync(
                    _activeConfigA,
                    _activeConfigB,
                    schemaA, tableA,
                    schemaB, tableB,
                    options, progress);

                _lastResult = result;
                _allDiffItems = result.DiffItems;

                // Build Dynamic Columns in DataGrid
                BuildDataGridColumns(result.Columns);

                // Apply initial filter
                ApplyFilter();

                // Update summary badges
                txtStatTotal.Text = $"Tổng: {result.DiffItems.Count:N0} dòng";
                txtStatDiff.Text = $"Sai lệch: {result.ModifiedCount:N0}";
                txtStatMatch.Text = $"Khớp: {result.MatchCount:N0}";

                txtStatus.Text = $"So sánh hoàn tất sau {result.ExecutionTime.TotalSeconds:N1}s. Tìm thấy {result.ModifiedCount} dòng sai lệch.";

                if (result.ModifiedCount == 0 && result.MissingInACount == 0 && result.MissingInBCount == 0)
                {
                    WpfMessageBox.Show("Tuyệt vời! Dữ liệu của 2 bảng hoàn toàn khớp 100% không có bất kỳ sai lệch nào.",
                                       "Khớp Hoàn Hảo", WpfMessageBoxButton.OK, WpfMessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                txtStatus.Text = $"Lỗi so sánh: {ex.Message}";
                WpfMessageBox.Show($"Quá trình đối soát phát sinh lỗi:\n{ex.Message}", "Lỗi So Sánh", WpfMessageBoxButton.OK, WpfMessageBoxImage.Error);
            }
            finally
            {
                _isComparing = false;
                btnStartCompare.IsEnabled = true;
                pbProgress.Visibility = Visibility.Collapsed;
                pbProgress.IsIndeterminate = false;
            }
        }

        private void BuildDataGridColumns(List<string> columns)
        {
            // Retain the first 4 fixed columns (#, Key, Status, Summary)
            while (dgDiffResults.Columns.Count > 4)
            {
                dgDiffResults.Columns.RemoveAt(dgDiffResults.Columns.Count - 1);
            }

            foreach (var col in columns)
            {
                var colA = new DataGridTextColumn
                {
                    Header = $"{col} (A)",
                    Binding = new WpfBinding($"RowValuesA[{col}]"),
                    Width = new DataGridLength(110)
                };

                var colB = new DataGridTextColumn
                {
                    Header = $"{col} (B)",
                    Binding = new WpfBinding($"RowValuesB[{col}]"),
                    Width = new DataGridLength(110)
                };

                dgDiffResults.Columns.Add(colA);
                dgDiffResults.Columns.Add(colB);
            }
        }

        #endregion

        #region Filter & Search Results

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
            if (_allDiffItems == null || _allDiffItems.Count == 0) return;

            IEnumerable<OracleRowDiffItem> filtered = _allDiffItems;

            if (rbFilterDiff.IsChecked == true)
            {
                filtered = filtered.Where(i => i.Status == OracleRowDiffStatus.Modified);
            }
            else if (rbFilterMissingB.IsChecked == true)
            {
                filtered = filtered.Where(i => i.Status == OracleRowDiffStatus.MissingInB);
            }
            else if (rbFilterMissingA.IsChecked == true)
            {
                filtered = filtered.Where(i => i.Status == OracleRowDiffStatus.MissingInA);
            }
            else if (rbFilterMatch.IsChecked == true)
            {
                filtered = filtered.Where(i => i.Status == OracleRowDiffStatus.Identical);
            }

            string search = txtSearchFilter.Text.Trim();
            if (!string.IsNullOrWhiteSpace(search))
            {
                filtered = filtered.Where(i =>
                    (i.KeyDisplay != null && i.KeyDisplay.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0) ||
                    (i.DifferingColumnsSummary != null && i.DifferingColumnsSummary.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0) ||
                    i.CellDiffs.Any(c =>
                        (c.ValueADisplay != null && c.ValueADisplay.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0) ||
                        (c.ValueBDisplay != null && c.ValueBDisplay.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0)));
            }

            dgDiffResults.ItemsSource = filtered.ToList();
        }

        private void RbMode_Checked(object sender, RoutedEventArgs e)
        {
            if (!IsLoaded) return;
            if (icKeyColumns != null)
            {
                icKeyColumns.Visibility = (rbModeKey.IsChecked == true) ? Visibility.Visible : Visibility.Collapsed;
            }
        }

        #endregion

        #region Export Actions

        private void BtnInsertActiveSelection_Click(object sender, RoutedEventArgs e)
        {
            if (_lastResult == null)
            {
                WpfMessageBox.Show("Chưa có kết quả so sánh để chèn vào Excel. Vui lòng nhấn 'Tiến Hành So Sánh' trước.",
                                   "Chưa có dữ liệu", WpfMessageBoxButton.OK, WpfMessageBoxImage.Warning);
                return;
            }

            try
            {
                var app = AddInEvents.Instance?.ExcelAppInstance;
                if (app == null)
                {
                    WpfMessageBox.Show("Không thể kết nối với ứng dụng Excel.", "Lỗi", WpfMessageBoxButton.OK, WpfMessageBoxImage.Error);
                    return;
                }

                txtStatus.Text = "Đang chèn dữ liệu so sánh vào vùng chọn hiện tại...";
                OracleDataCompareService.InsertDiffToActiveSelection(_lastResult, app);

                txtStatus.Text = "Đã chèn và tô màu sai lệch vào vị trí đang chọn thành công!";
                WpfMessageBox.Show("Dữ liệu so sánh đã được ghi thẳng vào vùng ô đang chọn và tô màu sai lệch thành công!",
                                   "Hoàn tất chèn dữ liệu", WpfMessageBoxButton.OK, WpfMessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Lỗi chèn dữ liệu: {ex.Message}", "Lỗi Excel Interop", WpfMessageBoxButton.OK, WpfMessageBoxImage.Error);
            }
        }

        private void BtnExportExcel_Click(object sender, RoutedEventArgs e)
        {
            if (_lastResult == null)
            {
                WpfMessageBox.Show("Chưa có kết quả so sánh để xuất Excel. Vui lòng nhấn 'Tiến Hành So Sánh' trước.",
                                   "Chưa có dữ liệu", WpfMessageBoxButton.OK, WpfMessageBoxImage.Warning);
                return;
            }

            try
            {
                var app = AddInEvents.Instance?.ExcelAppInstance;
                if (app == null)
                {
                    WpfMessageBox.Show("Không thể kết nối với ứng dụng Excel.", "Lỗi", WpfMessageBoxButton.OK, WpfMessageBoxImage.Error);
                    return;
                }

                txtStatus.Text = "Đang tạo Sheet báo cáo và tô màu sai lệch trên Excel...";
                OracleDataCompareService.ExportDiffReportToExcel(_lastResult, app);

                txtStatus.Text = "Xuất báo cáo Excel thành công!";
                WpfMessageBox.Show("Báo cáo so sánh dữ liệu Oracle đã được tạo thành công trên Sheet mới với đầy đủ định dạng màu sắc!",
                                   "Xuất Báo Cáo Thành Công", WpfMessageBoxButton.OK, WpfMessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Lỗi xuất Excel: {ex.Message}", "Lỗi Excel Interop", WpfMessageBoxButton.OK, WpfMessageBoxImage.Error);
            }
        }

        private void BtnCopySummary_Click(object sender, RoutedEventArgs e)
        {
            if (_lastResult == null)
            {
                WpfMessageBox.Show("Chưa có kết quả so sánh để sao chép.", "Chưa có dữ liệu", WpfMessageBoxButton.OK, WpfMessageBoxImage.Warning);
                return;
            }

            var sb = new StringBuilder();
            sb.AppendLine("=== KẾT QUẢ ĐỐI SOÁT DỮ LIỆU ORACLE ===");
            sb.AppendLine($"Database A: {_activeConfigA?.Host}/{_activeConfigA?.ServiceNameOrSid} - Bảng: {cboTableA.Text}");
            sb.AppendLine($"Database B: {_activeConfigB?.Host}/{_activeConfigB?.ServiceNameOrSid} - Bảng: {cboTableB.Text}");
            sb.AppendLine($"Tổng số dòng đối soát: {_lastResult.DiffItems.Count:N0}");
            sb.AppendLine($"Dòng khớp hoàn toàn: {_lastResult.MatchCount:N0}");
            sb.AppendLine($"Dòng có sai lệch giá trị: {_lastResult.ModifiedCount:N0}");
            sb.AppendLine($"Dòng chỉ có ở A: {_lastResult.MissingInBCount:N0}");
            sb.AppendLine($"Dòng chỉ có ở B: {_lastResult.MissingInACount:N0}");
            sb.AppendLine($"Thời gian thực thi: {_lastResult.ExecutionTime.TotalSeconds:N2} giây");

            try
            {
                WpfClipboard.SetText(sb.ToString());
                txtStatus.Text = "Đã sao chép tóm tắt kết quả vào Clipboard!";
            }
            catch (Exception ex)
            {
                txtStatus.Text = $"Lỗi sao chép: {ex.Message}";
            }
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        #endregion
    }
}
