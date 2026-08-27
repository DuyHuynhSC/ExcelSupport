using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using ExcelSupport.Helpers;
using ExcelSupport.Models;
using ExcelSupport.Services;
using Microsoft.Office.Interop.Excel;
using WpfMessageBox = System.Windows.MessageBox;
using WpfMessageBoxButton = System.Windows.MessageBoxButton;
using WpfMessageBoxImage = System.Windows.MessageBoxImage;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using ExcelRange = Microsoft.Office.Interop.Excel.Range;
using ExcelWorksheet = Microsoft.Office.Interop.Excel.Worksheet;
using Window = System.Windows.Window;

namespace ExcelSupport.Views
{
    public partial class OracleQuickQueryDialog : Window
    {
        private readonly ObservableCollection<OracleConnectionProfile> _profiles = new ObservableCollection<OracleConnectionProfile>();
        public static readonly DependencyProperty IsDarkThemeProperty =
            DependencyProperty.Register(nameof(IsDarkTheme), typeof(bool), typeof(OracleQuickQueryDialog),
                new PropertyMetadata(false));

        public bool IsDarkTheme
        {
            get => (bool)GetValue(IsDarkThemeProperty);
            set => SetValue(IsDarkThemeProperty, value);
        }

        private bool _isExecuting = false;
        private System.Data.DataTable? _currentDataTable;
        private string _lastExecutedSql = "";
        private OracleConnectionProfile? _lastExecutedProfile;

        private static OracleQuickQueryDialog? _currentInstance;

        public OracleQuickQueryDialog(bool isDarkTheme = false)
        {
            InitializeComponent();
            IsDarkTheme = isDarkTheme;

            LoadProfiles();
            InitTargetLocation();
            LoadQueryHistory();
            RestoreLastQuery();
        }

        public static void ShowWindow(bool? isDarkTheme = null)
        {
            try
            {
                bool isDark = isDarkTheme ?? (AddInEvents.MainViewModel?.IsDarkTheme ?? AiConfigManager.Current.IsDarkTheme);

                if (_currentInstance != null && _currentInstance.IsLoaded)
                {
                    _currentInstance.IsDarkTheme = isDark;
                    _currentInstance.Activate();
                    return;
                }

                _currentInstance = new OracleQuickQueryDialog(isDark);

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

                _currentInstance.ShowDialog();
                _currentInstance = null;
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Không thể mở hộp thoại Truy vấn nhanh:\n{ex.Message}",
                                   "Lỗi giao diện", WpfMessageBoxButton.OK, WpfMessageBoxImage.Error);
            }
        }

        private void LoadProfiles()
        {
            var list = OracleConnectionManager.GetProfiles();
            _profiles.Clear();
            foreach (var p in list)
            {
                _profiles.Add(p);
            }

            cboProfile.ItemsSource = _profiles;

            var defaultProfile = OracleConnectionManager.GetDefaultProfile();
            if (defaultProfile != null)
            {
                var match = _profiles.FirstOrDefault(p => p.Id == defaultProfile.Id);
                cboProfile.SelectedItem = match ?? _profiles.FirstOrDefault();
            }
            else if (_profiles.Count > 0)
            {
                cboProfile.SelectedIndex = 0;
            }
        }

        private void LoadQueryHistory()
        {
            try
            {
                var history = OracleConnectionManager.GetQueryHistory();
                cboQueryHistory.ItemsSource = null;
                cboQueryHistory.ItemsSource = history;
                cboQueryHistory.DisplayMemberPath = "DisplayText";
            }
            catch { }
        }

        private void CboQueryHistory_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cboQueryHistory.SelectedItem is OracleQueryHistoryItem item && !string.IsNullOrWhiteSpace(item.Sql))
            {
                txtSqlQuery.Text = item.Sql;
                txtSqlQuery.Focus();
                txtSqlQuery.CaretIndex = txtSqlQuery.Text.Length;
            }
        }

        private void BtnClearHistory_Click(object sender, RoutedEventArgs e)
        {
            var confirm = WpfMessageBox.Show("Bạn có muốn xóa toàn bộ lịch sử các câu SQL đã truy vấn không?",
                                             "Xác nhận xóa lịch sử", WpfMessageBoxButton.YesNo, WpfMessageBoxImage.Question);
            if (confirm == MessageBoxResult.Yes)
            {
                OracleConnectionManager.ClearQueryHistory();
                LoadQueryHistory();
                txtStatus.Text = "Đã xóa toàn bộ lịch sử truy vấn.";
            }
        }

        private void InitTargetLocation()
        {
            try
            {
                var app = AddInEvents.Instance?.ExcelAppInstance;
                if (app?.ActiveCell is ExcelRange cell)
                {
                    txtTargetLocation.Text = $"{cell.Worksheet.Name}!{cell.Address[false, false]}";
                }
            }
            catch { }
        }

        private void RestoreLastQuery()
        {
            try
            {
                var history = OracleConnectionManager.GetQueryHistory();
                if (history.Count > 0 && !string.IsNullOrWhiteSpace(history[0].Sql))
                {
                    txtSqlQuery.Text = history[0].Sql;
                    return;
                }

                var session = OracleConnectionManager.GetLastSession();
                if (session != null && !string.IsNullOrWhiteSpace(session.TableA))
                {
                    string where = string.IsNullOrWhiteSpace(session.WhereClauseA) ? "" : $" WHERE {session.WhereClauseA}";
                    txtSqlQuery.Text = $"SELECT * FROM {session.TableA}{where}";
                }
                else
                {
                    txtSqlQuery.Text = "SELECT * FROM DUAL";
                }
            }
            catch
            {
                txtSqlQuery.Text = "SELECT * FROM DUAL";
            }
        }

        private void CboProfile_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Update profile
        }

        private void BtnPickLocation_Click(object sender, RoutedEventArgs e)
        {
            var app = AddInEvents.Instance?.ExcelAppInstance;
            if (app == null) return;

            try
            {
                this.Visibility = Visibility.Hidden;

                dynamic excelApp = app;
                dynamic result = excelApp.InputBox(
                    Prompt: LocalizationService.Get("Oracle_PromptPickLocation") ?? "Chọn ô bắt đầu chèn dữ liệu:",
                    Title: LocalizationService.Get("Oracle_TitlePickLocation") ?? "Vị Trí Chèn Dữ Liệu",
                    Default: txtTargetLocation.Text.Trim(),
                    Type: 8);

                this.Visibility = Visibility.Visible;
                this.Activate();

                if (result is ExcelRange targetRange)
                {
                    txtTargetLocation.Text = $"{targetRange.Worksheet.Name}!{targetRange.Address[false, false]}";
                }
            }
            catch
            {
                this.Visibility = Visibility.Visible;
                this.Activate();
            }
        }

        private void BtnRefreshLocation_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var app = AddInEvents.Instance?.ExcelAppInstance;
                if (app?.ActiveCell is ExcelRange cell)
                {
                    txtTargetLocation.Text = $"{cell.Worksheet.Name}!{cell.Address[false, false]}";
                }
            }
            catch { }
        }

        private async void BtnExecute_Click(object sender, RoutedEventArgs e)
        {
            if (_isExecuting) return;

            string sql = txtSqlQuery.Text.Trim();
            if (string.IsNullOrWhiteSpace(sql))
            {
                WpfMessageBox.Show("Vui lòng nhập câu lệnh SQL cần truy vấn.", "Chưa nhập câu lệnh", WpfMessageBoxButton.OK, WpfMessageBoxImage.Warning);
                txtSqlQuery.Focus();
                return;
            }

            var profile = cboProfile.SelectedItem as OracleConnectionProfile;
            if (profile == null)
            {
                WpfMessageBox.Show("Vui lòng chọn hoặc cấu hình Kết nối Oracle trước khi truy vấn.", "Chưa chọn kết nối", WpfMessageBoxButton.OK, WpfMessageBoxImage.Warning);
                return;
            }

            _isExecuting = true;
            btnExecute.IsEnabled = false;
            btnInsert.IsEnabled = false;
            pbProgress.Visibility = Visibility.Visible;
            txtStatus.Text = "Đang kết nối Oracle và thực thi truy vấn...";

            int.TryParse(txtMaxRows.Text, out int maxRows);
            var sw = System.Diagnostics.Stopwatch.StartNew();

            try
            {
                var config = profile.ToConnectionConfig();
                var dt = await OracleQuickQueryService.ExecuteQueryAsync(config, sql, maxRows);
                sw.Stop();

                _currentDataTable = dt;
                _lastExecutedSql = sql;
                _lastExecutedProfile = profile;

                if (dt.Rows.Count == 0 && dt.Columns.Count == 0)
                {
                    dgPreview.ItemsSource = null;
                    txtPreviewPlaceholder.Text = "Câu lệnh thực thi thành công nhưng không trả về dữ liệu bảng.";
                    txtPreviewPlaceholder.Visibility = Visibility.Visible;
                    bdPreviewCount.Visibility = Visibility.Collapsed;
                    txtStatus.Text = "Câu lệnh thực thi thành công nhưng không có dữ liệu trả về.";
                    btnInsert.IsEnabled = false;
                    return;
                }

                // Bind to Preview DataGrid
                dgPreview.ItemsSource = _currentDataTable.DefaultView;
                txtPreviewPlaceholder.Visibility = Visibility.Collapsed;

                txtPreviewCount.Text = $"{dt.Rows.Count:N0} dòng, {dt.Columns.Count:N0} cột ({sw.Elapsed.TotalSeconds:F2}s)";
                bdPreviewCount.Visibility = Visibility.Visible;
                btnInsert.IsEnabled = true;

                // Save to history & reload history dropdown
                try
                {
                    OracleConnectionManager.AddQueryHistory(sql, dt.Rows.Count, profile.Name);
                    LoadQueryHistory();
                }
                catch { }

                txtStatus.Text = $"Đã tải {dt.Rows.Count:N0} dòng, {dt.Columns.Count:N0} cột ({sw.Elapsed.TotalSeconds:F2}s). Hãy kiểm tra dữ liệu xem trước và nhấn 'Chèn Vào Excel'.";
            }
            catch (Exception ex)
            {
                sw.Stop();
                txtStatus.Text = $"Lỗi thực thi: {ex.Message}";
                WpfMessageBox.Show($"Quá trình truy vấn phát sinh lỗi:\n\n{ex.Message}", "Lỗi Truy Vấn", WpfMessageBoxButton.OK, WpfMessageBoxImage.Error);
            }
            finally
            {
                _isExecuting = false;
                btnExecute.IsEnabled = true;
                pbProgress.Visibility = Visibility.Collapsed;
            }
        }

        private void BtnInsert_Click(object sender, RoutedEventArgs e)
        {
            if (_currentDataTable == null || (_currentDataTable.Rows.Count == 0 && _currentDataTable.Columns.Count == 0))
            {
                WpfMessageBox.Show("Chưa có dữ liệu để chèn. Vui lòng nhấn 'Thực Thi' để tải dữ liệu trước.", "Chưa có dữ liệu", WpfMessageBoxButton.OK, WpfMessageBoxImage.Warning);
                return;
            }

            var app = AddInEvents.Instance?.ExcelAppInstance;
            if (app == null)
            {
                WpfMessageBox.Show("Không thể kết nối với ứng dụng Excel.", "Lỗi Excel", WpfMessageBoxButton.OK, WpfMessageBoxImage.Error);
                return;
            }

            // Resolve Target Range
            ExcelWorksheet? targetWs = null;
            int startRow = 1;
            int startCol = 1;

            string loc = txtTargetLocation.Text.Trim();
            if (!string.IsNullOrWhiteSpace(loc))
            {
                try
                {
                    if (loc.Contains("!"))
                    {
                        var parts = loc.Split('!');
                        string wsName = parts[0].Replace("'", "");
                        string addr = parts[1];
                        targetWs = app.Worksheets[wsName];
                        ExcelRange rng = targetWs.Range[addr];
                        startRow = rng.Row;
                        startCol = rng.Column;
                    }
                    else
                    {
                        targetWs = app.ActiveSheet as ExcelWorksheet;
                        if (targetWs != null)
                        {
                            ExcelRange rng = targetWs.Range[loc];
                            startRow = rng.Row;
                            startCol = rng.Column;
                        }
                    }
                }
                catch
                {
                    targetWs = app.ActiveSheet as ExcelWorksheet;
                    if (app.ActiveCell != null)
                    {
                        startRow = app.ActiveCell.Row;
                        startCol = app.ActiveCell.Column;
                    }
                }
            }

            targetWs ??= app.ActiveSheet as ExcelWorksheet;
            if (targetWs == null)
            {
                WpfMessageBox.Show("Không tìm thấy Sheet làm việc để chèn dữ liệu.", "Lỗi Sheet", WpfMessageBoxButton.OK, WpfMessageBoxImage.Error);
                return;
            }

            int.TryParse(txtMaxRows.Text, out int maxRows);
            string headerColorHex = (cboHeaderColor.SelectedItem as ComboBoxItem)?.Tag?.ToString() ?? "#CCFFFF";

            var options = new OracleQuickQueryOptions
            {
                IncludeTitle = chkIncludeTitle.IsChecked == true,
                TitleColorHex = "#2563EB",
                IncludeHeaders = chkIncludeHeaders.IsChecked == true,
                HeaderBgColorHex = headerColorHex,
                AutoFitColumns = chkAutoFit.IsChecked == true,
                MaxRows = maxRows
            };

            var sw = System.Diagnostics.Stopwatch.StartNew();

            try
            {
                targetWs.Activate();

                // Extract Table Name for Title
                string sql = !string.IsNullOrWhiteSpace(_lastExecutedSql) ? _lastExecutedSql : txtSqlQuery.Text.Trim();
                string tableName = OracleQuickQueryService.ExtractTableName(sql);

                // Insert to Excel
                app.ScreenUpdating = false;
                var (rows, cols) = OracleQuickQueryService.InsertDataToWorksheet(targetWs, startRow, startCol, _currentDataTable, tableName, options);
                app.ScreenUpdating = true;
                sw.Stop();

                txtStatus.Text = $"Đã chèn thành công {rows:N0} dòng, {cols:N0} cột vào Sheet '{targetWs.Name}' ({sw.Elapsed.TotalSeconds:F2}s).";
                WpfMessageBox.Show($"Đã chèn thành công {rows:N0} dòng, {cols:N0} cột dữ liệu vào Sheet '{targetWs.Name}'!\nThời gian chèn: {sw.Elapsed.TotalSeconds:F2} giây.",
                                   "Chèn Dữ Liệu Thành Công", WpfMessageBoxButton.OK, WpfMessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                sw.Stop();
                try { app.ScreenUpdating = true; } catch { }
                txtStatus.Text = $"Lỗi chèn Excel: {ex.Message}";
                WpfMessageBox.Show($"Quá trình chèn dữ liệu vào Excel phát sinh lỗi:\n\n{ex.Message}", "Lỗi Chèn Excel", WpfMessageBoxButton.OK, WpfMessageBoxImage.Error);
            }
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void Window_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.Enter && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                e.Handled = true;
                BtnExecute_Click(btnExecute, new RoutedEventArgs());
            }
            else if (e.Key == Key.F5)
            {
                e.Handled = true;
                if (btnInsert.IsEnabled)
                {
                    BtnInsert_Click(btnInsert, new RoutedEventArgs());
                }
                else
                {
                    BtnExecute_Click(btnExecute, new RoutedEventArgs());
                }
            }
            else if (e.Key == Key.Escape)
            {
                Close();
            }
        }
    }
}
