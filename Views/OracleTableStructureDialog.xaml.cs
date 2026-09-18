using System;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using ExcelSupport.Helpers;
using ExcelSupport.Models;
using ExcelSupport.Services;
using Microsoft.Office.Interop.Excel;
using WpfMessageBox = System.Windows.MessageBox;
using WpfMessageBoxButton = System.Windows.MessageBoxButton;
using WpfMessageBoxImage = System.Windows.MessageBoxImage;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using ExcelWorksheet = Microsoft.Office.Interop.Excel.Worksheet;
using Window = System.Windows.Window;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;

namespace ExcelSupport.Views
{
    public partial class OracleTableStructureDialog : Window
    {
        public static readonly DependencyProperty IsDarkThemeProperty =
            DependencyProperty.Register(nameof(IsDarkTheme), typeof(bool), typeof(OracleTableStructureDialog),
                new PropertyMetadata(false));

        public bool IsDarkTheme
        {
            get => (bool)GetValue(IsDarkThemeProperty);
            set => SetValue(IsDarkThemeProperty, value);
        }

        private readonly OracleConnectionConfig _config;
        private readonly Action<string>? _onInsertSqlCallback;
        private OracleTableStructureResult? _currentStructure;
        private bool _isLoading = false;

        public OracleTableStructureDialog(
            OracleConnectionConfig config,
            string initialTableName,
            Action<string>? onInsertSqlCallback = null,
            bool isDarkTheme = false)
        {
            InitializeComponent();
            _config = config ?? throw new ArgumentNullException(nameof(config));
            _onInsertSqlCallback = onInsertSqlCallback;
            IsDarkTheme = isDarkTheme;

            txtTableName.Text = initialTableName ?? "";

            Loaded += async (s, e) =>
            {
                if (!string.IsNullOrWhiteSpace(txtTableName.Text))
                {
                    await LoadStructureAsync(txtTableName.Text.Trim());
                }
                else
                {
                    txtTableName.Focus();
                }
            };
        }

        public static void ShowDialog(
            Window owner,
            OracleConnectionConfig config,
            string initialTableName,
            Action<string>? onInsertSqlCallback = null,
            bool isDarkTheme = false)
        {
            try
            {
                var dlg = new OracleTableStructureDialog(config, initialTableName, onInsertSqlCallback, isDarkTheme);
                dlg.Owner = owner;
                dlg.ShowDialog();
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Không thể mở hộp thoại Cấu Trúc Bảng:\n{ex.Message}",
                    "Lỗi", WpfMessageBoxButton.OK, WpfMessageBoxImage.Error);
            }
        }

        private async Task LoadStructureAsync(string tableName)
        {
            if (_isLoading) return;
            if (string.IsNullOrWhiteSpace(tableName))
            {
                txtStatus.Text = LocalizationService.Get("Oracle_ErrEmptyTable", "Vui lòng nhập tên bảng cần tra cứu.");
                return;
            }

            try
            {
                _isLoading = true;
                btnRefresh.IsEnabled = false;
                pbLoading.Visibility = Visibility.Visible;
                txtStatus.Text = LocalizationService.Get("Oracle_StatusLoadingStructure", "Đang tải cấu trúc bảng từ Oracle Data Dictionary...");

                _currentStructure = await OracleQuickQueryService.GetTableStructureAsync(_config, tableName);

                dgColumns.ItemsSource = null;
                dgColumns.ItemsSource = _currentStructure.Columns;

                dgIndexes.ItemsSource = null;
                dgIndexes.ItemsSource = _currentStructure.Indexes;

                txtColCount.Text = $"{_currentStructure.Columns.Count} Cột";
                txtIdxCount.Text = $"{_currentStructure.Indexes.Count} Index";

                txtStatus.Text = LocalizationService.Get("Oracle_StatusStructureLoaded", 
                    "Đã tải thành công cấu trúc bảng '{0}' ({1} cột, {2} index).",
                    _currentStructure.FullTableName, _currentStructure.Columns.Count, _currentStructure.Indexes.Count);

                btnInsertSelect.IsEnabled = _currentStructure.Columns.Count > 0;
                btnExportExcel.IsEnabled = _currentStructure.Columns.Count > 0;
            }
            catch (Exception ex)
            {
                txtStatus.Text = $"Lỗi: {ex.Message}";
                WpfMessageBox.Show(this, 
                    $"Không thể tải cấu trúc bảng '{tableName}':\n{ex.Message}", 
                    "Lỗi tra cứu bảng", 
                    WpfMessageBoxButton.OK, 
                    WpfMessageBoxImage.Warning);
            }
            finally
            {
                _isLoading = false;
                btnRefresh.IsEnabled = true;
                pbLoading.Visibility = Visibility.Collapsed;
            }
        }

        private async void BtnRefresh_Click(object sender, RoutedEventArgs e)
        {
            await LoadStructureAsync(txtTableName.Text.Trim());
        }

        private async void TxtTableName_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                await LoadStructureAsync(txtTableName.Text.Trim());
            }
        }

        private void BtnInsertSelect_Click(object sender, RoutedEventArgs e)
        {
            if (_currentStructure == null || _currentStructure.Columns.Count == 0)
            {
                WpfMessageBox.Show(this, "Không có dữ liệu cột để sinh câu lệnh SELECT.", "Thông báo", WpfMessageBoxButton.OK, WpfMessageBoxImage.Information);
                return;
            }

            string sql = _currentStructure.GenerateSelectQuery();
            _onInsertSqlCallback?.Invoke(sql);
            Close();
        }

        private void BtnExportExcel_Click(object sender, RoutedEventArgs e)
        {
            if (_currentStructure == null || _currentStructure.Columns.Count == 0)
            {
                WpfMessageBox.Show(this, "Không có dữ liệu cấu trúc bảng để xuất ra Excel.", "Thông báo", WpfMessageBoxButton.OK, WpfMessageBoxImage.Information);
                return;
            }

            try
            {
                var app = AddInEvents.Instance?.ExcelAppInstance;
                if (app == null)
                {
                    WpfMessageBox.Show(this, "Không tìm thấy phiên bản Excel đang hoạt động.", "Lỗi Excel", WpfMessageBoxButton.OK, WpfMessageBoxImage.Warning);
                    return;
                }

                Workbook? wb = null;
                try { wb = app.ActiveWorkbook; } catch { }
                if (wb == null)
                {
                    wb = app.Workbooks.Add();
                }

                // Create a new sheet
                ExcelWorksheet newWs = (ExcelWorksheet)wb.Worksheets.Add();
                string sheetBaseName = $"STRUC_{_currentStructure.TableName}";
                if (sheetBaseName.Length > 28) sheetBaseName = sheetBaseName.Substring(0, 28);
                
                // Ensure unique sheet name
                string sheetName = sheetBaseName;
                int counter = 1;
                bool nameExists = true;
                while (nameExists)
                {
                    nameExists = false;
                    foreach (ExcelWorksheet sheet in wb.Worksheets)
                    {
                        if (string.Equals(sheet.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                        {
                            nameExists = true;
                            sheetName = $"{sheetBaseName}_{counter++}";
                            break;
                        }
                    }
                }

                newWs.Name = sheetName;
                var (rows, cols) = OracleQuickQueryService.ExportTableStructureToWorksheet(newWs, _currentStructure);
                newWs.Activate();

                txtStatus.Text = $"Đã xuất cấu trúc bảng sang Sheet '{sheetName}' ({rows} dòng).";
                WpfMessageBox.Show(this, 
                    $"Đã xuất cấu trúc bảng sang Sheet mới '{sheetName}' thành công!", 
                    "Xuất thành công", 
                    WpfMessageBoxButton.OK, 
                    WpfMessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show(this, $"Lỗi khi xuất bảng sang Excel:\n{ex.Message}", "Lỗi xuất Excel", WpfMessageBoxButton.OK, WpfMessageBoxImage.Error);
            }
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private async void Window_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                Close();
            }
            else if (e.Key == Key.F5)
            {
                e.Handled = true;
                await LoadStructureAsync(txtTableName.Text.Trim());
            }
        }
    }
}
