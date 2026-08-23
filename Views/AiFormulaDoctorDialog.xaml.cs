using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using ExcelSupport.Helpers;
using ExcelSupport.Models;
using ExcelSupport.Services;
using Microsoft.Office.Interop.Excel;
using WpfMessageBox = System.Windows.MessageBox;
using WpfMessageBoxButton = System.Windows.MessageBoxButton;
using WpfMessageBoxImage = System.Windows.MessageBoxImage;
using MediaColor = System.Windows.Media.Color;
using WpfWindow = System.Windows.Window;

namespace ExcelSupport.Views
{
    public partial class AiFormulaDoctorDialog : WpfWindow
    {
        private readonly ObservableCollection<FormulaCellItem> _errorItems = new ObservableCollection<FormulaCellItem>();
        private FormulaCellItem? _selectedErrorItem;
        private readonly bool _isDarkTheme;
        private static AiFormulaDoctorDialog? _currentInstance;

        public AiFormulaDoctorDialog(bool isDarkTheme = false)
        {
            InitializeComponent();
            _isDarkTheme = isDarkTheme;
            dgErrorCells.ItemsSource = _errorItems;

            if (_isDarkTheme)
            {
                ApplyDarkTheme();
            }

            Loaded += AiFormulaDoctorDialog_Loaded;
        }

        private void AiFormulaDoctorDialog_Loaded(object sender, RoutedEventArgs e)
        {
            // Auto scan current sheet on open
            BtnScanSheet_Click(this, new RoutedEventArgs());
        }

        public static void ShowWindow(bool isDarkTheme = false)
        {
            try
            {
                if (_currentInstance != null && _currentInstance.IsLoaded)
                {
                    _currentInstance.Activate();
                    return;
                }

                _currentInstance = new AiFormulaDoctorDialog(isDarkTheme);

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
                WpfMessageBox.Show($"Không thể mở Bác Sĩ Công Thức:\n{ex.Message}", "Lỗi giao diện", WpfMessageBoxButton.OK, WpfMessageBoxImage.Error);
            }
        }

        private void ApplyDarkTheme()
        {
            RootGrid.Background = new System.Windows.Media.SolidColorBrush(MediaColor.FromRgb(15, 23, 42));
        }

        #region Scanning Errors

        private async void BtnScanSheet_Click(object sender, RoutedEventArgs e)
        {
            await RunScanAsync(scanSelectionOnly: false);
        }

        private async void BtnScanSelection_Click(object sender, RoutedEventArgs e)
        {
            await RunScanAsync(scanSelectionOnly: true);
        }

        private async Task RunScanAsync(bool scanSelectionOnly)
        {
            var app = AddInEvents.Instance?.ExcelAppInstance;
            if (app == null)
            {
                txtStatus.Text = "Không thể kết nối với tiến trình Excel.";
                return;
            }

            txtStatus.Text = scanSelectionOnly ? "Đang quét lỗi trong vùng chọn..." : "Đang quét toàn bộ công thức trên Sheet...";
            pbProgress.Visibility = Visibility.Visible;
            pbProgress.IsIndeterminate = true;

            try
            {
                var scanResult = await Task.Run(() => AiFormulaDoctorService.ScanForErrors(app, scanSelectionOnly));

                _errorItems.Clear();
                foreach (var item in scanResult.ErrorItems)
                {
                    _errorItems.Add(item);
                }

                txtErrorBadge.Text = $"{_errorItems.Count} lỗi";
                txtStatus.Text = $"Quét hoàn tất trong {scanResult.ScanDuration.TotalSeconds:F2}s. Phát hiện {_errorItems.Count} ô bị lỗi.";

                if (_errorItems.Count > 0)
                {
                    dgErrorCells.SelectedIndex = 0;
                }
                else
                {
                    ClearDiagnosisCard();
                    txtAiDiagnosis.Text = "🎉 Tuyệt vời! Không phát hiện lỗi công thức nào trên Sheet.";
                }
            }
            catch (Exception ex)
            {
                txtStatus.Text = $"Lỗi quét: {ex.Message}";
            }
            finally
            {
                pbProgress.Visibility = Visibility.Collapsed;
                pbProgress.IsIndeterminate = false;
            }
        }

        #endregion

        #region Diagnosis & Fixing

        private async void DgErrorCells_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!IsLoaded) return;

            _selectedErrorItem = dgErrorCells.SelectedItem as FormulaCellItem;
            if (_selectedErrorItem == null)
            {
                ClearDiagnosisCard();
                return;
            }

            await DisplayAndDiagnoseItemAsync(_selectedErrorItem);
        }

        private async Task DisplayAndDiagnoseItemAsync(FormulaCellItem item)
        {
            txtSelectedCellBadge.Text = $"Ô: {item.CellAddress}";
            txtCurrentFormula.Text = item.Formula;
            txtProposedFormula.Text = item.ProposedFormula ?? string.Empty;
            txtAiDiagnosis.Text = item.AiDiagnosis ?? "Đang phân tích nguyên nhân lỗi...";
            txtFixExplanation.Text = item.FixExplanation ?? string.Empty;

            if (string.IsNullOrWhiteSpace(item.AiDiagnosis) || string.IsNullOrWhiteSpace(item.ProposedFormula))
            {
                pbProgress.Visibility = Visibility.Visible;
                pbProgress.IsIndeterminate = true;
                txtStatus.Text = $"Đang chuẩn đoán ô {item.CellAddress}...";

                try
                {
                    var aiConfig = AiConfigManager.Current;
                    var lang = LocalizationService.CurrentLanguage;

                    await Task.Run(() => AiFormulaDoctorService.DiagnoseAndProposeFixAsync(item, aiConfig, lang));

                    txtAiDiagnosis.Text = item.AiDiagnosis ?? "Không thể xác định lỗi.";
                    txtProposedFormula.Text = item.ProposedFormula ?? item.Formula;
                    txtFixExplanation.Text = item.FixExplanation ?? string.Empty;
                    txtStatus.Text = $"Đã chuẩn đoán xong ô {item.CellAddress}.";
                }
                catch (Exception ex)
                {
                    txtStatus.Text = $"Lỗi chuẩn đoán: {ex.Message}";
                }
                finally
                {
                    pbProgress.Visibility = Visibility.Collapsed;
                    pbProgress.IsIndeterminate = false;
                }
            }
        }

        private async void BtnReDiagnose_Click(object sender, RoutedEventArgs e)
        {
            if (_selectedErrorItem == null) return;
            _selectedErrorItem.AiDiagnosis = null;
            _selectedErrorItem.ProposedFormula = null;
            await DisplayAndDiagnoseItemAsync(_selectedErrorItem);
        }

        private void BtnApplyFix_Click(object sender, RoutedEventArgs e)
        {
            if (_selectedErrorItem == null)
            {
                WpfMessageBox.Show("Vui lòng chọn ô công thức cần sửa.", "Thông báo", WpfMessageBoxButton.OK, WpfMessageBoxImage.Information);
                return;
            }

            var app = AddInEvents.Instance?.ExcelAppInstance;
            if (app == null) return;

            _selectedErrorItem.ProposedFormula = txtProposedFormula.Text.Trim();
            bool success = AiFormulaDoctorService.ApplyFixToCell(app, _selectedErrorItem);

            if (success)
            {
                txtStatus.Text = $"Đã áp dụng công thức mới cho ô {_selectedErrorItem.CellAddress} thành công!";
                WpfMessageBox.Show($"Đã sửa công thức cho ô {_selectedErrorItem.CellAddress} thành công!", "Thành công", WpfMessageBoxButton.OK, WpfMessageBoxImage.Information);
            }
            else
            {
                WpfMessageBox.Show($"Không thể áp dụng công thức vào ô {_selectedErrorItem.CellAddress}.", "Lỗi áp dụng", WpfMessageBoxButton.OK, WpfMessageBoxImage.Error);
            }
        }

        private void BtnBatchFixColumn_Click(object sender, RoutedEventArgs e)
        {
            if (_selectedErrorItem == null)
            {
                WpfMessageBox.Show("Vui lòng chọn một ô mẫu trong cột để áp dụng sửa hàng loạt.", "Thông báo", WpfMessageBoxButton.OK, WpfMessageBoxImage.Information);
                return;
            }

            var app = AddInEvents.Instance?.ExcelAppInstance;
            if (app == null) return;

            _selectedErrorItem.ProposedFormula = txtProposedFormula.Text.Trim();
            int count = AiFormulaDoctorService.BatchApplyFixToColumn(app, _selectedErrorItem, _errorItems.ToList());

            txtStatus.Text = $"Đã tự động sửa {count} ô trong cùng cột thành công!";
            WpfMessageBox.Show($"Đã tự động áp dụng và sửa thành công {count} ô lỗi trong cột!", "Sửa hàng loạt thành công", WpfMessageBoxButton.OK, WpfMessageBoxImage.Information);
        }

        private void ClearDiagnosisCard()
        {
            txtSelectedCellBadge.Text = "Ô: --";
            txtCurrentFormula.Text = string.Empty;
            txtProposedFormula.Text = string.Empty;
            txtAiDiagnosis.Text = "Vui lòng chọn một ô lỗi từ danh sách bên trái để xem chuẩn đoán và đề xuất sửa.";
            txtFixExplanation.Text = string.Empty;
        }

        #endregion

        #region Explainer Tab

        private void BtnFetchActiveFormula_Click(object sender, RoutedEventArgs e)
        {
            var app = AddInEvents.Instance?.ExcelAppInstance;
            if (app?.ActiveCell != null)
            {
                string formula = app.ActiveCell.Formula?.ToString() ?? string.Empty;
                txtExplainFormulaInput.Text = formula;
            }
        }

        private async void BtnRunExplain_Click(object sender, RoutedEventArgs e)
        {
            string formula = txtExplainFormulaInput.Text.Trim();
            if (string.IsNullOrWhiteSpace(formula))
            {
                WpfMessageBox.Show("Vui lòng nhập công thức cần giải thích.", "Thông báo", WpfMessageBoxButton.OK, WpfMessageBoxImage.Information);
                return;
            }

            txtStatus.Text = "Đang phân tích và giải thích công thức...";
            pbProgress.Visibility = Visibility.Visible;
            pbProgress.IsIndeterminate = true;

            try
            {
                var aiConfig = AiConfigManager.Current;
                var lang = LocalizationService.CurrentLanguage;

                var result = await AiFormulaDoctorService.ExplainFormulaAsync(formula, aiConfig, lang);

                cardExplainPurpose.Visibility = Visibility.Visible;
                txtOverallPurpose.Text = result.OverallPurpose;
                icExplainSteps.ItemsSource = result.Steps;

                txtStatus.Text = "Giải thích công thức hoàn tất!";
            }
            catch (Exception ex)
            {
                txtStatus.Text = $"Lỗi giải thích: {ex.Message}";
            }
            finally
            {
                pbProgress.Visibility = Visibility.Collapsed;
                pbProgress.IsIndeterminate = false;
            }
        }

        #endregion

        #region Modernizer Tab

        private void BtnFetchActiveForModernize_Click(object sender, RoutedEventArgs e)
        {
            var app = AddInEvents.Instance?.ExcelAppInstance;
            if (app?.ActiveCell != null)
            {
                string formula = app.ActiveCell.Formula?.ToString() ?? string.Empty;
                txtModernizeInput.Text = formula;
            }
        }

        private async void BtnRunModernize_Click(object sender, RoutedEventArgs e)
        {
            string formula = txtModernizeInput.Text.Trim();
            if (string.IsNullOrWhiteSpace(formula))
            {
                WpfMessageBox.Show("Vui lòng nhập công thức cần tối ưu.", "Thông báo", WpfMessageBoxButton.OK, WpfMessageBoxImage.Information);
                return;
            }

            txtStatus.Text = "Đang tối ưu và hiện đại hóa công thức...";
            pbProgress.Visibility = Visibility.Visible;
            pbProgress.IsIndeterminate = true;

            try
            {
                var aiConfig = AiConfigManager.Current;
                var lang = LocalizationService.CurrentLanguage;

                var result = await AiFormulaDoctorService.ModernizeFormulaAsync(formula, aiConfig, lang);

                cardModernizeOutput.Visibility = Visibility.Visible;
                txtModernizedOutput.Text = result.ModernizedFormula;
                txtModernizeSummary.Text = result.ChangesSummary;

                txtStatus.Text = "Tối ưu hóa hoàn tất!";
            }
            catch (Exception ex)
            {
                txtStatus.Text = $"Lỗi tối ưu: {ex.Message}";
            }
            finally
            {
                pbProgress.Visibility = Visibility.Collapsed;
                pbProgress.IsIndeterminate = false;
            }
        }

        private void BtnApplyModernized_Click(object sender, RoutedEventArgs e)
        {
            var app = AddInEvents.Instance?.ExcelAppInstance;
            if (app?.ActiveCell == null) return;

            string modForm = txtModernizedOutput.Text.Trim();
            if (!string.IsNullOrWhiteSpace(modForm))
            {
                try
                {
                    app.ActiveCell.Formula = modForm;
                    txtStatus.Text = "Đã cập nhật công thức hiện đại vào ô Excel thành công!";
                    WpfMessageBox.Show("Đã cập nhật công thức vào ô Excel đang chọn thành công!", "Thành công", WpfMessageBoxButton.OK, WpfMessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    WpfMessageBox.Show($"Không thể gán công thức: {ex.Message}", "Lỗi", WpfMessageBoxButton.OK, WpfMessageBoxImage.Error);
                }
            }
        }

        #endregion

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
