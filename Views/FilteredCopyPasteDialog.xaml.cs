using System;
using System.IO;
using System.Linq;
using System.Windows;
using ExcelSupport.Host;
using ExcelSupport.Models;
using ExcelSupport.Services;
using Microsoft.Office.Interop.Excel;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using WpfMessageBox = System.Windows.MessageBox;

namespace ExcelSupport.Views
{
    public partial class FilteredCopyPasteDialog : System.Windows.Window
    {
        public static readonly DependencyProperty IsDarkThemeProperty =
            DependencyProperty.Register(nameof(IsDarkTheme), typeof(bool), typeof(FilteredCopyPasteDialog),
                new PropertyMetadata(false));

        public bool IsDarkTheme
        {
            get => (bool)GetValue(IsDarkThemeProperty);
            set => SetValue(IsDarkThemeProperty, value);
        }

        private static FilteredCopyPasteDialog? _currentInstance;

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

                _currentInstance = new FilteredCopyPasteDialog(app)
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
                WpfMessageBox.Show($"Lỗi mở màn hình Sao Chép & Dán Vùng Lọc:\n{ex.Message}",
                                   "ExcelSupport", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private readonly ExcelApp? _excelApp;

        public FilteredCopyPasteDialog(ExcelApp? app)
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
            if (_excelApp == null) return;

            try
            {
                Range? sel = _excelApp.Selection as Range;
                if (sel != null)
                {
                    string addr = sel.Address[false, false];
                    TxtSourceRange.Text = addr;
                    UpdateSourceInfo(sel);
                }
            }
            catch { }
        }

        private void OnPickSourceOnExcelClick(object sender, RoutedEventArgs e)
        {
            if (_excelApp == null) return;

            try
            {
                this.Visibility = Visibility.Hidden;

                dynamic app = _excelApp;
                dynamic result = app.InputBox(
                    Prompt: "Dùng chuột quét chọn Vùng Nguồn cần sao chép trên bảng tính Excel:",
                    Title: "Chọn Vùng Nguồn Cần Sao Chép",
                    Default: TxtSourceRange.Text.Trim(),
                    Type: 8);

                this.Visibility = Visibility.Visible;
                this.Activate();

                Range? srcRange = result as Range;
                if (srcRange != null)
                {
                    TxtSourceRange.Text = srcRange.Address[false, false];
                    UpdateSourceInfo(srcRange);
                }
            }
            catch
            {
                this.Visibility = Visibility.Visible;
                this.Activate();
            }
        }

        private void OnGetSourceSelectionClick(object sender, RoutedEventArgs e)
        {
            if (_excelApp == null) return;

            try
            {
                Range? sel = _excelApp.Selection as Range;
                if (sel == null)
                {
                    WpfMessageBox.Show(this, "Vui lòng chọn một vùng dữ liệu trên bảng tính Excel.", "Thông Báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                TxtSourceRange.Text = sel.Address[false, false];
                UpdateSourceInfo(sel);
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show(this, $"Lỗi nhận diện vùng chọn nguồn: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void UpdateSourceInfo(Range sel)
        {
            try
            {
                Range? visible = null;
                try { visible = sel.SpecialCells(XlCellType.xlCellTypeVisible); }
                catch { visible = sel; }

                int totalRows = sel.Rows.Count;
                int visibleRows = 0;

                if (visible != null)
                {
                    foreach (Range area in visible.Areas)
                    {
                        visibleRows += area.Rows.Count;
                    }
                }

                int hiddenRows = totalRows - visibleRows;
                TxtSourceInfo.Text = string.Format(LocalizationService.Get("FCP_SourceInfoFormat"), sel.Address[false, false], visibleRows, hiddenRows);
            }
            catch
            {
                TxtSourceInfo.Text = $"{LocalizationService.Get("FCP_SourceHeader")}: {sel.Address[false, false]}";
            }
        }

        private void OnPickTargetOnExcelClick(object sender, RoutedEventArgs e)
        {
            if (_excelApp == null) return;

            try
            {
                this.Visibility = Visibility.Hidden;

                dynamic app = _excelApp;
                dynamic result = app.InputBox(
                    Prompt: LocalizationService.Get("FCP_PromptPickTarget"),
                    Title: LocalizationService.Get("FCP_TitlePickTarget"),
                    Default: TxtTargetRange.Text.Trim(),
                    Type: 8);

                this.Visibility = Visibility.Visible;
                this.Activate();

                Range? targetRange = result as Range;
                if (targetRange != null)
                {
                    TxtTargetRange.Text = targetRange.Address[false, false];
                    UpdateTargetInfo(targetRange);
                }
            }
            catch
            {
                this.Visibility = Visibility.Visible;
                this.Activate();
            }
        }

        private void OnGetTargetSelectionClick(object sender, RoutedEventArgs e)
        {
            if (_excelApp == null) return;

            try
            {
                Range? sel = _excelApp.Selection as Range;
                if (sel == null)
                {
                    WpfMessageBox.Show(this, LocalizationService.Get("FCP_MsgSelectTarget"), LocalizationService.Get("Common_Notice"), MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                TxtTargetRange.Text = sel.Address[false, false];
                UpdateTargetInfo(sel);
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show(this, $"{LocalizationService.Get("Common_Error")}: {ex.Message}", LocalizationService.Get("Common_Error"), MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void UpdateTargetInfo(Range sel)
        {
            try
            {
                Worksheet? ws = sel.Worksheet;
                if (ws == null) return;

                int totalRows = sel.Rows.Count;
                if (totalRows == 1)
                {
                    TxtTargetInfo.Text = string.Format(LocalizationService.Get("FCP_TargetSingleFormat"), sel.Address[false, false], sel.Row);
                }
                else
                {
                    Range? visible = null;
                    try { visible = sel.SpecialCells(XlCellType.xlCellTypeVisible); }
                    catch { visible = sel; }

                    int visibleRows = 0;
                    if (visible != null)
                    {
                        foreach (Range area in visible.Areas)
                        {
                            visibleRows += area.Rows.Count;
                        }
                    }
                    int hiddenRows = totalRows - visibleRows;
                    TxtTargetInfo.Text = string.Format(LocalizationService.Get("FCP_TargetRangeFormat"), sel.Address[false, false], visibleRows, hiddenRows);
                }
            }
            catch
            {
                TxtTargetInfo.Text = $"{LocalizationService.Get("FCP_TargetHeader")}: {sel.Address[false, false]}";
            }
        }

        private void OnCopySourceClick(object sender, RoutedEventArgs e)
        {
            if (_excelApp == null) return;

            try
            {
                Range? srcRng = GetRangeFromText(TxtSourceRange.Text);
                if (srcRng == null)
                {
                    srcRng = _excelApp.Selection as Range;
                }

                if (srcRng == null)
                {
                    WpfMessageBox.Show(this, LocalizationService.Get("FCP_MsgSelectSource"), LocalizationService.Get("Common_Notice"), MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var result = FilteredCopyPasteService.CopyVisibleCells(_excelApp, srcRng);
                if (result.Success)
                {
                    TxtFooterStatus.Text = $"✅ {result.Message}";
                    WpfMessageBox.Show(this, result.Message, LocalizationService.Get("Common_Notice"), MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    TxtFooterStatus.Text = $"❌ {result.Message}";
                    WpfMessageBox.Show(this, result.Message, LocalizationService.Get("Common_Notice"), MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show(this, $"{LocalizationService.Get("Common_Error")}: {ex.Message}", LocalizationService.Get("Common_Error"), MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OnPasteTargetClick(object sender, RoutedEventArgs e)
        {
            if (_excelApp == null) return;

            try
            {
                Range? destRng = GetRangeFromText(TxtTargetRange.Text);
                if (destRng == null)
                {
                    destRng = _excelApp.Selection as Range;
                }

                if (destRng == null)
                {
                    WpfMessageBox.Show(this, LocalizationService.Get("FCP_MsgSelectTarget"), LocalizationService.Get("Common_Notice"), MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var options = BuildOptions();
                var result = FilteredCopyPasteService.PasteToVisibleCells(_excelApp, destRng, options);

                if (result.Success)
                {
                    TxtFooterStatus.Text = $"✅ {result.Message}";
                    WpfMessageBox.Show(this, result.Message, LocalizationService.Get("Common_Notice"), MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    TxtFooterStatus.Text = $"❌ {result.Message}";
                    WpfMessageBox.Show(this, result.Message, LocalizationService.Get("Common_Notice"), MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show(this, $"{LocalizationService.Get("Common_Error")}: {ex.Message}", LocalizationService.Get("Common_Error"), MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OnExecuteAllClick(object sender, RoutedEventArgs e)
        {
            if (_excelApp == null) return;

            try
            {
                Range? srcRng = GetRangeFromText(TxtSourceRange.Text) ?? (_excelApp.Selection as Range);
                Range? destRng = GetRangeFromText(TxtTargetRange.Text);

                if (srcRng == null)
                {
                    WpfMessageBox.Show(this, LocalizationService.Get("FCP_MsgSelectSource"), LocalizationService.Get("Common_Notice"), MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (destRng == null)
                {
                    WpfMessageBox.Show(this, LocalizationService.Get("FCP_MsgSelectTarget"), LocalizationService.Get("Common_Notice"), MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var options = BuildOptions();
                var result = FilteredCopyPasteService.ExecuteRangeToRangeCopyPaste(_excelApp, srcRng, destRng, options);

                if (result.Success)
                {
                    TxtFooterStatus.Text = $"✅ {result.Message}";
                    WpfMessageBox.Show(this, result.Message, LocalizationService.Get("Common_Notice"), MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    TxtFooterStatus.Text = $"❌ {result.Message}";
                    WpfMessageBox.Show(this, result.Message, LocalizationService.Get("Common_Notice"), MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show(this, $"{LocalizationService.Get("Common_Error")}: {ex.Message}", LocalizationService.Get("Common_Error"), MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private FilteredPasteOptions BuildOptions()
        {
            var pasteType = FilteredPasteType.ValuesOnly;
            if (RbPasteFormulas.IsChecked == true) pasteType = FilteredPasteType.Formulas;
            else if (RbPasteFormats.IsChecked == true) pasteType = FilteredPasteType.FormatsOnly;
            else if (RbPasteAll.IsChecked == true) pasteType = FilteredPasteType.All;

            return new FilteredPasteOptions
            {
                PasteType = pasteType,
                RepeatIfShorter = (ChkRepeatShorter.IsChecked == true),
                SkipBlanks = (ChkSkipBlanks.IsChecked == true),
                SourceAddress = TxtSourceRange.Text.Trim(),
                TargetAddress = TxtTargetRange.Text.Trim()
            };
        }

        private Range? GetRangeFromText(string text)
        {
            if (_excelApp == null || string.IsNullOrWhiteSpace(text)) return null;

            try
            {
                dynamic activeSheet = _excelApp.ActiveSheet;
                if (activeSheet != null)
                {
                    return activeSheet.Range[text.Trim()];
                }
            }
            catch { }

            return null;
        }

        private void OnCloseClick(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
