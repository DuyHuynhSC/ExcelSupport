using System;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using ExcelSupport.Models;
using ExcelSupport.Services;
using WpfMessageBox = System.Windows.MessageBox;

namespace ExcelSupport.Views
{
    public partial class DataCleaningWizardDialog : Window, INotifyPropertyChanged
    {
        private bool _isDarkTheme;
        private bool _isInitialized = false;
        private static DataCleaningWizardDialog? _currentInstance;

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

                _currentInstance = new DataCleaningWizardDialog(isDarkTheme);
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
                WpfMessageBox.Show($"Lỗi mở màn hình Dọn dẹp dữ liệu:\n{ex.Message}\n\nChi tiết:\n{ex.StackTrace}",
                                   "Lỗi Khởi Tạo", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        public DataCleaningWizardDialog(bool isDarkTheme = false)
        {
            InitializeComponent();
            IsDarkTheme = isDarkTheme;
            DataContext = this;
            _isInitialized = true;

            if (TxtSampleInput != null)
            {
                TxtSampleInput.TextChanged += OnSampleTextChanged;
            }

            Loaded += (s, e) => UpdatePreview();
        }

        private DataCleaningOptions BuildOptions()
        {
            var options = new DataCleaningOptions();
            if (!_isInitialized) return options;

            // Scope
            if (RbScopeSelection != null && RbScopeSelection.IsChecked == true) options.Scope = CleaningScope.SelectedRange;
            else if (RbScopeSheet != null && RbScopeSheet.IsChecked == true) options.Scope = CleaningScope.ActiveSheet;
            else if (RbScopeWorkbook != null && RbScopeWorkbook.IsChecked == true) options.Scope = CleaningScope.ActiveWorkbook;

            // Whitespace
            options.TrimSpaces = ChkTrim?.IsChecked == true;
            options.ReduceMultipleSpaces = ChkReduceSpaces?.IsChecked == true;
            options.RemoveNonBreakingSpaces = ChkRemoveNbsp?.IsChecked == true;
            options.RemoveLineBreaks = ChkRemoveLineBreaks?.IsChecked == true;
            options.RemoveUnprintableChars = ChkRemoveUnprintable?.IsChecked == true;

            // Case
            if (RbCaseUpper?.IsChecked == true) options.CaseOption = TextCaseOption.UpperCase;
            else if (RbCaseLower?.IsChecked == true) options.CaseOption = TextCaseOption.LowerCase;
            else if (RbCaseProper?.IsChecked == true) options.CaseOption = TextCaseOption.ProperCase;
            else if (RbCaseSentence?.IsChecked == true) options.CaseOption = TextCaseOption.SentenceCase;
            else options.CaseOption = TextCaseOption.None;

            // Vietnamese & Japanese
            options.RemoveVietnameseDiacritics = ChkRemoveAccents?.IsChecked == true;
            options.ConvertVietnameseToKatakana = ChkToKatakana?.IsChecked == true;
            options.KatakanaUseMiddleDot = RbKataDot?.IsChecked == true;
            options.JapaneseHalfWidthToFullWidth = ChkHankakuToZenkaku?.IsChecked == true;
            options.JapaneseFullWidthToHalfWidth = ChkZenkakuToHankaku?.IsChecked == true;
            options.RemoveDigits = ChkRemoveDigits?.IsChecked == true;
            options.RemoveLetters = ChkRemoveLetters?.IsChecked == true;
            options.RemoveSpecialSymbols = ChkRemoveSymbols?.IsChecked == true;

            // Numbers & Dates
            options.ConvertNumbersStoredAsText = ChkNumbersAsText?.IsChecked == true;
            options.StandardizeDates = ChkStandardizeDate?.IsChecked == true;
            if (CboDateFormat?.SelectedItem is ComboBoxItem item)
            {
                options.DateFormat = item.Content?.ToString() ?? "yyyy-MM-dd";
            }

            // Blanks & Errors
            if (ChkFillBlanks?.IsChecked == true)
            {
                if (RbFillDown?.IsChecked == true) options.FillBlanks = BlankFillOption.FillDownFromAbove;
                else
                {
                    options.FillBlanks = BlankFillOption.CustomValue;
                    options.CustomBlankValue = TxtCustomBlank?.Text ?? string.Empty;
                }
            }
            else
            {
                options.FillBlanks = BlankFillOption.None;
            }

            options.ReplaceErrorValues = ChkFixErrors?.IsChecked == true;
            options.CustomErrorReplacement = TxtCustomError?.Text ?? string.Empty;

            return options;
        }

        private void OnSampleTextChanged(object sender, TextChangedEventArgs e)
        {
            if (!_isInitialized) return;
            UpdatePreview();
        }

        private void OnRefreshPreviewClick(object sender, RoutedEventArgs e)
        {
            UpdatePreview();
        }

        private void UpdatePreview()
        {
            if (!_isInitialized) return;

            try
            {
                string input = TxtSampleInput?.Text ?? string.Empty;
                var opts = BuildOptions();

                // Run preview through cleaning logic
                string processed = input;

                if (opts.RemoveNonBreakingSpaces) processed = processed.Replace("\u00A0", " ").Replace("&nbsp;", " ");
                if (opts.RemoveLineBreaks) processed = processed.Replace("\r", " ").Replace("\n", " ");
                if (opts.TrimSpaces) processed = processed.Trim();
                if (opts.ReduceMultipleSpaces)
                {
                    while (processed.Contains("  ")) processed = processed.Replace("  ", " ");
                }

                if (opts.CaseOption == TextCaseOption.UpperCase) processed = processed.ToUpper();
                else if (opts.CaseOption == TextCaseOption.LowerCase) processed = processed.ToLower();
                else if (opts.CaseOption == TextCaseOption.ProperCase) processed = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(processed.ToLower());
                else if (opts.CaseOption == TextCaseOption.SentenceCase && processed.Length > 0)
                {
                    processed = char.ToUpper(processed[0]) + (processed.Length > 1 ? processed.Substring(1).ToLower() : "");
                }

                if (opts.RemoveVietnameseDiacritics) processed = VietnameseToKatakanaConverter.RemoveDiacritics(processed);
                if (opts.ConvertVietnameseToKatakana) processed = VietnameseToKatakanaConverter.ConvertToKatakana(processed, opts.KatakanaUseMiddleDot);

                if (TxtSampleOutput != null) TxtSampleOutput.Text = processed;
                if (TxtSampleKatakana != null) TxtSampleKatakana.Text = VietnameseToKatakanaConverter.ConvertToKatakana(input, true);
            }
            catch { }
        }

        private void OnExecuteClick(object sender, RoutedEventArgs e)
        {
            var addIn = AddInEvents.Instance;
            if (addIn == null) return;

            var options = BuildOptions();
            TxtStatus.Text = "⏳ Đang dọn dẹp và chuẩn hóa dữ liệu trên Excel...";

            bool success = addIn.ExecuteDataCleaning(options, out int modifiedCount, out string statusMsg);
            TxtStatus.Text = statusMsg;

            if (success)
            {
                WpfMessageBox.Show(statusMsg, "Dọn Dẹp Hoàn Tất", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                WpfMessageBox.Show(statusMsg, "Thông Báo", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void OnCloseClick(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
