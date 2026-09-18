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
    public partial class SpecialCopyDialog : System.Windows.Window
    {
        public static readonly DependencyProperty IsDarkThemeProperty =
            DependencyProperty.Register(nameof(IsDarkTheme), typeof(bool), typeof(SpecialCopyDialog),
                new PropertyMetadata(false));

        public bool IsDarkTheme
        {
            get => (bool)GetValue(IsDarkThemeProperty);
            set => SetValue(IsDarkThemeProperty, value);
        }

        private static SpecialCopyDialog? _currentInstance;

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

                _currentInstance = new SpecialCopyDialog(app)
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
                WpfMessageBox.Show($"Lỗi mở màn hình Special Copy:\n{ex.Message}",
                                   "ExcelSupport", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private readonly ExcelApp? _excelApp;
        private Range? _sourceRange;
        private bool _isInitializing = true;

        public SpecialCopyDialog(ExcelApp? app)
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
            _isInitializing = false;
            if (_excelApp == null) return;

            try
            {
                Range? sel = _excelApp.Selection as Range;
                if (sel != null)
                {
                    _sourceRange = sel;
                    TxtSourceRange.Text = sel.Address[false, false];
                    UpdateSourceInfo(sel);
                }
            }
            catch { }

            UpdatePreview();
        }

        private void OnPickSourceOnExcelClick(object sender, RoutedEventArgs e)
        {
            if (_excelApp == null) return;

            try
            {
                this.Visibility = Visibility.Hidden;

                dynamic app = _excelApp;
                dynamic result = app.InputBox(
                    Prompt: LocalizationService.Get("SC_PromptPickSource"),
                    Title: LocalizationService.Get("SC_TitlePickSource"),
                    Default: TxtSourceRange.Text.Trim(),
                    Type: 8);

                this.Visibility = Visibility.Visible;
                this.Activate();

                Range? srcRange = result as Range;
                if (srcRange != null)
                {
                    _sourceRange = srcRange;
                    TxtSourceRange.Text = srcRange.Address[false, false];
                    UpdateSourceInfo(srcRange);
                    UpdatePreview();
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
                    WpfMessageBox.Show(this, LocalizationService.Get("FCP_MsgSelectRange"), "Thông Báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                _sourceRange = sel;
                TxtSourceRange.Text = sel.Address[false, false];
                UpdateSourceInfo(sel);
                UpdatePreview();
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show(this, $"Lỗi nhận diện vùng chọn: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void UpdateSourceInfo(Range range)
        {
            try
            {
                int totalCells = range.Cells.Count;
                int totalRows = range.Rows.Count;
                int totalCols = range.Columns.Count;

                int visibleCells = totalCells;
                try
                {
                    Range vis = range.SpecialCells(XlCellType.xlCellTypeVisible);
                    visibleCells = vis.Cells.Count;
                }
                catch { }

                int hiddenCells = totalCells - visibleCells;

                if (hiddenCells > 0)
                {
                    TxtSourceInfo.Text = string.Format(LocalizationService.Get("SC_SourceInfoFiltered"), totalRows, totalCols, visibleCells, hiddenCells);
                }
                else
                {
                    TxtSourceInfo.Text = string.Format(LocalizationService.Get("SC_SourceInfoNormal"), totalRows, totalCols, totalCells);
                }
            }
            catch
            {
                TxtSourceInfo.Text = LocalizationService.Get("SC_SourceNoSelect");
            }
        }

        private SpecialCopyOptions BuildOptions()
        {
            var options = new SpecialCopyOptions();

            // 1. Delimiter
            if (RbCommaSpace.IsChecked == true) options.Delimiter = SpecialCopyDelimiter.CommaSpace;
            else if (RbSemicolon.IsChecked == true) options.Delimiter = SpecialCopyDelimiter.Semicolon;
            else if (RbPipe.IsChecked == true) options.Delimiter = SpecialCopyDelimiter.Pipe;
            else if (RbTab.IsChecked == true) options.Delimiter = SpecialCopyDelimiter.Tab;
            else if (RbSpace.IsChecked == true) options.Delimiter = SpecialCopyDelimiter.Space;
            else if (RbNewline.IsChecked == true) options.Delimiter = SpecialCopyDelimiter.Newline;
            else if (RbCustom.IsChecked == true)
            {
                options.Delimiter = SpecialCopyDelimiter.Custom;
                options.CustomDelimiter = TxtCustomDelimiter.Text;
            }
            else
            {
                options.Delimiter = SpecialCopyDelimiter.Comma;
            }

            // 2. Join Mode
            if (RbModeByRow.IsChecked == true) options.JoinMode = SpecialCopyJoinMode.ByRow;
            else if (RbModeByColumn.IsChecked == true) options.JoinMode = SpecialCopyJoinMode.ByColumn;
            else options.JoinMode = SpecialCopyJoinMode.AllCells;

            // 3. Quote Mode
            if (RbQuoteAutoCsv.IsChecked == true) options.QuoteMode = SpecialCopyQuoteMode.AutoCsv;
            else if (RbQuoteAlwaysDouble.IsChecked == true) options.QuoteMode = SpecialCopyQuoteMode.AlwaysDoubleQuote;
            else if (RbQuoteSingleSql.IsChecked == true) options.QuoteMode = SpecialCopyQuoteMode.SingleQuoteSql;
            else options.QuoteMode = SpecialCopyQuoteMode.None;

            // 4. Flags
            options.VisibleOnly = ChkVisibleOnly.IsChecked == true;
            options.SkipBlanks = ChkSkipBlanks.IsChecked == true;
            options.TrimSpaces = ChkTrimSpaces.IsChecked == true;

            return options;
        }

        private void UpdatePreview()
        {
            if (_isInitializing) return;

            try
            {
                Range? rng = _sourceRange;
                if (rng == null && _excelApp != null)
                {
                    try { rng = _excelApp.Selection as Range; } catch { }
                }

                if (rng == null)
                {
                    TxtPreview.Text = string.Empty;
                    TxtPreviewStats.Text = "0 ô | 0 dòng";
                    return;
                }

                var options = BuildOptions();

                Range targetRange = rng;
                if (options.VisibleOnly)
                {
                    try { targetRange = rng.SpecialCells(XlCellType.xlCellTypeVisible); }
                    catch { targetRange = rng; }
                }

                var rowDictValues = new SortedDictionary<int, SortedDictionary<int, object?>>();
                int totalCells = 0;

                foreach (Range area in targetRange.Areas)
                {
                    int rowCount = area.Rows.Count;
                    int colCount = area.Columns.Count;
                    int baseRow = area.Row;
                    int baseCol = area.Column;

                    object? rawValues = area.Value2;
                    object?[,]? valArray = rawValues as object[,];

                    for (int r = 1; r <= rowCount; r++)
                    {
                        int actualRow = baseRow + r - 1;
                        if (!rowDictValues.ContainsKey(actualRow))
                        {
                            rowDictValues[actualRow] = new SortedDictionary<int, object?>();
                        }

                        for (int c = 1; c <= colCount; c++)
                        {
                            int actualCol = baseCol + c - 1;
                            object? val = (valArray != null) ? valArray[r, c] : rawValues;
                            rowDictValues[actualRow][actualCol] = val;
                            totalCells++;
                        }
                    }
                }

                string delimiter = options.GetDelimiterString();
                var lines = new List<string>();

                if (options.JoinMode == SpecialCopyJoinMode.ByRow)
                {
                    foreach (var kvp in rowDictValues)
                    {
                        var cellStrings = new List<string>();
                        foreach (var colVal in kvp.Value)
                        {
                            object? val = colVal.Value;
                            if (options.SkipBlanks && (val == null || string.IsNullOrWhiteSpace(val.ToString()))) continue;
                            cellStrings.Add(FormatCellString(val, delimiter, options));
                        }

                        if (cellStrings.Count > 0 || !options.SkipBlanks)
                        {
                            lines.Add(string.Join(delimiter, cellStrings));
                        }
                    }
                }
                else if (options.JoinMode == SpecialCopyJoinMode.ByColumn)
                {
                    var colDictValues = new SortedDictionary<int, SortedDictionary<int, object?>>();
                    foreach (var rowKvp in rowDictValues)
                    {
                        int rowIdx = rowKvp.Key;
                        foreach (var colKvp in rowKvp.Value)
                        {
                            int colIdx = colKvp.Key;
                            if (!colDictValues.ContainsKey(colIdx))
                            {
                                colDictValues[colIdx] = new SortedDictionary<int, object?>();
                            }
                            colDictValues[colIdx][rowIdx] = colKvp.Value;
                        }
                    }

                    foreach (var kvp in colDictValues)
                    {
                        var cellStrings = new List<string>();
                        foreach (var rowVal in kvp.Value)
                        {
                            object? val = rowVal.Value;
                            if (options.SkipBlanks && (val == null || string.IsNullOrWhiteSpace(val.ToString()))) continue;
                            cellStrings.Add(FormatCellString(val, delimiter, options));
                        }

                        if (cellStrings.Count > 0 || !options.SkipBlanks)
                        {
                            lines.Add(string.Join(delimiter, cellStrings));
                        }
                    }
                }
                else // AllCells
                {
                    var allCellStrings = new List<string>();
                    foreach (var rowKvp in rowDictValues)
                    {
                        foreach (var colVal in rowKvp.Value)
                        {
                            object? val = colVal.Value;
                            if (options.SkipBlanks && (val == null || string.IsNullOrWhiteSpace(val.ToString()))) continue;
                            allCellStrings.Add(FormatCellString(val, delimiter, options));
                        }
                    }
                    lines.Add(string.Join(delimiter, allCellStrings));
                }

                string previewText = string.Join(Environment.NewLine, lines);
                TxtPreview.Text = previewText;
                TxtPreviewStats.Text = $"{totalCells} ô | {lines.Count} dòng | {previewText.Length} ký tự";
            }
            catch (Exception ex)
            {
                TxtPreview.Text = $"[Lỗi tạo xem trước: {ex.Message}]";
            }
        }

        private static string FormatCellString(object? val, string delimiter, SpecialCopyOptions options)
        {
            string s = val?.ToString() ?? string.Empty;
            if (options.TrimSpaces) s = s.Trim();

            switch (options.QuoteMode)
            {
                case SpecialCopyQuoteMode.AutoCsv:
                    if (s.Contains(delimiter) || s.Contains("\"") || s.Contains("\r") || s.Contains("\n"))
                    {
                        s = "\"" + s.Replace("\"", "\"\"") + "\"";
                    }
                    break;

                case SpecialCopyQuoteMode.AlwaysDoubleQuote:
                    s = "\"" + s.Replace("\"", "\"\"") + "\"";
                    break;

                case SpecialCopyQuoteMode.SingleQuoteSql:
                    s = "'" + s.Replace("'", "''") + "'";
                    break;
            }

            return s;
        }

        private void OnOptionChanged(object sender, RoutedEventArgs e)
        {
            UpdatePreview();
        }

        private void OnCustomDelimiterChanged(object sender, TextChangedEventArgs e)
        {
            if (RbCustom != null && RbCustom.IsChecked == true)
            {
                UpdatePreview();
            }
        }

        private void OnCopyToClipboardClick(object sender, RoutedEventArgs e)
        {
            var options = BuildOptions();
            var result = SpecialCopyService.ExecuteSpecialCopy(_excelApp, _sourceRange, options);

            if (result.Success)
            {
                TxtFooterStatus.Text = result.Message;
                WpfMessageBox.Show(this, 
                    $"{result.Message}\n\n👉 Bạn có thể chọn bất kỳ ô nào trên Excel hoặc cửa sổ ứng dụng khác và nhấn Ctrl + V để dán!", 
                    LocalizationService.Get("SC_Title"), 
                    MessageBoxButton.OK, 
                    MessageBoxImage.Information);
            }
            else
            {
                TxtFooterStatus.Text = result.Message;
                WpfMessageBox.Show(this, result.Message, "Thông Báo", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void OnPasteToExcelClick(object sender, RoutedEventArgs e)
        {
            if (_excelApp == null) return;

            var options = BuildOptions();
            var copyResult = SpecialCopyService.ExecuteSpecialCopy(_excelApp, _sourceRange, options);
            if (!copyResult.Success)
            {
                WpfMessageBox.Show(this, copyResult.Message, "Thông Báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            Range? targetRange = null;
            try
            {
                this.Visibility = Visibility.Hidden;

                dynamic app = _excelApp;
                dynamic result = app.InputBox(
                    Prompt: LocalizationService.Get("FCP_PromptPickTarget"),
                    Title: LocalizationService.Get("FCP_TitlePickTarget"),
                    Default: _excelApp.ActiveCell?.Address[false, false] ?? "A1",
                    Type: 8);

                targetRange = result as Range;
            }
            catch
            {
                // Người dùng bấm Cancel trên InputBox
                this.Visibility = Visibility.Visible;
                this.Activate();
                return;
            }
            finally
            {
                this.Visibility = Visibility.Visible;
                this.Activate();
            }

            if (targetRange == null) return;

            var pasteResult = SpecialCopyService.PasteSpecialCopy(_excelApp, targetRange);
            if (pasteResult.Success)
            {
                TxtFooterStatus.Text = pasteResult.Message;
                WpfMessageBox.Show(this, 
                    $"{pasteResult.Message}\n\n👉 Dữ liệu chuỗi nối đã được ghi thành công vào ô [{targetRange.Address[false, false]}]!", 
                    LocalizationService.Get("SC_Title"), 
                    MessageBoxButton.OK, 
                    MessageBoxImage.Information);
            }
            else
            {
                TxtFooterStatus.Text = pasteResult.Message;
                WpfMessageBox.Show(this, pasteResult.Message, "Thông Báo", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void OnCloseClick(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
