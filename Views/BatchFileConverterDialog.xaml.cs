using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using ExcelSupport.Host;
using ExcelSupport.Models;
using ExcelSupport.Services;
using Microsoft.Office.Interop.Excel;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using WpfMessageBox = System.Windows.MessageBox;
using Action = System.Action;

namespace ExcelSupport.Views
{
    public partial class BatchFileConverterDialog : System.Windows.Window
    {
        public static readonly DependencyProperty IsDarkThemeProperty =
            DependencyProperty.Register(nameof(IsDarkTheme), typeof(bool), typeof(BatchFileConverterDialog),
                new PropertyMetadata(false));

        public bool IsDarkTheme
        {
            get => (bool)GetValue(IsDarkThemeProperty);
            set => SetValue(IsDarkThemeProperty, value);
        }

        private static BatchFileConverterDialog? _currentInstance;

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

                _currentInstance = new BatchFileConverterDialog(app)
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
                WpfMessageBox.Show($"Lỗi mở màn hình Quản Trị & Chuyển Đổi File Hàng Loạt:\n{ex.Message}",
                                   "ExcelSupport", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private readonly ExcelApp? _excelApp;
        private readonly List<BatchFileItem> _fileItems = new List<BatchFileItem>();

        public BatchFileConverterDialog(ExcelApp? app)
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
            CboTargetFormat.ItemsSource = Enum.GetValues(typeof(ExcelOutputFormat));
            CboTargetFormat.SelectedItem = ExcelOutputFormat.PDF;

            string myDocs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            TxtOutputDir.Text = Path.Combine(myDocs, "Excel_Converted");

            // Mặc định nạp file Excel đang mở nếu có
            LoadActiveWorkbookIfAvailable();
        }

        private void LoadActiveWorkbookIfAvailable()
        {
            try
            {
                if (_excelApp != null)
                {
                    Workbook? activeWb = null;
                    try { activeWb = _excelApp.ActiveWorkbook; } catch { }

                    if (activeWb != null)
                    {
                        try
                        {
                            string fullName = string.Empty;
                            try { fullName = activeWb.FullName; } catch { }

                            if (!string.IsNullOrEmpty(fullName) && File.Exists(fullName))
                            {
                                AddFilesToGrid(new[] { fullName });

                                // Tự động đặt thư mục xuất là thư mục chứa file đang mở
                                string dir = Path.GetDirectoryName(fullName) ?? string.Empty;
                                if (!string.IsNullOrEmpty(dir) && Directory.Exists(dir))
                                {
                                    TxtOutputDir.Text = dir;
                                }
                            }
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(activeWb);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"LoadActiveWorkbookIfAvailable error: {ex.Message}");
            }
        }

        private void OnModeChanged(object sender, RoutedEventArgs e)
        {
            UpdateOptionsVisibility();
        }

        private void OnTargetFormatChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            UpdateOptionsVisibility();
        }

        private void UpdateOptionsVisibility()
        {
            if (PanelFormatOption == null || PanelMarkdownOptions == null || PanelMergeFileName == null || TxtModeHelp == null) return;

            if (RbModeConvert.IsChecked == true)
            {
                PanelFormatOption.Visibility = Visibility.Visible;
                PanelMergeFileName.Visibility = Visibility.Collapsed;

                bool isMarkdown = (CboTargetFormat.SelectedItem is ExcelOutputFormat fmt && fmt == ExcelOutputFormat.Markdown);
                PanelMarkdownOptions.Visibility = isMarkdown ? Visibility.Visible : Visibility.Collapsed;
                TxtModeHelp.Text = isMarkdown ? LocalizationService.Get("BFC_TipMarkdown") : LocalizationService.Get("BFC_TipConvert");
            }
            else if (RbModeSplit.IsChecked == true)
            {
                PanelFormatOption.Visibility = Visibility.Collapsed;
                PanelMarkdownOptions.Visibility = Visibility.Collapsed;
                PanelMergeFileName.Visibility = Visibility.Collapsed;
                TxtModeHelp.Text = LocalizationService.Get("BFC_TipSplit");
            }
            else if (RbModeMerge.IsChecked == true)
            {
                PanelFormatOption.Visibility = Visibility.Collapsed;
                PanelMarkdownOptions.Visibility = Visibility.Collapsed;
                PanelMergeFileName.Visibility = Visibility.Visible;
                TxtModeHelp.Text = LocalizationService.Get("BFC_TipMerge");
            }
        }

        private void OnMdSheetFilterModeChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (TxtMdSheetFilter == null || CboMdSheetFilterMode == null) return;
            bool isFilterActive = (CboMdSheetFilterMode.SelectedIndex > 0);
            TxtMdSheetFilter.IsEnabled = isFilterActive;
            if (BtnPickSheets != null) BtnPickSheets.IsEnabled = isFilterActive;
        }

        private List<string> GetCurrentWorkbookSheets()
        {
            var sheetNames = new List<string>();
            if (_excelApp == null) return sheetNames;

            string targetFilePath = string.Empty;
            if (GridFiles.SelectedItem is BatchFileItem selectedItem && !string.IsNullOrEmpty(selectedItem.FilePath))
            {
                targetFilePath = selectedItem.FilePath;
            }
            else if (_fileItems.Count > 0)
            {
                targetFilePath = _fileItems[0].FilePath;
            }

            Workbooks? workbooks = null;
            Workbook? matchedWb = null;
            bool openedTemp = false;

            try
            {
                workbooks = _excelApp.Workbooks;
                int wbCount = workbooks.Count;

                if (!string.IsNullOrEmpty(targetFilePath))
                {
                    for (int i = 1; i <= wbCount; i++)
                    {
                        Workbook? wb = null;
                        try
                        {
                            wb = workbooks[i];
                            if (string.Equals(wb.FullName, targetFilePath, StringComparison.OrdinalIgnoreCase) ||
                                string.Equals(wb.Name, Path.GetFileName(targetFilePath), StringComparison.OrdinalIgnoreCase))
                            {
                                matchedWb = wb;
                                break;
                            }
                        }
                        catch { }
                        finally
                        {
                            if (wb != null && matchedWb != wb)
                            {
                                Marshal.ReleaseComObject(wb);
                            }
                        }
                    }
                }

                if (matchedWb == null)
                {
                    try { matchedWb = _excelApp.ActiveWorkbook; } catch { }
                }

                if (matchedWb == null && !string.IsNullOrEmpty(targetFilePath) && File.Exists(targetFilePath))
                {
                    try
                    {
                        matchedWb = workbooks.Open(targetFilePath, ReadOnly: true, UpdateLinks: 0);
                        openedTemp = (matchedWb != null);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error opening temp wb in picker: {ex.Message}");
                    }
                }

                if (matchedWb != null)
                {
                    Sheets? sheets = null;
                    try
                    {
                        sheets = matchedWb.Worksheets;
                        int sCount = sheets.Count;
                        for (int s = 1; s <= sCount; s++)
                        {
                            _Worksheet? ws = null;
                            try
                            {
                                ws = (_Worksheet)sheets[s];
                                sheetNames.Add(ws.Name);
                            }
                            catch { }
                            finally
                            {
                                if (ws != null) Marshal.ReleaseComObject(ws);
                            }
                        }
                    }
                    finally
                    {
                        if (sheets != null) Marshal.ReleaseComObject(sheets);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"GetCurrentWorkbookSheets error: {ex.Message}");
            }
            finally
            {
                if (matchedWb != null)
                {
                    if (openedTemp)
                    {
                        try { matchedWb.Close(SaveChanges: false); } catch { }
                    }
                    Marshal.ReleaseComObject(matchedWb);
                    matchedWb = null;
                }

                if (workbooks != null)
                {
                    Marshal.ReleaseComObject(workbooks);
                    workbooks = null;
                }

                if (openedTemp)
                {
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }

            return sheetNames;
        }

        private void OnPickSheetsClick(object sender, RoutedEventArgs e)
        {
            var sheets = GetCurrentWorkbookSheets();
            if (sheets.Count == 0)
            {
                WpfMessageBox.Show(this, LocalizationService.Get("BFC_NoSheetsFound"), LocalizationService.Get("Common_Notice"), MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var currentSelected = new HashSet<string>(
                TxtMdSheetFilter.Text.Split(new[] { ',', ';', '|' }, StringSplitOptions.RemoveEmptyEntries)
                                    .Select(s => s.Trim()),
                StringComparer.OrdinalIgnoreCase);

            PanelPickerSheetList.Children.Clear();

            foreach (var sheet in sheets)
            {
                var chk = new System.Windows.Controls.CheckBox
                {
                    Content = sheet,
                    IsChecked = currentSelected.Contains(sheet),
                    Margin = new Thickness(0, 3, 0, 3),
                    Cursor = System.Windows.Input.Cursors.Hand
                };
                PanelPickerSheetList.Children.Add(chk);
            }

            PopupSheetPicker.IsOpen = true;
        }

        private void OnSelectAllPickerSheetsClick(object sender, RoutedEventArgs e)
        {
            foreach (var child in PanelPickerSheetList.Children)
            {
                if (child is System.Windows.Controls.CheckBox chk) chk.IsChecked = true;
            }
        }

        private void OnClearAllPickerSheetsClick(object sender, RoutedEventArgs e)
        {
            foreach (var child in PanelPickerSheetList.Children)
            {
                if (child is System.Windows.Controls.CheckBox chk) chk.IsChecked = false;
            }
        }

        private void OnApplyPickerSheetsClick(object sender, RoutedEventArgs e)
        {
            var selected = new List<string>();
            foreach (var child in PanelPickerSheetList.Children)
            {
                if (child is System.Windows.Controls.CheckBox chk && chk.IsChecked == true)
                {
                    selected.Add(chk.Content?.ToString() ?? string.Empty);
                }
            }

            TxtMdSheetFilter.Text = string.Join(", ", selected.Where(s => !string.IsNullOrEmpty(s)));
            PopupSheetPicker.IsOpen = false;
        }

        private void OnNumericPreviewTextInput(object sender, System.Windows.Input.TextCompositionEventArgs e)
        {
            e.Handled = !int.TryParse(e.Text, out _);
        }

        private void OnStartRowPreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Space)
            {
                e.Handled = true;
            }
        }

        private void OnStartRowLostFocus(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(TxtMdStartRow.Text) || !int.TryParse(TxtMdStartRow.Text.Trim(), out int val) || val < 1)
            {
                TxtMdStartRow.Text = "3";
            }
        }

        private void OnAddFilesClick(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Title = "Select Excel Files",
                Filter = "Excel Files (*.xlsx;*.xls;*.xlsb;*.xlsm;*.csv)|*.xlsx;*.xls;*.xlsb;*.xlsm;*.csv|All Files (*.*)|*.*",
                Multiselect = true
            };

            if (dlg.ShowDialog() == true)
            {
                AddFilesToGrid(dlg.FileNames);
            }
        }

        private void OnAddFolderClick(object sender, RoutedEventArgs e)
        {
            using (var dlg = new FolderBrowserDialog())
            {
                dlg.Description = "Select Folder";
                dlg.ShowNewFolderButton = false;

                if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    var extensions = new[] { ".xlsx", ".xls", ".xlsb", ".xlsm", ".csv" };
                    try
                    {
                        var files = Directory.GetFiles(dlg.SelectedPath, "*.*", SearchOption.TopDirectoryOnly)
                            .Where(f => extensions.Contains(Path.GetExtension(f).ToLowerInvariant()))
                            .ToArray();

                        AddFilesToGrid(files);
                    }
                    catch (Exception ex)
                    {
                        WpfMessageBox.Show($"Không thể quét thư mục:\n{ex.Message}", "ExcelSupport", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
            }
        }

        private void AddFilesToGrid(string[] filePaths)
        {
            var existing = new HashSet<string>(_fileItems.Select(f => f.FilePath), StringComparer.OrdinalIgnoreCase);

            foreach (var path in filePaths)
            {
                if (!File.Exists(path) || existing.Contains(path)) continue;

                var fi = new FileInfo(path);
                _fileItems.Add(new BatchFileItem
                {
                    FilePath = path,
                    FileName = fi.Name,
                    FileSize = FormatFileSize(fi.Length),
                    Status = LocalizationService.Get("Common_Ready")
                });
                existing.Add(path);
            }

            TxtFileCountBadge.Text = LocalizationService.Get("BFC_FileCountBadge", _fileItems.Count);
            GridFiles.ItemsSource = null;
            GridFiles.ItemsSource = _fileItems;
        }

        private static string FormatFileSize(long bytes)
        {
            if (bytes >= 1024 * 1024) return $"{bytes / (1024.0 * 1024.0):0.#} MB";
            if (bytes >= 1024) return $"{bytes / 1024.0:0.#} KB";
            return $"{bytes} B";
        }

        private void OnClearFilesClick(object sender, RoutedEventArgs e)
        {
            _fileItems.Clear();
            TxtFileCountBadge.Text = LocalizationService.Get("BFC_FileCountBadge", 0);
            GridFiles.ItemsSource = null;
        }

        private void OnBrowseOutputDirClick(object sender, RoutedEventArgs e)
        {
            using (var dlg = new FolderBrowserDialog())
            {
                dlg.Description = "Chọn thư mục lưu các file kết quả";
                dlg.ShowNewFolderButton = true;

                if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    TxtOutputDir.Text = dlg.SelectedPath;
                }
            }
        }

        private void OnExecuteBatchClick(object sender, RoutedEventArgs e)
        {
            if (_excelApp == null || _fileItems.Count == 0)
            {
                WpfMessageBox.Show(this, "Vui lòng thêm ít nhất một tập tin vào danh sách cần xử lý.", "Thông Báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string outDir = TxtOutputDir.Text.Trim();
            if (string.IsNullOrEmpty(outDir))
            {
                WpfMessageBox.Show(this, "Vui lòng chọn thư mục lưu kết quả.", "Thông Báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var mode = BatchConvertMode.ConvertFormat;
            if (RbModeSplit.IsChecked == true) mode = BatchConvertMode.SplitSheetsToFiles;
            else if (RbModeMerge.IsChecked == true) mode = BatchConvertMode.MergeFilesToOne;

            var targetFormat = (ExcelOutputFormat)(CboTargetFormat.SelectedItem ?? ExcelOutputFormat.PDF);

            var mdMode = (RbMdSeparateFiles.IsChecked == true)
                ? MarkdownSheetExportMode.SeparateFilePerSheet
                : MarkdownSheetExportMode.SingleFileWithHeadings;

            var filterMode = CboMdSheetFilterMode.SelectedIndex switch
            {
                1 => MarkdownSheetFilterMode.IncludeOnly,
                2 => MarkdownSheetFilterMode.Exclude,
                _ => MarkdownSheetFilterMode.All
            };

            int startRow = 3;
            if (int.TryParse(TxtMdStartRow.Text.Trim(), out int r) && r >= 1)
            {
                startRow = r;
            }

            var options = new BatchConvertOptions
            {
                Mode = mode,
                InputFiles = _fileItems.Select(f => f.FilePath).ToList(),
                OutputDirectory = outDir,
                TargetFormat = targetFormat,
                OverwriteExisting = (ChkOverwrite.IsChecked == true),
                MergedFileName = !string.IsNullOrWhiteSpace(TxtMergedFileName.Text) ? TxtMergedFileName.Text.Trim() : "Gop_Cac_File_Excel.xlsx",
                MarkdownMode = mdMode,
                SheetFilterMode = filterMode,
                SheetFilterPatterns = TxtMdSheetFilter.Text.Trim(),
                StartRow = startRow,
                IncludeMarkdownToc = (ChkMdToc.IsChecked == true),
                ConvertLineBreaksToBr = (ChkMdLineBreaks.IsChecked == true)
            };

            ProgressBarConvert.Visibility = Visibility.Visible;
            ProgressBarConvert.Value = 0;
            TxtFooterStatus.Text = "⏳ Đang thực thi xử lý file hàng loạt...";
            BtnExecute.IsEnabled = false;
            BtnClose.IsEnabled = false;

            AllowUiThreadToRender();

            try
            {
                var result = BatchFileConverterService.ExecuteBatchConversion(_excelApp, options, (current, total, fileName) =>
                {
                    ProgressBarConvert.Value = (double)current / total * 100.0;
                    TxtFooterStatus.Text = $"⏳ Đang xử lý ({current}/{total}): {fileName}...";
                    AllowUiThreadToRender();
                });

                ProgressBarConvert.Visibility = Visibility.Collapsed;
                TxtFooterStatus.Text = $"✅ {result.Message}";

                if (result.Success)
                {
                    var openFolder = WpfMessageBox.Show(
                        this,
                        $"{result.Message}\n\n• Thành công: {result.SuccessCount:N0} file\n• Thất bại: {result.FailCount:N0} file\n\nBạn có muốn mở thư mục kết quả ngay không?",
                        "Xử Lý Hoàn Tất",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Information);

                    if (openFolder == MessageBoxResult.Yes)
                    {
                        try { System.Diagnostics.Process.Start("explorer.exe", outDir); } catch { }
                    }
                }
                else
                {
                    WpfMessageBox.Show(this, result.Message, "Thông Báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                ProgressBarConvert.Visibility = Visibility.Collapsed;
                WpfMessageBox.Show(this, $"Lỗi xử lý file:\n{ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                TxtFooterStatus.Text = "❌ Đã xảy ra lỗi trong quá trình xử lý file.";
            }
            finally
            {
                BtnExecute.IsEnabled = true;
                BtnClose.IsEnabled = true;
            }
        }

        private void AllowUiThreadToRender()
        {
            try
            {
                Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Render, new Action(() => { }));
            }
            catch { }
        }

        private void OnCloseClick(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
