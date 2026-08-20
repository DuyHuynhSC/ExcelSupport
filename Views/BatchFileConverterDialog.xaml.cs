using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using ExcelSupport.Host;
using ExcelSupport.Models;
using ExcelSupport.Services;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using WpfMessageBox = System.Windows.MessageBox;

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
        }

        private void OnModeChanged(object sender, RoutedEventArgs e)
        {
            if (PanelFormatOption == null || PanelMergeFileName == null || TxtModeHelp == null) return;

            if (RbModeConvert.IsChecked == true)
            {
                PanelFormatOption.Visibility = Visibility.Visible;
                PanelMergeFileName.Visibility = Visibility.Collapsed;
                TxtModeHelp.Text = "💡 Chuyển đổi định dạng siêu tốc hàng loạt tập tin sang XLSX, XLS, XLSB, CSV, PDF.";
            }
            else if (RbModeSplit.IsChecked == true)
            {
                PanelFormatOption.Visibility = Visibility.Collapsed;
                PanelMergeFileName.Visibility = Visibility.Collapsed;
                TxtModeHelp.Text = "💡 Tách từng Sheet trong mỗi file Excel được chọn thành từng file .xlsx độc lập.";
            }
            else if (RbModeMerge.IsChecked == true)
            {
                PanelFormatOption.Visibility = Visibility.Collapsed;
                PanelMergeFileName.Visibility = Visibility.Visible;
                TxtModeHelp.Text = "💡 Gom toàn bộ các file Excel trong danh sách thành 1 file duy nhất (mỗi file nguồn tương ứng một Sheet).";
            }
        }

        private void OnAddFilesClick(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Title = "Chọn các tập tin Excel cần xử lý",
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
                dlg.Description = "Chọn thư mục chứa các file Excel cần xử lý";
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
                        WpfMessageBox.Show($"Không thể quét thư mục:\n{ex.Message}", "Thông Báo", MessageBoxButton.OK, MessageBoxImage.Warning);
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
                    Status = "Sẵn sàng"
                });
                existing.Add(path);
            }

            TxtFileCountBadge.Text = $"{_fileItems.Count:N0} file";
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
            TxtFileCountBadge.Text = "0 file";
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

        private async void OnExecuteBatchClick(object sender, RoutedEventArgs e)
        {
            if (_excelApp == null || _fileItems.Count == 0)
            {
                WpfMessageBox.Show("Vui lòng thêm ít nhất một tập tin vào danh sách cần xử lý.", "Thông Báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string outDir = TxtOutputDir.Text.Trim();
            if (string.IsNullOrEmpty(outDir))
            {
                WpfMessageBox.Show("Vui lòng chọn thư mục lưu kết quả.", "Thông Báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var mode = BatchConvertMode.ConvertFormat;
            if (RbModeSplit.IsChecked == true) mode = BatchConvertMode.SplitSheetsToFiles;
            else if (RbModeMerge.IsChecked == true) mode = BatchConvertMode.MergeFilesToOne;

            var targetFormat = (ExcelOutputFormat)(CboTargetFormat.SelectedItem ?? ExcelOutputFormat.PDF);

            var options = new BatchConvertOptions
            {
                Mode = mode,
                InputFiles = _fileItems.Select(f => f.FilePath).ToList(),
                OutputDirectory = outDir,
                TargetFormat = targetFormat,
                OverwriteExisting = (ChkOverwrite.IsChecked == true),
                MergedFileName = !string.IsNullOrWhiteSpace(TxtMergedFileName.Text) ? TxtMergedFileName.Text.Trim() : "Gop_Cac_File_Excel.xlsx"
            };

            ProgressBarConvert.Visibility = Visibility.Visible;
            ProgressBarConvert.Value = 0;
            TxtFooterStatus.Text = "⏳ Đang thực thi xử lý file hàng loạt...";

            try
            {
                var result = await Task.Run(() =>
                {
                    return BatchFileConverterService.ExecuteBatchConversion(_excelApp, options, (current, total, fileName) =>
                    {
                        Dispatcher.Invoke(() =>
                        {
                            ProgressBarConvert.Value = (double)current / total * 100.0;
                            TxtFooterStatus.Text = $"⏳ Đang xử lý ({current}/{total}): {fileName}...";
                        });
                    });
                });

                ProgressBarConvert.Visibility = Visibility.Collapsed;
                TxtFooterStatus.Text = $"✅ {result.Message}";

                if (result.Success)
                {
                    var openFolder = WpfMessageBox.Show(
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
                    WpfMessageBox.Show(result.Message, "Thông Báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                ProgressBarConvert.Visibility = Visibility.Collapsed;
                WpfMessageBox.Show($"Lỗi xử lý file:\n{ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                TxtFooterStatus.Text = "❌ Đã xảy ra lỗi trong quá trình xử lý file.";
            }
        }

        private void OnCloseClick(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
