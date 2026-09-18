using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Forms;
using ExcelSupport.ViewModels;
using MessageBox = System.Windows.MessageBox;

namespace ExcelSupport.Views
{
    public class SheetCheckItem : ViewModelBase
    {
        private bool _isChecked = true;
        public string SheetName { get; set; } = string.Empty;
        public bool IsChecked
        {
            get => _isChecked;
            set => SetProperty(ref _isChecked, value);
        }
    }

    public partial class SheetToolsDialog : Window
    {
        private readonly WorkbookNodeViewModel _workbook;
        private readonly ObservableCollection<SheetCheckItem> _splitItems = new ObservableCollection<SheetCheckItem>();
        private readonly ObservableCollection<SheetCheckItem> _mergeItems = new ObservableCollection<SheetCheckItem>();
        private readonly ObservableCollection<string> _importFiles = new ObservableCollection<string>();

        public SheetToolsDialog(WorkbookNodeViewModel workbook, int initialTabIndex = 0, bool isDarkTheme = false)
        {
            InitializeComponent();
            _workbook = workbook;

            TxtWorkbookTitle.Text = $"Workbook: {_workbook.WorkbookName} ({_workbook.SheetCount} sheets)";

            // Load danh sách Sheet
            foreach (var ws in _workbook.Worksheets)
            {
                _splitItems.Add(new SheetCheckItem { SheetName = ws.SheetName, IsChecked = true });
                _mergeItems.Add(new SheetCheckItem { SheetName = ws.SheetName, IsChecked = true });
            }

            LstSplitSheets.ItemsSource = _splitItems;
            LstMergeSheets.ItemsSource = _mergeItems;
            LstImportFiles.ItemsSource = _importFiles;

            // Default output folder: thư mục chứa file Excel hoặc Desktop
            string defaultFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            if (!string.IsNullOrEmpty(_workbook.FilePath) && File.Exists(_workbook.FilePath))
            {
                defaultFolder = Path.GetDirectoryName(_workbook.FilePath) ?? defaultFolder;
            }
            TxtSplitOutputFolder.Text = defaultFolder;

            MainTabControl.SelectedIndex = initialTabIndex;
        }

        #region Tab 1: Split Sheets

        private void OnSplitSelectAllClick(object sender, RoutedEventArgs e)
        {
            foreach (var item in _splitItems) item.IsChecked = true;
        }

        private void OnSplitUnselectAllClick(object sender, RoutedEventArgs e)
        {
            foreach (var item in _splitItems) item.IsChecked = false;
        }

        private void OnBrowseSplitFolderClick(object sender, RoutedEventArgs e)
        {
            using (var dlg = new FolderBrowserDialog())
            {
                dlg.Description = "Chọn thư mục lưu các file Excel tách ra";
                dlg.SelectedPath = TxtSplitOutputFolder.Text;
                if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    TxtSplitOutputFolder.Text = dlg.SelectedPath;
                }
            }
        }

        private void OnExecuteSplitClick(object sender, RoutedEventArgs e)
        {
            var selectedSheets = _splitItems.Where(i => i.IsChecked).Select(i => i.SheetName).ToList();
            if (selectedSheets.Count == 0)
            {
                MessageBox.Show("Vui lòng chọn ít nhất 1 sheet để tách.", "Tách Sheets", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string folder = TxtSplitOutputFolder.Text.Trim();
            if (string.IsNullOrEmpty(folder))
            {
                MessageBox.Show("Vui lòng chọn thư mục lưu file.", "Tách Sheets", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string wbName = _workbook.WorkbookName;
            bool keepOriginal = ChkKeepOriginalSheets.IsChecked == true;
            Close();

            ExcelDna.Integration.ExcelAsyncUtil.QueueAsMacro(() =>
            {
                AddInEvents.Instance?.SplitWorksheetsToFiles(wbName, selectedSheets, folder, keepOriginal);
            });
        }

        #endregion

        #region Tab 2: Merge Sheets

        private void OnMergeSelectAllClick(object sender, RoutedEventArgs e)
        {
            foreach (var item in _mergeItems) item.IsChecked = true;
        }

        private void OnMergeUnselectAllClick(object sender, RoutedEventArgs e)
        {
            foreach (var item in _mergeItems) item.IsChecked = false;
        }

        private void OnExecuteMergeClick(object sender, RoutedEventArgs e)
        {
            var selectedSheets = _mergeItems.Where(i => i.IsChecked).Select(i => i.SheetName).ToList();
            if (selectedSheets.Count == 0)
            {
                MessageBox.Show("Vui lòng chọn ít nhất 1 sheet để gộp.", "Gộp Sheets", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            bool hasHeader = RadMergeSkipSubHeaders.IsChecked == true;
            string wbName = _workbook.WorkbookName;
            Close();

            ExcelDna.Integration.ExcelAsyncUtil.QueueAsMacro(() =>
            {
                AddInEvents.Instance?.ConsolidateSheetsData(wbName, selectedSheets, hasHeader);
            });
        }

        #endregion

        #region Tab 3: Import Files

        private void OnBrowseImportFilesClick(object sender, RoutedEventArgs e)
        {
            using (var dlg = new OpenFileDialog())
            {
                dlg.Title = "Chọn các file Excel để nhập sheet";
                dlg.Filter = "Excel Files (*.xlsx;*.xls;*.xlsm;*.csv)|*.xlsx;*.xls;*.xlsm;*.csv|All Files (*.*)|*.*";
                dlg.Multiselect = true;

                if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    foreach (var f in dlg.FileNames)
                    {
                        if (!_importFiles.Contains(f))
                        {
                            _importFiles.Add(f);
                        }
                    }
                }
            }
        }

        private void OnExecuteImportClick(object sender, RoutedEventArgs e)
        {
            if (_importFiles.Count == 0)
            {
                MessageBox.Show("Vui lòng chọn ít nhất một file Excel để nhập.", "Nhập File", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string wbName = _workbook.WorkbookName;
            string[] files = _importFiles.ToArray();
            Close();

            ExcelDna.Integration.ExcelAsyncUtil.QueueAsMacro(() =>
            {
                AddInEvents.Instance?.ImportSheetsFromExternalFiles(wbName, files);
            });
        }

        #endregion

        #region Tab 4: Batch Rename

        private void OnExecuteBatchRenameClick(object sender, RoutedEventArgs e)
        {
            string prefix = TxtPrefix.Text ?? string.Empty;
            string suffix = TxtSuffix.Text ?? string.Empty;
            string find = TxtFindText.Text ?? string.Empty;
            string replace = TxtReplaceText.Text ?? string.Empty;

            if (string.IsNullOrEmpty(prefix) && string.IsNullOrEmpty(suffix) && string.IsNullOrEmpty(find))
            {
                MessageBox.Show("Vui lòng nhập ít nhất một quy tắc đổi tên (Tiền tố, Hậu tố hoặc Tìm & Thay thế).",
                                "Đổi Tên Hàng Loạt", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string wbName = _workbook.WorkbookName;
            Close();

            ExcelDna.Integration.ExcelAsyncUtil.QueueAsMacro(() =>
            {
                AddInEvents.Instance?.BatchRenameWorksheets(wbName, prefix, suffix, find, replace);
            });
        }

        #endregion

        private void OnCloseClick(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
