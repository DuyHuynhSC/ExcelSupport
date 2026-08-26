using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Input;
using ExcelSupport.Services;
using Microsoft.Win32;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using WpfMessageBox = System.Windows.MessageBox;
using WpfMessageBoxButton = System.Windows.MessageBoxButton;
using WpfMessageBoxImage = System.Windows.MessageBoxImage;

namespace ExcelSupport.ViewModels
{
    public class DesignPageCounterViewModel : ViewModelBase
    {
        private readonly ExcelApp _excelApp;

        private string _selectedTargetWorkbook = string.Empty;
        private string _selectedTemplateWorkbook = string.Empty;
        private string _customTargetFilePath = string.Empty;
        private string _customTemplateFilePath = string.Empty;

        private bool _ignoreCoverAndHistory = true;
        private bool _ignoreBlankPages = true;
        private bool _countShapesAndPictures = true;
        private int _minChangedCellsThreshold = 2;

        private bool _isAnalyzing;
        private string _progressMessage = string.Empty;
        private int _progressPercentage;

        private WorkbookPageCounterResult? _result;
        private SheetPageCounterResult? _selectedSheetResult;

        public ObservableCollection<string> OpenWorkbooks { get; } = new();

        public string SelectedTargetWorkbook
        {
            get => _selectedTargetWorkbook;
            set
            {
                if (SetProperty(ref _selectedTargetWorkbook, value))
                {
                    CustomTargetFilePath = string.Empty;
                }
            }
        }

        public string SelectedTemplateWorkbook
        {
            get => _selectedTemplateWorkbook;
            set
            {
                if (SetProperty(ref _selectedTemplateWorkbook, value))
                {
                    CustomTemplateFilePath = string.Empty;
                }
            }
        }

        public string CustomTargetFilePath
        {
            get => _customTargetFilePath;
            set => SetProperty(ref _customTargetFilePath, value);
        }

        public string CustomTemplateFilePath
        {
            get => _customTemplateFilePath;
            set => SetProperty(ref _customTemplateFilePath, value);
        }

        public bool IgnoreCoverAndHistory
        {
            get => _ignoreCoverAndHistory;
            set => SetProperty(ref _ignoreCoverAndHistory, value);
        }

        public bool IgnoreBlankPages
        {
            get => _ignoreBlankPages;
            set => SetProperty(ref _ignoreBlankPages, value);
        }

        public bool CountShapesAndPictures
        {
            get => _countShapesAndPictures;
            set => SetProperty(ref _countShapesAndPictures, value);
        }

        public int MinChangedCellsThreshold
        {
            get => _minChangedCellsThreshold;
            set => SetProperty(ref _minChangedCellsThreshold, value);
        }

        public bool IsAnalyzing
        {
            get => _isAnalyzing;
            set
            {
                if (SetProperty(ref _isAnalyzing, value))
                {
                    OnPropertyChanged(nameof(CanAnalyze));
                }
            }
        }

        public bool CanAnalyze => !IsAnalyzing;

        public string ProgressMessage
        {
            get => _progressMessage;
            set => SetProperty(ref _progressMessage, value);
        }

        public int ProgressPercentage
        {
            get => _progressPercentage;
            set => SetProperty(ref _progressPercentage, value);
        }

        public WorkbookPageCounterResult? Result
        {
            get => _result;
            set
            {
                if (SetProperty(ref _result, value))
                {
                    OnPropertyChanged(nameof(HasResult));
                }
            }
        }

        public bool HasResult => Result != null;

        public SheetPageCounterResult? SelectedSheetResult
        {
            get => _selectedSheetResult;
            set => SetProperty(ref _selectedSheetResult, value);
        }

        public ICommand RefreshWorkbooksCommand { get; }
        public ICommand BrowseTargetFileCommand { get; }
        public ICommand BrowseTemplateFileCommand { get; }
        public ICommand AnalyzeCommand { get; }
        public ICommand ExportReportCommand { get; }

        public DesignPageCounterViewModel(ExcelApp excelApp)
        {
            _excelApp = excelApp ?? throw new ArgumentNullException(nameof(excelApp));

            RefreshWorkbooksCommand = new RelayCommand(_ => RefreshWorkbooks());
            BrowseTargetFileCommand = new RelayCommand(_ => BrowseTargetFile());
            BrowseTemplateFileCommand = new RelayCommand(_ => BrowseTemplateFile());
            AnalyzeCommand = new RelayCommand(async _ => await AnalyzeAsync(), _ => CanAnalyze);
            ExportReportCommand = new RelayCommand(_ => ExportReport(), _ => HasResult && !IsAnalyzing);

            RefreshWorkbooks();
        }

        public void RefreshWorkbooks()
        {
            OpenWorkbooks.Clear();
            try
            {
                if (_excelApp.Workbooks != null)
                {
                    foreach (Microsoft.Office.Interop.Excel.Workbook wb in _excelApp.Workbooks)
                    {
                        OpenWorkbooks.Add(wb.Name);
                    }
                }
            }
            catch { }

            if (OpenWorkbooks.Count > 0)
            {
                if (string.IsNullOrEmpty(SelectedTargetWorkbook) || !OpenWorkbooks.Contains(SelectedTargetWorkbook))
                {
                    SelectedTargetWorkbook = OpenWorkbooks[0];
                }
            }
        }

        private void BrowseTargetFile()
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Excel Files (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls|All Files (*.*)|*.*",
                Title = "Chọn file thiết kế cần đếm số trang"
            };

            if (dlg.ShowDialog() == true)
            {
                CustomTargetFilePath = dlg.FileName;
                SelectedTargetWorkbook = Path.GetFileName(dlg.FileName);
            }
        }

        private void BrowseTemplateFile()
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Excel Files (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls|All Files (*.*)|*.*",
                Title = "Chọn file template gốc để đối chiếu"
            };

            if (dlg.ShowDialog() == true)
            {
                CustomTemplateFilePath = dlg.FileName;
                SelectedTemplateWorkbook = Path.GetFileName(dlg.FileName);
            }
        }

        public async Task AnalyzeAsync()
        {
            string targetSource = !string.IsNullOrEmpty(CustomTargetFilePath) ? CustomTargetFilePath : SelectedTargetWorkbook;
            string templateSource = !string.IsNullOrEmpty(CustomTemplateFilePath) ? CustomTemplateFilePath : SelectedTemplateWorkbook;

            if (string.IsNullOrWhiteSpace(targetSource))
            {
                WpfMessageBox.Show("Vui lòng chọn file thiết kế cần phân tích.", "Thiếu thông tin", WpfMessageBoxButton.OK, WpfMessageBoxImage.Warning);
                return;
            }

            IsAnalyzing = true;
            ProgressMessage = "Đang khởi tạo phân tích...";
            ProgressPercentage = 0;

            var options = new PageCounterOptions
            {
                IgnoreCoverAndHistory = IgnoreCoverAndHistory,
                IgnoreBlankPages = IgnoreBlankPages,
                CountShapesAndPictures = CountShapesAndPictures,
                MinChangedCellsThreshold = MinChangedCellsThreshold
            };

            try
            {
                var analysisResult = await Task.Run(() =>
                {
                    return DesignPageCounterService.AnalyzePages(
                        _excelApp,
                        targetSource,
                        templateSource,
                        options,
                        (msg, p) =>
                        {
                            System.Windows.Application.Current?.Dispatcher.Invoke(() =>
                            {
                                ProgressMessage = msg;
                                ProgressPercentage = p;
                            });
                        });
                });

                Result = analysisResult;
                if (Result.SheetResults.Count > 0)
                {
                    SelectedSheetResult = Result.SheetResults[0];
                }
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Đã xảy ra lỗi khi đếm số trang thiết kế:\n{ex.Message}", "Lỗi phân tích", WpfMessageBoxButton.OK, WpfMessageBoxImage.Error);
            }
            finally
            {
                IsAnalyzing = false;
                ProgressMessage = string.Empty;
            }
        }

        private void ExportReport()
        {
            if (Result == null) return;

            try
            {
                Microsoft.Office.Interop.Excel.Workbook? targetWb = null;
                string targetSource = !string.IsNullOrEmpty(CustomTargetFilePath) ? CustomTargetFilePath : SelectedTargetWorkbook;

                foreach (Microsoft.Office.Interop.Excel.Workbook wb in _excelApp.Workbooks)
                {
                    if (string.Equals(wb.Name, targetSource, StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(wb.FullName, targetSource, StringComparison.OrdinalIgnoreCase))
                    {
                        targetWb = wb;
                        break;
                    }
                }

                if (targetWb == null && _excelApp.ActiveWorkbook != null)
                {
                    targetWb = _excelApp.ActiveWorkbook;
                }

                if (targetWb == null)
                {
                    WpfMessageBox.Show("Không tìm thấy Workbook để xuất báo cáo.", "Lỗi xuất", WpfMessageBoxButton.OK, WpfMessageBoxImage.Warning);
                    return;
                }

                bool ok = DesignPageCounterService.ExportReportToExcel(_excelApp, targetWb, Result);
                if (ok)
                {
                    WpfMessageBox.Show("Đã xuất báo cáo thống kê số trang thiết kế ra Sheet mới thành công!", "Xuất thành công", WpfMessageBoxButton.OK, WpfMessageBoxImage.Information);
                }
                else
                {
                    WpfMessageBox.Show("Có lỗi khi tạo Sheet báo cáo.", "Lỗi xuất", WpfMessageBoxButton.OK, WpfMessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Lỗi xuất báo cáo: {ex.Message}", "Lỗi", WpfMessageBoxButton.OK, WpfMessageBoxImage.Error);
            }
        }
    }
}
