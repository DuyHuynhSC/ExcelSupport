using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Input;
using ExcelSupport.Services;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using WpfMessageBox = System.Windows.MessageBox;
using WpfMessageBoxButton = System.Windows.MessageBoxButton;
using WpfMessageBoxImage = System.Windows.MessageBoxImage;

namespace ExcelSupport.ViewModels
{
    public class SheetSelectorItem : ViewModelBase
    {
        private string _sheetName = string.Empty;
        private bool _isIncluded = true;

        public string SheetName
        {
            get => _sheetName;
            set => SetProperty(ref _sheetName, value);
        }

        public bool IsIncluded
        {
            get => _isIncluded;
            set => SetProperty(ref _isIncluded, value);
        }
    }

    public class ColorSwatchItem
    {
        public string Name { get; set; } = string.Empty;
        public string Hex { get; set; } = string.Empty;
    }

    public class DensityPresetItem
    {
        public string Name { get; set; } = string.Empty;
        public int Value { get; set; }
    }

    public class DesignPageCounterViewModel : ViewModelBase
    {
        private readonly ExcelApp _excelApp;

        private string _selectedTargetWorkbook = string.Empty;
        private string _selectedTemplateWorkbook = string.Empty;
        private string _customTargetFilePath = string.Empty;
        private string _customTemplateFilePath = string.Empty;

        private PageCounterMode _mode = PageCounterMode.CharacterAndHighlight;
        private int _charactersPerPage = 600;
        private bool _highlightChangedCells = true;
        private ColorSwatchItem _selectedHighlightColor;
        private DensityPresetItem _selectedDensityPreset;

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
        public ObservableCollection<SheetSelectorItem> AvailableSheets { get; } = new();
        public ObservableCollection<ColorSwatchItem> AvailableHighlightColors { get; } = new();
        public ObservableCollection<DensityPresetItem> AvailableDensityPresets { get; } = new();

        public PageCounterMode Mode
        {
            get => _mode;
            set
            {
                if (SetProperty(ref _mode, value))
                {
                    OnPropertyChanged(nameof(IsCharHighlightMode));
                    OnPropertyChanged(nameof(IsPrintBreakMode));
                }
            }
        }

        public bool IsCharHighlightMode
        {
            get => Mode == PageCounterMode.CharacterAndHighlight;
            set
            {
                if (value) Mode = PageCounterMode.CharacterAndHighlight;
            }
        }

        public bool IsPrintBreakMode
        {
            get => Mode == PageCounterMode.PrintBreakGrid;
            set
            {
                if (value) Mode = PageCounterMode.PrintBreakGrid;
            }
        }

        public int CharactersPerPage
        {
            get => _charactersPerPage;
            set => SetProperty(ref _charactersPerPage, value);
        }

        public bool HighlightChangedCells
        {
            get => _highlightChangedCells;
            set => SetProperty(ref _highlightChangedCells, value);
        }

        public ColorSwatchItem SelectedHighlightColor
        {
            get => _selectedHighlightColor;
            set => SetProperty(ref _selectedHighlightColor, value);
        }

        public DensityPresetItem SelectedDensityPreset
        {
            get => _selectedDensityPreset;
            set
            {
                if (SetProperty(ref _selectedDensityPreset, value) && value != null)
                {
                    CharactersPerPage = value.Value;
                }
            }
        }

        public string SelectedTargetWorkbook
        {
            get => _selectedTargetWorkbook;
            set
            {
                if (SetProperty(ref _selectedTargetWorkbook, value))
                {
                    if (!string.IsNullOrEmpty(value) && File.Exists(value))
                    {
                        CustomTargetFilePath = value;
                    }
                    else
                    {
                        CustomTargetFilePath = string.Empty;
                    }
                    LoadTargetSheets();
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
                    if (!string.IsNullOrEmpty(value) && File.Exists(value))
                    {
                        CustomTemplateFilePath = value;
                    }
                    else
                    {
                        CustomTemplateFilePath = string.Empty;
                    }
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

        public string SheetSelectionHeader
        {
            get
            {
                int total = AvailableSheets.Count;
                int included = AvailableSheets.Count(s => s.IsIncluded);
                if (total == 0) return "📑 Tùy chọn Sheet cần đếm (Đang tải...)";
                return $"📑 Tùy chọn Sheet cần đếm (Đã chọn {included}/{total} sheet)";
            }
        }

        public bool IgnoreCoverAndHistory
        {
            get => _ignoreCoverAndHistory;
            set
            {
                if (SetProperty(ref _ignoreCoverAndHistory, value))
                {
                    foreach (var item in AvailableSheets)
                    {
                        if (DesignPageCounterService.IsCoverOrHistorySheet(item.SheetName))
                        {
                            item.IsIncluded = !value;
                        }
                    }
                    OnPropertyChanged(nameof(SheetSelectionHeader));
                }
            }
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
                    OnPropertyChanged(nameof(HasEvidenceFile));
                }
            }
        }

        public bool HasResult => Result != null;
        public bool HasEvidenceFile => !string.IsNullOrEmpty(Result?.HighlightedClonedWorkbookPath) && File.Exists(Result!.HighlightedClonedWorkbookPath);

        public SheetPageCounterResult? SelectedSheetResult
        {
            get => _selectedSheetResult;
            set => SetProperty(ref _selectedSheetResult, value);
        }

        public ICommand RefreshWorkbooksCommand { get; }
        public ICommand BrowseTargetFileCommand { get; }
        public ICommand BrowseTemplateFileCommand { get; }
        public ICommand SelectAllSheetsCommand { get; }
        public ICommand DeselectAllSheetsCommand { get; }
        public ICommand InvertSheetsSelectionCommand { get; }
        public ICommand AnalyzeCommand { get; }
        public ICommand ExportReportCommand { get; }
        public ICommand OpenEvidenceWorkbookCommand { get; }

        public DesignPageCounterViewModel(ExcelApp excelApp)
        {
            _excelApp = excelApp ?? throw new ArgumentNullException(nameof(excelApp));

            // Presets
            AvailableDensityPresets.Add(new DensityPresetItem { Name = "600 ký tự (Chuẩn Tiếng Nhật / Kanji)", Value = 600 });
            AvailableDensityPresets.Add(new DensityPresetItem { Name = "1.200 ký tự (Tiếng Việt / Tiếng Anh)", Value = 1200 });
            AvailableDensityPresets.Add(new DensityPresetItem { Name = "800 ký tự / trang", Value = 800 });
            AvailableDensityPresets.Add(new DensityPresetItem { Name = "1.500 ký tự / trang", Value = 1500 });
            _selectedDensityPreset = AvailableDensityPresets[0];

            AvailableHighlightColors.Add(new ColorSwatchItem { Name = "Vàng nhạt", Hex = "#FEF08A" });
            AvailableHighlightColors.Add(new ColorSwatchItem { Name = "Xanh ngọc", Hex = "#BAE6FD" });
            AvailableHighlightColors.Add(new ColorSwatchItem { Name = "Xanh lá nhạt", Hex = "#BBF7D0" });
            AvailableHighlightColors.Add(new ColorSwatchItem { Name = "Cam nhạt", Hex = "#FED7AA" });
            _selectedHighlightColor = AvailableHighlightColors[0];

            RefreshWorkbooksCommand = new RelayCommand(_ => RefreshWorkbooks());
            BrowseTargetFileCommand = new RelayCommand(_ => BrowseTargetFile());
            BrowseTemplateFileCommand = new RelayCommand(_ => BrowseTemplateFile());
            SelectAllSheetsCommand = new RelayCommand(_ =>
            {
                foreach (var s in AvailableSheets) s.IsIncluded = true;
                OnPropertyChanged(nameof(SheetSelectionHeader));
            });
            DeselectAllSheetsCommand = new RelayCommand(_ =>
            {
                foreach (var s in AvailableSheets) s.IsIncluded = false;
                OnPropertyChanged(nameof(SheetSelectionHeader));
            });
            InvertSheetsSelectionCommand = new RelayCommand(_ =>
            {
                foreach (var s in AvailableSheets) s.IsIncluded = !s.IsIncluded;
                OnPropertyChanged(nameof(SheetSelectionHeader));
            });
            AnalyzeCommand = new RelayCommand(async _ => await AnalyzeAsync(), _ => CanAnalyze);
            ExportReportCommand = new RelayCommand(_ => ExportReport(), _ => HasResult && !IsAnalyzing);
            OpenEvidenceWorkbookCommand = new RelayCommand(_ => OpenEvidence(), _ => HasEvidenceFile);

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
                else
                {
                    LoadTargetSheets();
                }
            }
        }

        private void BrowseTargetFile()
        {
            try
            {
                using var dlg = new System.Windows.Forms.OpenFileDialog
                {
                    Filter = "Excel Files (*.xlsx;*.xlsm;*.xls;*.xlsb)|*.xlsx;*.xlsm;*.xls;*.xlsb|All Files (*.*)|*.*",
                    Title = "Chọn file thiết kế cần đếm số trang",
                    CheckFileExists = true
                };

                if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    string filePath = dlg.FileName;
                    if (!OpenWorkbooks.Contains(filePath))
                    {
                        OpenWorkbooks.Add(filePath);
                    }
                    CustomTargetFilePath = filePath;
                    _selectedTargetWorkbook = filePath;
                    OnPropertyChanged(nameof(SelectedTargetWorkbook));
                    LoadTargetSheets();
                }
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Không thể mở hộp thoại chọn file:\n{ex.Message}", "Lỗi chọn file", WpfMessageBoxButton.OK, WpfMessageBoxImage.Error);
            }
        }

        private void BrowseTemplateFile()
        {
            try
            {
                using var dlg = new System.Windows.Forms.OpenFileDialog
                {
                    Filter = "Excel Files (*.xlsx;*.xlsm;*.xls;*.xlsb)|*.xlsx;*.xlsm;*.xls;*.xlsb|All Files (*.*)|*.*",
                    Title = "Chọn file template gốc để đối chiếu",
                    CheckFileExists = true
                };

                if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    string filePath = dlg.FileName;
                    if (!OpenWorkbooks.Contains(filePath))
                    {
                        OpenWorkbooks.Add(filePath);
                    }
                    CustomTemplateFilePath = filePath;
                    _selectedTemplateWorkbook = filePath;
                    OnPropertyChanged(nameof(SelectedTemplateWorkbook));
                }
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Không thể mở hộp thoại chọn file:\n{ex.Message}", "Lỗi chọn file", WpfMessageBoxButton.OK, WpfMessageBoxImage.Error);
            }
        }

        public void LoadTargetSheets()
        {
            string targetSource = !string.IsNullOrEmpty(CustomTargetFilePath) ? CustomTargetFilePath : SelectedTargetWorkbook;
            if (string.IsNullOrWhiteSpace(targetSource))
            {
                AvailableSheets.Clear();
                OnPropertyChanged(nameof(SheetSelectionHeader));
                return;
            }

            AvailableSheets.Clear();
            try
            {
                Microsoft.Office.Interop.Excel.Workbook? targetWb = null;
                if (_excelApp.Workbooks != null)
                {
                    foreach (Microsoft.Office.Interop.Excel.Workbook wb in _excelApp.Workbooks)
                    {
                        if (string.Equals(wb.Name, targetSource, StringComparison.OrdinalIgnoreCase) ||
                            string.Equals(wb.FullName, targetSource, StringComparison.OrdinalIgnoreCase) ||
                            string.Equals(Path.GetFileName(wb.FullName), Path.GetFileName(targetSource), StringComparison.OrdinalIgnoreCase))
                        {
                            targetWb = wb;
                            break;
                        }
                    }
                }

                if (targetWb != null)
                {
                    foreach (Microsoft.Office.Interop.Excel.Worksheet ws in targetWb.Worksheets)
                    {
                        var item = new SheetSelectorItem
                        {
                            SheetName = ws.Name,
                            IsIncluded = !(IgnoreCoverAndHistory && DesignPageCounterService.IsCoverOrHistorySheet(ws.Name))
                        };
                        item.PropertyChanged += (s, e) => OnPropertyChanged(nameof(SheetSelectionHeader));
                        AvailableSheets.Add(item);
                    }
                }
                else if (File.Exists(targetSource))
                {
                    bool openedHere = false;
                    var tempWb = DesignPageCounterService.FindOrOpenWorkbook(_excelApp, targetSource, out openedHere);
                    if (tempWb != null)
                    {
                        try
                        {
                            foreach (Microsoft.Office.Interop.Excel.Worksheet ws in tempWb.Worksheets)
                            {
                                var item = new SheetSelectorItem
                                {
                                    SheetName = ws.Name,
                                    IsIncluded = !(IgnoreCoverAndHistory && DesignPageCounterService.IsCoverOrHistorySheet(ws.Name))
                                };
                                item.PropertyChanged += (s, e) => OnPropertyChanged(nameof(SheetSelectionHeader));
                                AvailableSheets.Add(item);
                            }
                        }
                        finally
                        {
                            if (openedHere)
                            {
                                try { tempWb.Close(false); Marshal.ReleaseComObject(tempWb); } catch { }
                            }
                        }
                    }
                }
            }
            catch { }

            OnPropertyChanged(nameof(SheetSelectionHeader));
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

            var excludedSet = new HashSet<string>(
                AvailableSheets.Where(s => !s.IsIncluded).Select(s => s.SheetName),
                StringComparer.OrdinalIgnoreCase);

            var options = new PageCounterOptions
            {
                Mode = Mode,
                CharactersPerPage = CharactersPerPage > 0 ? CharactersPerPage : 600,
                HighlightChangedCells = HighlightChangedCells,
                HighlightColorHex = SelectedHighlightColor?.Hex ?? "#FEF08A",
                IgnoreCoverAndHistory = IgnoreCoverAndHistory,
                IgnoreBlankPages = IgnoreBlankPages,
                CountShapesAndPictures = CountShapesAndPictures,
                MinChangedCellsThreshold = MinChangedCellsThreshold,
                ExcludedSheetNames = excludedSet
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

        private void OpenEvidence()
        {
            if (Result?.HighlightedClonedWorkbookPath == null || !File.Exists(Result.HighlightedClonedWorkbookPath))
            {
                WpfMessageBox.Show("Không tìm thấy file bản sao đã tô màu.", "Thông báo", WpfMessageBoxButton.OK, WpfMessageBoxImage.Warning);
                return;
            }

            try
            {
                bool ok = DesignPageCounterService.OpenEvidenceWorkbook(_excelApp, Result.HighlightedClonedWorkbookPath);
                if (!ok)
                {
                    WpfMessageBox.Show("Không thể mở file bằng Excel. Bạn có thể mở trực tiếp từ đường dẫn:\n" + Result.HighlightedClonedWorkbookPath, "Lỗi mở file", WpfMessageBoxButton.OK, WpfMessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Lỗi mở file: {ex.Message}", "Lỗi", WpfMessageBoxButton.OK, WpfMessageBoxImage.Error);
            }
        }

        private void ExportReport()
        {
            if (Result == null) return;

            try
            {
                Microsoft.Office.Interop.Excel.Workbook? targetWb = null;
                string targetSource = !string.IsNullOrEmpty(CustomTargetFilePath) ? CustomTargetFilePath : SelectedTargetWorkbook;

                if (_excelApp.Workbooks != null)
                {
                    foreach (Microsoft.Office.Interop.Excel.Workbook wb in _excelApp.Workbooks)
                    {
                        if (string.Equals(wb.Name, targetSource, StringComparison.OrdinalIgnoreCase) ||
                            string.Equals(wb.FullName, targetSource, StringComparison.OrdinalIgnoreCase) ||
                            string.Equals(Path.GetFileName(wb.FullName), Path.GetFileName(targetSource), StringComparison.OrdinalIgnoreCase))
                        {
                            targetWb = wb;
                            break;
                        }
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
