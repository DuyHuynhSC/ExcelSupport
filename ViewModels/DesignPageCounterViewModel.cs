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
        public bool IsAnyColor => Hex.Equals("ANY", StringComparison.OrdinalIgnoreCase);

        public System.Windows.Media.Brush SwatchBrush
        {
            get
            {
                if (IsAnyColor)
                {
                    return new System.Windows.Media.LinearGradientBrush(
                        new System.Windows.Media.GradientStopCollection
                        {
                            new(System.Windows.Media.Color.FromRgb(239, 68, 68), 0.0),
                            new(System.Windows.Media.Color.FromRgb(245, 158, 11), 0.33),
                            new(System.Windows.Media.Color.FromRgb(16, 185, 129), 0.66),
                            new(System.Windows.Media.Color.FromRgb(59, 130, 246), 1.0)
                        },
                        new System.Windows.Point(0, 0),
                        new System.Windows.Point(1, 1));
                }

                try
                {
                    var converter = new System.Windows.Media.BrushConverter();
                    return (System.Windows.Media.Brush)(converter.ConvertFromString(Hex) ?? System.Windows.Media.Brushes.Yellow);
                }
                catch
                {
                    return System.Windows.Media.Brushes.Yellow;
                }
            }
        }
    }

    public class DensityPresetItem
    {
        public string Name { get; set; } = string.Empty;
        public int Value { get; set; }
    }

    public class ShapeFactorPresetItem
    {
        public string Name { get; set; } = string.Empty;
        public double Value { get; set; }
    }

    public class ProjectProfilePresetItem
    {
        public string Name { get; set; } = string.Empty;
        public int CharsPerPage { get; set; }
        public double ShapeFactor { get; set; }
        public string Description { get; set; } = string.Empty;
    }

    public class DesignPageCounterViewModel : ViewModelBase
    {
        private readonly ExcelApp _excelApp;

        private string _selectedTargetWorkbook = string.Empty;
        private string _selectedTemplateWorkbook = string.Empty;
        private string _customTargetFilePath = string.Empty;
        private string _customTemplateFilePath = string.Empty;

        private PageCounterMode _mode = PageCounterMode.UserHighlightedColor;
        private int _charactersPerPage = 600;
        private double _shapePageFactor = 0.5;
        private bool _highlightChangedCells = true;
        private ColorSwatchItem _selectedHighlightColor = null!;
        private DensityPresetItem _selectedDensityPreset = null!;
        private ShapeFactorPresetItem _selectedShapeFactorPreset = null!;
        private ProjectProfilePresetItem? _selectedProfile;

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
        public ObservableCollection<ShapeFactorPresetItem> AvailableShapeFactors { get; } = new();
        public ObservableCollection<ProjectProfilePresetItem> AvailableProfiles { get; } = new();

        public PageCounterMode Mode
        {
            get => _mode;
            set
            {
                if (SetProperty(ref _mode, value))
                {
                    OnPropertyChanged(nameof(IsUserHighlightMode));
                    OnPropertyChanged(nameof(IsAutoDiffMode));
                    OnPropertyChanged(nameof(IsCharHighlightMode));
                    OnPropertyChanged(nameof(IsPrintBreakMode));
                }
            }
        }

        public bool IsUserHighlightMode
        {
            get => Mode == PageCounterMode.UserHighlightedColor;
            set
            {
                if (value) Mode = PageCounterMode.UserHighlightedColor;
            }
        }

        public bool IsAutoDiffMode
        {
            get => Mode == PageCounterMode.AutoDiffTemplate;
            set
            {
                if (value) Mode = PageCounterMode.AutoDiffTemplate;
            }
        }

        public bool IsCharHighlightMode
        {
            get => Mode == PageCounterMode.UserHighlightedColor || Mode == PageCounterMode.AutoDiffTemplate;
            set
            {
                if (value && Mode == PageCounterMode.PrintBreakGrid) Mode = PageCounterMode.UserHighlightedColor;
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

        public double ShapePageFactor
        {
            get => _shapePageFactor;
            set => SetProperty(ref _shapePageFactor, value);
        }

        public bool HighlightChangedCells
        {
            get => _highlightChangedCells;
            set => SetProperty(ref _highlightChangedCells, value);
        }

        public ColorSwatchItem SelectedHighlightColor
        {
            get => _selectedHighlightColor;
            set
            {
                if (SetProperty(ref _selectedHighlightColor, value))
                {
                    DesignPageCounterService.CurrentHighlightColorHex = value?.Hex ?? "#FEF08A";
                }
            }
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

        public ShapeFactorPresetItem SelectedShapeFactorPreset
        {
            get => _selectedShapeFactorPreset;
            set
            {
                if (SetProperty(ref _selectedShapeFactorPreset, value) && value != null)
                {
                    ShapePageFactor = value.Value;
                }
            }
        }

        public ProjectProfilePresetItem? SelectedProfile
        {
            get => _selectedProfile;
            set
            {
                if (SetProperty(ref _selectedProfile, value) && value != null)
                {
                    CharactersPerPage = value.CharsPerPage;
                    ShapePageFactor = value.ShapeFactor;

                    var density = AvailableDensityPresets.FirstOrDefault(d => d.Value == value.CharsPerPage);
                    if (density != null) _selectedDensityPreset = density;
                    OnPropertyChanged(nameof(SelectedDensityPreset));

                    var shapeFactor = AvailableShapeFactors.FirstOrDefault(s => Math.Abs(s.Value - value.ShapeFactor) < 0.01);
                    if (shapeFactor != null) _selectedShapeFactorPreset = shapeFactor;
                    OnPropertyChanged(nameof(SelectedShapeFactorPreset));
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
                if (total == 0) return LocalizationService.Get("PageCounter_SheetSelectHeaderLoading");
                return LocalizationService.Get("PageCounter_SheetSelectHeaderCount", included, total);
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
        public ICommand CreateNewCopyForHighlightCommand { get; }
        public ICommand HighlightSelectionCommand { get; }
        public ICommand ClearHighlightSelectionCommand { get; }
        public ICommand SelectAllSheetsCommand { get; }
        public ICommand DeselectAllSheetsCommand { get; }
        public ICommand InvertSheetsSelectionCommand { get; }
        public ICommand AnalyzeCommand { get; }
        public ICommand ExportReportCommand { get; }
        public ICommand OpenEvidenceWorkbookCommand { get; }

        public DesignPageCounterViewModel(ExcelApp excelApp)
        {
            _excelApp = excelApp ?? throw new ArgumentNullException(nameof(excelApp));

            LoadPresets();
            LocalizationService.LanguageChanged += _ =>
            {
                LoadPresets();
                OnPropertyChanged(nameof(SheetSelectionHeader));
            };

            RefreshWorkbooksCommand = new RelayCommand(_ => RefreshWorkbooks());
            BrowseTargetFileCommand = new RelayCommand(_ => BrowseTargetFile());
            BrowseTemplateFileCommand = new RelayCommand(_ => BrowseTemplateFile());
            CreateNewCopyForHighlightCommand = new RelayCommand(_ => CreateNewCopyForHighlight(), _ => CanAnalyze);
            HighlightSelectionCommand = new RelayCommand(_ => HighlightCurrentSelection());
            ClearHighlightSelectionCommand = new RelayCommand(_ => ClearCurrentSelection());
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

        public void LoadPresets()
        {
            int prevDensityVal = SelectedDensityPreset?.Value ?? _charactersPerPage;
            string prevColorHex = SelectedHighlightColor?.Hex ?? "ANY";
            double prevShapeFactorVal = SelectedShapeFactorPreset?.Value ?? _shapePageFactor;

            AvailableProfiles.Clear();
            AvailableProfiles.Add(new ProjectProfilePresetItem
            {
                Name = LocalizationService.Get("PageCounter_ProfileJIS"),
                CharsPerPage = 600,
                ShapeFactor = 0.5,
                Description = LocalizationService.Get("PageCounter_ProfileJIS_Desc")
            });
            AvailableProfiles.Add(new ProjectProfilePresetItem
            {
                Name = LocalizationService.Get("PageCounter_ProfileVN"),
                CharsPerPage = 1200,
                ShapeFactor = 0.5,
                Description = LocalizationService.Get("PageCounter_ProfileVN_Desc")
            });
            AvailableProfiles.Add(new ProjectProfilePresetItem
            {
                Name = LocalizationService.Get("PageCounter_ProfileBackend"),
                CharsPerPage = 800,
                ShapeFactor = 0.3,
                Description = LocalizationService.Get("PageCounter_ProfileBackend_Desc")
            });
            AvailableProfiles.Add(new ProjectProfilePresetItem
            {
                Name = LocalizationService.Get("PageCounter_ProfileWeb"),
                CharsPerPage = 500,
                ShapeFactor = 0.6,
                Description = LocalizationService.Get("PageCounter_ProfileWeb_Desc")
            });
            _selectedProfile = AvailableProfiles[0];
            OnPropertyChanged(nameof(SelectedProfile));

            AvailableDensityPresets.Clear();
            AvailableDensityPresets.Add(new DensityPresetItem { Name = LocalizationService.Get("PageCounter_DensityJapanese"), Value = 600 });
            AvailableDensityPresets.Add(new DensityPresetItem { Name = LocalizationService.Get("PageCounter_DensityVietnamese"), Value = 1200 });
            AvailableDensityPresets.Add(new DensityPresetItem { Name = LocalizationService.Get("PageCounter_Density800"), Value = 800 });
            AvailableDensityPresets.Add(new DensityPresetItem { Name = LocalizationService.Get("PageCounter_Density1500"), Value = 1500 });
            _selectedDensityPreset = AvailableDensityPresets.FirstOrDefault(p => p.Value == prevDensityVal) ?? AvailableDensityPresets[0];
            OnPropertyChanged(nameof(SelectedDensityPreset));

            AvailableHighlightColors.Clear();
            AvailableHighlightColors.Add(new ColorSwatchItem { Name = LocalizationService.Get("PageCounter_ColorAny"), Hex = "ANY" });
            AvailableHighlightColors.Add(new ColorSwatchItem { Name = LocalizationService.Get("PageCounter_ColorYellow"), Hex = "#FEF08A" });
            AvailableHighlightColors.Add(new ColorSwatchItem { Name = LocalizationService.Get("PageCounter_ColorCyan"), Hex = "#BAE6FD" });
            AvailableHighlightColors.Add(new ColorSwatchItem { Name = LocalizationService.Get("PageCounter_ColorGreen"), Hex = "#BBF7D0" });
            AvailableHighlightColors.Add(new ColorSwatchItem { Name = LocalizationService.Get("PageCounter_ColorOrange"), Hex = "#FED7AA" });
            AvailableHighlightColors.Add(new ColorSwatchItem { Name = LocalizationService.Get("PageCounter_ColorPink"), Hex = "#DDD6FE" });
            _selectedHighlightColor = AvailableHighlightColors.FirstOrDefault(c => c.Hex.Equals(prevColorHex, StringComparison.OrdinalIgnoreCase)) ?? AvailableHighlightColors[0];
            OnPropertyChanged(nameof(SelectedHighlightColor));

            AvailableShapeFactors.Clear();
            AvailableShapeFactors.Add(new ShapeFactorPresetItem { Name = LocalizationService.Get("PageCounter_ShapeFactor05"), Value = 0.5 });
            AvailableShapeFactors.Add(new ShapeFactorPresetItem { Name = LocalizationService.Get("PageCounter_ShapeFactor0"), Value = 0.0 });
            AvailableShapeFactors.Add(new ShapeFactorPresetItem { Name = LocalizationService.Get("PageCounter_ShapeFactor025"), Value = 0.25 });
            AvailableShapeFactors.Add(new ShapeFactorPresetItem { Name = LocalizationService.Get("PageCounter_ShapeFactor1"), Value = 1.0 });
            _selectedShapeFactorPreset = AvailableShapeFactors.FirstOrDefault(s => Math.Abs(s.Value - prevShapeFactorVal) < 0.01) ?? AvailableShapeFactors[0];
            OnPropertyChanged(nameof(SelectedShapeFactorPreset));
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

        private void CreateNewCopyForHighlight()
        {
            string targetSource = !string.IsNullOrEmpty(CustomTargetFilePath) ? CustomTargetFilePath : SelectedTargetWorkbook;
            if (string.IsNullOrWhiteSpace(targetSource))
            {
                WpfMessageBox.Show("Vui lòng chọn hoặc duyệt file thiết kế nguồn trước.", "Chưa chọn file", WpfMessageBoxButton.OK, WpfMessageBoxImage.Warning);
                return;
            }

            try
            {
                string? newPath = DesignPageCounterService.CreateAndOpenNewCopyForHighlighting(_excelApp, targetSource);
                if (!string.IsNullOrEmpty(newPath))
                {
                    string fileName = Path.GetFileName(newPath);
                    if (!OpenWorkbooks.Contains(newPath!))
                    {
                        OpenWorkbooks.Insert(0, newPath!);
                    }
                    CustomTargetFilePath = newPath ?? string.Empty;
                    _selectedTargetWorkbook = newPath ?? string.Empty;
                    OnPropertyChanged(nameof(SelectedTargetWorkbook));
                    LoadTargetSheets();

                    string msg = string.Format(
                        LocalizationService.Instance["PageCounter_MsgCopyCreated"] ??
                        "Đã tạo và mở file bản sao mới:\n{0}\n\n👉 Hướng dẫn thao tác nhanh:\n1. Bôi đen vùng ô bạn đã thiết kế và nhấn phím tắt [Ctrl + Shift + H] để tô màu tức thì (hoặc dùng Fill Color trên thanh Ribbon).\n2. Lưu file (Ctrl + S) nếu muốn.\n3. Quay lại đây và bấm nút 'Phân Tích & Đếm Trang' để hoàn tất!",
                        fileName);
                    string title = LocalizationService.Instance["PageCounter_MsgCopyCreatedTitle"] ?? "Đã mở file bản sao để tô màu";

                    WpfMessageBox.Show(msg, title, WpfMessageBoxButton.OK, WpfMessageBoxImage.Information);
                }
                else
                {
                    WpfMessageBox.Show("Không thể tạo bản sao file Excel. Vui lòng kiểm tra quyền truy cập file.", "Lỗi tạo bản sao", WpfMessageBoxButton.OK, WpfMessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Lỗi tạo bản sao: {ex.Message}", "Lỗi", WpfMessageBoxButton.OK, WpfMessageBoxImage.Error);
            }
        }

        public void HighlightCurrentSelection()
        {
            try
            {
                string hex = SelectedHighlightColor?.Hex ?? "ANY";
                bool ok = DesignPageCounterService.HighlightSelection(_excelApp, hex);
                if (!ok)
                {
                    WpfMessageBox.Show("Không thể tô màu vùng chọn. Vui lòng mở file Excel và chọn ít nhất một ô trên bảng tính.", "Chưa chọn ô", WpfMessageBoxButton.OK, WpfMessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Lỗi tô màu: {ex.Message}", "Lỗi", WpfMessageBoxButton.OK, WpfMessageBoxImage.Error);
            }
        }

        public void ClearCurrentSelection()
        {
            try
            {
                bool ok = DesignPageCounterService.ClearHighlightSelection(_excelApp);
                if (!ok)
                {
                    WpfMessageBox.Show("Không thể xóa màu vùng chọn. Vui lòng mở file Excel và chọn ít nhất một ô trên bảng tính.", "Chưa chọn ô", WpfMessageBoxButton.OK, WpfMessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Lỗi xóa màu: {ex.Message}", "Lỗi", WpfMessageBoxButton.OK, WpfMessageBoxImage.Error);
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

            var excludedSet = new HashSet<string>(
                AvailableSheets.Where(s => !s.IsIncluded).Select(s => s.SheetName),
                StringComparer.OrdinalIgnoreCase);

            string colorHex = SelectedHighlightColor?.Hex ?? "ANY";
            bool matchAnyColor = colorHex.Equals("ANY", StringComparison.OrdinalIgnoreCase);

            var options = new PageCounterOptions
            {
                Mode = Mode,
                CharactersPerPage = CharactersPerPage > 0 ? CharactersPerPage : 600,
                ShapePageFactor = ShapePageFactor,
                HighlightChangedCells = HighlightChangedCells,
                HighlightColorHex = colorHex,
                MatchAnyHighlightColor = matchAnyColor,
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
