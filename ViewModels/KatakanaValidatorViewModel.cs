using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Input;
using ExcelSupport.Services;
using Microsoft.Office.Interop.Excel;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using WpfMessageBox = System.Windows.MessageBox;
using WpfMessageBoxButton = System.Windows.MessageBoxButton;
using WpfMessageBoxImage = System.Windows.MessageBoxImage;

namespace ExcelSupport.ViewModels
{
    public class KatakanaValidatorViewModel : ViewModelBase
    {
        private readonly ExcelApp _excelApp;

        private ConversionScope _scope = ConversionScope.ActiveSheet;
        private bool _isScanning;
        private string _progressStatus = string.Empty;
        private int _progressPercent;
        private KatakanaInconsistencyGroup? _selectedGroup;
        private KatakanaCellLocation? _selectedLocation;
        private string _searchFilter = string.Empty;

        public ObservableCollection<KatakanaInconsistencyGroup> AllGroups { get; } = new();
        public ObservableCollection<KatakanaInconsistencyGroup> FilteredGroups { get; } = new();

        public ConversionScope Scope
        {
            get => _scope;
            set => SetProperty(ref _scope, value);
        }

        public bool IsScopeSelection
        {
            get => _scope == ConversionScope.Selection;
            set { if (value) Scope = ConversionScope.Selection; }
        }

        public bool IsScopeActiveSheet
        {
            get => _scope == ConversionScope.ActiveSheet;
            set { if (value) Scope = ConversionScope.ActiveSheet; }
        }

        public bool IsScopeActiveWorkbook
        {
            get => _scope == ConversionScope.ActiveWorkbook;
            set { if (value) Scope = ConversionScope.ActiveWorkbook; }
        }

        public bool IsScanning
        {
            get => _isScanning;
            set
            {
                if (SetProperty(ref _isScanning, value))
                {
                    OnPropertyChanged(nameof(CanScan));
                }
            }
        }

        public bool CanScan => !IsScanning;

        public string ProgressStatus
        {
            get => _progressStatus;
            set => SetProperty(ref _progressStatus, value);
        }

        public int ProgressPercent
        {
            get => _progressPercent;
            set => SetProperty(ref _progressPercent, value);
        }

        public KatakanaInconsistencyGroup? SelectedGroup
        {
            get => _selectedGroup;
            set
            {
                if (SetProperty(ref _selectedGroup, value))
                {
                    OnPropertyChanged(nameof(HasSelectedGroup));
                    SelectedLocation = value?.Variants.FirstOrDefault()?.Locations.FirstOrDefault();
                }
            }
        }

        public bool HasSelectedGroup => SelectedGroup != null;

        public KatakanaCellLocation? SelectedLocation
        {
            get => _selectedLocation;
            set => SetProperty(ref _selectedLocation, value);
        }

        public string SearchFilter
        {
            get => _searchFilter;
            set
            {
                if (SetProperty(ref _searchFilter, value))
                {
                    ApplyFilter();
                }
            }
        }

        public string TotalFoundSummary
        {
            get
            {
                if (AllGroups.Count == 0) return LocalizationService.Get("Katakana_NoInconsistenciesFound");
                int totalOccurrences = AllGroups.Sum(g => g.TotalOccurrences);
                return LocalizationService.Get("Katakana_SummaryFound", AllGroups.Count, totalOccurrences);
            }
        }

        public ICommand ScanCommand { get; }
        public ICommand StandardizeSelectedWithChouonCommand { get; }
        public ICommand StandardizeSelectedWithoutChouonCommand { get; }
        public ICommand StandardizeAllWithChouonCommand { get; }
        public ICommand NavigateToCellCommand { get; }

        public KatakanaValidatorViewModel(ExcelApp excelApp)
        {
            _excelApp = excelApp ?? throw new ArgumentNullException(nameof(excelApp));

            ScanCommand = new RelayCommand(async _ => await ScanAsync(), _ => CanScan);
            StandardizeSelectedWithChouonCommand = new RelayCommand(_ => StandardizeSelectedGroup(true), _ => HasSelectedGroup);
            StandardizeSelectedWithoutChouonCommand = new RelayCommand(_ => StandardizeSelectedGroup(false), _ => HasSelectedGroup);
            StandardizeAllWithChouonCommand = new RelayCommand(async _ => await StandardizeAllAsync(true), _ => FilteredGroups.Count > 0 && !IsScanning);
            NavigateToCellCommand = new RelayCommand(param => NavigateToLocation(param as KatakanaCellLocation));
        }

        public async Task ScanAsync()
        {
            IsScanning = true;
            ProgressStatus = "Đang quét các từ Katakana...";
            ProgressPercent = 0;
            AllGroups.Clear();
            FilteredGroups.Clear();
            SelectedGroup = null;

            try
            {
                var groups = await Task.Run(() =>
                {
                    return KatakanaValidatorService.ScanInconsistencies(
                        _excelApp,
                        Scope,
                        (msg, p) =>
                        {
                            System.Windows.Application.Current?.Dispatcher.Invoke(() =>
                            {
                                ProgressStatus = msg;
                                ProgressPercent = p;
                            });
                        });
                });

                foreach (var g in groups)
                {
                    AllGroups.Add(g);
                }

                ApplyFilter();
                OnPropertyChanged(nameof(TotalFoundSummary));

                if (FilteredGroups.Count > 0)
                {
                    SelectedGroup = FilteredGroups[0];
                }
                else
                {
                    WpfMessageBox.Show("Tuyệt vời! Không phát hiện bất kỳ sự bất đồng nhất Katakana nào trong phạm vi đã chọn.", "Không có lỗi", WpfMessageBoxButton.OK, WpfMessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Lỗi khi quét Katakana: {ex.Message}", "Lỗi", WpfMessageBoxButton.OK, WpfMessageBoxImage.Error);
            }
            finally
            {
                IsScanning = false;
                ProgressStatus = string.Empty;
            }
        }

        private void ApplyFilter()
        {
            FilteredGroups.Clear();
            string filter = SearchFilter.Trim().ToLowerInvariant();

            foreach (var g in AllGroups)
            {
                if (string.IsNullOrEmpty(filter) ||
                    g.BaseStem.ToLowerInvariant().Contains(filter) ||
                    g.Variants.Any(v => v.Word.ToLowerInvariant().Contains(filter)))
                {
                    FilteredGroups.Add(g);
                }
            }
        }

        public void StandardizeSelectedGroup(bool withChouon)
        {
            if (SelectedGroup == null) return;

            string targetWord = withChouon ? SelectedGroup.WithChouonWord : SelectedGroup.WithoutChouonWord;
            int totalReplaced = 0;

            foreach (var variant in SelectedGroup.Variants)
            {
                if (variant.Word != targetWord)
                {
                    totalReplaced += KatakanaValidatorService.ReplaceKatakanaWord(_excelApp, variant.Word, targetWord, variant.Locations);
                }
            }

            WpfMessageBox.Show($"Đã chuẩn hóa thành công cụm từ sang chuẩn [{targetWord}] ({totalReplaced} ô được cập nhật)!", "Chuẩn Hóa Thành Công", WpfMessageBoxButton.OK, WpfMessageBoxImage.Information);

            // Xóa nhóm khỏi danh sách sau khi đã chuẩn hóa
            AllGroups.Remove(SelectedGroup);
            FilteredGroups.Remove(SelectedGroup);
            SelectedGroup = FilteredGroups.FirstOrDefault();
            OnPropertyChanged(nameof(TotalFoundSummary));
        }

        public async Task StandardizeAllAsync(bool withChouon)
        {
            if (FilteredGroups.Count == 0) return;

            var res = WpfMessageBox.Show(
                $"Bạn có chắc chắn muốn chuẩn hóa toàn bộ {FilteredGroups.Count} nhóm từ Katakana sang chuẩn {(withChouon ? "CÓ TRƯỜNG ÂM (JIS Standard)" : "KHÔNG TRƯỜNG ÂM")}?",
                "Xác Nhận Chuẩn Hóa Hàng Loạt",
                WpfMessageBoxButton.YesNo,
                WpfMessageBoxImage.Question);

            if (res != System.Windows.MessageBoxResult.Yes) return;

            IsScanning = true;
            ProgressStatus = "Đang chuẩn hóa toàn bộ tài liệu...";
            int totalReplaced = 0;

            try
            {
                var groupsList = FilteredGroups.ToList();
                await Task.Run(() =>
                {
                    foreach (var g in groupsList)
                    {
                        string targetWord = withChouon ? g.WithChouonWord : g.WithoutChouonWord;
                        foreach (var variant in g.Variants)
                        {
                            if (variant.Word != targetWord)
                            {
                                totalReplaced += KatakanaValidatorService.ReplaceKatakanaWord(_excelApp, variant.Word, targetWord, variant.Locations);
                            }
                        }
                    }
                });

                WpfMessageBox.Show($"✅ Đã chuẩn hóa thành công toàn bộ {groupsList.Count} nhóm Katakana ({totalReplaced} ô đã cập nhật)!", "Hoàn Tất Chuẩn Hóa", WpfMessageBoxButton.OK, WpfMessageBoxImage.Information);

                AllGroups.Clear();
                FilteredGroups.Clear();
                SelectedGroup = null;
                OnPropertyChanged(nameof(TotalFoundSummary));
            }
            finally
            {
                IsScanning = false;
                ProgressStatus = string.Empty;
            }
        }

        public void NavigateToLocation(KatakanaCellLocation? loc)
        {
            if (loc == null || _excelApp == null) return;

            try
            {
                Workbook? targetWb = null;
                if (!string.IsNullOrEmpty(loc.WorkbookName))
                {
                    try { targetWb = _excelApp.Workbooks[loc.WorkbookName]; } catch { }
                }
                targetWb ??= _excelApp.ActiveWorkbook;
                if (targetWb == null) return;

                targetWb.Activate();
                Worksheet? ws = null;
                try { ws = targetWb.Worksheets[loc.SheetName]; } catch { }
                if (ws == null) return;

                ws.Activate();
                Range cell = ws.Range[loc.CellAddress];
                cell.Select();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error navigating to cell: {ex.Message}");
            }
        }
    }
}
