using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Windows.Data;
using System.Windows.Forms;
using System.Windows.Input;
using System.Runtime.InteropServices;

namespace ExcelSupport.ViewModels
{
    public enum SortOrder
    {
        Original,
        Ascending,
        Descending
    }

    public class TaskPaneViewModel : ViewModelBase
    {
        private WorkbookNodeViewModel? _selectedWorkbook;
        private WorksheetNodeViewModel? _selectedWorksheet;
        private bool _suppressActivationEvent;

        private string _workbookSearchText = string.Empty;
        private string _sheetSearchText = string.Empty;
        private SortOrder _workbookSortOrder = SortOrder.Original;
        private SortOrder _sheetSortOrder = SortOrder.Original;
        private bool _showHiddenSheets = false;
        private int _selectedTabIndex = 0;

        public int SelectedTabIndex
        {
            get => _selectedTabIndex;
            set => SetProperty(ref _selectedTabIndex, value);
        }

        public AiSettingsViewModel AiSettings { get; } = new AiSettingsViewModel();
        public AiAssistantViewModel AiAssistant { get; } = new AiAssistantViewModel();

        public ObservableCollection<WorkbookNodeViewModel> Workbooks { get; } 
            = new ObservableCollection<WorkbookNodeViewModel>();

        public ICollectionView WorkbooksView { get; }
        private ICollectionView? _sheetsView;

        public ICollectionView? SheetsView
        {
            get => _sheetsView;
            private set => SetProperty(ref _sheetsView, value);
        }

        public string WorkbookSearchText
        {
            get => _workbookSearchText;
            set
            {
                if (SetProperty(ref _workbookSearchText, value))
                {
                    WorkbooksView.Refresh();
                }
            }
        }

        public string SheetSearchText
        {
            get => _sheetSearchText;
            set
            {
                if (SetProperty(ref _sheetSearchText, value))
                {
                    SheetsView?.Refresh();
                }
            }
        }

        public bool ShowHiddenSheets
        {
            get => _showHiddenSheets;
            set
            {
                if (SetProperty(ref _showHiddenSheets, value))
                {
                    SheetsView?.Refresh();
                    OnPropertyChanged(nameof(ShowHiddenSheetsTooltip));
                }
            }
        }

        public string ShowHiddenSheetsTooltip => ShowHiddenSheets 
            ? "Đang hiển thị cả Sheet ẩn (Bấm để chỉ xem Sheet hiện)" 
            : "Đang ẩn các Sheet bị ẩn (Bấm để hiển thị tất cả)";

        public SortOrder WorkbookSortOrder
        {
            get => _workbookSortOrder;
            set
            {
                if (SetProperty(ref _workbookSortOrder, value))
                {
                    ApplyWorkbookSorting();
                    OnPropertyChanged(nameof(WorkbookSortLabel));
                }
            }
        }

        public SortOrder SheetSortOrder
        {
            get => _sheetSortOrder;
            set
            {
                if (SetProperty(ref _sheetSortOrder, value))
                {
                    ApplySheetSorting();
                    OnPropertyChanged(nameof(SheetSortLabel));
                }
            }
        }

        public string WorkbookSortLabel => WorkbookSortOrder switch
        {
            SortOrder.Ascending => "A-Z ↑",
            SortOrder.Descending => "Z-A ↓",
            _ => "A-Z"
        };

        public string SheetSortLabel => SheetSortOrder switch
        {
            SortOrder.Ascending => "A-Z ↑",
            SortOrder.Descending => "Z-A ↓",
            _ => "A-Z"
        };

        public WorkbookNodeViewModel? SelectedWorkbook
        {
            get => _selectedWorkbook;
            set
            {
                if (SetProperty(ref _selectedWorkbook, value))
                {
                    OnPropertyChanged(nameof(HasSelectedWorkbook));
                    UpdateSheetsView();

                    if (!_suppressActivationEvent && value != null)
                    {
                        RequestActivateWorkbook?.Invoke(value.WorkbookName);
                    }
                }
            }
        }

        public bool HasSelectedWorkbook => SelectedWorkbook != null;

        public WorksheetNodeViewModel? SelectedWorksheet
        {
            get => _selectedWorksheet;
            set
            {
                if (SetProperty(ref _selectedWorksheet, value))
                {
                    if (!_suppressActivationEvent && value != null)
                    {
                        RequestActivateWorksheet?.Invoke(value.WorkbookName, value.SheetName);
                    }
                }
            }
        }

        public string ThemeToggleIcon => IsDarkTheme ? "☀️" : "🌙";
        public string ThemeToggleTooltip => IsDarkTheme ? "Chuyển sang Giao diện Sáng (Light Theme)" : "Chuyển sang Giao diện Tối (Dark Theme)";

        public ICommand ActivateWorkbookCommand { get; }
        public ICommand ActivateWorksheetCommand { get; }
        public ICommand CloseWorkbookCommand { get; }
        public ICommand ToggleWorkbookSortCommand { get; }
        public ICommand ToggleSheetSortCommand { get; }
        public ICommand ToggleShowHiddenSheetsCommand { get; }
        public ICommand ToggleThemeCommand { get; }
        public ICommand CreateTocCommand { get; }
        public ICommand OpenToolsDialogCommand { get; }
        public ICommand SetWorkbookSortCommand { get; }
        public ICommand SetSheetSortCommand { get; }
        public ICommand ClearWorkbookSearchCommand { get; }
        public ICommand ClearSheetSearchCommand { get; }
        public ICommand CopyTextCommand { get; }
        public ICommand UnhideAllSheetsCommand { get; }
        public ICommand OpenVietnameseCheckDialogCommand { get; }
        public ICommand OpenWorkbookCompareDialogCommand { get; }

        public event Action<string>? RequestActivateWorkbook;
        public event Action<string, string>? RequestActivateWorksheet;
        public event Action<string>? RequestCloseWorkbook;
        public event Action<string, string, Color?>? RequestSetSheetTabColor;
        public event Action<string, string, int>? RequestSetSheetVisibility;
        public event Action<string>? RequestUnhideAllSheets;

        public TaskPaneViewModel()
        {
            // Load saved theme preference
            IsDarkTheme = ExcelSupport.Services.AiConfigManager.Current.IsDarkTheme;
            AiAssistant.IsDarkTheme = IsDarkTheme;
            AiSettings.IsDarkTheme = IsDarkTheme;

            ToggleThemeCommand = new RelayCommand(_ =>
            {
                IsDarkTheme = !IsDarkTheme;
                AiAssistant.IsDarkTheme = IsDarkTheme;
                AiSettings.IsDarkTheme = IsDarkTheme;
                OnPropertyChanged(nameof(ThemeToggleIcon));
                OnPropertyChanged(nameof(ThemeToggleTooltip));

                var cfg = ExcelSupport.Services.AiConfigManager.Current;
                cfg.IsDarkTheme = IsDarkTheme;
                ExcelSupport.Services.AiConfigManager.Save(cfg);
            });
            WorkbooksView = CollectionViewSource.GetDefaultView(Workbooks);
            WorkbooksView.Filter = FilterWorkbook;

            ActivateWorkbookCommand = new RelayCommand(param =>
            {
                if (param is WorkbookNodeViewModel wb)
                {
                    SelectedWorkbook = wb;
                    RequestActivateWorkbook?.Invoke(wb.WorkbookName);
                }
            });

            ActivateWorksheetCommand = new RelayCommand(param =>
            {
                if (param is WorksheetNodeViewModel ws)
                {
                    SelectedWorksheet = ws;
                    RequestActivateWorksheet?.Invoke(ws.WorkbookName, ws.SheetName);
                }
            });

            CloseWorkbookCommand = new RelayCommand(param =>
            {
                if (param is WorkbookNodeViewModel wb)
                {
                    RequestCloseWorkbook?.Invoke(wb.WorkbookName);
                }
                else if (param is string wbName)
                {
                    RequestCloseWorkbook?.Invoke(wbName);
                }
            });

            ToggleWorkbookSortCommand = new RelayCommand(_ =>
            {
                WorkbookSortOrder = (WorkbookSortOrder == SortOrder.Ascending)
                    ? SortOrder.Descending
                    : SortOrder.Ascending;
            });

            ToggleSheetSortCommand = new RelayCommand(_ =>
            {
                SheetSortOrder = (SheetSortOrder == SortOrder.Ascending)
                    ? SortOrder.Descending
                    : SortOrder.Ascending;
            });

            ToggleShowHiddenSheetsCommand = new RelayCommand(_ =>
            {
                ShowHiddenSheets = !ShowHiddenSheets;
            });

            CreateTocCommand = new RelayCommand(param =>
            {
                string? wbName = null;
                if (param is WorkbookNodeViewModel wb) wbName = wb.WorkbookName;
                else if (param is string name) wbName = name;
                else if (SelectedWorkbook != null) wbName = SelectedWorkbook.WorkbookName;

                ExcelDna.Integration.ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    AddInEvents.Instance?.CreateTableOfContents(wbName);
                });
            });

            OpenToolsDialogCommand = new RelayCommand(param =>
            {
                if (SelectedWorkbook != null)
                {
                    int tab = 0;
                    if (param is int t) tab = t;
                    else if (param is string s && int.TryParse(s, out int parsedTab)) tab = parsedTab;

                    var dlg = new Views.SheetToolsDialog(SelectedWorkbook, tab, IsDarkTheme);
                    dlg.ShowDialog();
                }
            });

            SetWorkbookSortCommand = new RelayCommand(param =>
            {
                if (param is string orderStr && Enum.TryParse<SortOrder>(orderStr, out var order))
                {
                    WorkbookSortOrder = order;
                }
            });

            SetSheetSortCommand = new RelayCommand(param =>
            {
                if (param is string orderStr && Enum.TryParse<SortOrder>(orderStr, out var order))
                {
                    SheetSortOrder = order;
                }
            });

            ClearWorkbookSearchCommand = new RelayCommand(_ => WorkbookSearchText = string.Empty);
            ClearSheetSearchCommand = new RelayCommand(_ => SheetSearchText = string.Empty);

            CopyTextCommand = new RelayCommand(param =>
            {
                if (param is string text)
                {
                    CopyToClipboard(text);
                }
            });

            UnhideAllSheetsCommand = new RelayCommand(param =>
            {
                if (param is WorkbookNodeViewModel wb)
                {
                    NotifyUnhideAllSheets(wb.WorkbookName);
                }
                else if (param is string wbName)
                {
                    NotifyUnhideAllSheets(wbName);
                }
                else if (SelectedWorkbook != null)
                {
                    NotifyUnhideAllSheets(SelectedWorkbook.WorkbookName);
                }
            });

            OpenVietnameseCheckDialogCommand = new RelayCommand(_ =>
            {
                Views.VietnameseCheckDialog.ShowWindow(IsDarkTheme);
            });

            OpenWorkbookCompareDialogCommand = new RelayCommand(param =>
            {
                string? defaultWb = null;
                if (param is WorkbookNodeViewModel wb) defaultWb = wb.WorkbookName;
                else if (param is string s) defaultWb = s;
                else if (SelectedWorkbook != null) defaultWb = SelectedWorkbook.WorkbookName;

                Views.WorkbookCompareDialog.ShowWindow(defaultWb, IsDarkTheme);
            });
        }

        private bool FilterWorkbook(object item)
        {
            if (string.IsNullOrWhiteSpace(WorkbookSearchText)) return true;
            if (item is WorkbookNodeViewModel wb)
            {
                return wb.WorkbookName.IndexOf(WorkbookSearchText, StringComparison.OrdinalIgnoreCase) >= 0;
            }
            return true;
        }

        private bool FilterSheet(object item)
        {
            if (item is WorksheetNodeViewModel ws)
            {
                if (!_showHiddenSheets && ws.IsHidden) return false;
                if (string.IsNullOrWhiteSpace(SheetSearchText)) return true;
                return ws.SheetName.IndexOf(SheetSearchText, StringComparison.OrdinalIgnoreCase) >= 0;
            }
            return true;
        }

        private void UpdateSheetsView()
        {
            if (SelectedWorkbook != null)
            {
                var view = CollectionViewSource.GetDefaultView(SelectedWorkbook.Worksheets);
                view.Filter = FilterSheet;
                SheetsView = view;
                ApplySheetSorting();
            }
            else
            {
                SheetsView = null;
            }
        }

        private void ApplyWorkbookSorting()
        {
            if (WorkbooksView is ListCollectionView lcv)
            {
                lcv.CustomSort = new WorkbookNodeComparer(WorkbookSortOrder);
                lcv.Refresh();
            }
            else if (WorkbooksView != null)
            {
                WorkbooksView.SortDescriptions.Clear();
                if (WorkbookSortOrder == SortOrder.Ascending)
                {
                    WorkbooksView.SortDescriptions.Add(new SortDescription(nameof(WorkbookNodeViewModel.WorkbookName), ListSortDirection.Ascending));
                }
                else if (WorkbookSortOrder == SortOrder.Descending)
                {
                    WorkbooksView.SortDescriptions.Add(new SortDescription(nameof(WorkbookNodeViewModel.WorkbookName), ListSortDirection.Descending));
                }
                WorkbooksView.Refresh();
            }
        }

        private void ApplySheetSorting()
        {
            if (SheetsView is ListCollectionView lcv)
            {
                lcv.CustomSort = new WorksheetNodeComparer(SheetSortOrder);
                lcv.Refresh();
            }
            else if (SheetsView != null)
            {
                SheetsView.SortDescriptions.Clear();
                if (SheetSortOrder == SortOrder.Ascending)
                {
                    SheetsView.SortDescriptions.Add(new SortDescription(nameof(WorksheetNodeViewModel.SheetName), ListSortDirection.Ascending));
                }
                else if (SheetSortOrder == SortOrder.Descending)
                {
                    SheetsView.SortDescriptions.Add(new SortDescription(nameof(WorksheetNodeViewModel.SheetName), ListSortDirection.Descending));
                }
                else
                {
                    SheetsView.SortDescriptions.Add(new SortDescription(nameof(WorksheetNodeViewModel.Index), ListSortDirection.Ascending));
                }
                SheetsView.Refresh();
            }
        }

        public void MergeWorkbooks(List<WorkbookNodeViewModel> incoming, string? activeWbName, string? activeWsName)
        {
            _suppressActivationEvent = true;
            try
            {
                // 1. Loại bỏ các Workbook đã đóng
                for (int i = Workbooks.Count - 1; i >= 0; i--)
                {
                    var existingWb = Workbooks[i];
                    if (!incoming.Any(w => string.Equals(w.WorkbookName, existingWb.WorkbookName, StringComparison.OrdinalIgnoreCase)))
                    {
                        Workbooks.RemoveAt(i);
                    }
                }

                // 2. Thêm mới hoặc cập nhật in-place các Workbook
                foreach (var inWb in incoming)
                {
                    var existingWb = Workbooks.FirstOrDefault(w => string.Equals(w.WorkbookName, inWb.WorkbookName, StringComparison.OrdinalIgnoreCase));
                    if (existingWb == null)
                    {
                        Workbooks.Add(inWb);
                    }
                    else
                    {
                        existingWb.FilePath = inWb.FilePath;
                        existingWb.IsActive = (existingWb.WorkbookName == activeWbName);

                        // Cập nhật danh sách Worksheets in-place
                        for (int j = existingWb.Worksheets.Count - 1; j >= 0; j--)
                        {
                            var existingWs = existingWb.Worksheets[j];
                            if (!inWb.Worksheets.Any(s => string.Equals(s.SheetName, existingWs.SheetName, StringComparison.OrdinalIgnoreCase)))
                            {
                                existingWb.Worksheets.RemoveAt(j);
                            }
                        }

                        foreach (var inWs in inWb.Worksheets)
                        {
                            var existingWs = existingWb.Worksheets.FirstOrDefault(s => string.Equals(s.SheetName, inWs.SheetName, StringComparison.OrdinalIgnoreCase));
                            if (existingWs == null)
                            {
                                existingWb.Worksheets.Add(inWs);
                            }
                            else
                            {
                                existingWs.Index = inWs.Index;
                                existingWs.IsActive = (existingWb.IsActive && existingWs.SheetName == activeWsName);
                                existingWs.IsHidden = inWs.IsHidden;
                                existingWs.IsVeryHidden = inWs.IsVeryHidden;
                                existingWs.IsProtected = inWs.IsProtected;
                                existingWs.TabColorHex = inWs.TabColorHex;
                                existingWs.WorkbookName = existingWb.WorkbookName;
                            }
                        }

                        existingWb.NotifyWorksheetsUpdated();
                    }
                }

                // 3. Đồng bộ lựa chọn hiển thị an toàn
                var targetWb = Workbooks.FirstOrDefault(w => string.Equals(w.WorkbookName, activeWbName, StringComparison.OrdinalIgnoreCase));
                if (SelectedWorkbook == null || !Workbooks.Contains(SelectedWorkbook))
                {
                    SelectedWorkbook = targetWb ?? Workbooks.FirstOrDefault();
                }
                else if (targetWb != null && SelectedWorkbook != targetWb)
                {
                    SelectedWorkbook = targetWb;
                }

                if (SelectedWorkbook != null)
                {
                    UpdateSheetsView();
                    var activeWs = SelectedWorkbook.Worksheets.FirstOrDefault(s => s.IsActive);
                    if (activeWs != null && SelectedWorksheet != activeWs)
                    {
                        SelectedWorksheet = activeWs;
                    }
                    else if (SelectedWorksheet == null || !SelectedWorkbook.Worksheets.Contains(SelectedWorksheet))
                    {
                        SelectedWorksheet = SelectedWorkbook.Worksheets.FirstOrDefault();
                    }
                }

                SheetsView?.Refresh();
                WorkbooksView.Refresh();
            }
            finally
            {
                _suppressActivationEvent = false;
            }
        }

        public void SetActiveSelection(string? activeWbName, string? activeWsName)
        {
            _suppressActivationEvent = true;
            try
            {
                foreach (var wb in Workbooks)
                {
                    bool isWbActive = string.Equals(wb.WorkbookName, activeWbName, StringComparison.OrdinalIgnoreCase);
                    wb.IsActive = isWbActive;

                    foreach (var ws in wb.Worksheets)
                    {
                        ws.IsActive = (isWbActive && string.Equals(ws.SheetName, activeWsName, StringComparison.OrdinalIgnoreCase));
                    }
                }

                var targetWb = Workbooks.FirstOrDefault(w => string.Equals(w.WorkbookName, activeWbName, StringComparison.OrdinalIgnoreCase));
                if (targetWb != null && SelectedWorkbook != targetWb)
                {
                    SelectedWorkbook = targetWb;
                }

                if (SelectedWorkbook != null)
                {
                    var activeWs = SelectedWorkbook.Worksheets.FirstOrDefault(s => s.IsActive);
                    if (activeWs != null && SelectedWorksheet != activeWs)
                    {
                        SelectedWorksheet = activeWs;
                    }
                }
            }
            finally
            {
                _suppressActivationEvent = false;
            }
        }

        public void NotifySetSheetColor(string wbName, string wsName, Color? color)
        {
            string? hex = color.HasValue ? $"#{color.Value.R:X2}{color.Value.G:X2}{color.Value.B:X2}" : null;
            foreach (var wb in Workbooks)
            {
                if (string.Equals(wb.WorkbookName, wbName, StringComparison.OrdinalIgnoreCase) ||
                    (string.IsNullOrEmpty(wbName) && wb.IsActive))
                {
                    foreach (var ws in wb.Worksheets)
                    {
                        if (string.Equals(ws.SheetName, wsName, StringComparison.OrdinalIgnoreCase))
                        {
                            ws.TabColorHex = hex;
                        }
                    }
                }
            }
            RequestSetSheetTabColor?.Invoke(wbName, wsName, color);
        }

        public void NotifySetSheetVisibility(string wbName, string wsName, int visibility)
        {
            RequestSetSheetVisibility?.Invoke(wbName, wsName, visibility);
        }

        public void NotifyUnhideAllSheets(string wbName)
        {
            RequestUnhideAllSheets?.Invoke(wbName);
        }
    }

    public class WorksheetNodeComparer : System.Collections.IComparer
    {
        private readonly SortOrder _order;

        [DllImport("shlwapi.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
        private static extern int StrCmpLogicalW(string psz1, string psz2);

        public WorksheetNodeComparer(SortOrder order)
        {
            _order = order;
        }

        public int Compare(object? x, object? y)
        {
            if (x is WorksheetNodeViewModel ws1 && y is WorksheetNodeViewModel ws2)
            {
                if (_order == SortOrder.Ascending)
                {
                    return StrCmpLogicalW(ws1.SheetName ?? string.Empty, ws2.SheetName ?? string.Empty);
                }
                if (_order == SortOrder.Descending)
                {
                    return StrCmpLogicalW(ws2.SheetName ?? string.Empty, ws1.SheetName ?? string.Empty);
                }
                // SortOrder.Original: Sắp xếp theo thứ tự Tab thực tế trong Excel
                return ws1.Index.CompareTo(ws2.Index);
            }
            return 0;
        }
    }

    public class WorkbookNodeComparer : System.Collections.IComparer
    {
        private readonly SortOrder _order;

        [DllImport("shlwapi.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
        private static extern int StrCmpLogicalW(string psz1, string psz2);

        public WorkbookNodeComparer(SortOrder order)
        {
            _order = order;
        }

        public int Compare(object? x, object? y)
        {
            if (x is WorkbookNodeViewModel wb1 && y is WorkbookNodeViewModel wb2)
            {
                if (_order == SortOrder.Ascending)
                {
                    return StrCmpLogicalW(wb1.WorkbookName ?? string.Empty, wb2.WorkbookName ?? string.Empty);
                }
                if (_order == SortOrder.Descending)
                {
                    return StrCmpLogicalW(wb2.WorkbookName ?? string.Empty, wb1.WorkbookName ?? string.Empty);
                }
            }
            return 0;
        }
    }
}
