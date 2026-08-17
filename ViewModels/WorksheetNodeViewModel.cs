using System;
using System.Drawing;
using System.Windows.Forms;
using System.Windows.Input;

namespace ExcelSupport.ViewModels
{
    public class WorksheetNodeViewModel : ViewModelBase
    {
        private string _sheetName = string.Empty;
        private bool _isActive;
        private int _index;
        private string? _tabColorHex;
        private bool _hasTabColor;
        private bool _isHidden;
        private bool _isVeryHidden;
        private bool _isProtected;

        public string WorkbookName { get; set; } = string.Empty;

        public string SheetName
        {
            get => _sheetName;
            set => SetProperty(ref _sheetName, value);
        }

        public int Index
        {
            get => _index;
            set => SetProperty(ref _index, value);
        }

        public bool IsActive
        {
            get => _isActive;
            set => SetProperty(ref _isActive, value);
        }

        public bool IsHidden
        {
            get => _isHidden;
            set
            {
                if (SetProperty(ref _isHidden, value))
                {
                    OnPropertyChanged(nameof(VisibilityStatusText));
                    OnPropertyChanged(nameof(IsVisibleSheet));
                }
            }
        }

        public bool IsVeryHidden
        {
            get => _isVeryHidden;
            set
            {
                if (SetProperty(ref _isVeryHidden, value))
                {
                    OnPropertyChanged(nameof(VisibilityStatusText));
                    OnPropertyChanged(nameof(IsVisibleSheet));
                }
            }
        }

        public bool IsProtected
        {
            get => _isProtected;
            set => SetProperty(ref _isProtected, value);
        }

        public bool IsVisibleSheet => !IsHidden && !IsVeryHidden;

        public string VisibilityStatusText
        {
            get
            {
                if (IsVeryHidden) return "Very Hidden";
                if (IsHidden) return "Hidden";
                return string.Empty;
            }
        }

        public string? TabColorHex
        {
            get => _tabColorHex;
            set
            {
                if (SetProperty(ref _tabColorHex, value))
                {
                    HasTabColor = !string.IsNullOrEmpty(value);
                    OnPropertyChanged(nameof(DisplayTabColor));
                }
            }
        }

        public string DisplayTabColor => !string.IsNullOrEmpty(_tabColorHex) ? _tabColorHex! : "#CBD5E1";

        public bool HasTabColor
        {
            get => _hasTabColor;
            private set => SetProperty(ref _hasTabColor, value);
        }

        public ICommand CopyNameCommand { get; }
        public ICommand RenameSheetCommand { get; }
        public ICommand SetColorCommand { get; }
        public ICommand PickCustomColorCommand { get; }
        public ICommand ClearColorCommand { get; }
        public ICommand UnhideSheetCommand { get; }
        public ICommand HideSheetCommand { get; }
        public ICommand VeryHideSheetCommand { get; }
        public ICommand CreateTocCommand { get; }
        public ICommand CheckVietnameseCommand { get; }
        public ICommand SortAscendingCommand { get; }
        public ICommand SortDescendingCommand { get; }

        public WorksheetNodeViewModel()
        {
            CopyNameCommand = new RelayCommand(_ => CopyToClipboard(SheetName));

            CheckVietnameseCommand = new RelayCommand(_ =>
            {
                Views.VietnameseCheckDialog.ShowWindow(IsDarkTheme);
            });

            RenameSheetCommand = new RelayCommand(_ =>
            {
                var dlg = new Views.RenameSheetDialog(SheetName, IsDarkTheme);
                if (dlg.ShowDialog() == true)
                {
                    string newName = dlg.NewSheetName;
                    string wbName = WorkbookName;
                    string oldName = SheetName;
                    ExcelDna.Integration.ExcelAsyncUtil.QueueAsMacro(() =>
                    {
                        AddInEvents.Instance?.RenameWorksheet(wbName, oldName, newName);
                    });
                }
            });

            SetColorCommand = new RelayCommand(param =>
            {
                if (param is string hex)
                {
                    TabColorHex = hex;
                    try
                    {
                        var color = ColorTranslator.FromHtml(hex);
                        AddInEvents.Instance?.ApplySheetColor(WorkbookName, SheetName, color);
                    }
                    catch { }
                }
            });

            PickCustomColorCommand = new RelayCommand(_ =>
            {
                using (var dlg = new ColorDialog())
                {
                    dlg.FullOpen = true;
                    if (!string.IsNullOrEmpty(TabColorHex))
                    {
                        try { dlg.Color = ColorTranslator.FromHtml(TabColorHex); } catch { }
                    }

                    if (dlg.ShowDialog() == DialogResult.OK)
                    {
                        var hex = $"#{dlg.Color.R:X2}{dlg.Color.G:X2}{dlg.Color.B:X2}";
                        TabColorHex = hex;
                        AddInEvents.Instance?.ApplySheetColor(WorkbookName, SheetName, dlg.Color);
                    }
                }
            });

            ClearColorCommand = new RelayCommand(_ =>
            {
                TabColorHex = null;
                AddInEvents.Instance?.ApplySheetColor(WorkbookName, SheetName, null);
            });

            UnhideSheetCommand = new RelayCommand(_ =>
            {
                AddInEvents.MainViewModel?.NotifySetSheetVisibility(WorkbookName, SheetName, -1); // xlSheetVisible = -1
            });

            HideSheetCommand = new RelayCommand(_ =>
            {
                AddInEvents.MainViewModel?.NotifySetSheetVisibility(WorkbookName, SheetName, 0); // xlSheetHidden = 0
            });

            VeryHideSheetCommand = new RelayCommand(_ =>
            {
                AddInEvents.MainViewModel?.NotifySetSheetVisibility(WorkbookName, SheetName, 2); // xlSheetVeryHidden = 2
            });

            CreateTocCommand = new RelayCommand(_ =>
            {
                string wbName = WorkbookName;
                ExcelDna.Integration.ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    AddInEvents.Instance?.CreateTableOfContents(wbName);
                });
            });

            SortAscendingCommand = new RelayCommand(_ =>
            {
                if (AddInEvents.MainViewModel != null)
                {
                    AddInEvents.MainViewModel.SheetSortOrder = SortOrder.Ascending;
                }
            });

            SortDescendingCommand = new RelayCommand(_ =>
            {
                if (AddInEvents.MainViewModel != null)
                {
                    AddInEvents.MainViewModel.SheetSortOrder = SortOrder.Descending;
                }
            });
        }
    }
}
