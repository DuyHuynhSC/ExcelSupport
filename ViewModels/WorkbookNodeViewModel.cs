using System.Collections.ObjectModel;
using System.Windows.Input;

namespace ExcelSupport.ViewModels
{
    public class WorkbookNodeViewModel : ViewModelBase
    {
        private string _workbookName = string.Empty;
        private string _filePath = string.Empty;
        private bool _isActive;

        public string WorkbookName
        {
            get => _workbookName;
            set => SetProperty(ref _workbookName, value);
        }

        public string FilePath
        {
            get => _filePath;
            set => SetProperty(ref _filePath, value);
        }

        public bool IsActive
        {
            get => _isActive;
            set => SetProperty(ref _isActive, value);
        }

        public int SheetCount => Worksheets.Count;

        public ObservableCollection<WorksheetNodeViewModel> Worksheets { get; } 
            = new ObservableCollection<WorksheetNodeViewModel>();

        public ICommand CopyNameCommand { get; }
        public ICommand CopyPathCommand { get; }
        public ICommand CreateTocCommand { get; }
        public ICommand UnhideAllSheetsCommand { get; }
        public ICommand CloseWorkbookCommand { get; }
        public ICommand SortAscendingCommand { get; }
        public ICommand SortDescendingCommand { get; }
        public ICommand OpenToolsDialogCommand { get; }
        public ICommand SplitWorkbookCommand { get; }
        public ICommand MergeSheetsCommand { get; }
        public ICommand BatchRenameCommand { get; }

        public WorkbookNodeViewModel()
        {
            CopyNameCommand = new RelayCommand(_ => CopyToClipboard(WorkbookName));
            CopyPathCommand = new RelayCommand(_ => CopyToClipboard(string.IsNullOrEmpty(FilePath) ? WorkbookName : FilePath));

            CreateTocCommand = new RelayCommand(_ =>
            {
                string wbName = WorkbookName;
                ExcelDna.Integration.ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    AddInEvents.Instance?.CreateTableOfContents(wbName);
                });
            });

            UnhideAllSheetsCommand = new RelayCommand(_ =>
            {
                AddInEvents.MainViewModel?.NotifyUnhideAllSheets(WorkbookName);
            });

            CloseWorkbookCommand = new RelayCommand(_ =>
            {
                AddInEvents.MainViewModel?.CloseWorkbookCommand.Execute(WorkbookName);
            });

            SortAscendingCommand = new RelayCommand(_ =>
            {
                if (AddInEvents.MainViewModel != null)
                {
                    AddInEvents.MainViewModel.WorkbookSortOrder = SortOrder.Ascending;
                }
            });

            SortDescendingCommand = new RelayCommand(_ =>
            {
                if (AddInEvents.MainViewModel != null)
                {
                    AddInEvents.MainViewModel.WorkbookSortOrder = SortOrder.Descending;
                }
            });

            OpenToolsDialogCommand = new RelayCommand(param =>
            {
                int tab = 0;
                if (param is int t) tab = t;
                else if (param is string s && int.TryParse(s, out int parsedTab)) tab = parsedTab;

                var dlg = new Views.SheetToolsDialog(this, tab, IsDarkTheme);
                dlg.ShowDialog();
            });

            SplitWorkbookCommand = new RelayCommand(_ =>
            {
                var dlg = new Views.SheetToolsDialog(this, 0, IsDarkTheme);
                dlg.ShowDialog();
            });

            MergeSheetsCommand = new RelayCommand(_ =>
            {
                var dlg = new Views.SheetToolsDialog(this, 1, IsDarkTheme);
                dlg.ShowDialog();
            });

            BatchRenameCommand = new RelayCommand(_ =>
            {
                var dlg = new Views.SheetToolsDialog(this, 3, IsDarkTheme);
                dlg.ShowDialog();
            });
        }

        public void NotifyWorksheetsUpdated()
        {
            OnPropertyChanged(nameof(SheetCount));
        }
    }
}
