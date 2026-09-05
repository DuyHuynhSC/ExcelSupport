using System;
using System.Windows;
using System.Windows.Interop;
using ExcelSupport.ViewModels;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;

namespace ExcelSupport.Views
{
    public partial class TableExportDialog : Window
    {
        public static bool IsDarkTheme => AddInEvents.MainViewModel?.IsDarkTheme ?? false;

        public TableExportDialog(ExcelApp excelApp)
        {
            InitializeComponent();
            DataContext = new TableExportViewModel(excelApp);
        }

        public static void ShowWindow(ExcelApp excelApp)
        {
            var dlg = new TableExportDialog(excelApp);
            try
            {
                if (excelApp != null)
                {
                    var helper = new WindowInteropHelper(dlg)
                    {
                        Owner = new IntPtr(excelApp.Hwnd)
                    };
                }
            }
            catch { }

            dlg.ShowDialog();
        }
    }
}
