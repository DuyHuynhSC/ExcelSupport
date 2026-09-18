using System;
using System.Windows;
using System.Windows.Interop;
using ExcelSupport.ViewModels;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;

namespace ExcelSupport.Views
{
    public partial class KatakanaValidatorDialog : Window
    {
        public static bool IsDarkTheme => AddInEvents.MainViewModel?.IsDarkTheme ?? false;

        public KatakanaValidatorDialog(ExcelApp excelApp)
        {
            InitializeComponent();
            DataContext = new KatakanaValidatorViewModel(excelApp);
        }

        public static void ShowWindow(ExcelApp excelApp)
        {
            var dlg = new KatakanaValidatorDialog(excelApp);
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
