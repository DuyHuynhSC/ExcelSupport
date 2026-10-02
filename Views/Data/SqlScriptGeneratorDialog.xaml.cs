using System;
using System.Windows;
using System.Windows.Interop;
using ExcelSupport.ViewModels;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;

namespace ExcelSupport.Views
{
    public partial class SqlScriptGeneratorDialog : Window
    {
        public static readonly DependencyProperty IsDarkThemeProperty =
            DependencyProperty.Register(nameof(IsDarkTheme), typeof(bool), typeof(SqlScriptGeneratorDialog),
                new PropertyMetadata(false));

        public bool IsDarkTheme
        {
            get => (bool)GetValue(IsDarkThemeProperty);
            set => SetValue(IsDarkThemeProperty, value);
        }

        private readonly SqlScriptGeneratorViewModel _viewModel;

        public SqlScriptGeneratorDialog(ExcelApp excelApp, bool isDarkTheme = false)
        {
            InitializeComponent();
            IsDarkTheme = isDarkTheme;

            _viewModel = new SqlScriptGeneratorViewModel(excelApp, isDarkTheme);
            _viewModel.OnSendToOracleQueryAction = (sql) =>
            {
                // Mở câu lệnh vừa sinh vào OracleQuickQueryDialog
                OracleQuickQueryDialog.ShowWindow(IsDarkTheme, sql);
            };

            DataContext = _viewModel;
        }

        public static void ShowWindow(ExcelApp? excelApp = null, bool? isDarkTheme = null)
        {
            var app = excelApp ?? AddInEvents.Instance?.ExcelAppInstance ?? (ExcelApp)ExcelDna.Integration.ExcelDnaUtil.Application;
            bool isDark = isDarkTheme ?? (AddInEvents.MainViewModel?.IsDarkTheme ?? false);

            var dlg = new SqlScriptGeneratorDialog(app, isDark);

            try
            {
                if (app != null)
                {
                    var helper = new WindowInteropHelper(dlg)
                    {
                        Owner = new IntPtr(app.Hwnd)
                    };
                }
            }
            catch { }

            dlg.ShowDialog();
        }

        private void OnCloseClick(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
