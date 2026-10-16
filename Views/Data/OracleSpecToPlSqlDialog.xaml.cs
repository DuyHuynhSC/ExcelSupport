using System;
using System.Windows;
using System.Windows.Interop;
using ExcelSupport.ViewModels;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;

namespace ExcelSupport.Views
{
    public partial class OracleSpecToPlSqlDialog : Window
    {
        public static readonly DependencyProperty IsDarkThemeProperty =
            DependencyProperty.Register(nameof(IsDarkTheme), typeof(bool), typeof(OracleSpecToPlSqlDialog),
                new PropertyMetadata(false));

        public bool IsDarkTheme
        {
            get => (bool)GetValue(IsDarkThemeProperty);
            set => SetValue(IsDarkThemeProperty, value);
        }

        private readonly OracleSpecToPlSqlViewModel _viewModel;

        public OracleSpecToPlSqlDialog(ExcelApp? excelApp = null, bool isDarkTheme = false)
        {
            InitializeComponent();
            IsDarkTheme = isDarkTheme;

            _viewModel = new OracleSpecToPlSqlViewModel(excelApp, isDarkTheme);
            _viewModel.OnSendToOracleQueryAction = (sql) =>
            {
                // Mở câu lệnh vừa sinh vào OracleQuickQueryDialog
                OracleQuickQueryDialog.ShowWindow(IsDarkTheme, sql);
            };

            DataContext = _viewModel;
        }

        public static void ShowWindow(ExcelApp? excelApp = null, bool? isDarkTheme = null)
        {
            try
            {
                var app = excelApp ?? AddInEvents.Instance?.ExcelAppInstance ?? (ExcelApp)ExcelDna.Integration.ExcelDnaUtil.Application;
                bool isDark = isDarkTheme ?? (AddInEvents.MainViewModel?.IsDarkTheme ?? false);

                var dlg = new OracleSpecToPlSqlDialog(app, isDark);

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
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(
                    $"Lỗi mở màn hình Spec to PL/SQL:\n{ex.Message}\n\nChi tiết:\n{ex.StackTrace}",
                    "ExcelSupport", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OnCloseClick(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
