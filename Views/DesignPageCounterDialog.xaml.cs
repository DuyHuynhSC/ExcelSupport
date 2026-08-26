using System;
using System.Windows;
using System.Windows.Input;
using ExcelSupport.Services;
using ExcelSupport.ViewModels;
using WpfMessageBox = System.Windows.MessageBox;
using WpfMessageBoxButton = System.Windows.MessageBoxButton;
using WpfMessageBoxImage = System.Windows.MessageBoxImage;
using WpfKeyEventArgs = System.Windows.Input.KeyEventArgs;

namespace ExcelSupport.Views
{
    public partial class DesignPageCounterDialog : Window
    {
        public static readonly DependencyProperty IsDarkThemeProperty =
            DependencyProperty.Register(
                nameof(IsDarkTheme),
                typeof(bool),
                typeof(DesignPageCounterDialog),
                new PropertyMetadata(false));

        public bool IsDarkTheme
        {
            get => (bool)GetValue(IsDarkThemeProperty);
            set => SetValue(IsDarkThemeProperty, value);
        }

        private static DesignPageCounterDialog? _currentInstance;

        public DesignPageCounterDialog()
        {
            InitializeComponent();
            var app = AddInEvents.Instance?.ExcelAppInstance;
            if (app != null)
            {
                DataContext = new DesignPageCounterViewModel(app);
            }
        }

        public static void ShowWindow(bool? isDarkTheme = null)
        {
            try
            {
                if (_currentInstance != null && _currentInstance.IsLoaded)
                {
                    _currentInstance.Activate();
                    return;
                }

                _currentInstance = new DesignPageCounterDialog();

                bool dark = isDarkTheme ?? (AddInEvents.MainViewModel?.IsDarkTheme ?? AiConfigManager.Current.IsDarkTheme);
                _currentInstance.IsDarkTheme = dark;

                try
                {
                    var addIn = AddInEvents.Instance;
                    if (addIn?.ExcelAppInstance != null)
                    {
                        var helper = new System.Windows.Interop.WindowInteropHelper(_currentInstance);
                        helper.Owner = new IntPtr(addIn.ExcelAppInstance.Hwnd);
                    }
                }
                catch { }

                _currentInstance.ShowDialog();
                _currentInstance = null;
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Không thể mở hộp thoại Đếm số trang thiết kế:\n{ex.Message}",
                                   "Lỗi giao diện", WpfMessageBoxButton.OK, WpfMessageBoxImage.Error);
            }
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void Window_KeyDown(object sender, WpfKeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                Close();
            }
        }
    }
}

