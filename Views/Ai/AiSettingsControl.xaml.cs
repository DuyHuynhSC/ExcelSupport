using System.Windows;
using ExcelSupport.ViewModels;
using WpfUserControl = System.Windows.Controls.UserControl;

namespace ExcelSupport.Views
{
    public partial class AiSettingsControl : WpfUserControl
    {
        public static readonly DependencyProperty IsDarkThemeProperty =
            DependencyProperty.Register(
                nameof(IsDarkTheme),
                typeof(bool),
                typeof(AiSettingsControl),
                new PropertyMetadata(false, OnIsDarkThemeChanged));

        public bool IsDarkTheme
        {
            get => (bool)GetValue(IsDarkThemeProperty);
            set => SetValue(IsDarkThemeProperty, value);
        }

        private static void OnIsDarkThemeChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is AiSettingsControl ctrl && ctrl.DataContext is ViewModelBase vm && e.NewValue is bool isDark)
            {
                vm.IsDarkTheme = isDark;
            }
        }

        public AiSettingsControl()
        {
            InitializeComponent();
        }
    }
}
