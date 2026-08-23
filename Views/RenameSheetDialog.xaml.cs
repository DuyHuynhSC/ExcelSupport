using System;
using System.Windows;
using MediaColor = System.Windows.Media.Color;
using MediaColorConverter = System.Windows.Media.ColorConverter;
using SolidColorBrush = System.Windows.Media.SolidColorBrush;
using WpfMessageBox = System.Windows.MessageBox;
using ExcelSupport.Services;

namespace ExcelSupport.Views
{
    public partial class RenameSheetDialog : Window
    {
        public string NewSheetName { get; private set; } = string.Empty;

        public RenameSheetDialog(string currentSheetName, bool isDarkTheme = false)
        {
            InitializeComponent();

            TxtPrompt.Text = string.Format(LocalizationService.Get("RenameSheet_PromptFormat"), currentSheetName);
            TxtNewSheetName.Text = currentSheetName;
            TxtNewSheetName.SelectAll();
            TxtNewSheetName.Focus();

            ApplyTheme(isDarkTheme);
        }

        private void ApplyTheme(bool isDarkTheme)
        {
            if (isDarkTheme)
            {
                Background = new SolidColorBrush((MediaColor)MediaColorConverter.ConvertFromString("#0F172A"));
                Resources["TextBrush"] = new SolidColorBrush((MediaColor)MediaColorConverter.ConvertFromString("#F8FAFC"));
                Resources["InputBgBrush"] = new SolidColorBrush((MediaColor)MediaColorConverter.ConvertFromString("#1E293B"));
                Resources["BorderBrush"] = new SolidColorBrush((MediaColor)MediaColorConverter.ConvertFromString("#475569"));
                BtnCancel.Background = new SolidColorBrush((MediaColor)MediaColorConverter.ConvertFromString("#334155"));
                BtnCancel.Foreground = new SolidColorBrush((MediaColor)MediaColorConverter.ConvertFromString("#F1F5F9"));
            }
            else
            {
                Background = new SolidColorBrush((MediaColor)MediaColorConverter.ConvertFromString("#FFFFFF"));
                Resources["TextBrush"] = new SolidColorBrush((MediaColor)MediaColorConverter.ConvertFromString("#1E293B"));
                Resources["InputBgBrush"] = new SolidColorBrush((MediaColor)MediaColorConverter.ConvertFromString("#FFFFFF"));
                Resources["BorderBrush"] = new SolidColorBrush((MediaColor)MediaColorConverter.ConvertFromString("#CBD5E1"));
                BtnCancel.Background = new SolidColorBrush((MediaColor)MediaColorConverter.ConvertFromString("#F1F5F9"));
                BtnCancel.Foreground = new SolidColorBrush((MediaColor)MediaColorConverter.ConvertFromString("#334155"));
            }
        }

        private void OnOkClick(object sender, RoutedEventArgs e)
        {
            string name = TxtNewSheetName.Text?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(name))
            {
                WpfMessageBox.Show(LocalizationService.Get("RenameSheet_MsgEmpty"), LocalizationService.Get("RenameSheet_WindowTitle"), MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            char[] invalidChars = { '\\', '/', '?', '*', '[', ']', ':' };
            if (name.IndexOfAny(invalidChars) >= 0)
            {
                WpfMessageBox.Show(LocalizationService.Get("RenameSheet_MsgInvalidChars"), LocalizationService.Get("RenameSheet_WindowTitle"), MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            NewSheetName = name;
            DialogResult = true;
            Close();
        }

        private void OnCancelClick(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
