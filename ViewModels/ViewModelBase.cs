using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using WpfClipboard = System.Windows.Clipboard;

namespace ExcelSupport.ViewModels
{
    public abstract class ViewModelBase : ExcelSupport.Models.ObservableModel
    {

        private bool _isDarkTheme;
        public bool IsDarkTheme
        {
            get => _isDarkTheme;
            set => SetProperty(ref _isDarkTheme, value);
        }

        internal static void CopyToClipboard(string text)
        {
            if (string.IsNullOrEmpty(text)) return;
            try
            {
                WpfClipboard.SetText(text);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Lỗi copy clipboard: {ex.Message}");
            }
        }
    }
}
