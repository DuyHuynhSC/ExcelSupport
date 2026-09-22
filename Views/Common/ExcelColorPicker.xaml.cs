using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using WpfUserControl = System.Windows.Controls.UserControl;
using WpfKeyEventArgs = System.Windows.Input.KeyEventArgs;
using WpfColor = System.Windows.Media.Color;
using WpfColorConverter = System.Windows.Media.ColorConverter;
using WpfBrush = System.Windows.Media.Brush;
using WpfBrushes = System.Windows.Media.Brushes;
using WpfButton = System.Windows.Controls.Button;

namespace ExcelSupport.Views.Common
{
    public partial class ExcelColorPicker : WpfUserControl, INotifyPropertyChanged
    {
        public static readonly DependencyProperty ColorHexProperty =
            DependencyProperty.Register(
                nameof(ColorHex),
                typeof(string),
                typeof(ExcelColorPicker),
                new FrameworkPropertyMetadata(string.Empty, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, OnColorHexChangedStatic));

        public static readonly DependencyProperty AllowTransparentProperty =
            DependencyProperty.Register(
                nameof(AllowTransparent),
                typeof(bool),
                typeof(ExcelColorPicker),
                new PropertyMetadata(false, OnAllowTransparentChangedStatic));

        public static readonly DependencyProperty IsDarkThemeProperty =
            DependencyProperty.Register(
                nameof(IsDarkTheme),
                typeof(bool),
                typeof(ExcelColorPicker),
                new PropertyMetadata(false));

        public event EventHandler<string>? ColorChanged;
        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged(string propertyName) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

        public string ColorHex
        {
            get => (string)GetValue(ColorHexProperty);
            set => SetValue(ColorHexProperty, value);
        }

        public bool AllowTransparent
        {
            get => (bool)GetValue(AllowTransparentProperty);
            set => SetValue(AllowTransparentProperty, value);
        }

        public bool IsDarkTheme
        {
            get => (bool)GetValue(IsDarkThemeProperty);
            set => SetValue(IsDarkThemeProperty, value);
        }

        public WpfBrush SwatchBrush
        {
            get
            {
                if (string.IsNullOrWhiteSpace(ColorHex) || string.Equals(ColorHex, "Transparent", StringComparison.OrdinalIgnoreCase))
                {
                    return WpfBrushes.Transparent;
                }
                return ParseBrush(ColorHex, Colors.Transparent);
            }
        }

        public string DisplayHex
        {
            get
            {
                if (string.IsNullOrWhiteSpace(ColorHex) || string.Equals(ColorHex, "Transparent", StringComparison.OrdinalIgnoreCase))
                {
                    return "None";
                }
                return ColorHex.ToUpperInvariant();
            }
        }

        public Visibility NoColorVisibility =>
            (string.IsNullOrWhiteSpace(ColorHex) || string.Equals(ColorHex, "Transparent", StringComparison.OrdinalIgnoreCase))
                ? Visibility.Visible : Visibility.Collapsed;

        public Visibility AllowTransparentVisibility =>
            AllowTransparent ? Visibility.Visible : Visibility.Collapsed;

        // Bảng màu Theme chuẩn Microsoft Office / Excel (10 cột x 6 hàng)
        public static readonly List<string> ThemeColors = new List<string>
        {
            // Row 1: Base colors
            "#FFFFFF", "#000000", "#E7E6E6", "#44546A", "#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47",
            // Row 2: Tints
            "#F2F2F2", "#7F7F7F", "#D9D9D9", "#D6DCE4", "#D9E1F2", "#FCE4D6", "#EDEDED", "#FFF2CC", "#DDEBF7", "#E2EFDA",
            // Row 3
            "#D9D9D9", "#595959", "#BFBFBF", "#ACB9CA", "#B4C6E7", "#F8CBAD", "#DBDBDB", "#FFE699", "#BDD7EE", "#C6E0B4",
            // Row 4
            "#BFBFBF", "#3F3F3F", "#A6A6A6", "#8497B0", "#8EA9DB", "#F4B084", "#C9C9C9", "#FFD966", "#9BC2E6", "#A9D08E",
            // Row 5: Shades
            "#A6A6A6", "#262626", "#808080", "#333F48", "#305496", "#C65911", "#7B7B7B", "#BF8F00", "#2F5597", "#548235",
            // Row 6: Deep shades
            "#7F7F7F", "#0C0C0C", "#595959", "#222A35", "#1F3864", "#833C0C", "#525252", "#7F6000", "#1E4E79", "#375623"
        };

        // Bảng 10 màu tiêu chuẩn Excel (Standard Colors)
        public static readonly List<string> StandardColors = new List<string>
        {
            "#C00000", "#FF0000", "#FFC000", "#FFFF00", "#92D050", "#00B050", "#00B0F0", "#0070C0", "#002060", "#7030A0"
        };

        public ExcelColorPicker()
        {
            InitializeComponent();
            icThemeColors.ItemsSource = ThemeColors;
            icStandardColors.ItemsSource = StandardColors;

            popupPicker.Opened += (s, e) =>
            {
                txtCustomHex.Text = ColorHex;
                txtCustomHex.Focus();
                txtCustomHex.SelectAll();
            };
        }

        private static void OnColorHexChangedStatic(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is ExcelColorPicker picker)
            {
                picker.OnPropertyChanged(nameof(SwatchBrush));
                picker.OnPropertyChanged(nameof(DisplayHex));
                picker.OnPropertyChanged(nameof(NoColorVisibility));
                picker.ColorChanged?.Invoke(picker, picker.ColorHex);
            }
        }

        private static void OnAllowTransparentChangedStatic(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is ExcelColorPicker picker)
            {
                picker.OnPropertyChanged(nameof(AllowTransparentVisibility));
            }
        }

        private void OnColorSwatchClick(object sender, RoutedEventArgs e)
        {
            if (sender is WpfButton btn && btn.Tag is string hex)
            {
                SetColorAndClose(hex);
            }
        }

        private void OnNoColorClick(object sender, RoutedEventArgs e)
        {
            SetColorAndClose(string.Empty);
        }

        private void OnApplyCustomHexClick(object sender, RoutedEventArgs e)
        {
            ApplyCustomHexInput();
        }

        private void OnCustomHexKeyDown(object sender, WpfKeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                ApplyCustomHexInput();
            }
        }

        private void ApplyCustomHexInput()
        {
            string input = txtCustomHex.Text?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(input))
            {
                if (AllowTransparent)
                {
                    SetColorAndClose(string.Empty);
                }
                return;
            }

            if (!input.StartsWith("#"))
            {
                input = "#" + input;
            }

            try
            {
                WpfColorConverter.ConvertFromString(input);
                SetColorAndClose(input);
            }
            catch
            {
                // Giữ nguyên nếu format không hợp lệ
            }
        }

        private void SetColorAndClose(string hex)
        {
            ColorHex = hex;
            popupPicker.IsOpen = false;
        }

        private static WpfBrush ParseBrush(string hex, WpfColor fallback)
        {
            try
            {
                string clean = hex.Trim();
                if (!clean.StartsWith("#")) clean = "#" + clean;
                var color = (WpfColor)WpfColorConverter.ConvertFromString(clean);
                return new SolidColorBrush(color);
            }
            catch
            {
                return new SolidColorBrush(fallback);
            }
        }
    }
}
