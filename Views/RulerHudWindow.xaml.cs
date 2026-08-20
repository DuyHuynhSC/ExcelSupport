using System;
using System.Globalization;
using System.Windows;
using System.Windows.Input;
using ExcelSupport.Host;
using ExcelSupport.Services;

namespace ExcelSupport.Views
{
    public partial class RulerHudWindow : Window
    {
        public static readonly DependencyProperty IsDarkThemeProperty =
            DependencyProperty.Register(nameof(IsDarkTheme), typeof(bool), typeof(RulerHudWindow),
                new PropertyMetadata(true));

        public bool IsDarkTheme
        {
            get => (bool)GetValue(IsDarkThemeProperty);
            set => SetValue(IsDarkThemeProperty, value);
        }

        public static readonly DependencyProperty DynamicFontSizeProperty =
            DependencyProperty.Register(nameof(DynamicFontSize), typeof(double), typeof(RulerHudWindow),
                new PropertyMetadata(13.5));

        public double DynamicFontSize
        {
            get => (double)GetValue(DynamicFontSizeProperty);
            set => SetValue(DynamicFontSizeProperty, value);
        }

        private static RulerHudWindow? _instance;
        private static double _savedFontSize = 13.5;
        private static double _savedLeft = -1;
        private static double _savedTop = -1;
        private static bool _userClosed = false;

        public static bool IsHudVisible => _instance != null && _instance.IsVisible;

        public static void ShowHud(bool isDarkTheme = false)
        {
            try
            {
                if (_userClosed) return;

                if (_instance == null)
                {
                    _instance = new RulerHudWindow
                    {
                        IsDarkTheme = isDarkTheme,
                        DynamicFontSize = _savedFontSize
                    };

                    if (_savedLeft >= 0 && _savedTop >= 0)
                    {
                        _instance.WindowStartupLocation = WindowStartupLocation.Manual;
                        _instance.Left = _savedLeft;
                        _instance.Top = _savedTop;
                    }
                    else
                    {
                        // Mặc định xuất hiện ở góc dưới bên phải màn hình làm việc
                        _instance.WindowStartupLocation = WindowStartupLocation.Manual;
                        var workArea = SystemParameters.WorkArea;
                        _instance.Left = workArea.Right - 540;
                        _instance.Top = workArea.Bottom - 140;
                    }

                    _instance.Show();
                }
                else
                {
                    _instance.IsDarkTheme = isDarkTheme;
                    if (!_instance.IsVisible) _instance.Show();
                }

                // Nạp ngay thông số thống kê gần nhất nếu có
                if (GridRulerService.LastQuickStats != null)
                {
                    _instance.UpdateStats(GridRulerService.LastQuickStats);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ShowHud error: {ex.Message}");
            }
        }

        public static void ForceOpenHud(bool isDarkTheme = false)
        {
            _userClosed = false;
            ShowHud(isDarkTheme);
        }

        public static void HideHud()
        {
            try
            {
                if (_instance != null)
                {
                    _savedLeft = _instance.Left;
                    _savedTop = _instance.Top;
                    _savedFontSize = _instance.DynamicFontSize;
                    _instance.Hide();
                }
            }
            catch { }
        }

        public static void ToggleHud(bool isDarkTheme = false)
        {
            if (IsHudVisible)
            {
                _userClosed = true;
                HideHud();
            }
            else
            {
                _userClosed = false;
                ShowHud(isDarkTheme);
            }
        }

        public static void UpdateCurrentStats(RulerQuickStats stats)
        {
            if (_instance != null && _instance.IsVisible)
            {
                _instance.Dispatcher.Invoke(() =>
                {
                    _instance.UpdateStats(stats);
                });
            }
        }

        public RulerHudWindow()
        {
            InitializeComponent();
            TxtFontSizeDisplay.Text = DynamicFontSize.ToString("0.#", CultureInfo.InvariantCulture);

            try
            {
                IsDarkTheme = AddInEvents.MainViewModel?.IsDarkTheme ?? true;
            }
            catch { }
        }

        public void UpdateStats(RulerQuickStats stats)
        {
            if (stats == null) return;

            TxtCellBadge.Text = !string.IsNullOrEmpty(stats.CellAddress) ? stats.CellAddress : $"[{stats.RowIndex}, {stats.ColLetter}]";

            // Row Stats
            TxtRowLabel.Text = $"DÒNG {stats.RowIndex}";
            if (stats.RowNumericCount > 0)
            {
                TxtRowSum.Text = FormatNumber(stats.RowSum);
                TxtRowAvg.Text = FormatNumber(stats.RowAvg);
                TxtRowCount.Text = $"{stats.RowNumericCount:N0} (Tổng {stats.RowNonEmptyCount})";
                TxtRowMaxMin.Text = $"{FormatNumber(stats.RowMax)} / {FormatNumber(stats.RowMin)}";
            }
            else
            {
                TxtRowSum.Text = "—";
                TxtRowAvg.Text = "—";
                TxtRowCount.Text = $"{stats.RowNonEmptyCount:N0} ô text";
                TxtRowMaxMin.Text = "—";
            }

            // Col Stats
            TxtColLabel.Text = $"CỘT {stats.ColLetter}";
            if (stats.ColNumericCount > 0)
            {
                TxtColSum.Text = FormatNumber(stats.ColSum);
                TxtColAvg.Text = FormatNumber(stats.ColAvg);
                TxtColCount.Text = $"{stats.ColNumericCount:N0} (Tổng {stats.ColNonEmptyCount})";
                TxtColMaxMin.Text = $"{FormatNumber(stats.ColMax)} / {FormatNumber(stats.ColMin)}";
            }
            else
            {
                TxtColSum.Text = "—";
                TxtColAvg.Text = "—";
                TxtColCount.Text = $"{stats.ColNonEmptyCount:N0} ô text";
                TxtColMaxMin.Text = "—";
            }
        }

        private static string FormatNumber(double val)
        {
            if (Math.Abs(val % 1) < 0.0001)
            {
                return val.ToString("N0", CultureInfo.CurrentCulture);
            }
            return val.ToString("N2", CultureInfo.CurrentCulture);
        }

        private void OnHeaderMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ButtonState == MouseButtonState.Pressed)
            {
                try
                {
                    DragMove();
                    _savedLeft = Left;
                    _savedTop = Top;
                }
                catch { }
            }
        }

        private void OnIncreaseFontSizeClick(object sender, RoutedEventArgs e)
        {
            if (DynamicFontSize < 26)
            {
                DynamicFontSize += 1.5;
                _savedFontSize = DynamicFontSize;
                TxtFontSizeDisplay.Text = DynamicFontSize.ToString("0.#", CultureInfo.InvariantCulture);
            }
        }

        private void OnDecreaseFontSizeClick(object sender, RoutedEventArgs e)
        {
            if (DynamicFontSize > 10.5)
            {
                DynamicFontSize -= 1.5;
                _savedFontSize = DynamicFontSize;
                TxtFontSizeDisplay.Text = DynamicFontSize.ToString("0.#", CultureInfo.InvariantCulture);
            }
        }

        private void OnCloseHudClick(object sender, RoutedEventArgs e)
        {
            _userClosed = true;
            HideHud();
        }
    }
}
