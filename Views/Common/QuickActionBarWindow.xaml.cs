using System;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using ExcelSupport.Host;
using ExcelSupport.Services;
using MouseEventArgs = System.Windows.Input.MouseEventArgs;
using Button = System.Windows.Controls.Button;

namespace ExcelSupport.Views
{
    public partial class QuickActionBarWindow : Window
    {
        #region Win32 Non-Activating Window Setup

        private const int GWL_EXSTYLE = -20;
        private const int WS_EX_NOACTIVATE = 0x08000000;
        private const int WS_EX_TOOLWINDOW = 0x00000080;
        private const int WM_MOUSEACTIVATE = 0x0021;
        private const int MA_NOACTIVATE = 3;

        [DllImport("user32.dll", EntryPoint = "GetWindowLong")]
        private static extern IntPtr GetWindowLong32(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", EntryPoint = "GetWindowLongPtr")]
        private static extern IntPtr GetWindowLongPtr64(IntPtr hWnd, int nIndex);

        private static IntPtr GetWindowLongPtr(IntPtr hWnd, int nIndex)
        {
            return IntPtr.Size == 8 ? GetWindowLongPtr64(hWnd, nIndex) : GetWindowLong32(hWnd, nIndex);
        }

        [DllImport("user32.dll", EntryPoint = "SetWindowLong")]
        private static extern IntPtr SetWindowLong32(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

        [DllImport("user32.dll", EntryPoint = "SetWindowLongPtr")]
        private static extern IntPtr SetWindowLongPtr64(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

        private static IntPtr SetWindowLongPtr(IntPtr hWnd, int nIndex, IntPtr dwNewLong)
        {
            return IntPtr.Size == 8 ? SetWindowLongPtr64(hWnd, nIndex, dwNewLong) : SetWindowLong32(hWnd, nIndex, dwNewLong);
        }

        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);

            try
            {
                var helper = new WindowInteropHelper(this);
                IntPtr hwnd = helper.Handle;

                long exStyle = GetWindowLongPtr(hwnd, GWL_EXSTYLE).ToInt64();
                exStyle |= WS_EX_NOACTIVATE | WS_EX_TOOLWINDOW;
                SetWindowLongPtr(hwnd, GWL_EXSTYLE, new IntPtr(exStyle));

                var source = HwndSource.FromHwnd(hwnd);
                source?.AddHook(WndProc);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[QuickActionBar.OnSourceInitialized] {ex.Message}");
            }
        }

        private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            if (msg == WM_MOUSEACTIVATE)
            {
                handled = true;
                return new IntPtr(MA_NOACTIVATE);
            }
            return IntPtr.Zero;
        }

        #endregion

        #region Dependency Properties & State

        public static readonly DependencyProperty IsDarkThemeProperty =
            DependencyProperty.Register(nameof(IsDarkTheme), typeof(bool), typeof(QuickActionBarWindow),
                new PropertyMetadata(true));

        public bool IsDarkTheme
        {
            get => (bool)GetValue(IsDarkThemeProperty);
            set => SetValue(IsDarkThemeProperty, value);
        }

        private bool _isActionRunning = false;

        #endregion

        public QuickActionBarWindow()
        {
            InitializeComponent();
            Opacity = 0.90;

            try
            {
                IsDarkTheme = AddInEvents.MainViewModel?.IsDarkTheme ?? true;
            }
            catch { }
        }

        #region Hover & Drag Interactions

        private void OnMouseEnter(object sender, MouseEventArgs e)
        {
            Opacity = 1.0;
        }

        private void OnMouseLeave(object sender, MouseEventArgs e)
        {
            if (!_isActionRunning)
            {
                Opacity = 0.88;
            }
        }

        private void OnDragGripMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                try
                {
                    DragMove();
                }
                catch { }
            }
        }

        #endregion

        #region Quick Actions Handlers

        private void OnCopyVisibleClick(object sender, RoutedEventArgs e)
        {
            _ = RunActionAsync(async () =>
            {
                bool success = QuickActionBarService.ExecuteCopyVisibleOnly();
                if (success)
                {
                    await ShowToastAsync("✓ " + LocalizationService.Get("QuickActionBar_Copied", "Đã sao chép ô hiển thị!"), autoClose: true);
                }
            });
        }

        private void OnZenkakuHankakuClick(object sender, RoutedEventArgs e)
        {
            _ = RunActionAsync(async () =>
            {
                bool success = QuickActionBarService.ExecuteConvertZenkakuHankaku(toHankaku: true);
                if (success)
                {
                    await ShowToastAsync("✓ " + LocalizationService.Get("QuickActionBar_ConvertedHankaku", "Đã chuyển sang Bán giác!"), autoClose: true);
                }
            });
        }

        private void OnZenkakuRightClick(object sender, MouseButtonEventArgs e)
        {
            if (sender is Button btn && btn.ContextMenu != null)
            {
                btn.ContextMenu.PlacementTarget = btn;
                btn.ContextMenu.IsOpen = true;
                e.Handled = true;
            }
        }

        private void OnToHankakuMenuItemClick(object sender, RoutedEventArgs e)
        {
            _ = RunActionAsync(async () =>
            {
                bool success = QuickActionBarService.ExecuteConvertZenkakuHankaku(toHankaku: true);
                if (success)
                {
                    await ShowToastAsync("✓ " + LocalizationService.Get("QuickActionBar_ConvertedHankaku", "Đã chuyển sang Bán giác!"), autoClose: true);
                }
            });
        }

        private void OnToZenkakuMenuItemClick(object sender, RoutedEventArgs e)
        {
            _ = RunActionAsync(async () =>
            {
                bool success = QuickActionBarService.ExecuteConvertZenkakuHankaku(toHankaku: false);
                if (success)
                {
                    await ShowToastAsync("✓ " + LocalizationService.Get("QuickActionBar_ConvertedZenkaku", "Đã chuyển sang Toàn giác!"), autoClose: true);
                }
            });
        }

        private void OnAiTranslateClick(object sender, RoutedEventArgs e)
        {
            QuickActionBarService.HideBar();
            QuickActionBarService.ExecuteAiQuickTranslate();
        }

        private void OnExportMarkdownClick(object sender, RoutedEventArgs e)
        {
            _ = RunActionAsync(async () =>
            {
                bool success = QuickActionBarService.ExecuteExportMarkdown();
                if (success)
                {
                    await ShowToastAsync("✓ " + LocalizationService.Get("QuickActionBar_MarkdownExported", "Đã chép bảng Markdown!"), autoClose: true);
                }
            });
        }

        private void OnGenerateSqlClick(object sender, RoutedEventArgs e)
        {
            QuickActionBarService.HideBar();
            SqlScriptGeneratorDialog.ShowWindow(null, IsDarkTheme);
        }

        private void OnTrimSpacesClick(object sender, RoutedEventArgs e)
        {
            _ = RunActionAsync(async () =>
            {
                var (success, count) = QuickActionBarService.ExecuteTrimSpaces();
                if (success)
                {
                    string msg = count > 0
                        ? string.Format(LocalizationService.Get("QuickActionBar_TrimSuccess", "Đã dọn dẹp {0} ô!"), count)
                        : LocalizationService.Get("QuickActionBar_TrimNone", "Dữ liệu đã chuẩn!");
                    await ShowToastAsync("✓ " + msg, autoClose: true);
                }
            });
        }

        private void OnDocTkctClick(object sender, RoutedEventArgs e)
        {
            _ = RunActionAsync(async () =>
            {
                var (success, msg) = QuickActionBarService.ExecuteOpenDetailedDesign();
                if (success)
                {
                    await ShowToastAsync("✓ TKCT", autoClose: true);
                }
            });
        }

        private void OnDocTkcbClick(object sender, RoutedEventArgs e)
        {
            _ = RunActionAsync(async () =>
            {
                var (success, msg) = QuickActionBarService.ExecuteOpenBasicDesign();
                if (success)
                {
                    await ShowToastAsync("✓ TKCB", autoClose: true);
                }
            });
        }

        private void OnDocUtClick(object sender, RoutedEventArgs e)
        {
            _ = RunActionAsync(async () =>
            {
                var (success, msg) = QuickActionBarService.ExecuteOpenTestSpec();
                if (success)
                {
                    await ShowToastAsync("✓ UT", autoClose: true);
                }
            });
        }

        private void OnDismissClick(object sender, RoutedEventArgs e)
        {
            QuickActionBarService.HideBar();
        }

        #endregion

        #region Helpers & Toast Feedback

        private async Task RunActionAsync(Func<Task> action)
        {
            if (_isActionRunning) return;
            _isActionRunning = true;
            try
            {
                await action();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[QuickActionBar Action Error] {ex.Message}");
            }
            finally
            {
                _isActionRunning = false;
            }
        }

        public async Task ShowToastAsync(string message, bool autoClose = true)
        {
            try
            {
                TxtToastMessage.Text = message;
                ActionsPanel.Visibility = Visibility.Collapsed;
                ToastFeedbackBorder.Visibility = Visibility.Visible;

                await Task.Delay(1100);

                if (autoClose)
                {
                    QuickActionBarService.HideBar();
                }
            }
            catch { }
            finally
            {
                ToastFeedbackBorder.Visibility = Visibility.Collapsed;
                ActionsPanel.Visibility = Visibility.Visible;
            }
        }

        #endregion
    }
}
