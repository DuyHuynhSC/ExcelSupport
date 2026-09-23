using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using System.Windows.Forms;
using ExcelSupport.Host;
using ExcelSupport.Models;
using ExcelSupport.Views;
using Microsoft.Office.Interop.Excel;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using WpfApplication = System.Windows.Application;
using SysAction = System.Action;

namespace ExcelSupport.Services
{
    [StructLayout(LayoutKind.Sequential)]
    public struct Win32Point
    {
        public int X;
        public int Y;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct Win32Rect
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct Win32MonitorInfo
    {
        public int cbSize;
        public Win32Rect rcMonitor;
        public Win32Rect rcWork;
        public uint dwFlags;
    }

    public static class QuickActionBarService
    {
        #region Win32 Native Cursor & Window Positioning Methods

        [DllImport("user32.dll")]
        private static extern bool GetCursorPos(out Win32Point lpPoint);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        [DllImport("user32.dll")]
        private static extern IntPtr MonitorFromPoint(Win32Point pt, uint dwFlags);

        [DllImport("user32.dll")]
        private static extern bool GetMonitorInfo(IntPtr hMonitor, ref Win32MonitorInfo lpmi);

        private const uint MONITOR_DEFAULTTONEAREST = 2;
        private static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
        private const uint SWP_NOSIZE = 0x0001;
        private const uint SWP_NOACTIVATE = 0x0010;
        private const uint SWP_SHOWWINDOW = 0x0040;

        #endregion

        private static QuickActionBarWindow? _barWindow;
        private static System.Threading.Timer? _debounceTimer;
        private static readonly object _lock = new object();
        private static bool _isProcessing = false;

        public static bool IsBarVisible => _barWindow != null && _barWindow.IsVisible;

        #region Event Hooking & Triggering

        public static void OnSheetSelectionChange(_Worksheet? ws, Range? target)
        {
            if (!AppSettings.IsQuickActionBarEnabled)
            {
                HideBar();
                return;
            }

            if (ws == null || target == null || _isProcessing)
            {
                return;
            }

            try
            {
                long cellCount = 0;
                try
                {
                    cellCount = target.CountLarge;
                }
                catch
                {
                    cellCount = target.Count;
                }

                // Nếu chỉ chọn 1 ô đơn lẻ -> ẩn thanh tác vụ
                if (cellCount <= 1)
                {
                    HideBar();
                    return;
                }

                // Quét chọn từ 2 ô trở lên -> Kích hoạt debounce timer 250ms
                lock (_lock)
                {
                    _debounceTimer?.Dispose();
                    _debounceTimer = new System.Threading.Timer(_ =>
                    {
                        try
                        {
                            var app = AddInEvents.Instance?.ExcelAppInstance;
                            if (app == null) return;

                            // Đảm bảo thao tác UI trên Dispatcher của WPF
                            WpfApplication.Current?.Dispatcher.BeginInvoke(new SysAction(() =>
                            {
                                ShowBarNearSelectionOrCursor(target);
                            }));
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"[QuickActionBar.DebounceTimer] {ex.Message}");
                        }
                    }, null, 250, Timeout.Infinite);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[QuickActionBar.OnSheetSelectionChange] {ex.Message}");
            }
        }

        public static void ShowBarNearSelectionOrCursor(Range? targetRange = null)
        {
            if (!AppSettings.IsQuickActionBarEnabled) return;

            try
            {
                var app = AddInEvents.Instance?.ExcelAppInstance;
                if (app == null) return;

                // 1. Lấy DPI scale của hệ thống / màn hình (DIP = PhysicalPixels / dpiScale)
                double dpiScaleX = 1.0;
                double dpiScaleY = 1.0;

                if (_barWindow != null && _barWindow.IsLoaded)
                {
                    try
                    {
                        var dpi = System.Windows.Media.VisualTreeHelper.GetDpi(_barWindow);
                        dpiScaleX = dpi.DpiScaleX;
                        dpiScaleY = dpi.DpiScaleY;
                    }
                    catch { }
                }

                if (dpiScaleX <= 0.1 || dpiScaleY <= 0.1)
                {
                    try
                    {
                        using (var g = System.Drawing.Graphics.FromHwnd(IntPtr.Zero))
                        {
                            dpiScaleX = g.DpiX / 96.0;
                            dpiScaleY = g.DpiY / 96.0;
                        }
                    }
                    catch { }
                }

                if (dpiScaleX <= 0.1) dpiScaleX = 1.0;
                if (dpiScaleY <= 0.1) dpiScaleY = 1.0;

                // 2. Xác định tọa độ (Physical Screen Pixels) của vùng chọn Excel
                int selLeftPx = -1;
                int selTopPx = -1;
                int selRightPx = -1;
                int selBottomPx = -1;
                bool hasRangeCoords = false;

                Range? rng = targetRange;
                if (rng == null)
                {
                    try { rng = app.Selection as Range; } catch { }
                }

                if (rng != null)
                {
                    try
                    {
                        var win = app.ActiveWindow;
                        if (win != null)
                        {
                            // Tọa độ gốc ô (0,0) của worksheet pane hiện tại trên màn hình
                            int originX = win.PointsToScreenPixelsX(0);
                            int originY = win.PointsToScreenPixelsY(0);

                            // Hệ số quy đổi từ Excel Points sang Physical Screen Pixels:
                            // 1 point = 1/72 inch. Tại 96 DPI: 1 point = 96/72 = 1.33333 physical pixels.
                            // Tính đúng hệ số Zoom của ActiveWindow và DPI Scale của màn hình.
                            double zoom = 1.0;
                            try
                            {
                                if (win.Zoom is int z && z > 0)
                                {
                                    zoom = z / 100.0;
                                }
                            }
                            catch { }

                            double ptToPxX = (dpiScaleX * 96.0) / 72.0 * zoom;
                            double ptToPxY = (dpiScaleY * 96.0) / 72.0 * zoom;

                            double rLeft = (double)rng.Left;
                            double rTop = (double)rng.Top;
                            double rWidth = (double)rng.Width;
                            double rHeight = (double)rng.Height;

                            selLeftPx = originX + (int)Math.Round(rLeft * ptToPxX);
                            selTopPx = originY + (int)Math.Round(rTop * ptToPxY);
                            selRightPx = selLeftPx + (int)Math.Round(rWidth * ptToPxX);
                            selBottomPx = selTopPx + (int)Math.Round(rHeight * ptToPxY);

                            hasRangeCoords = (selLeftPx >= 0 && selBottomPx >= 0);
                        }
                    }
                    catch { }
                }

                // 3. Kích thước thanh nổi (Physical Pixels)
                const double defaultBarWidthDip = 500;
                const double defaultBarHeightDip = 48;
                double barWidthDip = (_barWindow != null && _barWindow.ActualWidth > 100) ? _barWindow.ActualWidth : defaultBarWidthDip;
                double barHeightDip = (_barWindow != null && _barWindow.ActualHeight > 20) ? _barWindow.ActualHeight : defaultBarHeightDip;
                double barWidthPx = barWidthDip * dpiScaleX;
                double barHeightPx = barHeightDip * dpiScaleY;

                // 4. Tính toán vị trí hiển thị (Physical Pixels)
                int physicalX = -1;
                int physicalY = -1;

                bool hasMouse = GetCursorPos(out Win32Point mousePt);

                if (hasRangeCoords)
                {
                    // Canh lề ngang: Mặc định canh theo cạnh trái của vùng chọn
                    physicalX = selLeftPx;

                    // Nếu con trỏ chuột nằm gần vùng chọn và vùng chọn rộng hơn thanh bar:
                    // Di chuyển thanh bar theo ngang gần con trỏ chuột để người dùng dễ click
                    if (hasMouse && (selRightPx - selLeftPx) > barWidthPx)
                    {
                        int idealX = mousePt.X - (int)(barWidthPx / 2);
                        physicalX = Math.Max(selLeftPx, Math.Min(idealX, selRightPx - (int)barWidthPx));
                    }

                    // Canh lề dọc: Hiển thị NGAY BÊN DƯỚI vùng chọn (cách 4px)
                    physicalY = selBottomPx + (int)(4 * dpiScaleY);
                }
                else if (hasMouse)
                {
                    // Fallback theo con trỏ chuột
                    physicalX = mousePt.X + (int)(12 * dpiScaleX);
                    physicalY = mousePt.Y + (int)(16 * dpiScaleY);
                }

                if (physicalX < 0 || physicalY < 0) return;

                // 5. Lấy diện tích làm việc chuẩn của màn hình (Physical Pixels) qua Win32 Monitor Info
                Win32Rect workArea = new Win32Rect { Left = 0, Top = 0, Right = 1920, Bottom = 1040 };
                try
                {
                    IntPtr hMonitor = MonitorFromPoint(new Win32Point { X = physicalX, Y = physicalY }, MONITOR_DEFAULTTONEAREST);
                    if (hMonitor != IntPtr.Zero)
                    {
                        Win32MonitorInfo mi = new Win32MonitorInfo();
                        mi.cbSize = Marshal.SizeOf(typeof(Win32MonitorInfo));
                        if (GetMonitorInfo(hMonitor, ref mi))
                        {
                            workArea = mi.rcWork;
                        }
                    }
                }
                catch
                {
                    try
                    {
                        var screen = Screen.FromPoint(new System.Drawing.Point(physicalX, physicalY));
                        workArea = new Win32Rect
                        {
                            Left = screen.WorkingArea.Left,
                            Top = screen.WorkingArea.Top,
                            Right = screen.WorkingArea.Right,
                            Bottom = screen.WorkingArea.Bottom
                        };
                    }
                    catch { }
                }

                // Chống tràn viền phải / trái màn hình
                if (physicalX + barWidthPx > workArea.Right - (10 * dpiScaleX))
                {
                    physicalX = (int)(workArea.Right - barWidthPx - (10 * dpiScaleX));
                }
                if (physicalX < workArea.Left + (10 * dpiScaleX))
                {
                    physicalX = (int)(workArea.Left + (10 * dpiScaleX));
                }

                // Chống tràn đáy màn hình:
                // Nếu vùng chọn nằm sát đáy màn hình không đủ chỗ bên dưới -> Đảo lên NGAY TRÊN ĐẦU vùng chọn (cách 4px)
                if (physicalY + barHeightPx > workArea.Bottom - (10 * dpiScaleY))
                {
                    if (hasRangeCoords)
                    {
                        physicalY = selTopPx - (int)barHeightPx - (int)(4 * dpiScaleY);
                    }
                    else if (hasMouse)
                    {
                        physicalY = mousePt.Y - (int)barHeightPx - (int)(12 * dpiScaleY);
                    }
                }

                // Chống tràn mép trên màn hình
                if (physicalY < workArea.Top + (10 * dpiScaleY))
                {
                    physicalY = (int)(workArea.Top + (10 * dpiScaleY));
                }

                // 6. QUY ĐỔI SANG WPF DIPs (Device-Independent Pixels: wpf = physical / dpiScale)
                double wpfLeft = physicalX / dpiScaleX;
                double wpfTop = physicalY / dpiScaleY;

                if (_barWindow == null)
                {
                    _barWindow = new QuickActionBarWindow();
                }

                _barWindow.IsDarkTheme = AddInEvents.MainViewModel?.IsDarkTheme ?? false;
                _barWindow.Left = wpfLeft;
                _barWindow.Top = wpfTop;

                if (!_barWindow.IsVisible)
                {
                    _barWindow.Show();
                }

                try
                {
                    var helper = new System.Windows.Interop.WindowInteropHelper(_barWindow);
                    IntPtr hwnd = helper.EnsureHandle();
                    SetWindowPos(hwnd, HWND_TOPMOST, physicalX, physicalY, 0, 0, SWP_NOSIZE | SWP_NOACTIVATE | SWP_SHOWWINDOW);
                }
                catch { }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[QuickActionBar.ShowBarNearSelectionOrCursor] {ex.Message}");
            }
        }

        public static void HideBar()
        {
            try
            {
                lock (_lock)
                {
                    _debounceTimer?.Dispose();
                    _debounceTimer = null;
                }

                if (_barWindow != null && _barWindow.IsVisible)
                {
                    WpfApplication.Current?.Dispatcher.BeginInvoke(new SysAction(() =>
                    {
                        try
                        {
                            _barWindow?.Hide();
                        }
                        catch { }
                    }));
                }
            }
            catch { }
        }

        public static void CloseBar()
        {
            try
            {
                HideBar();
                _barWindow?.Close();
                _barWindow = null;
            }
            catch { }
        }

        public static void ToggleEnabled()
        {
            AppSettings.IsQuickActionBarEnabled = !AppSettings.IsQuickActionBarEnabled;
            if (!AppSettings.IsQuickActionBarEnabled)
            {
                HideBar();
            }
        }

        #endregion

        #region Action Execution Implementations

        public static bool ExecuteCopyVisibleOnly()
        {
            try
            {
                var app = AddInEvents.Instance?.ExcelAppInstance;
                if (app == null) return false;

                var res = FilteredCopyPasteService.CopyVisibleCells(app);
                return res.Success;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[QuickActionBar.ExecuteCopyVisibleOnly] {ex.Message}");
                return false;
            }
        }

        public static bool ExecuteConvertZenkakuHankaku(bool toHankaku)
        {
            try
            {
                var app = AddInEvents.Instance?.ExcelAppInstance;
                if (app == null) return false;

                var options = new JapaneseConversionOptions
                {
                    ToHankaku = toHankaku,
                    ConvertAlpha = true,
                    ConvertNumbers = true,
                    ConvertKatakana = true,
                    ConvertPunctuation = true,
                    ConvertSpace = true,
                    Scope = ConversionScope.Selection
                };

                var res = JapaneseTextConverterService.ExecuteConversion(app, options);
                return res != null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[QuickActionBar.ExecuteConvertZenkakuHankaku] {ex.Message}");
                return false;
            }
        }

        public static void ExecuteAiQuickTranslate()
        {
            try
            {
                AiQuickTranslatePopup.ShowPopup(AddInEvents.MainViewModel?.IsDarkTheme ?? false);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[QuickActionBar.ExecuteAiQuickTranslate] {ex.Message}");
            }
        }

        public static bool ExecuteExportMarkdown()
        {
            try
            {
                var app = AddInEvents.Instance?.ExcelAppInstance;
                if (app == null) return false;

                return TableExportService.QuickCopySelectionToMarkdown(app);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[QuickActionBar.ExecuteExportMarkdown] {ex.Message}");
                return false;
            }
        }

        public static (bool Success, int ModifiedCount) ExecuteTrimSpaces()
        {
            var app = AddInEvents.Instance?.ExcelAppInstance;
            if (app == null) return (false, 0);

            bool prevScreenUpdating = app.ScreenUpdating;
            try
            {
                _isProcessing = true;
                app.ScreenUpdating = false;

                dynamic sel = app.Selection;
                if (sel is not Range rng) return (false, 0);

                int rCount = rng.Rows.Count;
                int cCount = rng.Columns.Count;
                if (rCount <= 0 || cCount <= 0) return (false, 0);

                object? rawValues = rng.Value2;
                int modifiedCount = 0;

                if (rawValues is object[,] valArray)
                {
                    object[,] newArray = new object[rCount, cCount];
                    bool hasChanged = false;

                    for (int r = 1; r <= rCount; r++)
                    {
                        for (int c = 1; c <= cCount; c++)
                        {
                            object? val = valArray[r, c];
                            if (val is string str && !string.IsNullOrEmpty(str))
                            {
                                string trimmed = CleanWhitespace(str);
                                if (trimmed != str)
                                {
                                    newArray[r - 1, c - 1] = trimmed;
                                    modifiedCount++;
                                    hasChanged = true;
                                    continue;
                                }
                            }
                            newArray[r - 1, c - 1] = val!;
                        }
                    }

                    if (hasChanged)
                    {
                        rng.Value2 = newArray;
                    }
                }
                else if (rawValues is string singleStr && !string.IsNullOrEmpty(singleStr))
                {
                    string trimmed = CleanWhitespace(singleStr);
                    if (trimmed != singleStr)
                    {
                        rng.Value2 = trimmed;
                        modifiedCount = 1;
                    }
                }

                try
                {
                    app.StatusBar = $"✨ ExcelSupport: Đã cắt tỉa khoảng trắng cho {modifiedCount} ô tính!";
                }
                catch { }

                return (true, modifiedCount);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[QuickActionBar.ExecuteTrimSpaces] {ex.Message}");
                return (false, 0);
            }
            finally
            {
                try { app.ScreenUpdating = prevScreenUpdating; } catch { }
                _isProcessing = false;
            }
        }

        private static string CleanWhitespace(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;

            // Xóa non-breaking space & ký tự điều khiển
            string res = text.Replace("\u00A0", " ").Replace("&nbsp;", " ");
            res = res.Trim();

            // Rút gọn nhiều khoảng trắng liên tiếp giữa các từ thành 1 khoảng trắng
            while (res.Contains("  "))
            {
                res = res.Replace("  ", " ");
            }

            return res;
        }

        public static (bool Success, string Message) ExecuteOpenDetailedDesign()
        {
            try
            {
                var app = AddInEvents.Instance?.ExcelAppInstance;
                if (app == null) return (false, "Excel is not ready");

                var res = ProjectDocumentLauncherService.LaunchFromSelection(app, SpecDocumentType.DetailedDesign);
                return (res.Success, res.Message);
            }
            catch (Exception ex)
            {
                return (false, ex.Message);
            }
        }

        public static (bool Success, string Message) ExecuteOpenBasicDesign()
        {
            try
            {
                var app = AddInEvents.Instance?.ExcelAppInstance;
                if (app == null) return (false, "Excel is not ready");

                var res = ProjectDocumentLauncherService.LaunchFromSelection(app, SpecDocumentType.BasicDesign);
                return (res.Success, res.Message);
            }
            catch (Exception ex)
            {
                return (false, ex.Message);
            }
        }

        public static (bool Success, string Message) ExecuteOpenTestSpec()
        {
            try
            {
                var app = AddInEvents.Instance?.ExcelAppInstance;
                if (app == null) return (false, "Excel is not ready");

                var res = ProjectDocumentLauncherService.LaunchFromSelection(app, SpecDocumentType.TestSpec);
                return (res.Success, res.Message);
            }
            catch (Exception ex)
            {
                return (false, ex.Message);
            }
        }

        #endregion
    }
}
