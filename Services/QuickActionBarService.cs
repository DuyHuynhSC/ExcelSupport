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

    public static class QuickActionBarService
    {
        #region Win32 Native Cursor Methods

        [DllImport("user32.dll")]
        private static extern bool GetCursorPos(out Win32Point lpPoint);

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
                                ShowBarNearCursor();
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

        public static void ShowBarNearCursor()
        {
            if (!AppSettings.IsQuickActionBarEnabled) return;

            try
            {
                if (!GetCursorPos(out Win32Point pt))
                {
                    return;
                }

                // Lấy kích thước màn hình chứa con trỏ chuột
                var screen = Screen.FromPoint(new System.Drawing.Point(pt.X, pt.Y));
                var workArea = screen.WorkingArea;

                // Chiều dài ước tính của thanh tác vụ ~ 450px, chiều cao ~ 46px
                double targetLeft = pt.X + 16;
                double targetTop = pt.Y + 22;

                const double barWidth = 470;
                const double barHeight = 48;

                // Chống tràn viền màn hình bên phải
                if (targetLeft + barWidth > workArea.Right)
                {
                    targetLeft = workArea.Right - barWidth - 10;
                }
                if (targetLeft < workArea.Left)
                {
                    targetLeft = workArea.Left + 10;
                }

                // Chống tràn viền màn hình bên dưới (nổi lên trên con trỏ chuột nếu ở đáy màn hình)
                if (targetTop + barHeight > workArea.Bottom)
                {
                    targetTop = pt.Y - barHeight - 12;
                }
                if (targetTop < workArea.Top)
                {
                    targetTop = workArea.Top + 10;
                }

                if (_barWindow == null)
                {
                    _barWindow = new QuickActionBarWindow();
                }

                _barWindow.IsDarkTheme = AddInEvents.MainViewModel?.IsDarkTheme ?? true;
                _barWindow.Left = targetLeft;
                _barWindow.Top = targetTop;

                if (!_barWindow.IsVisible)
                {
                    _barWindow.Show();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[QuickActionBar.ShowBarNearCursor] {ex.Message}");
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
