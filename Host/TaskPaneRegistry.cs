using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using ExcelSupport.ViewModels;
using ExcelWindow = Microsoft.Office.Interop.Excel.Window;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using ExcelWorkbook = Microsoft.Office.Interop.Excel.Workbook;

namespace ExcelSupport.Host
{
    public class WindowPaneInfo
    {
        public int WindowHwnd { get; set; }
        public CustomTaskPane Pane { get; set; } = null!;
        public TaskPaneHostControl HostControl { get; set; } = null!;
    }

    public static class TaskPaneRegistry
    {
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsWindow(IntPtr hWnd);

        private static readonly List<WindowPaneInfo> _panes = new List<WindowPaneInfo>();
        private static readonly HashSet<int> _closingWindowHwnds = new HashSet<int>();
        private static readonly object _lock = new object();
        private static bool _isToggling = false;

        public static event Action<bool>? VisibilityChanged;

        public static bool IsTaskPaneVisible
        {
            get
            {
                var info = GetPaneInfoForActiveWindow();
                if (info == null) return AppSettings.IsTaskPaneAutoOpen;
                try
                {
                    return info.Pane.Visible;
                }
                catch
                {
                    return false;
                }
            }
        }

        private static WindowPaneInfo? GetPaneInfoForActiveWindow()
        {
            try
            {
                var app = (ExcelApp)ExcelDnaUtil.Application;
                ExcelWindow? activeWindow = null;
                try
                {
                    activeWindow = app.ActiveWindow;
                    if (activeWindow == null) return null;

                    int activeHwnd = activeWindow.Hwnd;

                    lock (_lock)
                    {
                        CleanUpDeadPanes();

                        foreach (var info in _panes)
                        {
                            if (info.WindowHwnd == activeHwnd)
                            {
                                return info;
                            }
                        }
                    }
                }
                finally
                {
                    if (activeWindow != null) Marshal.ReleaseComObject(activeWindow);
                }
            }
            catch { }

            return null;
        }

        private static void CleanUpDeadPanes()
        {
            for (int i = _panes.Count - 1; i >= 0; i--)
            {
                var item = _panes[i];
                bool isDead = false;
                try
                {
                    if (!IsWindow((IntPtr)item.WindowHwnd) || _closingWindowHwnds.Contains(item.WindowHwnd))
                    {
                        isDead = true;
                    }
                }
                catch
                {
                    isDead = true;
                }

                if (isDead)
                {
                    _panes.RemoveAt(i);
                }
            }
        }

        private static bool IsWindowAliveInExcel(int hwnd)
        {
            if (hwnd == 0 || !IsWindow((IntPtr)hwnd)) return false;

            try
            {
                var app = (ExcelApp)ExcelDnaUtil.Application;
                Microsoft.Office.Interop.Excel.Windows? windows = null;
                try
                {
                    windows = app.Windows;
                    if (windows == null) return false;

                    int count = windows.Count;
                    for (int i = 1; i <= count; i++)
                    {
                        ExcelWindow? win = null;
                        try
                        {
                            win = windows[i];
                            if (win != null && win.Hwnd == hwnd)
                            {
                                return true;
                            }
                        }
                        catch { }
                        finally
                        {
                            if (win != null) Marshal.ReleaseComObject(win);
                        }
                    }
                }
                finally
                {
                    if (windows != null) Marshal.ReleaseComObject(windows);
                }
            }
            catch { }

            return false;
        }

        public static CustomTaskPane? EnsureCreatedForWindow(ExcelWindow? window, TaskPaneViewModel? viewModel)
        {
            if (window == null) return null;
            int hwnd = 0;
            try
            {
                hwnd = window.Hwnd;
            }
            catch
            {
                return null;
            }

            lock (_lock)
            {
                CleanUpDeadPanes();

                foreach (var info in _panes)
                {
                    if (info.WindowHwnd == hwnd)
                    {
                        if (viewModel != null)
                        {
                            info.HostControl.BindViewModel(viewModel);
                        }
                        return info.Pane;
                    }
                }
            }

            try
            {
                var hostControl = new TaskPaneHostControl();
                if (viewModel != null)
                {
                    hostControl.BindViewModel(viewModel);
                }

                var newPane = CustomTaskPaneFactory.CreateCustomTaskPane(hostControl, "Workbook Navigator", window);
                if (newPane != null)
                {
                    newPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionLeft;
                    newPane.Width = 320;

                    newPane.VisibleStateChange += ctp =>
                    {
                        if (_isToggling) return;
                        try
                        {
                            bool isVisible = ctp.Visible;
                            if (isVisible)
                            {
                                AppSettings.IsTaskPaneAutoOpen = true;
                                VisibilityChanged?.Invoke(true);
                                return;
                            }

                            // Khi isVisible == false:
                            // Kiểm tra xem TaskPane bị đóng do cửa sổ/Workbook đang đóng hay do người dùng bấm 'X'
                            lock (_lock)
                            {
                                if (_closingWindowHwnds.Contains(hwnd) || !IsWindow((IntPtr)hwnd))
                                {
                                    CleanUpDeadPanes();
                                    return;
                                }
                            }

                            // Chờ kiểm tra trễ qua QueueAsMacro để Excel hoàn tất việc giải phóng/đóng cửa sổ (nếu có)
                            ExcelAsyncUtil.QueueAsMacro(() =>
                            {
                                if (_isToggling) return;

                                lock (_lock)
                                {
                                    if (_closingWindowHwnds.Contains(hwnd))
                                    {
                                        CleanUpDeadPanes();
                                        return;
                                    }
                                }

                                if (!IsWindowAliveInExcel(hwnd))
                                {
                                    lock (_lock)
                                    {
                                        CleanUpDeadPanes();
                                    }
                                    return;
                                }

                                // Cửa sổ vẫn đang mở bình thường trong Excel -> người dùng thật sự bấm nút 'X' để tắt TaskPane
                                try
                                {
                                    if (!ctp.Visible)
                                    {
                                        AppSettings.IsTaskPaneAutoOpen = false;
                                        VisibilityChanged?.Invoke(false);
                                        HideAllPanes();
                                    }
                                }
                                catch { }
                            });
                        }
                        catch { }
                    };

                    lock (_lock)
                    {
                        _panes.Add(new WindowPaneInfo
                        {
                            WindowHwnd = hwnd,
                            Pane = newPane,
                            HostControl = hostControl
                        });
                    }

                    return newPane;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error creating CustomTaskPane for window {hwnd}: {ex.Message}");
            }

            return null;
        }

        public static CustomTaskPane? EnsureCreatedForActiveWindow(TaskPaneViewModel? viewModel)
        {
            try
            {
                var app = (ExcelApp)ExcelDnaUtil.Application;
                ExcelWindow? activeWindow = null;
                try
                {
                    activeWindow = app.ActiveWindow;
                    if (activeWindow != null)
                    {
                        return EnsureCreatedForWindow(activeWindow, viewModel);
                    }
                }
                finally
                {
                    if (activeWindow != null) Marshal.ReleaseComObject(activeWindow);
                }
            }
            catch { }

            return null;
        }

        public static void ToggleTaskPane(TaskPaneViewModel viewModel, bool show)
        {
            _isToggling = true;
            try
            {
                AppSettings.IsTaskPaneAutoOpen = show;

                if (show)
                {
                    // 1. Duyệt qua tất cả các cửa sổ Excel đang mở để tạo và hiển thị TaskPane
                    try
                    {
                        var app = (ExcelApp)ExcelDnaUtil.Application;
                        Microsoft.Office.Interop.Excel.Windows? windows = null;
                        try
                        {
                            windows = app.Windows;
                            if (windows != null)
                            {
                                int count = windows.Count;
                                for (int i = 1; i <= count; i++)
                                {
                                    ExcelWindow? win = null;
                                    try
                                    {
                                        win = windows[i];
                                        if (win != null)
                                        {
                                            bool isWinVisible = false;
                                            try { isWinVisible = win.Visible; } catch { isWinVisible = true; }

                                            if (isWinVisible)
                                            {
                                                var pane = EnsureCreatedForWindow(win, viewModel);
                                                if (pane != null && !pane.Visible)
                                                {
                                                    pane.Visible = true;
                                                }
                                            }
                                        }
                                    }
                                    catch { }
                                    finally
                                    {
                                        if (win != null) Marshal.ReleaseComObject(win);
                                    }
                                }
                            }
                        }
                        finally
                        {
                            if (windows != null) Marshal.ReleaseComObject(windows);
                        }
                    }
                    catch { }

                    // 2. Đảm bảo các pane đã có trong danh sách cũng được hiển thị
                    lock (_lock)
                    {
                        CleanUpDeadPanes();
                        foreach (var info in _panes)
                        {
                            try
                            {
                                if (!info.Pane.Visible)
                                {
                                    info.Pane.Visible = true;
                                }
                            }
                            catch { }
                        }
                    }
                }
                else
                {
                    HideAllPanesInternal();
                }

                VisibilityChanged?.Invoke(show);
            }
            finally
            {
                _isToggling = false;
            }
        }

        public static void ToggleTaskPaneAuto()
        {
            if (AddInEvents.MainViewModel != null)
            {
                bool newVisibleState = !IsTaskPaneVisible;
                ToggleTaskPane(AddInEvents.MainViewModel, newVisibleState);
            }
        }

        public static void HideAllPanes()
        {
            _isToggling = true;
            try
            {
                AppSettings.IsTaskPaneAutoOpen = false;
                HideAllPanesInternal();
                VisibilityChanged?.Invoke(false);
            }
            finally
            {
                _isToggling = false;
            }
        }

        private static void HideAllPanesInternal()
        {
            lock (_lock)
            {
                CleanUpDeadPanes();
                foreach (var info in _panes)
                {
                    try
                    {
                        if (info.Pane.Visible)
                        {
                            info.Pane.Visible = false;
                        }
                    }
                    catch { }
                }
            }
        }

        public static void AutoRestoreForWindow(ExcelWindow? window, TaskPaneViewModel? viewModel)
        {
            if (!AppSettings.IsTaskPaneAutoOpen || window == null) return;

            try
            {
                bool isWinVisible = false;
                try { isWinVisible = window.Visible; } catch { isWinVisible = true; }
                if (!isWinVisible) return;

                var pane = EnsureCreatedForWindow(window, viewModel);
                if (pane != null && !pane.Visible)
                {
                    pane.Visible = true;
                    VisibilityChanged?.Invoke(true);
                }
            }
            catch { }
        }

        public static void AutoRestoreForActiveWindow(TaskPaneViewModel? viewModel)
        {
            if (!AppSettings.IsTaskPaneAutoOpen) return;

            try
            {
                var pane = EnsureCreatedForActiveWindow(viewModel);
                if (pane != null && !pane.Visible)
                {
                    pane.Visible = true;
                    VisibilityChanged?.Invoke(true);
                }
            }
            catch { }
        }

        public static void OnWorkbookBeforeClose(ExcelWorkbook? wb)
        {
            if (wb == null) return;
            try
            {
                string wbName = "";
                try { wbName = wb.Name; } catch { }

                var windows = wb.Windows;
                if (windows != null)
                {
                    int count = windows.Count;
                    var hwnds = new List<int>();
                    for (int i = 1; i <= count; i++)
                    {
                        ExcelWindow? win = null;
                        try
                        {
                            win = windows[i];
                            if (win != null)
                            {
                                int hwnd = win.Hwnd;
                                hwnds.Add(hwnd);
                            }
                        }
                        catch { }
                        finally
                        {
                            if (win != null) Marshal.ReleaseComObject(win);
                        }
                    }
                    Marshal.ReleaseComObject(windows);

                    lock (_lock)
                    {
                        foreach (var h in hwnds)
                        {
                            _closingWindowHwnds.Add(h);
                        }
                    }

                    // Nếu người dùng bấm Cancel khi hộp thoại hỏi lưu xuất hiện,
                    // sau khi Excel quay lại trạng thái nhàn rỗi (idle), kiểm tra lại để giải phóng hwnd
                    ExcelAsyncUtil.QueueAsMacro(() =>
                    {
                        lock (_lock)
                        {
                            bool stillOpen = false;
                            if (!string.IsNullOrEmpty(wbName))
                            {
                                try
                                {
                                    var app = (ExcelApp)ExcelDnaUtil.Application;
                                    var workbooks = app.Workbooks;
                                    if (workbooks != null)
                                    {
                                        try
                                        {
                                            int wbCount = workbooks.Count;
                                            for (int j = 1; j <= wbCount; j++)
                                            {
                                                ExcelWorkbook? openWb = null;
                                                try
                                                {
                                                    openWb = workbooks[j];
                                                    if (openWb != null && string.Equals(openWb.Name, wbName, StringComparison.OrdinalIgnoreCase))
                                                    {
                                                        stillOpen = true;
                                                        break;
                                                    }
                                                }
                                                catch { }
                                                finally
                                                {
                                                    if (openWb != null) Marshal.ReleaseComObject(openWb);
                                                }
                                            }
                                        }
                                        finally
                                        {
                                            Marshal.ReleaseComObject(workbooks);
                                        }
                                    }
                                }
                                catch { }
                            }

                            if (stillOpen)
                            {
                                // Người dùng đã hủy đóng (Cancel)
                                foreach (var h in hwnds)
                                {
                                    _closingWindowHwnds.Remove(h);
                                }
                            }
                            else
                            {
                                // Đã đóng thật sự -> dọn dẹp các pane của các hwnd này
                                CleanUpDeadPanes();
                                foreach (var h in hwnds)
                                {
                                    _closingWindowHwnds.Remove(h);
                                }
                            }
                        }
                    });
                }
            }
            catch { }
        }

        public static void DetachTaskPane()
        {
            lock (_lock)
            {
                foreach (var info in _panes.ToArray())
                {
                    try
                    {
                        info.Pane.Visible = false;
                        info.Pane.Delete();
                    }
                    catch { }
                }
                _panes.Clear();
            }
        }
    }
}
