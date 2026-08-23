using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using ExcelSupport.ViewModels;
using ExcelWindow = Microsoft.Office.Interop.Excel.Window;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;

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
        private static readonly object _lock = new object();

        public static event Action<bool>? VisibilityChanged;

        public static bool IsTaskPaneVisible
        {
            get
            {
                var info = GetPaneInfoForActiveWindow();
                if (info == null) return false;
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
                if (!IsWindow((IntPtr)item.WindowHwnd))
                {
                    try
                    {
                        item.Pane.Visible = false;
                    }
                    catch { }
                    _panes.RemoveAt(i);
                }
            }
        }

        public static CustomTaskPane? EnsureCreatedForActiveWindow(TaskPaneViewModel? viewModel)
        {
            var existingInfo = GetPaneInfoForActiveWindow();
            if (existingInfo != null)
            {
                if (viewModel != null)
                {
                    existingInfo.HostControl.BindViewModel(viewModel);
                }
                return existingInfo.Pane;
            }

            try
            {
                var app = (ExcelApp)ExcelDnaUtil.Application;
                ExcelWindow? activeWindow = null;

                try
                {
                    activeWindow = app.ActiveWindow;
                    if (activeWindow == null) return null;

                    int activeHwnd = activeWindow.Hwnd;

                    var hostControl = new TaskPaneHostControl();
                    if (viewModel != null)
                    {
                        hostControl.BindViewModel(viewModel);
                    }

                    var newPane = CustomTaskPaneFactory.CreateCustomTaskPane(hostControl, "Workbook Navigator", activeWindow);
                    if (newPane != null)
                    {
                        newPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionLeft;
                        newPane.Width = 320;

                        newPane.VisibleStateChange += ctp =>
                        {
                            try
                            {
                                var currentActiveInfo = GetPaneInfoForActiveWindow();
                                if (currentActiveInfo != null && currentActiveInfo.Pane == ctp)
                                {
                                    AppSettings.IsTaskPaneAutoOpen = ctp.Visible;
                                    VisibilityChanged?.Invoke(ctp.Visible);
                                }
                            }
                            catch { }
                        };

                        lock (_lock)
                        {
                            _panes.Add(new WindowPaneInfo
                            {
                                WindowHwnd = activeHwnd,
                                Pane = newPane,
                                HostControl = hostControl
                            });
                        }

                        return newPane;
                    }
                }
                finally
                {
                    if (activeWindow != null) Marshal.ReleaseComObject(activeWindow);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error creating CustomTaskPane: {ex.Message}");
            }

            return null;
        }

        public static void ToggleTaskPane(TaskPaneViewModel viewModel, bool show)
        {
            AppSettings.IsTaskPaneAutoOpen = show;
            var pane = EnsureCreatedForActiveWindow(viewModel);
            if (pane != null)
            {
                try
                {
                    if (pane.Visible != show)
                    {
                        pane.Visible = show;
                    }
                }
                catch { }
                VisibilityChanged?.Invoke(show);
            }
        }

        public static void AutoRestoreForActiveWindow(TaskPaneViewModel viewModel)
        {
            if (!AppSettings.IsTaskPaneAutoOpen) return;

            var pane = EnsureCreatedForActiveWindow(viewModel);
            if (pane != null)
            {
                try
                {
                    if (!pane.Visible)
                    {
                        pane.Visible = true;
                        VisibilityChanged?.Invoke(true);
                    }
                }
                catch { }
            }
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
