using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using ExcelSupport.ViewModels;
using ExcelWindow = Microsoft.Office.Interop.Excel.Window;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using WpfMessageBox = System.Windows.MessageBox;
using WpfMessageBoxButton = System.Windows.MessageBoxButton;
using WpfMessageBoxImage = System.Windows.MessageBoxImage;

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
        private static readonly List<WindowPaneInfo> _panes = new List<WindowPaneInfo>();

        public static event Action<bool>? VisibilityChanged;

        public static bool IsTaskPaneVisible
        {
            get
            {
                var info = GetPaneInfoForActiveWindow();
                return info != null && info.Pane.Visible;
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

                    CleanUpDeadPanes();

                    foreach (var info in _panes)
                    {
                        if (info.WindowHwnd == activeHwnd)
                        {
                            return info;
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
                try
                {
                    var win = item.Pane.Window as ExcelWindow;
                    if (win == null)
                    {
                        _panes.RemoveAt(i);
                        continue;
                    }

                    int testHwnd = win.Hwnd;
                }
                catch
                {
                    try { item.Pane.Delete(); } catch { }
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
                            // Lưu lại trạng thái người dùng (kể cả khi bấm dấu X trên Task Pane)
                            AppSettings.IsTaskPaneAutoOpen = ctp.Visible;
                            VisibilityChanged?.Invoke(ctp.Visible);
                        };

                        _panes.Add(new WindowPaneInfo
                        {
                            WindowHwnd = activeHwnd,
                            Pane = newPane,
                            HostControl = hostControl
                        });

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
                WpfMessageBox.Show($"Không thể tạo Task Pane cho cửa sổ này:\n{ex.Message}",
                                   "Excel-DNA Task Pane", WpfMessageBoxButton.OK, WpfMessageBoxImage.Error);
            }

            return null;
        }

        public static void ToggleTaskPane(TaskPaneViewModel viewModel, bool show)
        {
            var pane = EnsureCreatedForActiveWindow(viewModel);
            if (pane != null)
            {
                pane.Visible = show;
                AppSettings.IsTaskPaneAutoOpen = show;
                VisibilityChanged?.Invoke(show);
            }
        }

        public static void AutoRestoreForActiveWindow(TaskPaneViewModel viewModel)
        {
            if (AppSettings.IsTaskPaneAutoOpen)
            {
                var pane = EnsureCreatedForActiveWindow(viewModel);
                if (pane != null && !pane.Visible)
                {
                    pane.Visible = true;
                    VisibilityChanged?.Invoke(true);
                }
            }
        }

        public static void DetachTaskPane()
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
