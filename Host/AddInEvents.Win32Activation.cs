using System;
using System.Drawing;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using WpfMessageBox = System.Windows.MessageBox;

namespace ExcelSupport
{
    public partial class AddInEvents
    {
        #region Win32 Window Activation APIs (Multi-Monitor & Taskbar Support)

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsIconic(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, IntPtr ProcessId);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);

        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();

        [DllImport("user32.dll")]
        private static extern void SwitchToThisWindow(IntPtr hWnd, bool fAltTab);

        private const int SW_RESTORE = 9;
        private const int SW_SHOW = 5;

        private static void BringWindowToFront(IntPtr hWnd)
        {
            if (hWnd == IntPtr.Zero) return;

            try
            {
                if (IsIconic(hWnd))
                {
                    ShowWindow(hWnd, SW_RESTORE);
                }
                else
                {
                    ShowWindow(hWnd, SW_SHOW);
                }

                IntPtr foregroundHwnd = GetForegroundWindow();
                uint foregroundThread = GetWindowThreadProcessId(foregroundHwnd, IntPtr.Zero);
                uint currentThread = GetCurrentThreadId();

                if (foregroundThread != currentThread && foregroundThread != 0)
                {
                    AttachThreadInput(currentThread, foregroundThread, true);
                    SetForegroundWindow(hWnd);
                    AttachThreadInput(currentThread, foregroundThread, false);
                }
                else
                {
                    SetForegroundWindow(hWnd);
                }

                SwitchToThisWindow(hWnd, true);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"BringWindowToFront error: {ex.Message}");
            }
        }

        private void ActivateWorkbookAndWindow(Workbook? targetWb)
        {
            if (targetWb == null) return;

            try
            {
                // Kiểm tra nếu Workbook này đã là ActiveWorkbook thì không gọi targetWb.Activate() lại
                Workbook? currentActiveWb = null;
                bool isAlreadyActive = false;
                try
                {
                    currentActiveWb = _excelApp?.ActiveWorkbook;
                    if (currentActiveWb != null && string.Equals(currentActiveWb.Name, targetWb.Name, StringComparison.OrdinalIgnoreCase))
                    {
                        isAlreadyActive = true;
                    }
                }
                catch { }
                finally
                {
                    if (currentActiveWb != null) Marshal.ReleaseComObject(currentActiveWb);
                }

                if (!isAlreadyActive)
                {
                    targetWb.Activate();
                }

                // Khôi phục cửa sổ nếu bị minimize
                Windows? windows = null;
                try
                {
                    windows = targetWb.Windows;
                    if (windows != null && windows.Count > 0)
                    {
                        Window? win = null;
                        try
                        {
                            win = windows[1];
                            if (win != null)
                            {
                                if (win.WindowState == XlWindowState.xlMinimized)
                                {
                                    win.WindowState = XlWindowState.xlNormal;
                                }
                                if (!isAlreadyActive)
                                {
                                    win.Activate();
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
                catch { }
                finally
                {
                    if (windows != null) Marshal.ReleaseComObject(windows);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ActivateWorkbookAndWindow error: {ex.Message}");
            }
        }

        private void OnRequestActivateWorkbook(string workbookName)
        {
            if (_excelApp == null || string.IsNullOrEmpty(workbookName)) return;

            ExcelDna.Integration.ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    dynamic app = _excelApp;
                    dynamic? targetWb = null;
                    try { targetWb = app.Workbooks[workbookName]; } catch { }
                    if (targetWb != null)
                    {
                        ActivateWorkbookAndWindow(targetWb);
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Activate Workbook error: {ex.Message}");
                }
            });
        }

        private void OnRequestActivateWorksheet(string workbookName, string sheetName)
        {
            if (_excelApp == null || string.IsNullOrEmpty(sheetName)) return;

            ExcelDna.Integration.ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    dynamic app = _excelApp;
                    dynamic? targetWb = null;
                    if (!string.IsNullOrEmpty(workbookName))
                    {
                        try { targetWb = app.Workbooks[workbookName]; } catch { }
                    }
                    if (targetWb == null)
                    {
                        try { targetWb = app.ActiveWorkbook; } catch { }
                    }

                    if (targetWb != null)
                    {
                        ActivateWorkbookAndWindow(targetWb);

                        dynamic? ws = null;
                        try { ws = targetWb.Worksheets[sheetName]; } catch { }
                        if (ws == null)
                        {
                            try { ws = targetWb.Sheets[sheetName]; } catch { }
                        }

                        if (ws != null)
                        {
                            if ((int)ws.Visible != (int)XlSheetVisibility.xlSheetVisible)
                            {
                                ws.Visible = (int)XlSheetVisibility.xlSheetVisible;
                            }

                            // Kiểm tra nếu sheet này đã đang là ActiveSheet thì không activate lại
                            object? currentActiveSheet = null;
                            string? currentSheetName = null;
                            try
                            {
                                currentActiveSheet = _excelApp?.ActiveSheet;
                                if (currentActiveSheet is _Worksheet curWs)
                                {
                                    currentSheetName = curWs.Name;
                                }
                            }
                            catch { }
                            finally
                            {
                                if (currentActiveSheet != null) Marshal.ReleaseComObject(currentActiveSheet);
                            }

                            if (!string.Equals(currentSheetName, sheetName, StringComparison.OrdinalIgnoreCase))
                            {
                                try { ws.Activate(); } catch { }
                            }

                            // Đồng bộ ngay huy hiệu Active trên giao diện
                            MainViewModel?.SetActiveSelection(targetWb.Name, sheetName);
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Activate Sheet error: {ex.Message}");
                }
            });
        }

        private void OnRequestCloseWorkbook(string workbookName)
        {
            if (_excelApp == null) return;

            try
            {
                dynamic app = _excelApp;
                dynamic? targetWb = null;
                if (!string.IsNullOrEmpty(workbookName))
                {
                    try { targetWb = app.Workbooks[workbookName]; } catch { }
                }

                if (targetWb != null)
                {
                    targetWb.Close();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Close Workbook error: {ex.Message}");
            }
        }

        private void OnRequestSetSheetTabColor(string workbookName, string sheetName, Color? color)
        {
            if (_excelApp == null) return;

            try
            {
                dynamic app = _excelApp;
                dynamic? targetWb = null;

                if (!string.IsNullOrEmpty(workbookName))
                {
                    try { targetWb = app.Workbooks[workbookName]; } catch { }
                }
                if (targetWb == null)
                {
                    try { targetWb = app.ActiveWorkbook; } catch { }
                }

                if (targetWb != null)
                {
                    dynamic? ws = null;
                    try { ws = targetWb.Worksheets[sheetName]; } catch { }
                    if (ws == null)
                    {
                        try { ws = targetWb.Sheets[sheetName]; } catch { }
                    }

                    if (ws != null)
                    {
                        if (color == null)
                        {
                            try { ws.Tab.ColorIndex = -4142; } catch { }
                        }
                        else
                        {
                            int oleColor = ColorTranslator.ToOle(color.Value);
                            ws.Tab.Color = oleColor;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Set Sheet Tab Color error: {ex.Message}");
            }

            // Cập nhật lại toàn bộ UI ngay lập tức
            RefreshWorkbookTree();
        }

        private void OnRequestSetSheetVisibility(string workbookName, string sheetName, int visibility)
        {
            if (_excelApp == null) return;

            try
            {
                dynamic app = _excelApp;
                dynamic? targetWb = null;

                if (!string.IsNullOrEmpty(workbookName))
                {
                    try { targetWb = app.Workbooks[workbookName]; } catch { }
                }
                if (targetWb == null)
                {
                    try { targetWb = app.ActiveWorkbook; } catch { }
                }

                if (targetWb != null)
                {
                    dynamic? ws = null;
                    try { ws = targetWb.Worksheets[sheetName]; } catch { }
                    if (ws == null)
                    {
                        try { ws = targetWb.Sheets[sheetName]; } catch { }
                    }

                    if (ws != null)
                    {
                        ws.Visible = visibility;
                        if (visibility == (int)XlSheetVisibility.xlSheetVisible)
                        {
                            try { ws.Activate(); } catch { }
                        }
                    }
                }

                RefreshWorkbookTree();
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Không thể thay đổi trạng thái ẩn/hiện của sheet:\n{ex.Message}",
                                   "Sheet Navigator", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
            }
        }

        private void OnRequestUnhideAllSheets(string workbookName)
        {
            if (_excelApp == null) return;

            try
            {
                dynamic app = _excelApp;
                dynamic? targetWb = null;

                if (!string.IsNullOrEmpty(workbookName))
                {
                    try { targetWb = app.Workbooks[workbookName]; } catch { }
                }
                if (targetWb == null)
                {
                    try { targetWb = app.ActiveWorkbook; } catch { }
                }

                if (targetWb != null)
                {
                    int sheetCount = targetWb.Sheets.Count;
                    for (int j = 1; j <= sheetCount; j++)
                    {
                        try
                        {
                            dynamic ws = targetWb.Sheets[j];
                            if ((int)ws.Visible != (int)XlSheetVisibility.xlSheetVisible)
                            {
                                ws.Visible = (int)XlSheetVisibility.xlSheetVisible;
                            }
                        }
                        catch { }
                    }
                }

                RefreshWorkbookTree();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Unhide All Sheets error: {ex.Message}");
            }
        }

        #endregion
    }
}
