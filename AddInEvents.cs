using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using ExcelSupport.Host;
using ExcelSupport.Models;
using ExcelSupport.Services;
using ExcelSupport.ViewModels;
using CompareOptions = ExcelSupport.Models.CompareOptions;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using WpfApplication = System.Windows.Application;
using WpfMessageBox = System.Windows.MessageBox;
using SysAction = System.Action;

namespace ExcelSupport
{
    public partial class AddInEvents : IExcelAddIn
    {
        public static AddInEvents? Instance { get; private set; }
        public static TaskPaneViewModel? MainViewModel { get; private set; }

        private ExcelApp? _excelApp;
        public ExcelApp? ExcelAppInstance => _excelApp;
        private bool _isBatchProcessing = false;

        public void AutoOpen()
        {
            Instance = this;

            // Khởi tạo WPF Application runtime với ShutdownMode = OnExplicitShutdown
            // Đảm bảo không bao giờ đóng tiến trình Excel khi một hộp thoại WPF đóng lại
            if (WpfApplication.Current == null)
            {
                try
                {
                    new WpfApplication
                    {
                        ShutdownMode = System.Windows.ShutdownMode.OnExplicitShutdown
                    };
                }
                catch { }
            }

            _excelApp = (ExcelApp)ExcelDnaUtil.Application;
            MainViewModel = new TaskPaneViewModel();

            MainViewModel.RequestActivateWorkbook += OnRequestActivateWorkbook;
            MainViewModel.RequestActivateWorksheet += OnRequestActivateWorksheet;
            MainViewModel.RequestCloseWorkbook += OnRequestCloseWorkbook;
            MainViewModel.RequestSetSheetTabColor += OnRequestSetSheetTabColor;
            MainViewModel.RequestSetSheetVisibility += OnRequestSetSheetVisibility;
            MainViewModel.RequestUnhideAllSheets += OnRequestUnhideAllSheets;

            HookExcelEvents();

            try
            {
                _excelApp?.OnKey("^+Q", "OracleQuickQueryCommand");
                _excelApp?.OnKey("^+q", "OracleQuickQueryCommand");
                _excelApp?.OnKey("^+H", "ApplyDesignHighlightSelectionCommand");
                _excelApp?.OnKey("^+h", "ApplyDesignHighlightSelectionCommand");
                _excelApp?.OnKey("^+!H", "ClearDesignHighlightSelectionCommand");
                _excelApp?.OnKey("^+!h", "ClearDesignHighlightSelectionCommand");
                _excelApp?.OnKey("^+M", "ExportMarkdownTableCommand");
                _excelApp?.OnKey("^+m", "ExportMarkdownTableCommand");
                _excelApp?.OnKey("^+W", "ToggleTaskPaneCommand");
                _excelApp?.OnKey("^+w", "ToggleTaskPaneCommand");
            }
            catch { }

            QueueRefresh();
        }

        public void AutoClose()
        {
            UnhookExcelEvents();

            try
            {
                _excelApp?.OnKey("^+Q");
                _excelApp?.OnKey("^+q");
                _excelApp?.OnKey("^+H");
                _excelApp?.OnKey("^+h");
                _excelApp?.OnKey("^+!H");
                _excelApp?.OnKey("^+!h");
                _excelApp?.OnKey("^+M");
                _excelApp?.OnKey("^+m");
                _excelApp?.OnKey("^+W");
                _excelApp?.OnKey("^+w");
            }
            catch { }

            if (MainViewModel != null)
            {
                MainViewModel.RequestActivateWorkbook -= OnRequestActivateWorkbook;
                MainViewModel.RequestActivateWorksheet -= OnRequestActivateWorksheet;
                MainViewModel.RequestCloseWorkbook -= OnRequestCloseWorkbook;
                MainViewModel.RequestSetSheetTabColor -= OnRequestSetSheetTabColor;
                MainViewModel.RequestSetSheetVisibility -= OnRequestSetSheetVisibility;
                MainViewModel.RequestUnhideAllSheets -= OnRequestUnhideAllSheets;
                MainViewModel = null;
            }

            lock (_refreshLock)
            {
                _refreshDebounceTimer?.Dispose();
                _refreshDebounceTimer = null;
            }

            lock (_selectionLock)
            {
                _selectionDebounceTimer?.Dispose();
                _selectionDebounceTimer = null;
            }

            TaskPaneRegistry.DetachTaskPane();

            if (_excelApp != null)
            {
                try
                {
                    Marshal.ReleaseComObject(_excelApp);
                }
                catch { }
                finally
                {
                    _excelApp = null;
                }
            }

            Instance = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        public void RefreshWorkbookTreePublic()
        {
            QueueRefresh();
        }

        #region Event Hooking & Realtime Synchronization

        private void HookExcelEvents()
        {
            if (_excelApp == null) return;
            try
            {
                var events = (AppEvents_Event)_excelApp;

                events.NewWorkbook += ExcelApp_NewWorkbook;
                _excelApp.WorkbookOpen += ExcelApp_WorkbookOpen;
                _excelApp.WorkbookBeforeClose += ExcelApp_WorkbookBeforeClose;
                _excelApp.WorkbookAfterSave += ExcelApp_WorkbookAfterSave;
                _excelApp.WorkbookActivate += ExcelApp_WorkbookActivate;
                _excelApp.WorkbookDeactivate += ExcelApp_WorkbookDeactivate;

                _excelApp.WorkbookNewSheet += ExcelApp_WorkbookNewSheet;
                _excelApp.SheetActivate += ExcelApp_SheetActivate;
                _excelApp.SheetDeactivate += ExcelApp_SheetDeactivate;
                _excelApp.SheetChange += ExcelApp_SheetChange;
                _excelApp.SheetSelectionChange += ExcelApp_SheetSelectionChange;
                _excelApp.WorkbookBeforeSave += ExcelApp_WorkbookBeforeSave;

                _excelApp.WindowActivate += ExcelApp_WindowActivate;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Lỗi hook sự kiện Excel: {ex.Message}");
            }
        }

        private void UnhookExcelEvents()
        {
            if (_excelApp == null) return;
            try
            {
                var events = (AppEvents_Event)_excelApp;

                events.NewWorkbook -= ExcelApp_NewWorkbook;
                _excelApp.WorkbookOpen -= ExcelApp_WorkbookOpen;
                _excelApp.WorkbookBeforeClose -= ExcelApp_WorkbookBeforeClose;
                _excelApp.WorkbookAfterSave -= ExcelApp_WorkbookAfterSave;
                _excelApp.WorkbookActivate -= ExcelApp_WorkbookActivate;
                _excelApp.WorkbookDeactivate -= ExcelApp_WorkbookDeactivate;

                _excelApp.WorkbookNewSheet -= ExcelApp_WorkbookNewSheet;
                _excelApp.SheetActivate -= ExcelApp_SheetActivate;
                _excelApp.SheetDeactivate -= ExcelApp_SheetDeactivate;
                _excelApp.SheetChange -= ExcelApp_SheetChange;
                _excelApp.SheetSelectionChange -= ExcelApp_SheetSelectionChange;
                _excelApp.WorkbookBeforeSave -= ExcelApp_WorkbookBeforeSave;

                _excelApp.WindowActivate -= ExcelApp_WindowActivate;
            }
            catch { }
        }

        private void ExcelApp_NewWorkbook(Workbook Wb) => QueueRefresh();
        private void ExcelApp_WorkbookOpen(Workbook Wb) => QueueRefresh();
        private void ExcelApp_WorkbookBeforeClose(Workbook Wb, ref bool Cancel)
        {
            Services.GridRulerService.OnWorkbookBeforeSave(Wb);
            QueueRefresh();
        }
        private void ExcelApp_WorkbookBeforeSave(Workbook Wb, bool SaveAsUI, ref bool Cancel)
        {
            Services.GridRulerService.OnWorkbookBeforeSave(Wb);
        }
        private void ExcelApp_WorkbookAfterSave(Workbook Wb, bool Success) => QueueRefresh();
        private void ExcelApp_WorkbookNewSheet(Workbook Wb, object Sh) => QueueRefresh();
        private void ExcelApp_SheetActivate(object Sh)
        {
            QueueActiveSelectionSync();
            Services.GridRulerService.OnSheetActivate(Sh as _Worksheet);
        }
        private void ExcelApp_SheetDeactivate(object Sh) { }
        private void ExcelApp_SheetChange(object Sh, Range Target) { }
        private void ExcelApp_SheetSelectionChange(object Sh, Range Target)
        {
            Services.GridRulerService.OnSheetSelectionChange(Sh as _Worksheet, Target);
        }
        private void ExcelApp_WorkbookActivate(Workbook Wb) => QueueActiveSelectionSync();
        private void ExcelApp_WorkbookDeactivate(Workbook Wb) { }
        private void ExcelApp_WindowActivate(Workbook Wb, Window Wn) => QueueActiveSelectionSync();

        private System.Threading.Timer? _refreshDebounceTimer;
        private readonly object _refreshLock = new object();

        public void QueueRefresh()
        {
            if (_isBatchProcessing) return;

            lock (_refreshLock)
            {
                _refreshDebounceTimer?.Dispose();
                _refreshDebounceTimer = new System.Threading.Timer(_ =>
                {
                    ExcelAsyncUtil.QueueAsMacro(() =>
                    {
                        if (_isBatchProcessing) return;
                        RefreshWorkbookTree();
                    });
                }, null, 150, System.Threading.Timeout.Infinite);
            }
        }

        private System.Threading.Timer? _selectionDebounceTimer;
        private readonly object _selectionLock = new object();

        private void QueueActiveSelectionSync()
        {
            if (_isBatchProcessing) return;

            lock (_selectionLock)
            {
                _selectionDebounceTimer?.Dispose();
                _selectionDebounceTimer = new System.Threading.Timer(_ =>
                {
                    ExcelAsyncUtil.QueueAsMacro(() =>
                    {
                        if (_isBatchProcessing) return;
                        UpdateActiveSheetState();
                    });
                }, null, 60, System.Threading.Timeout.Infinite);
            }
        }

        private void RefreshWorkbookTree()
        {
            if (_excelApp == null || MainViewModel == null) return;

            var treeData = new List<WorkbookNodeViewModel>();
            Workbooks? workbooks = null;
            string? activeWbName = null;
            string? activeWsName = null;

            try
            {
                workbooks = _excelApp.Workbooks;
                int wbCount = workbooks.Count;

                Workbook? activeWb = null;
                object? activeSheetObj = null;
                try
                {
                    activeWb = _excelApp.ActiveWorkbook;
                    if (activeWb != null) activeWbName = activeWb.Name;

                    activeSheetObj = _excelApp.ActiveSheet;
                    if (activeSheetObj is _Worksheet ws)
                    {
                        activeWsName = ws.Name;
                    }
                }
                catch { }
                finally
                {
                    if (activeSheetObj != null) Marshal.ReleaseComObject(activeSheetObj);
                    if (activeWb != null) Marshal.ReleaseComObject(activeWb);
                }

                for (int i = 1; i <= wbCount; i++)
                {
                    Workbook? wb = null;
                    Sheets? sheets = null;
                    try
                    {
                        wb = workbooks[i];
                        var wbNode = new WorkbookNodeViewModel
                        {
                            WorkbookName = wb.Name,
                            FilePath = wb.FullName,
                            IsActive = (wb.Name == activeWbName)
                        };

                        sheets = wb.Sheets;
                        int sheetCount = sheets.Count;

                        for (int j = 1; j <= sheetCount; j++)
                        {
                            object? sheetObj = null;
                            Tab? tab = null;
                            try
                            {
                                sheetObj = sheets[j];
                                if (sheetObj is _Worksheet ws)
                                {
                                    var wsNode = new WorksheetNodeViewModel
                                    {
                                        WorkbookName = wb.Name,
                                        SheetName = ws.Name,
                                        Index = j,
                                        IsActive = (wb.Name == activeWbName && ws.Name == activeWsName)
                                    };

                                    // Đọc trạng thái Ẩn/Hiện của Sheet
                                    try
                                    {
                                        int vis = Convert.ToInt32(ws.Visible);
                                        wsNode.IsHidden = (vis != (int)XlSheetVisibility.xlSheetVisible);
                                        wsNode.IsVeryHidden = (vis == (int)XlSheetVisibility.xlSheetVeryHidden);
                                    }
                                    catch { }

                                    // Đọc trạng thái Protect của Sheet
                                    try
                                    {
                                        wsNode.IsProtected = ws.ProtectContents;
                                    }
                                    catch { }

                                    // Đọc màu Tab của Sheet
                                    try
                                    {
                                        tab = ws.Tab;
                                        object rawColor = tab.Color;
                                        int colorIndex = -4142;
                                        try { colorIndex = (int)tab.ColorIndex; } catch { }

                                        if (rawColor is bool b && !b)
                                        {
                                            wsNode.TabColorHex = null;
                                        }
                                        else if (colorIndex == -4142 && (rawColor == null || Convert.ToInt64(rawColor) == 0 || Convert.ToInt64(rawColor) == 16777215))
                                        {
                                            wsNode.TabColorHex = null;
                                        }
                                        else if (rawColor != null)
                                        {
                                            int oleColor = Convert.ToInt32(rawColor);
                                            var sysColor = ColorTranslator.FromOle(oleColor);
                                            wsNode.TabColorHex = $"#{sysColor.R:X2}{sysColor.G:X2}{sysColor.B:X2}";
                                        }
                                        else
                                        {
                                            wsNode.TabColorHex = null;
                                        }
                                    }
                                    catch
                                    {
                                        wsNode.TabColorHex = null;
                                    }

                                    wbNode.Worksheets.Add(wsNode);
                                }
                            }
                            finally
                            {
                                if (tab != null) Marshal.ReleaseComObject(tab);
                                if (sheetObj != null) Marshal.ReleaseComObject(sheetObj);
                            }
                        }

                        wbNode.NotifyWorksheetsUpdated();
                        treeData.Add(wbNode);
                    }
                    finally
                    {
                        if (sheets != null) Marshal.ReleaseComObject(sheets);
                        if (wb != null) Marshal.ReleaseComObject(wb);
                    }
                }
            }
            catch (COMException)
            {
                // Bỏ qua khi Excel đang trong Cell Edit Mode
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"RefreshWorkbookTree error: {ex.Message}");
            }
            finally
            {
                if (workbooks != null) Marshal.ReleaseComObject(workbooks);
            }

            WpfApplication.Current?.Dispatcher.BeginInvoke(new SysAction(() =>
            {
                MainViewModel.MergeWorkbooks(treeData, activeWbName, activeWsName);
            }));
        }

        public void ApplySheetColor(string workbookName, string sheetName, Color? color)
        {
            OnRequestSetSheetTabColor(workbookName, sheetName, color);
        }

        private void UpdateActiveSheetState()
        {
            if (_excelApp == null || MainViewModel == null) return;

            Workbook? activeWb = null;
            object? activeSheetObj = null;
            string? activeWbName = null;
            string? activeSheetName = null;

            try
            {
                activeWb = _excelApp.ActiveWorkbook;
                if (activeWb != null) activeWbName = activeWb.Name;

                activeSheetObj = _excelApp.ActiveSheet;
                if (activeSheetObj is _Worksheet ws)
                {
                    activeSheetName = ws.Name;
                }
            }
            catch (COMException) { }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"UpdateActiveSheetState error: {ex.Message}");
            }
            finally
            {
                if (activeSheetObj != null) Marshal.ReleaseComObject(activeSheetObj);
                if (activeWb != null) Marshal.ReleaseComObject(activeWb);
            }

            WpfApplication.Current?.Dispatcher.BeginInvoke(new SysAction(() =>
            {
                MainViewModel.SetActiveSelection(activeWbName, activeSheetName);
            }));
        }

        #endregion
    }
}

