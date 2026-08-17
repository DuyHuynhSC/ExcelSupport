using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using ExcelSupport.Host;
using ExcelSupport.Models;
using ExcelSupport.ViewModels;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using WpfApplication = System.Windows.Application;
using WpfMessageBox = System.Windows.MessageBox;
using SysAction = System.Action;

namespace ExcelSupport
{
    public class AddInEvents : IExcelAddIn
    {
        public static AddInEvents? Instance { get; private set; }
        public static TaskPaneViewModel? MainViewModel { get; private set; }

        private ExcelApp? _excelApp;
        public ExcelApp? ExcelAppInstance => _excelApp;
        private bool _isBatchProcessing = false;

        public void AutoOpen()
        {
            Instance = this;
            _excelApp = (ExcelApp)ExcelDnaUtil.Application;
            MainViewModel = new TaskPaneViewModel();

            MainViewModel.RequestActivateWorkbook += OnRequestActivateWorkbook;
            MainViewModel.RequestActivateWorksheet += OnRequestActivateWorksheet;
            MainViewModel.RequestCloseWorkbook += OnRequestCloseWorkbook;
            MainViewModel.RequestSetSheetTabColor += OnRequestSetSheetTabColor;
            MainViewModel.RequestSetSheetVisibility += OnRequestSetSheetVisibility;
            MainViewModel.RequestUnhideAllSheets += OnRequestUnhideAllSheets;

            HookExcelEvents();

            QueueRefresh();
        }

        public void AutoClose()
        {
            UnhookExcelEvents();

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

                _excelApp.WindowActivate -= ExcelApp_WindowActivate;
            }
            catch { }
        }

        private void ExcelApp_NewWorkbook(Workbook Wb) => QueueRefresh();
        private void ExcelApp_WorkbookOpen(Workbook Wb) => QueueRefresh();
        private void ExcelApp_WorkbookBeforeClose(Workbook Wb, ref bool Cancel) => QueueRefresh();
        private void ExcelApp_WorkbookAfterSave(Workbook Wb, bool Success) => QueueRefresh();
        private void ExcelApp_WorkbookNewSheet(Workbook Wb, object Sh) => QueueRefresh();
        private void ExcelApp_SheetActivate(object Sh) => QueueRefresh();
        private void ExcelApp_SheetDeactivate(object Sh) => QueueRefresh();
        private void ExcelApp_SheetChange(object Sh, Range Target) => QueueActiveSelectionSync();
        private void ExcelApp_WorkbookActivate(Workbook Wb) => QueueRefresh();
        private void ExcelApp_WorkbookDeactivate(Workbook Wb) => QueueRefresh();
        private void ExcelApp_WindowActivate(Workbook Wb, Window Wn)
        {
            QueueRefresh();
            ExcelSupport.Ribbon.RibbonController.Instance?.InvalidateRibbon();
        }

        public void QueueRefresh()
        {
            if (_isBatchProcessing) return;
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                if (_isBatchProcessing) return;
                RefreshWorkbookTree();
                if (MainViewModel != null)
                {
                    TaskPaneRegistry.AutoRestoreForActiveWindow(MainViewModel);
                }
                ExcelSupport.Ribbon.RibbonController.Instance?.InvalidateRibbon();
            });
        }

        private void QueueActiveSelectionSync()
        {
            if (_isBatchProcessing) return;
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                if (_isBatchProcessing) return;
                UpdateActiveSheetState();
            });
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

        private void OnRequestActivateWorkbook(string workbookName)
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
                    targetWb.Activate();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Activate Workbook error: {ex.Message}");
            }
        }

        private void OnRequestActivateWorksheet(string workbookName, string sheetName)
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
                    targetWb.Activate();
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
                        ws.Activate();
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Activate Sheet error: {ex.Message}");
            }
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

        #region Excel COM Helper Methods for AI Assistant & Translation

        public class ActiveCellInfo
        {
            public string WorkbookName { get; set; } = string.Empty;
            public string SheetName { get; set; } = string.Empty;
            public string CellAddress { get; set; } = string.Empty;
            public string Formula { get; set; } = string.Empty;
            public string Value { get; set; } = string.Empty;
            public string ErrorText { get; set; } = string.Empty;
            public bool HasError => !string.IsNullOrEmpty(ErrorText);
        }

        public class CellTextItem
        {
            public int Row { get; set; }
            public int Column { get; set; }
            public string Address { get; set; } = string.Empty;
            public string OriginalText { get; set; } = string.Empty;
            public string TranslatedText { get; set; } = string.Empty;
        }

        public ActiveCellInfo? GetActiveCellInfo()
        {
            if (_excelApp == null) return null;

            try
            {
                Range? cell = null;
                _Worksheet? ws = null;
                Workbook? wb = null;
                try
                {
                    cell = _excelApp.ActiveCell;
                    if (cell == null) return null;

                    ws = cell.Worksheet;
                    wb = ws.Parent as Workbook;

                    var info = new ActiveCellInfo
                    {
                        WorkbookName = wb?.Name ?? string.Empty,
                        SheetName = ws.Name,
                        CellAddress = cell.Address[false, false],
                        Formula = cell.Formula?.ToString() ?? string.Empty,
                        Value = cell.Text?.ToString() ?? string.Empty
                    };

                    // Kiểm tra nếu ô đang chứa mã lỗi Excel (#N/A, #VALUE!, #REF!, #DIV/0!, ...)
                    string valStr = info.Value.Trim();
                    if (valStr.StartsWith("#") && (valStr.EndsWith("!") || valStr.EndsWith("?")))
                    {
                        info.ErrorText = valStr;
                    }

                    return info;
                }
                finally
                {
                    if (cell != null) Marshal.ReleaseComObject(cell);
                    if (ws != null) Marshal.ReleaseComObject(ws);
                    if (wb != null) Marshal.ReleaseComObject(wb);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"GetActiveCellInfo error: {ex.Message}");
                return null;
            }
        }

        public bool InsertFormulaToActiveCell(string formula)
        {
            if (_excelApp == null || string.IsNullOrWhiteSpace(formula)) return false;

            try
            {
                string cleanFormula = formula.Trim();
                if (!cleanFormula.StartsWith("=") && !cleanFormula.StartsWith("@"))
                {
                    cleanFormula = "=" + cleanFormula;
                }

                Range? cell = null;
                try
                {
                    cell = _excelApp.ActiveCell;
                    if (cell != null)
                    {
                        cell.Formula = cleanFormula;
                        return true;
                    }
                }
                finally
                {
                    if (cell != null) Marshal.ReleaseComObject(cell);
                }
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Không thể chèn công thức vào ô Excel:\n{ex.Message}", "Trợ Lý AI",
                                   System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
            }
            return false;
        }

        public List<CellTextItem> GetSelectedCellsText(int maxCells = 500)
        {
            var list = new List<CellTextItem>();
            if (_excelApp == null) return list;

            try
            {
                Range? selection = null;
                try
                {
                    selection = _excelApp.Selection as Range;
                    if (selection == null) return list;

                    int count = 0;
                    foreach (Range cell in selection.Cells)
                    {
                        try
                        {
                            string text = cell.Text?.ToString() ?? string.Empty;
                            if (!string.IsNullOrWhiteSpace(text))
                            {
                                list.Add(new CellTextItem
                                {
                                    Row = cell.Row,
                                    Column = cell.Column,
                                    Address = cell.Address[false, false],
                                    OriginalText = text.Trim()
                                });

                                count++;
                                if (count >= maxCells) break;
                            }
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(cell);
                        }
                    }
                }
                finally
                {
                    if (selection != null) Marshal.ReleaseComObject(selection);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"GetSelectedCellsText error: {ex.Message}");
            }

            return list;
        }

        public bool WriteTranslatedCells(List<CellTextItem> items, bool writeToAdjacentColumn)
        {
            if (_excelApp == null || items == null || items.Count == 0) return false;

            try
            {
                _Worksheet? ws = null;
                try
                {
                    ws = _excelApp.ActiveSheet as _Worksheet;
                    if (ws == null) return false;

                    _excelApp.ScreenUpdating = false;

                    foreach (var item in items)
                    {
                        if (string.IsNullOrEmpty(item.TranslatedText)) continue;

                        int targetCol = writeToAdjacentColumn ? (item.Column + 1) : item.Column;
                        Range? cell = null;
                        try
                        {
                            cell = ws.Cells[item.Row, targetCol] as Range;
                            if (cell != null)
                            {
                                cell.Value2 = item.TranslatedText;
                            }
                        }
                        finally
                        {
                            if (cell != null) Marshal.ReleaseComObject(cell);
                        }
                    }

                    return true;
                }
                finally
                {
                    _excelApp.ScreenUpdating = true;
                    if (ws != null) Marshal.ReleaseComObject(ws);
                }
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Lỗi ghi dữ liệu dịch vào bảng tính:\n{ex.Message}", "Dịch Thuật AI",
                                   System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                return false;
            }
        }

        public bool CreateTableOfContents(string? workbookName, bool addBackLinkToSheets = true)
        {
            if (_excelApp == null) return false;

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

                if (targetWb == null)
                {
                    WpfMessageBox.Show("Không tìm thấy Workbook đang mở.", "Tạo Mục Lục",
                                       System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                    return false;
                }

                try { app.ScreenUpdating = false; } catch { }

                string indexSheetName = "Mục Lục";
                dynamic? wsIndex = null;

                // Tìm xem đã có sheet Mục Lục chưa
                try
                {
                    wsIndex = targetWb.Worksheets[indexSheetName];
                }
                catch { }

                if (wsIndex == null)
                {
                    try
                    {
                        dynamic firstSheet = targetWb.Sheets[1];
                        wsIndex = targetWb.Worksheets.Add(firstSheet);
                    }
                    catch
                    {
                        wsIndex = targetWb.Worksheets.Add();
                    }

                    if (wsIndex != null)
                    {
                        wsIndex.Name = indexSheetName;
                    }
                }
                else
                {
                    // Di chuyển lên đầu tiên
                    try
                    {
                        dynamic firstSheet = targetWb.Sheets[1];
                        wsIndex.Move(firstSheet);
                    }
                    catch { }

                    try
                    {
                        if ((int)wsIndex.Visible != (int)XlSheetVisibility.xlSheetVisible)
                        {
                            wsIndex.Visible = (int)XlSheetVisibility.xlSheetVisible;
                        }
                    }
                    catch { }

                    try
                    {
                        wsIndex.Cells.Clear();
                    }
                    catch { }
                }

                if (wsIndex == null) return false;

                // 1. Tiêu đề lớn
                dynamic titleRange = wsIndex.Range["A1:E1"];
                titleRange.Merge();
                titleRange.Value2 = $"📋 BẢNG MỤC LỤC CÁC SHEET - {targetWb.Name}";
                titleRange.Font.Size = 14;
                titleRange.Font.Bold = true;
                titleRange.Font.Name = "Segoe UI";
                titleRange.Font.Color = ColorTranslator.ToOle(Color.White);
                titleRange.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(16, 124, 65)); // Office Dark Green
                titleRange.HorizontalAlignment = -4108; // xlCenter
                titleRange.VerticalAlignment = -4108; // xlCenter
                titleRange.RowHeight = 36;

                // 2. Tiêu đề các cột
                wsIndex.Cells[3, 1] = "STT";
                wsIndex.Cells[3, 2] = "Tên Sheet (Click để mở)";
                wsIndex.Cells[3, 3] = "Trạng Thái";
                wsIndex.Cells[3, 4] = "Màu Tab";
                wsIndex.Cells[3, 5] = "Ghi Chú";

                dynamic headerRange = wsIndex.Range["A3:E3"];
                headerRange.Font.Bold = true;
                headerRange.Font.Size = 11;
                headerRange.Font.Name = "Segoe UI";
                headerRange.Font.Color = ColorTranslator.ToOle(Color.White);
                headerRange.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(30, 41, 59)); // Slate 800
                headerRange.HorizontalAlignment = -4108; // xlCenter
                headerRange.VerticalAlignment = -4108; // xlCenter
                headerRange.RowHeight = 26;

                int currentRow = 4;
                int stt = 1;
                int sheetCount = targetWb.Sheets.Count;

                for (int i = 1; i <= sheetCount; i++)
                {
                    try
                    {
                        dynamic ws = targetWb.Sheets[i];
                        string wsName = ws.Name;
                        if (string.Equals(wsName, indexSheetName, StringComparison.OrdinalIgnoreCase))
                        {
                            continue; // Bỏ qua chính sheet Mục Lục
                        }

                        // Cột 1: STT
                        dynamic sttCell = wsIndex.Cells[currentRow, 1];
                        sttCell.Value2 = stt;
                        sttCell.HorizontalAlignment = -4108; // xlCenter

                        // Cột 2: Tên Sheet + Hyperlink
                        dynamic linkCell = wsIndex.Cells[currentRow, 2];
                        wsIndex.Hyperlinks.Add(
                            linkCell,
                            "",
                            $"'{wsName}'!A1",
                            $"Chuyển đến sheet [{wsName}]",
                            wsName
                        );
                        linkCell.Font.Name = "Segoe UI";
                        linkCell.Font.Size = 11;

                        // Cột 3: Trạng thái Ẩn/Hiện
                        dynamic statusCell = wsIndex.Cells[currentRow, 3];
                        int vis = (int)ws.Visible;
                        if (vis == (int)XlSheetVisibility.xlSheetVisible)
                        {
                            statusCell.Value2 = "Hiển thị";
                            statusCell.Font.Color = ColorTranslator.ToOle(Color.FromArgb(22, 101, 52));
                        }
                        else if (vis == (int)XlSheetVisibility.xlSheetHidden)
                        {
                            statusCell.Value2 = "Bị ẩn (Hidden)";
                            statusCell.Font.Color = ColorTranslator.ToOle(Color.FromArgb(180, 83, 9));
                        }
                        else
                        {
                            statusCell.Value2 = "Ẩn sâu (Very Hidden)";
                            statusCell.Font.Color = ColorTranslator.ToOle(Color.FromArgb(220, 38, 38));
                        }
                        statusCell.HorizontalAlignment = -4108;

                        // Cột 4: Màu Tab
                        dynamic colorCell = wsIndex.Cells[currentRow, 4];
                        try
                        {
                            object rawColor = ws.Tab.Color;
                            int colorIndex = -4142;
                            try { colorIndex = (int)ws.Tab.ColorIndex; } catch { }

                            if (colorIndex != -4142 && rawColor != null && !(rawColor is bool))
                            {
                                colorCell.Interior.Color = Convert.ToInt32(rawColor);
                                colorCell.Value2 = "   ";
                            }
                            else
                            {
                                colorCell.Value2 = "(Mặc định)";
                                colorCell.Font.Color = ColorTranslator.ToOle(Color.FromArgb(148, 163, 184));
                                colorCell.HorizontalAlignment = -4108;
                            }
                        }
                        catch
                        {
                            colorCell.Value2 = "-";
                            colorCell.HorizontalAlignment = -4108;
                        }

                        // Cột 5: Ghi chú
                        dynamic noteCell = wsIndex.Cells[currentRow, 5];
                        noteCell.Value2 = "";

                        stt++;
                        currentRow++;
                    }
                    catch (Exception exSheet)
                    {
                        System.Diagnostics.Debug.WriteLine($"Sheet index error: {exSheet.Message}");
                    }
                }

                // Định dạng toàn bộ bảng dữ liệu
                if (currentRow > 4)
                {
                    dynamic dataTable = wsIndex.Range[$"A3:E{currentRow - 1}"];
                    dataTable.Borders.LineStyle = 1; // xlContinuous
                    dataTable.Borders.Color = ColorTranslator.ToOle(Color.FromArgb(203, 213, 225));
                    dataTable.Font.Name = "Segoe UI";
                }

                // Tự động căn chỉnh độ rộng cột
                try
                {
                    wsIndex.Range["A:E"].EntireColumn.AutoFit();
                    wsIndex.Range["B:B"].ColumnWidth = Math.Max(25.0, Convert.ToDouble(wsIndex.Range["B:B"].ColumnWidth) + 5.0);
                    wsIndex.Range["E:E"].ColumnWidth = 20.0;
                }
                catch { }

                // Kích hoạt Sheet Mục Lục
                try
                {
                    wsIndex.Activate();
                }
                catch { }

                // Đổi màu Tab cho Sheet Mục Lục thành xanh lá nổi bật
                try
                {
                    wsIndex.Tab.Color = ColorTranslator.ToOle(Color.FromArgb(16, 124, 65));
                }
                catch { }

                RefreshWorkbookTree();

                WpfMessageBox.Show($"✅ Đã tạo thành công Bảng Mục Lục cho {stt - 1} sheet trong [{targetWb.Name}]!",
                                   "Tạo Mục Lục Sheet", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
                return true;
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Lỗi tạo bảng mục lục:\n{ex.Message}", "Tạo Mục Lục",
                                   System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                return false;
            }
            finally
            {
                if (_excelApp != null)
                {
                    try { _excelApp.ScreenUpdating = true; } catch { }
                }
            }
        }

        #endregion

        #region Excel COM Helper Methods for Rename, Split, and Merge Sheets

        public bool RenameWorksheet(string wbName, string oldSheetName, string newSheetName)
        {
            if (_excelApp == null || string.IsNullOrWhiteSpace(newSheetName)) return false;

            // Kiểm tra quy chuẩn đặt tên sheet Excel: tối đa 31 ký tự, không chứa \ / ? * [ ] :
            string cleanName = newSheetName.Trim();
            if (cleanName.Length > 31)
            {
                WpfMessageBox.Show("Tên Sheet không được vượt quá 31 ký tự.", "Đổi Tên Sheet",
                                   System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                return false;
            }

            char[] invalidChars = { '\\', '/', '?', '*', '[', ']', ':' };
            if (cleanName.IndexOfAny(invalidChars) >= 0)
            {
                WpfMessageBox.Show("Tên Sheet không được chứa các ký tự đặc biệt: \\ / ? * [ ] :", "Đổi Tên Sheet",
                                   System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                return false;
            }

            try
            {
                dynamic app = _excelApp;
                dynamic? targetWb = null;
                if (!string.IsNullOrEmpty(wbName))
                {
                    try { targetWb = app.Workbooks[wbName]; } catch { }
                }
                if (targetWb == null)
                {
                    try { targetWb = app.ActiveWorkbook; } catch { }
                }

                if (targetWb == null) return false;

                dynamic ws = targetWb.Worksheets[oldSheetName];
                if (ws != null)
                {
                    ws.Name = cleanName;
                    RefreshWorkbookTree();
                    return true;
                }
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Không thể đổi tên sheet:\n{ex.Message}", "Đổi Tên Sheet",
                                   System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
            }
            return false;
        }



        public bool BatchRenameWorksheets(string wbName, string prefix, string suffix, string findText, string replaceText)
        {
            if (_excelApp == null) return false;

            _isBatchProcessing = true;
            dynamic app = _excelApp;

            try
            {
                try { app.EnableEvents = false; } catch { }
                try { app.ScreenUpdating = false; } catch { }

                dynamic? targetWb = null;
                if (!string.IsNullOrEmpty(wbName))
                {
                    try { targetWb = app.Workbooks[wbName]; } catch { }
                }
                if (targetWb == null)
                {
                    try { targetWb = app.ActiveWorkbook; } catch { }
                }

                if (targetWb == null) return false;

                int count = targetWb.Sheets.Count;
                int renamedCount = 0;

                for (int i = 1; i <= count; i++)
                {
                    try
                    {
                        dynamic ws = targetWb.Sheets[i];
                        string currentName = ws.Name;
                        string newName = currentName;

                        if (!string.IsNullOrEmpty(findText))
                        {
                            newName = newName.Replace(findText, replaceText ?? string.Empty);
                        }

                        if (!string.IsNullOrEmpty(prefix))
                        {
                            newName = prefix + newName;
                        }

                        if (!string.IsNullOrEmpty(suffix))
                        {
                            newName = newName + suffix;
                        }

                        if (newName.Length > 31)
                        {
                            newName = newName.Substring(0, 31);
                        }

                        if (newName != currentName)
                        {
                            ws.Name = newName;
                            renamedCount++;
                        }
                    }
                    catch { }
                }

                _isBatchProcessing = false;
                try { app.EnableEvents = true; } catch { }
                try { app.ScreenUpdating = true; } catch { }

                RefreshWorkbookTree();
                WpfMessageBox.Show($"✅ Đã đổi tên thành công cho {renamedCount} sheet!", "Đổi Tên Hàng Loạt",
                                   System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
                return true;
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Lỗi đổi tên hàng loạt:\n{ex.Message}", "Đổi Tên Hàng Loạt",
                                   System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                return false;
            }
            finally
            {
                _isBatchProcessing = false;
                if (_excelApp != null)
                {
                    try { _excelApp.EnableEvents = true; } catch { }
                    try { _excelApp.ScreenUpdating = true; } catch { }
                }
            }
        }

        public bool SplitWorksheetsToFiles(string wbName, List<string>? sheetNames, string outputFolder, bool keepOriginalSheets = true)
        {
            if (_excelApp == null || string.IsNullOrWhiteSpace(outputFolder)) return false;

            if (!System.IO.Directory.Exists(outputFolder))
            {
                try { System.IO.Directory.CreateDirectory(outputFolder); }
                catch (Exception ex)
                {
                    WpfMessageBox.Show($"Không thể tạo thư mục lưu:\n{ex.Message}", "Tách Sheet",
                                       System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                    return false;
                }
            }

            _isBatchProcessing = true;
            dynamic app = _excelApp;

            try
            {
                try { app.EnableEvents = false; } catch { }
                try { app.DisplayAlerts = false; } catch { }
                try { app.ScreenUpdating = false; } catch { }

                dynamic? targetWb = null;
                if (!string.IsNullOrEmpty(wbName))
                {
                    try { targetWb = app.Workbooks[wbName]; } catch { }
                }
                if (targetWb == null)
                {
                    try { targetWb = app.ActiveWorkbook; } catch { }
                }

                if (targetWb == null) return false;

                int totalSheets = targetWb.Sheets.Count;
                int exportedCount = 0;
                var exportedSheetNames = new List<string>();

                for (int i = 1; i <= totalSheets; i++)
                {
                    try
                    {
                        dynamic ws = targetWb.Sheets[i];
                        string sheetName = ws.Name;

                        if (sheetNames != null && sheetNames.Count > 0 && !sheetNames.Contains(sheetName))
                        {
                            continue;
                        }

                        // Sao chép sheet sang một Workbook mới hoàn toàn
                        ws.Copy();
                        dynamic newWb = app.ActiveWorkbook;

                        if (newWb != null)
                        {
                            // Chuẩn hóa tên file xuất
                            string cleanSheetName = string.Join("_", sheetName.Split(System.IO.Path.GetInvalidFileNameChars()));
                            string filePath = System.IO.Path.Combine(outputFolder, $"{cleanSheetName}.xlsx");

                            // Lưu file và đóng lại
                            newWb.SaveAs(filePath, 51); // 51 = xlOpenXMLWorkbook (.xlsx)
                            newWb.Close(false);
                            exportedCount++;
                            exportedSheetNames.Add(sheetName);
                        }
                    }
                    catch (Exception exSheet)
                    {
                        System.Diagnostics.Debug.WriteLine($"Split sheet error: {exSheet.Message}");
                    }
                }

                // Nếu người dùng chọn xóa sheet gốc sau khi tách
                if (!keepOriginalSheets && exportedSheetNames.Count > 0)
                {
                    try
                    {
                        // Excel bắt buộc phải có ít nhất 1 sheet hiển thị
                        if (targetWb.Sheets.Count <= exportedSheetNames.Count)
                        {
                            dynamic emptyWs = targetWb.Worksheets.Add();
                            emptyWs.Name = "Sheet1";
                        }

                        foreach (var sName in exportedSheetNames)
                        {
                            try
                            {
                                dynamic wsDel = targetWb.Sheets[sName];
                                if (wsDel != null)
                                {
                                    wsDel.Delete();
                                }
                            }
                            catch { }
                        }
                    }
                    catch (Exception exDel)
                    {
                        System.Diagnostics.Debug.WriteLine($"Delete sheet error: {exDel.Message}");
                    }
                }

                // Kích hoạt lại Workbook ban đầu
                try { targetWb.Activate(); } catch { }

                _isBatchProcessing = false;
                try { app.EnableEvents = true; } catch { }
                try { app.DisplayAlerts = true; } catch { }
                try { app.ScreenUpdating = true; } catch { }

                RefreshWorkbookTree();

                string msg = keepOriginalSheets
                    ? $"✅ Đã tách và lưu thành công {exportedCount} file Excel (.xlsx) vào thư mục:\n{outputFolder}"
                    : $"✅ Đã tách thành công {exportedCount} file Excel (.xlsx) và xóa {exportedSheetNames.Count} sheet tương ứng khỏi file hiện tại!\nThư mục lưu: {outputFolder}";

                WpfMessageBox.Show(msg, "Tách Sheet Thành Công", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
                return true;
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Lỗi khi tách sheet:\n{ex.Message}", "Tách Sheet",
                                   System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                return false;
            }
            finally
            {
                _isBatchProcessing = false;
                if (_excelApp != null)
                {
                    try { _excelApp.EnableEvents = true; } catch { }
                    try { _excelApp.DisplayAlerts = true; } catch { }
                    try { _excelApp.ScreenUpdating = true; } catch { }
                }
            }
        }

        public bool ConsolidateSheetsData(string wbName, List<string>? sheetNames, bool hasHeaderRow = true)
        {
            if (_excelApp == null) return false;

            _isBatchProcessing = true;
            dynamic app = _excelApp;

            try
            {
                try { app.EnableEvents = false; } catch { }
                try { app.DisplayAlerts = false; } catch { }
                try { app.ScreenUpdating = false; } catch { }

                dynamic? targetWb = null;
                if (!string.IsNullOrEmpty(wbName))
                {
                    try { targetWb = app.Workbooks[wbName]; } catch { }
                }
                if (targetWb == null)
                {
                    try { targetWb = app.ActiveWorkbook; } catch { }
                }

                if (targetWb == null) return false;

                string resultSheetName = "Tong_Hop";
                dynamic? wsSummary = null;

                try { wsSummary = targetWb.Worksheets[resultSheetName]; } catch { }

                if (wsSummary == null)
                {
                    // Thêm Sheet Tong_Hop vào vị trí cuối cùng
                    dynamic lastSheet = targetWb.Sheets[targetWb.Sheets.Count];
                    wsSummary = targetWb.Worksheets.Add(After: lastSheet);
                    wsSummary.Name = resultSheetName;
                }
                else
                {
                    wsSummary.Cells.Clear();
                }

                int totalSheets = targetWb.Sheets.Count;
                int currentDestRow = 1;
                bool isFirstSheet = true;
                int sheetsMerged = 0;

                for (int i = 1; i <= totalSheets; i++)
                {
                    try
                    {
                        dynamic ws = targetWb.Sheets[i];
                        string sheetName = ws.Name;

                        if (sheetName == resultSheetName || (sheetNames != null && sheetNames.Count > 0 && !sheetNames.Contains(sheetName)))
                        {
                            continue;
                        }

                        dynamic usedRange = ws.UsedRange;
                        int firstRow = usedRange.Row;
                        int rowsCount = usedRange.Rows.Count;
                        int firstCol = usedRange.Column;
                        int colsCount = usedRange.Columns.Count;

                        if (rowsCount <= 0 || colsCount <= 0) continue;

                        int lastRow = firstRow + rowsCount - 1;
                        int lastCol = firstCol + colsCount - 1;

                        int startRow = firstRow;
                        if (!isFirstSheet && hasHeaderRow)
                        {
                            if (rowsCount > 1)
                            {
                                startRow = firstRow + 1; // Bỏ qua dòng tiêu đề nếu sheet có nhiều hơn 1 dòng
                            }
                            else
                            {
                                startRow = firstRow; // Nếu sheet chỉ có 1 dòng thì vẫn lấy dòng đó
                            }
                        }

                        if (startRow <= lastRow)
                        {
                            string startColLetter = GetExcelColumnLetter(firstCol);
                            string endColLetter = GetExcelColumnLetter(lastCol);
                            string rangeAddress = $"{startColLetter}{startRow}:{endColLetter}{lastRow}";

                            dynamic sourceRange = ws.Range[rangeAddress];
                            dynamic destCell = wsSummary.Range[$"A{currentDestRow}"];

                            sourceRange.Copy(destCell);
                            try { app.CutCopyMode = false; } catch { }

                            int rowsCopied = (lastRow - startRow + 1);
                            currentDestRow += rowsCopied;
                            sheetsMerged++;
                            isFirstSheet = false;
                        }
                    }
                    catch (Exception exSheet)
                    {
                        System.Diagnostics.Debug.WriteLine($"Consolidate sheet error: {exSheet.Message}");
                    }
                }

                // Kích hoạt Sheet Tổng Hợp
                try { wsSummary.Activate(); } catch { }
                try { wsSummary.Range["A:Z"].EntireColumn.AutoFit(); } catch { }

                _isBatchProcessing = false;
                try { app.EnableEvents = true; } catch { }
                try { app.DisplayAlerts = true; } catch { }
                try { app.ScreenUpdating = true; } catch { }

                RefreshWorkbookTree();

                WpfMessageBox.Show($"✅ Đã gộp thành công dữ liệu từ {sheetsMerged} sheet vào sheet [{resultSheetName}] ({currentDestRow - 1} dòng dữ liệu)!",
                                   "Gộp Sheet Thành Công", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
                return true;
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Lỗi gộp dữ liệu các sheet:\n{ex.Message}", "Gộp Sheet",
                                   System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                return false;
            }
            finally
            {
                _isBatchProcessing = false;
                if (_excelApp != null)
                {
                    try { _excelApp.EnableEvents = true; } catch { }
                    try { _excelApp.DisplayAlerts = true; } catch { }
                    try { _excelApp.ScreenUpdating = true; } catch { }
                }
            }
        }

        private static string GetExcelColumnLetter(int colIndex)
        {
            int div = colIndex;
            string colLetter = string.Empty;
            while (div > 0)
            {
                int mod = (div - 1) % 26;
                colLetter = (char)(65 + mod) + colLetter;
                div = (div - mod) / 26;
            }
            return string.IsNullOrEmpty(colLetter) ? "A" : colLetter;
        }

        public bool ImportSheetsFromExternalFiles(string targetWbName, string[] filePaths)
        {
            if (_excelApp == null || filePaths == null || filePaths.Length == 0) return false;

            _isBatchProcessing = true;
            dynamic app = _excelApp;

            try
            {
                try { app.EnableEvents = false; } catch { }
                try { app.DisplayAlerts = false; } catch { }
                try { app.ScreenUpdating = false; } catch { }

                dynamic? targetWb = null;
                if (!string.IsNullOrEmpty(targetWbName))
                {
                    try { targetWb = app.Workbooks[targetWbName]; } catch { }
                }
                if (targetWb == null)
                {
                    try { targetWb = app.ActiveWorkbook; } catch { }
                }

                if (targetWb == null) return false;

                int importedCount = 0;

                foreach (var path in filePaths)
                {
                    if (!System.IO.File.Exists(path)) continue;

                    try
                    {
                        dynamic sourceWb = app.Workbooks.Open(path, ReadOnly: true);
                        if (sourceWb != null)
                        {
                            int sourceSheetCount = sourceWb.Sheets.Count;
                            for (int i = 1; i <= sourceSheetCount; i++)
                            {
                                dynamic ws = sourceWb.Sheets[i];
                                dynamic lastSheet = targetWb.Sheets[targetWb.Sheets.Count];
                                ws.Copy(After: lastSheet);
                                importedCount++;
                            }
                            sourceWb.Close(false);
                        }
                    }
                    catch (Exception exFile)
                    {
                        System.Diagnostics.Debug.WriteLine($"Import file error: {exFile.Message}");
                    }
                }

                _isBatchProcessing = false;
                try { app.EnableEvents = true; } catch { }
                try { app.DisplayAlerts = true; } catch { }
                try { app.ScreenUpdating = true; } catch { }

                RefreshWorkbookTree();
                WpfMessageBox.Show($"✅ Đã nhập thành công {importedCount} sheet vào [{targetWb.Name}]!",
                                   "Nhập File Thành Công", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
                return true;
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Lỗi nhập sheet từ file:\n{ex.Message}", "Nhập File",
                                   System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                return false;
            }
            finally
            {
                _isBatchProcessing = false;
                if (_excelApp != null)
                {
                    try { _excelApp.EnableEvents = true; } catch { }
                    try { _excelApp.ScreenUpdating = true; } catch { }
                    try { _excelApp.DisplayAlerts = true; } catch { }
                }
            }
        }

        #endregion

        #region Vietnamese Text Scanner & Auditor

        public enum VietnameseScanScope
        {
            ActiveSheet,
            ActiveWorkbook,
            AllWorkbooks
        }

        private static readonly HashSet<char> VietnameseCharSet = new HashSet<char>(
            "àáảãạăằắẳẵặâầấẩẫậèéẻẽẹêềếểễệìíỉĩịòóỏõọôồốổỗộơờớởỡợùúủũụưừứửữựỳýỷỹỵđ" +
            "ÀÁẢÃẠĂẰẮẲẴẶÂẦẤẨẪẬÈÉẺẼẸÊỀẾỂỄỆÌÍỈĨỊÒÓỎÕỌÔỒỐỔỖỘƠỜỚỞỠỢÙÚỦŨỤƯỪỨỬỮỰỲÝỶỸỴĐ"
        );

        internal static bool HasVietnamese(string? text)
        {
            if (text == null || text.Length == 0) return false;

            // 1. Kiểm tra trực tiếp từng ký tự
            foreach (char c in text)
            {
                if (VietnameseCharSet.Contains(c)) return true;
                // Dấu tổ hợp Unicode (Combining Diacritical Marks)
                if (c >= '\u0300' && c <= '\u036F') return true;
            }

            // 2. Chuẩn hóa sang NFC (FormC) để xử lý bảng mã Unicode tổ hợp
            try
            {
                string normalized = text.Normalize(System.Text.NormalizationForm.FormC);
                foreach (char c in normalized)
                {
                    if (VietnameseCharSet.Contains(c)) return true;
                }
            }
            catch { }

            return false;
        }

        public List<VietnameseLocationItem> ScanVietnameseLocations(VietnameseScanScope scope, Action<string>? progressCallback = null)
        {
            var results = new List<VietnameseLocationItem>();

            if (_excelApp == null)
            {
                try { _excelApp = (ExcelApp)ExcelDnaUtil.Application; } catch { }
            }
            if (_excelApp == null) return results;

            dynamic app = _excelApp;
            var targetWorkbooks = new List<dynamic>();

            try
            {
                int wbCount = 0;
                try { wbCount = app.Workbooks.Count; } catch { }
                if (wbCount == 0) return results;

                if (scope == VietnameseScanScope.AllWorkbooks)
                {
                    for (int i = 1; i <= wbCount; i++)
                    {
                        try { targetWorkbooks.Add(app.Workbooks[i]); } catch { }
                    }
                }
                else
                {
                    dynamic? activeWb = null;
                    try { activeWb = app.ActiveWorkbook; } catch { }
                    if (activeWb != null)
                    {
                        targetWorkbooks.Add(activeWb);
                    }
                    else if (wbCount > 0)
                    {
                        try { targetWorkbooks.Add(app.Workbooks[1]); } catch { }
                    }
                }

                int totalFound = 0;

                foreach (var wb in targetWorkbooks)
                {
                    string wbName = string.Empty;
                    try { wbName = wb.Name; } catch { }

                    var targetSheets = new List<dynamic>();
                    if (scope == VietnameseScanScope.ActiveSheet)
                    {
                        dynamic? activeWs = null;
                        try { activeWs = app.ActiveSheet; } catch { }
                        if (activeWs != null)
                        {
                            targetSheets.Add(activeWs);
                        }
                        else
                        {
                            try { targetSheets.Add(wb.ActiveSheet ?? wb.Sheets[1]); } catch { }
                        }
                    }
                    else
                    {
                        int sheetCount = 0;
                        try { sheetCount = wb.Sheets.Count; } catch { }
                        for (int s = 1; s <= sheetCount; s++)
                        {
                            try { targetSheets.Add(wb.Sheets[s]); } catch { }
                        }
                    }

                    foreach (var ws in targetSheets)
                    {
                        string wsName = string.Empty;
                        try { wsName = ws.Name; } catch { }

                        progressCallback?.Invoke($"Đang quét: {wbName} ➔ {wsName}...");

                        // 1. Kiểm tra tên Sheet
                        if (HasVietnamese(wsName))
                        {
                            totalFound++;
                            results.Add(new VietnameseLocationItem
                            {
                                Index = totalFound,
                                WorkbookName = wbName,
                                SheetName = wsName,
                                CellAddress = "-",
                                TextContent = wsName,
                                Type = VietnameseLocationType.SheetName
                            });
                        }

                        // 2. Kiểm tra dữ liệu trong UsedRange
                        try
                        {
                            dynamic usedRange = ws.UsedRange;
                            if (usedRange != null)
                            {
                                int startRow = 1;
                                int startCol = 1;
                                try { startRow = usedRange.Row; } catch { }
                                try { startCol = usedRange.Column; } catch { }

                                object? valObj = null;
                                try { valObj = usedRange.Value2; } catch { }

                                if (valObj is Array arr)
                                {
                                    if (arr.Rank == 2)
                                    {
                                        int r1 = arr.GetLowerBound(0);
                                        int r2 = arr.GetUpperBound(0);
                                        int c1 = arr.GetLowerBound(1);
                                        int c2 = arr.GetUpperBound(1);

                                        for (int r = r1; r <= r2; r++)
                                        {
                                            for (int c = c1; c <= c2; c++)
                                            {
                                                object? cellVal = arr.GetValue(r, c);
                                                if (cellVal != null)
                                                {
                                                    string str = cellVal.ToString() ?? string.Empty;
                                                    if (HasVietnamese(str))
                                                    {
                                                        totalFound++;
                                                        string addr = GetExcelColumnLetter(startCol + (c - c1)) + (startRow + (r - r1));
                                                        results.Add(new VietnameseLocationItem
                                                        {
                                                            Index = totalFound,
                                                            WorkbookName = wbName,
                                                            SheetName = wsName,
                                                            CellAddress = addr,
                                                            TextContent = str,
                                                            Type = VietnameseLocationType.Cell
                                                        });
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    else if (arr.Rank == 1)
                                    {
                                        int i1 = arr.GetLowerBound(0);
                                        int i2 = arr.GetUpperBound(0);
                                        for (int i = i1; i <= i2; i++)
                                        {
                                            object? cellVal = arr.GetValue(i);
                                            if (cellVal != null)
                                            {
                                                string str = cellVal.ToString() ?? string.Empty;
                                                if (HasVietnamese(str))
                                                {
                                                    totalFound++;
                                                    string addr = GetExcelColumnLetter(startCol) + (startRow + (i - i1));
                                                    results.Add(new VietnameseLocationItem
                                                    {
                                                        Index = totalFound,
                                                        WorkbookName = wbName,
                                                        SheetName = wsName,
                                                        CellAddress = addr,
                                                        TextContent = str,
                                                        Type = VietnameseLocationType.Cell
                                                    });
                                                }
                                            }
                                        }
                                    }
                                }
                                else if (valObj != null)
                                {
                                    string str = valObj.ToString() ?? string.Empty;
                                    if (HasVietnamese(str))
                                    {
                                        totalFound++;
                                        string addr = GetExcelColumnLetter(startCol) + startRow;
                                        results.Add(new VietnameseLocationItem
                                        {
                                            Index = totalFound,
                                            WorkbookName = wbName,
                                            SheetName = wsName,
                                            CellAddress = addr,
                                            TextContent = str,
                                            Type = VietnameseLocationType.Cell
                                        });
                                    }
                                }
                            }
                        }
                        catch (Exception exUsed)
                        {
                            System.Diagnostics.Debug.WriteLine($"UsedRange scan error in {wsName}: {exUsed.Message}");
                        }

                        // 3. Kiểm tra Comments / Ghi chú
                        try
                        {
                            object? commentsObj = ws?.Comments;
                            if (commentsObj != null)
                            {
                                dynamic comments = commentsObj;
                                int commentCount = (int)comments.Count;
                                for (int ci = 1; ci <= commentCount; ci++)
                                {
                                    try
                                    {
                                        dynamic comment = comments[ci];
                                        string commentText = comment.Text();
                                        if (HasVietnamese(commentText))
                                        {
                                            totalFound++;
                                            string commentAddr = comment.Parent.Address.Replace("$", "");
                                            results.Add(new VietnameseLocationItem
                                            {
                                                Index = totalFound,
                                                WorkbookName = wbName,
                                                SheetName = wsName,
                                                CellAddress = commentAddr,
                                                TextContent = commentText,
                                                Type = VietnameseLocationType.Comment
                                            });
                                        }
                                    }
                                    catch { }
                                }
                            }
                        }
                        catch { }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ScanVietnameseLocations error: {ex.Message}");
            }

            return results;
        }

        public bool NavigateToCell(string workbookName, string sheetName, string cellAddress)
        {
            if (_excelApp == null)
            {
                try { _excelApp = (ExcelApp)ExcelDnaUtil.Application; } catch { }
            }
            if (_excelApp == null) return false;

            ExcelDna.Integration.ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    dynamic app = _excelApp;
                    dynamic targetWb = app.Workbooks[workbookName];
                    if (targetWb == null) return;

                    try { targetWb.Activate(); } catch { }

                    dynamic targetWs = targetWb.Sheets[sheetName];
                    if (targetWs == null) return;

                    // Nếu Sheet đang bị ẩn, tự động hiện ra để người dùng xem
                    try
                    {
                        if ((int)targetWs.Visible != (int)XlSheetVisibility.xlSheetVisible)
                        {
                            targetWs.Visible = (int)XlSheetVisibility.xlSheetVisible;
                        }
                    }
                    catch { }

                    try { targetWs.Activate(); } catch { }

                    if (!string.IsNullOrEmpty(cellAddress) && cellAddress != "-")
                    {
                        string cleanAddr = cellAddress.Split(' ')[0];
                        if (!string.IsNullOrEmpty(cleanAddr) && !cleanAddr.StartsWith("Dòng"))
                        {
                            dynamic rng = targetWs.Range[cleanAddr];
                            if (rng != null)
                            {
                                try { rng.Select(); } catch { }
                                try
                                {
                                    app.ActiveWindow.ScrollRow = rng.Row;
                                    app.ActiveWindow.ScrollColumn = rng.Column;
                                }
                                catch { }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"NavigateToCell error: {ex.Message}");
                }
            });

            return true;
        }

        public bool CreateVietnameseReportSheet(List<VietnameseLocationItem> items)
        {
            if (_excelApp == null || items == null || items.Count == 0) return false;

            _isBatchProcessing = true;
            dynamic app = _excelApp;

            try
            {
                try { app.EnableEvents = false; } catch { }
                try { app.DisplayAlerts = false; } catch { }
                try { app.ScreenUpdating = false; } catch { }

                dynamic wb = app.ActiveWorkbook;
                if (wb == null) return false;

                string reportSheetName = $"VN_Check_{DateTime.Now:yyyyMMdd_HHmm}";
                if (reportSheetName.Length > 31) reportSheetName = reportSheetName.Substring(0, 31);

                dynamic firstSheet = wb.Sheets[1];
                dynamic wsReport = wb.Worksheets.Add(Before: firstSheet);
                wsReport.Name = reportSheetName;
                wsReport.Tab.Color = 0x2563EB; // Màu xanh nổi bật

                // Tiêu đề
                wsReport.Range["A1"].Value = "BÁO CÁO RÀ SOÁT CÁC VỊ TRÍ CHỨA TIẾNG VIỆT";
                wsReport.Range["A1"].Font.Bold = true;
                wsReport.Range["A1"].Font.Size = 14;
                wsReport.Range["A1"].Font.Color = 0x1E3A8A;

                wsReport.Range["A2"].Value = $"Thời gian quét: {DateTime.Now:dd/MM/yyyy HH:mm:ss} | Tổng số vị trí: {items.Count}";
                wsReport.Range["A2"].Font.Italic = true;
                wsReport.Range["A2"].Font.Size = 10;
                wsReport.Range["A2"].Font.Color = 0x64748B;

                // Headers
                string[] headers = new string[] { "STT", "Tên File (Workbook)", "Tên Sheet", "Địa Chỉ Ô", "Loại Vị Trí", "Nội Dung Tiếng Việt" };
                for (int h = 0; h < headers.Length; h++)
                {
                    dynamic cell = wsReport.Cells[4, h + 1];
                    cell.Value = headers[h];
                    cell.Font.Bold = true;
                    cell.Interior.Color = 0x107C41; // Green
                    cell.Font.Color = 0xFFFFFF; // White
                }

                // Ghi dữ liệu
                for (int i = 0; i < items.Count; i++)
                {
                    var item = items[i];
                    int r = 5 + i;

                    wsReport.Cells[r, 1].Value = i + 1;
                    wsReport.Cells[r, 2].Value = item.WorkbookName;
                    wsReport.Cells[r, 3].Value = item.SheetName;
                    wsReport.Cells[r, 4].Value = item.CellAddress;
                    wsReport.Cells[r, 5].Value = item.TypeDescription;
                    wsReport.Cells[r, 6].Value = item.TextContent;

                    // Thêm Hyperlink nếu có địa chỉ ô
                    if (!string.IsNullOrEmpty(item.CellAddress) && item.CellAddress != "-")
                    {
                        try
                        {
                            string subAddress = $"'{item.SheetName}'!{item.CellAddress}";
                            dynamic cellLink = wsReport.Cells[r, 4];
                            wsReport.Hyperlinks.Add(Anchor: cellLink, Address: "", SubAddress: subAddress, TextToDisplay: item.CellAddress);
                        }
                        catch { }
                    }
                }

                // Định dạng kẻ bảng
                int lastRow = 4 + items.Count;
                dynamic tableRange = wsReport.Range[$"A4:F{lastRow}"];
                tableRange.Borders.LineStyle = 1; // xlContinuous
                tableRange.Borders.Color = 0xCBD5E1;

                // Tự động căn chỉnh độ rộng cột
                wsReport.Columns["A:F"].AutoFit();

                _isBatchProcessing = false;
                try { app.EnableEvents = true; } catch { }
                try { app.DisplayAlerts = true; } catch { }
                try { app.ScreenUpdating = true; } catch { }

                RefreshWorkbookTree();
                WpfMessageBox.Show($"✅ Đã tạo thành công Sheet báo cáo [{reportSheetName}] với {items.Count} vị trí tiếng Việt!",
                                   "Tạo Báo Cáo Thành Công", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
                return true;
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Lỗi tạo sheet báo cáo:\n{ex.Message}", "Lỗi Báo Cáo",
                                   System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                return false;
            }
            finally
            {
                _isBatchProcessing = false;
                if (_excelApp != null)
                {
                    try { _excelApp.EnableEvents = true; } catch { }
                    try { _excelApp.ScreenUpdating = true; } catch { }
                    try { _excelApp.DisplayAlerts = true; } catch { }
                }
            }
        }

        #endregion

        #region Workbook & Sheet Diff & Compare Engine

        public List<string> GetOpenWorkbookNamesList()
        {
            var list = new List<string>();
            if (_excelApp == null)
            {
                try { _excelApp = (ExcelApp)ExcelDnaUtil.Application; } catch { }
            }
            if (_excelApp == null) return list;

            try
            {
                dynamic app = _excelApp;
                int count = app.Workbooks.Count;
                for (int i = 1; i <= count; i++)
                {
                    try { list.Add(app.Workbooks[i].Name); } catch { }
                }
            }
            catch { }
            return list;
        }

        public List<string> GetWorksheetNamesList(string workbookName)
        {
            var list = new List<string>();
            if (_excelApp == null || string.IsNullOrEmpty(workbookName)) return list;

            try
            {
                dynamic app = _excelApp;
                dynamic wb = app.Workbooks[workbookName];
                if (wb != null)
                {
                    int count = wb.Sheets.Count;
                    for (int i = 1; i <= count; i++)
                    {
                        try { list.Add(wb.Sheets[i].Name); } catch { }
                    }
                }
            }
            catch { }
            return list;
        }

        public List<CompareDiffItem> CompareWorkbooksOrSheets(
            string wb1Name, 
            string? ws1Name, 
            string wb2Name, 
            string? ws2Name, 
            CompareOptions options, 
            Action<string>? progressCallback = null)
        {
            var results = new List<CompareDiffItem>();
            if (_excelApp == null)
            {
                try { _excelApp = (ExcelApp)ExcelDnaUtil.Application; } catch { }
            }
            if (_excelApp == null) return results;

            dynamic app = _excelApp;

            try
            {
                object? rawWb1 = null;
                object? rawWb2 = null;

                try { rawWb1 = app.Workbooks[wb1Name]; } catch { }
                try { rawWb2 = app.Workbooks[wb2Name]; } catch { }

                if (rawWb1 == null || rawWb2 == null)
                {
                    progressCallback?.Invoke("Không tìm thấy Workbook đã chọn.");
                    return results;
                }

                dynamic wb1 = rawWb1;
                dynamic wb2 = rawWb2;

                var sheetPairs = new List<(object ws1, object ws2, string sName1, string sName2)>();

                // Nếu người dùng chọn đích danh Sheet1 vs Sheet2
                if (!string.IsNullOrEmpty(ws1Name) && !string.IsNullOrEmpty(ws2Name))
                {
                    object? s1 = null;
                    object? s2 = null;
                    try { s1 = wb1.Sheets[ws1Name!]; } catch { }
                    try { s2 = wb2.Sheets[ws2Name!]; } catch { }
                    if (s1 != null && s2 != null)
                    {
                        sheetPairs.Add((s1, s2, ws1Name!, ws2Name!));
                    }
                }
                else
                {
                    // So sánh toàn bộ các Sheet có cùng tên
                    dynamic sheets1 = wb1.Sheets;
                    int count1 = (int)sheets1.Count;
                    for (int i = 1; i <= count1; i++)
                    {
                        try
                        {
                            dynamic s1 = sheets1[i];
                            string sName = (string)s1.Name;
                            object? s2 = null;
                            try { s2 = wb2.Sheets[sName]; } catch { }
                            if (s2 != null)
                            {
                                sheetPairs.Add((s1, s2, sName, sName));
                            }
                        }
                        catch { }
                    }
                }

                if (sheetPairs.Count == 0)
                {
                    progressCallback?.Invoke("Không tìm thấy Sheet nào phù hợp để so sánh.");
                    return results;
                }

                int totalDiffCount = 0;

                foreach (var pair in sheetPairs)
                {
                    progressCallback?.Invoke($"Đang so sánh Sheet [{pair.sName1}]...");
                    dynamic ws1 = pair.ws1;
                    dynamic ws2 = pair.ws2;

                    if (options.Mode == CompareMode.CellByCell)
                    {
                        CompareSheetCellByCell(ws1, ws2, pair.sName1, wb1Name, wb2Name, options, results, ref totalDiffCount);
                    }
                    else
                    {
                        CompareSheetByKeyColumn(ws1, ws2, pair.sName1, wb1Name, wb2Name, options, results, ref totalDiffCount);
                    }
                }

                progressCallback?.Invoke($"Hoàn tất! Tìm thấy {results.Count} điểm sai khác.");
            }
            catch (Exception ex)
            {
                progressCallback?.Invoke($"Lỗi so sánh: {ex.Message}");
            }

            return results;
        }

        private void CompareSheetCellByCell(
            dynamic ws1, dynamic ws2, 
            string sheetName, string wb1Name, string wb2Name, 
            CompareOptions options, 
            List<CompareDiffItem> results, 
            ref int totalDiffCount)
        {
            try
            {
                dynamic u1 = ws1.UsedRange;
                dynamic u2 = ws2.UsedRange;

                int r1Start = 1, c1Start = 1, r1Count = 0, c1Count = 0;
                int r2Start = 1, c2Start = 1, r2Count = 0, c2Count = 0;

                if (u1 != null)
                {
                    try { r1Start = u1.Row; c1Start = u1.Column; r1Count = u1.Rows.Count; c1Count = u1.Columns.Count; } catch { }
                }
                if (u2 != null)
                {
                    try { r2Start = u2.Row; c2Start = u2.Column; r2Count = u2.Rows.Count; c2Count = u2.Columns.Count; } catch { }
                }

                if (r1Count == 0 && r2Count == 0) return;

                int minRow = Math.Min(r1Start, r2Start);
                int minCol = Math.Min(c1Start, c2Start);
                int maxRow = Math.Max(r1Start + r1Count - 1, r2Start + r2Count - 1);
                int maxCol = Math.Max(c1Start + c1Count - 1, c2Start + c2Count - 1);

                object? valObj1 = options.CompareFormulas ? u1?.Formula : u1?.Value2;
                object? valObj2 = options.CompareFormulas ? u2?.Formula : u2?.Value2;

                for (int r = minRow; r <= maxRow; r++)
                {
                    for (int c = minCol; c <= maxCol; c++)
                    {
                        string str1 = GetValueFromArray(valObj1, r1Start, c1Start, r1Count, c1Count, r, c);
                        string str2 = GetValueFromArray(valObj2, r2Start, c2Start, r2Count, c2Count, r, c);

                        string norm1 = NormalizeCompareString(str1, options);
                        string norm2 = NormalizeCompareString(str2, options);

                        if (!string.Equals(norm1, norm2, options.CaseInsensitive ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal))
                        {
                            totalDiffCount++;
                            DiffType diffType = DiffType.Modified;

                            if (string.IsNullOrEmpty(norm1) && !string.IsNullOrEmpty(norm2))
                            {
                                diffType = DiffType.Added;
                            }
                            else if (!string.IsNullOrEmpty(norm1) && string.IsNullOrEmpty(norm2))
                            {
                                diffType = DiffType.Deleted;
                            }

                            string addr = GetExcelColumnLetter(c) + r;

                            results.Add(new CompareDiffItem
                            {
                                Index = totalDiffCount,
                                SheetName = sheetName,
                                CellAddress = addr,
                                KeyIdentifier = addr,
                                Type = diffType,
                                OldValue = str1,
                                NewValue = str2,
                                Workbook1Name = wb1Name,
                                Workbook2Name = wb2Name
                            });
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error CompareSheetCellByCell: {ex.Message}");
            }
        }

        private void CompareSheetByKeyColumn(
            dynamic ws1, dynamic ws2, 
            string sheetName, string wb1Name, string wb2Name, 
            CompareOptions options, 
            List<CompareDiffItem> results, 
            ref int totalDiffCount)
        {
            try
            {
                dynamic u1 = ws1.UsedRange;
                dynamic u2 = ws2.UsedRange;

                int keyCol = options.KeyColumnIndex; // 1-based index (e.g. 1 = col A)

                var dict1 = ExtractRowsByKey(u1, keyCol, options);
                var dict2 = ExtractRowsByKey(u2, keyCol, options);

                // 1. Kiểm tra các dòng có trong cả 2 hoặc bị sửa đổi
                foreach (var kvp in dict1)
                {
                    string key = kvp.Key;
                    int row1Index = kvp.Value.row;
                    var row1Vals = kvp.Value.values;

                    if (dict2.TryGetValue(key, out (int row, string[] values) row2Data))
                    {
                        int row2Index = row2Data.row;
                        var row2Vals = row2Data.values;

                        int maxCols = Math.Max(row1Vals.Length, row2Vals.Length);
                        for (int col = 0; col < maxCols; col++)
                        {
                            string s1 = col < row1Vals.Length ? row1Vals[col] : string.Empty;
                            string s2 = col < row2Vals.Length ? row2Vals[col] : string.Empty;

                            string n1 = NormalizeCompareString(s1, options);
                            string n2 = NormalizeCompareString(s2, options);

                            if (!string.Equals(n1, n2, options.CaseInsensitive ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal))
                            {
                                totalDiffCount++;
                                string colLetter = GetExcelColumnLetter(col + 1);
                                string addr = $"{colLetter}{row2Index} (A:{colLetter}{row1Index})";

                                results.Add(new CompareDiffItem
                                {
                                    Index = totalDiffCount,
                                    SheetName = sheetName,
                                    CellAddress = addr,
                                    KeyIdentifier = $"Khóa: [{key}] - Cột {colLetter}",
                                    Type = DiffType.Modified,
                                    OldValue = s1,
                                    NewValue = s2,
                                    Workbook1Name = wb1Name,
                                    Workbook2Name = wb2Name
                                });
                            }
                        }
                    }
                    else
                    {
                        // Dòng có ở File A nhưng bị xóa ở File B
                        totalDiffCount++;
                        results.Add(new CompareDiffItem
                        {
                            Index = totalDiffCount,
                            SheetName = sheetName,
                            CellAddress = $"Dòng {row1Index}",
                            KeyIdentifier = $"Khóa: [{key}]",
                            Type = DiffType.Deleted,
                            OldValue = string.Join(" | ", row1Vals),
                            NewValue = "(Đã bị xóa khỏi File B)",
                            Workbook1Name = wb1Name,
                            Workbook2Name = wb2Name
                        });
                    }
                }

                // 2. Kiểm tra các dòng thêm mới ở File B
                foreach (var kvp in dict2)
                {
                    string key = kvp.Key;
                    if (!dict1.ContainsKey(key))
                    {
                        totalDiffCount++;
                        int row2Index = kvp.Value.row;
                        var row2Vals = kvp.Value.values;

                        results.Add(new CompareDiffItem
                        {
                            Index = totalDiffCount,
                            SheetName = sheetName,
                            CellAddress = $"Dòng {row2Index}",
                            KeyIdentifier = $"Khóa: [{key}]",
                            Type = DiffType.Added,
                            OldValue = "(Không có trong File A)",
                            NewValue = string.Join(" | ", row2Vals),
                            Workbook1Name = wb1Name,
                            Workbook2Name = wb2Name
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error CompareSheetByKeyColumn: {ex.Message}");
            }
        }

        private Dictionary<string, (int row, string[] values)> ExtractRowsByKey(dynamic? usedRange, int keyCol, CompareOptions options)
        {
            var dict = new Dictionary<string, (int row, string[] values)>(StringComparer.OrdinalIgnoreCase);
            if (usedRange == null) return dict;

            try
            {
                int startRow = usedRange.Row;
                int startCol = usedRange.Column;
                int rowCount = usedRange.Rows.Count;
                int colCount = usedRange.Columns.Count;

                object? valObj = options.CompareFormulas ? usedRange.Formula : usedRange.Value2;

                for (int r = startRow; r < startRow + rowCount; r++)
                {
                    var rowVals = new string[colCount];
                    string keyVal = string.Empty;

                    for (int c = startCol; c < startCol + colCount; c++)
                    {
                        string str = GetValueFromArray(valObj, startRow, startCol, rowCount, colCount, r, c);
                        int colIdx = c - startCol;
                        rowVals[colIdx] = str;

                        if (c == keyCol)
                        {
                            keyVal = NormalizeCompareString(str, options);
                        }
                    }

                    if (!string.IsNullOrEmpty(keyVal) && !dict.ContainsKey(keyVal))
                    {
                        dict[keyVal] = (r, rowVals);
                    }
                }
            }
            catch { }

            return dict;
        }

        private string GetValueFromArray(object? valObj, int startRow, int startCol, int rowCount, int colCount, int targetRow, int targetCol)
        {
            if (valObj == null) return string.Empty;

            if (valObj is Array arr)
            {
                if (arr.Rank == 2)
                {
                    int r1 = arr.GetLowerBound(0);
                    int c1 = arr.GetLowerBound(1);
                    int rIdx = r1 + (targetRow - startRow);
                    int cIdx = c1 + (targetCol - startCol);

                    if (rIdx >= arr.GetLowerBound(0) && rIdx <= arr.GetUpperBound(0) &&
                        cIdx >= arr.GetLowerBound(1) && cIdx <= arr.GetUpperBound(1))
                    {
                        object? item = arr.GetValue(rIdx, cIdx);
                        return item?.ToString() ?? string.Empty;
                    }
                }
                else if (arr.Rank == 1)
                {
                    int i1 = arr.GetLowerBound(0);
                    int idx = i1 + (targetRow - startRow);
                    if (idx >= arr.GetLowerBound(0) && idx <= arr.GetUpperBound(0))
                    {
                        object? item = arr.GetValue(idx);
                        return item?.ToString() ?? string.Empty;
                    }
                }
            }
            else if (targetRow == startRow && targetCol == startCol)
            {
                return valObj.ToString() ?? string.Empty;
            }

            return string.Empty;
        }

        private string NormalizeCompareString(string? text, CompareOptions options)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;
            string res = text!;
            if (options.IgnoreWhitespace)
            {
                res = res.Trim();
            }
            return res;
        }

        public bool HighlightDiffInWorksheet(List<CompareDiffItem> diffs, string wbName, string wsName)
        {
            if (_excelApp == null || diffs == null || diffs.Count == 0) return false;

            try
            {
                dynamic app = _excelApp;
                dynamic wb = app.Workbooks[wbName];
                dynamic ws = wb.Sheets[wsName];

                _isBatchProcessing = true;
                try { app.ScreenUpdating = false; } catch { }

                foreach (var diff in diffs)
                {
                    if (!string.Equals(diff.SheetName, wsName, StringComparison.OrdinalIgnoreCase)) continue;

                    string cleanAddr = diff.CellAddress.Split(' ')[0];
                    if (string.IsNullOrEmpty(cleanAddr) || cleanAddr.StartsWith("Dòng")) continue;

                    try
                    {
                        dynamic rng = ws.Range[cleanAddr];
                        if (rng != null)
                        {
                            switch (diff.Type)
                            {
                                case DiffType.Modified:
                                    rng.Interior.Color = 0x82E0FF; // BGR: Light Amber/Yellow (RGB: 255, 224, 130)
                                    break;
                                case DiffType.Added:
                                    rng.Interior.Color = 0xC8E6C9; // BGR: Light Green (RGB: 200, 230, 201)
                                    break;
                                case DiffType.Deleted:
                                    rng.Interior.Color = 0xCCD2FF; // BGR: Light Red/Pink (RGB: 255, 210, 204)
                                    break;
                            }
                        }
                    }
                    catch { }
                }

                _isBatchProcessing = false;
                try { app.ScreenUpdating = true; } catch { }

                WpfMessageBox.Show($"✅ Đã tô màu nổi bật các ô sai khác trên Sheet [{wsName}] thành công!\n- Màu Vàng: Ô thay đổi giá trị\n- Màu Xanh: Ô/Dòng thêm mới\n- Màu Đỏ: Ô/Dòng bị xóa",
                                   "Tô Màu Hoàn Tất", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
                return true;
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Lỗi tô màu sai khác:\n{ex.Message}", "Lỗi Tô Màu",
                                   System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                return false;
            }
            finally
            {
                _isBatchProcessing = false;
                if (_excelApp != null)
                {
                    try { _excelApp.ScreenUpdating = true; } catch { }
                }
            }
        }

        public bool CreateDiffReportSheet(string targetWbName, List<CompareDiffItem> diffs)
        {
            if (_excelApp == null || diffs == null || diffs.Count == 0) return false;

            try
            {
                dynamic app = _excelApp;
                dynamic wb = app.Workbooks[targetWbName];
                if (wb == null) return false;

                _isBatchProcessing = true;
                try { app.ScreenUpdating = false; } catch { }
                try { app.DisplayAlerts = false; } catch { }

                string reportSheetName = $"Diff_Report_{DateTime.Now:yyyyMMdd_HHmm}";
                dynamic wsReport = wb.Sheets.Add(Before: wb.Sheets[1]);
                wsReport.Name = reportSheetName;
                wsReport.Tab.Color = 0xD97706;

                // Tiêu đề lớn
                wsReport.Range["A1:F1"].Merge();
                wsReport.Cells[1, 1] = $"BÁO CÁO SAI KHÁC DỮ LIỆU (WORKBOOK COMPARE REPORT) — {DateTime.Now:dd/MM/yyyy HH:mm:ss}";
                dynamic titleRange = wsReport.Range["A1:F1"];
                titleRange.Font.Bold = true;
                titleRange.Font.Size = 13;
                titleRange.Font.Color = 0xFFFFFF;
                titleRange.Interior.Color = 0x1E293B; // Dark Slate Header
                titleRange.HorizontalAlignment = -4108; // xlCenter

                // Thông tin file so sánh
                wsReport.Cells[2, 1] = $"File Gốc (A): {diffs[0].Workbook1Name}  ⇋  File Mới (B): {diffs[0].Workbook2Name}  |  Tổng số sai khác: {diffs.Count}";
                dynamic subRange = wsReport.Range["A2:F2"];
                subRange.Merge();
                subRange.Font.Italic = true;
                subRange.Font.Size = 10.5;
                subRange.Font.Color = 0x475569;

                // Header Cột
                string[] headers = { "STT", "Tên Sheet", "Vị Trí / Khóa", "Loại Sai Khác", "Giá Trị File A (Gốc)", "Giá Trị File B (Mới)" };
                for (int c = 0; c < headers.Length; c++)
                {
                    dynamic headerCell = wsReport.Cells[3, c + 1];
                    headerCell.Value2 = headers[c];
                    headerCell.Font.Bold = true;
                    headerCell.Font.Color = 0x0F172A;
                    headerCell.Interior.Color = 0xE2E8F0; // Slate 200
                    headerCell.HorizontalAlignment = -4108; // xlCenter
                }

                // Điền dữ liệu
                for (int i = 0; i < diffs.Count; i++)
                {
                    var item = diffs[i];
                    int r = 4 + i;

                    wsReport.Cells[r, 1] = item.Index;
                    wsReport.Cells[r, 2] = item.SheetName;
                    wsReport.Cells[r, 3] = item.CellAddress;
                    wsReport.Cells[r, 4] = item.TypeDescription;
                    wsReport.Cells[r, 5] = item.OldValue;
                    wsReport.Cells[r, 6] = item.NewValue;

                    // Màu badge loại sai khác
                    dynamic typeCell = wsReport.Cells[r, 4];
                    typeCell.HorizontalAlignment = -4108;
                    switch (item.Type)
                    {
                        case DiffType.Modified:
                            typeCell.Interior.Color = 0xFEF3C7;
                            typeCell.Font.Color = 0x92400E;
                            break;
                        case DiffType.Added:
                            typeCell.Interior.Color = 0xDCFCE7;
                            typeCell.Font.Color = 0x166534;
                            break;
                        case DiffType.Deleted:
                            typeCell.Interior.Color = 0xFEE2E2;
                            typeCell.Font.Color = 0x991B1B;
                            break;
                    }

                    // Hyperlink đến vị trí ô trên Sheet
                    string cleanAddr = item.CellAddress.Split(' ')[0];
                    if (!string.IsNullOrEmpty(cleanAddr) && !cleanAddr.StartsWith("Dòng"))
                    {
                        try
                        {
                            string subAddress = $"'{item.SheetName}'!{cleanAddr}";
                            dynamic cellLink = wsReport.Cells[r, 3];
                            wsReport.Hyperlinks.Add(Anchor: cellLink, Address: "", SubAddress: subAddress, TextToDisplay: item.CellAddress);
                        }
                        catch { }
                    }
                }

                // Định dạng kẻ bảng
                int lastRow = 3 + diffs.Count;
                dynamic tableRange = wsReport.Range[$"A3:F{lastRow}"];
                tableRange.Borders.LineStyle = 1;
                tableRange.Borders.Color = 0xCBD5E1;

                // Tự động căn chỉnh độ rộng cột
                wsReport.Columns["A:F"].AutoFit();

                _isBatchProcessing = false;
                try { app.ScreenUpdating = true; } catch { }
                try { app.DisplayAlerts = true; } catch { }

                RefreshWorkbookTree();
                WpfMessageBox.Show($"✅ Đã tạo thành công Sheet báo cáo sai khác [{reportSheetName}] với {diffs.Count} mục!",
                                   "Tạo Báo Cáo Thành Công", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
                return true;
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Lỗi tạo sheet báo cáo sai khác:\n{ex.Message}", "Lỗi Báo Cáo",
                                   System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                return false;
            }
            finally
            {
                _isBatchProcessing = false;
                if (_excelApp != null)
                {
                    try { _excelApp.ScreenUpdating = true; } catch { }
                    try { _excelApp.DisplayAlerts = true; } catch { }
                }
            }
        }

        #endregion
    }
}

