using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;

namespace ExcelSupport.Services
{
    public static class TranslationUndoHelper
    {
        public class CellBackupItem
        {
            public int Row { get; set; }
            public int Column { get; set; }
            public object? OldValue { get; set; }
            public object? NewValue { get; set; }
        }

        public class UndoRecord
        {
            public string SheetName { get; set; } = string.Empty;
            public string? WorkbookName { get; set; }
            public List<CellBackupItem> Cells { get; set; } = new List<CellBackupItem>();
        }

        private static readonly Stack<UndoRecord> _undoStack = new Stack<UndoRecord>();
        private static readonly Stack<UndoRecord> _redoStack = new Stack<UndoRecord>();

        public const string UndoMacroName = "ExcelSupport_UndoTranslation";
        public const string RedoMacroName = "ExcelSupport_RedoTranslation";

        public static void RecordAndApply(_Worksheet ws, List<CellBackupItem> items, string actionName = "Dịch Thuật AI")
        {
            if (ws == null || items == null || items.Count == 0) return;

            var record = new UndoRecord
            {
                SheetName = ws.Name,
                Cells = items
            };

            try
            {
                var wb = ws.Parent as Workbook;
                record.WorkbookName = wb?.Name;
            }
            catch { }

            _undoStack.Push(record);
            _redoStack.Clear();

            try
            {
                var app = (ExcelApp)ExcelDnaUtil.Application;

                // 1. Đăng ký với menu Undo chuẩn của Excel
                try
                {
                    app.OnUndo($"Hoàn tác {actionName}", UndoMacroName);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"OnUndo error: {ex.Message}");
                }

                // 2. Gán trực tiếp phím tắt Ctrl+Z & Ctrl+z với tên macro XLL đã đăng ký
                try
                {
                    app.OnKey("^z", UndoMacroName);
                    app.OnKey("^Z", UndoMacroName);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"OnKey registration error: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"RecordAndApply error: {ex.Message}");
            }
        }

        [ExcelCommand(Name = UndoMacroName)]
        public static void UndoTranslation()
        {
            ExcelApp? app = null;
            try
            {
                app = (ExcelApp)ExcelDnaUtil.Application;
                if (app != null)
                {
                    // Trả lại phím tắt Ctrl+Z mặc định của Excel
                    app.OnKey("^z");
                    app.OnKey("^Z");
                }

                if (_undoStack.Count == 0)
                {
                    try { app?.Undo(); } catch { }
                    return;
                }

                var record = _undoStack.Pop();
                if (app == null) return;

                _Worksheet? ws = null;
                try
                {
                    if (!string.IsNullOrEmpty(record.WorkbookName))
                    {
                        var wb = app.Workbooks[record.WorkbookName];
                        ws = wb.Sheets[record.SheetName] as _Worksheet;
                    }
                    else
                    {
                        ws = app.ActiveSheet as _Worksheet;
                    }
                }
                catch
                {
                    ws = app.ActiveSheet as _Worksheet;
                }

                if (ws == null) return;

                app.ScreenUpdating = false;

                foreach (var cellItem in record.Cells)
                {
                    Range? cell = null;
                    try
                    {
                        cell = ws.Cells[cellItem.Row, cellItem.Column] as Range;
                        if (cell != null)
                        {
                            cell.Value2 = cellItem.OldValue;
                        }
                    }
                    finally
                    {
                        if (cell != null) Marshal.ReleaseComObject(cell);
                    }
                }

                _redoStack.Push(record);

                // Thiết lập Redo (Ctrl+Y)
                try
                {
                    app.OnRepeat("Làm lại Dịch Thuật AI", RedoMacroName);
                }
                catch { }

                try
                {
                    app.OnKey("^y", RedoMacroName);
                    app.OnKey("^Y", RedoMacroName);
                }
                catch { }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"UndoTranslation error: {ex.Message}");
            }
            finally
            {
                try
                {
                    if (app != null) app.ScreenUpdating = true;
                }
                catch { }
            }
        }

        [ExcelCommand(Name = RedoMacroName)]
        public static void RedoTranslation()
        {
            ExcelApp? app = null;
            try
            {
                app = (ExcelApp)ExcelDnaUtil.Application;
                if (app != null)
                {
                    // Trả lại phím tắt Ctrl+Y mặc định của Excel
                    app.OnKey("^y");
                    app.OnKey("^Y");
                }

                if (_redoStack.Count == 0) return;

                var record = _redoStack.Pop();
                if (app == null) return;

                _Worksheet? ws = null;
                try
                {
                    if (!string.IsNullOrEmpty(record.WorkbookName))
                    {
                        var wb = app.Workbooks[record.WorkbookName];
                        ws = wb.Sheets[record.SheetName] as _Worksheet;
                    }
                    else
                    {
                        ws = app.ActiveSheet as _Worksheet;
                    }
                }
                catch
                {
                    ws = app.ActiveSheet as _Worksheet;
                }

                if (ws == null) return;

                app.ScreenUpdating = false;

                foreach (var cellItem in record.Cells)
                {
                    Range? cell = null;
                    try
                    {
                        cell = ws.Cells[cellItem.Row, cellItem.Column] as Range;
                        if (cell != null)
                        {
                            cell.Value2 = cellItem.NewValue;
                        }
                    }
                    finally
                    {
                        if (cell != null) Marshal.ReleaseComObject(cell);
                    }
                }

                _undoStack.Push(record);

                // Đăng ký lại Undo (Ctrl+Z)
                try
                {
                    app.OnUndo("Hoàn tác Dịch Thuật AI", UndoMacroName);
                }
                catch { }

                try
                {
                    app.OnKey("^z", UndoMacroName);
                    app.OnKey("^Z", UndoMacroName);
                }
                catch { }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"RedoTranslation error: {ex.Message}");
            }
            finally
            {
                try
                {
                    if (app != null) app.ScreenUpdating = true;
                }
                catch { }
            }
        }
    }
}
