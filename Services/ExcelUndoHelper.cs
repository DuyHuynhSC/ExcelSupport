using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelSupport.Models;
using Microsoft.Office.Interop.Excel;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;

namespace ExcelSupport.Services
{
    public interface IUndoableAction
    {
        string ActionName { get; }
        void Undo(ExcelApp app);
        void Redo(ExcelApp app);
    }

    public class BorderSnapshot
    {
        public object? LineStyle { get; set; }
        public object? Weight { get; set; }
        public object? Color { get; set; }
    }

    public class CellFormatSnapshot
    {
        public int Row { get; set; }
        public int Column { get; set; }

        // Font
        public string? FontName { get; set; }
        public object? FontSize { get; set; }
        public object? FontBold { get; set; }
        public object? FontItalic { get; set; }
        public object? FontUnderline { get; set; }
        public object? FontColor { get; set; }

        // Interior
        public object? InteriorColor { get; set; }
        public object? InteriorColorIndex { get; set; }

        // Alignment
        public object? HorizontalAlignment { get; set; }
        public object? VerticalAlignment { get; set; }
        public object? WrapText { get; set; }

        // NumberFormat
        public object? NumberFormat { get; set; }

        // Borders
        public BorderSnapshot? Top { get; set; }
        public BorderSnapshot? Bottom { get; set; }
        public BorderSnapshot? Left { get; set; }
        public BorderSnapshot? Right { get; set; }
    }

    public class JapaneseFormatUndoAction : IUndoableAction
    {
        public string SheetName { get; set; } = string.Empty;
        public string? WorkbookName { get; set; }
        public string RangeAddress { get; set; } = string.Empty;
        public JapaneseFormatPreset? Preset { get; set; }
        public string ActionName { get; set; } = string.Empty;
        public List<CellFormatSnapshot> Snapshots { get; set; } = new List<CellFormatSnapshot>();

        public void Undo(ExcelApp app)
        {
            _Worksheet? ws = GetWorksheet(app);
            if (ws == null) return;

            try
            {
                ExcelUndoHelper.RestoreCellFormatting(ws, Snapshots);
            }
            finally
            {
                Marshal.ReleaseComObject(ws);
            }
        }

        public void Redo(ExcelApp app)
        {
            _Worksheet? ws = GetWorksheet(app);
            if (ws == null) return;

            Range? targetRange = null;
            try
            {
                if (!string.IsNullOrEmpty(RangeAddress) && Preset != null)
                {
                    targetRange = ws.Range[RangeAddress];
                    if (targetRange != null)
                    {
                        JapanesePresetFormatService.ApplyPresetDirect(app, targetRange, Preset);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[JapaneseFormatUndoAction] Redo error: {ex.Message}");
            }
            finally
            {
                if (targetRange != null) Marshal.ReleaseComObject(targetRange);
                Marshal.ReleaseComObject(ws);
            }
        }

        private _Worksheet? GetWorksheet(ExcelApp app)
        {
            try
            {
                if (!string.IsNullOrEmpty(WorkbookName))
                {
                    var wb = app.Workbooks[WorkbookName];
                    return wb.Sheets[SheetName] as _Worksheet;
                }
                return app.ActiveSheet as _Worksheet;
            }
            catch
            {
                try { return app.ActiveSheet as _Worksheet; } catch { return null; }
            }
        }
    }

    public class TranslationUndoAction : IUndoableAction
    {
        public string SheetName { get; set; } = string.Empty;
        public string? WorkbookName { get; set; }
        public string ActionName { get; set; } = "Dịch Thuật AI";
        public List<TranslationUndoHelper.CellBackupItem> Cells { get; set; } = new List<TranslationUndoHelper.CellBackupItem>();

        public void Undo(ExcelApp app)
        {
            _Worksheet? ws = GetWorksheet(app);
            if (ws == null) return;

            try
            {
                foreach (var cellItem in Cells)
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
            }
            finally
            {
                Marshal.ReleaseComObject(ws);
            }
        }

        public void Redo(ExcelApp app)
        {
            _Worksheet? ws = GetWorksheet(app);
            if (ws == null) return;

            try
            {
                foreach (var cellItem in Cells)
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
            }
            finally
            {
                Marshal.ReleaseComObject(ws);
            }
        }

        private _Worksheet? GetWorksheet(ExcelApp app)
        {
            try
            {
                if (!string.IsNullOrEmpty(WorkbookName))
                {
                    var wb = app.Workbooks[WorkbookName];
                    return wb.Sheets[SheetName] as _Worksheet;
                }
                return app.ActiveSheet as _Worksheet;
            }
            catch
            {
                try { return app.ActiveSheet as _Worksheet; } catch { return null; }
            }
        }
    }

    public static class ExcelUndoHelper
    {
        public const string GlobalUndoMacroName = "ExcelSupport_GlobalUndo";
        public const string GlobalRedoMacroName = "ExcelSupport_GlobalRedo";

        private static readonly Stack<IUndoableAction> _undoStack = new Stack<IUndoableAction>();
        private static readonly Stack<IUndoableAction> _redoStack = new Stack<IUndoableAction>();

        public static void RecordAction(IUndoableAction action)
        {
            if (action == null) return;

            _undoStack.Push(action);
            _redoStack.Clear();

            try
            {
                var app = (ExcelApp)ExcelDnaUtil.Application;
                if (app == null) return;

                // 1. Đăng ký tên hiển thị với menu Undo của Excel
                try
                {
                    app.OnUndo($"Hoàn tác {action.ActionName}", GlobalUndoMacroName);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[ExcelUndoHelper] OnUndo error: {ex.Message}");
                }

                // 2. Gán trực tiếp phím tắt Ctrl+Z
                try
                {
                    app.OnKey("^z", GlobalUndoMacroName);
                    app.OnKey("^Z", GlobalUndoMacroName);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[ExcelUndoHelper] OnKey ^z error: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ExcelUndoHelper] RecordAction error: {ex.Message}");
            }
        }

        [ExcelCommand(Name = GlobalUndoMacroName)]
        public static void GlobalUndo()
        {
            ExcelApp? app = null;
            try
            {
                app = (ExcelApp)ExcelDnaUtil.Application;
                if (app != null)
                {
                    // Tạm giải phóng OnKey ^z để tránh vòng lặp
                    app.OnKey("^z");
                    app.OnKey("^Z");
                }

                if (_undoStack.Count == 0)
                {
                    try { app?.Undo(); } catch { }
                    return;
                }

                var action = _undoStack.Pop();
                if (app == null) return;

                bool oldScreenUpdating = app.ScreenUpdating;
                XlCalculation oldCalculation = app.Calculation;
                bool oldEnableEvents = app.EnableEvents;

                try
                {
                    app.ScreenUpdating = false;
                    app.EnableEvents = false;
                    if (app.Calculation != XlCalculation.xlCalculationManual)
                    {
                        app.Calculation = XlCalculation.xlCalculationManual;
                    }

                    action.Undo(app);
                }
                finally
                {
                    try
                    {
                        app.Calculation = oldCalculation;
                        app.EnableEvents = oldEnableEvents;
                        app.ScreenUpdating = oldScreenUpdating;
                    }
                    catch { }
                }

                _redoStack.Push(action);

                // Cấu hình Undo tiếp theo nếu còn phần tử trong Stack
                if (_undoStack.Count > 0)
                {
                    var nextUndo = _undoStack.Peek();
                    try
                    {
                        app.OnUndo($"Hoàn tác {nextUndo.ActionName}", GlobalUndoMacroName);
                    }
                    catch { }

                    try
                    {
                        app.OnKey("^z", GlobalUndoMacroName);
                        app.OnKey("^Z", GlobalUndoMacroName);
                    }
                    catch { }
                }

                // Cấu hình Redo (Ctrl+Y)
                try
                {
                    app.OnRepeat($"Làm lại {action.ActionName}", GlobalRedoMacroName);
                }
                catch { }

                try
                {
                    app.OnKey("^y", GlobalRedoMacroName);
                    app.OnKey("^Y", GlobalRedoMacroName);
                }
                catch { }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ExcelUndoHelper] GlobalUndo error: {ex.Message}");
            }
        }

        [ExcelCommand(Name = GlobalRedoMacroName)]
        public static void GlobalRedo()
        {
            ExcelApp? app = null;
            try
            {
                app = (ExcelApp)ExcelDnaUtil.Application;
                if (app != null)
                {
                    // Tạm giải phóng OnKey ^y để tránh vòng lặp
                    app.OnKey("^y");
                    app.OnKey("^Y");
                }

                if (_redoStack.Count == 0) return;

                var action = _redoStack.Pop();
                if (app == null) return;

                bool oldScreenUpdating = app.ScreenUpdating;
                XlCalculation oldCalculation = app.Calculation;
                bool oldEnableEvents = app.EnableEvents;

                try
                {
                    app.ScreenUpdating = false;
                    app.EnableEvents = false;
                    if (app.Calculation != XlCalculation.xlCalculationManual)
                    {
                        app.Calculation = XlCalculation.xlCalculationManual;
                    }

                    action.Redo(app);
                }
                finally
                {
                    try
                    {
                        app.Calculation = oldCalculation;
                        app.EnableEvents = oldEnableEvents;
                        app.ScreenUpdating = oldScreenUpdating;
                    }
                    catch { }
                }

                _undoStack.Push(action);

                // Cấu hình Redo tiếp theo nếu còn
                if (_redoStack.Count > 0)
                {
                    var nextRedo = _redoStack.Peek();
                    try
                    {
                        app.OnRepeat($"Làm lại {nextRedo.ActionName}", GlobalRedoMacroName);
                    }
                    catch { }

                    try
                    {
                        app.OnKey("^y", GlobalRedoMacroName);
                        app.OnKey("^Y", GlobalRedoMacroName);
                    }
                    catch { }
                }

                // Cấu hình Undo (Ctrl+Z)
                try
                {
                    app.OnUndo($"Hoàn tác {action.ActionName}", GlobalUndoMacroName);
                }
                catch { }

                try
                {
                    app.OnKey("^z", GlobalUndoMacroName);
                    app.OnKey("^Z", GlobalUndoMacroName);
                }
                catch { }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ExcelUndoHelper] GlobalRedo error: {ex.Message}");
            }
        }

        public static List<CellFormatSnapshot> CaptureRangeFormatting(Range range, int maxCells = 5000)
        {
            var list = new List<CellFormatSnapshot>();
            if (range == null) return list;

            long count = 0;
            try { count = range.Cells.CountLarge; }
            catch { count = range.Cells.Count; }

            if (count > maxCells) return list;

            var areas = range.Areas;
            int areaCount = areas.Count;

            for (int a = 1; a <= areaCount; a++)
            {
                Range area = areas[a];
                int rowCount = area.Rows.Count;
                int colCount = area.Columns.Count;
                int startRow = area.Row;
                int startCol = area.Column;

                for (int r = 1; r <= rowCount; r++)
                {
                    for (int c = 1; c <= colCount; c++)
                    {
                        Range? cell = null;
                        try
                        {
                            cell = area.Cells[r, c] as Range;
                            if (cell == null) continue;

                            var snapshot = new CellFormatSnapshot
                            {
                                Row = startRow + r - 1,
                                Column = startCol + c - 1
                            };

                            // Font
                            var font = cell.Font;
                            if (font != null)
                            {
                                try { snapshot.FontName = font.Name?.ToString(); } catch { }
                                try { snapshot.FontSize = font.Size; } catch { }
                                try { snapshot.FontBold = font.Bold; } catch { }
                                try { snapshot.FontItalic = font.Italic; } catch { }
                                try { snapshot.FontUnderline = font.Underline; } catch { }
                                try { snapshot.FontColor = font.Color; } catch { }
                                Marshal.ReleaseComObject(font);
                            }

                            // Interior
                            var interior = cell.Interior;
                            if (interior != null)
                            {
                                try { snapshot.InteriorColor = interior.Color; } catch { }
                                try { snapshot.InteriorColorIndex = interior.ColorIndex; } catch { }
                                Marshal.ReleaseComObject(interior);
                            }

                            // Alignment
                            try { snapshot.HorizontalAlignment = cell.HorizontalAlignment; } catch { }
                            try { snapshot.VerticalAlignment = cell.VerticalAlignment; } catch { }
                            try { snapshot.WrapText = cell.WrapText; } catch { }

                            // NumberFormat
                            try { snapshot.NumberFormat = cell.NumberFormat; } catch { }

                            // Borders
                            var borders = cell.Borders;
                            if (borders != null)
                            {
                                snapshot.Top = CaptureBorder(borders[XlBordersIndex.xlEdgeTop]);
                                snapshot.Bottom = CaptureBorder(borders[XlBordersIndex.xlEdgeBottom]);
                                snapshot.Left = CaptureBorder(borders[XlBordersIndex.xlEdgeLeft]);
                                snapshot.Right = CaptureBorder(borders[XlBordersIndex.xlEdgeRight]);
                                Marshal.ReleaseComObject(borders);
                            }

                            list.Add(snapshot);
                        }
                        finally
                        {
                            if (cell != null) Marshal.ReleaseComObject(cell);
                        }
                    }
                }
                Marshal.ReleaseComObject(area);
            }
            Marshal.ReleaseComObject(areas);

            return list;
        }

        public static void RestoreCellFormatting(_Worksheet ws, List<CellFormatSnapshot> snapshots)
        {
            if (ws == null || snapshots == null || snapshots.Count == 0) return;

            foreach (var item in snapshots)
            {
                Range? cell = null;
                try
                {
                    cell = ws.Cells[item.Row, item.Column] as Range;
                    if (cell == null) continue;

                    // Font
                    var font = cell.Font;
                    if (font != null)
                    {
                        try
                        {
                            if (!string.IsNullOrEmpty(item.FontName)) font.Name = item.FontName;
                            if (item.FontSize != null) font.Size = item.FontSize;
                            if (item.FontBold != null) font.Bold = item.FontBold;
                            if (item.FontItalic != null) font.Italic = item.FontItalic;
                            if (item.FontUnderline != null) font.Underline = item.FontUnderline;
                            if (item.FontColor != null) font.Color = item.FontColor;
                        }
                        catch { }
                        finally
                        {
                            Marshal.ReleaseComObject(font);
                        }
                    }

                    // Interior
                    var interior = cell.Interior;
                    if (interior != null)
                    {
                        try
                        {
                            if (item.InteriorColorIndex != null && Convert.ToInt32(item.InteriorColorIndex) == (int)XlColorIndex.xlColorIndexNone)
                            {
                                interior.ColorIndex = XlColorIndex.xlColorIndexNone;
                            }
                            else if (item.InteriorColor != null)
                            {
                                interior.Color = item.InteriorColor;
                            }
                        }
                        catch { }
                        finally
                        {
                            Marshal.ReleaseComObject(interior);
                        }
                    }

                    // Alignment
                    try
                    {
                        if (item.HorizontalAlignment != null) cell.HorizontalAlignment = item.HorizontalAlignment;
                        if (item.VerticalAlignment != null) cell.VerticalAlignment = item.VerticalAlignment;
                        if (item.WrapText != null) cell.WrapText = item.WrapText;
                    }
                    catch { }

                    // NumberFormat
                    try
                    {
                        if (item.NumberFormat != null) cell.NumberFormat = item.NumberFormat;
                    }
                    catch { }

                    // Borders
                    var borders = cell.Borders;
                    if (borders != null)
                    {
                        try
                        {
                            RestoreBorder(borders[XlBordersIndex.xlEdgeTop], item.Top);
                            RestoreBorder(borders[XlBordersIndex.xlEdgeBottom], item.Bottom);
                            RestoreBorder(borders[XlBordersIndex.xlEdgeLeft], item.Left);
                            RestoreBorder(borders[XlBordersIndex.xlEdgeRight], item.Right);
                        }
                        catch { }
                        finally
                        {
                            Marshal.ReleaseComObject(borders);
                        }
                    }
                }
                catch { }
                finally
                {
                    if (cell != null) Marshal.ReleaseComObject(cell);
                }
            }
        }

        private static BorderSnapshot? CaptureBorder(Border? border)
        {
            if (border == null) return null;
            try
            {
                return new BorderSnapshot
                {
                    LineStyle = border.LineStyle,
                    Weight = border.Weight,
                    Color = border.Color
                };
            }
            catch
            {
                return null;
            }
            finally
            {
                Marshal.ReleaseComObject(border);
            }
        }

        private static void RestoreBorder(Border? border, BorderSnapshot? snapshot)
        {
            if (border == null) return;
            try
            {
                if (snapshot == null || snapshot.LineStyle == null || Convert.ToInt32(snapshot.LineStyle) == (int)XlLineStyle.xlLineStyleNone)
                {
                    border.LineStyle = XlLineStyle.xlLineStyleNone;
                }
                else
                {
                    border.LineStyle = snapshot.LineStyle;
                    if (snapshot.Weight != null) border.Weight = snapshot.Weight;
                    if (snapshot.Color != null) border.Color = snapshot.Color;
                }
            }
            catch { }
            finally
            {
                Marshal.ReleaseComObject(border);
            }
        }
    }
}
