using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using Shape = Microsoft.Office.Interop.Excel.Shape;

namespace ExcelSupport.Services
{
    public enum GridRulerMode
    {
        BothRowAndCol, // Cả Dòng & Cột (Chữ thập)
        RowOnly,       // Chỉ Dòng
        ColOnly        // Chỉ Cột
    }

    public static class GridRulerService
    {
        public const string RowShapeName = "__ES_GridRuler_Row__";
        public const string ColShapeName = "__ES_GridRuler_Col__";

        public static bool IsEnabled { get; set; } = false;
        public static GridRulerMode CurrentMode { get; set; } = GridRulerMode.BothRowAndCol;
        public static float Transparency { get; set; } = 0.75f; // 75% trong suốt siêu dịu mắt
        public static string CurrentColorKey { get; set; } = "Yellow";

        // Bảng màu hỗ trợ
        public static readonly Dictionary<string, Color> ColorPalette = new Dictionary<string, Color>(StringComparer.OrdinalIgnoreCase)
        {
            { "Yellow", Color.FromArgb(253, 224, 71) },   // Vàng dịu (#FDE047)
            { "Sky", Color.FromArgb(56, 189, 248) },      // Xanh biển lơ (#38BDF8)
            { "Emerald", Color.FromArgb(74, 222, 128) },  // Xanh ngọc lục (#4ADE80)
            { "Orange", Color.FromArgb(251, 146, 60) },   // Cam đào (#FB923C)
            { "Purple", Color.FromArgb(192, 132, 252) },  // Tím Lavender (#C084FC)
            { "Pink", Color.FromArgb(244, 114, 182) },    // Hồng phấn (#F472B6)
            { "Gray", Color.FromArgb(148, 163, 184) }     // Xám thanh lịch (#94A3B8)
        };

        public static Color CurrentColor => ColorPalette.TryGetValue(CurrentColorKey, out var c) ? c : ColorPalette["Yellow"];

        #region Event Handlers

        public static void OnSheetSelectionChange(_Worksheet? ws, Range? target)
        {
            if (!IsEnabled || ws == null || target == null) return;

            try
            {
                UpdateRuler(ws, target);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"GridRuler SelectionChange error: {ex.Message}");
            }
        }

        public static void OnSheetActivate(_Worksheet? ws)
        {
            if (!IsEnabled || ws == null) return;

            try
            {
                var activeCell = ws.Application.ActiveCell;
                if (activeCell != null)
                {
                    UpdateRuler(ws, activeCell);
                    Marshal.ReleaseComObject(activeCell);
                }
            }
            catch { }
        }

        public static void OnWorkbookBeforeSave(Workbook? wb)
        {
            if (wb == null) return;
            try
            {
                // Dọn dẹp sạch sẽ shape trước khi lưu để không lưu lại rác vào file
                RemoveRulerFromAllWorksheets(wb);
            }
            catch { }
        }

        #endregion

        #region Core Ruler Update

        public static void UpdateRuler(_Worksheet ws, Range target)
        {
            if (ws == null || target == null || !IsEnabled) return;

            try
            {
                ws.Application.ScreenUpdating = false;

                int oleColor = ColorTranslator.ToOle(CurrentColor);

                // Tính toán vùng bao phủ của cửa sổ hiển thị
                double leftBound = 0;
                double topBound = 0;
                double fullWidth = 15000;
                double fullHeight = 15000;

                try
                {
                    var visRange = ws.Application.ActiveWindow?.VisibleRange;
                    if (visRange != null)
                    {
                        leftBound = Math.Max(0, (double)visRange.Left - 500);
                        topBound = Math.Max(0, (double)visRange.Top - 500);
                        fullWidth = Math.Max(10000, (double)visRange.Width + 2000);
                        fullHeight = Math.Max(10000, (double)visRange.Height + 2000);
                        Marshal.ReleaseComObject(visRange);
                    }
                }
                catch { }

                double targetTop = (double)target.Top;
                double targetHeight = (double)target.Height;
                double targetLeft = (double)target.Left;
                double targetWidth = (double)target.Width;

                // 1. Thước Ngang (Row Ruler)
                if (CurrentMode == GridRulerMode.BothRowAndCol || CurrentMode == GridRulerMode.RowOnly)
                {
                    var rowShape = GetOrCreateShape(ws, RowShapeName, leftBound, targetTop, fullWidth, targetHeight, oleColor);
                    if (rowShape != null)
                    {
                        rowShape.Top = (float)targetTop;
                        rowShape.Height = (float)targetHeight;
                        rowShape.Left = (float)leftBound;
                        rowShape.Width = (float)fullWidth;
                        rowShape.Visible = MsoTriState.msoTrue;
                        Marshal.ReleaseComObject(rowShape);
                    }
                }
                else
                {
                    HideShape(ws, RowShapeName);
                }

                // 2. Thước Dọc (Column Ruler)
                if (CurrentMode == GridRulerMode.BothRowAndCol || CurrentMode == GridRulerMode.ColOnly)
                {
                    var colShape = GetOrCreateShape(ws, ColShapeName, targetLeft, topBound, targetWidth, fullHeight, oleColor);
                    if (colShape != null)
                    {
                        colShape.Left = (float)targetLeft;
                        colShape.Width = (float)targetWidth;
                        colShape.Top = (float)topBound;
                        colShape.Height = (float)fullHeight;
                        colShape.Visible = MsoTriState.msoTrue;
                        Marshal.ReleaseComObject(colShape);
                    }
                }
                else
                {
                    HideShape(ws, ColShapeName);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"UpdateRuler error: {ex.Message}");
            }
            finally
            {
                try { ws.Application.ScreenUpdating = true; } catch { }
            }
        }

        private static Shape? GetOrCreateShape(_Worksheet ws, string shapeName, double left, double top, double width, double height, int oleColor)
        {
            Shape? foundShape = null;

            try
            {
                foreach (Shape s in ws.Shapes)
                {
                    if (string.Equals(s.Name, shapeName, StringComparison.OrdinalIgnoreCase))
                    {
                        foundShape = s;
                        break;
                    }
                    Marshal.ReleaseComObject(s);
                }
            }
            catch { }

            if (foundShape == null)
            {
                try
                {
                    foundShape = ws.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, (float)left, (float)top, (float)width, (float)height);
                    foundShape.Name = shapeName;
                    foundShape.Placement = XlPlacement.xlFreeFloating;
                    foundShape.Locked = false;
                    foundShape.Line.Visible = MsoTriState.msoFalse;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Create shape error: {ex.Message}");
                    return null;
                }
            }

            try
            {
                foundShape.Fill.Solid();
                foundShape.Fill.ForeColor.RGB = oleColor;
                foundShape.Fill.Transparency = Transparency;
            }
            catch { }

            return foundShape;
        }

        private static void HideShape(_Worksheet ws, string shapeName)
        {
            try
            {
                foreach (Shape s in ws.Shapes)
                {
                    if (string.Equals(s.Name, shapeName, StringComparison.OrdinalIgnoreCase))
                    {
                        s.Visible = MsoTriState.msoFalse;
                    }
                    Marshal.ReleaseComObject(s);
                }
            }
            catch { }
        }

        public static void RemoveRuler(_Worksheet ws)
        {
            if (ws == null) return;
            try
            {
                var toDelete = new List<Shape>();
                foreach (Shape s in ws.Shapes)
                {
                    if (string.Equals(s.Name, RowShapeName, StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(s.Name, ColShapeName, StringComparison.OrdinalIgnoreCase))
                    {
                        toDelete.Add(s);
                    }
                    else
                    {
                        Marshal.ReleaseComObject(s);
                    }
                }

                foreach (var s in toDelete)
                {
                    try { s.Delete(); } catch { }
                    Marshal.ReleaseComObject(s);
                }
            }
            catch { }
        }

        public static void RemoveRulerFromAllWorksheets(Workbook wb)
        {
            if (wb == null) return;
            try
            {
                foreach (_Worksheet s in wb.Worksheets)
                {
                    RemoveRuler(s);
                    Marshal.ReleaseComObject(s);
                }
            }
            catch { }
        }

        #endregion

        #region Configuration & Toggles

        public static void Toggle(ExcelApp? app)
        {
            IsEnabled = !IsEnabled;

            if (app == null) return;

            try
            {
                var activeWs = app.ActiveSheet as _Worksheet;
                if (activeWs != null)
                {
                    if (IsEnabled)
                    {
                        var cell = app.ActiveCell;
                        if (cell != null)
                        {
                            UpdateRuler(activeWs, cell);
                            Marshal.ReleaseComObject(cell);
                        }
                    }
                    else
                    {
                        RemoveRuler(activeWs);
                    }
                    Marshal.ReleaseComObject(activeWs);
                }
            }
            catch { }
        }

        public static void SetColor(string colorKey, ExcelApp? app)
        {
            CurrentColorKey = colorKey;

            if (!IsEnabled || app == null) return;

            try
            {
                var activeWs = app.ActiveSheet as _Worksheet;
                if (activeWs != null)
                {
                    var cell = app.ActiveCell;
                    if (cell != null)
                    {
                        UpdateRuler(activeWs, cell);
                        Marshal.ReleaseComObject(cell);
                    }
                    Marshal.ReleaseComObject(activeWs);
                }
            }
            catch { }
        }

        public static void SetMode(GridRulerMode mode, ExcelApp? app)
        {
            CurrentMode = mode;

            if (!IsEnabled || app == null) return;

            try
            {
                var activeWs = app.ActiveSheet as _Worksheet;
                if (activeWs != null)
                {
                    var cell = app.ActiveCell;
                    if (cell != null)
                    {
                        UpdateRuler(activeWs, cell);
                        Marshal.ReleaseComObject(cell);
                    }
                    Marshal.ReleaseComObject(activeWs);
                }
            }
            catch { }
        }

        #endregion
    }
}
