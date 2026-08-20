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

    public class RulerQuickStats
    {
        public string CellAddress { get; set; } = string.Empty;
        public int RowIndex { get; set; }
        public int ColIndex { get; set; }
        public string ColLetter { get; set; } = string.Empty;

        // Row Stats
        public double RowSum { get; set; }
        public double RowAvg { get; set; }
        public int RowNumericCount { get; set; }
        public int RowNonEmptyCount { get; set; }
        public double RowMax { get; set; }
        public double RowMin { get; set; }

        // Col Stats
        public double ColSum { get; set; }
        public double ColAvg { get; set; }
        public int ColNumericCount { get; set; }
        public int ColNonEmptyCount { get; set; }
        public double ColMax { get; set; }
        public double ColMin { get; set; }

        public string RowStatsSummary => RowNumericCount > 0
            ? $"Dòng {RowIndex}: Tổng = {RowSum:N0} | TB = {RowAvg:N2} | Đếm = {RowNonEmptyCount} | Max = {RowMax:N0} | Min = {RowMin:N0}"
            : $"Dòng {RowIndex}: {RowNonEmptyCount} ô có dữ liệu";

        public string ColStatsSummary => ColNumericCount > 0
            ? $"Cột {ColLetter}: Tổng = {ColSum:N0} | TB = {ColAvg:N2} | Đếm = {ColNonEmptyCount} | Max = {ColMax:N0} | Min = {ColMin:N0}"
            : $"Cột {ColLetter}: {ColNonEmptyCount} ô có dữ liệu";
    }

    public static class GridRulerService
    {
        public const string RowShapeName = "__ES_GridRuler_Row__";
        public const string ColShapeName = "__ES_GridRuler_Col__";

        public static bool IsEnabled { get; set; } = false;
        public static bool ShowQuickStatsInStatusBar { get; set; } = true;
        public static bool ShowFloatingHud { get; set; } = true;
        public static GridRulerMode CurrentMode { get; set; } = GridRulerMode.BothRowAndCol;
        public static float Transparency { get; set; } = 0.75f; // 75% trong suốt siêu dịu mắt
        public static string CurrentColorKey { get; set; } = "Yellow";

        public static event Action<RulerQuickStats>? QuickStatsUpdated;
        public static RulerQuickStats? LastQuickStats { get; private set; }

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

                // 3. Tính toán Quick Stats siêu tốc cho dòng & cột đang chọn
                var stats = CalculateQuickStats(ws, target);
                LastQuickStats = stats;
                QuickStatsUpdated?.Invoke(stats);

                // Cập nhật lên Bảng Thống Kê Nổi (Floating HUD)
                Views.RulerHudWindow.UpdateCurrentStats(stats);
                if (ShowFloatingHud && !Views.RulerHudWindow.IsHudVisible)
                {
                    Views.RulerHudWindow.ShowHud(AddInEvents.MainViewModel?.IsDarkTheme ?? true);
                }

                if (ShowQuickStatsInStatusBar)
                {
                    try
                    {
                        ws.Application.StatusBar = $"📍 [Ruler Plus] {stats.RowStatsSummary}  ||  {stats.ColStatsSummary}";
                    }
                    catch { }
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

        public static RulerQuickStats CalculateQuickStats(_Worksheet ws, Range target)
        {
            var stats = new RulerQuickStats
            {
                CellAddress = target.Address[false, false],
                RowIndex = target.Row,
                ColIndex = target.Column,
                ColLetter = GetColumnLetter(target.Column)
            };

            Range? usedRange = null;
            try
            {
                usedRange = ws.UsedRange;
                if (usedRange == null || usedRange.Rows.Count == 0) return stats;

                int startRow = usedRange.Row;
                int startCol = usedRange.Column;
                int numRows = usedRange.Rows.Count;
                int numCols = usedRange.Columns.Count;

                // 1. TÍNH TOÁN THEO DÒNG (Chỉ tính các ô hiển thị - Bỏ qua cột bị ẩn)
                Range? rowRange = null;
                Range? rowVisible = null;
                try
                {
                    rowRange = ws.Range[ws.Cells[stats.RowIndex, startCol], ws.Cells[stats.RowIndex, startCol + numCols - 1]];
                    try
                    {
                        rowVisible = rowRange.SpecialCells(XlCellType.xlCellTypeVisible);
                    }
                    catch
                    {
                        rowVisible = rowRange;
                    }

                    if (rowVisible != null)
                    {
                        int rowNonEmpty = 0;
                        int rowNumeric = 0;
                        double rowSum = 0;
                        double rowMax = 0;
                        double rowMin = 0;
                        bool firstNum = true;

                        foreach (Range area in rowVisible.Areas)
                        {
                            object? rawVal = area.Value2;
                            if (rawVal is object[,] val2D)
                            {
                                int areaRows = val2D.GetLength(0);
                                int areaCols = val2D.GetLength(1);
                                for (int r = 1; r <= areaRows; r++)
                                {
                                    for (int c = 1; c <= areaCols; c++)
                                    {
                                        ProcessStatValue(val2D[r, c], ref rowNonEmpty, ref rowNumeric, ref rowSum, ref rowMax, ref rowMin, ref firstNum);
                                    }
                                }
                            }
                            else if (rawVal != null)
                            {
                                ProcessStatValue(rawVal, ref rowNonEmpty, ref rowNumeric, ref rowSum, ref rowMax, ref rowMin, ref firstNum);
                            }
                            Marshal.ReleaseComObject(area);
                        }

                        stats.RowNonEmptyCount = rowNonEmpty;
                        stats.RowNumericCount = rowNumeric;
                        stats.RowSum = rowSum;
                        stats.RowMax = rowMax;
                        stats.RowMin = rowMin;

                        if (stats.RowNumericCount > 0)
                        {
                            stats.RowAvg = stats.RowSum / stats.RowNumericCount;
                        }
                    }
                }
                finally
                {
                    if (rowVisible != null && rowVisible != rowRange) Marshal.ReleaseComObject(rowVisible);
                    if (rowRange != null) Marshal.ReleaseComObject(rowRange);
                }

                // 2. TÍNH TOÁN THEO CỘT (Chỉ tính các ô hiển thị - Bỏ qua các dòng bị ẩn hoặc bị lọc bởi AutoFilter)
                Range? colRange = null;
                Range? colVisible = null;
                try
                {
                    colRange = ws.Range[ws.Cells[startRow, stats.ColIndex], ws.Cells[startRow + numRows - 1, stats.ColIndex]];
                    try
                    {
                        colVisible = colRange.SpecialCells(XlCellType.xlCellTypeVisible);
                    }
                    catch
                    {
                        colVisible = colRange;
                    }

                    if (colVisible != null)
                    {
                        int colNonEmpty = 0;
                        int colNumeric = 0;
                        double colSum = 0;
                        double colMax = 0;
                        double colMin = 0;
                        bool firstNum = true;

                        foreach (Range area in colVisible.Areas)
                        {
                            object? rawVal = area.Value2;
                            if (rawVal is object[,] val2D)
                            {
                                int areaRows = val2D.GetLength(0);
                                int areaCols = val2D.GetLength(1);
                                for (int r = 1; r <= areaRows; r++)
                                {
                                    for (int c = 1; c <= areaCols; c++)
                                    {
                                        ProcessStatValue(val2D[r, c], ref colNonEmpty, ref colNumeric, ref colSum, ref colMax, ref colMin, ref firstNum);
                                    }
                                }
                            }
                            else if (rawVal != null)
                            {
                                ProcessStatValue(rawVal, ref colNonEmpty, ref colNumeric, ref colSum, ref colMax, ref colMin, ref firstNum);
                            }
                            Marshal.ReleaseComObject(area);
                        }

                        stats.ColNonEmptyCount = colNonEmpty;
                        stats.ColNumericCount = colNumeric;
                        stats.ColSum = colSum;
                        stats.ColMax = colMax;
                        stats.ColMin = colMin;

                        if (stats.ColNumericCount > 0)
                        {
                            stats.ColAvg = stats.ColSum / stats.ColNumericCount;
                        }
                    }
                }
                finally
                {
                    if (colVisible != null && colVisible != colRange) Marshal.ReleaseComObject(colVisible);
                    if (colRange != null) Marshal.ReleaseComObject(colRange);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"CalculateQuickStats error: {ex.Message}");
            }
            finally
            {
                if (usedRange != null) Marshal.ReleaseComObject(usedRange);
            }

            return stats;
        }

        private static void ProcessStatValue(object? cellObj, ref int nonEmptyCount, ref int numericCount, ref double sum, ref double max, ref double min, ref bool firstNum)
        {
            if (cellObj == null) return;
            string str = cellObj.ToString()?.Trim() ?? string.Empty;
            if (string.IsNullOrEmpty(str)) return;

            nonEmptyCount++;

            if (cellObj is double d)
            {
                numericCount++;
                sum += d;
                if (firstNum) { max = d; min = d; firstNum = false; }
                else { if (d > max) max = d; if (d < min) min = d; }
            }
            else if (cellObj is int i)
            {
                numericCount++;
                sum += i;
                if (firstNum) { max = i; min = i; firstNum = false; }
                else { if (i > max) max = i; if (i < min) min = i; }
            }
            else if (cellObj is long l)
            {
                numericCount++;
                sum += l;
                if (firstNum) { max = l; min = l; firstNum = false; }
                else { if (l > max) max = l; if (l < min) min = l; }
            }
            else if (cellObj is float f)
            {
                numericCount++;
                sum += f;
                if (firstNum) { max = f; min = f; firstNum = false; }
                else { if (f > max) max = f; if (f < min) min = f; }
            }
            else if (cellObj is decimal dec)
            {
                double num = (double)dec;
                numericCount++;
                sum += num;
                if (firstNum) { max = num; min = num; firstNum = false; }
                else { if (num > max) max = num; if (num < min) min = num; }
            }
            else if (double.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.CurrentCulture, out double parsed) ||
                     double.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out parsed))
            {
                numericCount++;
                sum += parsed;
                if (firstNum) { max = parsed; min = parsed; firstNum = false; }
                else { if (parsed > max) max = parsed; if (parsed < min) min = parsed; }
            }
        }

        private static string GetColumnLetter(int colIndex)
        {
            string colLetter = string.Empty;
            while (colIndex > 0)
            {
                int modulo = (colIndex - 1) % 26;
                colLetter = Convert.ToChar('A' + modulo) + colLetter;
                colIndex = (colIndex - modulo) / 26;
            }
            return colLetter;
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

                try { ws.Application.StatusBar = false; } catch { }
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
                        if (ShowFloatingHud)
                        {
                            Views.RulerHudWindow.ShowHud(AddInEvents.MainViewModel?.IsDarkTheme ?? true);
                        }
                    }
                    else
                    {
                        RemoveRuler(activeWs);
                        Views.RulerHudWindow.HideHud();
                        try { app.StatusBar = false; } catch { }
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
