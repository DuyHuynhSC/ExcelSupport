using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using ExcelSupport.Models;
using Microsoft.Office.Interop.Excel;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;

namespace ExcelSupport.Services
{
    public static class SheetSnapshotService
    {
        private static readonly List<SheetSnapshotItem> Snapshots = new List<SheetSnapshotItem>();
        private static readonly object SyncLock = new object();
        private const int MaxSnapshotCount = 30;

        public static IReadOnlyList<SheetSnapshotItem> GetSnapshots()
        {
            lock (SyncLock)
            {
                return Snapshots.OrderByDescending(s => s.Timestamp).ToList();
            }
        }

        public static SheetSnapshotItem? TakeSnapshot(ExcelApp app, string? description = null, bool isAuto = false, Worksheet? targetSheet = null)
        {
            try
            {
                var ws = targetSheet ?? app.ActiveSheet as Worksheet;
                if (ws == null) return null;

                var wb = ws.Parent as Workbook;
                var usedRange = ws.UsedRange;
                if (usedRange == null) return null;

                int startRow = usedRange.Row;
                int startCol = usedRange.Column;
                int rowCount = usedRange.Rows.Count;
                int colCount = usedRange.Columns.Count;

                if (rowCount <= 0 || colCount <= 0) return null;

                var snapshot = new SheetSnapshotItem
                {
                    WorkbookName = wb?.Name ?? "Workbook",
                    SheetName = ws.Name,
                    Description = string.IsNullOrWhiteSpace(description)
                        ? (isAuto ? "Tự động sao lưu trước tác vụ" : "Sao lưu thủ công")
                        : description.Trim(),
                    StartRow = startRow,
                    StartColumn = startCol,
                    RowCount = rowCount,
                    ColumnCount = colCount,
                    IsAutoSnapshot = isAuto,
                    Timestamp = DateTime.Now
                };

                // Bulk read values, formulas, and number formats
                object[,] valArray = new object[rowCount, colCount];
                object[,] formArray = new object[rowCount, colCount];
                object[,] numFormatArray = new object[rowCount, colCount];

                if (rowCount == 1 && colCount == 1)
                {
                    valArray[0, 0] = usedRange.Value2;
                    formArray[0, 0] = usedRange.Formula;
                    numFormatArray[0, 0] = usedRange.NumberFormat;
                }
                else
                {
                    if (usedRange.Value2 is object[,] rawVals)
                    {
                        valArray = ConvertToOneBasedToZeroBased(rawVals, rowCount, colCount);
                    }
                    if (usedRange.Formula is object[,] rawForms)
                    {
                        formArray = ConvertToOneBasedToZeroBased(rawForms, rowCount, colCount);
                    }
                    if (usedRange.NumberFormat is object[,] rawFormats)
                    {
                        numFormatArray = ConvertToOneBasedToZeroBased(rawFormats, rowCount, colCount);
                    }
                }

                snapshot.Values = valArray;
                snapshot.Formulas = formArray;
                snapshot.NumberFormats = numFormatArray;

                // Capture column widths
                for (int c = 1; c <= Math.Min(colCount, 100); c++)
                {
                    try
                    {
                        var colRange = ws.Columns[startCol + c - 1] as Range;
                        if (colRange != null && colRange.ColumnWidth is double w)
                        {
                            snapshot.ColumnWidths[startCol + c - 1] = w;
                        }
                    }
                    catch { }
                }

                lock (SyncLock)
                {
                    Snapshots.Insert(0, snapshot);
                    while (Snapshots.Count > MaxSnapshotCount)
                    {
                        Snapshots.RemoveAt(Snapshots.Count - 1);
                    }
                }

                return snapshot;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[SheetSnapshotService] TakeSnapshot error: {ex.Message}");
                return null;
            }
        }

        public static bool RestoreSnapshot(ExcelApp app, SheetSnapshotItem snapshot, bool restoreToNewSheet = false)
        {
            try
            {
                var wb = app.ActiveWorkbook;
                if (wb == null) return false;

                Worksheet? targetSheet = null;

                if (restoreToNewSheet)
                {
                    // Create new sheet with unique name
                    string baseName = $"Restored_{snapshot.SheetName}";
                    if (baseName.Length > 22) baseName = baseName.Substring(0, 22);
                    string sheetName = baseName;
                    int suffix = 1;

                    while (SheetExists(wb, sheetName))
                    {
                        sheetName = $"{baseName}_{suffix++}";
                    }

                    targetSheet = wb.Worksheets.Add() as Worksheet;
                    if (targetSheet != null)
                    {
                        targetSheet.Name = sheetName;
                    }
                }
                else
                {
                    // Find original sheet or use active
                    try
                    {
                        targetSheet = wb.Worksheets[snapshot.SheetName] as Worksheet;
                    }
                    catch
                    {
                        targetSheet = app.ActiveSheet as Worksheet;
                    }
                }

                if (targetSheet == null) return false;

                // Turn off screen updating for high performance restore
                app.ScreenUpdating = false;
                app.Calculation = XlCalculation.xlCalculationManual;
                app.EnableEvents = false;

                try
                {
                    int rows = snapshot.RowCount;
                    int cols = snapshot.ColumnCount;

                    var cell1 = targetSheet.Cells[snapshot.StartRow, snapshot.StartColumn] as Range;
                    var cell2 = targetSheet.Cells[snapshot.StartRow + rows - 1, snapshot.StartColumn + cols - 1] as Range;
                    var destRange = targetSheet.Range[cell1, cell2];

                    // Convert 0-based to 1-based arrays for Excel COM bulk write
                    object[,] writeVals = ConvertTo1Based(snapshot.Values, rows, cols);
                    object[,] writeForms = ConvertTo1Based(snapshot.Formulas, rows, cols);

                    // If formulas exist, write formulas; otherwise write values
                    bool hasFormulas = false;
                    for (int r = 0; r < rows; r++)
                    {
                        for (int c = 0; c < cols; c++)
                        {
                            string? formStr = snapshot.Formulas[r, c]?.ToString();
                            if (!string.IsNullOrWhiteSpace(formStr) && formStr.StartsWith("="))
                            {
                                hasFormulas = true;
                                break;
                            }
                        }
                        if (hasFormulas) break;
                    }

                    if (hasFormulas)
                    {
                        destRange.Formula = writeForms;
                    }
                    else
                    {
                        destRange.Value2 = writeVals;
                    }

                    // Restore Column Widths
                    foreach (var kvp in snapshot.ColumnWidths)
                    {
                        try
                        {
                            var colRange = targetSheet.Columns[kvp.Key] as Range;
                            if (colRange != null) colRange.ColumnWidth = kvp.Value;
                        }
                        catch { }
                    }

                    targetSheet.Activate();
                    return true;
                }
                finally
                {
                    app.Calculation = XlCalculation.xlCalculationAutomatic;
                    app.EnableEvents = true;
                    app.ScreenUpdating = true;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[SheetSnapshotService] RestoreSnapshot error: {ex.Message}");
                return false;
            }
        }

        public static List<SnapshotCellDiff> CompareWithCurrent(ExcelApp app, SheetSnapshotItem snapshot)
        {
            var diffs = new List<SnapshotCellDiff>();

            try
            {
                var ws = app.ActiveSheet as Worksheet;
                if (ws == null) return diffs;

                int rows = snapshot.RowCount;
                int cols = snapshot.ColumnCount;

                var cell1 = ws.Cells[snapshot.StartRow, snapshot.StartColumn] as Range;
                var cell2 = ws.Cells[snapshot.StartRow + rows - 1, snapshot.StartColumn + cols - 1] as Range;
                var currentRange = ws.Range[cell1, cell2];

                object[,] curVals = new object[rows, cols];
                object[,] curForms = new object[rows, cols];

                if (rows == 1 && cols == 1)
                {
                    curVals[0, 0] = currentRange.Value2;
                    curForms[0, 0] = currentRange.Formula;
                }
                else
                {
                    if (currentRange.Value2 is object[,] rawV)
                    {
                        curVals = ConvertToOneBasedToZeroBased(rawV, rows, cols);
                    }
                    if (currentRange.Formula is object[,] rawF)
                    {
                        curForms = ConvertToOneBasedToZeroBased(rawF, rows, cols);
                    }
                }

                for (int r = 0; r < rows; r++)
                {
                    for (int c = 0; c < cols; c++)
                    {
                        string oldVal = snapshot.Values[r, c]?.ToString() ?? string.Empty;
                        string newVal = curVals[r, c]?.ToString() ?? string.Empty;

                        string oldForm = snapshot.Formulas[r, c]?.ToString() ?? string.Empty;
                        string newForm = curForms[r, c]?.ToString() ?? string.Empty;

                        if (oldVal != newVal || oldForm != newForm)
                        {
                            int actualRow = snapshot.StartRow + r;
                            int actualCol = snapshot.StartColumn + c;

                            var diff = new SnapshotCellDiff
                            {
                                Row = actualRow,
                                Column = actualCol,
                                CellAddress = GetExcelAddress(actualRow, actualCol),
                                SnapshotValue = oldVal,
                                CurrentValue = newVal,
                                SnapshotFormula = oldForm,
                                CurrentFormula = newForm,
                                DiffType = oldForm != newForm ? SnapshotDiffType.FormulaChanged : SnapshotDiffType.ValueChanged
                            };

                            diffs.Add(diff);
                            if (diffs.Count >= 500) return diffs; // Limit diff preview to 500 cells
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[SheetSnapshotService] Compare error: {ex.Message}");
            }

            return diffs;
        }

        public static void DeleteSnapshot(string snapshotId)
        {
            lock (SyncLock)
            {
                Snapshots.RemoveAll(s => s.Id == snapshotId);
            }
        }

        public static void ClearAllSnapshots()
        {
            lock (SyncLock)
            {
                Snapshots.Clear();
            }
        }

        #region Array Helpers

        private static object[,] ConvertToOneBasedToZeroBased(object[,] raw, int rows, int cols)
        {
            var res = new object[rows, cols];
            int rLower = raw.GetLowerBound(0);
            int cLower = raw.GetLowerBound(1);

            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    res[r, c] = raw[rLower + r, cLower + c];
                }
            }
            return res;
        }

        private static object[,] ConvertTo1Based(object[,] raw0, int rows, int cols)
        {
            // Excel COM interop expects 1-based multi-dimensional array
            Array arr = Array.CreateInstance(typeof(object), new int[] { rows, cols }, new int[] { 1, 1 });
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    arr.SetValue(raw0[r, c], r + 1, c + 1);
                }
            }
            return (object[,])arr;
        }

        private static bool SheetExists(Workbook wb, string name)
        {
            return wb.Worksheets.Cast<Worksheet>().Any(ws => string.Equals(ws.Name, name, StringComparison.OrdinalIgnoreCase));
        }

        private static string GetExcelAddress(int row, int col)
        {
            string colLetter = string.Empty;
            while (col > 0)
            {
                int modulo = (col - 1) % 26;
                colLetter = Convert.ToChar('A' + modulo) + colLetter;
                col = (col - modulo) / 26;
            }
            return $"{colLetter}{row}";
        }

        #endregion
    }
}
