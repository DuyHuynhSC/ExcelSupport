using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Office.Interop.Excel;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;

namespace ExcelSupport.Services
{
    #region Enums & Data Models

    public enum BlankCleanupScope
    {
        ActiveSheet,
        AllSheetsInActiveWb,
        AllOpenWorkbooks
    }

    public enum BlankCleanupTarget
    {
        EntirelyBlankRows,
        BlankRowsInKeyColumn,
        EntirelyBlankColumns
    }

    public enum BlankCleanupAction
    {
        Delete,
        Hide,
        Highlight
    }

    public enum SafeMergeDirection
    {
        AcrossRows,    // Gộp theo từng dòng (Row by row)
        DownColumns,   // Gộp theo từng cột (Column by column)
        AllToOneCell   // Gộp toàn bộ thành 1 ô duy nhất
    }

    public enum MergeSeparatorType
    {
        Space,
        Comma,
        Semicolon,
        NewLine,
        Pipe,
        Custom
    }

    public class SafeMergeOptions
    {
        public SafeMergeDirection Direction { get; set; } = SafeMergeDirection.AcrossRows;
        public MergeSeparatorType SeparatorType { get; set; } = MergeSeparatorType.Comma;
        public string CustomSeparator { get; set; } = ", ";
        public bool IgnoreBlankCells { get; set; } = true;
        public bool TrimSpaces { get; set; } = true;

        public string GetActualSeparator()
        {
            switch (SeparatorType)
            {
                case MergeSeparatorType.Space: return " ";
                case MergeSeparatorType.Comma: return ", ";
                case MergeSeparatorType.Semicolon: return "; ";
                case MergeSeparatorType.NewLine: return "\n";
                case MergeSeparatorType.Pipe: return " | ";
                case MergeSeparatorType.Custom: return CustomSeparator ?? "";
                default: return " ";
            }
        }
    }

    public class CombineSheetsOptions
    {
        public List<string> SelectedSheetNames { get; set; } = new List<string>();
        public bool HasHeaderRow { get; set; } = true;
        public int HeaderRowCount { get; set; } = 1;
        public bool AddSourceColumn { get; set; } = true;
        public string SourceColumnHeader { get; set; } = "Tên Sheet Nguồn";
        public bool SkipBlankRows { get; set; } = true;
    }

    public class SheetItemInfo
    {
        public string SheetName { get; set; } = string.Empty;
        public string WorkbookName { get; set; } = string.Empty;
        public int TotalRows { get; set; }
        public int TotalCols { get; set; }
        public bool IsSelected { get; set; } = true;

        public string DisplayText => $"{SheetName} ({TotalRows:N0} dòng x {TotalCols:N0} cột)";
    }

    #endregion

    public static class BatchCleanerAndMergeService
    {
        #region 1. XÓA DÒNG & CỘT TRỐNG HÀNG LOẠT

        /// <summary>
        /// Xóa hoặc ẩn các dòng/cột trống trên 1 Worksheet
        /// </summary>
        public static (int ProcessedCount, int TotalRowsOrCols) ProcessBlankInSheet(
            _Worksheet ws,
            BlankCleanupTarget target,
            BlankCleanupAction action,
            int keyColumnIndex = 1,
            System.Drawing.Color? highlightColor = null)
        {
            if (ws == null) return (0, 0);

            Range? usedRange = null;
            try
            {
                usedRange = ws.UsedRange;
                if (usedRange == null || usedRange.Rows.Count == 0) return (0, 0);

                int startRow = usedRange.Row;
                int totalRows = usedRange.Rows.Count;
                int startCol = usedRange.Column;
                int totalCols = usedRange.Columns.Count;

                object[,]? values2D = null;
                if (totalRows == 1 && totalCols == 1)
                {
                    values2D = new object[2, 2];
                    values2D[1, 1] = usedRange.Value2;
                }
                else
                {
                    values2D = (object[,])usedRange.Value2;
                }

                if (values2D == null) return (0, 0);

                ws.Application.ScreenUpdating = false;

                if (target == BlankCleanupTarget.EntirelyBlankColumns)
                {
                    // Xử lý Cột trống
                    var blankCols = new List<int>();

                    for (int c = 1; c <= totalCols; c++)
                    {
                        bool isColBlank = true;
                        for (int r = 1; r <= totalRows; r++)
                        {
                            object? v = values2D[r, c];
                            if (v != null && !string.IsNullOrWhiteSpace(v.ToString()))
                            {
                                isColBlank = false;
                                break;
                            }
                        }
                        if (isColBlank)
                        {
                            blankCols.Add(startCol + c - 1);
                        }
                    }

                    if (blankCols.Count == 0) return (0, totalCols);

                    // Thực thi hành động từ phải qua trái (để không làm lệch index)
                    blankCols.Sort();
                    blankCols.Reverse();

                    int colorOle = ColorTranslator.ToOle(highlightColor ?? System.Drawing.Color.FromArgb(254, 202, 202)); // Đỏ nhạt

                    foreach (int colIdx in blankCols)
                    {
                        Range? colRange = null;
                        try
                        {
                            colRange = ws.Columns[colIdx] as Range;
                            if (colRange != null)
                            {
                                if (action == BlankCleanupAction.Delete)
                                {
                                    colRange.Delete(XlDeleteShiftDirection.xlShiftToLeft);
                                }
                                else if (action == BlankCleanupAction.Hide)
                                {
                                    colRange.Hidden = true;
                                }
                                else if (action == BlankCleanupAction.Highlight)
                                {
                                    colRange.Interior.Color = colorOle;
                                }
                            }
                        }
                        catch { }
                        finally
                        {
                            if (colRange != null) Marshal.ReleaseComObject(colRange);
                        }
                    }

                    return (blankCols.Count, totalCols);
                }
                else
                {
                    // Xử lý Dòng trống
                    var blankRows = new List<int>();

                    for (int r = 1; r <= totalRows; r++)
                    {
                        bool isRowBlank = false;

                        if (target == BlankCleanupTarget.EntirelyBlankRows)
                        {
                            isRowBlank = true;
                            for (int c = 1; c <= totalCols; c++)
                            {
                                object? v = values2D[r, c];
                                if (v != null && !string.IsNullOrWhiteSpace(v.ToString()))
                                {
                                    isRowBlank = false;
                                    break;
                                }
                            }
                        }
                        else // BlankRowsInKeyColumn
                        {
                            int keyColOffset = keyColumnIndex - startCol + 1;
                            if (keyColOffset >= 1 && keyColOffset <= totalCols)
                            {
                                object? v = values2D[r, keyColOffset];
                                isRowBlank = (v == null || string.IsNullOrWhiteSpace(v.ToString()));
                            }
                        }

                        if (isRowBlank)
                        {
                            blankRows.Add(startRow + r - 1);
                        }
                    }

                    if (blankRows.Count == 0) return (0, totalRows);

                    int colorOle = ColorTranslator.ToOle(highlightColor ?? System.Drawing.Color.FromArgb(254, 202, 202));

                    if (action == BlankCleanupAction.Delete)
                    {
                        // Gom thành các khối liên tiếp để xóa từ dưới lên siêu tốc
                        var blocks = GroupContiguousNumbers(blankRows);
                        blocks.Reverse();

                        foreach (var (fromRow, toRow) in blocks)
                        {
                            Range? rowRange = null;
                            try
                            {
                                rowRange = ws.Range[$"A{fromRow}:A{toRow}"];
                                if (rowRange != null)
                                {
                                    rowRange.EntireRow.Delete(XlDeleteShiftDirection.xlShiftUp);
                                }
                            }
                            catch { }
                            finally
                            {
                                if (rowRange != null) Marshal.ReleaseComObject(rowRange);
                            }
                        }
                    }
                    else if (action == BlankCleanupAction.Hide)
                    {
                        var blocks = GroupContiguousNumbers(blankRows);
                        foreach (var (fromRow, toRow) in blocks)
                        {
                            Range? rowRange = null;
                            try
                            {
                                rowRange = ws.Range[$"A{fromRow}:A{toRow}"];
                                if (rowRange != null)
                                {
                                    rowRange.EntireRow.Hidden = true;
                                }
                            }
                            catch { }
                            finally
                            {
                                if (rowRange != null) Marshal.ReleaseComObject(rowRange);
                            }
                        }
                    }
                    else if (action == BlankCleanupAction.Highlight)
                    {
                        foreach (int r in blankRows)
                        {
                            Range? rowRange = null;
                            try
                            {
                                rowRange = ws.Range[ws.Cells[r, startCol], ws.Cells[r, startCol + totalCols - 1]];
                                if (rowRange != null)
                                {
                                    rowRange.Interior.Color = colorOle;
                                }
                            }
                            catch { }
                            finally
                            {
                                if (rowRange != null) Marshal.ReleaseComObject(rowRange);
                            }
                        }
                    }

                    return (blankRows.Count, totalRows);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ProcessBlankInSheet error: {ex.Message}");
                return (0, 0);
            }
            finally
            {
                try { ws.Application.ScreenUpdating = true; } catch { }
                if (usedRange != null) Marshal.ReleaseComObject(usedRange);
            }
        }

        private static List<(int From, int To)> GroupContiguousNumbers(List<int> numbers)
        {
            var result = new List<(int From, int To)>();
            if (numbers.Count == 0) return result;

            numbers.Sort();
            int start = numbers[0];
            int prev = numbers[0];

            for (int i = 1; i < numbers.Count; i++)
            {
                if (numbers[i] == prev + 1)
                {
                    prev = numbers[i];
                }
                else
                {
                    result.Add((start, prev));
                    start = numbers[i];
                    prev = numbers[i];
                }
            }
            result.Add((start, prev));

            return result;
        }

        #endregion

        #region 2. GỘP Ô BẢO TOÀN DỮ LIỆU (SAFE MERGE CELLS)

        /// <summary>
        /// Gộp vùng ô đang chọn mà không làm mất dữ liệu của bất kỳ ô nào
        /// </summary>
        public static (bool Success, string Message, int MergedCount) MergeSelectedCellsSafely(_Worksheet ws, Range selection, SafeMergeOptions options)
        {
            if (ws == null || selection == null || options == null)
            {
                return (false, "Vui lòng chọn một vùng ô (nhiều hơn 1 ô) để thực hiện gộp.", 0);
            }

            try
            {
                int rowCount = selection.Rows.Count;
                int colCount = selection.Columns.Count;

                if (rowCount <= 1 && colCount <= 1)
                {
                    return (false, "Vùng chọn chỉ có 1 ô duy nhất. Vui lòng chọn ít nhất 2 ô trở lên để gộp.", 0);
                }

                ws.Application.ScreenUpdating = false;
                ws.Application.DisplayAlerts = false; // Tắt popup cảnh báo mất dữ liệu của Excel

                string separator = options.GetActualSeparator();
                int mergedGroups = 0;

                object[,] rawValues;
                if (rowCount == 1 && colCount == 1)
                {
                    rawValues = new object[2, 2];
                    rawValues[1, 1] = selection.Value2;
                }
                else
                {
                    rawValues = (object[,])selection.Value2;
                }

                if (options.Direction == SafeMergeDirection.AcrossRows)
                {
                    // Gộp theo từng dòng (Row by row)
                    for (int r = 1; r <= rowCount; r++)
                    {
                        var texts = new List<string>();
                        for (int c = 1; c <= colCount; c++)
                        {
                            object? val = rawValues[r, c];
                            string s = val?.ToString() ?? "";
                            if (options.TrimSpaces) s = s.Trim();

                            if (!options.IgnoreBlankCells || !string.IsNullOrEmpty(s))
                            {
                                texts.Add(s);
                            }
                        }

                        string mergedText = string.Join(separator, texts);

                        Range? rowRange = null;
                        try
                        {
                            rowRange = selection.Rows[r] as Range;
                            if (rowRange != null)
                            {
                                rowRange.ClearContents();
                                rowRange.Merge();
                                rowRange.Value2 = mergedText;
                                if (separator == "\n") rowRange.WrapText = true;
                                mergedGroups++;
                            }
                        }
                        catch { }
                        finally
                        {
                            if (rowRange != null) Marshal.ReleaseComObject(rowRange);
                        }
                    }
                }
                else if (options.Direction == SafeMergeDirection.DownColumns)
                {
                    // Gộp theo từng cột (Column by column)
                    for (int c = 1; c <= colCount; c++)
                    {
                        var texts = new List<string>();
                        for (int r = 1; r <= rowCount; r++)
                        {
                            object? val = rawValues[r, c];
                            string s = val?.ToString() ?? "";
                            if (options.TrimSpaces) s = s.Trim();

                            if (!options.IgnoreBlankCells || !string.IsNullOrEmpty(s))
                            {
                                texts.Add(s);
                            }
                        }

                        string mergedText = string.Join(separator, texts);

                        Range? colRange = null;
                        try
                        {
                            colRange = selection.Columns[c] as Range;
                            if (colRange != null)
                            {
                                colRange.ClearContents();
                                colRange.Merge();
                                colRange.Value2 = mergedText;
                                if (separator == "\n") colRange.WrapText = true;
                                mergedGroups++;
                            }
                        }
                        catch { }
                        finally
                        {
                            if (colRange != null) Marshal.ReleaseComObject(colRange);
                        }
                    }
                }
                else // AllToOneCell
                {
                    // Gộp toàn bộ vùng chọn thành 1 ô duy nhất
                    var texts = new List<string>();
                    for (int r = 1; r <= rowCount; r++)
                    {
                        for (int c = 1; c <= colCount; c++)
                        {
                            object? val = rawValues[r, c];
                            string s = val?.ToString() ?? "";
                            if (options.TrimSpaces) s = s.Trim();

                            if (!options.IgnoreBlankCells || !string.IsNullOrEmpty(s))
                            {
                                texts.Add(s);
                            }
                        }
                    }

                    string mergedText = string.Join(separator, texts);

                    selection.ClearContents();
                    selection.Merge();
                    selection.Value2 = mergedText;
                    if (separator == "\n") selection.WrapText = true;
                    mergedGroups = 1;
                }

                return (true, $"Đã gộp thành công {mergedGroups} nhóm ô bảo toàn 100% dữ liệu!", mergedGroups);
            }
            catch (Exception ex)
            {
                return (false, $"Lỗi khi gộp ô: {ex.Message}", 0);
            }
            finally
            {
                try
                {
                    ws.Application.DisplayAlerts = true;
                    ws.Application.ScreenUpdating = true;
                }
                catch { }
            }
        }

        #endregion

        #region 3. GỘP NHIỀU SHEETS THÀNH 1 SHEET (COMBINE SHEETS)

        /// <summary>
        /// Gộp dữ liệu từ nhiều Sheets trong Workbook thành 1 Sheet Tổng Hợp duy nhất
        /// </summary>
        public static (bool Success, string Message, int TotalRowsConsolidated) CombineSheetsIntoOne(
            Workbook targetWb,
            List<_Worksheet> sourceSheets,
            CombineSheetsOptions options)
        {
            if (targetWb == null || sourceSheets == null || sourceSheets.Count == 0)
            {
                return (false, "Không có sheet nào được chọn để gộp.", 0);
            }

            _Worksheet? newSheet = null;

            try
            {
                targetWb.Application.ScreenUpdating = false;

                // 1. Quét và thu thập dữ liệu từ tất cả các sheet nguồn
                var allRowsData = new List<List<object?>>();
                var headerRowData = new List<object?>();
                int maxCols = 0;

                for (int sIdx = 0; sIdx < sourceSheets.Count; sIdx++)
                {
                    var ws = sourceSheets[sIdx];
                    Range? usedRange = null;

                    try
                    {
                        usedRange = ws.UsedRange;
                        if (usedRange == null || usedRange.Rows.Count == 0) continue;

                        int totalRows = usedRange.Rows.Count;
                        int totalCols = usedRange.Columns.Count;
                        if (totalRows == 0 || totalCols == 0) continue;

                        object[,] values = (object[,])usedRange.Value2;
                        maxCols = Math.Max(maxCols, totalCols);

                        int startDataRow = 1;

                        // Nếu có dòng tiêu đề và đây là sheet đầu tiên -> Lấy tiêu đề
                        if (options.HasHeaderRow)
                        {
                            if (headerRowData.Count == 0)
                            {
                                for (int c = 1; c <= totalCols; c++)
                                {
                                    headerRowData.Add(values[1, c]);
                                }
                            }
                            startDataRow = options.HeaderRowCount + 1;
                        }

                        // Lấy các dòng dữ liệu
                        for (int r = startDataRow; r <= totalRows; r++)
                        {
                            var rowList = new List<object?>();
                            bool hasData = false;

                            for (int c = 1; c <= totalCols; c++)
                            {
                                object? cellVal = values[r, c];
                                rowList.Add(cellVal);
                                if (cellVal != null && !string.IsNullOrWhiteSpace(cellVal.ToString()))
                                {
                                    hasData = true;
                                }
                            }

                            if (options.SkipBlankRows && !hasData) continue;

                            // Thêm tên sheet nguồn nếu được chọn
                            if (options.AddSourceColumn)
                            {
                                rowList.Add(ws.Name);
                            }

                            allRowsData.Add(rowList);
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Read sheet {ws.Name} error: {ex.Message}");
                    }
                    finally
                    {
                        if (usedRange != null) Marshal.ReleaseComObject(usedRange);
                    }
                }

                if (allRowsData.Count == 0)
                {
                    return (false, "Không tìm thấy dữ liệu nào trong các sheet được chọn.", 0);
                }

                // 2. Tạo sheet tổng hợp mới
                string baseName = "TongHop_DuLieu";
                string sheetName = baseName;
                int suffix = 1;

                while (true)
                {
                    bool exists = false;
                    foreach (_Worksheet s in targetWb.Worksheets)
                    {
                        if (string.Equals(s.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                        {
                            exists = true;
                            Marshal.ReleaseComObject(s);
                            break;
                        }
                        Marshal.ReleaseComObject(s);
                    }
                    if (!exists) break;
                    sheetName = $"{baseName}_{suffix++}";
                }

                newSheet = targetWb.Worksheets.Add() as _Worksheet;
                if (newSheet == null) return (false, "Không thể tạo sheet mới.", 0);
                newSheet.Name = sheetName;

                int finalCols = maxCols + (options.AddSourceColumn ? 1 : 0);
                int totalOutputRows = allRowsData.Count + (options.HasHeaderRow ? 1 : 0);

                object?[,] outputArray = new object?[totalOutputRows + 1, finalCols + 1];

                int currentOutRow = 1;

                // Ghi tiêu đề
                if (options.HasHeaderRow)
                {
                    for (int c = 0; c < headerRowData.Count; c++)
                    {
                        outputArray[1, c + 1] = headerRowData[c];
                    }
                    if (options.AddSourceColumn)
                    {
                        outputArray[1, finalCols] = options.SourceColumnHeader;
                    }
                    currentOutRow = 2;
                }

                // Ghi dữ liệu
                for (int i = 0; i < allRowsData.Count; i++)
                {
                    var rowList = allRowsData[i];
                    for (int c = 0; c < rowList.Count; c++)
                    {
                        outputArray[currentOutRow, c + 1] = rowList[c];
                    }
                    currentOutRow++;
                }

                // Gán vào sheet mới qua 1 mảng 2D duy nhất
                Range destRange = newSheet.Range[newSheet.Cells[1, 1], newSheet.Cells[totalOutputRows, finalCols]];
                destRange.Value2 = outputArray;

                // Định dạng tiêu đề đẹp mắt
                if (options.HasHeaderRow)
                {
                    Range headerRange = newSheet.Range[newSheet.Cells[1, 1], newSheet.Cells[1, finalCols]];
                    headerRange.Font.Bold = true;
                    headerRange.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.FromArgb(16, 124, 65)); // Xanh Excel
                    headerRange.Font.Color = ColorTranslator.ToOle(System.Drawing.Color.White);
                    Marshal.ReleaseComObject(headerRange);
                }

                newSheet.Columns.AutoFit();
                Marshal.ReleaseComObject(destRange);

                return (true, $"Đã gộp thành công {sourceSheets.Count} sheet với tổng cộng {allRowsData.Count:N0} dòng vào sheet '{sheetName}'!", allRowsData.Count);
            }
            catch (Exception ex)
            {
                return (false, $"Lỗi khi gộp sheet: {ex.Message}", 0);
            }
            finally
            {
                try { targetWb.Application.ScreenUpdating = true; } catch { }
                if (newSheet != null) Marshal.ReleaseComObject(newSheet);
            }
        }

        #endregion
    }
}
