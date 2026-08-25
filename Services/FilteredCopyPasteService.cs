using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelSupport.Models;
using Microsoft.Office.Interop.Excel;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;

namespace ExcelSupport.Services
{
    public static class FilteredCopyPasteService
    {
        // Bộ nhớ đệm lưu dữ liệu đã copy gần nhất
        public static List<List<object?>>? CachedValues { get; private set; }
        public static List<List<string?>>? CachedFormulas { get; private set; }
        public static int CachedRowCount => CachedValues?.Count ?? 0;
        public static int CachedColCount => (CachedValues != null && CachedValues.Count > 0) ? CachedValues[0].Count : 0;

        /// <summary>
        /// Sao chép chỉ các ô hiển thị (Visible Cells Only) vào Clipboard và bộ nhớ đệm
        /// </summary>
        public static FilteredPasteResult CopyVisibleCells(ExcelApp? app, Range? sourceRange = null)
        {
            var result = new FilteredPasteResult();
            if (app == null)
            {
                result.Success = false;
                result.Message = LocalizationService.Get("FCP_MsgExcelNotReady");
                return result;
            }

            try
            {
                Range? rng = sourceRange ?? (app.Selection as Range);
                if (rng == null)
                {
                    result.Success = false;
                    result.Message = LocalizationService.Get("FCP_MsgSelectRange");
                    return result;
                }

                Range? visibleRange = null;
                try
                {
                    visibleRange = rng.SpecialCells(XlCellType.xlCellTypeVisible);
                }
                catch
                {
                    visibleRange = rng; // Nếu không có ô ẩn thì lấy nguyên vùng
                }

                if (visibleRange == null)
                {
                    result.Success = false;
                    result.Message = LocalizationService.Get("FCP_MsgNoVisibleCells");
                    return result;
                }

                // 1. Lưu vào bộ đệm in-memory
                ExtractVisibleData(visibleRange);

                // 2. Đưa vào Clipboard chuẩn của Excel
                visibleRange.Copy();

                int rowCount = CachedRowCount;
                int colCount = CachedColCount;

                result.Success = true;
                result.SourceRowCount = rowCount;
                result.RowsPasted = rowCount;
                result.Message = string.Format(LocalizationService.Get("FCP_MsgCopySuccess"), rowCount, colCount);
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"{LocalizationService.Get("Common_Error")}: {ex.Message}";
            }

            return result;
        }

        /// <summary>
        /// Dán dữ liệu vào các dòng hiển thị (Bỏ qua các dòng bị ẩn/bị lọc)
        /// </summary>
        public static FilteredPasteResult PasteToVisibleCells(ExcelApp? app, Range? targetRange = null, FilteredPasteOptions? options = null)
        {
            var result = new FilteredPasteResult();
            if (app == null)
            {
                result.Success = false;
                result.Message = LocalizationService.Get("FCP_MsgExcelNotReady");
                return result;
            }

            options ??= new FilteredPasteOptions();

            bool prevScreen = app.ScreenUpdating;
            bool prevEvents = app.EnableEvents;
            XlCalculation prevCalc = app.Calculation;

            try
            {
                Range? destRng = targetRange ?? (app.Selection as Range);
                if (destRng == null)
                {
                    result.Success = false;
                    result.Message = LocalizationService.Get("FCP_MsgSelectTarget");
                    return result;
                }

                Worksheet? ws = destRng.Worksheet;
                if (ws == null)
                {
                    result.Success = false;
                    result.Message = LocalizationService.Get("FCP_MsgTargetSheetNotFound");
                    return result;
                }

                // Lấy dữ liệu nguồn (ưu tiên từ Cache, nếu không có lấy từ Windows Clipboard)
                List<List<object?>> sourceData = GetSourceData(options);
                if (sourceData.Count == 0)
                {
                    result.Success = false;
                    result.Message = LocalizationService.Get("FCP_MsgCacheEmpty");
                    return result;
                }

                int srcRows = sourceData.Count;
                int srcCols = sourceData[0].Count;

                // Thu thập danh sách các dòng hiển thị tại đích
                int startRow = destRng.Row;
                int startCol = destRng.Column;
                int destTotalRows = destRng.Rows.Count;
                int maxRowsToScan = ws.Rows.Count;

                var visibleTargetRows = new List<int>();
                int hiddenCount = 0;

                if (destTotalRows > 1)
                {
                    Range? visibleDest = null;
                    try { visibleDest = destRng.SpecialCells(XlCellType.xlCellTypeVisible); }
                    catch { visibleDest = destRng; }

                    if (visibleDest != null)
                    {
                        var rowSet = new SortedSet<int>();
                        foreach (Range area in visibleDest.Areas)
                        {
                            int areaStart = area.Row;
                            int areaCount = area.Rows.Count;
                            for (int r = 0; r < areaCount; r++)
                            {
                                rowSet.Add(areaStart + r);
                            }
                        }
                        visibleTargetRows.AddRange(rowSet);
                    }
                    hiddenCount = destTotalRows - visibleTargetRows.Count;
                }
                else
                {
                    // Nếu chọn 1 ô đơn, tìm đủ srcRows dòng hiển thị từ startRow trở xuống
                    int currentRow = startRow;
                    while (visibleTargetRows.Count < srcRows && currentRow <= maxRowsToScan)
                    {
                        Range rowRng = ws.Rows[currentRow];
                        bool isHidden = false;
                        try { isHidden = Convert.ToBoolean(rowRng.EntireRow.Hidden); }
                        catch { }

                        if (!isHidden)
                        {
                            visibleTargetRows.Add(currentRow);
                        }
                        else
                        {
                            hiddenCount++;
                        }
                        currentRow++;
                    }
                }

                if (visibleTargetRows.Count == 0)
                {
                    result.Success = false;
                    result.Message = LocalizationService.Get("FCP_MsgNoVisibleCells");
                    return result;
                }

                // Tạm tắt cập nhật màn hình để tăng tốc tối đa
                app.ScreenUpdating = false;
                app.EnableEvents = false;
                app.Calculation = XlCalculation.xlCalculationManual;

                int pastedRows = 0;

                for (int i = 0; i < visibleTargetRows.Count; i++)
                {
                    int targetRowIdx = visibleTargetRows[i];
                    int srcRowIdx = i;

                    if (srcRowIdx >= srcRows)
                    {
                        if (options.RepeatIfShorter)
                        {
                            srcRowIdx = i % srcRows;
                        }
                        else
                        {
                            break; // Hết dữ liệu nguồn
                        }
                    }

                    var rowData = sourceData[srcRowIdx];

                    for (int c = 0; c < srcCols; c++)
                    {
                        int targetColIdx = startCol + c;
                        object? val = (c < rowData.Count) ? rowData[c] : null;

                        if (options.SkipBlanks && (val == null || string.IsNullOrWhiteSpace(val.ToString())))
                        {
                            continue;
                        }

                        Range cell = ws.Cells[targetRowIdx, targetColIdx];

                        if (options.PasteType == FilteredPasteType.Formulas && CachedFormulas != null && srcRowIdx < CachedFormulas.Count && c < CachedFormulas[srcRowIdx].Count)
                        {
                            string? formula = CachedFormulas[srcRowIdx][c];
                            if (!string.IsNullOrEmpty(formula))
                            {
                                cell.Formula = formula;
                            }
                            else
                            {
                                cell.Value2 = val;
                            }
                        }
                        else
                        {
                            cell.Value2 = val;
                        }
                    }

                    pastedRows++;
                }

                result.Success = true;
                result.SourceRowCount = srcRows;
                result.TargetVisibleRowCount = visibleTargetRows.Count;
                result.RowsPasted = pastedRows;
                result.HiddenRowsProtected = hiddenCount;
                result.Message = string.Format(LocalizationService.Get("FCP_MsgPasteSuccess"), pastedRows, hiddenCount);
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"{LocalizationService.Get("Common_Error")}: {ex.Message}";
            }
            finally
            {
                try
                {
                    app.ScreenUpdating = prevScreen;
                    app.EnableEvents = prevEvents;
                    app.Calculation = prevCalc;
                }
                catch { }
            }

            return result;
        }

        /// <summary>
        /// Thực thi sao chép trực tiếp từ Vùng Nguồn sang Vùng Đích
        /// </summary>
        public static FilteredPasteResult ExecuteRangeToRangeCopyPaste(
            ExcelApp app, 
            Range sourceRange, 
            Range targetRange, 
            FilteredPasteOptions options)
        {
            var copyResult = CopyVisibleCells(app, sourceRange);
            if (!copyResult.Success) return copyResult;

            return PasteToVisibleCells(app, targetRange, options);
        }

        private static void ExtractVisibleData(Range visibleRange)
        {
            CachedValues = new List<List<object?>>();
            CachedFormulas = new List<List<string?>>();

            // Duyệt qua từng Area của SpecialCells(xlCellTypeVisible)
            // Nhóm theo hàng để tạo ma trận dòng x cột chuẩn
            var rowDictValues = new SortedDictionary<int, SortedDictionary<int, object?>>();
            var rowDictFormulas = new SortedDictionary<int, SortedDictionary<int, string?>>();

            foreach (Range area in visibleRange.Areas)
            {
                int rowCount = area.Rows.Count;
                int colCount = area.Columns.Count;
                int baseRow = area.Row;
                int baseCol = area.Column;

                object? rawValues = area.Value2;
                object? rawFormulas = area.Formula;

                object?[,]? valArray = rawValues as object[,];
                object?[,]? formulaArray = rawFormulas as object[,];

                for (int r = 1; r <= rowCount; r++)
                {
                    int actualRow = baseRow + r - 1;
                    if (!rowDictValues.ContainsKey(actualRow))
                    {
                        rowDictValues[actualRow] = new SortedDictionary<int, object?>();
                        rowDictFormulas[actualRow] = new SortedDictionary<int, string?>();
                    }

                    for (int c = 1; c <= colCount; c++)
                    {
                        int actualCol = baseCol + c - 1;
                        object? val = (valArray != null) ? valArray[r, c] : rawValues;
                        string? formula = (formulaArray != null) ? formulaArray[r, c]?.ToString() : rawFormulas?.ToString();

                        rowDictValues[actualRow][actualCol] = val;
                        rowDictFormulas[actualRow][actualCol] = formula;
                    }
                }
            }

            foreach (var kvp in rowDictValues)
            {
                var rowList = new List<object?>();
                foreach (var colVal in kvp.Value)
                {
                    rowList.Add(colVal.Value);
                }
                CachedValues.Add(rowList);
            }

            foreach (var kvp in rowDictFormulas)
            {
                var rowList = new List<string?>();
                foreach (var colVal in kvp.Value)
                {
                    rowList.Add(colVal.Value);
                }
                CachedFormulas.Add(rowList);
            }
        }

        private static List<List<object?>> GetSourceData(FilteredPasteOptions options)
        {
            if (CachedValues != null && CachedValues.Count > 0)
            {
                return CachedValues;
            }

            // Fallback đọc từ Clipboard hệ thống nếu người dùng đã copy từ ngoài
            var list = new List<List<object?>>();
            try
            {
                if (System.Windows.Forms.Clipboard.ContainsText())
                {
                    string text = System.Windows.Forms.Clipboard.GetText();
                    var lines = text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
                    foreach (var line in lines)
                    {
                        if (string.IsNullOrEmpty(line) && line == lines.Last()) continue;
                        var cols = line.Split('\t');
                        var row = new List<object?>();
                        foreach (var c in cols)
                        {
                            row.Add(c);
                        }
                        list.Add(row);
                    }
                }
            }
            catch { }

            return list;
        }
    }
}
