using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using ExcelSupport.Models;
using Microsoft.Office.Interop.Excel;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;

namespace ExcelSupport.Services
{
    public static class SpecialCopyService
    {
        public static string? CachedSpecialCopyText { get; private set; }
        public static List<string>? CachedSpecialCopyLines { get; private set; }
        public static SpecialCopyOptions LastOptions { get; set; } = new();

        /// <summary>
        /// Thực hiện sao chép nối chuỗi từ vùng chọn Excel và lưu vào Clipboard
        /// </summary>
        public static SpecialCopyResult ExecuteSpecialCopy(ExcelApp? app, Range? sourceRange = null, SpecialCopyOptions? options = null)
        {
            var result = new SpecialCopyResult();
            if (app == null)
            {
                result.Success = false;
                result.Message = LocalizationService.Get("FCP_MsgExcelNotReady");
                return result;
            }

            options ??= LastOptions ?? new SpecialCopyOptions();
            LastOptions = options;

            try
            {
                // Hủy chế độ CutCopyMode trước đó của Excel (xóa đường viền nhấp nháy marching ants)
                try { app.CutCopyMode = (XlCutCopyMode)0; } catch { }

                Range? rng = sourceRange ?? (app.Selection as Range);
                if (rng == null)
                {
                    result.Success = false;
                    result.Message = LocalizationService.Get("FCP_MsgSelectRange");
                    return result;
                }

                Range targetRange = rng;
                if (options.VisibleOnly)
                {
                    try
                    {
                        targetRange = rng.SpecialCells(XlCellType.xlCellTypeVisible);
                    }
                    catch
                    {
                        targetRange = rng; // Không có ô ẩn
                    }
                }

                if (targetRange == null)
                {
                    result.Success = false;
                    result.Message = LocalizationService.Get("FCP_MsgNoVisibleCells");
                    return result;
                }

                // Trích xuất ma trận giá trị theo dòng và cột (kế thừa logic Copy Filtered Cells)
                var rowDictValues = new SortedDictionary<int, SortedDictionary<int, object?>>();
                int totalCells = 0;

                foreach (Range area in targetRange.Areas)
                {
                    int rowCount = area.Rows.Count;
                    int colCount = area.Columns.Count;
                    int baseRow = area.Row;
                    int baseCol = area.Column;

                    object? rawValues = area.Value2;
                    object?[,]? valArray = rawValues as object[,];

                    for (int r = 1; r <= rowCount; r++)
                    {
                        int actualRow = baseRow + r - 1;
                        if (!rowDictValues.ContainsKey(actualRow))
                        {
                            rowDictValues[actualRow] = new SortedDictionary<int, object?>();
                        }

                        for (int c = 1; c <= colCount; c++)
                        {
                            int actualCol = baseCol + c - 1;
                            object? val = (valArray != null) ? valArray[r, c] : rawValues;
                            rowDictValues[actualRow][actualCol] = val;
                            totalCells++;
                        }
                    }
                }

                if (rowDictValues.Count == 0)
                {
                    result.Success = false;
                    result.Message = LocalizationService.Get("FCP_MsgNoVisibleCells");
                    return result;
                }

                string delimiter = options.GetDelimiterString();
                var lines = new List<string>();

                if (options.JoinMode == SpecialCopyJoinMode.AllCells)
                {
                    // Nối tất cả các ô (toàn bộ các dòng và các cột) thành 1 chuỗi / 1 dòng duy nhất
                    var allCellStrings = new List<string>();
                    foreach (var rowKvp in rowDictValues)
                    {
                        foreach (var colVal in rowKvp.Value)
                        {
                            object? val = colVal.Value;
                            if (options.SkipBlanks && IsBlank(val)) continue;
                            allCellStrings.Add(FormatCellString(val, delimiter, options));
                        }
                    }
                    lines.Add(string.Join(delimiter, allCellStrings));
                }
                else if (options.JoinMode == SpecialCopyJoinMode.ByColumn)
                {
                    // Nhóm theo từng cột: mỗi cột nối thành 1 dòng
                    var colDictValues = new SortedDictionary<int, SortedDictionary<int, object?>>();
                    foreach (var rowKvp in rowDictValues)
                    {
                        int rowIdx = rowKvp.Key;
                        foreach (var colKvp in rowKvp.Value)
                        {
                            int colIdx = colKvp.Key;
                            if (!colDictValues.ContainsKey(colIdx))
                            {
                                colDictValues[colIdx] = new SortedDictionary<int, object?>();
                            }
                            colDictValues[colIdx][rowIdx] = colKvp.Value;
                        }
                    }

                    foreach (var kvp in colDictValues)
                    {
                        var cellStrings = new List<string>();
                        foreach (var rowVal in kvp.Value)
                        {
                            object? val = rowVal.Value;
                            if (options.SkipBlanks && IsBlank(val)) continue;
                            cellStrings.Add(FormatCellString(val, delimiter, options));
                        }

                        if (cellStrings.Count > 0 || !options.SkipBlanks)
                        {
                            lines.Add(string.Join(delimiter, cellStrings));
                        }
                    }
                }
                else // SpecialCopyJoinMode.ByRow
                {
                    // Nhóm theo từng dòng: mỗi dòng các cột nối lại với nhau
                    foreach (var kvp in rowDictValues)
                    {
                        var cellStrings = new List<string>();
                        foreach (var colVal in kvp.Value)
                        {
                            object? val = colVal.Value;
                            if (options.SkipBlanks && IsBlank(val)) continue;
                            cellStrings.Add(FormatCellString(val, delimiter, options));
                        }

                        if (cellStrings.Count > 0 || !options.SkipBlanks)
                        {
                            lines.Add(string.Join(delimiter, cellStrings));
                        }
                    }
                }

                string fullText = string.Join(Environment.NewLine, lines);

                // Lưu vào bộ nhớ đệm SpecialCopy
                CachedSpecialCopyText = fullText;
                CachedSpecialCopyLines = lines;

                // Đồng bộ vào FilteredCopyPasteService để nếu bấm "Paste to Visible Cells" thì vẫn nhận chuỗi nối
                var cacheMatrix = new List<List<object?>>();
                foreach (var line in lines)
                {
                    cacheMatrix.Add(new List<object?> { line });
                }
                FilteredCopyPasteService.SetCustomCache(cacheMatrix);

                // Đưa vào Windows Clipboard chuẩn Unicode (có cơ chế thử lại nhiều lần nếu clipboard bị lock)
                bool clipOk = SetClipboardTextWithRetry(fullText);

                // Đảm bảo Excel CutCopyMode được giải phóng để Ctrl+V dán từ Windows Clipboard
                try { app.CutCopyMode = (XlCutCopyMode)0; } catch { }

                result.Success = clipOk;
                result.ResultText = fullText;
                result.TotalCellsProcessed = totalCells;
                result.TotalLines = lines.Count;
                result.Message = string.Format(
                    LocalizationService.Get("SC_MsgCopySuccess"), 
                    totalCells, 
                    lines.Count, 
                    GetDelimiterDisplay(options.Delimiter));
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"{LocalizationService.Get("Common_Error")}: {ex.Message}";
            }

            return result;
        }

        /// <summary>
        /// Sao chép nhanh 1-chạm với dấu nối chỉ định (mặc định nối toàn bộ các ô/dòng được chọn thành 1 dòng CSV)
        /// </summary>
        public static SpecialCopyResult QuickCopy(ExcelApp? app, SpecialCopyDelimiter delimiter)
        {
            var options = new SpecialCopyOptions
            {
                Delimiter = delimiter,
                JoinMode = SpecialCopyJoinMode.AllCells, // Nối toàn bộ các ô đã chọn (cả dòng và cột) thành 1 dòng CSV
                QuoteMode = SpecialCopyQuoteMode.None,
                SkipBlanks = false,
                TrimSpaces = false,
                VisibleOnly = true
            };

            return ExecuteSpecialCopy(app, null, options);
        }

        /// <summary>
        /// Dán chuỗi đã sao chép vào ô hiện tại hoặc các ô hiển thị của vùng đích
        /// </summary>
        public static FilteredPasteResult PasteSpecialCopy(ExcelApp? app, Range? targetRange = null)
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

                List<string> linesToPaste = CachedSpecialCopyLines ?? new List<string>();
                if (linesToPaste.Count == 0)
                {
                    // Thử đọc từ Clipboard nếu cache trống
                    if (System.Windows.Forms.Clipboard.ContainsText())
                    {
                        string text = System.Windows.Forms.Clipboard.GetText();
                        var parts = text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
                        linesToPaste = parts.ToList();
                    }
                }

                if (linesToPaste.Count == 0)
                {
                    result.Success = false;
                    result.Message = LocalizationService.Get("SC_MsgClipboardEmpty");
                    return result;
                }

                int startRow = destRng.Row;
                int startCol = destRng.Column;
                int maxRows = ws.Rows.Count;

                bool prevScreen = app.ScreenUpdating;
                bool prevEvents = app.EnableEvents;
                XlCalculation prevCalc = app.Calculation;

                try
                {
                    app.ScreenUpdating = false;
                    app.EnableEvents = false;
                    app.Calculation = XlCalculation.xlCalculationManual;

                    int pastedCount = 0;
                    int currentRow = startRow;

                    for (int i = 0; i < linesToPaste.Count && currentRow <= maxRows; i++)
                    {
                        // Bỏ qua các dòng bị ẩn
                        while (currentRow <= maxRows)
                        {
                            Range rowRng = ws.Rows[currentRow];
                            bool isHidden = false;
                            try { isHidden = Convert.ToBoolean(rowRng.EntireRow.Hidden); }
                            catch { }

                            if (!isHidden) break;
                            currentRow++;
                        }

                        if (currentRow > maxRows) break;

                        Range cell = ws.Cells[currentRow, startCol];
                        cell.Value2 = linesToPaste[i];
                        pastedCount++;
                        currentRow++;
                    }

                    // Giải phóng CutCopyMode
                    try { app.CutCopyMode = (XlCutCopyMode)0; } catch { }

                    result.Success = true;
                    result.RowsPasted = pastedCount;
                    result.Message = string.Format(LocalizationService.Get("SC_MsgPasteSuccess"), pastedCount);
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
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"{LocalizationService.Get("Common_Error")}: {ex.Message}";
            }

            return result;
        }

        private static string FormatCellString(object? val, string delimiter, SpecialCopyOptions options)
        {
            string s = val?.ToString() ?? string.Empty;
            if (options.TrimSpaces)
            {
                s = s.Trim();
            }

            switch (options.QuoteMode)
            {
                case SpecialCopyQuoteMode.AutoCsv:
                    if (s.Contains(delimiter) || s.Contains("\"") || s.Contains("\r") || s.Contains("\n"))
                    {
                        s = "\"" + s.Replace("\"", "\"\"") + "\"";
                    }
                    break;

                case SpecialCopyQuoteMode.AlwaysDoubleQuote:
                    s = "\"" + s.Replace("\"", "\"\"") + "\"";
                    break;

                case SpecialCopyQuoteMode.SingleQuoteSql:
                    s = "'" + s.Replace("'", "''") + "'";
                    break;

                case SpecialCopyQuoteMode.None:
                default:
                    break;
            }

            return s;
        }

        private static bool IsBlank(object? val)
        {
            return val == null || string.IsNullOrWhiteSpace(val.ToString());
        }

        public static bool SetClipboardTextWithRetry(string text, int retries = 10, int delayMs = 50)
        {
            for (int i = 0; i < retries; i++)
            {
                try
                {
                    var dataObj = new System.Windows.Forms.DataObject();
                    dataObj.SetData(System.Windows.Forms.DataFormats.UnicodeText, true, text);
                    dataObj.SetData(System.Windows.Forms.DataFormats.Text, true, text);
                    dataObj.SetData(System.Windows.Forms.DataFormats.StringFormat, true, text);
                    System.Windows.Forms.Clipboard.SetDataObject(dataObj, true, 5, 50);
                    return true;
                }
                catch
                {
                    try
                    {
                        System.Windows.Clipboard.SetText(text);
                        return true;
                    }
                    catch
                    {
                        Thread.Sleep(delayMs);
                    }
                }
            }
            return false;
        }

        public static string GetDelimiterDisplay(SpecialCopyDelimiter delimiter)
        {
            return delimiter switch
            {
                SpecialCopyDelimiter.Comma => "Dấu phẩy (,)",
                SpecialCopyDelimiter.CommaSpace => "Phẩy kèm cách (, )",
                SpecialCopyDelimiter.Semicolon => "Chấm phẩy (;)",
                SpecialCopyDelimiter.Pipe => "Gạch đứng (|)",
                SpecialCopyDelimiter.Tab => "Tab",
                SpecialCopyDelimiter.Space => "Khoảng cách",
                SpecialCopyDelimiter.Newline => "Xuống dòng",
                SpecialCopyDelimiter.Custom => "Tùy chỉnh",
                _ => ","
            };
        }
    }
}
