using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using ExcelSupport.Models;
using Microsoft.Office.Interop.Excel;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;

namespace ExcelSupport.Services
{
    public static class TableMergeService
    {
        public static List<MergeColumnItem> GetSheetColumns(ExcelApp? app, string wbName, string sheetName, int headerRow = 1)
        {
            var list = new List<MergeColumnItem>();
            if (app == null || string.IsNullOrEmpty(wbName) || string.IsNullOrEmpty(sheetName)) return list;

            Workbook? wb = null;
            _Worksheet? ws = null;
            Range? usedRange = null;

            try
            {
                foreach (Workbook w in app.Workbooks)
                {
                    if (string.Equals(w.Name, wbName, StringComparison.OrdinalIgnoreCase))
                    {
                        wb = w;
                        break;
                    }
                }

                if (wb == null) return list;

                foreach (_Worksheet s in wb.Worksheets)
                {
                    if (string.Equals(s.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        ws = s;
                        break;
                    }
                }

                if (ws == null) return list;

                usedRange = ws.UsedRange;
                if (usedRange == null || usedRange.Rows.Count == 0 || usedRange.Columns.Count == 0) return list;

                int startCol = usedRange.Column;
                int totalCols = usedRange.Columns.Count;
                int startRow = usedRange.Row;

                int targetRowOffset = headerRow - startRow + 1;
                object? rawVal = usedRange.Value2;

                if (rawVal is object[,] allVals)
                {
                    int maxRows = allVals.GetLength(0);
                    int actualRow = (targetRowOffset >= 1 && targetRowOffset <= maxRows) ? targetRowOffset : 1;

                    for (int c = 1; c <= totalCols; c++)
                    {
                        int absoluteCol = startCol + c - 1;
                        string colLetter = ConvertColIndexToLetter(absoluteCol);
                        string headerText = allVals[actualRow, c]?.ToString()?.Trim() ?? string.Empty;

                        list.Add(new MergeColumnItem
                        {
                            ColumnIndex = absoluteCol,
                            ColumnLetter = colLetter,
                            HeaderText = headerText,
                            OutputHeaderName = !string.IsNullOrEmpty(headerText) ? headerText : $"Cột_{colLetter}",
                            IsSelected = true
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"GetSheetColumns error: {ex.Message}");
            }
            finally
            {
                if (usedRange != null) Marshal.ReleaseComObject(usedRange);
                if (ws != null) Marshal.ReleaseComObject(ws);
            }

            return list;
        }

        public static TableMergeResult ExecuteTableMerge(ExcelApp app, TableMergeOptions options)
        {
            var result = new TableMergeResult();
            if (app == null || options == null)
            {
                result.Success = false;
                result.Message = "Cấu hình ghép bảng không hợp lệ.";
                return result;
            }

            var selectedCols2 = options.SelectedColumnsFromTable2.Where(c => c.IsSelected).ToList();
            if (selectedCols2.Count == 0)
            {
                result.Success = false;
                result.Message = "Vui lòng chọn ít nhất một cột từ Bảng 2 cần ghép sang.";
                return result;
            }

            Workbook? wb1 = null, wb2 = null;
            _Worksheet? ws1 = null, ws2 = null;
            Range? usedRange1 = null, usedRange2 = null;

            try
            {
                app.ScreenUpdating = false;
                app.Calculation = XlCalculation.xlCalculationManual;

                // Tìm Workbook & Sheet 1
                foreach (Workbook w in app.Workbooks)
                {
                    if (string.Equals(w.Name, options.Table1WorkbookName, StringComparison.OrdinalIgnoreCase)) { wb1 = w; break; }
                }
                if (wb1 == null) wb1 = app.ActiveWorkbook;
                if (wb1 == null) { result.Success = false; result.Message = "Không tìm thấy file của Bảng 1."; return result; }

                foreach (_Worksheet s in wb1.Worksheets)
                {
                    if (string.Equals(s.Name, options.Table1SheetName, StringComparison.OrdinalIgnoreCase)) { ws1 = s; break; }
                }
                if (ws1 == null) ws1 = wb1.ActiveSheet as _Worksheet;
                if (ws1 == null) { result.Success = false; result.Message = "Không tìm thấy Sheet của Bảng 1."; return result; }

                // Tìm Workbook & Sheet 2
                foreach (Workbook w in app.Workbooks)
                {
                    if (string.Equals(w.Name, options.Table2WorkbookName, StringComparison.OrdinalIgnoreCase)) { wb2 = w; break; }
                }
                if (wb2 == null) wb2 = wb1;

                foreach (_Worksheet s in wb2.Worksheets)
                {
                    if (string.Equals(s.Name, options.Table2SheetName, StringComparison.OrdinalIgnoreCase)) { ws2 = s; break; }
                }
                if (ws2 == null) { result.Success = false; result.Message = "Không tìm thấy Sheet của Bảng 2."; return result; }

                // Đọc dữ liệu Bảng 1
                usedRange1 = ws1.UsedRange;
                if (usedRange1 == null || usedRange1.Rows.Count == 0 || usedRange1.Columns.Count == 0)
                {
                    result.Success = false; result.Message = "Bảng 1 không có dữ liệu."; return result;
                }

                // Đọc dữ liệu Bảng 2
                usedRange2 = ws2.UsedRange;
                if (usedRange2 == null || usedRange2.Rows.Count == 0 || usedRange2.Columns.Count == 0)
                {
                    result.Success = false; result.Message = "Bảng 2 không có dữ liệu."; return result;
                }

                int startRow1 = usedRange1.Row, startCol1 = usedRange1.Column, numRows1 = usedRange1.Rows.Count, numCols1 = usedRange1.Columns.Count;
                int startRow2 = usedRange2.Row, startCol2 = usedRange2.Column, numRows2 = usedRange2.Rows.Count, numCols2 = usedRange2.Columns.Count;

                object?[,]? val1 = usedRange1.Value2 as object[,];
                object?[,]? val2 = usedRange2.Value2 as object[,];

                if (val1 == null || val2 == null)
                {
                    result.Success = false; result.Message = "Không thể đọc dữ liệu dạng mảng 2D từ bảng tính."; return result;
                }

                // Xây dựng Dictionary tra cứu cho Bảng 2
                int keyColOffset2 = options.Table2KeyColIndex - startCol2 + 1;
                int headerRowOffset2 = options.Table2HeaderRow - startRow2 + 1;
                int startDataRow2 = Math.Max(1, headerRowOffset2 + 1);

                var stringComparer = options.MatchCase ? StringComparer.Ordinal : StringComparer.OrdinalIgnoreCase;
                var table2Map = new Dictionary<string, object?[]>(stringComparer);
                var table2UnmatchedSet = new HashSet<string>(stringComparer);

                for (int r = startDataRow2; r <= numRows2; r++)
                {
                    string rawKey = val2[r, keyColOffset2]?.ToString() ?? string.Empty;
                    string normalizedKey = NormalizeKey(rawKey, options.TrimSpaces, options.IgnoreAccent);
                    if (string.IsNullOrEmpty(normalizedKey)) continue;

                    if (!table2Map.ContainsKey(normalizedKey))
                    {
                        var rowData = new object?[selectedCols2.Count];
                        for (int c = 0; c < selectedCols2.Count; c++)
                        {
                            int colOffset = selectedCols2[c].ColumnIndex - startCol2 + 1;
                            rowData[c] = (colOffset >= 1 && colOffset <= numCols2) ? val2[r, colOffset] : null;
                        }
                        table2Map[normalizedKey] = rowData;
                        table2UnmatchedSet.Add(normalizedKey);
                    }
                }

                // Xử lý ghép với Bảng 1
                int keyColOffset1 = options.Table1KeyColIndex - startCol1 + 1;
                int headerRowOffset1 = options.Table1HeaderRow - startRow1 + 1;
                int startDataRow1 = Math.Max(1, headerRowOffset1 + 1);

                int matchedCount = 0;
                int unmatchedCount = 0;

                if (options.OutputTarget == TableMergeOutputTarget.InsertAdjacentToTable1 && options.JoinType != TableJoinType.FullOuterJoin)
                {
                    // Chèn trực tiếp vào các cột bên phải Bảng 1
                    int targetStartCol = startCol1 + numCols1;
                    int outputRowCount = numRows1 - headerRowOffset1 + 1;
                    object[,] insertData = new object[outputRowCount, selectedCols2.Count];

                    // Header
                    for (int c = 0; c < selectedCols2.Count; c++)
                    {
                        insertData[1, c + 1] = !string.IsNullOrEmpty(selectedCols2[c].OutputHeaderName) 
                            ? selectedCols2[c].OutputHeaderName 
                            : selectedCols2[c].HeaderText;
                    }

                    // Data Rows
                    for (int r = startDataRow1; r <= numRows1; r++)
                    {
                        int outRow = r - headerRowOffset1 + 1;
                        string rawKey = val1[r, keyColOffset1]?.ToString() ?? string.Empty;
                        string normalizedKey = NormalizeKey(rawKey, options.TrimSpaces, options.IgnoreAccent);

                        if (!string.IsNullOrEmpty(normalizedKey) && table2Map.TryGetValue(normalizedKey, out var rowData2))
                        {
                            matchedCount++;
                            for (int c = 0; c < selectedCols2.Count; c++)
                            {
                                insertData[outRow, c + 1] = rowData2[c] ?? string.Empty;
                            }
                        }
                        else
                        {
                            unmatchedCount++;
                            for (int c = 0; c < selectedCols2.Count; c++)
                            {
                                insertData[outRow, c + 1] = string.Empty;
                            }
                        }
                    }

                    // Ghi mảng vào Excel
                    Range writeRange = ws1.Range[ws1.Cells[options.Table1HeaderRow, targetStartCol],
                                                 ws1.Cells[options.Table1HeaderRow + outputRowCount - 1, targetStartCol + selectedCols2.Count - 1]];
                    writeRange.Value2 = insertData;
                    writeRange.Columns.AutoFit();
                    Marshal.ReleaseComObject(writeRange);

                    result.Success = true;
                    result.TotalRowsMerged = numRows1 - startDataRow1 + 1;
                    result.MatchedRows = matchedCount;
                    result.UnmatchedRows = unmatchedCount;
                    result.OutputSheetName = ws1.Name;
                    result.Message = $"Đã chèn {selectedCols2.Count} cột ghép vào Sheet '{ws1.Name}' (Khớp: {matchedCount:N0}, Không khớp: {unmatchedCount:N0})!";
                }
                else
                {
                    // Tạo Sheet mới
                    var rowsList = new List<object?[]>();

                    // Header Row
                    var headerList = new List<object?>();
                    for (int c = 1; c <= numCols1; c++)
                    {
                        headerList.Add(val1[headerRowOffset1, c] ?? $"Cột_{ConvertColIndexToLetter(startCol1 + c - 1)}");
                    }
                    foreach (var c in selectedCols2)
                    {
                        headerList.Add(!string.IsNullOrEmpty(c.OutputHeaderName) ? c.OutputHeaderName : c.HeaderText);
                    }
                    rowsList.Add(headerList.ToArray());

                    // Data Rows
                    for (int r = startDataRow1; r <= numRows1; r++)
                    {
                        string rawKey = val1[r, keyColOffset1]?.ToString() ?? string.Empty;
                        string normalizedKey = NormalizeKey(rawKey, options.TrimSpaces, options.IgnoreAccent);

                        object?[]? rowData2 = null;
                        bool isMatch = !string.IsNullOrEmpty(normalizedKey) && table2Map.TryGetValue(normalizedKey, out rowData2);

                        if (options.JoinType == TableJoinType.InnerJoin && !isMatch)
                        {
                            unmatchedCount++;
                            continue; // Bỏ qua dòng không khớp trong Inner Join
                        }

                        if (isMatch)
                        {
                            matchedCount++;
                            table2UnmatchedSet.Remove(normalizedKey);
                        }
                        else
                        {
                            unmatchedCount++;
                        }

                        var rowItems = new List<object?>();
                        for (int c = 1; c <= numCols1; c++) rowItems.Add(val1[r, c]);
                        for (int c = 0; c < selectedCols2.Count; c++)
                        {
                            rowItems.Add(isMatch ? (rowData2?[c] ?? string.Empty) : string.Empty);
                        }
                        rowsList.Add(rowItems.ToArray());
                    }

                    // Nếu là Full Outer Join: Thêm các dòng chỉ có ở Bảng 2
                    if (options.JoinType == TableJoinType.FullOuterJoin)
                    {
                        for (int r = startDataRow2; r <= numRows2; r++)
                        {
                            string rawKey = val2[r, keyColOffset2]?.ToString() ?? string.Empty;
                            string normalizedKey = NormalizeKey(rawKey, options.TrimSpaces, options.IgnoreAccent);

                            if (!string.IsNullOrEmpty(normalizedKey) && table2UnmatchedSet.Contains(normalizedKey))
                            {
                                var rowItems = new List<object?>();
                                for (int c = 1; c <= numCols1; c++)
                                {
                                    // Gán khóa vào cột khóa bảng 1, các cột khác để trống
                                    rowItems.Add((c == keyColOffset1) ? val2[r, keyColOffset2] : string.Empty);
                                }
                                for (int c = 0; c < selectedCols2.Count; c++)
                                {
                                    int colOffset = selectedCols2[c].ColumnIndex - startCol2 + 1;
                                    rowItems.Add(val2[r, colOffset]);
                                }
                                rowsList.Add(rowItems.ToArray());
                                table2UnmatchedSet.Remove(normalizedKey);
                            }
                        }
                    }

                    // Xuất ra Sheet mới
                    string sheetPrefix = "Ghep_Bang_";
                    string newSheetName = sheetPrefix + DateTime.Now.ToString("HHmmss");

                    _Worksheet newWs = (_Worksheet)wb1.Worksheets.Add(After: wb1.Worksheets[wb1.Worksheets.Count]);
                    newWs.Name = newSheetName;

                    int totalRows = rowsList.Count;
                    int totalCols = headerList.Count;
                    object[,] exportArr = new object[totalRows, totalCols];

                    for (int r = 0; r < totalRows; r++)
                    {
                        var rowArr = rowsList[r];
                        for (int c = 0; c < totalCols; c++)
                        {
                            exportArr[r + 1, c + 1] = (c < rowArr.Length && rowArr[c] != null) ? rowArr[c]! : string.Empty;
                        }
                    }

                    Range exportRange = newWs.Range[newWs.Cells[1, 1], newWs.Cells[totalRows, totalCols]];
                    exportRange.Value2 = exportArr;

                    // Định dạng Header màu xanh Navy chuyên nghiệp
                    Range headerRange = newWs.Range[newWs.Cells[1, 1], newWs.Cells[1, totalCols]];
                    headerRange.Font.Bold = true;
                    headerRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(16, 124, 65));
                    headerRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                    Marshal.ReleaseComObject(headerRange);

                    exportRange.Columns.AutoFit();
                    Marshal.ReleaseComObject(exportRange);
                    Marshal.ReleaseComObject(newWs);

                    result.Success = true;
                    result.TotalRowsMerged = totalRows - 1;
                    result.MatchedRows = matchedCount;
                    result.UnmatchedRows = unmatchedCount;
                    result.OutputSheetName = newSheetName;
                    result.Message = $"Đã tạo Sheet mới '{newSheetName}' với {totalRows - 1:N0} dòng (Khớp: {matchedCount:N0}, Không khớp: {unmatchedCount:N0})!";
                }
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"Lỗi thực thi ghép bảng: {ex.Message}";
            }
            finally
            {
                try
                {
                    app.Calculation = XlCalculation.xlCalculationAutomatic;
                    app.ScreenUpdating = true;
                }
                catch { }

                if (usedRange1 != null) Marshal.ReleaseComObject(usedRange1);
                if (usedRange2 != null) Marshal.ReleaseComObject(usedRange2);
                if (ws1 != null) Marshal.ReleaseComObject(ws1);
                if (ws2 != null) Marshal.ReleaseComObject(ws2);
            }

            return result;
        }

        private static string NormalizeKey(string key, bool trim, bool ignoreAccent)
        {
            if (string.IsNullOrEmpty(key)) return string.Empty;
            string s = trim ? key.Trim() : key;
            if (ignoreAccent)
            {
                s = RemoveDiacritics(s);
            }
            return s;
        }

        public static string RemoveDiacritics(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;
            var normalizedString = text.Normalize(System.Text.NormalizationForm.FormD);
            var stringBuilder = new System.Text.StringBuilder(capacity: normalizedString.Length);

            for (int i = 0; i < normalizedString.Length; i++)
            {
                char c = normalizedString[i];
                var unicodeCategory = System.Globalization.CharUnicodeInfo.GetUnicodeCategory(c);
                if (unicodeCategory != System.Globalization.UnicodeCategory.NonSpacingMark)
                {
                    stringBuilder.Append(c);
                }
            }

            return stringBuilder
                .ToString()
                .Normalize(System.Text.NormalizationForm.FormC)
                .Replace('đ', 'd')
                .Replace('Đ', 'D');
        }

        private static string ConvertColIndexToLetter(int colIndex)
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
    }
}
