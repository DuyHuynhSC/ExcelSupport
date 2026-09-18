using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using ExcelSupport.Models;
using Microsoft.Office.Interop.Excel;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;

namespace ExcelSupport.Services
{
    public static class BatchFindReplaceService
    {
        /// <summary>
        /// Phân tích chuỗi văn bản thành danh sách cặp [Tìm -> Thay thế]
        /// Hỗ trợ định dạng phân cách: Tab, =>, ->, Dấu phẩy, Dấu chấm phẩy, Dấu gạch đứng (|)
        /// </summary>
        public static List<FindReplacePair> ParseDictionaryText(string rawText)
        {
            var list = new List<FindReplacePair>();
            if (string.IsNullOrWhiteSpace(rawText)) return list;

            var lines = rawText.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var line in lines)
            {
                var trimmed = line.Trim();
                if (string.IsNullOrEmpty(trimmed) || trimmed.StartsWith("#")) continue;

                string find = string.Empty;
                string replace = string.Empty;

                if (trimmed.Contains("=>"))
                {
                    var parts = trimmed.Split(new[] { "=>" }, 2, StringSplitOptions.None);
                    find = parts[0].Trim();
                    replace = parts.Length > 1 ? parts[1].Trim() : string.Empty;
                }
                else if (trimmed.Contains("->"))
                {
                    var parts = trimmed.Split(new[] { "->" }, 2, StringSplitOptions.None);
                    find = parts[0].Trim();
                    replace = parts.Length > 1 ? parts[1].Trim() : string.Empty;
                }
                else if (trimmed.Contains("\t"))
                {
                    var parts = trimmed.Split(new[] { '\t' }, 2, StringSplitOptions.None);
                    find = parts[0].Trim();
                    replace = parts.Length > 1 ? parts[1].Trim() : string.Empty;
                }
                else if (trimmed.Contains("|"))
                {
                    var parts = trimmed.Split(new[] { '|' }, 2, StringSplitOptions.None);
                    find = parts[0].Trim();
                    replace = parts.Length > 1 ? parts[1].Trim() : string.Empty;
                }
                else if (trimmed.Contains(";"))
                {
                    var parts = trimmed.Split(new[] { ';' }, 2, StringSplitOptions.None);
                    find = parts[0].Trim();
                    replace = parts.Length > 1 ? parts[1].Trim() : string.Empty;
                }
                else if (trimmed.Contains(","))
                {
                    var parts = trimmed.Split(new[] { ',' }, 2, StringSplitOptions.None);
                    find = parts[0].Trim();
                    replace = parts.Length > 1 ? parts[1].Trim() : string.Empty;
                }
                else
                {
                    find = trimmed;
                    replace = string.Empty; // Xóa từ
                }

                if (!string.IsNullOrEmpty(find) && !seen.Contains(find))
                {
                    seen.Add(find);
                    list.Add(new FindReplacePair { FindText = find, ReplaceText = replace });
                }
            }

            return list;
        }

        /// <summary>
        /// Đọc bảng từ điển từ 1 vùng chọn trên Excel (2 cột: Cột 1 = Từ cũ, Cột 2 = Từ mới)
        /// </summary>
        public static List<FindReplacePair> LoadDictionaryFromExcelRange(Range? range)
        {
            var list = new List<FindReplacePair>();
            if (range == null) return list;

            try
            {
                int rows = range.Rows.Count;
                int cols = range.Columns.Count;
                if (rows == 0) return list;

                object? rawVal = range.Value2;
                if (rawVal is object[,] val2D)
                {
                    var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    for (int r = 1; r <= rows; r++)
                    {
                        string find = val2D[r, 1]?.ToString()?.Trim() ?? string.Empty;
                        string replace = (cols >= 2) ? (val2D[r, 2]?.ToString() ?? string.Empty) : string.Empty;

                        if (!string.IsNullOrEmpty(find) && !seen.Contains(find))
                        {
                            seen.Add(find);
                            list.Add(new FindReplacePair { FindText = find, ReplaceText = replace });
                        }
                    }
                }
                else if (rawVal != null)
                {
                    string singleFind = rawVal.ToString()?.Trim() ?? string.Empty;
                    if (!string.IsNullOrEmpty(singleFind))
                    {
                        list.Add(new FindReplacePair { FindText = singleFind, ReplaceText = string.Empty });
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"LoadDictionaryFromExcelRange error: {ex.Message}");
            }

            return list;
        }

        /// <summary>
        /// Thực thi tìm và thay thế hàng loạt siêu tốc
        /// </summary>
        public static BatchFindReplaceResult ExecuteBatchReplace(ExcelApp app, BatchFindReplaceOptions options)
        {
            var result = new BatchFindReplaceResult();
            if (app == null || options == null || options.Pairs.Count == 0)
            {
                result.Success = false;
                result.Message = "Không có danh sách từ khóa tra cứu để thay thế.";
                return result;
            }

            try
            {
                app.ScreenUpdating = false;
                app.Calculation = XlCalculation.xlCalculationManual;

                var targetSheets = new List<_Worksheet>();

                switch (options.Scope)
                {
                    case FindReplaceScope.Selection:
                    case FindReplaceScope.ActiveSheet:
                        if (app.ActiveSheet is _Worksheet actWs) targetSheets.Add(actWs);
                        break;

                    case FindReplaceScope.AllSheetsCurrentWorkbook:
                        if (app.ActiveWorkbook != null)
                        {
                            foreach (_Worksheet ws in app.ActiveWorkbook.Worksheets)
                            {
                                targetSheets.Add(ws);
                            }
                        }
                        break;

                    case FindReplaceScope.AllOpenWorkbooks:
                        foreach (Workbook wb in app.Workbooks)
                        {
                            foreach (_Worksheet ws in wb.Worksheets)
                            {
                                targetSheets.Add(ws);
                            }
                        }
                        break;
                }

                if (targetSheets.Count == 0)
                {
                    result.Success = false;
                    result.Message = "Không tìm thấy bảng tính hợp lệ để thực thi.";
                    return result;
                }

                var pairCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                foreach (var p in options.Pairs) pairCounts[p.FindText] = 0;

                int totalCellsModified = 0;
                int sheetsModifiedCount = 0;
                int highlightColorOle = ColorTranslator.ToOle(options.HighlightColor);
                var comp = options.MatchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

                foreach (var ws in targetSheets)
                {
                    Range? targetRange = null;
                    bool isSelectionOnly = (options.Scope == FindReplaceScope.Selection && ws == app.ActiveSheet);

                    try
                    {
                        targetRange = isSelectionOnly ? (app.Selection as Range) : ws.UsedRange;
                        if (targetRange == null || targetRange.Rows.Count == 0) continue;

                        int numRows = targetRange.Rows.Count;
                        int numCols = targetRange.Columns.Count;
                        int startRow = targetRange.Row;
                        int startCol = targetRange.Column;

                        object? rawVal = (options.LookIn == FindReplaceLookIn.Formulas) ? targetRange.Formula : targetRange.Value2;
                        if (rawVal == null) continue;

                        bool sheetChanged = false;

                        if (rawVal is object[,] allVals)
                        {
                            var modifiedCellsToHighlight = new List<Range>();

                            for (int r = 1; r <= numRows; r++)
                            {
                                for (int c = 1; c <= numCols; c++)
                                {
                                    object? cellObj = allVals[r, c];
                                    if (cellObj == null) continue;
                                    string originalStr = cellObj.ToString() ?? string.Empty;
                                    if (string.IsNullOrEmpty(originalStr)) continue;

                                    string currentStr = originalStr;
                                    bool cellModified = false;

                                    foreach (var pair in options.Pairs)
                                    {
                                        if (options.MatchEntireCell)
                                        {
                                            if (string.Equals(currentStr, pair.FindText, comp))
                                            {
                                                currentStr = pair.ReplaceText;
                                                pairCounts[pair.FindText]++;
                                                cellModified = true;
                                            }
                                        }
                                        else
                                        {
                                            if (currentStr.IndexOf(pair.FindText, comp) >= 0)
                                            {
                                                int countBefore = (currentStr.Length - currentStr.Replace(pair.FindText, "").Length) / Math.Max(1, pair.FindText.Length);
                                                currentStr = ReplaceString(currentStr, pair.FindText, pair.ReplaceText, comp);
                                                pairCounts[pair.FindText] += countBefore;
                                                cellModified = true;
                                            }
                                        }
                                    }

                                    if (cellModified && currentStr != originalStr)
                                    {
                                        allVals[r, c] = currentStr;
                                        totalCellsModified++;
                                        sheetChanged = true;

                                        if (options.HighlightReplacedCells)
                                        {
                                            try
                                            {
                                                Range cellRange = ws.Cells[startRow + r - 1, startCol + c - 1];
                                                cellRange.Interior.Color = highlightColorOle;
                                                Marshal.ReleaseComObject(cellRange);
                                            }
                                            catch { }
                                        }
                                    }
                                }
                            }

                            if (sheetChanged)
                            {
                                sheetsModifiedCount++;
                                if (options.LookIn == FindReplaceLookIn.Formulas)
                                    targetRange.Formula = allVals;
                                else
                                    targetRange.Value2 = allVals;
                            }
                        }
                        else
                        {
                            // Đơn lẻ 1 ô
                            string originalStr = rawVal.ToString() ?? string.Empty;
                            string currentStr = originalStr;
                            bool cellModified = false;

                            foreach (var pair in options.Pairs)
                            {
                                if (options.MatchEntireCell)
                                {
                                    if (string.Equals(currentStr, pair.FindText, comp))
                                    {
                                        currentStr = pair.ReplaceText;
                                        pairCounts[pair.FindText]++;
                                        cellModified = true;
                                    }
                                }
                                else
                                {
                                    if (currentStr.IndexOf(pair.FindText, comp) >= 0)
                                    {
                                        currentStr = ReplaceString(currentStr, pair.FindText, pair.ReplaceText, comp);
                                        pairCounts[pair.FindText]++;
                                        cellModified = true;
                                    }
                                }
                            }

                            if (cellModified && currentStr != originalStr)
                            {
                                if (options.LookIn == FindReplaceLookIn.Formulas)
                                    targetRange.Formula = currentStr;
                                else
                                    targetRange.Value2 = currentStr;

                                totalCellsModified++;
                                sheetsModifiedCount++;

                                if (options.HighlightReplacedCells)
                                {
                                    try { targetRange.Interior.Color = highlightColorOle; } catch { }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Batch replace on sheet error: {ex.Message}");
                    }
                    finally
                    {
                        if (targetRange != null) Marshal.ReleaseComObject(targetRange);
                        Marshal.ReleaseComObject(ws);
                    }
                }

                int totalReplacements = 0;
                var pairResults = new List<FindReplacePair>();
                foreach (var p in options.Pairs)
                {
                    int cnt = pairCounts.TryGetValue(p.FindText, out int c) ? c : 0;
                    totalReplacements += cnt;
                    pairResults.Add(new FindReplacePair
                    {
                        FindText = p.FindText,
                        ReplaceText = p.ReplaceText,
                        MatchCount = cnt
                    });
                }

                result.Success = true;
                result.TotalReplacements = totalReplacements;
                result.TotalCellsModified = totalCellsModified;
                result.SheetsModified = sheetsModifiedCount;
                result.PairResults = pairResults;
                result.Message = $"Đã thay thế thành công {totalReplacements:N0} lần trên {totalCellsModified:N0} ô ({sheetsModifiedCount:N0} Sheet)!";
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"Lỗi thực thi thay thế: {ex.Message}";
            }
            finally
            {
                try
                {
                    app.Calculation = XlCalculation.xlCalculationAutomatic;
                    app.ScreenUpdating = true;
                }
                catch { }
            }

            return result;
        }

        private static string ReplaceString(string str, string oldValue, string newValue, StringComparison comp)
        {
            if (string.IsNullOrEmpty(str) || string.IsNullOrEmpty(oldValue)) return str;

            int index = str.IndexOf(oldValue, comp);
            if (index < 0) return str;

            var sb = new System.Text.StringBuilder();
            int previousIndex = 0;

            while (index >= 0)
            {
                sb.Append(str.Substring(previousIndex, index - previousIndex));
                sb.Append(newValue);
                previousIndex = index + oldValue.Length;
                index = str.IndexOf(oldValue, previousIndex, comp);
            }

            sb.Append(str.Substring(previousIndex));
            return sb.ToString();
        }
    }
}
