using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using ExcelSupport.Models;
using Microsoft.Office.Interop.Excel;
using SysDataTable = System.Data.DataTable;

namespace ExcelSupport.Services
{
    public static class AdvancedFilterService
    {
        private static readonly Regex VietnameseCharRegex = new Regex(@"[àáạảãâầấậẩẫăằắặẳẵèéẹẻẽêềếệểễìíịỉĩòóọỏõôồốộổỗơờớợởỡùúụủũưừứựửữỳýỵỷỹđÀÁẠẢÃÂẦẤẬẨẪĂẰẮẶẲẴÈÉẸẺẼÊỀẾỆỂỄÌÍỊỈĨÒÓỌỎÕÔỒỐỘỔỖƠỜỚỢỞỠÙÚỤỦŨƯỪỨỰỬỮỲÝỴỶỸĐ]", RegexOptions.Compiled);

        #region 1. Phân Tách Danh Sách & Biểu Thức

        /// <summary>
        /// Phân tách danh sách paste từ clipboard theo các dấu phân cách thông dụng
        /// </summary>
        public static List<string> ParseBatchList(string? rawText)
        {
            if (string.IsNullOrWhiteSpace(rawText)) return new List<string>();

            char[] separators = { '\r', '\n', '\t', ',', ';', '|' };
            var tokens = rawText!.Split(separators, StringSplitOptions.RemoveEmptyEntries);

            var result = new List<string>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var token in tokens)
            {
                string clean = token.Trim();
                if (!string.IsNullOrEmpty(clean) && seen.Add(clean))
                {
                    result.Add(clean);
                }
            }

            return result;
        }

        /// <summary>
        /// Phân tích cú pháp biểu thức nhanh thành cấu trúc AdvancedFilterCriteria
        /// Ví dụ: "(> 0 and < 50) or > 250" hoặc "(>= 100 and <= 500) or == 0"
        /// </summary>
        public static AdvancedFilterCriteria ParseQuickExpression(string expression, int targetColumnIndex, string targetColumnName = "")
        {
            var criteria = new AdvancedFilterCriteria { OuterOperator = LogicalOperator.Or };
            if (string.IsNullOrWhiteSpace(expression)) return criteria;

            string clean = expression.Trim();

            // Tách theo "OR" hoặc "hoặc"
            var groupTokens = Regex.Split(clean, @"\s+(?:OR|hoặc|\|\|)\s+", RegexOptions.IgnoreCase);

            int groupNum = 1;
            foreach (var groupToken in groupTokens)
            {
                string gText = groupToken.Trim().Trim('(', ')');
                if (string.IsNullOrWhiteSpace(gText)) continue;

                var ruleGroup = new FilterRuleGroup
                {
                    GroupTitle = $"Nhóm {groupNum++}",
                    InnerOperator = LogicalOperator.And
                };

                // Tách các điều kiện bên trong theo "AND" hoặc "và"
                var ruleTokens = Regex.Split(gText, @"\s+(?:AND|và|&&)\s+", RegexOptions.IgnoreCase);

                foreach (var rToken in ruleTokens)
                {
                    string ruleStr = rToken.Trim();
                    var rule = ParseSingleRule(ruleStr, targetColumnIndex, targetColumnName);
                    if (rule != null)
                    {
                        ruleGroup.Rules.Add(rule);
                    }
                }

                if (ruleGroup.Rules.Count > 0)
                {
                    criteria.Groups.Add(ruleGroup);
                }
            }

            return criteria;
        }

        private static FilterRule? ParseSingleRule(string ruleStr, int defaultColIndex, string defaultColName)
        {
            if (string.IsNullOrWhiteSpace(ruleStr)) return null;

            string s = ruleStr.Trim();

            // Bỏ chữ "x" hoặc "X" hoặc [Tên cột]
            if (s.StartsWith("x ", StringComparison.OrdinalIgnoreCase) || s.StartsWith("x=", StringComparison.OrdinalIgnoreCase) ||
                s.StartsWith("x>", StringComparison.OrdinalIgnoreCase) || s.StartsWith("x<", StringComparison.OrdinalIgnoreCase) ||
                s.StartsWith("x!", StringComparison.OrdinalIgnoreCase))
            {
                s = s.Substring(1).Trim();
            }

            FilterOperator op = FilterOperator.Equals;
            string val1 = string.Empty;
            string val2 = string.Empty;

            if (s.StartsWith(">=")) { op = FilterOperator.GreaterThanOrEqual; val1 = s.Substring(2).Trim(); }
            else if (s.StartsWith("<=")) { op = FilterOperator.LessThanOrEqual; val1 = s.Substring(2).Trim(); }
            else if (s.StartsWith("!=")) { op = FilterOperator.NotEquals; val1 = s.Substring(2).Trim(); }
            else if (s.StartsWith("<>")) { op = FilterOperator.NotEquals; val1 = s.Substring(2).Trim(); }
            else if (s.StartsWith("==")) { op = FilterOperator.Equals; val1 = s.Substring(2).Trim(); }
            else if (s.StartsWith("=")) { op = FilterOperator.Equals; val1 = s.Substring(1).Trim(); }
            else if (s.StartsWith(">")) { op = FilterOperator.GreaterThan; val1 = s.Substring(1).Trim(); }
            else if (s.StartsWith("<")) { op = FilterOperator.LessThan; val1 = s.Substring(1).Trim(); }
            else if (s.StartsWith("contains(", StringComparison.OrdinalIgnoreCase) && s.EndsWith(")"))
            {
                op = FilterOperator.Contains;
                val1 = s.Substring(9, s.Length - 10).Trim('"', '\'');
            }
            else if (s.StartsWith("startswith(", StringComparison.OrdinalIgnoreCase) && s.EndsWith(")"))
            {
                op = FilterOperator.StartsWith;
                val1 = s.Substring(11, s.Length - 12).Trim('"', '\'');
            }
            else if (s.StartsWith("endswith(", StringComparison.OrdinalIgnoreCase) && s.EndsWith(")"))
            {
                op = FilterOperator.EndsWith;
                val1 = s.Substring(9, s.Length - 10).Trim('"', '\'');
            }
            else if (s.StartsWith("regex(", StringComparison.OrdinalIgnoreCase) && s.EndsWith(")"))
            {
                op = FilterOperator.MatchesRegex;
                val1 = s.Substring(6, s.Length - 7).Trim('"', '\'');
            }
            else
            {
                // Mặc định là so sánh bằng
                op = FilterOperator.Equals;
                val1 = s.Trim('"', '\'');
            }

            return new FilterRule
            {
                ColumnIndex = defaultColIndex,
                ColumnName = defaultColName,
                Operator = op,
                Value1 = val1,
                Value2 = val2
            };
        }

        #endregion

        #region 2. Đọc Thông Tin Cột Từ Sheet

        public static List<ColumnHeaderItem> GetSheetColumns(_Worksheet ws)
        {
            var list = new List<ColumnHeaderItem>();
            if (ws == null) return list;

            Range? usedRange = null;
            try
            {
                usedRange = ws.UsedRange;
                if (usedRange == null) return list;

                int startCol = usedRange.Column;
                int colCount = usedRange.Columns.Count;
                int startRow = usedRange.Row;

                for (int c = 0; c < colCount; c++)
                {
                    int colIdx = startCol + c;
                    string colLetter = GetColumnLetter(colIdx);
                    string headerText = string.Empty;

                    Range? headerCell = null;
                    try
                    {
                        headerCell = ws.Cells[startRow, colIdx] as Range;
                        headerText = headerCell?.Value2?.ToString()?.Trim() ?? string.Empty;
                    }
                    catch { }
                    finally
                    {
                        if (headerCell != null) Marshal.ReleaseComObject(headerCell);
                    }

                    list.Add(new ColumnHeaderItem
                    {
                        ColumnIndex = colIdx,
                        ColumnLetter = colLetter,
                        HeaderText = headerText
                    });
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"GetSheetColumns error: {ex.Message}");
            }
            finally
            {
                if (usedRange != null) Marshal.ReleaseComObject(usedRange);
            }

            return list;
        }

        public static string GetColumnLetter(int col)
        {
            int dividend = col;
            string colName = string.Empty;
            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                colName = Convert.ToChar(65 + modulo) + colName;
                dividend = (dividend - modulo) / 26;
            }
            return colName;
        }

        #endregion

        #region 3. Kiểm Tra & Đánh Giá Dữ Liệu (Evaluation Engine)

        /// <summary>
        /// Đánh giá 1 ô dữ liệu với 1 quy tắc đơn lẻ
        /// </summary>
        public static bool EvaluateCell(object? cellValue, FilterRule rule)
        {
            if (rule == null) return true;

            string cellStr = cellValue?.ToString() ?? string.Empty;
            bool isCellEmpty = string.IsNullOrWhiteSpace(cellStr);

            switch (rule.Operator)
            {
                case FilterOperator.IsEmpty:
                    return isCellEmpty;

                case FilterOperator.IsNotEmpty:
                    return !isCellEmpty;

                case FilterOperator.ContainsVietnamese:
                    return !isCellEmpty && VietnameseCharRegex.IsMatch(cellStr);

                case FilterOperator.IsEven:
                    if (double.TryParse(cellStr, NumberStyles.Any, CultureInfo.InvariantCulture, out double numEven) ||
                        double.TryParse(cellStr, NumberStyles.Any, CultureInfo.CurrentCulture, out numEven))
                    {
                        return Math.Abs(numEven % 2) < 0.0001;
                    }
                    return false;

                case FilterOperator.IsOdd:
                    if (double.TryParse(cellStr, NumberStyles.Any, CultureInfo.InvariantCulture, out double numOdd) ||
                        double.TryParse(cellStr, NumberStyles.Any, CultureInfo.CurrentCulture, out numOdd))
                    {
                        return Math.Abs(numOdd % 2) > 0.0001;
                    }
                    return false;

                case FilterOperator.MatchesRegex:
                    if (string.IsNullOrEmpty(rule.Value1)) return true;
                    try
                    {
                        var regexOpts = rule.MatchCase ? RegexOptions.None : RegexOptions.IgnoreCase;
                        return Regex.IsMatch(cellStr, rule.Value1, regexOpts);
                    }
                    catch { return false; }

                case FilterOperator.Contains:
                    if (string.IsNullOrEmpty(rule.Value1)) return true;
                    return cellStr.IndexOf(rule.Value1, rule.MatchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase) >= 0;

                case FilterOperator.NotContains:
                    if (string.IsNullOrEmpty(rule.Value1)) return true;
                    return cellStr.IndexOf(rule.Value1, rule.MatchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase) < 0;

                case FilterOperator.StartsWith:
                    if (string.IsNullOrEmpty(rule.Value1)) return true;
                    return cellStr.StartsWith(rule.Value1, rule.MatchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase);

                case FilterOperator.EndsWith:
                    if (string.IsNullOrEmpty(rule.Value1)) return true;
                    return cellStr.EndsWith(rule.Value1, rule.MatchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase);

                case FilterOperator.Today:
                    if (DateTime.TryParse(cellStr, out DateTime dtToday))
                    {
                        return dtToday.Date == DateTime.Today;
                    }
                    return false;

                case FilterOperator.ThisMonth:
                    if (DateTime.TryParse(cellStr, out DateTime dtMonth))
                    {
                        return dtMonth.Year == DateTime.Today.Year && dtMonth.Month == DateTime.Today.Month;
                    }
                    return false;

                case FilterOperator.ThisYear:
                    if (DateTime.TryParse(cellStr, out DateTime dtYear))
                    {
                        return dtYear.Year == DateTime.Today.Year;
                    }
                    return false;
            }

            // Xử lý so sánh số
            bool isCellNum = double.TryParse(cellStr, NumberStyles.Any, CultureInfo.InvariantCulture, out double cNum) ||
                             double.TryParse(cellStr, NumberStyles.Any, CultureInfo.CurrentCulture, out cNum);
            bool isRule1Num = double.TryParse(rule.Value1, NumberStyles.Any, CultureInfo.InvariantCulture, out double r1Num) ||
                              double.TryParse(rule.Value1, NumberStyles.Any, CultureInfo.CurrentCulture, out r1Num);
            bool isRule2Num = double.TryParse(rule.Value2, NumberStyles.Any, CultureInfo.InvariantCulture, out double r2Num) ||
                              double.TryParse(rule.Value2, NumberStyles.Any, CultureInfo.CurrentCulture, out r2Num);

            if (isCellNum && isRule1Num)
            {
                switch (rule.Operator)
                {
                    case FilterOperator.GreaterThan:
                        return cNum > r1Num;
                    case FilterOperator.GreaterThanOrEqual:
                        return cNum >= r1Num;
                    case FilterOperator.LessThan:
                        return cNum < r1Num;
                    case FilterOperator.LessThanOrEqual:
                        return cNum <= r1Num;
                    case FilterOperator.Equals:
                        return Math.Abs(cNum - r1Num) < 0.000001;
                    case FilterOperator.NotEquals:
                        return Math.Abs(cNum - r1Num) >= 0.000001;
                    case FilterOperator.Between:
                        if (isRule2Num) return cNum >= Math.Min(r1Num, r2Num) && cNum <= Math.Max(r1Num, r2Num);
                        return cNum >= r1Num;
                    case FilterOperator.NotBetween:
                        if (isRule2Num) return cNum < Math.Min(r1Num, r2Num) || cNum > Math.Max(r1Num, r2Num);
                        return cNum < r1Num;
                }
            }

            // Xử lý so sánh chuỗi
            var strComp = rule.MatchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
            switch (rule.Operator)
            {
                case FilterOperator.Equals:
                    return string.Equals(cellStr, rule.Value1, strComp);
                case FilterOperator.NotEquals:
                    return !string.Equals(cellStr, rule.Value1, strComp);
                case FilterOperator.Between:
                    return string.Compare(cellStr, rule.Value1, strComp) >= 0 && string.Compare(cellStr, rule.Value2, strComp) <= 0;
                case FilterOperator.NotBetween:
                    return string.Compare(cellStr, rule.Value1, strComp) < 0 || string.Compare(cellStr, rule.Value2, strComp) > 0;
                case FilterOperator.GreaterThan:
                    return string.Compare(cellStr, rule.Value1, strComp) > 0;
                case FilterOperator.GreaterThanOrEqual:
                    return string.Compare(cellStr, rule.Value1, strComp) >= 0;
                case FilterOperator.LessThan:
                    return string.Compare(cellStr, rule.Value1, strComp) < 0;
                case FilterOperator.LessThanOrEqual:
                    return string.Compare(cellStr, rule.Value1, strComp) <= 0;
            }

            return false;
        }

        /// <summary>
        /// Đánh giá 1 dòng dữ liệu trong mảng 2D với tiêu chí đa điều kiện AdvancedFilterCriteria
        /// </summary>
        public static bool EvaluateRow(object[,] values2D, int rowIdx, int startCol, AdvancedFilterCriteria criteria)
        {
            if (criteria == null || criteria.Groups.Count == 0) return true;

            int totalCols = values2D.GetLength(1);

            bool outerResult = criteria.OuterOperator == LogicalOperator.And;

            foreach (var group in criteria.Groups)
            {
                if (group.Rules.Count == 0) continue;

                bool groupResult = group.InnerOperator == LogicalOperator.And;

                foreach (var rule in group.Rules)
                {
                    int colOffset = rule.ColumnIndex - startCol + 1;
                    object? cellVal = (colOffset >= 1 && colOffset <= totalCols) ? values2D[rowIdx, colOffset] : null;

                    bool match = EvaluateCell(cellVal, rule);

                    if (group.InnerOperator == LogicalOperator.And)
                    {
                        groupResult = groupResult && match;
                        if (!groupResult) break; // Short-circuit AND
                    }
                    else // OR
                    {
                        groupResult = groupResult || match;
                        if (groupResult) break; // Short-circuit OR
                    }
                }

                if (criteria.OuterOperator == LogicalOperator.Or)
                {
                    if (groupResult) return true; // Short-circuit OR
                }
                else // AND
                {
                    if (!groupResult) return false; // Short-circuit AND
                }
            }

            return criteria.OuterOperator == LogicalOperator.Or ? false : true;
        }

        /// <summary>
        /// Đánh giá 1 dòng dữ liệu trong mảng 2D với tiêu chí lọc danh sách Paste (Batch List)
        /// </summary>
        public static bool EvaluateBatchListRow(object[,] values2D, int rowIdx, int targetCol, int startCol, BatchListFilterCriteria criteria, HashSet<string>? setExact, List<string>? itemsList)
        {
            if (criteria == null || criteria.ParsedItems.Count == 0) return true;

            int colOffset = targetCol - startCol + 1;
            int totalCols = values2D.GetLength(1);
            object? cellVal = (colOffset >= 1 && colOffset <= totalCols) ? values2D[rowIdx, colOffset] : null;
            string cellStr = cellVal?.ToString()?.Trim() ?? string.Empty;

            bool isMatch = false;

            if (criteria.IsExactMatch)
            {
                if (setExact != null)
                {
                    isMatch = setExact.Contains(cellStr);
                }
            }
            else
            {
                // So khớp Contains
                var comp = criteria.MatchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
                var list = itemsList ?? criteria.ParsedItems;
                foreach (var item in list)
                {
                    if (cellStr.IndexOf(item, comp) >= 0)
                    {
                        isMatch = true;
                        break;
                    }
                }
            }

            // Nếu là Blacklist (ExcludeList), đảo ngược kết quả
            return criteria.ExcludeList ? !isMatch : isMatch;
        }

        #endregion

        #region 4. Thực Thi Hành Động Lọc Trên Excel (COM Interop)

        /// <summary>
        /// Lọc trực tiếp trên Sheet bằng cách ẩn các dòng không thỏa (In-Place Row Hiding theo Contiguous Blocks)
        /// </summary>
        public static FilterExecutionResult ApplyInPlaceFilter(_Worksheet ws, Func<int, bool> rowMatcher)
        {
            var result = new FilterExecutionResult();
            if (ws == null || rowMatcher == null) return result;

            Range? usedRange = null;
            try
            {
                usedRange = ws.UsedRange;
                if (usedRange == null) return result;

                int startRow = usedRange.Row;
                int totalRows = usedRange.Rows.Count;

                // Bỏ qua dòng tiêu đề đầu tiên nếu có nhiều hơn 1 dòng
                int dataStartRow = startRow + 1;
                int dataRowCount = totalRows - 1;

                if (dataRowCount <= 0)
                {
                    result.TotalRows = 0;
                    result.MatchedRows = 0;
                    result.Success = true;
                    return result;
                }

                result.TotalRows = dataRowCount;
                int matchedCount = 0;

                ws.Application.ScreenUpdating = false;

                // Hiện lại toàn bộ trước khi áp dụng bộ lọc mới
                try { ws.Rows.Hidden = false; } catch { }

                int blockStart = -1;

                for (int i = 0; i < dataRowCount; i++)
                {
                    int actualRow = dataStartRow + i;
                    bool isMatch = rowMatcher(i + 2); // 1-based index tương ứng với mảng 2D (dòng 1 là tiêu đề)

                    if (isMatch)
                    {
                        matchedCount++;
                        // Kết thúc dải dòng cần ẩn nếu có
                        if (blockStart != -1)
                        {
                            HideRowRange(ws, blockStart, actualRow - 1);
                            blockStart = -1;
                        }
                    }
                    else
                    {
                        // Dòng không thỏa -> Đánh dấu để ẩn
                        if (blockStart == -1)
                        {
                            blockStart = actualRow;
                        }
                    }
                }

                // Ẩn dải cuối cùng
                if (blockStart != -1)
                {
                    HideRowRange(ws, blockStart, dataStartRow + dataRowCount - 1);
                }

                result.MatchedRows = matchedCount;
                result.Success = true;
                result.Message = $"Đã lọc thành công: Hiển thị {matchedCount:N0} / {dataRowCount:N0} dòng ({result.MatchPercentage:F1}%)";
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"Lỗi thực thi lọc: {ex.Message}";
                System.Diagnostics.Debug.WriteLine($"ApplyInPlaceFilter error: {ex.Message}");
            }
            finally
            {
                try { if (ws != null) ws.Application.ScreenUpdating = true; } catch { }
                if (usedRange != null) Marshal.ReleaseComObject(usedRange);
            }

            return result;
        }

        private static void HideRowRange(_Worksheet ws, int fromRow, int toRow)
        {
            Range? r = null;
            try
            {
                r = ws.Range[$"A{fromRow}:A{toRow}"];
                if (r != null)
                {
                    r.EntireRow.Hidden = true;
                }
            }
            catch { }
            finally
            {
                if (r != null) Marshal.ReleaseComObject(r);
            }
        }

        /// <summary>
        /// Hiện lại toàn bộ các dòng trên Sheet (Clear Filter)
        /// </summary>
        public static bool ClearFilter(_Worksheet ws)
        {
            if (ws == null) return false;
            try
            {
                ws.Application.ScreenUpdating = false;
                ws.Rows.Hidden = false;
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ClearFilter error: {ex.Message}");
                return false;
            }
            finally
            {
                try { ws.Application.ScreenUpdating = true; } catch { }
            }
        }

        /// <summary>
        /// Tô màu nền Highlight các dòng thỏa mãn
        /// </summary>
        public static int HighlightMatchingRows(_Worksheet ws, Func<int, bool> rowMatcher, System.Drawing.Color color)
        {
            if (ws == null || rowMatcher == null) return 0;
            int highlighted = 0;

            Range? usedRange = null;
            try
            {
                usedRange = ws.UsedRange;
                if (usedRange == null) return 0;

                int startRow = usedRange.Row;
                int totalRows = usedRange.Rows.Count;
                int startCol = usedRange.Column;
                int totalCols = usedRange.Columns.Count;

                int dataStartRow = startRow + 1;
                int dataRowCount = totalRows - 1;

                if (dataRowCount <= 0) return 0;

                ws.Application.ScreenUpdating = false;
                int oleColor = System.Drawing.ColorTranslator.ToOle(color);

                for (int i = 0; i < dataRowCount; i++)
                {
                    int actualRow = dataStartRow + i;
                    if (rowMatcher(i + 2))
                    {
                        Range? rowRange = null;
                        try
                        {
                            rowRange = ws.Range[ws.Cells[actualRow, startCol], ws.Cells[actualRow, startCol + totalCols - 1]];
                            rowRange.Interior.Color = oleColor;
                            highlighted++;
                        }
                        catch { }
                        finally
                        {
                            if (rowRange != null) Marshal.ReleaseComObject(rowRange);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"HighlightMatchingRows error: {ex.Message}");
            }
            finally
            {
                try { if (ws != null) ws.Application.ScreenUpdating = true; } catch { }
                if (usedRange != null) Marshal.ReleaseComObject(usedRange);
            }

            return highlighted;
        }

        /// <summary>
        /// Trích xuất toàn bộ dòng thỏa mãn sang một Sheet mới
        /// </summary>
        public static bool ExtractMatchingRowsToNewSheet(Workbook wb, _Worksheet ws, Func<int, bool> rowMatcher, out int extractedRowsCount)
        {
            extractedRowsCount = 0;
            if (wb == null || ws == null || rowMatcher == null) return false;

            Range? usedRange = null;
            _Worksheet? newSheet = null;

            try
            {
                usedRange = ws.UsedRange;
                if (usedRange == null) return false;

                int totalRows = usedRange.Rows.Count;
                int totalCols = usedRange.Columns.Count;
                if (totalRows <= 1) return false;

                object[,] allValues = (object[,])usedRange.Value2;

                var matchedRowIndices = new List<int>();
                for (int r = 2; r <= totalRows; r++)
                {
                    if (rowMatcher(r))
                    {
                        matchedRowIndices.Add(r);
                    }
                }

                if (matchedRowIndices.Count == 0) return false;

                extractedRowsCount = matchedRowIndices.Count;

                // Tạo mảng trích xuất gồm dòng tiêu đề + các dòng thỏa
                object[,] extractedData = new object[matchedRowIndices.Count + 1, totalCols];

                // Copy tiêu đề (dòng 1)
                for (int c = 1; c <= totalCols; c++)
                {
                    extractedData[1, c] = allValues[1, c];
                }

                // Copy các dòng thỏa mãn
                for (int i = 0; i < matchedRowIndices.Count; i++)
                {
                    int srcRow = matchedRowIndices[i];
                    for (int c = 1; c <= totalCols; c++)
                    {
                        extractedData[i + 2, c] = allValues[srcRow, c];
                    }
                }

                // Tạo sheet mới
                string baseName = "TrichXuat_Loc";
                string sheetName = baseName;
                int suffix = 1;

                while (true)
                {
                    bool exists = false;
                    foreach (_Worksheet s in wb.Worksheets)
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

                newSheet = wb.Worksheets.Add() as _Worksheet;
                if (newSheet == null) return false;
                newSheet.Name = sheetName;

                Range destRange = newSheet.Range[newSheet.Cells[1, 1], newSheet.Cells[matchedRowIndices.Count + 1, totalCols]];
                destRange.Value2 = extractedData;

                // Định dạng tiêu đề nổi bật
                Range headerRange = newSheet.Range[newSheet.Cells[1, 1], newSheet.Cells[1, totalCols]];
                headerRange.Font.Bold = true;
                headerRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(16, 124, 65));
                headerRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                Marshal.ReleaseComObject(headerRange);

                newSheet.Columns.AutoFit();
                Marshal.ReleaseComObject(destRange);

                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ExtractMatchingRowsToNewSheet error: {ex.Message}");
                return false;
            }
            finally
            {
                if (newSheet != null) Marshal.ReleaseComObject(newSheet);
                if (usedRange != null) Marshal.ReleaseComObject(usedRange);
            }
        }

        #endregion

        #region 5. Xem Trước Trực Tiếp (Live Preview)

        public static (SysDataTable PreviewTable, int TotalCount, int MatchedCount) GetPreviewData(_Worksheet ws, Func<int, bool> rowMatcher, int maxPreviewRows = 12)
        {
            var dt = new SysDataTable();
            if (ws == null || rowMatcher == null) return (dt, 0, 0);

            Range? usedRange = null;
            try
            {
                usedRange = ws.UsedRange;
                if (usedRange == null) return (dt, 0, 0);

                int totalRows = usedRange.Rows.Count;
                int totalCols = usedRange.Columns.Count;
                int startCol = usedRange.Column;

                if (totalRows <= 1) return (dt, 0, 0);

                object[,] allValues = (object[,])usedRange.Value2;

                // Thêm cột số thứ tự dòng
                dt.Columns.Add("Dòng", typeof(int));

                // Thêm các cột dữ liệu
                for (int c = 1; c <= totalCols; c++)
                {
                    string colLetter = GetColumnLetter(startCol + c - 1);
                    string header = allValues[1, c]?.ToString()?.Trim() ?? string.Empty;
                    string colName = string.IsNullOrEmpty(header) ? $"Cột {colLetter}" : $"{colLetter}: {header}";
                    dt.Columns.Add(colName, typeof(string));
                }

                int matchedCount = 0;

                for (int r = 2; r <= totalRows; r++)
                {
                    if (rowMatcher(r))
                    {
                        matchedCount++;
                        if (dt.Rows.Count < maxPreviewRows)
                        {
                            var row = dt.NewRow();
                            row["Dòng"] = r;
                            for (int c = 1; c <= totalCols; c++)
                            {
                                row[c] = allValues[r, c]?.ToString() ?? string.Empty;
                            }
                            dt.Rows.Add(row);
                        }
                    }
                }

                return (dt, totalRows - 1, matchedCount);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"GetPreviewData error: {ex.Message}");
                return (dt, 0, 0);
            }
            finally
            {
                if (usedRange != null) Marshal.ReleaseComObject(usedRange);
            }
        }

        #endregion
    }
}
