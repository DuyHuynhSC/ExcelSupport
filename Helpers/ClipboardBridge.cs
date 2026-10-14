using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace ExcelSupport.Helpers
{
    /// <summary>
    /// Deep Module đại diện cho bộ trung chuyển Clipboard an toàn giữa Excel và Windows Clipboard.
    /// Đóng gói kín:
    /// 1. Cơ chế retry khi Windows Clipboard bị tiến trình khác lock (Anti-lock STA retry loop).
    /// 2. Hỗ trợ đa định dạng Unicode/Text đồng thời.
    /// 3. Trích xuất ma trận vùng ô hiển thị (Visible Cells / Non-contiguous Areas) chuẩn hóa.
    /// 4. Đọc/ghi dữ liệu có cấu trúc từ/vào Clipboard (TSV / Line-delimited).
    /// </summary>
    public static class ClipboardBridge
    {
        /// <summary>
        /// Ghi chuỗi văn bản vào Windows Clipboard kèm cơ chế thử lại tự động khi bị khóa bởi tiến trình khác.
        /// </summary>
        public static bool SetText(string text, int retries = 10, int delayMs = 50)
        {
            if (text == null) return false;

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

        /// <summary>
        /// Đọc chuỗi văn bản từ Windows Clipboard một cách an toàn.
        /// </summary>
        public static string GetText()
        {
            try
            {
                if (System.Windows.Forms.Clipboard.ContainsText())
                {
                    return System.Windows.Forms.Clipboard.GetText();
                }
                if (System.Windows.Clipboard.ContainsText())
                {
                    return System.Windows.Clipboard.GetText();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ClipboardBridge] GetText error: {ex.Message}");
            }
            return string.Empty;
        }

        /// <summary>
        /// Phân tích nội dung Clipboard thành bảng ma trận 2 chiều (tách theo dòng xuống hàng và cột qua Tab).
        /// </summary>
        public static List<List<object?>> ParseTableFromClipboard()
        {
            var list = new List<List<object?>>();
            string text = GetText();
            if (string.IsNullOrEmpty(text)) return list;

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
            return list;
        }

        /// <summary>
        /// Trích xuất dữ liệu ma trận từ một vùng ô (kể cả vùng chọn rời rạc nhiều Areas hoặc ô hiển thị Visible Only)
        /// thành danh sách các dòng và cột có thứ tự.
        /// </summary>
        public static (List<List<object?>> Values, List<List<string?>> Formulas, int TotalCells) ExtractRangeAreas(Range targetRange, bool includeFormulas = true)
        {
            var values = new List<List<object?>>();
            var formulas = new List<List<string?>>();
            if (targetRange == null) return (values, formulas, 0);

            var rowDictValues = new SortedDictionary<int, SortedDictionary<int, object?>>();
            var rowDictFormulas = new SortedDictionary<int, SortedDictionary<int, string?>>();
            int totalCells = 0;

            foreach (Range area in targetRange.Areas)
            {
                int rowCount = area.Rows.Count;
                int colCount = area.Columns.Count;
                int baseRow = area.Row;
                int baseCol = area.Column;

                object? rawValues = area.Value2;
                object? rawFormulas = includeFormulas ? area.Formula : null;

                object?[,]? valArray = rawValues as object[,];
                object?[,]? formulaArray = rawFormulas as object[,];

                for (int r = 1; r <= rowCount; r++)
                {
                    int actualRow = baseRow + r - 1;
                    if (!rowDictValues.ContainsKey(actualRow))
                    {
                        rowDictValues[actualRow] = new SortedDictionary<int, object?>();
                        if (includeFormulas)
                        {
                            rowDictFormulas[actualRow] = new SortedDictionary<int, string?>();
                        }
                    }

                    for (int c = 1; c <= colCount; c++)
                    {
                        int actualCol = baseCol + c - 1;
                        object? val = (valArray != null) ? valArray[r, c] : rawValues;
                        rowDictValues[actualRow][actualCol] = val;

                        if (includeFormulas)
                        {
                            string? formula = (formulaArray != null) ? formulaArray[r, c]?.ToString() : rawFormulas?.ToString();
                            rowDictFormulas[actualRow][actualCol] = formula;
                        }

                        totalCells++;
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
                values.Add(rowList);
            }

            if (includeFormulas)
            {
                foreach (var kvp in rowDictFormulas)
                {
                    var rowList = new List<string?>();
                    foreach (var colVal in kvp.Value)
                    {
                        rowList.Add(colVal.Value);
                    }
                    formulas.Add(rowList);
                }
            }

            return (values, formulas, totalCells);
        }
    }
}
