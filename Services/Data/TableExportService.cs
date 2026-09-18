using System;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;

namespace ExcelSupport.Services
{
    public class TableExportOptions
    {
        public bool FirstRowAsHeader { get; set; } = true;
        public bool AlignNumbersRight { get; set; } = true;
        public bool CompactFormat { get; set; } = false;
        public bool IncludeHtmlStyles { get; set; } = true;
        public bool ConvertLineBreaksToBr { get; set; } = true;
    }

    public static class TableExportService
    {
        /// <summary>
        /// Sao chép nhanh vùng chọn hiện tại trong Excel sang định dạng Markdown Table vào Clipboard (Ctrl + Shift + M).
        /// </summary>
        public static bool QuickCopySelectionToMarkdown(ExcelApp app)
        {
            if (app == null) return false;

            try
            {
                dynamic sel = app.Selection;
                if (sel is not Range rng)
                {
                    app.StatusBar = "⚠️ ExcelSupport: Vui lòng chọn một vùng ô trên bảng tính để sao chép bảng Markdown!";
                    return false;
                }

                string markdown = RangeToMarkdown(rng, new TableExportOptions());
                if (!string.IsNullOrEmpty(markdown))
                {
                    Clipboard.SetText(markdown);
                    app.StatusBar = $"📋 ExcelSupport: Đã sao chép bảng Markdown ({rng.Rows.Count}x{rng.Columns.Count}) vào Clipboard (Ctrl + Shift + M)!";
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                try { app.StatusBar = $"❌ ExcelSupport: Lỗi xuất Markdown: {ex.Message}"; } catch { }
                return false;
            }
        }

        public static string RangeToMarkdown(Range rng, TableExportOptions options)
        {
            if (rng == null) return string.Empty;

            object[,] values = GetRangeValues2D(rng);
            int rows = values.GetLength(0);
            int cols = values.GetLength(1);

            if (rows == 0 || cols == 0) return string.Empty;

            // Xác định kiểu dữ liệu của từng cột để căn lề (Left / Right)
            bool[] isNumberCol = new bool[cols];
            int[] maxColWidths = new int[cols];

            for (int c = 0; c < cols; c++)
            {
                isNumberCol[c] = options.AlignNumbersRight;
                maxColWidths[c] = 3; // Độ rộng tối thiểu
            }

            // Chuẩn bị ma trận chuỗi
            string[,] strGrid = new string[rows, cols];

            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    object? val = values[r + 1, c + 1];
                    string cellText = val?.ToString() ?? string.Empty;

                    // Thoát ký tự pipe '|' và xử lý xuống dòng
                    cellText = cellText.Replace("|", "\\|");
                    if (options.ConvertLineBreaksToBr)
                    {
                        cellText = cellText.Replace("\r\n", "<br>").Replace("\n", "<br>").Replace("\r", "<br>");
                    }
                    else
                    {
                        cellText = cellText.Replace("\r\n", " ").Replace("\n", " ").Replace("\r", " ");
                    }

                    cellText = cellText.Trim();
                    strGrid[r, c] = cellText;

                    if (cellText.Length > maxColWidths[c])
                    {
                        maxColWidths[c] = cellText.Length;
                    }

                    // Kiểm tra xem cột có phải toàn số không (bỏ qua dòng tiêu đề)
                    if (options.AlignNumbersRight && r > 0 && !string.IsNullOrEmpty(cellText))
                    {
                        if (!double.TryParse(cellText, out _))
                        {
                            isNumberCol[c] = false;
                        }
                    }
                }
            }

            var sb = new StringBuilder();

            int headerRowIndex = 0;
            // 1. Dòng Header
            sb.Append("|");
            for (int c = 0; c < cols; c++)
            {
                string text = options.FirstRowAsHeader ? strGrid[headerRowIndex, c] : $"Cột {c + 1}";
                if (options.CompactFormat)
                {
                    sb.Append($" {text} |");
                }
                else
                {
                    sb.Append($" {text.PadRight(maxColWidths[c])} |");
                }
            }
            sb.AppendLine();

            // 2. Dòng phân cách (Separator với căn lề)
            sb.Append("|");
            for (int c = 0; c < cols; c++)
            {
                if (options.CompactFormat)
                {
                    sb.Append(isNumberCol[c] ? " ---: |" : " :--- |");
                }
                else
                {
                    int width = Math.Max(3, maxColWidths[c]);
                    if (isNumberCol[c])
                    {
                        sb.Append($" {new string('-', width - 1)}: |");
                    }
                    else
                    {
                        sb.Append($" :{new string('-', width - 1)} |");
                    }
                }
            }
            sb.AppendLine();

            // 3. Các dòng dữ liệu
            int startDataRow = options.FirstRowAsHeader ? 1 : 0;
            for (int r = startDataRow; r < rows; r++)
            {
                sb.Append("|");
                for (int c = 0; c < cols; c++)
                {
                    string text = strGrid[r, c];
                    if (options.CompactFormat)
                    {
                        sb.Append($" {text} |");
                    }
                    else
                    {
                        if (isNumberCol[c])
                        {
                            sb.Append($" {text.PadLeft(maxColWidths[c])} |");
                        }
                        else
                        {
                            sb.Append($" {text.PadRight(maxColWidths[c])} |");
                        }
                    }
                }
                sb.AppendLine();
            }

            return sb.ToString();
        }

        public static string RangeToHtml(Range rng, TableExportOptions options)
        {
            if (rng == null) return string.Empty;

            object[,] values = GetRangeValues2D(rng);
            int rows = values.GetLength(0);
            int cols = values.GetLength(1);

            if (rows == 0 || cols == 0) return string.Empty;

            var sb = new StringBuilder();

            string tableStyle = options.IncludeHtmlStyles
                ? "style=\"border-collapse: collapse; width: 100%; font-family: Segoe UI, sans-serif; font-size: 13px; border: 1px solid #cbd5e1;\""
                : "border=\"1\"";

            string thStyle = options.IncludeHtmlStyles
                ? "style=\"background-color: #f1f5f9; color: #1e293b; font-weight: 600; padding: 8px 12px; border: 1px solid #cbd5e1; text-align: left;\""
                : "";

            string tdStyle = options.IncludeHtmlStyles
                ? "style=\"padding: 6px 12px; border: 1px solid #e2e8f0; color: #334155;\""
                : "";

            string tdNumStyle = options.IncludeHtmlStyles
                ? "style=\"padding: 6px 12px; border: 1px solid #e2e8f0; color: #334155; text-align: right;\""
                : "align=\"right\"";

            sb.AppendLine($"<table {tableStyle}>");

            int startDataRow = 1;

            if (options.FirstRowAsHeader)
            {
                sb.AppendLine("  <thead>");
                sb.AppendLine("    <tr>");
                for (int c = 1; c <= cols; c++)
                {
                    string text = FormatHtmlCell(values[1, c], options);
                    sb.AppendLine($"      <th {thStyle}>{text}</th>");
                }
                sb.AppendLine("    </tr>");
                sb.AppendLine("  </thead>");
                startDataRow = 2;
            }

            sb.AppendLine("  <tbody>");
            for (int r = startDataRow; r <= rows; r++)
            {
                string rowBg = (options.IncludeHtmlStyles && r % 2 == 0) ? " style=\"background-color: #f8fafc;\"" : "";
                sb.AppendLine($"    <tr{rowBg}>");

                for (int c = 1; c <= cols; c++)
                {
                    object? cellVal = values[r, c];
                    string text = FormatHtmlCell(cellVal, options);
                    bool isNum = options.AlignNumbersRight && double.TryParse(cellVal?.ToString(), out _);
                    string styleToUse = isNum ? tdNumStyle : tdStyle;

                    sb.AppendLine($"      <td {styleToUse}>{text}</td>");
                }
                sb.AppendLine("    </tr>");
            }
            sb.AppendLine("  </tbody>");
            sb.AppendLine("</table>");

            return sb.ToString();
        }

        private static string FormatHtmlCell(object? val, TableExportOptions options)
        {
            if (val == null) return string.Empty;
            string text = val.ToString() ?? string.Empty;

            text = text.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;");
            if (options.ConvertLineBreaksToBr)
            {
                text = text.Replace("\r\n", "<br>").Replace("\n", "<br>").Replace("\r", "<br>");
            }
            return text;
        }

        private static object[,] GetRangeValues2D(Range rng)
        {
            try
            {
                object raw = rng.Value2;
                if (raw is object[,] arr)
                {
                    return arr;
                }
                else if (raw != null)
                {
                    object[,] single = new object[2, 2];
                    single[1, 1] = raw;
                    return single;
                }
            }
            catch { }

            return new object[1, 1];
        }
    }
}
