using System;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using ExcelSupport.Models;
using Microsoft.Office.Interop.Excel;
using Oracle.ManagedDataAccess.Client;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using ExcelRange = Microsoft.Office.Interop.Excel.Range;
using ExcelWorksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace ExcelSupport.Services
{
    public static class OracleQuickQueryService
    {
        public static async Task<System.Data.DataTable> ExecuteQueryAsync(OracleConnectionConfig config, string sql, int maxRows = 0)
        {
            string connStr = config.BuildConnectionString();
            return await Task.Run(() =>
            {
                var dt = new System.Data.DataTable();
                using (var conn = new OracleConnection(connStr))
                {
                    conn.Open();

                    string execSql = sql.Trim().TrimEnd(';');
                    if (maxRows > 0 && !execSql.ToUpperInvariant().Contains("ROWNUM"))
                    {
                        execSql = $"SELECT * FROM ({execSql}) WHERE ROWNUM <= {maxRows}";
                    }

                    using (var cmd = new OracleCommand(execSql, conn))
                    {
                        cmd.CommandTimeout = 120;
                        using (var adapter = new OracleDataAdapter(cmd))
                        {
                            adapter.Fill(dt);
                        }
                    }
                }
                return dt;
            });
        }

        public static string ExtractTableName(string sql)
        {
            if (string.IsNullOrWhiteSpace(sql)) return "QUERY_RESULT";

            try
            {
                // Match "FROM [schema.]tableName"
                var match = Regex.Match(sql, @"\bFROM\s+([""']?(?<schema>[a-zA-Z0-9_]+)[""']?\.)?([""']?(?<table>[a-zA-Z0-9_]+)[""']?)", RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    string table = match.Groups["table"].Value;
                    if (!string.IsNullOrWhiteSpace(table))
                    {
                        return table.ToUpperInvariant();
                    }
                }
            }
            catch { }

            return "QUERY_RESULT";
        }

        public static (int RowsInserted, int ColsInserted) InsertDataToWorksheet(
            ExcelWorksheet ws,
            int startRow,
            int startCol,
            System.Data.DataTable dt,
            string tableName,
            OracleQuickQueryOptions options)
        {
            if (dt == null) return (0, 0);

            int numCols = dt.Columns.Count;
            int numRows = dt.Rows.Count;
            if (numCols == 0) return (0, 0);

            Color titleColor = Color.FromArgb(37, 99, 235);       // Blue (#2563EB) default
            if (!string.IsNullOrWhiteSpace(options.TitleColorHex))
            {
                try
                {
                    titleColor = ColorTranslator.FromHtml(options.TitleColorHex);
                }
                catch { }
            }

            Color headerBgColor = Color.FromArgb(204, 255, 255);   // Pastel Cyan (#CCFFFF) default
            if (!string.IsNullOrWhiteSpace(options.HeaderBgColorHex))
            {
                try
                {
                    headerBgColor = ColorTranslator.FromHtml(options.HeaderBgColorHex);
                }
                catch { }
            }

            // Tính toán màu chữ Header tương phản tự động
            double luminance = (0.299 * headerBgColor.R + 0.587 * headerBgColor.G + 0.114 * headerBgColor.B) / 255.0;
            Color headerTextColor = luminance < 0.5 ? Color.White : Color.FromArgb(15, 23, 42); // Dark Slate for light bg, White for dark bg
            Color borderColor = Color.FromArgb(156, 163, 175);     // Gray 400

            int curRow = startRow;

            // 1. Tiêu đề bảng (Chỉ hiển thị tên Table theo yêu cầu người dùng)
            if (options.IncludeTitle && !string.IsNullOrWhiteSpace(tableName))
            {
                ExcelRange titleRange = ws.Cells[curRow, startCol];
                titleRange.Value2 = tableName;
                titleRange.Font.Bold = true;
                titleRange.Font.Size = 11;
                titleRange.Font.Color = ColorTranslator.ToOle(titleColor);
                curRow++;
            }

            // 2. Dòng Header
            int headerRow = curRow;
            if (options.IncludeHeaders)
            {
                object[,] headerArray = new object[1, numCols];
                for (int c = 0; c < numCols; c++)
                {
                    headerArray[0, c] = dt.Columns[c].ColumnName;
                }

                ExcelRange headerRange = ws.Range[ws.Cells[headerRow, startCol], ws.Cells[headerRow, startCol + numCols - 1]];
                headerRange.Value2 = headerArray;
                headerRange.Font.Bold = true;
                headerRange.Font.Size = 10;
                headerRange.Font.Color = ColorTranslator.ToOle(headerTextColor);
                headerRange.Interior.Color = ColorTranslator.ToOle(headerBgColor);
                headerRange.HorizontalAlignment = -4108; // xlCenter
                curRow++;
            }

            // 3. Dữ liệu các dòng
            int dataStartRow = curRow;
            if (numRows > 0)
            {
                object[,] dataArray = new object[numRows, numCols];
                for (int r = 0; r < numRows; r++)
                {
                    DataRow row = dt.Rows[r];
                    for (int c = 0; c < numCols; c++)
                    {
                        object val = row[c];
                        if (val == null || val is DBNull)
                        {
                            dataArray[r, c] = "";
                        }
                        else if (val is DateTime dtVal)
                        {
                            dataArray[r, c] = dtVal.ToString("yyyy-MM-dd HH:mm:ss");
                        }
                        else
                        {
                            dataArray[r, c] = val.ToString() ?? "";
                        }
                    }
                }

                ExcelRange dataRange = ws.Range[ws.Cells[dataStartRow, startCol], ws.Cells[dataStartRow + numRows - 1, startCol + numCols - 1]];
                dataRange.Value2 = dataArray;

                // Kẻ viền toàn bộ khối bảng (Header + Data)
                int startBorderRow = options.IncludeHeaders ? headerRow : dataStartRow;
                ExcelRange fullTable = ws.Range[ws.Cells[startBorderRow, startCol], ws.Cells[dataStartRow + numRows - 1, startCol + numCols - 1]];
                fullTable.Borders.LineStyle = 1; // xlContinuous
                fullTable.Borders.Color = ColorTranslator.ToOle(borderColor);

                curRow = dataStartRow + numRows;
            }

            // 4. AutoFit cột
            if (options.AutoFitColumns)
            {
                ExcelRange allColsRange = ws.Range[ws.Cells[startRow, startCol], ws.Cells[curRow, startCol + numCols - 1]];
                allColsRange.Columns.AutoFit();
            }

            return (numRows, numCols);
        }
    }
}
