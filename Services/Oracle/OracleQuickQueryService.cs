using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
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
        public static Task<System.Data.DataTable> ExecuteQueryAsync(OracleConnectionConfig config, string sql, int maxRows = 0)
            => OracleSchemaInspector.ExecuteQueryAsync(config, sql, maxRows);

        public static string ExtractTableName(string sql)
            => OracleSchemaInspector.ExtractTableName(sql);

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

        public static Task<OracleTableStructureResult> GetTableStructureAsync(OracleConnectionConfig config, string fullTableName)
            => OracleSchemaInspector.GetTableStructureAsync(config, fullTableName);

        public static (int RowsInserted, int ColsInserted) ExportTableStructureToWorksheet(
            ExcelWorksheet ws,
            OracleTableStructureResult structure)
        {
            if (structure == null) return (0, 0);

            Color titleColor = Color.FromArgb(37, 99, 235);         // Blue (#2563EB)
            Color colHeaderBgColor = Color.FromArgb(204, 255, 255); // Pastel Cyan (#CCFFFF)
            Color idxHeaderBgColor = Color.FromArgb(191, 219, 254); // Light Blue (#BFDBFE)
            Color borderColor = Color.FromArgb(156, 163, 175);       // Gray 400

            int startRow = 2;
            int startCol = 2;
            int curRow = startRow;

            // 1. Tiêu đề Bảng
            ExcelRange titleRange = ws.Cells[curRow, startCol];
            titleRange.Value2 = $"CẤU TRÚC BẢNG ORACLE: {structure.FullTableName}";
            titleRange.Font.Bold = true;
            titleRange.Font.Size = 13;
            titleRange.Font.Color = ColorTranslator.ToOle(titleColor);
            curRow++;

            ExcelRange timeRange = ws.Cells[curRow, startCol];
            timeRange.Value2 = $"Thời gian xuất: {DateTime.Now:yyyy-MM-dd HH:mm:ss} | Tổng số cột: {structure.Columns.Count} | Tổng số index: {structure.Indexes.Count}";
            timeRange.Font.Italic = true;
            timeRange.Font.Size = 9.5;
            timeRange.Font.Color = ColorTranslator.ToOle(Color.FromArgb(100, 116, 139));
            curRow += 2;

            // 2. PHẦN 1: BẢNG CỘT (COLUMNS)
            ExcelRange sec1Title = ws.Cells[curRow, startCol];
            sec1Title.Value2 = $"📋 1. DANH SÁCH CỘT ({structure.Columns.Count} CỘT)";
            sec1Title.Font.Bold = true;
            sec1Title.Font.Size = 11;
            sec1Title.Font.Color = ColorTranslator.ToOle(Color.FromArgb(30, 41, 59));
            curRow++;

            int colTableStartRow = curRow;
            string[] colHeaders = { "STT", "Khóa Chính (PK)", "Tên Cột", "Kiểu Dữ Liệu", "Nullable", "Giá Trị Mặc Định", "Chú Thích (Comments)" };
            int numColAttrs = colHeaders.Length;

            object[,] colHeaderArr = new object[1, numColAttrs];
            for (int i = 0; i < numColAttrs; i++) colHeaderArr[0, i] = colHeaders[i];

            ExcelRange colHeaderRange = ws.Range[ws.Cells[curRow, startCol], ws.Cells[curRow, startCol + numColAttrs - 1]];
            colHeaderRange.Value2 = colHeaderArr;
            colHeaderRange.Font.Bold = true;
            colHeaderRange.Font.Size = 10;
            colHeaderRange.Interior.Color = ColorTranslator.ToOle(colHeaderBgColor);
            colHeaderRange.HorizontalAlignment = -4108; // xlCenter
            curRow++;

            int numCols = structure.Columns.Count;
            if (numCols > 0)
            {
                object[,] colDataArr = new object[numCols, numColAttrs];
                for (int r = 0; r < numCols; r++)
                {
                    var col = structure.Columns[r];
                    colDataArr[r, 0] = col.ColumnId;
                    colDataArr[r, 1] = col.IsPrimaryKey ? $"⭐ PK ({col.PrimaryKeyPosition})" : "";
                    colDataArr[r, 2] = col.ColumnName;
                    colDataArr[r, 3] = col.DisplayDataType;
                    colDataArr[r, 4] = col.NullableText;
                    colDataArr[r, 5] = col.DataDefault ?? "";
                    colDataArr[r, 6] = col.Comments ?? "";
                }

                ExcelRange colDataRange = ws.Range[ws.Cells[curRow, startCol], ws.Cells[curRow + numCols - 1, startCol + numColAttrs - 1]];
                colDataRange.Value2 = colDataArr;
                colDataRange.Font.Size = 9.5;

                ExcelRange fullColTable = ws.Range[ws.Cells[colTableStartRow, startCol], ws.Cells[curRow + numCols - 1, startCol + numColAttrs - 1]];
                fullColTable.Borders.LineStyle = 1;
                fullColTable.Borders.Color = ColorTranslator.ToOle(borderColor);

                curRow += numCols;
            }

            curRow += 2; // khoảng cách

            // 3. PHẦN 2: BẢNG INDEX (INDEXES)
            ExcelRange sec2Title = ws.Cells[curRow, startCol];
            sec2Title.Value2 = $"⚡ 2. DANH SÁCH INDEX ({structure.Indexes.Count} INDEX)";
            sec2Title.Font.Bold = true;
            sec2Title.Font.Size = 11;
            sec2Title.Font.Color = ColorTranslator.ToOle(Color.FromArgb(30, 41, 59));
            curRow++;

            int idxTableStartRow = curRow;
            string[] idxHeaders = { "STT", "Tên Index", "Phân Loại", "Duy Nhất (Unique)", "Cột Tham Gia", "Kiểu Index", "Trạng Thái" };
            int numIdxAttrs = idxHeaders.Length;

            object[,] idxHeaderArr = new object[1, numIdxAttrs];
            for (int i = 0; i < numIdxAttrs; i++) idxHeaderArr[0, i] = idxHeaders[i];

            ExcelRange idxHeaderRange = ws.Range[ws.Cells[curRow, startCol], ws.Cells[curRow, startCol + numIdxAttrs - 1]];
            idxHeaderRange.Value2 = idxHeaderArr;
            idxHeaderRange.Font.Bold = true;
            idxHeaderRange.Font.Size = 10;
            idxHeaderRange.Interior.Color = ColorTranslator.ToOle(idxHeaderBgColor);
            idxHeaderRange.HorizontalAlignment = -4108; // xlCenter
            curRow++;

            int numIdx = structure.Indexes.Count;
            if (numIdx > 0)
            {
                object[,] idxDataArr = new object[numIdx, numIdxAttrs];
                for (int r = 0; r < numIdx; r++)
                {
                    var idx = structure.Indexes[r];
                    idxDataArr[r, 0] = r + 1;
                    idxDataArr[r, 1] = idx.IndexName;
                    idxDataArr[r, 2] = idx.IndexCategory;
                    idxDataArr[r, 3] = idx.IsUnique ? "YES" : "NO";
                    idxDataArr[r, 4] = idx.ColumnsDisplay;
                    idxDataArr[r, 5] = idx.IndexType;
                    idxDataArr[r, 6] = idx.Status;
                }

                ExcelRange idxDataRange = ws.Range[ws.Cells[curRow, startCol], ws.Cells[curRow + numIdx - 1, startCol + numIdxAttrs - 1]];
                idxDataRange.Value2 = idxDataArr;
                idxDataRange.Font.Size = 9.5;

                ExcelRange fullIdxTable = ws.Range[ws.Cells[idxTableStartRow, startCol], ws.Cells[curRow + numIdx - 1, startCol + numIdxAttrs - 1]];
                fullIdxTable.Borders.LineStyle = 1;
                fullIdxTable.Borders.Color = ColorTranslator.ToOle(borderColor);

                curRow += numIdx;
            }

            // AutoFit các cột
            int maxCols = Math.Max(numColAttrs, numIdxAttrs);
            ExcelRange allColsRange = ws.Range[ws.Cells[startRow, startCol], ws.Cells[curRow, startCol + maxCols - 1]];
            allColsRange.Columns.AutoFit();

            return (curRow - startRow, maxCols);
        }

        #region AI & Rule-Based SQL Query Generation from Excel Selection

        /// <summary>
        /// Đọc nội dung văn bản từ Range hoặc Selection hiện tại của Excel
        /// </summary>
        public static string GetExcelSelectionText(ExcelApp? app)
        {
            if (app == null) return string.Empty;

            try
            {
                var sel = app.Selection as ExcelRange;
                if (sel == null) return string.Empty;

                object[,] arr = SqlScriptGeneratorService.GetRangeValues2D(sel);
                int rows = arr.GetLength(0);
                int cols = arr.GetLength(1);

                if (rows == 0 || cols == 0) return string.Empty;

                var sb = new StringBuilder();
                for (int r = 1; r <= rows; r++)
                {
                    var rowTokens = new List<string>();
                    for (int c = 1; c <= cols; c++)
                    {
                        string val = arr[r, c]?.ToString()?.Trim() ?? string.Empty;
                        if (!string.IsNullOrEmpty(val))
                        {
                            rowTokens.Add(val);
                        }
                    }
                    if (rowTokens.Count > 0)
                    {
                        sb.AppendLine(string.Join("\t", rowTokens));
                    }
                }

                return sb.ToString().Trim();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[GetExcelSelectionText] error: {ex.Message}");
                return string.Empty;
            }
        }

        /// <summary>
        /// Phân tích văn bản/bảng điều kiện từ vùng chọn Excel và sinh câu lệnh SQL SELECT chuẩn Oracle
        /// Kết hợp giữa AI và Fallback Parser cục bộ
        /// </summary>
        public static async Task<string> GenerateSqlFromSelectionAsync(string selectionText, AiConfig? aiConfig = null)
        {
            if (string.IsNullOrWhiteSpace(selectionText)) return string.Empty;

            aiConfig ??= AiConfigManager.Current;

            // 1. Thử gọi AI nếu đã cấu hình API Key
            if (aiConfig != null && !string.IsNullOrWhiteSpace(aiConfig.ApiKey))
            {
                try
                {
                    string systemPrompt =
                        "Bạn là chuyên gia cơ sở dữ liệu Oracle và chuyên viên BSE/Developer phân tích tài liệu thiết kế phần mềm Nhật Bản (TKCT/TKCB).\n" +
                        "Nhiệm vụ: Phân tích bảng/danh sách điều kiện trích xuất dữ liệu từ các ô Excel và sinh ra câu lệnh SQL SELECT chuẩn xác cho Oracle Database.\n\n" +
                        "QUY TẮC BẮT BUỘC:\n" +
                        "1. Xác định đúng tên Bảng từ tiêu đề (ví dụ: 'table 受注見出トラン[JDNTHA]' -> table là JDNTHA; hoặc 'Bảng ADDMTA' -> ADDMTA).\n" +
                        "2. Cấu trúc câu lệnh:\n" +
                        "   SELECT *\n" +
                        "   FROM <TÊN_BẢNG>\n" +
                        "   WHERE <ĐIỀU_KIỆN_1>\n" +
                        "     AND <ĐIỀU_KIỆN_2>\n" +
                        "     ...\n" +
                        "   ORDER BY <CÁC_CỘT>;\n" +
                        "3. Chuyển đổi giá trị:\n" +
                        "   - Giá trị hằng số trong nháy kép \"...\" chuyển thành nháy đơn '...'. (Ví dụ: [DATKB] = \"1\" -> DATKB = '1').\n" +
                        "   - Giá trị tham chiếu biến màn hình như '画面.処理区分[HD_HANKB]' hoặc '[HD_KTNCD]' -> chuyển thành Oracle bind variable :HD_HANKB, :HD_KTNCD.\n" +
                        "   - Toán tử so sánh: =, >=, <=, <>, !=, LIKE, IN, IS NULL, IS NOT NULL.\n" +
                        "   - Nếu có dòng 'ORDER BY ...' thì đưa vào mệnh đề ORDER BY ở cuối.\n" +
                        "   - Các dòng ghi chú mô tả bằng chữ không phải điều kiện SQL -> chuyển thành comment SQL: -- Ghi chú: ...\n" +
                        "4. Chỉ trả về mã SQL bên trong khối ```sql\n...\n```, không kèm lời chào hay giải thích ngoài khối code.";

                    string userPrompt = $"Dưới đây là nội dung các ô Excel chứa điều kiện trích xuất dữ liệu:\n\n{selectionText}\n\nHãy sinh câu lệnh SQL SELECT Oracle tương ứng:";
                    string aiReply = await OpenAiClientService.SendChatAsync(aiConfig, userPrompt, systemPrompt);

                    string extractedSql = ExtractSqlBlock(aiReply);
                    if (!string.IsNullOrWhiteSpace(extractedSql))
                    {
                        return extractedSql;
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[GenerateSqlFromSelectionAsync] AI error: {ex.Message}. Falling back to local parser.");
                }
            }

            // 2. Fallback Parser cục bộ bằng thuật toán Rule-based & Regex
            return ParseSpecConditionsToSql(selectionText);
        }

        /// <summary>
        /// Trích xuất khối mã SQL từ phản hồi văn bản của AI
        /// </summary>
        public static string ExtractSqlBlock(string response)
        {
            if (string.IsNullOrWhiteSpace(response)) return string.Empty;

            var match = Regex.Match(response, @"```(?:sql)?\s*([\s\S]*?)\s*```", RegexOptions.IgnoreCase);
            if (match.Success)
            {
                return match.Groups[1].Value.Trim();
            }

            // Nếu trả về trực tiếp câu SELECT ... FROM
            if (response.IndexOf("SELECT", StringComparison.OrdinalIgnoreCase) >= 0 &&
                response.IndexOf("FROM", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return response.Trim();
            }

            return string.Empty;
        }

        /// <summary>
        /// Bộ phân tích cục bộ (Rule-based Spec Parser) phân tích đặc tả điều kiện trích xuất sang câu lệnh SQL Oracle
        /// </summary>
        public static string ParseSpecConditionsToSql(string rawText)
        {
            if (string.IsNullOrWhiteSpace(rawText)) return string.Empty;

            var lines = rawText.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries)
                               .Select(l => l.Trim())
                               .Where(l => !string.IsNullOrWhiteSpace(l))
                               .ToList();

            if (lines.Count == 0) return string.Empty;

            string tableName = "MY_TABLE";
            var conditions = new List<string>();
            var notes = new List<string>();
            string orderByClause = string.Empty;

            // 1. Nhận diện tên Bảng
            bool foundTable = false;
            foreach (var line in lines)
            {
                // Mẫu 1: table 受注見出トラン[JDNTHA] hoặc table [JDNTHA] hoặc table JDNTHA
                var tblMatch = Regex.Match(line, @"(?:table|bảng|data\s+table)\s*[^\[\r\n]*\[(?<table>[A-Za-z0-9_]+)\]", RegexOptions.IgnoreCase);
                if (tblMatch.Success)
                {
                    tableName = tblMatch.Groups["table"].Value.Trim().ToUpperInvariant();
                    foundTable = true;
                    break;
                }

                // Mẫu 2: FROM [TABLE] hoặc FROM TABLE
                var fromMatch = Regex.Match(line, @"\bFROM\s+\[?(?<table>[A-Za-z0-9_]+)\]?", RegexOptions.IgnoreCase);
                if (fromMatch.Success)
                {
                    tableName = fromMatch.Groups["table"].Value.Trim().ToUpperInvariant();
                    foundTable = true;
                    break;
                }
            }

            // Nếu chưa tìm thấy bảng, tìm tên bảng trong ngoặc vuông đầu tiên ở 3 dòng đầu (không chứa toán tử)
            if (!foundTable)
            {
                for (int i = 0; i < Math.Min(3, lines.Count); i++)
                {
                    string l = lines[i];
                    if (!l.Contains("=") && !l.Contains(">") && !l.Contains("<"))
                    {
                        var bracketMatch = Regex.Match(l, @"\[(?<table>[A-Za-z0-9_]{3,30})\]");
                        if (bracketMatch.Success)
                        {
                            tableName = bracketMatch.Groups["table"].Value.Trim().ToUpperInvariant();
                            break;
                        }
                    }
                }
            }

            // 2. Nhận diện các điều kiện cột và Order By
            foreach (var line in lines)
            {
                // Kiểm tra ORDER BY
                var orderMatch = Regex.Match(line, @"^\s*ORDER\s+BY\s+(?<cols>.+)$", RegexOptions.IgnoreCase);
                if (orderMatch.Success)
                {
                    orderByClause = $"ORDER BY {orderMatch.Groups["cols"].Value.Trim()}";
                    continue;
                }

                // Kiểm tra dòng điều kiện cột có ngoặc vuông: [DATKB] = "1" hoặc 伝票削除区分 [DATKB] = "1"
                var colMatch = Regex.Match(line, @"\[(?<col>[A-Za-z0-9_]+)\]\s*(?<op>>=|<=|<>|!=|=|>|<|LIKE|IN)\s*(?<val>.+)", RegexOptions.IgnoreCase);
                if (!colMatch.Success)
                {
                    // Fallback: không có ngoặc vuông ở cột: DATKB = "1"
                    colMatch = Regex.Match(line, @"\b(?<col>[A-Za-z0-9_]{2,30})\s*(?<op>>=|<=|<>|!=|=|>|<|LIKE|IN)\s*(?<val>.+)", RegexOptions.IgnoreCase);
                }

                if (colMatch.Success)
                {
                    string col = colMatch.Groups["col"].Value.Trim().ToUpperInvariant();
                    string op = colMatch.Groups["op"].Value.Trim().ToUpperInvariant();
                    string rawVal = colMatch.Groups["val"].Value.Trim();

                    string formattedVal = FormatConditionValue(rawVal, op);
                    if (formattedVal == "IS NULL" || formattedVal == "IS NOT NULL")
                    {
                        conditions.Add($"{col} {formattedVal}");
                    }
                    else
                    {
                        conditions.Add($"{col} {op} {formattedVal}");
                    }
                    continue;
                }

                // Nếu là dòng tiêu đề chứa tên bảng thì bỏ qua
                if (line.IndexOf(tableName, StringComparison.OrdinalIgnoreCase) >= 0 &&
                    (line.IndexOf("table", StringComparison.OrdinalIgnoreCase) >= 0 || line.IndexOf("bảng", StringComparison.OrdinalIgnoreCase) >= 0 || line.IndexOf("trích xuất", StringComparison.OrdinalIgnoreCase) >= 0))
                {
                    continue;
                }

                // Dòng ghi chú mô tả bằng chữ (ví dụ: '部門制御 thì giống nguồn tham chiếu')
                if (!line.StartsWith("▲") && !line.StartsWith("/*") && line.Length > 2)
                {
                    notes.Add($"-- Ghi chú: {line.Trim()}");
                }
            }

            // 3. Xây dựng câu SQL hoàn chỉnh
            var sb = new StringBuilder();
            sb.AppendLine("SELECT *");
            sb.AppendLine($"FROM {tableName}");

            if (conditions.Count > 0)
            {
                sb.Append("WHERE ");
                for (int i = 0; i < conditions.Count; i++)
                {
                    if (i > 0)
                    {
                        sb.Append("  AND ");
                    }
                    sb.AppendLine(conditions[i]);
                }
            }

            if (notes.Count > 0)
            {
                foreach (var note in notes)
                {
                    sb.AppendLine($"  {note}");
                }
            }

            if (!string.IsNullOrWhiteSpace(orderByClause))
            {
                sb.AppendLine(orderByClause);
            }

            return sb.ToString().Trim();
        }

        private static string FormatConditionValue(string rawVal, string op)
        {
            if (string.IsNullOrWhiteSpace(rawVal)) return "''";

            string s = rawVal.Trim();

            // Xử lý NULL
            if (s.Equals("NULL", StringComparison.OrdinalIgnoreCase))
            {
                return (op == "!=" || op == "<>") ? "IS NOT NULL" : "IS NULL";
            }

            // Xử lý tham chiếu màn hình: 画面.処理区分[HD_HANKB] hoặc [HD_KTNCD]
            var screenVarMatch = Regex.Match(s, @"(?:画面[^\r\n\[]*\[|\[)(?<var>[A-Za-z0-9_]+)\]");
            if (screenVarMatch.Success)
            {
                string varName = screenVarMatch.Groups["var"].Value.Trim();
                return $":{varName}";
            }

            // Xử lý giá trị trong nháy kép "1" -> '1'
            if (s.StartsWith("\"") && s.EndsWith("\"") && s.Length >= 2)
            {
                string inner = s.Substring(1, s.Length - 2);
                return $"'{inner}'";
            }

            // Xử lý giá trị trong nháy đơn '1' -> giữ nguyên
            if (s.StartsWith("'") && s.EndsWith("'") && s.Length >= 2)
            {
                return s;
            }

            // Xử lý danh sách IN: ("1", "2") hoặc ('1', '2')
            if (s.StartsWith("(") && s.EndsWith(")"))
            {
                return s.Replace("\"", "'");
            }

            // Số nguyên hoặc số thập phân
            if (long.TryParse(s, out _) || double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out _))
            {
                return s;
            }

            // Chuỗi không có nháy -> bao quanh bằng nháy đơn
            return $"'{s}'";
        }

        #endregion
    }
}
