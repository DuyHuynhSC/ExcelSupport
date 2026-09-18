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

        public static async Task<OracleTableStructureResult> GetTableStructureAsync(OracleConnectionConfig config, string fullTableName)
        {
            if (string.IsNullOrWhiteSpace(fullTableName))
                throw new ArgumentException("Tên bảng không được để trống.", nameof(fullTableName));

            string connStr = config.BuildConnectionString();
            return await Task.Run(() =>
            {
                var result = new OracleTableStructureResult();
                string cleaned = fullTableName.Trim().Replace("\"", "").Replace("'", "");
                string owner = "";
                string tableName = cleaned;

                if (cleaned.Contains("."))
                {
                    var parts = cleaned.Split('.');
                    owner = parts[0].Trim().ToUpperInvariant();
                    tableName = parts[1].Trim().ToUpperInvariant();
                }
                else
                {
                    tableName = cleaned.Trim().ToUpperInvariant();
                }

                using (var conn = new OracleConnection(connStr))
                {
                    conn.Open();

                    // Nếu chưa có owner, tự động tìm kiếm owner trong ALL_TABLES / ALL_VIEWS
                    if (string.IsNullOrEmpty(owner))
                    {
                        using (var cmdOwner = new OracleCommand(
                            @"SELECT OWNER FROM ALL_TABLES 
                              WHERE TABLE_NAME = :tbl 
                              ORDER BY CASE 
                                WHEN OWNER = SYS_CONTEXT('USERENV', 'CURRENT_SCHEMA') THEN 0 
                                WHEN OWNER = :usr THEN 1 
                                ELSE 2 END", conn))
                        {
                            cmdOwner.Parameters.Add(new OracleParameter("tbl", tableName));
                            cmdOwner.Parameters.Add(new OracleParameter("usr", (config.Username ?? "").ToUpperInvariant()));
                            using (var rdr = cmdOwner.ExecuteReader())
                            {
                                if (rdr.Read())
                                {
                                    owner = rdr.GetString(0);
                                }
                            }
                        }

                        if (string.IsNullOrEmpty(owner))
                        {
                            using (var cmdViewOwner = new OracleCommand(
                                @"SELECT OWNER FROM ALL_VIEWS 
                                  WHERE VIEW_NAME = :tbl 
                                  ORDER BY CASE 
                                    WHEN OWNER = SYS_CONTEXT('USERENV', 'CURRENT_SCHEMA') THEN 0 
                                    WHEN OWNER = :usr THEN 1 
                                    ELSE 2 END", conn))
                            {
                                cmdViewOwner.Parameters.Add(new OracleParameter("tbl", tableName));
                                cmdViewOwner.Parameters.Add(new OracleParameter("usr", (config.Username ?? "").ToUpperInvariant()));
                                using (var rdr = cmdViewOwner.ExecuteReader())
                                {
                                    if (rdr.Read())
                                    {
                                        owner = rdr.GetString(0);
                                    }
                                }
                            }
                        }

                        if (string.IsNullOrEmpty(owner))
                        {
                            owner = (config.Username ?? "").ToUpperInvariant();
                        }
                    }

                    result.Owner = owner;
                    result.TableName = tableName;

                    // 1. Query Primary Key columns
                    var pkDict = new System.Collections.Generic.Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                    using (var cmdPk = new OracleCommand(
                        @"SELECT cc.COLUMN_NAME, cc.POSITION
                          FROM ALL_CONSTRAINTS con
                          JOIN ALL_CONS_COLUMNS cc 
                            ON con.OWNER = cc.OWNER AND con.CONSTRAINT_NAME = cc.CONSTRAINT_NAME
                          WHERE con.CONSTRAINT_TYPE = 'P'
                            AND con.OWNER = :owner
                            AND con.TABLE_NAME = :tableName
                          ORDER BY cc.POSITION", conn))
                    {
                        cmdPk.Parameters.Add(new OracleParameter("owner", owner));
                        cmdPk.Parameters.Add(new OracleParameter("tableName", tableName));
                        using (var rdr = cmdPk.ExecuteReader())
                        {
                            while (rdr.Read())
                            {
                                string colName = rdr.GetString(0);
                                int pos = rdr.IsDBNull(1) ? 1 : Convert.ToInt32(rdr.GetValue(1));
                                pkDict[colName] = pos;
                            }
                        }
                    }

                    // 2. Query Columns & Comments
                    using (var cmdCols = new OracleCommand(
                        @"SELECT 
                            c.COLUMN_ID,
                            c.COLUMN_NAME,
                            c.DATA_TYPE,
                            c.DATA_LENGTH,
                            c.DATA_PRECISION,
                            c.DATA_SCALE,
                            c.NULLABLE,
                            c.DATA_DEFAULT,
                            cm.COMMENTS
                          FROM ALL_TAB_COLUMNS c
                          LEFT JOIN ALL_COL_COMMENTS cm 
                            ON c.OWNER = cm.OWNER AND c.TABLE_NAME = cm.TABLE_NAME AND c.COLUMN_NAME = cm.COLUMN_NAME
                          WHERE c.OWNER = :owner AND c.TABLE_NAME = :tableName
                          ORDER BY c.COLUMN_ID", conn))
                    {
                        cmdCols.Parameters.Add(new OracleParameter("owner", owner));
                        cmdCols.Parameters.Add(new OracleParameter("tableName", tableName));
                        using (var rdr = cmdCols.ExecuteReader())
                        {
                            while (rdr.Read())
                            {
                                var col = new OracleTableColumnMetadata
                                {
                                    ColumnId = rdr.IsDBNull(0) ? 0 : Convert.ToInt32(rdr.GetValue(0)),
                                    ColumnName = rdr.IsDBNull(1) ? "" : rdr.GetString(1),
                                    DataType = rdr.IsDBNull(2) ? "" : rdr.GetString(2),
                                    DataLength = rdr.IsDBNull(3) ? (int?)null : Convert.ToInt32(rdr.GetValue(3)),
                                    DataPrecision = rdr.IsDBNull(4) ? (int?)null : Convert.ToInt32(rdr.GetValue(4)),
                                    DataScale = rdr.IsDBNull(5) ? (int?)null : Convert.ToInt32(rdr.GetValue(5)),
                                    IsNullable = !rdr.IsDBNull(6) && rdr.GetString(6).Equals("Y", StringComparison.OrdinalIgnoreCase),
                                    DataDefault = rdr.IsDBNull(7) ? null : rdr.GetString(7)?.Trim(),
                                    Comments = rdr.IsDBNull(8) ? null : rdr.GetString(8)?.Trim()
                                };

                                if (pkDict.TryGetValue(col.ColumnName, out int pkPos))
                                {
                                    col.IsPrimaryKey = true;
                                    col.PrimaryKeyPosition = pkPos;
                                }

                                result.Columns.Add(col);
                            }
                        }
                    }

                    // 3. Query Indexes
                    var indexMap = new System.Collections.Generic.Dictionary<string, OracleTableIndexMetadata>(StringComparer.OrdinalIgnoreCase);
                    using (var cmdIdx = new OracleCommand(
                        @"SELECT 
                            i.INDEX_NAME,
                            i.UNIQUENESS,
                            i.STATUS,
                            i.INDEX_TYPE,
                            (SELECT COUNT(*) 
                             FROM ALL_CONSTRAINTS con 
                             WHERE con.OWNER = i.OWNER 
                               AND con.CONSTRAINT_TYPE = 'P' 
                               AND con.CONSTRAINT_NAME = i.INDEX_NAME) AS IS_PK_CONSTRAINT
                          FROM ALL_INDEXES i
                          WHERE i.TABLE_OWNER = :owner AND i.TABLE_NAME = :tableName
                          ORDER BY IS_PK_CONSTRAINT DESC, i.INDEX_NAME", conn))
                    {
                        cmdIdx.Parameters.Add(new OracleParameter("owner", owner));
                        cmdIdx.Parameters.Add(new OracleParameter("tableName", tableName));
                        using (var rdr = cmdIdx.ExecuteReader())
                        {
                            while (rdr.Read())
                            {
                                string idxName = rdr.GetString(0);
                                string uniqueness = rdr.IsDBNull(1) ? "NONUNIQUE" : rdr.GetString(1);
                                string status = rdr.IsDBNull(2) ? "VALID" : rdr.GetString(2);
                                string idxType = rdr.IsDBNull(3) ? "NORMAL" : rdr.GetString(3);
                                int isPkCon = rdr.IsDBNull(4) ? 0 : Convert.ToInt32(rdr.GetValue(4));

                                var idxMeta = new OracleTableIndexMetadata
                                {
                                    IndexName = idxName,
                                    TableName = tableName,
                                    TableOwner = owner,
                                    IsPrimaryKey = isPkCon > 0,
                                    IsUnique = uniqueness.Equals("UNIQUE", StringComparison.OrdinalIgnoreCase),
                                    Status = status,
                                    IndexType = idxType
                                };
                                indexMap[idxName] = idxMeta;
                                result.Indexes.Add(idxMeta);
                            }
                        }
                    }

                    // 4. Query Index Columns
                    if (indexMap.Count > 0)
                    {
                        using (var cmdIndCols = new OracleCommand(
                            @"SELECT 
                                ic.INDEX_NAME,
                                ic.COLUMN_NAME,
                                ic.COLUMN_POSITION,
                                ic.DESCEND
                              FROM ALL_IND_COLUMNS ic
                              WHERE ic.TABLE_OWNER = :owner AND ic.TABLE_NAME = :tableName
                              ORDER BY ic.INDEX_NAME, ic.COLUMN_POSITION", conn))
                        {
                            cmdIndCols.Parameters.Add(new OracleParameter("owner", owner));
                            cmdIndCols.Parameters.Add(new OracleParameter("tableName", tableName));
                            using (var rdr = cmdIndCols.ExecuteReader())
                            {
                                var groupedCols = new System.Collections.Generic.Dictionary<string, System.Collections.Generic.List<string>>(StringComparer.OrdinalIgnoreCase);
                                while (rdr.Read())
                                {
                                    string idxName = rdr.GetString(0);
                                    string colName = rdr.GetString(1);
                                    string descend = rdr.IsDBNull(3) ? "ASC" : rdr.GetString(3);
                                    string colDesc = $"{colName} {descend}".Trim();

                                    if (!groupedCols.TryGetValue(idxName, out var list))
                                    {
                                        list = new System.Collections.Generic.List<string>();
                                        groupedCols[idxName] = list;
                                    }
                                    list.Add(colDesc);

                                    if (indexMap.TryGetValue(idxName, out var idxMeta))
                                    {
                                        idxMeta.ColumnNames.Add(colName);
                                    }
                                }

                                foreach (var kvp in groupedCols)
                                {
                                    if (indexMap.TryGetValue(kvp.Key, out var idxMeta))
                                    {
                                        idxMeta.ColumnsDisplay = string.Join(", ", kvp.Value);
                                    }
                                }
                            }
                        }
                    }
                }

                return result;
            });
        }

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
    }
}
