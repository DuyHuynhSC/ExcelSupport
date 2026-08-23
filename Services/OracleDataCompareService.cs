using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using ExcelSupport.Models;
using Oracle.ManagedDataAccess.Client;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using ExcelWorksheet = Microsoft.Office.Interop.Excel._Worksheet;
using ExcelRange = Microsoft.Office.Interop.Excel.Range;

namespace ExcelSupport.Services
{
    public static class OracleDataCompareService
    {
        #region Database Metadata & Connection Testing

        public static async Task<(bool Success, string Message, string ServerVersion)> TestConnectionAsync(OracleConnectionConfig config)
        {
            return await Task.Run(() =>
            {
                try
                {
                    string connStr = config.BuildConnectionString();
                    using var conn = new OracleConnection(connStr);
                    conn.Open();

                    string version = conn.ServerVersion ?? "Oracle Database";
                    string banner = version;

                    try
                    {
                        using var cmd = conn.CreateCommand();
                        cmd.CommandText = "SELECT BANNER FROM V$VERSION WHERE ROWNUM = 1";
                        var result = cmd.ExecuteScalar();
                        if (result != null && result != DBNull.Value)
                        {
                            banner = result.ToString() ?? version;
                        }
                    }
                    catch { }

                    return (true, "Kết nối thành công!", banner);
                }
                catch (Exception ex)
                {
                    return (false, $"Lỗi kết nối: {ex.Message}", string.Empty);
                }
            });
        }

        public static async Task<List<string>> GetSchemasAsync(OracleConnectionConfig config)
        {
            return await Task.Run(() =>
            {
                var schemas = new List<string>();
                try
                {
                    string connStr = config.BuildConnectionString();
                    using var conn = new OracleConnection(connStr);
                    conn.Open();

                    using var cmd = conn.CreateCommand();
                    cmd.CommandText = "SELECT USERNAME FROM ALL_USERS ORDER BY USERNAME";
                    using var reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        schemas.Add(reader.GetString(0));
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"GetSchemas error: {ex.Message}");
                }
                return schemas;
            });
        }

        public static async Task<List<string>> GetTablesAndViewsAsync(OracleConnectionConfig config, string schema)
        {
            return await Task.Run(() =>
            {
                var tables = new List<string>();
                try
                {
                    string connStr = config.BuildConnectionString();
                    using var conn = new OracleConnection(connStr);
                    conn.Open();

                    using var cmd = conn.CreateCommand();
                    if (!string.IsNullOrWhiteSpace(schema))
                    {
                        cmd.CommandText = @"
                            SELECT TABLE_NAME FROM ALL_TABLES WHERE OWNER = :pOwner
                            UNION
                            SELECT VIEW_NAME AS TABLE_NAME FROM ALL_VIEWS WHERE OWNER = :pOwner
                            ORDER BY 1";
                        cmd.Parameters.Add(new OracleParameter("pOwner", schema.Trim().ToUpperInvariant()));
                    }
                    else
                    {
                        cmd.CommandText = @"
                            SELECT TABLE_NAME FROM USER_TABLES
                            UNION
                            SELECT VIEW_NAME AS TABLE_NAME FROM USER_VIEWS
                            ORDER BY 1";
                    }

                    using var reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        tables.Add(reader.GetString(0));
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"GetTablesAndViews error: {ex.Message}");
                }
                return tables;
            });
        }

        public static async Task<List<OracleTableColumnInfo>> GetTableColumnsAsync(OracleConnectionConfig config, string schema, string tableName)
        {
            return await Task.Run(() =>
            {
                var columns = new List<OracleTableColumnInfo>();
                try
                {
                    string connStr = config.BuildConnectionString();
                    using var conn = new OracleConnection(connStr);
                    conn.Open();

                    string owner = string.IsNullOrWhiteSpace(schema) ? config.Username.Trim().ToUpperInvariant() : schema.Trim().ToUpperInvariant();
                    string tbl = tableName.Trim().ToUpperInvariant();

                    // 1. Get Primary Key Columns
                    var pkCols = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    try
                    {
                        using var pkCmd = conn.CreateCommand();
                        pkCmd.CommandText = @"
                            SELECT cols.COLUMN_NAME
                            FROM ALL_CONSTRAINTS cons
                            JOIN ALL_CONS_COLUMNS cols ON cons.CONSTRAINT_NAME = cols.CONSTRAINT_NAME AND cons.OWNER = cols.OWNER
                            WHERE cons.CONSTRAINT_TYPE = 'P'
                              AND cons.OWNER = :pOwner
                              AND cons.TABLE_NAME = :pTable
                            ORDER BY cols.POSITION";
                        pkCmd.Parameters.Add(new OracleParameter("pOwner", owner));
                        pkCmd.Parameters.Add(new OracleParameter("pTable", tbl));

                        using var pkReader = pkCmd.ExecuteReader();
                        while (pkReader.Read())
                        {
                            pkCols.Add(pkReader.GetString(0));
                        }
                    }
                    catch { }

                    // 2. Get Column Definitions
                    using var colCmd = conn.CreateCommand();
                    colCmd.CommandText = @"
                        SELECT COLUMN_NAME, DATA_TYPE, DATA_LENGTH, NULLABLE
                        FROM ALL_TAB_COLUMNS
                        WHERE OWNER = :pOwner AND TABLE_NAME = :pTable
                        ORDER BY COLUMN_ID";
                    colCmd.Parameters.Add(new OracleParameter("pOwner", owner));
                    colCmd.Parameters.Add(new OracleParameter("pTable", tbl));

                    using var colReader = colCmd.ExecuteReader();
                    while (colReader.Read())
                    {
                        string colName = colReader.GetString(0);
                        string dataType = colReader.GetString(1);
                        int dataLength = Convert.ToInt32(colReader.GetValue(2));
                        string nullable = colReader.GetString(3);

                        bool isPk = pkCols.Contains(colName);
                        columns.Add(new OracleTableColumnInfo
                        {
                            ColumnName = colName,
                            DataType = dataType,
                            DataLength = dataLength,
                            Nullable = (nullable == "Y"),
                            IsPrimaryKey = isPk,
                            IsSelectedKey = isPk,
                            IsSelectedCompare = true
                        });
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"GetTableColumns error: {ex.Message}");
                }
                return columns;
            });
        }

        #endregion

        #region Data Retrieval & Comparison Engine

        public static async Task<OracleCompareResult> CompareTablesAsync(
            OracleConnectionConfig configA,
            OracleConnectionConfig configB,
            string schemaA,
            string tableA,
            string schemaB,
            string tableB,
            OracleCompareOptions options,
            IProgress<(string StatusText, double ProgressPercent)>? progress = null)
        {
            var stopwatch = System.Diagnostics.Stopwatch.StartNew();

            progress?.Report(("Đang kết nối và nạp dữ liệu từ Database A...", 10));
            var dataTableATask = FetchTableDataAsync(configA, schemaA, tableA, options, isTableA: true);

            progress?.Report(("Đang kết nối và nạp dữ liệu từ Database B...", 25));
            var dataTableBTask = FetchTableDataAsync(configB, schemaB, tableB, options, isTableA: false);

            await Task.WhenAll(dataTableATask, dataTableBTask);

            var dtA = await dataTableATask;
            var dtB = await dataTableBTask;

            progress?.Report(("Đang tiến hành thuật toán đối soát dữ liệu...", 60));

            var result = new OracleCompareResult
            {
                SchemaA = schemaA,
                TableA = tableA,
                SchemaB = schemaB,
                TableB = tableB,
                TotalRowsA = dtA.Rows.Count,
                TotalRowsB = dtB.Rows.Count
            };

            // Xác định danh sách cột chung cần so sánh
            var allColsA = dtA.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToList();
            var allColsB = dtB.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToList();
            var commonCols = allColsA.Intersect(allColsB, StringComparer.OrdinalIgnoreCase).ToList();

            // Nếu user có chọn danh sách cột so sánh cụ thể
            var compareCols = (options.SelectedCompareColumns.Count > 0)
                ? commonCols.Intersect(options.SelectedCompareColumns, StringComparer.OrdinalIgnoreCase).ToList()
                : commonCols;

            if (compareCols.Count == 0)
            {
                compareCols = commonCols;
            }

            result.Columns = compareCols;
            result.KeyColumns = options.SelectedKeyColumns;

            var diffItems = new List<OracleRowDiffItem>();

            if (options.Mode == OracleCompareMode.ByKeyColumns && options.SelectedKeyColumns.Count > 0)
            {
                // SO SÁNH THEO KHÓA CHÍNH (PRIMARY KEY / COMPOSITE KEY)
                var keyCols = options.SelectedKeyColumns.Intersect(commonCols, StringComparer.OrdinalIgnoreCase).ToList();
                if (keyCols.Count == 0)
                {
                    keyCols = new List<string> { commonCols.First() };
                }

                var mapA = new Dictionary<string, DataRow>(StringComparer.OrdinalIgnoreCase);
                foreach (DataRow row in dtA.Rows)
                {
                    string k = BuildKeyString(row, keyCols, options);
                    if (!mapA.ContainsKey(k)) mapA[k] = row;
                }

                var mapB = new Dictionary<string, DataRow>(StringComparer.OrdinalIgnoreCase);
                foreach (DataRow row in dtB.Rows)
                {
                    string k = BuildKeyString(row, keyCols, options);
                    if (!mapB.ContainsKey(k)) mapB[k] = row;
                }

                var allKeys = new HashSet<string>(mapA.Keys, StringComparer.OrdinalIgnoreCase);
                allKeys.UnionWith(mapB.Keys);

                int rowNum = 1;
                int totalKeys = allKeys.Count;
                int processed = 0;

                foreach (var k in allKeys)
                {
                    processed++;
                    if (processed % 500 == 0)
                    {
                        double p = 60 + (35.0 * processed / totalKeys);
                        progress?.Report(($"Đang so khớp bản ghi ({processed}/{totalKeys})...", p));
                    }

                    bool inA = mapA.TryGetValue(k, out var rowA);
                    bool inB = mapB.TryGetValue(k, out var rowB);

                    var item = new OracleRowDiffItem
                    {
                        RowNumber = rowNum++,
                        KeyDisplay = k
                    };

                    if (inA && inB)
                    {
                        PopulateRowData(item.RowValuesA, rowA!, commonCols);
                        PopulateRowData(item.RowValuesB, rowB!, commonCols);

                        bool hasDiff = false;
                        foreach (var col in compareCols)
                        {
                            object? valA = rowA![col];
                            object? valB = rowB![col];

                            bool isColDiff = IsValueDifferent(valA, valB, options);
                            if (isColDiff)
                            {
                                hasDiff = true;
                                item.DifferingColumns.Add(col);
                            }

                            item.CellDiffs.Add(new OracleCellDiff
                            {
                                ColumnName = col,
                                ValueA = valA,
                                ValueB = valB,
                                IsDifferent = isColDiff
                            });
                        }

                        item.Status = hasDiff ? OracleRowDiffStatus.Modified : OracleRowDiffStatus.Identical;
                    }
                    else if (inA && !inB)
                    {
                        PopulateRowData(item.RowValuesA, rowA!, commonCols);
                        item.Status = OracleRowDiffStatus.MissingInB; // Chỉ có ở A, thiếu ở B

                        foreach (var col in compareCols)
                        {
                            item.CellDiffs.Add(new OracleCellDiff
                            {
                                ColumnName = col,
                                ValueA = rowA![col],
                                ValueB = null,
                                IsDifferent = true
                            });
                            item.DifferingColumns.Add(col);
                        }
                    }
                    else if (!inA && inB)
                    {
                        PopulateRowData(item.RowValuesB, rowB!, commonCols);
                        item.Status = OracleRowDiffStatus.MissingInA; // Chỉ có ở B, thiếu ở A

                        foreach (var col in compareCols)
                        {
                            item.CellDiffs.Add(new OracleCellDiff
                            {
                                ColumnName = col,
                                ValueA = null,
                                ValueB = rowB![col],
                                IsDifferent = true
                            });
                            item.DifferingColumns.Add(col);
                        }
                    }

                    diffItems.Add(item);
                }
            }
            else
            {
                // SO SÁNH THEO THỨ TỰ BẢN GHI (SEQUENTIAL ORDER)
                int maxCount = Math.Max(dtA.Rows.Count, dtB.Rows.Count);
                for (int i = 0; i < maxCount; i++)
                {
                    if (i % 500 == 0)
                    {
                        double p = 60 + (35.0 * (i + 1) / maxCount);
                        progress?.Report(($"Đang so khớp bản ghi theo thứ tự ({i + 1}/{maxCount})...", p));
                    }

                    DataRow? rowA = (i < dtA.Rows.Count) ? dtA.Rows[i] : null;
                    DataRow? rowB = (i < dtB.Rows.Count) ? dtB.Rows[i] : null;

                    var item = new OracleRowDiffItem
                    {
                        RowNumber = i + 1,
                        KeyDisplay = $"Dòng #{i + 1}"
                    };

                    if (rowA != null && rowB != null)
                    {
                        PopulateRowData(item.RowValuesA, rowA, commonCols);
                        PopulateRowData(item.RowValuesB, rowB, commonCols);

                        bool hasDiff = false;
                        foreach (var col in compareCols)
                        {
                            object? valA = rowA[col];
                            object? valB = rowB[col];

                            bool isColDiff = IsValueDifferent(valA, valB, options);
                            if (isColDiff)
                            {
                                hasDiff = true;
                                item.DifferingColumns.Add(col);
                            }

                            item.CellDiffs.Add(new OracleCellDiff
                            {
                                ColumnName = col,
                                ValueA = valA,
                                ValueB = valB,
                                IsDifferent = isColDiff
                            });
                        }

                        item.Status = hasDiff ? OracleRowDiffStatus.Modified : OracleRowDiffStatus.Identical;
                    }
                    else if (rowA != null && rowB == null)
                    {
                        PopulateRowData(item.RowValuesA, rowA, commonCols);
                        item.Status = OracleRowDiffStatus.MissingInB;
                        foreach (var col in compareCols)
                        {
                            item.CellDiffs.Add(new OracleCellDiff
                            {
                                ColumnName = col,
                                ValueA = rowA[col],
                                ValueB = null,
                                IsDifferent = true
                            });
                            item.DifferingColumns.Add(col);
                        }
                    }
                    else if (rowA == null && rowB != null)
                    {
                        PopulateRowData(item.RowValuesB, rowB, commonCols);
                        item.Status = OracleRowDiffStatus.MissingInA;
                        foreach (var col in compareCols)
                        {
                            item.CellDiffs.Add(new OracleCellDiff
                            {
                                ColumnName = col,
                                ValueA = null,
                                ValueB = rowB[col],
                                IsDifferent = true
                            });
                            item.DifferingColumns.Add(col);
                        }
                    }

                    diffItems.Add(item);
                }
            }

            result.DiffItems = diffItems;
            stopwatch.Stop();
            result.ExecutionTime = stopwatch.Elapsed;

            progress?.Report(("Hoàn tất đối soát dữ liệu!", 100));
            return result;
        }

        private static async Task<DataTable> FetchTableDataAsync(OracleConnectionConfig config, string schema, string tableName, OracleCompareOptions options, bool isTableA)
        {
            return await Task.Run(() =>
            {
                var dt = new DataTable();
                try
                {
                    string connStr = config.BuildConnectionString();
                    using var conn = new OracleConnection(connStr);
                    conn.Open();

                    using var cmd = conn.CreateCommand();

                    if (options.UseCustomQuery)
                    {
                        string customSql = isTableA ? options.CustomQueryA : options.CustomQueryB;
                        if (string.IsNullOrWhiteSpace(customSql)) customSql = $"SELECT * FROM {tableName}";
                        cmd.CommandText = customSql;
                    }
                    else
                    {
                        string fullTableName = !string.IsNullOrWhiteSpace(schema) ? $"{schema.Trim()}.{tableName.Trim()}" : tableName.Trim();
                        string whereClause = isTableA ? options.WhereClauseA : options.WhereClauseB;

                        var sb = new StringBuilder();
                        sb.Append($"SELECT * FROM {fullTableName}");
                        if (!string.IsNullOrWhiteSpace(whereClause))
                        {
                            string w = whereClause.Trim();
                            if (!w.StartsWith("WHERE", StringComparison.OrdinalIgnoreCase))
                            {
                                sb.Append($" WHERE {w}");
                            }
                            else
                            {
                                sb.Append($" {w}");
                            }
                        }

                        if (options.MaxRows > 0)
                        {
                            // Oracle 12c+ FETCH FIRST syntax
                            sb.Append($" FETCH FIRST {options.MaxRows} ROWS ONLY");
                        }

                        cmd.CommandText = sb.ToString();
                    }

                    using var adapter = new OracleDataAdapter(cmd);
                    adapter.Fill(dt);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"FetchTableData error: {ex.Message}");
                    throw new Exception($"Không thể tải dữ liệu từ Database {(isTableA ? "A" : "B")}:\n{ex.Message}", ex);
                }
                return dt;
            });
        }

        private static string BuildKeyString(DataRow row, List<string> keyCols, OracleCompareOptions options)
        {
            var parts = new List<string>();
            foreach (var col in keyCols)
            {
                object? val = row[col];
                string s = FormatValueForCompare(val, options);
                parts.Add(s);
            }
            return string.Join(" | ", parts);
        }

        private static void PopulateRowData(Dictionary<string, object?> target, DataRow row, List<string> columns)
        {
            foreach (var col in columns)
            {
                if (row.Table.Columns.Contains(col))
                {
                    target[col] = row[col];
                }
            }
        }

        private static string FormatValueForCompare(object? val, OracleCompareOptions options)
        {
            if (val == null || val is DBNull)
            {
                return options.TreatNullAsEmpty ? "" : "<NULL>";
            }

            if (val is DateTime dt)
            {
                return dt.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
            }

            string s = val.ToString() ?? "";
            if (options.TrimStrings) s = s.Trim();
            if (options.IgnoreWhitespace) s = System.Text.RegularExpressions.Regex.Replace(s, @"\s+", " ");
            if (options.IgnoreCase) s = s.ToLowerInvariant();
            return s;
        }

        public static bool IsValueDifferent(object? valA, object? valB, OracleCompareOptions options)
        {
            bool isNullA = valA == null || valA is DBNull || (options.TreatNullAsEmpty && string.IsNullOrEmpty(valA.ToString()?.Trim()));
            bool isNullB = valB == null || valB is DBNull || (options.TreatNullAsEmpty && string.IsNullOrEmpty(valB.ToString()?.Trim()));

            if (isNullA && isNullB) return false;
            if (isNullA || isNullB) return true;

            // Numeric comparison with tolerance
            if (options.NumericTolerance > 0 &&
                double.TryParse(valA!.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out double numA) &&
                double.TryParse(valB!.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out double numB))
            {
                return Math.Abs(numA - numB) > options.NumericTolerance;
            }

            string strA = FormatValueForCompare(valA, options);
            string strB = FormatValueForCompare(valB, options);

            return !string.Equals(strA, strB, options.IgnoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal);
        }

        #endregion

        #region Excel Export & Active Cell Insertion with Color Highlighting

        public static void ExportDiffReportToExcel(OracleCompareResult result, ExcelApp? app, bool highlightOnlyDiffs = false)
        {
            if (app == null) return;

            try
            {
                app.ScreenUpdating = false;
                app.DisplayAlerts = false;

                var wb = app.ActiveWorkbook;
                if (wb == null)
                {
                    wb = app.Workbooks.Add();
                }

                // Tạo sheet mới
                string baseSheetName = "Oracle_Diff_Report";
                string sheetName = baseSheetName;
                int counter = 1;
                while (SheetExists(wb, sheetName))
                {
                    sheetName = $"{baseSheetName}_{counter++}";
                }

                ExcelWorksheet ws = wb.Worksheets.Add();
                ws.Name = sheetName;

                RenderDiffDataToWorksheet(ws, 1, 1, result, highlightOnlyDiffs);

                ws.Activate();
                app.ScreenUpdating = true;
            }
            catch (Exception ex)
            {
                app.ScreenUpdating = true;
                throw new Exception($"Lỗi xuất báo cáo Excel: {ex.Message}", ex);
            }
            finally
            {
                app.ScreenUpdating = true;
            }
        }

        public static void InsertDiffToActiveSelection(OracleCompareResult result, ExcelApp? app, bool highlightOnlyDiffs = false)
        {
            if (app == null) return;

            try
            {
                app.ScreenUpdating = false;
                var activeSheet = app.ActiveSheet as ExcelWorksheet;
                var activeCell = app.ActiveCell;

                if (activeSheet == null || activeCell == null)
                {
                    ExportDiffReportToExcel(result, app, highlightOnlyDiffs);
                    return;
                }

                int startRow = activeCell.Row;
                int startCol = activeCell.Column;

                RenderDiffDataToWorksheet(activeSheet, startRow, startCol, result, highlightOnlyDiffs);

                app.ScreenUpdating = true;
            }
            catch (Exception ex)
            {
                app.ScreenUpdating = true;
                throw new Exception($"Lỗi chèn vào vị trí đang chọn: {ex.Message}", ex);
            }
            finally
            {
                app.ScreenUpdating = true;
            }
        }

        private static void RenderDiffDataToWorksheet(ExcelWorksheet ws, int startRow, int startCol, OracleCompareResult result, bool highlightOnlyDiffs)
        {
            var filterItems = highlightOnlyDiffs
                ? result.DiffItems.Where(r => r.Status != OracleRowDiffStatus.Identical).ToList()
                : result.DiffItems;

            int curRow = startRow;

            // 1. BANNER TIÊU ĐỀ
            ExcelRange titleRange = ws.Range[ws.Cells[curRow, startCol], ws.Cells[curRow, startCol + 4]];
            titleRange.Merge();
            titleRange.Value2 = "BÁO CÁO ĐỐI SOÁT DỮ LIỆU BẢNG ORACLE (ORACLE TABLE DIFF REPORT)";
            titleRange.Font.Bold = true;
            titleRange.Font.Size = 14;
            titleRange.Font.Color = ColorTranslator.ToOle(Color.FromArgb(30, 41, 59));
            curRow += 2;

            // 2. THỐNG KÊ TỔNG QUAN
            ws.Cells[curRow, startCol].Value2 = "Bảng DB A (Gốc):";
            ws.Cells[curRow, startCol + 1].Value2 = $"{result.SchemaA}.{result.TableA} ({result.TotalRowsA:N0} dòng)";
            ws.Cells[curRow, startCol + 3].Value2 = "Trùng khớp:";
            ws.Cells[curRow, startCol + 4].Value2 = $"{result.MatchCount:N0} dòng";
            curRow++;

            ws.Cells[curRow, startCol].Value2 = "Bảng DB B (Đối chiếu):";
            ws.Cells[curRow, startCol + 1].Value2 = $"{result.SchemaB}.{result.TableB} ({result.TotalRowsB:N0} dòng)";
            ws.Cells[curRow, startCol + 3].Value2 = "Sai lệch giá trị:";
            ws.Cells[curRow, startCol + 4].Value2 = $"{result.ModifiedCount:N0} dòng";
            curRow++;

            ws.Cells[curRow, startCol].Value2 = "Thời gian xử lý:";
            ws.Cells[curRow, startCol + 1].Value2 = $"{result.ExecutionTime.TotalSeconds:F2} giây";
            ws.Cells[curRow, startCol + 3].Value2 = "Chỉ có ở A / B:";
            ws.Cells[curRow, startCol + 4].Value2 = $"-{result.MissingInBCount:N0} / +{result.MissingInACount:N0}";
            curRow += 2;

            // 3. XÂY DỰNG BẢNG DỮ LIỆU ĐỐI SOÁT
            // Cột: [STT] [Khóa / Vị Trí] [Trạng Thái] [Cột Sai Lệch] + Mỗi cột so sánh gồm: [Cột (DB A)] [Cột (DB B)]
            var headers = new List<string> { "STT", "Khóa / Bản ghi", "Trạng Thái", "Cột Sai Khác" };
            foreach (var col in result.Columns)
            {
                headers.Add($"{col} (DB A)");
                headers.Add($"{col} (DB B)");
            }

            int headerRow = curRow;
            int totalCols = headers.Count;

            // Ghi Header
            object[,] headerArray = new object[1, totalCols];
            for (int c = 0; c < totalCols; c++)
            {
                headerArray[0, c] = headers[c];
            }

            ExcelRange headerRange = ws.Range[ws.Cells[headerRow, startCol], ws.Cells[headerRow, startCol + totalCols - 1]];
            headerRange.Value2 = headerArray;
            headerRange.Font.Bold = true;
            headerRange.Font.Color = ColorTranslator.ToOle(Color.White);
            headerRange.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(30, 41, 59)); // Slate 800
            headerRange.HorizontalAlignment = -4108; // xlCenter

            curRow++;
            int dataStartRow = curRow;
            int rowCount = filterItems.Count;

            if (rowCount > 0)
            {
                object[,] dataArray = new object[rowCount, totalCols];
                for (int i = 0; i < rowCount; i++)
                {
                    var item = filterItems[i];
                    dataArray[i, 0] = item.RowNumber;
                    dataArray[i, 1] = item.KeyDisplay;
                    dataArray[i, 2] = item.StatusBadge;
                    dataArray[i, 3] = item.DifferingColumnsSummary;

                    int colIdx = 4;
                    foreach (var col in result.Columns)
                    {
                        item.RowValuesA.TryGetValue(col, out var valA);
                        item.RowValuesB.TryGetValue(col, out var valB);

                        dataArray[i, colIdx] = FormatValueDisplay(valA);
                        dataArray[i, colIdx + 1] = FormatValueDisplay(valB);
                        colIdx += 2;
                    }
                }

                ExcelRange dataRange = ws.Range[ws.Cells[dataStartRow, startCol], ws.Cells[dataStartRow + rowCount - 1, startCol + totalCols - 1]];
                dataRange.Value2 = dataArray;

                // TÔ MÀU TRỰC QUAN CÁC Ô SAI LỆCH VÀ BẢN GHI
                // Palette:
                int colorModifiedCell = ColorTranslator.ToOle(Color.FromArgb(254, 240, 138)); // Vàng chanh sáng
                int colorModifiedText = ColorTranslator.ToOle(Color.FromArgb(180, 83, 9));    // Cam đậm
                int colorMissingA = ColorTranslator.ToOle(Color.FromArgb(254, 226, 226));     // Đỏ nhạt (Chỉ có ở A)
                int colorMissingB = ColorTranslator.ToOle(Color.FromArgb(219, 234, 254));     // Xanh dương nhạt (Chỉ có ở B)

                for (int i = 0; i < rowCount; i++)
                {
                    var item = filterItems[i];
                    int r = dataStartRow + i;

                    if (item.Status == OracleRowDiffStatus.MissingInB)
                    {
                        // Chỉ có ở A -> Tô màu đỏ nhạt cho toàn dòng
                        ExcelRange rowRange = ws.Range[ws.Cells[r, startCol], ws.Cells[r, startCol + totalCols - 1]];
                        rowRange.Interior.Color = colorMissingA;
                    }
                    else if (item.Status == OracleRowDiffStatus.MissingInA)
                    {
                        // Chỉ có ở B -> Tô màu xanh nhạt cho toàn dòng
                        ExcelRange rowRange = ws.Range[ws.Cells[r, startCol], ws.Cells[r, startCol + totalCols - 1]];
                        rowRange.Interior.Color = colorMissingB;
                    }
                    else if (item.Status == OracleRowDiffStatus.Modified)
                    {
                        // Dòng có ô sửa đổi -> Tô màu đúng tại các ô có giá trị khác nhau
                        int colIdx = 4;
                        foreach (var col in result.Columns)
                        {
                            if (item.DifferingColumns.Contains(col))
                            {
                                ExcelRange cellA = ws.Cells[r, startCol + colIdx];
                                ExcelRange cellB = ws.Cells[r, startCol + colIdx + 1];

                                cellA.Interior.Color = colorModifiedCell;
                                cellA.Font.Bold = true;
                                cellA.Font.Color = colorModifiedText;

                                cellB.Interior.Color = colorModifiedCell;
                                cellB.Font.Bold = true;
                                cellB.Font.Color = colorModifiedText;
                            }
                            colIdx += 2;
                        }
                    }
                }

                // Kẻ viền bảng (Borders)
                ExcelRange fullTableRange = ws.Range[ws.Cells[headerRow, startCol], ws.Cells[dataStartRow + rowCount - 1, startCol + totalCols - 1]];
                fullTableRange.Borders.LineStyle = 1; // xlContinuous
                fullTableRange.Borders.Color = ColorTranslator.ToOle(Color.FromArgb(203, 213, 225)); // Slate 300
            }

            // Tự động căn chỉnh độ rộng cột
            ExcelRange allColsRange = ws.Range[ws.Cells[headerRow, startCol], ws.Cells[headerRow + rowCount + 2, startCol + totalCols - 1]];
            allColsRange.Columns.AutoFit();
        }

        private static string FormatValueDisplay(object? val)
        {
            if (val == null || val is DBNull) return "";
            if (val is DateTime dt) return dt.ToString("yyyy-MM-dd HH:mm:ss");
            return val.ToString() ?? "";
        }

        private static bool SheetExists(Microsoft.Office.Interop.Excel.Workbook wb, string sheetName)
        {
            foreach (ExcelWorksheet ws in wb.Worksheets)
            {
                if (string.Equals(ws.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
            return false;
        }

        #endregion
    }
}
