using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelSupport.Models;
using Microsoft.Office.Interop.Excel;
using Oracle.ManagedDataAccess.Client;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using ExcelRange = Microsoft.Office.Interop.Excel.Range;
using ExcelWorksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace ExcelSupport.Services
{
    public static class OracleDataCompareService
    {
        #region Connection Testing & Schema Metadata Discovery

        public static async Task<(bool Success, string Message, string ServerVersion)> TestConnectionAsync(OracleConnectionConfig config)
        {
            string connStr = config.BuildConnectionString();
            return await Task.Run(() =>
            {
                try
                {
                    using (var conn = new OracleConnection(connStr))
                    {
                        conn.Open();
                        string version = conn.ServerVersion;
                        return (true, "Kết nối thành công!", version);
                    }
                }
                catch (OracleException oex)
                {
                    return (false, $"Lỗi Oracle (ORA-{oex.Number}): {oex.Message}", string.Empty);
                }
                catch (Exception ex)
                {
                    return (false, $"Lỗi kết nối: {ex.Message}", string.Empty);
                }
            });
        }

        public static async Task<List<string>> GetSchemasAsync(OracleConnectionConfig config)
        {
            string connStr = config.BuildConnectionString();
            return await Task.Run(() =>
            {
                var schemas = new List<string>();
                try
                {
                    using (var conn = new OracleConnection(connStr))
                    {
                        conn.Open();
                        string sql = "SELECT username FROM all_users ORDER BY username";
                        using (var cmd = new OracleCommand(sql, conn))
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                schemas.Add(reader.GetString(0));
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[GetSchemasAsync] Error: {ex.Message}");
                }
                return schemas;
            });
        }

        public static async Task<List<string>> GetTablesAndViewsAsync(OracleConnectionConfig config, string schema)
        {
            string connStr = config.BuildConnectionString();
            return await Task.Run(() =>
            {
                var tables = new List<string>();
                if (string.IsNullOrWhiteSpace(schema)) return tables;

                try
                {
                    using (var conn = new OracleConnection(connStr))
                    {
                        conn.Open();
                        string sql = @"
                            SELECT table_name FROM all_tables WHERE owner = :owner
                            UNION
                            SELECT view_name FROM all_views WHERE owner = :owner
                            ORDER BY 1";

                        using (var cmd = new OracleCommand(sql, conn))
                        {
                            cmd.Parameters.Add(new OracleParameter("owner", schema.ToUpperInvariant()));
                            using (var reader = cmd.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    tables.Add(reader.GetString(0));
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[GetTablesAndViewsAsync] Error: {ex.Message}");
                }
                return tables;
            });
        }

        public static async Task<List<OracleTableColumnInfo>> GetTableColumnsAsync(OracleConnectionConfig config, string schema, string tableName)
        {
            string connStr = config.BuildConnectionString();
            return await Task.Run(() =>
            {
                var columns = new List<OracleTableColumnInfo>();
                if (string.IsNullOrWhiteSpace(schema) || string.IsNullOrWhiteSpace(tableName)) return columns;

                try
                {
                    using (var conn = new OracleConnection(connStr))
                    {
                        conn.Open();

                        // 1. Lấy Primary Keys của bảng
                        var pkCols = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                        string pkSql = @"
                            SELECT cols.column_name
                            FROM all_constraints cons
                            JOIN all_cons_columns cols 
                              ON cons.constraint_name = cols.constraint_name 
                             AND cons.owner = cols.owner
                            WHERE cons.constraint_type = 'P'
                              AND cons.owner = :owner
                              AND cons.table_name = :tname
                            ORDER BY cols.position";

                        using (var pkCmd = new OracleCommand(pkSql, conn))
                        {
                            pkCmd.Parameters.Add(new OracleParameter("owner", schema.ToUpperInvariant()));
                            pkCmd.Parameters.Add(new OracleParameter("tname", tableName.ToUpperInvariant()));
                            using (var reader = pkCmd.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    pkCols.Add(reader.GetString(0));
                                }
                            }
                        }

                        // 2. Lấy thông tin cấu trúc cột
                        string colSql = @"
                            SELECT column_name, data_type, data_length, nullable, column_id
                            FROM all_tab_columns
                            WHERE owner = :owner AND table_name = :tname
                            ORDER BY column_id";

                        using (var colCmd = new OracleCommand(colSql, conn))
                        {
                            colCmd.Parameters.Add(new OracleParameter("owner", schema.ToUpperInvariant()));
                            colCmd.Parameters.Add(new OracleParameter("tname", tableName.ToUpperInvariant()));
                            using (var reader = colCmd.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    string colName = reader.GetString(0);
                                    string dataType = reader.GetString(1);
                                    int dataLen = Convert.ToInt32(reader.GetValue(2));
                                    string nullableStr = reader.GetString(3);

                                    bool isPk = pkCols.Contains(colName);

                                    columns.Add(new OracleTableColumnInfo
                                    {
                                        ColumnName = colName,
                                        DataType = dataType,
                                        DataLength = dataLen,
                                        Nullable = (nullableStr == "Y"),
                                        IsPrimaryKey = isPk,
                                        IsSelectedKey = isPk,
                                        IsSelectedCompare = true
                                    });
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[GetTableColumnsAsync] Error: {ex.Message}");
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
            IProgress<(string StatusText, double ProgressPercent)>? progress = null,
            string connectionNameA = "",
            string connectionNameB = "")
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
                ConnectionNameA = connectionNameA,
                SchemaB = schemaB,
                TableB = tableB,
                ConnectionNameB = connectionNameB,
                TotalRowsA = dtA.Rows.Count,
                TotalRowsB = dtB.Rows.Count,
                Options = options
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

            // Thực thi thuật toán so sánh
            if (options.Mode == OracleCompareMode.ByKeyColumns && options.SelectedKeyColumns.Count > 0)
            {
                CompareByKeyColumns(dtA, dtB, result, options, progress);
            }
            else
            {
                CompareSequentially(dtA, dtB, result, options, progress);
            }

            stopwatch.Stop();
            result.ExecutionTime = stopwatch.Elapsed;

            progress?.Report(($"Hoàn tất đối soát trong {result.ExecutionTime.TotalSeconds:F2}s.", 100));
            return result;
        }

        private static void CompareByKeyColumns(
            System.Data.DataTable dtA,
            System.Data.DataTable dtB,
            OracleCompareResult result,
            OracleCompareOptions options,
            IProgress<(string StatusText, double ProgressPercent)>? progress)
        {
            var keyCols = options.SelectedKeyColumns;

            // 1. Lập chỉ mục Dict cho Bảng B
            var dictB = new Dictionary<string, DataRow>(StringComparer.Ordinal);

            foreach (DataRow rowB in dtB.Rows)
            {
                string key = BuildRowKey(rowB, keyCols, options);
                if (!dictB.ContainsKey(key))
                {
                    dictB.Add(key, rowB);
                }
            }

            var processedKeysB = new HashSet<string>(StringComparer.Ordinal);
            int rowIdx = 1;
            int totalRowsA = dtA.Rows.Count;

            // 2. So khớp từng dòng từ Bảng A sang Bảng B
            for (int i = 0; i < totalRowsA; i++)
            {
                DataRow rowA = dtA.Rows[i];
                string key = BuildRowKey(rowA, keyCols, options);

                var diffItem = new OracleRowDiffItem
                {
                    RowNumber = rowIdx++,
                    KeyDisplay = key
                };

                // Nạp giá trị dòng A
                foreach (DataColumn col in dtA.Columns)
                {
                    diffItem.RowValuesA[col.ColumnName] = rowA[col];
                }

                if (dictB.TryGetValue(key, out DataRow? rowB))
                {
                    processedKeysB.Add(key);

                    // Nạp giá trị dòng B
                    foreach (DataColumn col in dtB.Columns)
                    {
                        diffItem.RowValuesB[col.ColumnName] = rowB[col];
                    }

                    // So sánh từng cột
                    bool hasDiff = false;
                    foreach (var colName in result.Columns)
                    {
                        object? valA = rowA.Table.Columns.Contains(colName) ? rowA[colName] : DBNull.Value;
                        object? valB = rowB.Table.Columns.Contains(colName) ? rowB[colName] : DBNull.Value;

                        bool isColDiff = IsValueDifferent(valA, valB, options);
                        if (isColDiff)
                        {
                            hasDiff = true;
                            diffItem.DifferingColumns.Add(colName);
                        }

                        diffItem.CellDiffs.Add(new OracleCellDiff
                        {
                            ColumnName = colName,
                            ValueA = valA,
                            ValueB = valB,
                            IsDifferent = isColDiff
                        });
                    }

                    diffItem.Status = hasDiff ? OracleRowDiffStatus.Modified : OracleRowDiffStatus.Identical;
                }
                else
                {
                    // Có trong A nhưng thiếu trong B
                    diffItem.Status = OracleRowDiffStatus.MissingInB;
                    foreach (var colName in result.Columns)
                    {
                        object? valA = rowA.Table.Columns.Contains(colName) ? rowA[colName] : DBNull.Value;
                        diffItem.CellDiffs.Add(new OracleCellDiff
                        {
                            ColumnName = colName,
                            ValueA = valA,
                            ValueB = DBNull.Value,
                            IsDifferent = true
                        });
                    }
                }

                result.DiffItems.Add(diffItem);

                if (i % 500 == 0 && totalRowsA > 0)
                {
                    double pct = 60 + ((double)i / totalRowsA) * 30;
                    progress?.Report(($"Đang đối soát bản ghi {i:N0} / {totalRowsA:N0}...", pct));
                }
            }

            // 3. Tìm các dòng chỉ có trong Bảng B mà không có trong A
            foreach (DataRow rowB in dtB.Rows)
            {
                string key = BuildRowKey(rowB, keyCols, options);
                if (!processedKeysB.Contains(key))
                {
                    var diffItem = new OracleRowDiffItem
                    {
                        RowNumber = rowIdx++,
                        KeyDisplay = key,
                        Status = OracleRowDiffStatus.MissingInA
                    };

                    foreach (DataColumn col in dtB.Columns)
                    {
                        diffItem.RowValuesB[col.ColumnName] = rowB[col];
                    }

                    foreach (var colName in result.Columns)
                    {
                        object? valB = rowB.Table.Columns.Contains(colName) ? rowB[colName] : DBNull.Value;
                        diffItem.CellDiffs.Add(new OracleCellDiff
                        {
                            ColumnName = colName,
                            ValueA = DBNull.Value,
                            ValueB = valB,
                            IsDifferent = true
                        });
                    }

                    result.DiffItems.Add(diffItem);
                }
            }
        }

        private static void CompareSequentially(
            System.Data.DataTable dtA,
            System.Data.DataTable dtB,
            OracleCompareResult result,
            OracleCompareOptions options,
            IProgress<(string StatusText, double ProgressPercent)>? progress)
        {
            int maxCount = Math.Max(dtA.Rows.Count, dtB.Rows.Count);

            for (int i = 0; i < maxCount; i++)
            {
                DataRow? rowA = i < dtA.Rows.Count ? dtA.Rows[i] : null;
                DataRow? rowB = i < dtB.Rows.Count ? dtB.Rows[i] : null;

                var diffItem = new OracleRowDiffItem
                {
                    RowNumber = i + 1,
                    KeyDisplay = $"Dòng {i + 1}"
                };

                if (rowA != null && rowB != null)
                {
                    foreach (DataColumn col in dtA.Columns) diffItem.RowValuesA[col.ColumnName] = rowA[col];
                    foreach (DataColumn col in dtB.Columns) diffItem.RowValuesB[col.ColumnName] = rowB[col];

                    bool hasDiff = false;
                    foreach (var colName in result.Columns)
                    {
                        object? valA = rowA.Table.Columns.Contains(colName) ? rowA[colName] : DBNull.Value;
                        object? valB = rowB.Table.Columns.Contains(colName) ? rowB[colName] : DBNull.Value;

                        bool isColDiff = IsValueDifferent(valA, valB, options);
                        if (isColDiff)
                        {
                            hasDiff = true;
                            diffItem.DifferingColumns.Add(colName);
                        }

                        diffItem.CellDiffs.Add(new OracleCellDiff
                        {
                            ColumnName = colName,
                            ValueA = valA,
                            ValueB = valB,
                            IsDifferent = isColDiff
                        });
                    }

                    diffItem.Status = hasDiff ? OracleRowDiffStatus.Modified : OracleRowDiffStatus.Identical;
                }
                else if (rowA != null)
                {
                    diffItem.Status = OracleRowDiffStatus.MissingInB;
                    foreach (DataColumn col in dtA.Columns) diffItem.RowValuesA[col.ColumnName] = rowA[col];
                    foreach (var colName in result.Columns)
                    {
                        object? valA = rowA.Table.Columns.Contains(colName) ? rowA[colName] : DBNull.Value;
                        diffItem.CellDiffs.Add(new OracleCellDiff
                        {
                            ColumnName = colName,
                            ValueA = valA,
                            ValueB = DBNull.Value,
                            IsDifferent = true
                        });
                    }
                }
                else if (rowB != null)
                {
                    diffItem.Status = OracleRowDiffStatus.MissingInA;
                    foreach (DataColumn col in dtB.Columns) diffItem.RowValuesB[col.ColumnName] = rowB[col];
                    foreach (var colName in result.Columns)
                    {
                        object? valB = rowB.Table.Columns.Contains(colName) ? rowB[colName] : DBNull.Value;
                        diffItem.CellDiffs.Add(new OracleCellDiff
                        {
                            ColumnName = colName,
                            ValueA = DBNull.Value,
                            ValueB = valB,
                            IsDifferent = true
                        });
                    }
                }

                result.DiffItems.Add(diffItem);

                if (i % 500 == 0 && maxCount > 0)
                {
                    double pct = 60 + ((double)i / maxCount) * 30;
                    progress?.Report(($"Đang đối soát thứ tự dòng {i:N0} / {maxCount:N0}...", pct));
                }
            }
        }

        private static string BuildRowKey(DataRow row, List<string> keyCols, OracleCompareOptions options)
        {
            var sb = new StringBuilder();
            for (int i = 0; i < keyCols.Count; i++)
            {
                if (i > 0) sb.Append("|");
                string colName = keyCols[i];
                if (row.Table.Columns.Contains(colName))
                {
                    object val = row[colName];
                    string str = FormatValueForCompare(val, options);
                    sb.Append(str);
                }
            }
            return sb.ToString();
        }

        private static async Task<System.Data.DataTable> FetchTableDataAsync(
            OracleConnectionConfig config,
            string schema,
            string tableName,
            OracleCompareOptions options,
            bool isTableA)
        {
            string connStr = config.BuildConnectionString();
            return await Task.Run(() =>
            {
                var dt = new System.Data.DataTable();
                using (var conn = new OracleConnection(connStr))
                {
                    conn.Open();

                    string sql;
                    if (options.UseCustomQuery)
                    {
                        sql = isTableA ? options.CustomQueryA : options.CustomQueryB;
                    }
                    else
                    {
                        string where = isTableA ? options.WhereClauseA : options.WhereClauseB;
                        string whereSql = string.IsNullOrWhiteSpace(where) ? "" : $"WHERE {where.Trim()}";

                        string fullTable = string.IsNullOrWhiteSpace(schema) ? tableName : $"\"{schema.ToUpperInvariant()}\".\"{tableName.ToUpperInvariant()}\"";

                        if (options.MaxRows > 0)
                        {
                            sql = $"SELECT * FROM (SELECT * FROM {fullTable} {whereSql}) WHERE ROWNUM <= {options.MaxRows}";
                        }
                        else
                        {
                            sql = $"SELECT * FROM {fullTable} {whereSql}";
                        }
                    }

                    using (var cmd = new OracleCommand(sql, conn))
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

        private static string FormatValueForCompare(object? val, OracleCompareOptions options)
        {
            if (val == null || val is DBNull) return "";

            string str;
            if (val is DateTime dt) str = dt.ToString("yyyy-MM-dd HH:mm:ss");
            else if (val is double d) str = d.ToString("G15");
            else if (val is decimal dec) str = dec.ToString("G29");
            else if (val is float f) str = f.ToString("G7");
            else str = val.ToString() ?? "";

            if (options.TrimStrings) str = str.Trim();
            if (options.IgnoreWhitespace) str = System.Text.RegularExpressions.Regex.Replace(str, @"\s+", " ").Trim();
            if (options.IgnoreCase) str = str.ToLowerInvariant();

            return str;
        }

        public static bool IsValueDifferent(object? valA, object? valB, OracleCompareOptions options)
        {
            bool isNullA = valA == null || valA is DBNull;
            bool isNullB = valB == null || valB is DBNull;

            if (isNullA && isNullB) return false;

            if (options.TreatNullAsEmpty)
            {
                string sA = isNullA ? "" : valA!.ToString()!.Trim();
                string sB = isNullB ? "" : valB!.ToString()!.Trim();
                if (string.IsNullOrEmpty(sA) && string.IsNullOrEmpty(sB)) return false;
            }

            if (isNullA != isNullB) return true;

            if (options.NumericTolerance > 0 && double.TryParse(valA?.ToString(), out double dA) && double.TryParse(valB?.ToString(), out double dB))
            {
                return Math.Abs(dA - dB) > options.NumericTolerance;
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

                var wb = app.ActiveWorkbook ?? app.Workbooks.Add();

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

                RenderDiffDataToWorksheet(activeSheet, activeCell.Row, activeCell.Column, result, highlightOnlyDiffs);
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
            if (result.Options?.ReportLayout == OracleReportLayout.SideBySide)
            {
                RenderSideBySideDiff(ws, startRow, startCol, result, highlightOnlyDiffs);
            }
            else
            {
                RenderStackedTopBottomDiff(ws, startRow, startCol, result, highlightOnlyDiffs);
            }
        }

        private static void RenderStackedTopBottomDiff(ExcelWorksheet ws, int startRow, int startCol, OracleCompareResult result, bool highlightOnlyDiffs)
        {
            var filterItems = highlightOnlyDiffs
                ? result.DiffItems.Where(r => r.Status != OracleRowDiffStatus.Identical).ToList()
                : result.DiffItems;

            if (result.Columns.Count == 0) return;

            Color highlightColor = ParseColorHex(result.Options?.HighlightColorHex, Color.FromArgb(239, 68, 68));
            Color headerBgColor = Color.FromArgb(110, 231, 183);
            Color headerTextColor = Color.FromArgb(6, 78, 59);
            Color borderColor = Color.FromArgb(156, 163, 175);

            string labelA = string.IsNullOrWhiteSpace(result.ConnectionNameA) ? result.TableA : $"{result.TableA}({result.ConnectionNameA})";
            string labelB = string.IsNullOrWhiteSpace(result.ConnectionNameB) ? result.TableB : $"{result.TableB}({result.ConnectionNameB})";

            int curRow = RenderSingleTableBlock(ws, startRow, startCol, labelA, Color.FromArgb(37, 99, 235), headerBgColor, headerTextColor, borderColor, highlightColor, result.Columns, filterItems, isTableA: true);

            curRow++; // Dòng trống ngăn cách

            curRow = RenderSingleTableBlock(ws, curRow, startCol, labelB, Color.FromArgb(190, 24, 93), headerBgColor, headerTextColor, borderColor, highlightColor, result.Columns, filterItems, isTableA: false);

            ws.Range[ws.Cells[startRow, startCol], ws.Cells[curRow, startCol + result.Columns.Count - 1]].Columns.AutoFit();
        }

        private static int RenderSingleTableBlock(
            ExcelWorksheet ws,
            int curRow,
            int startCol,
            string title,
            Color titleColor,
            Color headerBgColor,
            Color headerTextColor,
            Color borderColor,
            Color highlightColor,
            List<string> columns,
            List<OracleRowDiffItem> items,
            bool isTableA)
        {
            int numCols = columns.Count;
            int rowCount = items.Count;

            // 1. Tiêu đề bảng
            ExcelRange titleRange = ws.Cells[curRow, startCol];
            titleRange.Value2 = title;
            titleRange.Font.Bold = true;
            titleRange.Font.Size = 11;
            titleRange.Font.Color = ColorTranslator.ToOle(titleColor);
            curRow++;

            // 2. Dòng Header cột
            int headerRow = curRow;
            object[,] headerArray = new object[1, numCols];
            for (int c = 0; c < numCols; c++) headerArray[0, c] = columns[c];

            ExcelRange headerRange = ws.Range[ws.Cells[headerRow, startCol], ws.Cells[headerRow, startCol + numCols - 1]];
            headerRange.Value2 = headerArray;
            headerRange.Font.Bold = true;
            headerRange.Font.Size = 10;
            headerRange.Font.Color = ColorTranslator.ToOle(headerTextColor);
            headerRange.Interior.Color = ColorTranslator.ToOle(headerBgColor);
            headerRange.HorizontalAlignment = -4108; // xlCenter
            curRow++;

            // 3. Dữ liệu & Tô màu sai khác
            if (rowCount > 0)
            {
                int dataStartRow = curRow;
                object[,] dataArray = new object[rowCount, numCols];

                for (int i = 0; i < rowCount; i++)
                {
                    var dict = isTableA ? items[i].RowValuesA : items[i].RowValuesB;
                    for (int c = 0; c < numCols; c++)
                    {
                        dict.TryGetValue(columns[c], out var val);
                        dataArray[i, c] = FormatValueDisplay(val);
                    }
                }

                ExcelRange dataRange = ws.Range[ws.Cells[dataStartRow, startCol], ws.Cells[dataStartRow + rowCount - 1, startCol + numCols - 1]];
                dataRange.Value2 = dataArray;

                // Tô màu sai khác
                var missingStatus = isTableA ? OracleRowDiffStatus.MissingInB : OracleRowDiffStatus.MissingInA;
                Color missingBg = isTableA ? Color.FromArgb(254, 226, 226) : Color.FromArgb(219, 234, 254);

                for (int i = 0; i < rowCount; i++)
                {
                    var item = items[i];
                    int r = dataStartRow + i;

                    if (item.Status == missingStatus)
                    {
                        ws.Range[ws.Cells[r, startCol], ws.Cells[r, startCol + numCols - 1]].Interior.Color = ColorTranslator.ToOle(missingBg);
                    }
                    else if (item.Status == OracleRowDiffStatus.Modified)
                    {
                        foreach (var diffCol in item.DifferingColumns)
                        {
                            int colIdx = columns.IndexOf(diffCol);
                            if (colIdx >= 0)
                            {
                                ExcelRange cell = ws.Cells[r, startCol + colIdx];
                                cell.Interior.Color = ColorTranslator.ToOle(highlightColor);
                                cell.Font.Bold = true;
                            }
                        }
                    }
                }

                // Kẻ viền bảng
                ExcelRange fullTable = ws.Range[ws.Cells[headerRow, startCol], ws.Cells[dataStartRow + rowCount - 1, startCol + numCols - 1]];
                fullTable.Borders.LineStyle = 1;
                fullTable.Borders.Color = ColorTranslator.ToOle(borderColor);

                curRow = dataStartRow + rowCount;
            }

            return curRow;
        }

        private static void RenderSideBySideDiff(ExcelWorksheet ws, int startRow, int startCol, OracleCompareResult result, bool highlightOnlyDiffs)
        {
            var filterItems = highlightOnlyDiffs
                ? result.DiffItems.Where(r => r.Status != OracleRowDiffStatus.Identical).ToList()
                : result.DiffItems;

            int curRow = startRow;

            // Tiêu đề
            ExcelRange titleRange = ws.Range[ws.Cells[curRow, startCol], ws.Cells[curRow, startCol + 4]];
            titleRange.Merge();
            titleRange.Value2 = "BÁO CÁO ĐỐI SOÁT DỮ LIỆU BẢNG ORACLE (SIDE-BY-SIDE)";
            titleRange.Font.Bold = true;
            titleRange.Font.Size = 14;
            titleRange.Font.Color = ColorTranslator.ToOle(Color.FromArgb(30, 41, 59));
            curRow += 2;

            // Thống kê tổng quan
            ws.Cells[curRow, startCol].Value2 = "Bảng DB A:";
            ws.Cells[curRow, startCol + 1].Value2 = $"{result.SchemaA}.{result.TableA} ({result.TotalRowsA:N0} dòng)";
            ws.Cells[curRow, startCol + 3].Value2 = "Trùng khớp:";
            ws.Cells[curRow, startCol + 4].Value2 = $"{result.MatchCount:N0} dòng";
            curRow++;

            ws.Cells[curRow, startCol].Value2 = "Bảng DB B:";
            ws.Cells[curRow, startCol + 1].Value2 = $"{result.SchemaB}.{result.TableB} ({result.TotalRowsB:N0} dòng)";
            ws.Cells[curRow, startCol + 3].Value2 = "Sai lệch:";
            ws.Cells[curRow, startCol + 4].Value2 = $"{result.ModifiedCount:N0} dòng";
            curRow += 2;

            var headers = new List<string> { "STT", "Khóa / Bản ghi", "Trạng Thái", "Cột Sai Khác" };
            foreach (var col in result.Columns)
            {
                headers.Add($"{col} (A)");
                headers.Add($"{col} (B)");
            }

            int headerRow = curRow;
            int totalCols = headers.Count;
            object[,] headerArray = new object[1, totalCols];
            for (int c = 0; c < totalCols; c++) headerArray[0, c] = headers[c];

            ExcelRange headerRange = ws.Range[ws.Cells[headerRow, startCol], ws.Cells[headerRow, startCol + totalCols - 1]];
            headerRange.Value2 = headerArray;
            headerRange.Font.Bold = true;
            headerRange.Font.Color = ColorTranslator.ToOle(Color.White);
            headerRange.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(30, 41, 59));
            headerRange.HorizontalAlignment = -4108;
            curRow++;

            int dataStartRow = curRow;
            int rowCount = filterItems.Count;
            Color highlightColor = ParseColorHex(result.Options?.HighlightColorHex, Color.FromArgb(254, 240, 138));

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

                for (int i = 0; i < rowCount; i++)
                {
                    var item = filterItems[i];
                    int r = dataStartRow + i;

                    if (item.Status == OracleRowDiffStatus.MissingInB)
                    {
                        ws.Range[ws.Cells[r, startCol], ws.Cells[r, startCol + totalCols - 1]].Interior.Color = ColorTranslator.ToOle(Color.FromArgb(254, 226, 226));
                    }
                    else if (item.Status == OracleRowDiffStatus.MissingInA)
                    {
                        ws.Range[ws.Cells[r, startCol], ws.Cells[r, startCol + totalCols - 1]].Interior.Color = ColorTranslator.ToOle(Color.FromArgb(219, 234, 254));
                    }
                    else if (item.Status == OracleRowDiffStatus.Modified)
                    {
                        int colIdx = 4;
                        foreach (var col in result.Columns)
                        {
                            if (item.DifferingColumns.Contains(col))
                            {
                                ws.Cells[r, startCol + colIdx].Interior.Color = ColorTranslator.ToOle(highlightColor);
                                ws.Cells[r, startCol + colIdx + 1].Interior.Color = ColorTranslator.ToOle(highlightColor);
                            }
                            colIdx += 2;
                        }
                    }
                }

                ExcelRange fullTableRange = ws.Range[ws.Cells[headerRow, startCol], ws.Cells[dataStartRow + rowCount - 1, startCol + totalCols - 1]];
                fullTableRange.Borders.LineStyle = 1;
                fullTableRange.Borders.Color = ColorTranslator.ToOle(Color.FromArgb(203, 213, 225));
            }

            ExcelRange allColsRange = ws.Range[ws.Cells[headerRow, startCol], ws.Cells[headerRow + rowCount + 2, startCol + totalCols - 1]];
            allColsRange.Columns.AutoFit();
        }

        private static Color ParseColorHex(string? hex, Color defaultColor)
        {
            if (string.IsNullOrWhiteSpace(hex)) return defaultColor;
            try { return ColorTranslator.FromHtml(hex); }
            catch { return defaultColor; }
        }

        private static string FormatValueDisplay(object? val)
        {
            if (val == null || val is DBNull) return "";
            if (val is DateTime dt) return dt.ToString("yyyy-MM-dd HH:mm:ss");
            return val.ToString() ?? "";
        }

        private static bool SheetExists(Microsoft.Office.Interop.Excel.Workbook wb, string sheetName)
        {
            return wb.Worksheets.Cast<ExcelWorksheet>().Any(s => string.Equals(s.Name, sheetName, StringComparison.OrdinalIgnoreCase));
        }

        #endregion
    }
}
