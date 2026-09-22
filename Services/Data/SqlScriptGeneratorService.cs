using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using ExcelSupport.Models;
using Microsoft.Office.Interop.Excel;

namespace ExcelSupport.Services
{
    public static class SqlScriptGeneratorService
    {
        /// <summary>
        /// Đọc dữ liệu từ Range sang mảng 2 chiều 1-based [1..rows, 1..cols]
        /// </summary>
        public static object[,] GetRangeValues2D(Range rng)
        {
            if (rng == null) return new object[1, 1];

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

        /// <summary>
        /// Phân tích và tự động nhận diện danh sách cột và kiểu dữ liệu từ Range
        /// </summary>
        public static List<SqlColumnDefinition> DetectColumnsFromRange(Range rng, bool firstRowIsHeader)
        {
            var result = new List<SqlColumnDefinition>();
            if (rng == null) return result;

            object[,] values = GetRangeValues2D(rng);
            int rows = values.GetLength(0);
            int cols = values.GetLength(1);

            if (rows == 0 || cols == 0) return result;

            int dataStartRow = firstRowIsHeader ? 2 : 1;

            for (int c = 1; c <= cols; c++)
            {
                string headerName = string.Empty;
                if (firstRowIsHeader)
                {
                    object? hVal = values[1, c];
                    headerName = hVal?.ToString()?.Trim() ?? string.Empty;
                }

                if (string.IsNullOrWhiteSpace(headerName))
                {
                    headerName = $"COL{c}";
                }

                string sanitizedName = SanitizeColumnName(headerName);

                // Nhận diện kiểu dữ liệu mẫu từ các dòng dữ liệu (tối đa 50 dòng)
                SqlDataType detectedType = DetectColumnDataType(values, c, dataStartRow, rows);

                bool isPkCandidate = c == 1 && (
                    sanitizedName.EndsWith("_ID", StringComparison.OrdinalIgnoreCase) ||
                    sanitizedName.Equals("ID", StringComparison.OrdinalIgnoreCase) ||
                    sanitizedName.EndsWith("_CODE", StringComparison.OrdinalIgnoreCase) ||
                    sanitizedName.Equals("CODE", StringComparison.OrdinalIgnoreCase)
                );

                result.Add(new SqlColumnDefinition
                {
                    ColumnIndex = c,
                    OriginalHeader = headerName,
                    ColumnName = sanitizedName,
                    DataType = detectedType,
                    IsSelected = true,
                    IsPrimaryKey = isPkCandidate,
                    AllowNull = true
                });
            }

            return result;
        }

        /// <summary>
        /// Chuẩn hóa tên cột thành định dạng hợp lệ trong SQL (chữ hoa, thay khoảng trắng và ký tự lạ bằng dấu gạch dưới)
        /// </summary>
        public static string SanitizeColumnName(string rawName)
        {
            if (string.IsNullOrWhiteSpace(rawName)) return "COL";

            string s = rawName.Trim();
            // Thay thế ký tự khoảng trắng hoặc dấu phân cách bằng _
            s = Regex.Replace(s, @"[\s\.\,\;\:\-\/\\]+", "_");
            // Bỏ các ký tự đặc biệt không phải chữ, số hoặc _
            s = Regex.Replace(s, @"[^\w]", "");
            if (string.IsNullOrEmpty(s)) return "COL";

            // Nếu bắt đầu bằng số, thêm tiền tố COL_
            if (char.IsDigit(s[0]))
            {
                s = "COL_" + s;
            }

            return s.ToUpperInvariant();
        }

        /// <summary>
        /// Phân tích mẫu giá trị trong cột để đoán định kiểu dữ liệu SQL phù hợp nhất
        /// </summary>
        private static SqlDataType DetectColumnDataType(object[,] values, int colIndex, int startRow, int maxRows)
        {
            int checkLimit = Math.Min(startRow + 50, maxRows);
            int numberCount = 0;
            int dateCount = 0;
            int boolCount = 0;
            int sampleCount = 0;

            for (int r = startRow; r <= checkLimit; r++)
            {
                object? val = values[r, colIndex];
                if (val == null || string.IsNullOrWhiteSpace(val.ToString())) continue;

                sampleCount++;
                string s = val.ToString()!.Trim();

                if (val is bool || bool.TryParse(s, out _))
                {
                    boolCount++;
                    continue;
                }

                // Kiểm tra DateTime trong chuỗi
                if (DateTime.TryParse(s, CultureInfo.CurrentCulture, DateTimeStyles.None, out _) ||
                    DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.None, out _))
                {
                    // Nếu chuỗi chứa dấu / hoặc - hoặc : thì khả năng cao là ngày giờ
                    if (s.Contains('/') || s.Contains('-') || s.Contains(':'))
                    {
                        dateCount++;
                        continue;
                    }
                }

                // Kiểm tra số (nguyên hoặc thực)
                if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out double numVal) ||
                    double.TryParse(s, NumberStyles.Any, CultureInfo.CurrentCulture, out numVal))
                {
                    // Excel OADate thường là số từ 20000 đến 60000 (khoảng năm 1954 đến 2064)
                    // nhưng nếu không có dấu hiệu rõ ràng ta coi là Number
                    numberCount++;
                    continue;
                }
            }

            if (sampleCount == 0) return SqlDataType.Text;

            if (dateCount == sampleCount) return SqlDataType.Date;
            if (numberCount == sampleCount) return SqlDataType.Number;
            if (boolCount == sampleCount) return SqlDataType.Boolean;

            return SqlDataType.Text;
        }

        /// <summary>
        /// Sinh script SQL từ Range và các tùy chọn cấu hình
        /// </summary>
        public static SqlScriptGeneratorResult GenerateScript(Range rng, SqlScriptGeneratorOptions options)
        {
            if (rng == null)
            {
                return new SqlScriptGeneratorResult
                {
                    Success = false,
                    ErrorMessage = "Vui lòng chọn một vùng dữ liệu trên bảng tính Excel."
                };
            }

            object[,] rawValues = GetRangeValues2D(rng);
            return GenerateScriptFromValues(rawValues, options);
        }

        /// <summary>
        /// Sinh script SQL từ mảng 2 chiều dữ liệu và các tùy chọn
        /// </summary>
        public static SqlScriptGeneratorResult GenerateScriptFromValues(object[,] values, SqlScriptGeneratorOptions options)
        {
            var sw = Stopwatch.StartNew();
            int rows = values.GetLength(0);
            int cols = values.GetLength(1);

            if (rows == 0 || cols == 0)
            {
                return new SqlScriptGeneratorResult
                {
                    Success = false,
                    ErrorMessage = "Vùng chọn không chứa bất kỳ dữ liệu nào."
                };
            }

            var selectedCols = options.Columns.Where(c => c.IsSelected).ToList();
            if (selectedCols.Count == 0)
            {
                return new SqlScriptGeneratorResult
                {
                    Success = false,
                    ErrorMessage = "Vui lòng chọn ít nhất 1 cột để sinh script SQL."
                };
            }

            int dataStartRow = options.FirstRowIsHeader ? 2 : 1;
            int totalDataRows = rows - dataStartRow + 1;

            if (totalDataRows <= 0)
            {
                return new SqlScriptGeneratorResult
                {
                    Success = false,
                    ErrorMessage = "Vùng chọn chỉ có dòng tiêu đề, không có dòng dữ liệu nào bên dưới."
                };
            }

            // Kiểm tra yêu cầu khóa chính cho MERGE / UPDATE
            if (options.StatementType == SqlStatementType.MergeUpsert || options.StatementType == SqlStatementType.Update)
            {
                var pkCols = selectedCols.Where(c => c.IsPrimaryKey).ToList();
                if (pkCols.Count == 0)
                {
                    return new SqlScriptGeneratorResult
                    {
                        Success = false,
                        ErrorMessage = $"Cú pháp {options.StatementType} bắt buộc phải chọn ít nhất 1 cột Khóa chính (PK / Key) để làm điều kiện so khớp."
                    };
                }
            }

            string tableName = FormatTableName(options.SchemaName, options.TableName, options.Dialect);
            var sb = new StringBuilder(totalDataRows * selectedCols.Count * 25);
            int statementCount = 0;

            // Header chú thích đầu file script
            sb.AppendLine($"-- =============================================================================");
            sb.AppendLine($"-- Script Generated by ExcelSupport Add-in");
            sb.AppendLine($"-- Target Dialect: {options.Dialect}");
            sb.AppendLine($"-- Statement Type: {options.StatementType}");
            sb.AppendLine($"-- Target Table  : {tableName}");
            sb.AppendLine($"-- Total Rows    : {totalDataRows:N0} rows | Total Columns: {selectedCols.Count}");
            sb.AppendLine($"-- Generated At  : {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            sb.AppendLine($"-- =============================================================================");
            sb.AppendLine();

            // Transaction Start
            if (options.IncludeTransaction)
            {
                AppendTransactionStart(sb, options.Dialect);
            }

            // Identity Insert ON (SQL Server)
            if (options.Dialect == SqlDialect.SqlServer && options.IncludeIdentityInsert)
            {
                sb.AppendLine($"SET IDENTITY_INSERT {tableName} ON;");
                sb.AppendLine();
            }

            // Sinh nội dung chính theo Statement Type
            switch (options.StatementType)
            {
                case SqlStatementType.InsertSingle:
                    statementCount = GenerateInsertSingle(sb, values, dataStartRow, rows, selectedCols, tableName, options);
                    break;

                case SqlStatementType.InsertBatch:
                    statementCount = GenerateInsertBatch(sb, values, dataStartRow, rows, selectedCols, tableName, options);
                    break;

                case SqlStatementType.MergeUpsert:
                    statementCount = GenerateMergeUpsert(sb, values, dataStartRow, rows, selectedCols, tableName, options);
                    break;

                case SqlStatementType.Update:
                    statementCount = GenerateUpdate(sb, values, dataStartRow, rows, selectedCols, tableName, options);
                    break;
            }

            // Identity Insert OFF (SQL Server)
            if (options.Dialect == SqlDialect.SqlServer && options.IncludeIdentityInsert)
            {
                sb.AppendLine();
                sb.AppendLine($"SET IDENTITY_INSERT {tableName} OFF;");
            }

            // Transaction End
            if (options.IncludeTransaction)
            {
                AppendTransactionEnd(sb, options.Dialect);
            }

            sw.Stop();

            return new SqlScriptGeneratorResult
            {
                Success = true,
                SqlScript = sb.ToString(),
                RowCount = totalDataRows,
                ColumnCount = selectedCols.Count,
                StatementCount = statementCount,
                ExecutionTimeMs = sw.ElapsedMilliseconds
            };
        }

        #region Statement Generators

        private static int GenerateInsertSingle(
            StringBuilder sb,
            object[,] values,
            int startRow,
            int endRow,
            List<SqlColumnDefinition> cols,
            string tableName,
            SqlScriptGeneratorOptions options)
        {
            int count = 0;
            string colList = string.Join(", ", cols.Select(c => FormatColumnName(c.ColumnName, options.Dialect)));

            for (int r = startRow; r <= endRow; r++)
            {
                var valList = new List<string>(cols.Count);
                for (int i = 0; i < cols.Count; i++)
                {
                    var col = cols[i];
                    object? cellVal = values[r, col.ColumnIndex];
                    valList.Add(FormatSqlValue(cellVal, col, options));
                }

                sb.Append($"INSERT INTO {tableName} ({colList}) VALUES ({string.Join(", ", valList)});");
                sb.AppendLine();
                count++;
            }

            return count;
        }

        private static int GenerateInsertBatch(
            StringBuilder sb,
            object[,] values,
            int startRow,
            int endRow,
            List<SqlColumnDefinition> cols,
            string tableName,
            SqlScriptGeneratorOptions options)
        {
            int batchSize = Math.Max(1, options.BatchSize);
            int totalRows = endRow - startRow + 1;
            int count = 0;
            string colList = string.Join(", ", cols.Select(c => FormatColumnName(c.ColumnName, options.Dialect)));

            if (options.Dialect == SqlDialect.Oracle)
            {
                // Oracle sử dụng cú pháp: INSERT ALL INTO table (cols) VALUES (...) INTO ... SELECT 1 FROM DUAL;
                int currentInBatch = 0;
                bool batchOpened = false;

                for (int r = startRow; r <= endRow; r++)
                {
                    if (!batchOpened)
                    {
                        sb.AppendLine($"INSERT ALL");
                        batchOpened = true;
                        currentInBatch = 0;
                    }

                    var valList = new List<string>(cols.Count);
                    for (int i = 0; i < cols.Count; i++)
                    {
                        var col = cols[i];
                        object? cellVal = values[r, col.ColumnIndex];
                        valList.Add(FormatSqlValue(cellVal, col, options));
                    }

                    sb.AppendLine($"  INTO {tableName} ({colList}) VALUES ({string.Join(", ", valList)})");
                    currentInBatch++;

                    if (currentInBatch >= batchSize || r == endRow)
                    {
                        sb.AppendLine($"SELECT 1 FROM DUAL;");
                        sb.AppendLine();
                        batchOpened = false;
                        count++;
                    }
                }
            }
            else
            {
                // Các hệ quản trị khác: INSERT INTO table (cols) VALUES (...), (...);
                int currentInBatch = 0;
                bool batchOpened = false;

                for (int r = startRow; r <= endRow; r++)
                {
                    if (!batchOpened)
                    {
                        sb.AppendLine($"INSERT INTO {tableName} ({colList}) VALUES");
                        batchOpened = true;
                        currentInBatch = 0;
                    }

                    var valList = new List<string>(cols.Count);
                    for (int i = 0; i < cols.Count; i++)
                    {
                        var col = cols[i];
                        object? cellVal = values[r, col.ColumnIndex];
                        valList.Add(FormatSqlValue(cellVal, col, options));
                    }

                    currentInBatch++;
                    bool isLastInBatch = (currentInBatch >= batchSize) || (r == endRow);

                    if (isLastInBatch)
                    {
                        sb.AppendLine($"  ({string.Join(", ", valList)});");
                        sb.AppendLine();
                        batchOpened = false;
                        count++;
                    }
                    else
                    {
                        sb.AppendLine($"  ({string.Join(", ", valList)}),");
                    }
                }
            }

            return count;
        }

        private static int GenerateMergeUpsert(
            StringBuilder sb,
            object[,] values,
            int startRow,
            int endRow,
            List<SqlColumnDefinition> cols,
            string tableName,
            SqlScriptGeneratorOptions options)
        {
            var pkCols = cols.Where(c => c.IsPrimaryKey).ToList();
            var nonPkCols = cols.Where(c => !c.IsPrimaryKey).ToList();
            int count = 0;

            for (int r = startRow; r <= endRow; r++)
            {
                var valMap = new Dictionary<string, string>();
                foreach (var col in cols)
                {
                    object? cellVal = values[r, col.ColumnIndex];
                    valMap[col.ColumnName] = FormatSqlValue(cellVal, col, options);
                }

                switch (options.Dialect)
                {
                    case SqlDialect.Oracle:
                        GenerateOracleMergeRow(sb, tableName, cols, pkCols, nonPkCols, valMap);
                        break;

                    case SqlDialect.SqlServer:
                        GenerateSqlServerMergeRow(sb, tableName, cols, pkCols, nonPkCols, valMap);
                        break;

                    case SqlDialect.PostgreSql:
                        GeneratePostgresUpsertRow(sb, tableName, cols, pkCols, nonPkCols, valMap, options);
                        break;

                    case SqlDialect.MySql:
                        GenerateMySqlUpsertRow(sb, tableName, cols, pkCols, nonPkCols, valMap, options);
                        break;

                    case SqlDialect.Sqlite:
                        GenerateSqliteUpsertRow(sb, tableName, cols, pkCols, nonPkCols, valMap, options);
                        break;

                    default:
                        GenerateGenericMergeRow(sb, tableName, cols, pkCols, nonPkCols, valMap);
                        break;
                }

                count++;
            }

            return count;
        }

        private static int GenerateUpdate(
            StringBuilder sb,
            object[,] values,
            int startRow,
            int endRow,
            List<SqlColumnDefinition> cols,
            string tableName,
            SqlScriptGeneratorOptions options)
        {
            var pkCols = cols.Where(c => c.IsPrimaryKey).ToList();
            var nonPkCols = cols.Where(c => !c.IsPrimaryKey).ToList();
            int count = 0;

            if (nonPkCols.Count == 0)
            {
                sb.AppendLine("-- Tất cả các cột đều là Khóa chính, không có cột dữ liệu nào để cập nhật (UPDATE SET).");
                return 0;
            }

            for (int r = startRow; r <= endRow; r++)
            {
                var setClauses = new List<string>(nonPkCols.Count);
                foreach (var col in nonPkCols)
                {
                    object? cellVal = values[r, col.ColumnIndex];
                    string formattedVal = FormatSqlValue(cellVal, col, options);
                    setClauses.Add($"{FormatColumnName(col.ColumnName, options.Dialect)} = {formattedVal}");
                }

                var whereClauses = new List<string>(pkCols.Count);
                foreach (var col in pkCols)
                {
                    object? cellVal = values[r, col.ColumnIndex];
                    string formattedVal = FormatSqlValue(cellVal, col, options);
                    whereClauses.Add($"{FormatColumnName(col.ColumnName, options.Dialect)} = {formattedVal}");
                }

                sb.AppendLine($"UPDATE {tableName}");
                sb.AppendLine($"SET {string.Join(", ", setClauses)}");
                sb.AppendLine($"WHERE {string.Join(" AND ", whereClauses)};");
                sb.AppendLine();
                count++;
            }

            return count;
        }

        #endregion

        #region Dialect-Specific Upsert Helpers

        private static void GenerateOracleMergeRow(
            StringBuilder sb,
            string tableName,
            List<SqlColumnDefinition> allCols,
            List<SqlColumnDefinition> pkCols,
            List<SqlColumnDefinition> nonPkCols,
            Dictionary<string, string> valMap)
        {
            sb.AppendLine($"MERGE INTO {tableName} tgt");
            sb.Append("USING (SELECT ");
            sb.Append(string.Join(", ", allCols.Select(c => $"{valMap[c.ColumnName]} AS {c.ColumnName}")));
            sb.AppendLine(" FROM DUAL) src");

            string onClause = string.Join(" AND ", pkCols.Select(c => $"tgt.{c.ColumnName} = src.{c.ColumnName}"));
            sb.AppendLine($"ON ({onClause})");

            if (nonPkCols.Count > 0)
            {
                sb.AppendLine("WHEN MATCHED THEN");
                sb.AppendLine($"  UPDATE SET {string.Join(", ", nonPkCols.Select(c => $"tgt.{c.ColumnName} = src.{c.ColumnName}"))}");
            }

            sb.AppendLine("WHEN NOT MATCHED THEN");
            sb.AppendLine($"  INSERT ({string.Join(", ", allCols.Select(c => c.ColumnName))})");
            sb.AppendLine($"  VALUES ({string.Join(", ", allCols.Select(c => $"src.{c.ColumnName}"))});");
            sb.AppendLine();
        }

        private static void GenerateSqlServerMergeRow(
            StringBuilder sb,
            string tableName,
            List<SqlColumnDefinition> allCols,
            List<SqlColumnDefinition> pkCols,
            List<SqlColumnDefinition> nonPkCols,
            Dictionary<string, string> valMap)
        {
            sb.AppendLine($"MERGE INTO {tableName} AS tgt");
            sb.Append("USING (VALUES (");
            sb.Append(string.Join(", ", allCols.Select(c => valMap[c.ColumnName])));
            sb.Append(")) AS src (");
            sb.Append(string.Join(", ", allCols.Select(c => $"[{c.ColumnName}]")));
            sb.AppendLine(")");

            string onClause = string.Join(" AND ", pkCols.Select(c => $"tgt.[{c.ColumnName}] = src.[{c.ColumnName}]"));
            sb.AppendLine($"ON {onClause}");

            if (nonPkCols.Count > 0)
            {
                sb.AppendLine("WHEN MATCHED THEN");
                sb.AppendLine($"  UPDATE SET {string.Join(", ", nonPkCols.Select(c => $"tgt.[{c.ColumnName}] = src.[{c.ColumnName}]"))}");
            }

            sb.AppendLine("WHEN NOT MATCHED THEN");
            sb.AppendLine($"  INSERT ({string.Join(", ", allCols.Select(c => $"[{c.ColumnName}]"))})");
            sb.AppendLine($"  VALUES ({string.Join(", ", allCols.Select(c => $"src.[{c.ColumnName}]"))});");
            sb.AppendLine();
        }

        private static void GeneratePostgresUpsertRow(
            StringBuilder sb,
            string tableName,
            List<SqlColumnDefinition> allCols,
            List<SqlColumnDefinition> pkCols,
            List<SqlColumnDefinition> nonPkCols,
            Dictionary<string, string> valMap,
            SqlScriptGeneratorOptions options)
        {
            string colList = string.Join(", ", allCols.Select(c => FormatColumnName(c.ColumnName, options.Dialect)));
            string valList = string.Join(", ", allCols.Select(c => valMap[c.ColumnName]));
            string pkList = string.Join(", ", pkCols.Select(c => FormatColumnName(c.ColumnName, options.Dialect)));

            sb.Append($"INSERT INTO {tableName} ({colList}) VALUES ({valList})");
            sb.AppendLine();

            if (nonPkCols.Count > 0)
            {
                string updateSet = string.Join(", ", nonPkCols.Select(c =>
                {
                    string colFormatted = FormatColumnName(c.ColumnName, options.Dialect);
                    return $"{colFormatted} = EXCLUDED.{colFormatted}";
                }));
                sb.AppendLine($"ON CONFLICT ({pkList}) DO UPDATE SET {updateSet};");
            }
            else
            {
                sb.AppendLine($"ON CONFLICT ({pkList}) DO NOTHING;");
            }
            sb.AppendLine();
        }

        private static void GenerateMySqlUpsertRow(
            StringBuilder sb,
            string tableName,
            List<SqlColumnDefinition> allCols,
            List<SqlColumnDefinition> pkCols,
            List<SqlColumnDefinition> nonPkCols,
            Dictionary<string, string> valMap,
            SqlScriptGeneratorOptions options)
        {
            string colList = string.Join(", ", allCols.Select(c => FormatColumnName(c.ColumnName, options.Dialect)));
            string valList = string.Join(", ", allCols.Select(c => valMap[c.ColumnName]));

            sb.Append($"INSERT INTO {tableName} ({colList}) VALUES ({valList})");

            if (nonPkCols.Count > 0)
            {
                sb.AppendLine();
                string updateSet = string.Join(", ", nonPkCols.Select(c =>
                {
                    string colFormatted = FormatColumnName(c.ColumnName, options.Dialect);
                    return $"{colFormatted} = VALUES({colFormatted})";
                }));
                sb.AppendLine($"ON DUPLICATE KEY UPDATE {updateSet};");
            }
            else
            {
                sb.AppendLine(";");
            }
            sb.AppendLine();
        }

        private static void GenerateSqliteUpsertRow(
            StringBuilder sb,
            string tableName,
            List<SqlColumnDefinition> allCols,
            List<SqlColumnDefinition> pkCols,
            List<SqlColumnDefinition> nonPkCols,
            Dictionary<string, string> valMap,
            SqlScriptGeneratorOptions options)
        {
            string colList = string.Join(", ", allCols.Select(c => FormatColumnName(c.ColumnName, options.Dialect)));
            string valList = string.Join(", ", allCols.Select(c => valMap[c.ColumnName]));
            string pkList = string.Join(", ", pkCols.Select(c => FormatColumnName(c.ColumnName, options.Dialect)));

            sb.Append($"INSERT INTO {tableName} ({colList}) VALUES ({valList})");
            sb.AppendLine();

            if (nonPkCols.Count > 0)
            {
                string updateSet = string.Join(", ", nonPkCols.Select(c =>
                {
                    string colFormatted = FormatColumnName(c.ColumnName, options.Dialect);
                    return $"{colFormatted} = excluded.{colFormatted}";
                }));
                sb.AppendLine($"ON CONFLICT({pkList}) DO UPDATE SET {updateSet};");
            }
            else
            {
                sb.AppendLine($"ON CONFLICT({pkList}) DO NOTHING;");
            }
            sb.AppendLine();
        }

        private static void GenerateGenericMergeRow(
            StringBuilder sb,
            string tableName,
            List<SqlColumnDefinition> allCols,
            List<SqlColumnDefinition> pkCols,
            List<SqlColumnDefinition> nonPkCols,
            Dictionary<string, string> valMap)
        {
            sb.AppendLine($"MERGE INTO {tableName} tgt");
            sb.Append("USING (SELECT ");
            sb.Append(string.Join(", ", allCols.Select(c => $"{valMap[c.ColumnName]} AS {c.ColumnName}")));
            sb.AppendLine(") src");

            string onClause = string.Join(" AND ", pkCols.Select(c => $"tgt.{c.ColumnName} = src.{c.ColumnName}"));
            sb.AppendLine($"ON ({onClause})");

            if (nonPkCols.Count > 0)
            {
                sb.AppendLine("WHEN MATCHED THEN");
                sb.AppendLine($"  UPDATE SET {string.Join(", ", nonPkCols.Select(c => $"tgt.{c.ColumnName} = src.{c.ColumnName}"))}");
            }

            sb.AppendLine("WHEN NOT MATCHED THEN");
            sb.AppendLine($"  INSERT ({string.Join(", ", allCols.Select(c => c.ColumnName))})");
            sb.AppendLine($"  VALUES ({string.Join(", ", allCols.Select(c => $"src.{c.ColumnName}"))});");
            sb.AppendLine();
        }

        #endregion

        #region Formatting & Escaping Utilities

        public static string FormatSqlValue(object? rawVal, SqlColumnDefinition col, SqlScriptGeneratorOptions options)
        {
            if (rawVal == null) return "NULL";

            string sVal = rawVal.ToString()?.Trim() ?? string.Empty;

            if (string.IsNullOrEmpty(sVal))
            {
                return options.TreatBlankAsNull ? "NULL" : "''";
            }

            if (options.TreatBlankAsNull && sVal.Equals("NULL", StringComparison.OrdinalIgnoreCase))
            {
                return "NULL";
            }

            switch (col.DataType)
            {
                case SqlDataType.Number:
                    if (double.TryParse(sVal, NumberStyles.Any, CultureInfo.InvariantCulture, out double dVal) ||
                        double.TryParse(sVal, NumberStyles.Any, CultureInfo.CurrentCulture, out dVal))
                    {
                        return dVal.ToString(CultureInfo.InvariantCulture);
                    }
                    return EscapeTextValue(sVal, options);

                case SqlDataType.Boolean:
                    bool bVal = false;
                    if (rawVal is bool b) bVal = b;
                    else if (bool.TryParse(sVal, out bool parsedB)) bVal = parsedB;
                    else if (sVal == "1" || sVal.Equals("TRUE", StringComparison.OrdinalIgnoreCase) || sVal.Equals("YES", StringComparison.OrdinalIgnoreCase)) bVal = true;

                    return options.Dialect switch
                    {
                        SqlDialect.PostgreSql => bVal ? "TRUE" : "FALSE",
                        _ => bVal ? "1" : "0"
                    };

                case SqlDataType.Date:
                case SqlDataType.Timestamp:
                    DateTime dt = ExtractDateTime(rawVal, sVal);
                    if (dt == DateTime.MinValue)
                    {
                        return EscapeTextValue(sVal, options);
                    }
                    return FormatDateTimeLiteral(dt, col.DataType == SqlDataType.Timestamp, options.Dialect);

                case SqlDataType.Raw:
                    return sVal;

                case SqlDataType.Text:
                case SqlDataType.Auto:
                default:
                    return EscapeTextValue(sVal, options);
            }
        }

        private static DateTime ExtractDateTime(object rawVal, string sVal)
        {
            if (rawVal is DateTime dtDirect) return dtDirect;

            // Xử lý Excel OADate (số ngày kể từ 1900-01-01)
            if (rawVal is double d && d > 1.0 && d < 2958465.0)
            {
                try { return DateTime.FromOADate(d); } catch { }
            }

            if (double.TryParse(sVal, NumberStyles.Any, CultureInfo.InvariantCulture, out double parsedD) && parsedD > 1.0 && parsedD < 2958465.0)
            {
                try { return DateTime.FromOADate(parsedD); } catch { }
            }

            if (DateTime.TryParse(sVal, CultureInfo.CurrentCulture, DateTimeStyles.None, out DateTime dtParsed))
            {
                return dtParsed;
            }

            if (DateTime.TryParse(sVal, CultureInfo.InvariantCulture, DateTimeStyles.None, out dtParsed))
            {
                return dtParsed;
            }

            return DateTime.MinValue;
        }

        private static string FormatDateTimeLiteral(DateTime dt, bool isTimestamp, SqlDialect dialect)
        {
            bool hasTime = dt.Hour != 0 || dt.Minute != 0 || dt.Second != 0 || dt.Millisecond != 0 || isTimestamp;
            string dateFmt = hasTime ? dt.ToString("yyyy-MM-dd HH:mm:ss") : dt.ToString("yyyy-MM-dd");

            return dialect switch
            {
                SqlDialect.Oracle => hasTime
                    ? $"TO_DATE('{dateFmt}', 'YYYY-MM-DD HH24:MI:SS')"
                    : $"TO_DATE('{dateFmt}', 'YYYY-MM-DD')",

                SqlDialect.PostgreSql => hasTime
                    ? $"'{dateFmt}'::timestamp"
                    : $"'{dateFmt}'::date",

                SqlDialect.SqlServer => $"'{dateFmt}'",
                SqlDialect.MySql => $"'{dateFmt}'",
                SqlDialect.Sqlite => $"'{dateFmt}'",
                _ => $"'{dateFmt}'"
            };
        }

        private static string EscapeTextValue(string text, SqlScriptGeneratorOptions options)
        {
            // Escape ký tự nháy đơn SQL chuẩn ' -> ''
            string escaped = text.Replace("'", "''");

            bool useNPrefix = options.UseUnicodeNPrefix && (options.Dialect == SqlDialect.SqlServer || options.Dialect == SqlDialect.Oracle);
            return useNPrefix ? $"N'{escaped}'" : $"'{escaped}'";
        }

        public static string FormatTableName(string? schema, string table, SqlDialect dialect)
        {
            string cleanTable = string.IsNullOrWhiteSpace(table) ? "MY_TABLE" : table.Trim();
            string cleanSchema = schema?.Trim() ?? string.Empty;

            if (string.IsNullOrEmpty(cleanSchema))
            {
                return FormatIdentifier(cleanTable, dialect);
            }

            return $"{FormatIdentifier(cleanSchema, dialect)}.{FormatIdentifier(cleanTable, dialect)}";
        }

        public static string FormatColumnName(string colName, SqlDialect dialect)
        {
            return FormatIdentifier(colName, dialect);
        }

        private static string FormatIdentifier(string identifier, SqlDialect dialect)
        {
            return dialect switch
            {
                SqlDialect.Oracle => identifier.ToUpperInvariant(),
                SqlDialect.SqlServer => $"[{identifier}]",
                SqlDialect.PostgreSql => $"\"{identifier.ToLowerInvariant()}\"",
                SqlDialect.MySql => $"`{identifier}`",
                SqlDialect.Sqlite => $"\"{identifier}\"",
                _ => identifier
            };
        }

        private static void AppendTransactionStart(StringBuilder sb, SqlDialect dialect)
        {
            switch (dialect)
            {
                case SqlDialect.Oracle:
                    sb.AppendLine("-- Transaction Start (Oracle uses implicit transactions)");
                    sb.AppendLine();
                    break;
                case SqlDialect.SqlServer:
                    sb.AppendLine("BEGIN TRANSACTION;");
                    sb.AppendLine();
                    break;
                case SqlDialect.PostgreSql:
                    sb.AppendLine("BEGIN;");
                    sb.AppendLine();
                    break;
                case SqlDialect.MySql:
                    sb.AppendLine("START TRANSACTION;");
                    sb.AppendLine();
                    break;
                case SqlDialect.Sqlite:
                    sb.AppendLine("BEGIN TRANSACTION;");
                    sb.AppendLine();
                    break;
                default:
                    sb.AppendLine("BEGIN TRANSACTION;");
                    sb.AppendLine();
                    break;
            }
        }

        private static void AppendTransactionEnd(StringBuilder sb, SqlDialect dialect)
        {
            sb.AppendLine();
            sb.AppendLine("COMMIT;");
        }

        #endregion
    }
}
