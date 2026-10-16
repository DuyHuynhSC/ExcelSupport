using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using ExcelSupport.Models;
using Microsoft.Office.Interop.Excel;

namespace ExcelSupport.Services
{
    public class OracleSpecToPlSqlService
    {
        public static object[,]? GetRangeValues2D(Range range)
        {
            if (range == null) return null;

            object val = range.Value2;
            if (val is object[,] arr)
            {
                return arr;
            }

            if (val != null)
            {
                object[,] singleArr = new object[2, 2];
                singleArr[1, 1] = val;
                return singleArr;
            }

            return null;
        }

        public static OracleSpecParseResult ParseRangeSpec(object[,]? data, string fallbackSheetName)
        {
            var result = new OracleSpecParseResult();
            if (data == null)
            {
                result.Success = false;
                result.ErrorMessage = "Dữ liệu vùng chọn trống.";
                return result;
            }

            int rowCount = data.GetLength(0);
            int colCount = data.GetLength(1);

            if (rowCount < 1 || colCount < 1)
            {
                result.Success = false;
                result.ErrorMessage = "Vùng chọn không có dòng hoặc cột hợp lệ.";
                return result;
            }

            // Gợi ý tên bảng mặc định từ tên Sheet
            string defaultTableName = SanitizeIdentifier(fallbackSheetName);
            if (string.IsNullOrWhiteSpace(defaultTableName))
            {
                defaultTableName = "MY_TABLE";
            }
            result.DetectedTableName = defaultTableName;

            // 1. Quét tìm Table Name hoặc Logical Table Name ở các dòng trên cùng (nếu có)
            for (int r = 1; r <= Math.Min(5, rowCount); r++)
            {
                for (int c = 1; c <= colCount; c++)
                {
                    string cellStr = data[r, c]?.ToString()?.Trim() ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(cellStr)) continue;

                    // Nhận diện Table Name: Table:, テーブル名:, Bảng:
                    var matchTable = Regex.Match(cellStr, @"(?:table|tbl|bảng|テーブル名?|table\s*name)[:：\s]+([A-Za-z0-9_]+)", RegexOptions.IgnoreCase);
                    if (matchTable.Success && matchTable.Groups.Count > 1)
                    {
                        result.DetectedTableName = matchTable.Groups[1].Value.ToUpperInvariant();
                    }

                    // Nhận diện Logical Table Name: 論理名:, Tên logic:
                    var matchLogical = Regex.Match(cellStr, @"(?:論理名|logical|tên\s*bảng|mô\s*tả)[:：\s]+([^\r\n]+)", RegexOptions.IgnoreCase);
                    if (matchLogical.Success && matchLogical.Groups.Count > 1)
                    {
                        result.DetectedTableLogicalName = matchLogical.Groups[1].Value.Trim();
                    }
                }
            }

            // 2. Nhận diện dòng Header (Header Row Detection)
            int bestHeaderRow = -1;
            int maxScore = 0;
            int colPhysical = -1;
            int colLogical = -1;
            int colType = -1;
            int colLength = -1;
            int colNullable = -1;
            int colPk = -1;
            int colDefault = -1;
            int colComment = -1;

            int scanLimit = Math.Min(10, rowCount);
            for (int r = 1; r <= scanLimit; r++)
            {
                int score = 0;
                int curPhysical = -1, curLogical = -1, curType = -1, curLength = -1;
                int curNullable = -1, curPk = -1, curDefault = -1, curComment = -1;

                for (int c = 1; c <= colCount; c++)
                {
                    string h = data[r, c]?.ToString()?.Trim().ToLowerInvariant() ?? string.Empty;
                    if (string.IsNullOrEmpty(h)) continue;

                    // Physical Column Name
                    if (MatchesHeader(h, "physical", "物理名", "tên cột", "cột", "column name", "col_name", "field name", "column", "field", "item_id", "項目名")
                        && !MatchesHeader(h, "logical", "論理"))
                    {
                        curPhysical = c;
                        score += 3;
                    }
                    // Logical Name
                    else if (MatchesHeader(h, "logical", "論理名", "tên logic", "tên trường", "tiêu đề", "和名", "tên tiếng việt", "tên tiếng nhật"))
                    {
                        curLogical = c;
                        score += 3;
                    }
                    // Data Type
                    else if (MatchesHeader(h, "data type", "datatype", "type", "データ型", "kiểu dữ liệu", "kiểu", "型"))
                    {
                        curType = c;
                        score += 3;
                    }
                    // Length / Scale
                    else if (MatchesHeader(h, "length", "size", "precision", "scale", "桁数", "桁", "độ dài", "kích thước", "長さ"))
                    {
                        curLength = c;
                        score += 2;
                    }
                    // Nullable
                    else if (MatchesHeader(h, "nullable", "allow null", "null", "not null", "必須", "cho phép null", "bắt buộc", "null?"))
                    {
                        curNullable = c;
                        score += 2;
                    }
                    // Primary Key
                    else if (MatchesHeader(h, "pk", "primary key", "primary", "主キー", "khóa chính", "key"))
                    {
                        curPk = c;
                        score += 2;
                    }
                    // Default
                    else if (MatchesHeader(h, "default", "initial", "初期値", "mặc định", "default value"))
                    {
                        curDefault = c;
                        score += 1;
                    }
                    // Comments / Remarks
                    else if (MatchesHeader(h, "comment", "remarks", "description", "note", "備考", "説明", "ghi chú", "mô tả"))
                    {
                        curComment = c;
                        score += 2;
                    }
                }

                if (score > maxScore)
                {
                    maxScore = score;
                    bestHeaderRow = r;
                    colPhysical = curPhysical;
                    colLogical = curLogical;
                    colType = curType;
                    colLength = curLength;
                    colNullable = curNullable;
                    colPk = curPk;
                    colDefault = curDefault;
                    colComment = curComment;
                }
            }

            // Fallback nếu không phát hiện được rõ ràng dòng tiêu đề
            if (bestHeaderRow == -1 || maxScore < 2)
            {
                bestHeaderRow = 1;
                // Heuristic mapping theo thứ tự cột thông dụng
                if (colPhysical == -1 && colCount >= 1) colPhysical = 1;
                if (colLogical == -1 && colCount >= 2) colLogical = 2;
                if (colType == -1 && colCount >= 3) colType = 3;
                if (colLength == -1 && colCount >= 4) colLength = 4;
            }

            result.HeaderRowIndex = bestHeaderRow;

            // 3. Quét các dòng dữ liệu đặc tả cột
            var list = new List<OracleColumnSpec>();
            for (int r = bestHeaderRow + 1; r <= rowCount; r++)
            {
                string rawPhysical = colPhysical > 0 ? (data[r, colPhysical]?.ToString()?.Trim() ?? string.Empty) : string.Empty;
                string rawLogical = colLogical > 0 ? (data[r, colLogical]?.ToString()?.Trim() ?? string.Empty) : string.Empty;
                string rawType = colType > 0 ? (data[r, colType]?.ToString()?.Trim() ?? string.Empty) : string.Empty;
                string rawLength = colLength > 0 ? (data[r, colLength]?.ToString()?.Trim() ?? string.Empty) : string.Empty;
                string rawNullable = colNullable > 0 ? (data[r, colNullable]?.ToString()?.Trim() ?? string.Empty) : string.Empty;
                string rawPk = colPk > 0 ? (data[r, colPk]?.ToString()?.Trim() ?? string.Empty) : string.Empty;
                string rawDefault = colDefault > 0 ? (data[r, colDefault]?.ToString()?.Trim() ?? string.Empty) : string.Empty;
                string rawComment = colComment > 0 ? (data[r, colComment]?.ToString()?.Trim() ?? string.Empty) : string.Empty;

                // Nếu dòng trống hoàn toàn hoặc không có tên cột physical lẫn logical -> bỏ qua
                if (string.IsNullOrWhiteSpace(rawPhysical) && string.IsNullOrWhiteSpace(rawLogical) && string.IsNullOrWhiteSpace(rawType))
                {
                    continue;
                }

                // Nếu physical name trống nhưng có logical name, thử tạo physical name
                if (string.IsNullOrWhiteSpace(rawPhysical))
                {
                    rawPhysical = "COL_" + r;
                }

                string physicalName = SanitizeIdentifier(rawPhysical);
                if (string.IsNullOrWhiteSpace(physicalName)) continue;

                // Chuẩn hóa Data Type & Length
                var (dataType, dataLength) = NormalizeOracleType(rawType, rawLength);

                // Chuẩn hóa Nullable
                bool isNullable = ParseNullable(rawNullable);

                // Chuẩn hóa PK
                bool isPk = ParsePrimaryKey(rawPk, rawPhysical, rawLogical);
                if (isPk)
                {
                    isNullable = false; // Primary Key bắt buộc NOT NULL trong Oracle
                }

                var colSpec = new OracleColumnSpec
                {
                    RowIndex = r,
                    PhysicalName = physicalName,
                    LogicalName = rawLogical,
                    OracleDataType = dataType,
                    DataLength = dataLength,
                    IsNullable = isNullable,
                    IsPrimaryKey = isPk,
                    DefaultValue = CleanDefaultValue(rawDefault),
                    Comments = !string.IsNullOrWhiteSpace(rawComment) ? rawComment : rawLogical,
                    IsSelected = true
                };

                list.Add(colSpec);
            }

            result.Columns = list;
            return result;
        }

        private static bool MatchesHeader(string text, params string[] keywords)
        {
            return keywords.Any(k => text.Equals(k, StringComparison.OrdinalIgnoreCase) || text.Contains(k));
        }

        public static string SanitizeIdentifier(string? name)
        {
            if (string.IsNullOrWhiteSpace(name)) return string.Empty;
            string s = name.Trim();
            s = Regex.Replace(s, @"\s+", "_");
            s = Regex.Replace(s, @"[^A-Za-z0-9_#$]", "");
            return s.ToUpperInvariant();
        }

        public static (string Type, string Length) NormalizeOracleType(string rawType, string rawLength)
        {
            string t = rawType.Trim().ToUpperInvariant();
            string l = rawLength.Trim();

            // Nếu rawType có chứa length dạng VARCHAR2(50) hay NUMBER(10,2)
            var matchParen = Regex.Match(t, @"^([A-Z0-9_]+)\s*\(([^)]+)\)$");
            if (matchParen.Success)
            {
                t = matchParen.Groups[1].Value.Trim();
                if (string.IsNullOrWhiteSpace(l))
                {
                    l = matchParen.Groups[2].Value.Trim();
                }
            }

            // Ánh xạ sang chuẩn Oracle
            switch (t)
            {
                case "VARCHAR":
                case "VARCHAR2":
                case "STRING":
                    t = "VARCHAR2";
                    if (string.IsNullOrWhiteSpace(l)) l = "100";
                    break;

                case "NVARCHAR":
                case "NVARCHAR2":
                    t = "NVARCHAR2";
                    if (string.IsNullOrWhiteSpace(l)) l = "100";
                    break;

                case "CHAR":
                    t = "CHAR";
                    if (string.IsNullOrWhiteSpace(l)) l = "1";
                    break;

                case "NCHAR":
                    t = "NCHAR";
                    if (string.IsNullOrWhiteSpace(l)) l = "1";
                    break;

                case "INT":
                case "INTEGER":
                    t = "NUMBER";
                    if (string.IsNullOrWhiteSpace(l)) l = "10";
                    break;

                case "BIGINT":
                    t = "NUMBER";
                    if (string.IsNullOrWhiteSpace(l)) l = "19";
                    break;

                case "SMALLINT":
                case "TINYINT":
                    t = "NUMBER";
                    if (string.IsNullOrWhiteSpace(l)) l = "5";
                    break;

                case "NUMBER":
                case "NUMERIC":
                case "DECIMAL":
                    t = "NUMBER";
                    break;

                case "FLOAT":
                case "DOUBLE":
                case "REAL":
                    t = "NUMBER";
                    break;

                case "DATE":
                    t = "DATE";
                    l = string.Empty;
                    break;

                case "DATETIME":
                case "TIMESTAMP":
                    t = "TIMESTAMP";
                    l = string.Empty;
                    break;

                case "BOOLEAN":
                case "BOOL":
                    t = "CHAR";
                    if (string.IsNullOrWhiteSpace(l)) l = "1";
                    break;

                case "CLOB":
                case "TEXT":
                case "LONGTEXT":
                case "NTEXT":
                    t = "CLOB";
                    l = string.Empty;
                    break;

                case "BLOB":
                case "BYTEA":
                case "IMAGE":
                case "BINARY":
                case "VARBINARY":
                    t = "BLOB";
                    l = string.Empty;
                    break;

                default:
                    if (string.IsNullOrWhiteSpace(t))
                    {
                        t = "VARCHAR2";
                        if (string.IsNullOrWhiteSpace(l)) l = "100";
                    }
                    break;
            }

            return (t, l);
        }

        private static bool ParseNullable(string val)
        {
            if (string.IsNullOrWhiteSpace(val)) return true;
            string v = val.Trim().ToUpperInvariant();

            if (v == "NOT NULL" || v == "N" || v == "NO" || v == "×" || v == "0" || v == "FALSE"
                || v == "必須" || v == "BẮT BUỘC" || v == "REQUIRED" || v == "MANDATORY")
            {
                return false;
            }

            return true;
        }

        private static bool ParsePrimaryKey(string pkVal, string physical, string logical)
        {
            string v = pkVal.Trim().ToUpperInvariant();
            if (v == "PK" || v == "Y" || v == "YES" || v == "○" || v == "1" || v == "TRUE"
                || v == "KHÓA CHÍNH" || v == "主キー" || v == "PRIMARY" || v == "PRIMARY KEY")
            {
                return true;
            }

            // Gợi ý nếu tên cột là ID, USER_ID, CODE hoặc có ghi chú PK
            if (physical.EndsWith("_ID", StringComparison.OrdinalIgnoreCase) && physical.Length <= 10 && v.Length > 0)
            {
                return true;
            }

            return false;
        }

        private static string CleanDefaultValue(string val)
        {
            if (string.IsNullOrWhiteSpace(val)) return string.Empty;
            string v = val.Trim();
            if (v.Equals("NULL", StringComparison.OrdinalIgnoreCase)) return string.Empty;
            return v;
        }

        // =========================================================================================
        // CODE & PL/SQL GENERATION LOGIC
        // =========================================================================================

        public static string GenerateScript(OracleSpecOptions options)
        {
            if (options == null) return string.Empty;

            var cols = options.Columns.Where(c => c.IsSelected).ToList();
            if (cols.Count == 0)
            {
                return "-- Không có cột nào được chọn để sinh mã.";
            }

            string fullTableName = string.IsNullOrWhiteSpace(options.SchemaName)
                ? options.TableName.ToUpperInvariant()
                : $"{options.SchemaName.ToUpperInvariant()}.{options.TableName.ToUpperInvariant()}";

            string pureTableName = options.TableName.ToUpperInvariant();

            switch (options.GenMode)
            {
                case OraclePlSqlGenMode.DdlScript:
                    return GenerateDdlScript(options, fullTableName, pureTableName, cols);

                case OraclePlSqlGenMode.PlSqlSafeBlock:
                    return GeneratePlSqlSafeBlock(options, pureTableName, cols);

                case OraclePlSqlGenMode.PlSqlCrudProcedures:
                    return GeneratePlSqlCrudProcedures(options, pureTableName, cols);

                case OraclePlSqlGenMode.PlSqlMergeUpsert:
                    return GeneratePlSqlMergeUpsert(options, pureTableName, cols);

                case OraclePlSqlGenMode.CSharpEntity:
                    return GenerateCSharpEntity(options, pureTableName, cols);

                case OraclePlSqlGenMode.TypeScriptDto:
                    return GenerateTypeScriptDto(options, pureTableName, cols);

                default:
                    return GenerateDdlScript(options, fullTableName, pureTableName, cols);
            }
        }

        private static string GenerateDdlScript(OracleSpecOptions opt, string fullTableName, string pureTableName, List<OracleColumnSpec> cols)
        {
            var sb = new StringBuilder();
            sb.AppendLine("-- =============================================================================");
            sb.AppendLine($"-- TABLE DDL SCRIPT: {fullTableName}");
            if (!string.IsNullOrWhiteSpace(opt.TableLogicalName))
            {
                sb.AppendLine($"-- Description      : {opt.TableLogicalName}");
            }
            sb.AppendLine($"-- Generated at     : {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            sb.AppendLine("-- Tool             : ExcelSupport Add-In (Spec to PL/SQL Generator)");
            sb.AppendLine("-- =============================================================================");
            sb.AppendLine();

            sb.AppendLine($"CREATE TABLE {fullTableName} (");

            var pkCols = cols.Where(c => c.IsPrimaryKey).ToList();

            for (int i = 0; i < cols.Count; i++)
            {
                var col = cols[i];
                string typeDef = col.GetFullDataTypeDefinition();

                // Identity cho Oracle 12c+
                string identityClause = string.Empty;
                if (opt.UseIdentityColumn && col.IsPrimaryKey && (typeDef.StartsWith("NUMBER", StringComparison.OrdinalIgnoreCase) || typeDef == "INT"))
                {
                    identityClause = " GENERATED ALWAYS AS IDENTITY";
                }

                string defaultClause = string.Empty;
                if (!string.IsNullOrWhiteSpace(col.DefaultValue) && string.IsNullOrEmpty(identityClause))
                {
                    defaultClause = $" DEFAULT {FormatDefaultValue(col.DefaultValue, col.OracleDataType)}";
                }

                string nullClause = col.IsNullable ? " NULL" : " NOT NULL";
                if (!string.IsNullOrEmpty(identityClause)) nullClause = string.Empty; // Identity tự mang NOT NULL

                string line = $"    {col.PhysicalName,-30} {typeDef,-18}{defaultClause}{identityClause}{nullClause}";

                // Dấu phẩy nếu còn cột tiếp theo hoặc có Primary Key Constraint
                if (i < cols.Count - 1 || pkCols.Count > 0)
                {
                    line += ",";
                }

                // Ghi chú inline tiếng Nhật / Việt nếu có
                if (!string.IsNullOrWhiteSpace(col.LogicalName))
                {
                    line += $" -- {col.LogicalName}";
                }

                sb.AppendLine(line);
            }

            // Khóa chính Constraint
            if (pkCols.Count > 0)
            {
                string pkColList = string.Join(", ", pkCols.Select(c => c.PhysicalName));
                sb.AppendLine($"    CONSTRAINT PK_{pureTableName} PRIMARY KEY ({pkColList})");
            }

            sb.AppendLine(");");
            sb.AppendLine();

            // Table Comment
            if (opt.IncludeComments)
            {
                string tblComment = !string.IsNullOrWhiteSpace(opt.TableLogicalName) ? opt.TableLogicalName : pureTableName;
                sb.AppendLine($"COMMENT ON TABLE {fullTableName} IS '{EscapeSqlString(tblComment)}';");
                sb.AppendLine();

                // Column Comments
                sb.AppendLine("-- Column Comments");
                foreach (var col in cols)
                {
                    string commentText = !string.IsNullOrWhiteSpace(col.Comments)
                        ? col.Comments
                        : (!string.IsNullOrWhiteSpace(col.LogicalName) ? col.LogicalName : col.PhysicalName);

                    sb.AppendLine($"COMMENT ON COLUMN {fullTableName}.{col.PhysicalName} IS '{EscapeSqlString(commentText)}';");
                }
                sb.AppendLine();
            }

            // Sequence & Trigger nếu người dùng chọn và không dùng Identity
            if (opt.IncludeSequenceTrigger && !opt.UseIdentityColumn && pkCols.Count > 0)
            {
                var firstPk = pkCols[0];
                string seqName = $"SEQ_{pureTableName}";
                string trgName = $"TRG_{pureTableName}_BI";

                sb.AppendLine("-- Sequence & Auto-increment Trigger");
                sb.AppendLine($"CREATE SEQUENCE {seqName} START WITH 1 INCREMENT BY 1 NOCACHE NOCYCLE;");
                sb.AppendLine();
                sb.AppendLine($"CREATE OR REPLACE TRIGGER {trgName}");
                sb.AppendLine($"BEFORE INSERT ON {fullTableName}");
                sb.AppendLine("FOR EACH ROW");
                sb.AppendLine("BEGIN");
                sb.AppendLine($"    IF :NEW.{firstPk.PhysicalName} IS NULL THEN");
                sb.AppendLine($"        SELECT {seqName}.NEXTVAL INTO :NEW.{firstPk.PhysicalName} FROM DUAL;");
                sb.AppendLine("    END IF;");
                sb.AppendLine("END;");
                sb.AppendLine("/");
                sb.AppendLine();
            }

            return sb.ToString();
        }

        private static string GeneratePlSqlSafeBlock(OracleSpecOptions opt, string pureTableName, List<OracleColumnSpec> cols)
        {
            var sb = new StringBuilder();
            sb.AppendLine("-- =============================================================================");
            sb.AppendLine($"-- PL/SQL IDEMPOTENT CREATION BLOCK: {pureTableName}");
            sb.AppendLine("-- Kiểm tra và tạo bảng an toàn nếu chưa tồn tại trong Schema");
            sb.AppendLine("-- =============================================================================");
            sb.AppendLine("DECLARE");
            sb.AppendLine("    v_table_exists NUMBER := 0;");
            sb.AppendLine("BEGIN");
            sb.AppendLine($"    SELECT COUNT(*) INTO v_table_exists FROM user_tables WHERE table_name = '{pureTableName}';");
            sb.AppendLine();
            sb.AppendLine("    IF v_table_exists = 0 THEN");
            sb.AppendLine($"        DBMS_OUTPUT.PUT_LINE('Starting creation of table {pureTableName}...');");
            sb.AppendLine();
            sb.AppendLine("        EXECUTE IMMEDIATE '");

            var pkCols = cols.Where(c => c.IsPrimaryKey).ToList();
            sb.AppendLine($"            CREATE TABLE {pureTableName} (");

            for (int i = 0; i < cols.Count; i++)
            {
                var col = cols[i];
                string typeDef = col.GetFullDataTypeDefinition();

                string defaultClause = string.Empty;
                if (!string.IsNullOrWhiteSpace(col.DefaultValue))
                {
                    defaultClause = $" DEFAULT {FormatDefaultValue(col.DefaultValue, col.OracleDataType)}";
                }

                string nullClause = col.IsNullable ? " NULL" : " NOT NULL";
                string comma = (i < cols.Count - 1 || pkCols.Count > 0) ? "," : "";

                sb.AppendLine($"                {col.PhysicalName,-28} {typeDef,-16}{defaultClause}{nullClause}{comma}");
            }

            if (pkCols.Count > 0)
            {
                string pkColList = string.Join(", ", pkCols.Select(c => c.PhysicalName));
                sb.AppendLine($"                CONSTRAINT PK_{pureTableName} PRIMARY KEY ({pkColList})");
            }

            sb.AppendLine("            )';");
            sb.AppendLine();

            if (opt.IncludeComments)
            {
                string tblComment = !string.IsNullOrWhiteSpace(opt.TableLogicalName) ? opt.TableLogicalName : pureTableName;
                sb.AppendLine($"        EXECUTE IMMEDIATE 'COMMENT ON TABLE {pureTableName} IS ''{EscapeSqlString(tblComment)}''';");

                foreach (var col in cols)
                {
                    string commentText = !string.IsNullOrWhiteSpace(col.Comments)
                        ? col.Comments
                        : (!string.IsNullOrWhiteSpace(col.LogicalName) ? col.LogicalName : col.PhysicalName);

                    sb.AppendLine($"        EXECUTE IMMEDIATE 'COMMENT ON COLUMN {pureTableName}.{col.PhysicalName} IS ''{EscapeSqlString(commentText)}''';");
                }
            }

            sb.AppendLine();
            sb.AppendLine($"        DBMS_OUTPUT.PUT_LINE('SUCCESS: Table {pureTableName} created successfully.');");
            sb.AppendLine("    ELSE");
            sb.AppendLine($"        DBMS_OUTPUT.PUT_LINE('INFO: Table {pureTableName} already exists. Skipping creation.');");
            sb.AppendLine("    END IF;");
            sb.AppendLine("EXCEPTION");
            sb.AppendLine("    WHEN OTHERS THEN");
            sb.AppendLine($"        DBMS_OUTPUT.PUT_LINE('ERROR creating table {pureTableName}: ' || SQLERRM);");
            sb.AppendLine("        RAISE;");
            sb.AppendLine("END;");
            sb.AppendLine("/");

            return sb.ToString();
        }

        private static string GeneratePlSqlCrudProcedures(OracleSpecOptions opt, string pureTableName, List<OracleColumnSpec> cols)
        {
            var sb = new StringBuilder();
            sb.AppendLine("-- =============================================================================");
            sb.AppendLine($"-- PL/SQL CRUD PROCEDURES SKELETON: {pureTableName}");
            sb.AppendLine("-- Generated based on Excel Table Specification");
            sb.AppendLine("-- =============================================================================");
            sb.AppendLine();

            var pkCols = cols.Where(c => c.IsPrimaryKey).ToList();
            if (pkCols.Count == 0)
            {
                pkCols = cols.Take(1).ToList();
            }

            var nonPkCols = cols.Where(c => !c.IsPrimaryKey).ToList();

            // 1. Procedure INSERT
            sb.AppendLine($"-- 1. INSERT PROCEDURE: SP_INS_{pureTableName}");
            sb.AppendLine($"CREATE OR REPLACE PROCEDURE SP_INS_{pureTableName} (");
            for (int i = 0; i < cols.Count; i++)
            {
                var col = cols[i];
                string paramType = col.GetFullDataTypeDefinition();
                string comma = (i < cols.Count - 1) ? "," : "";
                sb.AppendLine($"    p_{col.PhysicalName,-28} IN  {pureTableName}.{col.PhysicalName}%TYPE{comma}");
            }
            sb.AppendLine(")");
            sb.AppendLine("IS");
            sb.AppendLine("BEGIN");
            sb.AppendLine($"    INSERT INTO {pureTableName} (");
            sb.AppendLine("        " + string.Join(",\n        ", cols.Select(c => c.PhysicalName)));
            sb.AppendLine("    ) VALUES (");
            sb.AppendLine("        " + string.Join(",\n        ", cols.Select(c => $"p_{c.PhysicalName}")));
            sb.AppendLine("    );");
            sb.AppendLine("EXCEPTION");
            sb.AppendLine("    WHEN OTHERS THEN");
            sb.AppendLine($"        RAISE_APPLICATION_ERROR(-20001, 'Error in SP_INS_{pureTableName}: ' || SQLERRM);");
            sb.AppendLine("END;");
            sb.AppendLine("/");
            sb.AppendLine();

            // 2. Procedure UPDATE
            sb.AppendLine($"-- 2. UPDATE PROCEDURE: SP_UPD_{pureTableName}");
            sb.AppendLine($"CREATE OR REPLACE PROCEDURE SP_UPD_{pureTableName} (");
            for (int i = 0; i < cols.Count; i++)
            {
                var col = cols[i];
                string comma = (i < cols.Count - 1) ? "," : "";
                sb.AppendLine($"    p_{col.PhysicalName,-28} IN  {pureTableName}.{col.PhysicalName}%TYPE{comma}");
            }
            sb.AppendLine(")");
            sb.AppendLine("IS");
            sb.AppendLine("BEGIN");
            sb.AppendLine($"    UPDATE {pureTableName} SET");
            sb.AppendLine("        " + string.Join(",\n        ", nonPkCols.Select(c => $"{c.PhysicalName} = p_{c.PhysicalName}")));
            sb.AppendLine("    WHERE");
            sb.AppendLine("        " + string.Join(" AND\n        ", pkCols.Select(c => $"{c.PhysicalName} = p_{c.PhysicalName}")) + ";");
            sb.AppendLine("EXCEPTION");
            sb.AppendLine("    WHEN OTHERS THEN");
            sb.AppendLine($"        RAISE_APPLICATION_ERROR(-20002, 'Error in SP_UPD_{pureTableName}: ' || SQLERRM);");
            sb.AppendLine("END;");
            sb.AppendLine("/");
            sb.AppendLine();

            // 3. Procedure DELETE
            sb.AppendLine($"-- 3. DELETE PROCEDURE: SP_DEL_{pureTableName}");
            sb.AppendLine($"CREATE OR REPLACE PROCEDURE SP_DEL_{pureTableName} (");
            for (int i = 0; i < pkCols.Count; i++)
            {
                var col = pkCols[i];
                string comma = (i < pkCols.Count - 1) ? "," : "";
                sb.AppendLine($"    p_{col.PhysicalName,-28} IN  {pureTableName}.{col.PhysicalName}%TYPE{comma}");
            }
            sb.AppendLine(")");
            sb.AppendLine("IS");
            sb.AppendLine("BEGIN");
            sb.AppendLine($"    DELETE FROM {pureTableName}");
            sb.AppendLine("    WHERE");
            sb.AppendLine("        " + string.Join(" AND\n        ", pkCols.Select(c => $"{c.PhysicalName} = p_{c.PhysicalName}")) + ";");
            sb.AppendLine("EXCEPTION");
            sb.AppendLine("    WHEN OTHERS THEN");
            sb.AppendLine($"        RAISE_APPLICATION_ERROR(-20003, 'Error in SP_DEL_{pureTableName}: ' || SQLERRM);");
            sb.AppendLine("END;");
            sb.AppendLine("/");
            sb.AppendLine();

            return sb.ToString();
        }

        private static string GeneratePlSqlMergeUpsert(OracleSpecOptions opt, string pureTableName, List<OracleColumnSpec> cols)
        {
            var sb = new StringBuilder();
            sb.AppendLine("-- =============================================================================");
            sb.AppendLine($"-- PL/SQL MERGE INTO (UPSERT) TEMPLATE: {pureTableName}");
            sb.AppendLine("-- Cập nhật nếu đã tồn tại, chèn mới nếu chưa có theo Khóa Chính");
            sb.AppendLine("-- =============================================================================");
            sb.AppendLine();

            var pkCols = cols.Where(c => c.IsPrimaryKey).ToList();
            if (pkCols.Count == 0)
            {
                pkCols = cols.Take(1).ToList();
            }

            var nonPkCols = cols.Where(c => !c.IsPrimaryKey).ToList();

            sb.AppendLine($"MERGE INTO {pureTableName} t");
            sb.AppendLine("USING (");
            sb.AppendLine("    SELECT");
            for (int i = 0; i < cols.Count; i++)
            {
                var col = cols[i];
                string comma = (i < cols.Count - 1) ? "," : "";
                sb.AppendLine($"        :p_{col.PhysicalName,-28} AS {col.PhysicalName}{comma}");
            }
            sb.AppendLine("    FROM DUAL");
            sb.AppendLine(") s");
            sb.AppendLine("ON (");
            sb.AppendLine("    " + string.Join(" AND\n    ", pkCols.Select(c => $"t.{c.PhysicalName} = s.{c.PhysicalName}")));
            sb.AppendLine(")");

            if (nonPkCols.Count > 0)
            {
                sb.AppendLine("WHEN MATCHED THEN");
                sb.AppendLine("    UPDATE SET");
                sb.AppendLine("        " + string.Join(",\n        ", nonPkCols.Select(c => $"t.{c.PhysicalName} = s.{c.PhysicalName}")));
            }

            sb.AppendLine("WHEN NOT MATCHED THEN");
            sb.AppendLine("    INSERT (");
            sb.AppendLine("        " + string.Join(",\n        ", cols.Select(c => $"t.{c.PhysicalName}")));
            sb.AppendLine("    ) VALUES (");
            sb.AppendLine("        " + string.Join(",\n        ", cols.Select(c => $"s.{c.PhysicalName}")));
            sb.AppendLine("    );");
            sb.AppendLine("/");

            return sb.ToString();
        }

        private static string GenerateCSharpEntity(OracleSpecOptions opt, string pureTableName, List<OracleColumnSpec> cols)
        {
            var sb = new StringBuilder();
            string className = ToPascalCase(pureTableName);

            sb.AppendLine("using System;");
            sb.AppendLine("using System.ComponentModel.DataAnnotations;");
            sb.AppendLine("using System.ComponentModel.DataAnnotations.Schema;");
            sb.AppendLine();
            sb.AppendLine("namespace YourProject.Entities");
            sb.AppendLine("{");

            if (!string.IsNullOrWhiteSpace(opt.TableLogicalName))
            {
                sb.AppendLine("    /// <summary>");
                sb.AppendLine($"    /// {EscapeXml(opt.TableLogicalName)}");
                sb.AppendLine("    /// </summary>");
            }
            sb.AppendLine($"    [Table(\"{pureTableName}\")]");
            sb.AppendLine($"    public class {className}");
            sb.AppendLine("    {");

            foreach (var col in cols)
            {
                string propName = ToPascalCase(col.PhysicalName);
                string csType = MapOracleToCSharpType(col.OracleDataType, col.IsNullable);

                string comment = !string.IsNullOrWhiteSpace(col.LogicalName) ? col.LogicalName : col.Comments;
                if (!string.IsNullOrWhiteSpace(comment))
                {
                    sb.AppendLine("        /// <summary>");
                    sb.AppendLine($"        /// {EscapeXml(comment)}");
                    sb.AppendLine("        /// </summary>");
                }

                if (col.IsPrimaryKey)
                {
                    sb.AppendLine("        [Key]");
                }

                sb.AppendLine($"        [Column(\"{col.PhysicalName}\")]");
                sb.AppendLine($"        public {csType} {propName} {{ get; set; }}");
                sb.AppendLine();
            }

            sb.AppendLine("    }");
            sb.AppendLine("}");

            return sb.ToString();
        }

        private static string GenerateTypeScriptDto(OracleSpecOptions opt, string pureTableName, List<OracleColumnSpec> cols)
        {
            var sb = new StringBuilder();
            string interfaceName = "I" + ToPascalCase(pureTableName);

            if (!string.IsNullOrWhiteSpace(opt.TableLogicalName))
            {
                sb.AppendLine("/**");
                sb.AppendLine($" * {opt.TableLogicalName}");
                sb.AppendLine(" */");
            }
            sb.AppendLine($"export interface {interfaceName} {{");

            foreach (var col in cols)
            {
                string propName = ToCamelCase(col.PhysicalName);
                string tsType = MapOracleToTypeScriptType(col.OracleDataType);
                string optMark = col.IsNullable ? "?" : "";

                string comment = !string.IsNullOrWhiteSpace(col.LogicalName) ? col.LogicalName : col.Comments;
                if (!string.IsNullOrWhiteSpace(comment))
                {
                    sb.AppendLine($"  /** {comment} */");
                }

                sb.AppendLine($"  {propName}{optMark}: {tsType};");
            }

            sb.AppendLine("}");
            return sb.ToString();
        }

        private static string FormatDefaultValue(string defVal, string dataType)
        {
            if (string.IsNullOrWhiteSpace(defVal)) return string.Empty;
            string v = defVal.Trim();

            // Nếu là hàm hệ thống
            if (v.Equals("SYSDATE", StringComparison.OrdinalIgnoreCase) ||
                v.Equals("CURRENT_TIMESTAMP", StringComparison.OrdinalIgnoreCase) ||
                v.Equals("SYSTIMESTAMP", StringComparison.OrdinalIgnoreCase))
            {
                return v.ToUpperInvariant();
            }

            // Nếu đã có dấu nháy đơn
            if (v.StartsWith("'") && v.EndsWith("'"))
            {
                return v;
            }

            // Nếu là kiểu số
            if (dataType == "NUMBER" || dataType == "INT" || dataType == "INTEGER" || dataType == "FLOAT")
            {
                return v;
            }

            return $"'{EscapeSqlString(v)}'";
        }

        private static string EscapeSqlString(string str)
        {
            return str.Replace("'", "''");
        }

        private static string EscapeXml(string str)
        {
            return str.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");
        }

        public static string ToPascalCase(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return string.Empty;
            var parts = text.Split(new[] { '_', ' ', '-' }, StringSplitOptions.RemoveEmptyEntries);
            var sb = new StringBuilder();
            foreach (var p in parts)
            {
                if (p.Length == 1)
                {
                    sb.Append(char.ToUpperInvariant(p[0]));
                }
                else if (p.Length > 1)
                {
                    sb.Append(char.ToUpperInvariant(p[0]));
                    sb.Append(p.Substring(1).ToLowerInvariant());
                }
            }
            return sb.ToString();
        }

        public static string ToCamelCase(string text)
        {
            string p = ToPascalCase(text);
            if (p.Length == 0) return p;
            return char.ToLowerInvariant(p[0]) + p.Substring(1);
        }

        private static string MapOracleToCSharpType(string oracleType, bool isNullable)
        {
            string t = oracleType.ToUpperInvariant();
            string nullMark = isNullable ? "?" : "";

            if (t.StartsWith("VARCHAR") || t.StartsWith("CHAR") || t == "CLOB" || t.StartsWith("NVARCHAR"))
            {
                return "string";
            }
            if (t.StartsWith("NUMBER") || t == "DECIMAL" || t == "NUMERIC")
            {
                return $"decimal{nullMark}";
            }
            if (t == "INT" || t == "INTEGER")
            {
                return $"int{nullMark}";
            }
            if (t == "DATE" || t.StartsWith("TIMESTAMP"))
            {
                return $"DateTime{nullMark}";
            }
            if (t == "BLOB")
            {
                return "byte[]";
            }

            return "string";
        }

        private static string MapOracleToTypeScriptType(string oracleType)
        {
            string t = oracleType.ToUpperInvariant();
            if (t.StartsWith("VARCHAR") || t.StartsWith("CHAR") || t == "CLOB" || t.StartsWith("NVARCHAR"))
            {
                return "string";
            }
            if (t.StartsWith("NUMBER") || t == "DECIMAL" || t == "NUMERIC" || t == "INT" || t == "INTEGER")
            {
                return "number";
            }
            if (t == "DATE" || t.StartsWith("TIMESTAMP"))
            {
                return "Date | string";
            }
            if (t == "BLOB")
            {
                return "Uint8Array";
            }

            return "string";
        }
    }
}
