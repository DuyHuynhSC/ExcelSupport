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

        public static List<string> ExtractLinesFrom2D(object[,]? data)
        {
            var lines = new List<string>();
            if (data == null) return lines;

            int rowCount = data.GetLength(0);
            int colCount = data.GetLength(1);

            for (int r = 1; r <= rowCount; r++)
            {
                var rowTokens = new List<string>();
                for (int c = 1; c <= colCount; c++)
                {
                    string cellStr = data[r, c]?.ToString()?.Trim() ?? string.Empty;
                    if (!string.IsNullOrEmpty(cellStr))
                    {
                        rowTokens.Add(cellStr);
                    }
                }

                if (rowTokens.Count == 0) continue;

                // Kiểm tra xem có cell nào chứa ký tự xuống dòng \n (như Alt+Enter trong Excel)
                bool hasNewline = rowTokens.Any(t => t.Contains("\n"));
                if (hasNewline)
                {
                    foreach (var token in rowTokens)
                    {
                        var subLines = token.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                        foreach (var sl in subLines)
                        {
                            if (!string.IsNullOrWhiteSpace(sl))
                            {
                                lines.Add(sl.Trim());
                            }
                        }
                    }
                }
                else
                {
                    // Ghép các cột trong cùng 1 dòng thành 1 chuỗi hoàn chỉnh
                    string joined = string.Join(" ", rowTokens).Trim();
                    if (!string.IsNullOrWhiteSpace(joined))
                    {
                        lines.Add(joined);
                    }
                }
            }

            return lines;
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

            var lines = ExtractLinesFrom2D(data);
            if (lines.Count == 0)
            {
                result.Success = false;
                result.ErrorMessage = "Vùng chọn không có dòng dữ liệu hợp lệ.";
                return result;
            }

            // 1. Phân loại xem vùng chọn là Đặc tả Truy vấn / TKCT (QueryLogicSpec) hay Đặc tả DDL (TableDefinitionDdl)
            bool isQueryLogic = DetectIsQueryLogicSpec(lines);

            if (isQueryLogic)
            {
                result.SpecType = OracleSpecType.QueryLogicSpec;
                return ParseQueryLogicSpec(lines, fallbackSheetName);
            }
            else
            {
                result.SpecType = OracleSpecType.TableDefinitionDdl;
                return ParseTableDefinitionDdl(data, fallbackSheetName);
            }
        }

        private static bool DetectIsQueryLogicSpec(List<string> lines)
        {
            int queryScore = 0;
            foreach (var line in lines)
            {
                // Có chứa [COLUMN] theo sau bởi toán tử so sánh: = , <> , >= , <= , != , LIKE , IN
                if (Regex.IsMatch(line, @"\[[A-Za-z0-9_#$]+\]\s*(=|<>|!=|>=|<=|>|<|LIKE|IN)", RegexOptions.IgnoreCase))
                {
                    queryScore += 3;
                }
                // Có chứa từ khóa trích xuất / query
                if (Regex.IsMatch(line, @"(?:trích\s*xuất|lấy\s*data|kiểm\s*tra|tồn\s*tại|抽出|取得|data\s*table|画面\.)", RegexOptions.IgnoreCase))
                {
                    queryScore += 2;
                }
            }

            return queryScore >= 2;
        }

        // =========================================================================================
        // PARSER CHO QUERY / LOGIC SPEC (ĐẶC TẢ TRÍCH XUẤT TKCT)
        // =========================================================================================

        public static OracleSpecParseResult ParseQueryLogicSpec(List<string> lines, string fallbackSheetName)
        {
            var result = new OracleSpecParseResult
            {
                Success = true,
                SpecType = OracleSpecType.QueryLogicSpec
            };

            var queryInfo = new OracleQuerySpecInfo();
            string detectedTable = string.Empty;
            string detectedLogicalTable = string.Empty;

            // 1. Quét tìm Table Name và Action Type
            foreach (var line in lines)
            {
                // Match: Trích xuất data table 受注見出トラン[JDNTHA]
                var matchTable = Regex.Match(line, @"(?:data\s*table|table|bảng|テーブル)?\s*([^\r\n\[\]]+)?\s*\[([A-Za-z0-9_#$]+)\]", RegexOptions.IgnoreCase);
                if (matchTable.Success)
                {
                    string candidateLogical = matchTable.Groups[1].Value.Trim();
                    string candidatePhysical = matchTable.Groups[2].Value.Trim().ToUpperInvariant();

                    // Loại bỏ các từ tiền tố thừa như "Trích xuất data table ", "・ "
                    candidateLogical = Regex.Replace(candidateLogical, @"^[・\-\*\s]*(?:trích\s*xuất|cập\s*nhật|xóa|thêm|data\s*table|table|bảng|テーブル)?\s*", "", RegexOptions.IgnoreCase).Trim();

                    // Đảm bảo không phải là dòng điều kiện cột (dòng điều kiện thường có toán tử = hoặc <>)
                    if (!line.Contains("=") && !line.Contains("<>") && !line.Contains(">=") && !line.Contains("<="))
                    {
                        detectedTable = candidatePhysical;
                        detectedLogicalTable = candidateLogical;
                    }
                    else if (string.IsNullOrEmpty(detectedTable))
                    {
                        // Fallback lấy table name đầu tiên nếu chưa có
                        detectedTable = candidatePhysical;
                    }
                }

                // Nhận diện Action Type
                if (Regex.IsMatch(line, @"(?:trích\s*xuất|lấy|kiểm\s*tra|check|tìm\s*kiếm|抽出|取得|検索|select)", RegexOptions.IgnoreCase))
                {
                    queryInfo.ActionType = "SELECT";
                }
                else if (Regex.IsMatch(line, @"(?:cập\s*nhật|sửa|更新|update)", RegexOptions.IgnoreCase))
                {
                    queryInfo.ActionType = "UPDATE";
                }
                else if (Regex.IsMatch(line, @"(?:xóa|hủy|削除|delete)", RegexOptions.IgnoreCase))
                {
                    queryInfo.ActionType = "DELETE";
                }
                else if (Regex.IsMatch(line, @"(?:thêm|chèn|tạo|登録|insert)", RegexOptions.IgnoreCase))
                {
                    queryInfo.ActionType = "INSERT";
                }
            }

            if (string.IsNullOrWhiteSpace(detectedTable))
            {
                detectedTable = SanitizeIdentifier(fallbackSheetName);
                if (string.IsNullOrWhiteSpace(detectedTable)) detectedTable = "MY_TABLE";
            }

            queryInfo.TableName = detectedTable;
            queryInfo.TableLogicalName = detectedLogicalTable;
            result.DetectedTableName = detectedTable;
            result.DetectedTableLogicalName = detectedLogicalTable;

            // 2. Quét các dòng điều kiện (Conditions) và Hậu điều kiện (Post-actions)
            int condIndex = 1;
            foreach (var rawLine in lines)
            {
                string line = rawLine.Trim();
                if (string.IsNullOrEmpty(line)) continue;

                // A. Nhận diện Post-Condition: Trường hợp tồn tại data...
                if (Regex.IsMatch(line, @"(?:tồn\s*tại\s*data|có\s*data|データが存在する場合)", RegexOptions.IgnoreCase))
                {
                    queryInfo.OnExistsAction = line.TrimStart('・', '-', '*', ' ');
                    continue;
                }

                // B. Nhận diện Post-Condition: Trường hợp không tồn tại data...
                if (Regex.IsMatch(line, @"(?:không\s*tồn\s*tại\s*data|không\s*có\s*data|データが存在しない場合)", RegexOptions.IgnoreCase))
                {
                    queryInfo.OnNotExistsAction = line.TrimStart('・', '-', '*', ' ');
                    continue;
                }

                // C. Nhận diện Điều Kiện Cột: [COLUMN] toán_tử giá_trị
                // Ví dụ: 伝票削除区分 [DATKB] = "1"
                // Hoặc: 処理区分 [HANKB] = 画面.処理区分[HD_HANKB]
                var matchCond = Regex.Match(line,
                    @"^[・\-\*\s]*(?<logical>[^\[\r\n=<>]+)?\s*\[(?<physical>[A-Za-z0-9_#$]+)\]\s*(?<op>=|<>|!=|>=|<=|>|<|LIKE|NOT\s+LIKE|IN|NOT\s+IN|IS\s+NULL|IS\s+NOT\s+NULL)\s*(?<val>.+)$",
                    RegexOptions.IgnoreCase);

                if (matchCond.Success)
                {
                    string logicalName = matchCond.Groups["logical"].Value.Trim();
                    string physicalCol = matchCond.Groups["physical"].Value.Trim().ToUpperInvariant();
                    string op = matchCond.Groups["op"].Value.Trim().ToUpperInvariant();
                    string rawVal = matchCond.Groups["val"].Value.Trim();

                    // Chuẩn hóa SQL Value & Bind Variable
                    var (targetSql, bindVar) = ParseSqlValueAndBindVar(rawVal);

                    var cond = new OracleConditionSpec
                    {
                        RowIndex = condIndex++,
                        LogicalName = logicalName,
                        PhysicalColumn = physicalCol,
                        Operator = op,
                        RawValue = rawVal,
                        TargetSqlValue = targetSql,
                        BindVariableName = bindVar,
                        IsSelected = true
                    };

                    queryInfo.Conditions.Add(cond);
                }
                else
                {
                    // D. Ghi chú nghiệp vụ khác (VD: 部門制御 thì giống chuẩn)
                    if (!line.StartsWith("・") && !line.Contains(detectedTable))
                    {
                        queryInfo.OtherNotes.Add(line);
                    }
                }
            }

            result.QuerySpec = queryInfo;
            return result;
        }

        private static (string TargetSql, string BindVar) ParseSqlValueAndBindVar(string rawVal)
        {
            if (string.IsNullOrWhiteSpace(rawVal)) return ("NULL", string.Empty);
            string v = rawVal.Trim();

            // 1. Nếu là giá trị chuỗi cố định trong dấu ngoặc kép: "1", "0", "9"
            var matchQuote = Regex.Match(v, @"^[""'](.*)[""']$");
            if (matchQuote.Success)
            {
                string inner = matchQuote.Groups[1].Value;
                return ($"'{EscapeSqlString(inner)}'", string.Empty);
            }

            // 2. Nếu là item từ màn hình: 画面.処理区分[HD_HANKB] hoặc 画面.出荷予定日[HD_SYUYTIDT]
            var matchScreen = Regex.Match(v, @"(?:画面[.\s]*[^\r\n\[\]]*)?\[([A-Za-z0-9_#$]+)\]");
            if (matchScreen.Success)
            {
                string varName = matchScreen.Groups[1].Value.Trim().ToUpperInvariant();
                return ($":{varName}", varName);
            }

            // 3. Nếu là số nguyên hoặc số thập phân
            if (decimal.TryParse(v, out _))
            {
                return (v, string.Empty);
            }

            // 4. Nếu là hàm SQL hoặc từ khóa hệ thống: SYSDATE, NULL, CURRENT_TIMESTAMP
            if (v.Equals("SYSDATE", StringComparison.OrdinalIgnoreCase) ||
                v.Equals("NULL", StringComparison.OrdinalIgnoreCase) ||
                v.Equals("CURRENT_TIMESTAMP", StringComparison.OrdinalIgnoreCase))
            {
                return (v.ToUpperInvariant(), string.Empty);
            }

            // 5. Mặc định coi là biến bind hoặc chuỗi
            if (Regex.IsMatch(v, @"^[A-Za-z0-9_#$]+$"))
            {
                return ($":{v.ToUpperInvariant()}", v.ToUpperInvariant());
            }

            return ($"'{EscapeSqlString(v)}'", string.Empty);
        }

        // =========================================================================================
        // PARSER CHO TABLE DEFINITION DDL SPEC (ĐẶC TẢ CẤU TRÚC BẢNG)
        // =========================================================================================

        public static OracleSpecParseResult ParseTableDefinitionDdl(object[,] data, string fallbackSheetName)
        {
            var result = new OracleSpecParseResult
            {
                Success = true,
                SpecType = OracleSpecType.TableDefinitionDdl
            };

            int rowCount = data.GetLength(0);
            int colCount = data.GetLength(1);

            string defaultTableName = SanitizeIdentifier(fallbackSheetName);
            if (string.IsNullOrWhiteSpace(defaultTableName)) defaultTableName = "MY_TABLE";
            result.DetectedTableName = defaultTableName;

            // 1. Quét tìm Table Name hoặc Logical Table Name ở các dòng trên cùng
            for (int r = 1; r <= Math.Min(5, rowCount); r++)
            {
                for (int c = 1; c <= colCount; c++)
                {
                    string cellStr = data[r, c]?.ToString()?.Trim() ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(cellStr)) continue;

                    var matchTable = Regex.Match(cellStr, @"(?:table|tbl|bảng|テーブル名?|table\s*name)[:：\s]+([A-Za-z0-9_]+)", RegexOptions.IgnoreCase);
                    if (matchTable.Success && matchTable.Groups.Count > 1)
                    {
                        result.DetectedTableName = matchTable.Groups[1].Value.ToUpperInvariant();
                    }

                    var matchLogical = Regex.Match(cellStr, @"(?:論理名|logical|tên\s*bảng|mô\s*tả)[:：\s]+([^\r\n]+)", RegexOptions.IgnoreCase);
                    if (matchLogical.Success && matchLogical.Groups.Count > 1)
                    {
                        result.DetectedTableLogicalName = matchLogical.Groups[1].Value.Trim();
                    }
                }
            }

            // 2. Nhận diện dòng Header
            int bestHeaderRow = -1;
            int maxScore = 0;
            int colPhysical = -1, colLogical = -1, colType = -1, colLength = -1;
            int colNullable = -1, colPk = -1, colDefault = -1, colComment = -1;

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

                    if (MatchesHeader(h, "physical", "物理名", "tên cột", "cột", "column name", "col_name", "field name", "column", "field", "item_id", "項目名")
                        && !MatchesHeader(h, "logical", "論理"))
                    {
                        curPhysical = c;
                        score += 3;
                    }
                    else if (MatchesHeader(h, "logical", "論理名", "tên logic", "tên trường", "tiêu đề", "和名", "tên tiếng việt", "tên tiếng nhật"))
                    {
                        curLogical = c;
                        score += 3;
                    }
                    else if (MatchesHeader(h, "data type", "datatype", "type", "データ型", "kiểu dữ liệu", "kiểu", "型"))
                    {
                        curType = c;
                        score += 3;
                    }
                    else if (MatchesHeader(h, "length", "size", "precision", "scale", "桁数", "桁", "độ dài", "kích thước", "長さ"))
                    {
                        curLength = c;
                        score += 2;
                    }
                    else if (MatchesHeader(h, "nullable", "allow null", "null", "not null", "必須", "cho phép null", "bắt buộc", "null?"))
                    {
                        curNullable = c;
                        score += 2;
                    }
                    else if (MatchesHeader(h, "pk", "primary key", "primary", "主キー", "khóa chính", "key"))
                    {
                        curPk = c;
                        score += 2;
                    }
                    else if (MatchesHeader(h, "default", "initial", "初期値", "mặc định", "default value"))
                    {
                        curDefault = c;
                        score += 1;
                    }
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

            if (bestHeaderRow == -1 || maxScore < 2)
            {
                bestHeaderRow = 1;
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

                if (string.IsNullOrWhiteSpace(rawPhysical) && string.IsNullOrWhiteSpace(rawLogical) && string.IsNullOrWhiteSpace(rawType))
                {
                    continue;
                }

                if (string.IsNullOrWhiteSpace(rawPhysical))
                {
                    rawPhysical = "COL_" + r;
                }

                string physicalName = SanitizeIdentifier(rawPhysical);
                if (string.IsNullOrWhiteSpace(physicalName)) continue;

                var (dataType, dataLength) = NormalizeOracleType(rawType, rawLength);
                bool isNullable = ParseNullable(rawNullable);
                bool isPk = ParsePrimaryKey(rawPk, rawPhysical, rawLogical);
                if (isPk) isNullable = false;

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
            string s = name!.Trim();
            s = Regex.Replace(s, @"\s+", "_");
            s = Regex.Replace(s, @"[^A-Za-z0-9_#$]", "");
            return s.ToUpperInvariant();
        }

        public static (string Type, string Length) NormalizeOracleType(string rawType, string rawLength)
        {
            string t = rawType.Trim().ToUpperInvariant();
            string l = rawLength.Trim();

            var matchParen = Regex.Match(t, @"^([A-Z0-9_]+)\s*\(([^)]+)\)$");
            if (matchParen.Success)
            {
                t = matchParen.Groups[1].Value.Trim();
                if (string.IsNullOrWhiteSpace(l))
                {
                    l = matchParen.Groups[2].Value.Trim();
                }
            }

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
        // CODE & PL/SQL GENERATION DISPATCHER
        // =========================================================================================

        public static string GenerateScript(OracleSpecOptions options)
        {
            if (options == null) return string.Empty;

            if (options.SpecType == OracleSpecType.QueryLogicSpec)
            {
                return GenerateQueryLogicScript(options);
            }
            else
            {
                return GenerateDdlSpecScript(options);
            }
        }

        // =========================================================================================
        // GENERATION CHO QUERY / LOGIC SPEC (TKCT)
        // =========================================================================================

        private static string GenerateQueryLogicScript(OracleSpecOptions opt)
        {
            var q = opt.QuerySpec;
            var conds = q.Conditions.Where(c => c.IsSelected).ToList();

            string fullTableName = string.IsNullOrWhiteSpace(opt.SchemaName)
                ? opt.TableName.ToUpperInvariant()
                : $"{opt.SchemaName.ToUpperInvariant()}.{opt.TableName.ToUpperInvariant()}";

            string pureTableName = opt.TableName.ToUpperInvariant();

            switch (opt.GenMode)
            {
                case OraclePlSqlGenMode.QueryCheckBlock:
                    return GeneratePlSqlCheckBlock(opt, fullTableName, pureTableName, conds);

                case OraclePlSqlGenMode.QueryPureSelect:
                    return GeneratePureSelectQuery(opt, fullTableName, pureTableName, conds);

                case OraclePlSqlGenMode.QueryCursorLoop:
                    return GenerateCursorLoopBlock(opt, fullTableName, pureTableName, conds);

                case OraclePlSqlGenMode.QueryProcedure:
                    return GenerateQueryProcedure(opt, fullTableName, pureTableName, conds);

                case OraclePlSqlGenMode.QueryCSharpCode:
                    return GenerateQueryCSharpCode(opt, pureTableName, conds);

                default:
                    return GeneratePlSqlCheckBlock(opt, fullTableName, pureTableName, conds);
            }
        }

        private static string GeneratePlSqlCheckBlock(OracleSpecOptions opt, string fullTableName, string pureTableName, List<OracleConditionSpec> conds)
        {
            var sb = new StringBuilder();
            var q = opt.QuerySpec;
            string curName = $"csr{pureTableName}";
            string recName = $"rec{pureTableName}";

            sb.AppendLine("-- =============================================================================");
            sb.AppendLine($"-- ORACLE PL/SQL: Trích xuất / Kiểm tra data table {q.TableLogicalName} [{pureTableName}]");
            sb.AppendLine($"-- Cursor Name  : {curName}");
            sb.AppendLine("-- Sinh tự động từ đặc tả xử lý bảng tính (TKCT - 詳細設計書)");
            sb.AppendLine($"-- Thời gian sinh: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            sb.AppendLine("-- =============================================================================");
            sb.AppendLine("DECLARE");
            sb.AppendLine($"    -- 1. Khai báo cursor trích xuất data table {q.TableLogicalName} [{pureTableName}]");
            sb.AppendLine($"    CURSOR {curName} IS");
            sb.AppendLine("        SELECT *");
            sb.AppendLine($"        FROM {fullTableName}");

            if (conds.Count > 0)
            {
                sb.AppendLine("        WHERE");
                for (int i = 0; i < conds.Count; i++)
                {
                    var c = conds[i];
                    string prefix = (i == 0) ? "        " : "    AND ";
                    string comment = !string.IsNullOrWhiteSpace(c.LogicalName) ? $" -- {c.LogicalName}" : "";
                    sb.AppendLine($"        {prefix}{c.PhysicalColumn,-24} {c.Operator,-3} {c.TargetSqlValue}{comment}");
                }
            }
            sb.AppendLine("        ;");
            sb.AppendLine();
            sb.AppendLine($"    {recName} {curName}%ROWTYPE;");
            sb.AppendLine("BEGIN");
            sb.AppendLine($"    -- 2. Mở cursor {curName} và trích xuất dữ liệu từ bảng {pureTableName}");
            sb.AppendLine($"    OPEN {curName};");
            sb.AppendLine($"    FETCH {curName} INTO {recName};");
            sb.AppendLine();
            sb.AppendLine($"    -- 3. Kiểm tra kết quả trích xuất dữ liệu từ cursor");
            sb.AppendLine($"    IF {curName}%FOUND THEN");
            sb.AppendLine($"        CLOSE {curName};");

            if (!string.IsNullOrWhiteSpace(q.OnExistsAction))
            {
                sb.AppendLine($"        -- {q.OnExistsAction}");
                // Nếu có nhắc tới mã lỗi No.X thì sinh mã lỗi tương ứng
                var matchErr = Regex.Match(q.OnExistsAction, @"No\.?(\d+)", RegexOptions.IgnoreCase);
                int errNum = matchErr.Success ? int.Parse(matchErr.Groups[1].Value) : 5;
                sb.AppendLine($"        RAISE_APPLICATION_ERROR(-{20000 + errNum}, '{EscapeSqlString(q.OnExistsAction)}');");
            }
            else
            {
                sb.AppendLine($"        DBMS_OUTPUT.PUT_LINE('Dữ liệu đã tồn tại trong {pureTableName}');");
            }

            sb.AppendLine("    ELSE");
            sb.AppendLine($"        CLOSE {curName};");

            if (!string.IsNullOrWhiteSpace(q.OnNotExistsAction))
            {
                sb.AppendLine($"        -- {q.OnNotExistsAction}");
            }
            sb.AppendLine("        DBMS_OUTPUT.PUT_LINE('Thỏa mãn điều kiện: Không tìm thấy data, tiếp tục xử lý.');");
            sb.AppendLine("    END IF;");

            if (q.OtherNotes.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("    -- Ghi chú bổ sung từ đặc tả:");
                foreach (var note in q.OtherNotes)
                {
                    sb.AppendLine($"    -- * {note}");
                }
            }

            sb.AppendLine();
            sb.AppendLine("EXCEPTION");
            sb.AppendLine("    WHEN OTHERS THEN");
            sb.AppendLine($"        IF {curName}%ISOPEN THEN");
            sb.AppendLine($"            CLOSE {curName};");
            sb.AppendLine("        END IF;");
            sb.AppendLine($"        DBMS_OUTPUT.PUT_LINE('Lỗi khi trích xuất dữ liệu {pureTableName} qua cursor {curName}: ' || SQLERRM);");
            sb.AppendLine("        RAISE;");
            sb.AppendLine("END;");
            sb.AppendLine("/");

            return sb.ToString();
        }

        private static string GeneratePureSelectQuery(OracleSpecOptions opt, string fullTableName, string pureTableName, List<OracleConditionSpec> conds)
        {
            var sb = new StringBuilder();
            var q = opt.QuerySpec;

            sb.AppendLine("-- =============================================================================");
            sb.AppendLine($"-- SELECT QUERY: {fullTableName} ({q.TableLogicalName})");
            sb.AppendLine("-- =============================================================================");
            sb.AppendLine("SELECT");
            sb.AppendLine("    *");
            sb.AppendLine($"FROM {fullTableName}");

            if (conds.Count > 0)
            {
                sb.AppendLine("WHERE");
                for (int i = 0; i < conds.Count; i++)
                {
                    var c = conds[i];
                    string prefix = (i == 0) ? "    " : "AND ";
                    string comment = !string.IsNullOrWhiteSpace(c.LogicalName) ? $" -- {c.LogicalName}" : "";
                    sb.AppendLine($"    {prefix}{c.PhysicalColumn,-24} {c.Operator,-3} {c.TargetSqlValue}{comment}");
                }
            }
            sb.AppendLine(";");

            return sb.ToString();
        }

        private static string GenerateCursorLoopBlock(OracleSpecOptions opt, string fullTableName, string pureTableName, List<OracleConditionSpec> conds)
        {
            var sb = new StringBuilder();
            var q = opt.QuerySpec;
            string curName = $"csr{pureTableName}";
            string recName = $"rec{pureTableName}";

            sb.AppendLine("-- =============================================================================");
            sb.AppendLine($"-- PL/SQL CURSOR LOOP: {curName} ({pureTableName})");
            sb.AppendLine($"-- Table Logical : {q.TableLogicalName}");
            sb.AppendLine("-- =============================================================================");
            sb.AppendLine("DECLARE");
            sb.AppendLine($"    -- 1. Khai báo cursor trích xuất data table {q.TableLogicalName} [{pureTableName}]");
            sb.AppendLine($"    CURSOR {curName} IS");
            sb.AppendLine("        SELECT *");
            sb.AppendLine($"        FROM {fullTableName}");

            if (conds.Count > 0)
            {
                sb.AppendLine("        WHERE");
                for (int i = 0; i < conds.Count; i++)
                {
                    var c = conds[i];
                    string prefix = (i == 0) ? "        " : "    AND ";
                    string comment = !string.IsNullOrWhiteSpace(c.LogicalName) ? $" -- {c.LogicalName}" : "";
                    sb.AppendLine($"        {prefix}{c.PhysicalColumn,-24} {c.Operator,-3} {c.TargetSqlValue}{comment}");
                }
            }
            sb.AppendLine("        ;");
            sb.AppendLine();
            sb.AppendLine($"    {recName} {curName}%ROWTYPE;");
            sb.AppendLine("BEGIN");
            sb.AppendLine($"    -- 2. Mở cursor {curName} và duyệt từng dòng dữ liệu trích xuất");
            sb.AppendLine($"    OPEN {curName};");
            sb.AppendLine("    LOOP");
            sb.AppendLine($"        FETCH {curName} INTO {recName};");
            sb.AppendLine($"        EXIT WHEN {curName}%NOTFOUND;");
            sb.AppendLine();
            sb.AppendLine($"        -- Xử lý nghiệp vụ với từng dòng dữ liệu ({recName}.<COL>):");
            sb.AppendLine($"        DBMS_OUTPUT.PUT_LINE('Đang xử lý bản ghi từ {curName}...');");
            sb.AppendLine("    END LOOP;");
            sb.AppendLine($"    CLOSE {curName};");
            sb.AppendLine();
            sb.AppendLine($"    /*");
            sb.AppendLine($"    -- Hoặc dùng cú pháp FOR loop ngắn gọn:");
            sb.AppendLine($"    FOR {recName} IN {curName} LOOP");
            sb.AppendLine($"        -- Xử lý từng bản ghi {recName}...");
            sb.AppendLine($"    END LOOP;");
            sb.AppendLine($"    */");
            sb.AppendLine();
            sb.AppendLine("EXCEPTION");
            sb.AppendLine("    WHEN OTHERS THEN");
            sb.AppendLine($"        IF {curName}%ISOPEN THEN");
            sb.AppendLine($"            CLOSE {curName};");
            sb.AppendLine("        END IF;");
            sb.AppendLine($"        DBMS_OUTPUT.PUT_LINE('Lỗi khi duyệt cursor {curName}: ' || SQLERRM);");
            sb.AppendLine("        RAISE;");
            sb.AppendLine("END;");
            sb.AppendLine("/");

            return sb.ToString();
        }

        private static string GenerateQueryProcedure(OracleSpecOptions opt, string fullTableName, string pureTableName, List<OracleConditionSpec> conds)
        {
            var sb = new StringBuilder();
            var q = opt.QuerySpec;
            string spName = $"SP_CHECK_{pureTableName}";
            string curName = $"csr{pureTableName}";
            string recName = $"rec{pureTableName}";

            // Lấy danh sách các biến bind
            var bindParams = conds.Where(c => !string.IsNullOrEmpty(c.BindVariableName))
                                  .GroupBy(c => c.BindVariableName)
                                  .Select(g => g.First())
                                  .ToList();

            sb.AppendLine("-- =============================================================================");
            sb.AppendLine($"-- STORED PROCEDURE / FUNCTION: {spName}");
            sb.AppendLine($"-- Cursor Name: {curName}");
            sb.AppendLine("-- =============================================================================");
            sb.AppendLine($"CREATE OR REPLACE PROCEDURE {spName} (");

            for (int i = 0; i < bindParams.Count; i++)
            {
                var bp = bindParams[i];
                string comma = (i < bindParams.Count - 1 || true) ? "," : "";
                string comment = !string.IsNullOrWhiteSpace(bp.LogicalName) ? $" -- {bp.LogicalName}" : "";
                sb.AppendLine($"    p_{bp.BindVariableName,-26} IN  VARCHAR2{comma}{comment}");
            }

            sb.AppendLine("    p_exists_flag              OUT NUMBER   -- 1: Tồn tại, 0: Không tồn tại");
            sb.AppendLine(")");
            sb.AppendLine("IS");
            sb.AppendLine($"    -- 1. Khai báo cursor trích xuất data table {q.TableLogicalName} [{pureTableName}]");
            sb.AppendLine($"    CURSOR {curName} IS");
            sb.AppendLine("        SELECT *");
            sb.AppendLine($"        FROM {fullTableName}");

            if (conds.Count > 0)
            {
                sb.AppendLine("        WHERE");
                for (int i = 0; i < conds.Count; i++)
                {
                    var c = conds[i];
                    string prefix = (i == 0) ? "        " : "    AND ";
                    string valExpr = !string.IsNullOrEmpty(c.BindVariableName) ? $"p_{c.BindVariableName}" : c.TargetSqlValue;
                    string comment = !string.IsNullOrWhiteSpace(c.LogicalName) ? $" -- {c.LogicalName}" : "";
                    sb.AppendLine($"{prefix}{c.PhysicalColumn,-24} {c.Operator,-3} {valExpr}{comment}");
                }
            }
            sb.AppendLine("        ;");
            sb.AppendLine();
            sb.AppendLine($"    {recName} {curName}%ROWTYPE;");
            sb.AppendLine("BEGIN");
            sb.AppendLine($"    -- 2. Mở cursor {curName} và kiểm tra kết quả trích xuất dữ liệu");
            sb.AppendLine($"    OPEN {curName};");
            sb.AppendLine($"    FETCH {curName} INTO {recName};");
            sb.AppendLine();
            sb.AppendLine($"    IF {curName}%FOUND THEN");
            sb.AppendLine("        p_exists_flag := 1;");
            sb.AppendLine("    ELSE");
            sb.AppendLine("        p_exists_flag := 0;");
            sb.AppendLine("    END IF;");
            sb.AppendLine();
            sb.AppendLine($"    CLOSE {curName};");
            sb.AppendLine("EXCEPTION");
            sb.AppendLine("    WHEN OTHERS THEN");
            sb.AppendLine($"        IF {curName}%ISOPEN THEN");
            sb.AppendLine($"            CLOSE {curName};");
            sb.AppendLine("        END IF;");
            sb.AppendLine($"        RAISE_APPLICATION_ERROR(-20001, 'Lỗi trong {spName}: ' || SQLERRM);");
            sb.AppendLine("END;");
            sb.AppendLine("/");

            return sb.ToString();
        }

        private static string GenerateQueryCSharpCode(OracleSpecOptions opt, string pureTableName, List<OracleConditionSpec> conds)
        {
            var sb = new StringBuilder();
            sb.AppendLine("// =============================================================================");
            sb.AppendLine($"// C# DAPPER / ADO.NET QUERY: {pureTableName}");
            sb.AppendLine("// =============================================================================");
            sb.AppendLine("const string sql = @\"");
            sb.AppendLine("    SELECT *");
            sb.AppendLine($"    FROM {pureTableName}");

            if (conds.Count > 0)
            {
                sb.AppendLine("    WHERE");
                for (int i = 0; i < conds.Count; i++)
                {
                    var c = conds[i];
                    string prefix = (i == 0) ? "        " : "    AND ";
                    sb.AppendLine($"{prefix}{c.PhysicalColumn,-24} {c.Operator,-3} {c.TargetSqlValue}");
                }
            }
            sb.AppendLine("\";");
            sb.AppendLine();
            sb.AppendLine("var parameters = new DynamicParameters();");
            var bindParams = conds.Where(c => !string.IsNullOrEmpty(c.BindVariableName))
                                  .GroupBy(c => c.BindVariableName)
                                  .Select(g => g.First())
                                  .ToList();

            foreach (var bp in bindParams)
            {
                sb.AppendLine($"parameters.Add(\"{bp.BindVariableName}\", model.{ToPascalCase(bp.BindVariableName)}); // {bp.LogicalName}");
            }

            return sb.ToString();
        }

        // =========================================================================================
        // GENERATION CHO TABLE DEFINITION DDL SPEC (ĐẶC TẢ BẢNG)
        // =========================================================================================

        private static string GenerateDdlSpecScript(OracleSpecOptions options)
        {
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
                if (!string.IsNullOrEmpty(identityClause)) nullClause = string.Empty;

                string line = $"    {col.PhysicalName,-30} {typeDef,-18}{defaultClause}{identityClause}{nullClause}";

                if (i < cols.Count - 1 || pkCols.Count > 0)
                {
                    line += ",";
                }

                if (!string.IsNullOrWhiteSpace(col.LogicalName))
                {
                    line += $" -- {col.LogicalName}";
                }

                sb.AppendLine(line);
            }

            if (pkCols.Count > 0)
            {
                string pkColList = string.Join(", ", pkCols.Select(c => c.PhysicalName));
                sb.AppendLine($"    CONSTRAINT PK_{pureTableName} PRIMARY KEY ({pkColList})");
            }

            sb.AppendLine(");");
            sb.AppendLine();

            if (opt.IncludeComments)
            {
                string tblComment = !string.IsNullOrWhiteSpace(opt.TableLogicalName) ? opt.TableLogicalName : pureTableName;
                sb.AppendLine($"COMMENT ON TABLE {fullTableName} IS '{EscapeSqlString(tblComment)}';");
                sb.AppendLine();

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
            sb.AppendLine("-- =============================================================================");
            sb.AppendLine();

            var pkCols = cols.Where(c => c.IsPrimaryKey).ToList();
            if (pkCols.Count == 0) pkCols = cols.Take(1).ToList();
            var nonPkCols = cols.Where(c => !c.IsPrimaryKey).ToList();

            // INSERT
            sb.AppendLine($"CREATE OR REPLACE PROCEDURE SP_INS_{pureTableName} (");
            for (int i = 0; i < cols.Count; i++)
            {
                var col = cols[i];
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

            // UPDATE
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

            // DELETE
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
            sb.AppendLine("-- =============================================================================");
            sb.AppendLine();

            var pkCols = cols.Where(c => c.IsPrimaryKey).ToList();
            if (pkCols.Count == 0) pkCols = cols.Take(1).ToList();
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

            if (v.Equals("SYSDATE", StringComparison.OrdinalIgnoreCase) ||
                v.Equals("CURRENT_TIMESTAMP", StringComparison.OrdinalIgnoreCase) ||
                v.Equals("SYSTIMESTAMP", StringComparison.OrdinalIgnoreCase))
            {
                return v.ToUpperInvariant();
            }

            if (v.StartsWith("'") && v.EndsWith("'"))
            {
                return v;
            }

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
            if (t.StartsWith("NUMBER") || t == "DECIMAL" || t == "NUMERIC")
            {
                return "number";
            }
            if (t == "INT" || t == "INTEGER")
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
