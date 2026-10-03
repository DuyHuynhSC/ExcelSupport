using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using ExcelSupport.Models;
using Microsoft.Office.Interop.Excel;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;

namespace ExcelSupport.Services
{
    /// <summary>
    /// Service phân tích và chuyển đổi ngôn ngữ tự nhiên (hoặc nội dung ô tính Excel) thành câu truy vấn SQL Oracle.
    /// Hỗ trợ cả Engine Quy tắc nội bộ (Deterministic Rule-Based) cực nhanh 0ms và Engine AI (OpenAI / Qwen / Local LLM).
    /// </summary>
    public static class NaturalLanguageSqlService
    {
        private static readonly Regex TablePatternVi = new Regex(
            @"(?i)(?:trích\s+xuất|truy\s+vấn|lấy\s+dữ\s+liệu|chọn|xem)\s+(?:dữ\s+liệu\s+)?(?:từ\s+)?(?:bảng|table)?\s*[:\s]*([a-zA-Z0-9_#$]+)",
            RegexOptions.Compiled);

        private static readonly Regex TablePatternExplicit = new Regex(
            @"(?i)\b(?:table|bảng)\s*[:=\s]\s*([a-zA-Z0-9_#$]+)",
            RegexOptions.Compiled);

        private static readonly Regex TablePatternJa = new Regex(
            @"([a-zA-Z0-9_#$]+)\s*(?:テーブル|table)\s*(?:から|の|を)?\s*(?:抽出|取得|検索|参照)?",
            RegexOptions.Compiled);

        private static readonly Regex TablePatternJaExplicit = new Regex(
            @"(?:テーブル|table)\s*[:：\s]\s*([a-zA-Z0-9_#$]+)",
            RegexOptions.Compiled);

        private static readonly Regex TablePatternEn = new Regex(
            @"(?i)\b(?:extract|query|select|fetch|get)\s+(?:from\s+)?(?:table\s+)?([a-zA-Z0-9_#$]+)",
            RegexOptions.Compiled);

        private static readonly Regex TablePatternFrom = new Regex(
            @"(?i)\bFROM\s+([a-zA-Z0-9_#$]+)",
            RegexOptions.Compiled);

        private static readonly Regex BoilerplateHeaderPattern = new Regex(
            @"(?i)^(?:theo\s+key\s+như\s+sau|theo\s+điều\s+kiện|điều\s+kiện|dưới\s+đây|như\s+sau|以下のキーで抽出|以下のキー|以下の条件|検索条件|where|conditions|key|keys)[:：]?$",
            RegexOptions.Compiled);

        private static readonly Regex ConditionOperatorPattern = new Regex(
            @"(?i)\b(?:=|<(?!=)|<=|>(?!=)|>=|!=|<>|LIKE|NOT\s+LIKE|IN|NOT\s+IN|IS\s+NULL|IS\s+NOT\s+NULL|BETWEEN)\b|=",
            RegexOptions.Compiled);

        private static readonly Regex ColonConditionPattern = new Regex(
            @"^([a-zA-Z0-9_#$]+)\s*[:：]\s*(.+)$",
            RegexOptions.Compiled);

        /// <summary>
        /// Trích xuất nội dung từ vùng ô đang chọn (Selection) trên Excel và chuyển đổi thành câu lệnh SQL.
        /// An toàn bộ nhớ, tự động thu hẹp theo UsedRange nếu người dùng chọn toàn bộ cột/dòng.
        /// </summary>
        public static bool TryExtractFromSelection(ExcelApp? app, out string generatedSql, out string infoMsg, out string detectedTable)
        {
            generatedSql = string.Empty;
            infoMsg = string.Empty;
            detectedTable = string.Empty;

            if (app == null) return false;

            try
            {
                dynamic? sel = null;
                try { sel = app.Selection; } catch { return false; }
                if (sel is not Range rng) return false;

                // Giới hạn vùng ô nếu chọn quá rộng (tránh OutOfMemory khi bấm chọn cả cột)
                Range evalRange = rng;
                try
                {
                    if (rng.CountLarge > 1000)
                    {
                        var used = rng.Worksheet?.UsedRange;
                        if (used != null)
                        {
                            var inter = app.Intersect(rng, used);
                            if (inter != null) evalRange = inter;
                            else return false;
                        }
                    }
                }
                catch { }

                var rawRows = ExtractRowsFromRange(evalRange);
                if (rawRows.Count == 0) return false;

                return TryParseRowsToSql(rawRows, out generatedSql, out infoMsg, out detectedTable);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[NaturalLanguageSqlService] TryExtractFromSelection error: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Phân tích văn bản tự do (ví dụ dán vào hoặc từ Editor) thành câu lệnh SQL.
        /// </summary>
        public static bool TryParseTextToSql(string text, out string generatedSql, out string infoMsg, out string detectedTable)
        {
            generatedSql = string.Empty;
            infoMsg = string.Empty;
            detectedTable = string.Empty;

            if (string.IsNullOrWhiteSpace(text)) return false;

            var lines = text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries)
                            .Select(l => l.Trim())
                            .Where(l => !string.IsNullOrWhiteSpace(l))
                            .ToList();

            var rows = lines.Select(l => new List<string> { l } as IList<string>).ToList();
            return TryParseRowsToSql(rows, out generatedSql, out infoMsg, out detectedTable);
        }

        /// <summary>
        /// Phân tích danh sách các dòng ô dữ liệu thành câu lệnh SQL Oracle.
        /// </summary>
        public static bool TryParseRowsToSql(IEnumerable<IList<string>> rows, out string generatedSql, out string infoMsg, out string detectedTable)
        {
            generatedSql = string.Empty;
            infoMsg = string.Empty;
            detectedTable = string.Empty;

            if (rows == null) return false;

            // 1. Kiểm tra xem toàn bộ các ô đã là câu lệnh SQL thuần túy (SELECT ... FROM ...) chưa
            var combinedTextBuilder = new StringBuilder();
            foreach (var r in rows)
            {
                foreach (var cell in r)
                {
                    if (!string.IsNullOrWhiteSpace(cell))
                    {
                        combinedTextBuilder.AppendLine(cell);
                    }
                }
            }
            string combinedAll = combinedTextBuilder.ToString().Trim();
            if (combinedAll.StartsWith("SELECT ", StringComparison.OrdinalIgnoreCase) ||
                combinedAll.StartsWith("WITH ", StringComparison.OrdinalIgnoreCase))
            {
                generatedSql = combinedAll.TrimEnd(';');
                detectedTable = ExtractTableNameFromSql(generatedSql);
                infoMsg = LocalizationService.Get("Oracle_NlDetectedStatus", detectedTable);
                return true;
            }

            // 2. Chuyển đổi các ô thành các dòng lệnh logic
            var logicalLines = new List<string>();
            foreach (var row in rows)
            {
                if (row == null || row.Count == 0) continue;

                var nonBlankCells = row.Where(c => !string.IsNullOrWhiteSpace(c)).Select(c => c.Trim()).ToList();
                if (nonBlankCells.Count == 0) continue;

                if (nonBlankCells.Count == 1)
                {
                    // 1 cell có thể chứa nhiều dòng (Alt+Enter)
                    string cellContent = nonBlankCells[0];
                    var subLines = cellContent.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (var sl in subLines)
                    {
                        string trimmed = CleanBulletPrefix(sl);
                        if (!string.IsNullOrWhiteSpace(trimmed))
                        {
                            logicalLines.Add(trimmed);
                        }
                    }
                }
                else if (nonBlankCells.Count >= 2)
                {
                    // Dạng bảng 2 cột: Cột 1 là Tên Cột / Khóa, Cột 2 là Giá trị
                    string c1 = CleanBulletPrefix(nonBlankCells[0]);
                    string c2 = CleanBulletPrefix(nonBlankCells[1]);

                    // Nếu Cột 1 chứa khai báo bảng (VD: "Trích xuất table XXXX")
                    if (IsTableDeclaration(c1))
                    {
                        logicalLines.Add(c1);
                        if (!string.IsNullOrWhiteSpace(c2) && !IsTableDeclaration(c2))
                        {
                            logicalLines.Add(c2);
                        }
                    }
                    else if (ConditionOperatorPattern.IsMatch(c1))
                    {
                        // Bản thân c1 đã là biểu thức điều kiện (VD: "AAAA = '1'")
                        logicalLines.Add(c1);
                    }
                    else if (IsValidColumnIdentifier(c1))
                    {
                        // C1 là tên cột, C2 là giá trị -> ghép thành "C1 = C2"
                        string formattedVal = FormatConditionValue(c2);
                        logicalLines.Add($"{c1} = {formattedVal}");
                    }
                    else
                    {
                        logicalLines.Add(c1);
                        logicalLines.Add(c2);
                    }
                }
            }

            if (logicalLines.Count == 0) return false;

            // 3. Tìm tên bảng trong các dòng logic
            string? foundTableName = null;
            int tableLineIndex = -1;

            for (int i = 0; i < logicalLines.Count; i++)
            {
                string line = logicalLines[i];
                string? table = ExtractTableNameFromLine(line);
                if (!string.IsNullOrEmpty(table))
                {
                    foundTableName = table;
                    tableLineIndex = i;
                    break;
                }
            }

            if (string.IsNullOrEmpty(foundTableName))
            {
                return false;
            }

            detectedTable = foundTableName!.ToUpperInvariant();

            // 4. Trích xuất các điều kiện WHERE từ các dòng còn lại
            var conditions = new List<string>();

            for (int i = 0; i < logicalLines.Count; i++)
            {
                if (i == tableLineIndex) continue; // Bỏ qua dòng khai báo tên bảng

                string line = logicalLines[i].Trim();
                if (string.IsNullOrWhiteSpace(line)) continue;

                // Bỏ qua các dòng tiêu đề dẫn chuyện: "theo key như sau", "điều kiện:", "như sau:"
                if (BoilerplateHeaderPattern.IsMatch(line)) continue;

                // Chuẩn hóa dấu nháy kép thành nháy đơn SQL
                line = NormalizeQuotes(line);

                // Dạng 1: Đã chứa toán tử so sánh (=, >, <, LIKE, IN, IS NULL...)
                if (ConditionOperatorPattern.IsMatch(line))
                {
                    string cond = CleanConditionLine(line);
                    if (!string.IsNullOrWhiteSpace(cond) && !conditions.Contains(cond))
                    {
                        conditions.Add(cond);
                    }
                    continue;
                }

                // Dạng 2: Dạng dấu hai chấm (VD: "AAAA: '1'" hoặc "AAAA: 123")
                var colonMatch = ColonConditionPattern.Match(line);
                if (colonMatch.Success)
                {
                    string col = colonMatch.Groups[1].Value.Trim().ToUpperInvariant();
                    string val = FormatConditionValue(colonMatch.Groups[2].Value.Trim());
                    string cond = $"{col} = {val}";
                    if (!conditions.Contains(cond))
                    {
                        conditions.Add(cond);
                    }
                    continue;
                }
            }

            // 5. Sinh câu lệnh SQL Oracle chuẩn
            var sb = new StringBuilder();
            sb.Append($"SELECT * FROM {detectedTable}");

            if (conditions.Count > 0)
            {
                sb.Append(" WHERE ");
                sb.Append(string.Join(" AND ", conditions));
            }

            generatedSql = sb.ToString();
            infoMsg = LocalizationService.Get("Oracle_NlDetectedStatus", detectedTable);
            return true;
        }

        /// <summary>
        /// Sử dụng AI nội bộ hoặc OpenAI/Qwen để chuyển đổi câu hỏi/ngôn ngữ tự nhiên phức tạp thành câu lệnh Oracle SQL.
        /// </summary>
        public static async Task<string> ConvertWithAiAsync(string naturalLanguage, AiConfig? config = null)
        {
            var targetConfig = config ?? AiConfigManager.Current;
            if (targetConfig == null || string.IsNullOrWhiteSpace(targetConfig.BaseUrl))
            {
                throw new InvalidOperationException("Chưa cấu hình máy chủ AI (Base URL). Vui lòng cấu hình trong Ribbon -> Settings -> Cấu hình AI.");
            }

            string systemPrompt = 
                "Bạn là chuyên gia cơ sở dữ liệu Oracle SQL cao cấp. " +
                "Nhiệm vụ của bạn là chuyển đổi yêu cầu ngôn ngữ tự nhiên từ người dùng bảng tính Excel thành một câu lệnh Oracle SQL chuẩn xác, an toàn và tối ưu.\n" +
                "QUY TẮC BẮT BUỘC:\n" +
                "1. Chỉ trả về duy nhất chuỗi SQL thuần túy. TUYỆT ĐỐI KHÔNG dùng khối markdown (không dùng ```sql hay ```), không giải thích, không thêm bất kỳ văn bản nào khác.\n" +
                "2. Sử dụng cú pháp Oracle SQL: Tên bảng và tên cột viết hoa (UPPERCASE). Chuỗi ký tự đặt trong dấu nháy đơn '...'.\n" +
                "3. Nếu có giới hạn dòng hoặc top, sử dụng điều kiện ROWNUM <= N hoặc FETCH FIRST N ROWS ONLY.";

            string userPrompt = naturalLanguage.Trim();

            string rawReply = await OpenAiClientService.SendChatAsync(targetConfig, userPrompt, systemPrompt);
            if (string.IsNullOrWhiteSpace(rawReply))
            {
                throw new InvalidOperationException("Máy chủ AI không trả về phản hồi nào.");
            }

            // Dọn dẹp nếu AI vẫn bọc trong markdown code fence ```sql
            string cleaned = rawReply.Trim();
            if (cleaned.StartsWith("```sql", StringComparison.OrdinalIgnoreCase))
            {
                cleaned = cleaned.Substring(6).Trim();
            }
            else if (cleaned.StartsWith("```", StringComparison.OrdinalIgnoreCase))
            {
                cleaned = cleaned.Substring(3).Trim();
            }

            if (cleaned.EndsWith("```", StringComparison.OrdinalIgnoreCase))
            {
                cleaned = cleaned.Substring(0, cleaned.Length - 3).Trim();
            }

            return cleaned.Trim().TrimEnd(';');
        }

        #region Helper Methods

        private static List<IList<string>> ExtractRowsFromRange(Range rng)
        {
            var result = new List<IList<string>>();
            if (rng == null) return result;

            try
            {
                int areaCount = rng.Areas.Count;
                for (int a = 1; a <= areaCount; a++)
                {
                    Range area = rng.Areas[a];
                    object raw = area.Value2;

                    if (raw is object[,] arr)
                    {
                        int rows = arr.GetLength(0);
                        int cols = arr.GetLength(1);

                        for (int r = 1; r <= rows; r++)
                        {
                            var rowList = new List<string>(cols);
                            for (int c = 1; c <= cols; c++)
                            {
                                object val = arr[r, c];
                                rowList.Add(val?.ToString() ?? string.Empty);
                            }
                            if (rowList.Any(s => !string.IsNullOrWhiteSpace(s)))
                            {
                                result.Add(rowList);
                            }
                        }
                    }
                    else if (raw != null)
                    {
                        string str = raw.ToString() ?? string.Empty;
                        if (!string.IsNullOrWhiteSpace(str))
                        {
                            result.Add(new List<string> { str });
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[NaturalLanguageSqlService] ExtractRowsFromRange error: {ex.Message}");
            }

            return result;
        }

        private static string CleanBulletPrefix(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return string.Empty;
            // Xóa gạch đầu dòng: - , * , ・ , + , • , 1. , v.v.
            return Regex.Replace(input.Trim(), @"^[\s\-\*・•+>#]+|^\d+[\.\)]\s*", "").Trim();
        }

        private static bool IsTableDeclaration(string line)
        {
            return TablePatternVi.IsMatch(line) ||
                   TablePatternExplicit.IsMatch(line) ||
                   TablePatternJa.IsMatch(line) ||
                   TablePatternJaExplicit.IsMatch(line) ||
                   TablePatternEn.IsMatch(line) ||
                   TablePatternFrom.IsMatch(line);
        }

        private static string? ExtractTableNameFromLine(string line)
        {
            if (string.IsNullOrWhiteSpace(line)) return null;

            // Pattern Tiếng Việt: "Trích xuất table XXXX theo key như sau"
            var mVi = TablePatternVi.Match(line);
            if (mVi.Success && IsValidTableName(mVi.Groups[1].Value))
            {
                return mVi.Groups[1].Value;
            }

            // Pattern Rõ ràng: "Table: XXXX" hoặc "Bảng: XXXX"
            var mExp = TablePatternExplicit.Match(line);
            if (mExp.Success && IsValidTableName(mExp.Groups[1].Value))
            {
                return mExp.Groups[1].Value;
            }

            // Pattern Tiếng Nhật: "XXXXテーブルから以下のキーで抽出"
            var mJa = TablePatternJa.Match(line);
            if (mJa.Success && IsValidTableName(mJa.Groups[1].Value))
            {
                return mJa.Groups[1].Value;
            }

            // Pattern Tiếng Nhật: "テーブル: XXXX"
            var mJaExp = TablePatternJaExplicit.Match(line);
            if (mJaExp.Success && IsValidTableName(mJaExp.Groups[1].Value))
            {
                return mJaExp.Groups[1].Value;
            }

            // Pattern Tiếng Anh: "Extract table XXXX with key:"
            var mEn = TablePatternEn.Match(line);
            if (mEn.Success && IsValidTableName(mEn.Groups[1].Value))
            {
                return mEn.Groups[1].Value;
            }

            // Pattern SQL FROM: "FROM XXXX"
            var mFrom = TablePatternFrom.Match(line);
            if (mFrom.Success && IsValidTableName(mFrom.Groups[1].Value))
            {
                return mFrom.Groups[1].Value;
            }

            return null;
        }

        private static bool IsValidTableName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return false;
            // Bỏ qua các từ khóa gây nhiễu
            string upper = name.Trim().ToUpperInvariant();
            if (upper == "TABLE" || upper == "THEO" || upper == "KEY" || upper == "NHU" || upper == "SAU" || upper == "DU" || upper == "LIEU")
            {
                return false;
            }
            return Regex.IsMatch(upper, @"^[A-Z0-9_#$]+$");
        }

        private static bool IsValidColumnIdentifier(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return false;
            return Regex.IsMatch(name.Trim(), @"^[a-zA-Z0-9_#$]+$");
        }

        private static string NormalizeQuotes(string text)
        {
            // Thay dấu ngoặc kép hoặc ngoặc nghiêng kiểu Word (“ ”) thành nháy đơn chuẩn SQL '
            return text.Replace('"', '\'')
                       .Replace('“', '\'')
                       .Replace('”', '\'')
                       .Replace('‘', '\'')
                       .Replace('’', '\'');
        }

        private static string CleanConditionLine(string line)
        {
            string cleaned = line.Trim().TrimEnd(';', ',');
            // Viết hoa tên cột ở vế trái dấu so sánh nếu là định danh hợp lệ
            var match = Regex.Match(cleaned, @"^([a-zA-Z0-9_#$]+)\s*(=|<(?!=)|<=|>(?!=)|>=|!=|<>|LIKE|IN)\s*(.+)$", RegexOptions.IgnoreCase);
            if (match.Success)
            {
                string col = match.Groups[1].Value.ToUpperInvariant();
                string op = match.Groups[2].Value.ToUpperInvariant();
                string val = match.Groups[3].Value.Trim();
                return $"{col} {op} {val}";
            }
            return cleaned;
        }

        private static string FormatConditionValue(string rawVal)
        {
            if (string.IsNullOrWhiteSpace(rawVal)) return "NULL";

            string trimmed = rawVal.Trim();
            // Nếu đã có dấu nháy đơn
            if (trimmed.StartsWith("'") && trimmed.EndsWith("'"))
            {
                return trimmed;
            }

            // Nếu là số thuần túy (nguyên hoặc thập phân)
            if (decimal.TryParse(trimmed, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out _))
            {
                return trimmed;
            }

            // Nếu là NULL, SYSDATE, CURRENT_DATE
            string upper = trimmed.ToUpperInvariant();
            if (upper == "NULL" || upper == "SYSDATE" || upper.StartsWith("TO_DATE(") || upper.StartsWith("DATE '"))
            {
                return upper;
            }

            // Mặc định bọc dấu nháy đơn cho kiểu chuỗi
            return $"'{trimmed.Replace("'", "''")}'";
        }

        private static string ExtractTableNameFromSql(string sql)
        {
            try
            {
                var match = Regex.Match(sql, @"\bFROM\s+([""']?(?<schema>[a-zA-Z0-9_]+)[""']?\.)?([""']?(?<table>[a-zA-Z0-9_]+)[""']?)", RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    return match.Groups["table"].Value.ToUpperInvariant();
                }
            }
            catch { }
            return "QUERY_RESULT";
        }

        #endregion
    }
}
