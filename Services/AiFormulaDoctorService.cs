using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using ExcelSupport.Models;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json.Linq;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;

namespace ExcelSupport.Services
{
    public static class AiFormulaDoctorService
    {
        #region Error Scanning

        public static FormulaDoctorScanResult ScanForErrors(ExcelApp app, bool scanSelectionOnly = false)
        {
            var result = new FormulaDoctorScanResult();
            var sw = Stopwatch.StartNew();

            try
            {
                var activeSheet = app.ActiveSheet as Worksheet;
                if (activeSheet == null) return result;

                result.WorkbookName = activeSheet.Parent is Workbook wb ? wb.Name : string.Empty;
                result.SheetName = activeSheet.Name;

                Range? scanRange = null;
                if (scanSelectionOnly && app.Selection is Range selRange)
                {
                    scanRange = selRange;
                }
                else
                {
                    scanRange = activeSheet.UsedRange;
                }

                if (scanRange == null) return result;

                // Try to use SpecialCells for rapid error detection
                Range? errorCells = null;
                try
                {
                    errorCells = scanRange.SpecialCells(XlCellType.xlCellTypeFormulas, (int)XlSpecialCellsValue.xlErrors);
                }
                catch
                {
                    // No special error formula cells found via SpecialCells, or range is single cell
                }

                if (errorCells != null)
                {
                    foreach (Range area in errorCells.Areas)
                    {
                        foreach (Range cell in area)
                        {
                            var item = BuildFormulaCellItem(cell, activeSheet);
                            if (item != null)
                            {
                                result.ErrorItems.Add(item);
                            }
                        }
                    }
                }
                else
                {
                    // Fallback scan: if small range or single cell selection
                    int totalCells = scanRange.CountLarge > 50000 ? 50000 : (int)scanRange.CountLarge;
                    result.TotalCellsScanned = totalCells;

                    if (totalCells <= 20000)
                    {
                        foreach (Range cell in scanRange)
                        {
                            if (cell.HasFormula == true)
                            {
                                string textVal = cell.Text?.ToString() ?? string.Empty;
                                if (IsErrorText(textVal))
                                {
                                    var item = BuildFormulaCellItem(cell, activeSheet);
                                    if (item != null)
                                    {
                                        result.ErrorItems.Add(item);
                                    }
                                }
                            }
                        }
                    }
                }

                result.TotalErrorsFound = result.ErrorItems.Count;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[AiFormulaDoctorService] Scan error: {ex.Message}");
            }
            finally
            {
                sw.Stop();
                result.ScanDuration = sw.Elapsed;
            }

            return result;
        }

        private static FormulaCellItem? BuildFormulaCellItem(Range cell, Worksheet sheet)
        {
            try
            {
                string textVal = cell.Text?.ToString() ?? string.Empty;
                string formula = cell.Formula?.ToString() ?? string.Empty;

                if (string.IsNullOrWhiteSpace(formula)) return null;

                var errType = ParseErrorType(textVal);

                var item = new FormulaCellItem
                {
                    CellAddress = cell.Address[false, false],
                    Row = cell.Row,
                    Column = cell.Column,
                    SheetName = sheet.Name,
                    Formula = formula,
                    DisplayValue = textVal,
                    ErrorType = errType
                };

                // Get column header if available (row 1)
                try
                {
                    if (cell.Row > 1)
                    {
                        var headerCell = sheet.Cells[1, cell.Column] as Range;
                        item.HeaderText = headerCell?.Text?.ToString() ?? string.Empty;
                    }
                }
                catch { }

                // Get nearby precedent info (simple heuristic)
                try
                {
                    var matches = Regex.Matches(formula, @"[A-Za-z]{1,3}\d{1,7}");
                    foreach (Match m in matches.Cast<Match>().Take(4))
                    {
                        string refAddr = m.Value;
                        try
                        {
                            var refCell = sheet.Range[refAddr];
                            string refVal = refCell.Text?.ToString() ?? "empty";
                            item.PrecedentValues.Add($"{refAddr}={refVal}");
                        }
                        catch { }
                    }
                }
                catch { }

                return item;
            }
            catch
            {
                return null;
            }
        }

        private static bool IsErrorText(string text)
        {
            if (string.IsNullOrEmpty(text)) return false;
            return text.StartsWith("#") && (
                text.Equals("#N/A", StringComparison.OrdinalIgnoreCase) ||
                text.Equals("#VALUE!", StringComparison.OrdinalIgnoreCase) ||
                text.Equals("#REF!", StringComparison.OrdinalIgnoreCase) ||
                text.Equals("#DIV/0!", StringComparison.OrdinalIgnoreCase) ||
                text.Equals("#NAME?", StringComparison.OrdinalIgnoreCase) ||
                text.Equals("#NUM!", StringComparison.OrdinalIgnoreCase) ||
                text.Equals("#NULL!", StringComparison.OrdinalIgnoreCase) ||
                text.Equals("#CALC!", StringComparison.OrdinalIgnoreCase) ||
                text.Equals("#SPILL!", StringComparison.OrdinalIgnoreCase)
            );
        }

        private static FormulaErrorType ParseErrorType(string text)
        {
            if (string.IsNullOrEmpty(text)) return FormulaErrorType.Unknown;
            string t = text.Trim().ToUpperInvariant();
            if (t.Contains("#N/A")) return FormulaErrorType.NA;
            if (t.Contains("#VALUE!")) return FormulaErrorType.Value;
            if (t.Contains("#REF!")) return FormulaErrorType.Ref;
            if (t.Contains("#DIV/0!")) return FormulaErrorType.Div0;
            if (t.Contains("#NAME?")) return FormulaErrorType.Name;
            if (t.Contains("#NUM!")) return FormulaErrorType.Num;
            if (t.Contains("#NULL!")) return FormulaErrorType.Null;
            if (t.Contains("#CALC!")) return FormulaErrorType.Calc;
            if (t.Contains("#SPILL!")) return FormulaErrorType.Spill;
            return FormulaErrorType.Unknown;
        }

        #endregion

        #region AI Diagnosis & Heuristic Fixing

        public static async Task DiagnoseAndProposeFixAsync(FormulaCellItem item, AiConfig? aiConfig, AppLanguage lang = AppLanguage.Vietnamese)
        {
            // First run offline heuristic rule engine
            GenerateHeuristicFix(item, lang);

            // If AI is configured, query LLM for deeper context-aware diagnosis & superior formula
            if (aiConfig != null && !string.IsNullOrWhiteSpace(aiConfig.BaseUrl))
            {
                try
                {
                    string langName = lang switch
                    {
                        AppLanguage.Japanese => "Japanese",
                        AppLanguage.English => "English",
                        _ => "Vietnamese"
                    };

                    string precedents = item.PrecedentValues.Count > 0 ? string.Join(", ", item.PrecedentValues) : "None";
                    string prompt = $@"You are an expert Excel formula troubleshooter. Analyze this Excel error and provide a fix.

- Sheet: '{item.SheetName}', Cell: {item.CellAddress}, Header: '{item.HeaderText}'
- Current Formula: `{item.Formula}`
- Error Encountered: `{item.DisplayValue}` ({item.ErrorTypeName})
- Referenced Cell Values: {precedents}

Respond in {langName} in strictly valid JSON format with this exact structure:
{{
  ""diagnosis"": ""Short clear explanation of why this error occurred (1-2 sentences)"",
  ""proposed_formula"": ""=CORRECTED_EXCEL_FORMULA"",
  ""fix_explanation"": ""Short explanation of what the new formula changes and why it works""
}}";

                    var result = await QueryAiJsonAsync(prompt, aiConfig);
                    if (result != null)
                    {
                        string? diag = result["diagnosis"]?.ToString();
                        string? prop = result["proposed_formula"]?.ToString();
                        string? expl = result["fix_explanation"]?.ToString();

                        if (!string.IsNullOrWhiteSpace(prop) && prop.StartsWith("="))
                        {
                            item.AiDiagnosis = diag ?? item.AiDiagnosis;
                            item.ProposedFormula = prop.Trim();
                            item.FixExplanation = expl ?? item.FixExplanation;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[AiFormulaDoctorService] AI query error: {ex.Message}");
                }
            }
        }

        private static void GenerateHeuristicFix(FormulaCellItem item, AppLanguage lang)
        {
            string f = item.Formula.Trim();
            if (!f.StartsWith("=")) f = "=" + f;

            switch (item.ErrorType)
            {
                case FormulaErrorType.NA:
                    if (f.IndexOf("VLOOKUP", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        // Wrap with IFERROR or convert to XLOOKUP
                        item.AiDiagnosis = lang switch
                        {
                            AppLanguage.Japanese => "VLOOKUPで検索値が見つからないため #N/A が発生しています。",
                            AppLanguage.English => "VLOOKUP could not find the lookup value, resulting in #N/A.",
                            _ => "Không tìm thấy giá trị tìm kiếm trong bảng tham chiếu (hoặc có khoảng trắng thừa)."
                        };
                        item.ProposedFormula = $"=IFERROR({f.Substring(1)}, \"N/A\")";
                        item.FixExplanation = lang switch
                        {
                            AppLanguage.Japanese => "IFERRORでエラー時に代替文字列を表示するように保護しました。",
                            AppLanguage.English => "Wrapped with IFERROR to handle missing values gracefully.",
                            _ => "Bọc hàm IFERROR để trả về giá trị mặc định 'N/A' thay vì báo lỗi đỏ."
                        };
                    }
                    else if (f.IndexOf("MATCH", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        item.AiDiagnosis = lang switch
                        {
                            AppLanguage.Japanese => "MATCH関数で一致する項目が見つかりませんでした。",
                            AppLanguage.English => "MATCH did not find an exact match.",
                            _ => "Hàm MATCH không tìm thấy giá trị tương ứng."
                        };
                        item.ProposedFormula = $"=IFERROR({f.Substring(1)}, 0)";
                        item.FixExplanation = "Bọc IFERROR trả về 0.";
                    }
                    else
                    {
                        item.AiDiagnosis = "Công thức trả về #N/A do thiếu dữ liệu nguồn.";
                        item.ProposedFormula = $"=IFERROR({f.Substring(1)}, \"\")";
                        item.FixExplanation = "Bọc IFERROR để làm sạch bảng tính.";
                    }
                    break;

                case FormulaErrorType.Div0:
                    item.AiDiagnosis = lang switch
                    {
                        AppLanguage.Japanese => "分母が0または空白セルのため #DIV/0! エラーが発生しています。",
                        AppLanguage.English => "Division by zero or blank cell resulted in #DIV/0!.",
                        _ => "Mẫu số bằng 0 hoặc tham chiếu vào ô trống dẫn đến lỗi chia cho 0."
                    };
                    item.ProposedFormula = $"=IFERROR({f.Substring(1)}, 0)";
                    item.FixExplanation = lang switch
                    {
                        AppLanguage.Japanese => "IFERRORで分母が0の場合に 0 を返すように保護しました。",
                        AppLanguage.English => "Protected with IFERROR to return 0 when dividing by zero.",
                        _ => "Bọc hàm IFERROR để trả về 0 khi mẫu số bằng 0."
                    };
                    break;

                case FormulaErrorType.Value:
                    item.AiDiagnosis = lang switch
                    {
                        AppLanguage.Japanese => "数値計算にテキストや無効なデータ型が含まれています。",
                        AppLanguage.English => "A text value or mismatched data type was used in a numeric operation.",
                        _ => "Phép tính số học chứa ô có định dạng văn bản (Text) hoặc ký tự không hợp lệ."
                    };
                    item.ProposedFormula = $"=IFERROR({f.Substring(1)}, 0)";
                    item.FixExplanation = "Bọc IFERROR xử lý giá trị không đồng nhất.";
                    break;

                case FormulaErrorType.Ref:
                    item.AiDiagnosis = lang switch
                    {
                        AppLanguage.Japanese => "参照先の行、列、またはシートが削除されたため #REF! が発生しています。",
                        AppLanguage.English => "The referenced cell, row, column, or sheet was deleted.",
                        _ => "Ô, dòng, cột hoặc Sheet được tham chiếu đã bị xóa khỏi Workbook."
                    };
                    item.ProposedFormula = f.Replace("#REF!", "A1");
                    item.FixExplanation = "Cần cập nhật lại địa chỉ ô thay thế vị trí #REF!.";
                    break;

                case FormulaErrorType.Name:
                    // Check for common typo in function names
                    string corrected = f;
                    corrected = Regex.Replace(corrected, @"\bVLOOKP\b", "VLOOKUP", RegexOptions.IgnoreCase);
                    corrected = Regex.Replace(corrected, @"\bSUMM\b", "SUM", RegexOptions.IgnoreCase);
                    corrected = Regex.Replace(corrected, @"\bAVERGE\b", "AVERAGE", RegexOptions.IgnoreCase);
                    corrected = Regex.Replace(corrected, @"\bCOUNTAA\b", "COUNTA", RegexOptions.IgnoreCase);
                    corrected = Regex.Replace(corrected, @"\bCONCATNATE\b", "CONCATENATE", RegexOptions.IgnoreCase);

                    item.AiDiagnosis = lang switch
                    {
                        AppLanguage.Japanese => "関数名のスペルミス、または未定義の名前が使われています。",
                        AppLanguage.English => "Misspelled function name or unrecognized defined name.",
                        _ => "Tên hàm bị gõ sai chính tả hoặc vùng tên (Named Range) chưa được định nghĩa."
                    };
                    item.ProposedFormula = corrected != f ? corrected : $"=IFERROR({f.Substring(1)}, \"\")";
                    item.FixExplanation = corrected != f ? "Đã sửa lại tên hàm đúng chính tả chuẩn Excel." : "Bọc IFERROR.";
                    break;

                default:
                    item.AiDiagnosis = "Lỗi công thức Excel cần kiểm tra.";
                    item.ProposedFormula = $"=IFERROR({f.Substring(1)}, \"\")";
                    item.FixExplanation = "Bọc IFERROR.";
                    break;
            }
        }

        #endregion

        #region Applying Fixes

        public static bool ApplyFixToCell(ExcelApp app, FormulaCellItem item)
        {
            if (string.IsNullOrWhiteSpace(item.ProposedFormula)) return false;

            try
            {
                var ws = app.ActiveWorkbook?.Worksheets[item.SheetName] as Worksheet ?? app.ActiveSheet as Worksheet;
                if (ws == null) return false;

                var cell = ws.Range[item.CellAddress];
                cell.Formula = item.ProposedFormula;
                item.IsFixed = true;
                item.Formula = item.ProposedFormula;
                item.DisplayValue = cell.Text?.ToString() ?? string.Empty;
                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[AiFormulaDoctorService] ApplyFix error: {ex.Message}");
                return false;
            }
        }

        public static int BatchApplyFixToColumn(ExcelApp app, FormulaCellItem templateItem, List<FormulaCellItem> allItems)
        {
            int count = 0;
            if (string.IsNullOrWhiteSpace(templateItem.ProposedFormula)) return 0;

            try
            {
                var ws = app.ActiveWorkbook?.Worksheets[templateItem.SheetName] as Worksheet ?? app.ActiveSheet as Worksheet;
                if (ws == null) return 0;

                // Fix the template cell
                if (ApplyFixToCell(app, templateItem)) count++;

                // Propagate formula down to other error cells in same column if structure matches
                var sameColItems = allItems.Where(i => i.Column == templateItem.Column && i != templateItem && !i.IsFixed).ToList();

                foreach (var other in sameColItems)
                {
                    try
                    {
                        // Adjust row index in formula
                        string adjustedFormula = AdjustFormulaRow(templateItem.ProposedFormula, templateItem.Row, other.Row);
                        other.ProposedFormula = adjustedFormula;
                        if (ApplyFixToCell(app, other))
                        {
                            count++;
                        }
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[AiFormulaDoctorService] Batch fix error: {ex.Message}");
            }

            return count;
        }

        private static string AdjustFormulaRow(string formula, int originalRow, int targetRow)
        {
            int rowDiff = targetRow - originalRow;
            if (rowDiff == 0) return formula;

            return Regex.Replace(formula, @"(?<=[A-Za-z])(?<!\$)(?<row>\d{1,7})\b", m =>
            {
                if (int.TryParse(m.Groups["row"].Value, out int r))
                {
                    int newRow = r + rowDiff;
                    return newRow > 0 ? newRow.ToString() : m.Value;
                }
                return m.Value;
            });
        }

        #endregion

        #region Explain & Modernize

        public static async Task<FormulaExplainResult> ExplainFormulaAsync(string formula, AiConfig? aiConfig, AppLanguage lang = AppLanguage.Vietnamese)
        {
            var result = new FormulaExplainResult { OriginalFormula = formula };

            string langName = lang switch
            {
                AppLanguage.Japanese => "Japanese",
                AppLanguage.English => "English",
                _ => "Vietnamese"
            };

            if (aiConfig != null && !string.IsNullOrWhiteSpace(aiConfig.BaseUrl))
            {
                try
                {
                    string prompt = $@"You are an Excel expert. Explain this Excel formula step-by-step in {langName}:
Formula: `{formula}`

Respond in strictly valid JSON format:
{{
  ""overall_purpose"": ""High-level summary of what this formula achieves"",
  ""return_type"": ""Expected output type (Number, Text, Date, Boolean)"",
  ""optimization_advice"": ""Tip on how to make it cleaner or faster"",
  ""steps"": [
    {{ ""step"": 1, ""sub_expression"": ""Inner function"", ""description"": ""What this part does"" }},
    {{ ""step"": 2, ""sub_expression"": ""Outer function"", ""description"": ""What the outer part does"" }}
  ]
}}";

                    var json = await QueryAiJsonAsync(prompt, aiConfig);
                    if (json != null)
                    {
                        result.OverallPurpose = json["overall_purpose"]?.ToString() ?? "Công thức tính toán Excel";
                        result.ReturnTypeInfo = json["return_type"]?.ToString() ?? "Kết quả tự động";
                        result.OptimizationAdvice = json["optimization_advice"]?.ToString() ?? string.Empty;

                        if (json["steps"] is JArray arr)
                        {
                            foreach (var s in arr)
                            {
                                result.Steps.Add(new FormulaStepExplanation
                                {
                                    StepNumber = s["step"]?.ToObject<int>() ?? (result.Steps.Count + 1),
                                    SubExpression = s["sub_expression"]?.ToString() ?? string.Empty,
                                    Description = s["description"]?.ToString() ?? string.Empty
                                });
                            }
                        }
                        return result;
                    }
                }
                catch { }
            }

            // Fallback rule explainer
            result.OverallPurpose = $"Công thức chứa các hàm: {ExtractFunctionNames(formula)}";
            result.ReturnTypeInfo = "Dữ liệu tính toán";
            result.Steps.Add(new FormulaStepExplanation
            {
                StepNumber = 1,
                SubExpression = formula,
                Description = "Được thực thi tuần tự từ trong ra ngoài theo thứ tự ưu tiên của toán tử Excel."
            });

            return result;
        }

        public static async Task<FormulaModernizeResult> ModernizeFormulaAsync(string formula, AiConfig? aiConfig, AppLanguage lang = AppLanguage.Vietnamese)
        {
            var result = new FormulaModernizeResult { OriginalFormula = formula };

            // Quick rule modernization
            string modernized = formula;

            // 1. VLOOKUP -> XLOOKUP
            if (modernized.IndexOf("VLOOKUP", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                // VLOOKUP(A2, D:E, 2, FALSE) -> XLOOKUP(A2, D:D, E:E, "N/A")
                result.ChangesSummary = "Nâng cấp từ VLOOKUP sang XLOOKUP (hiệu năng cao hơn, không bị ảnh hưởng khi chèn/xóa cột).";
            }

            // 2. IF(ISERROR()) -> IFERROR()
            if (modernized.IndexOf("IF(ISERROR(", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                result.ChangesSummary = "Thay thế cấu trúc lồng IF(ISERROR()) rườm rà bằng hàm IFERROR() gọn nhẹ.";
            }

            result.ModernizedFormula = modernized;
            result.IsModernized = modernized != formula;

            if (aiConfig != null && !string.IsNullOrWhiteSpace(aiConfig.BaseUrl))
            {
                try
                {
                    string prompt = $@"Convert this Excel formula into modern Excel 365 best-practice formula (e.g. use XLOOKUP, LET, IFS, IFERROR, FILTER, UNIQUE if applicable):
Original Formula: `{formula}`

Respond in strictly valid JSON:
{{
  ""modernized_formula"": ""=MODERN_FORMULA"",
  ""summary"": ""Summary of enhancements made""
}}";

                    var json = await QueryAiJsonAsync(prompt, aiConfig);
                    if (json != null)
                    {
                        string? m = json["modernized_formula"]?.ToString();
                        string? s = json["summary"]?.ToString();
                        if (!string.IsNullOrWhiteSpace(m) && m.StartsWith("="))
                        {
                            result.ModernizedFormula = m.Trim();
                            result.ChangesSummary = s ?? result.ChangesSummary;
                            result.IsModernized = true;
                        }
                    }
                }
                catch { }
            }

            return result;
        }

        private static string ExtractFunctionNames(string formula)
        {
            var matches = Regex.Matches(formula, @"\b([A-Z]{2,15})\s*\(");
            var names = matches.Cast<Match>().Select(m => m.Groups[1].Value).Distinct().ToList();
            return names.Count > 0 ? string.Join(", ", names) : "Biểu thức số học";
        }

        private static async Task<JObject?> QueryAiJsonAsync(string prompt, AiConfig config)
        {
            string baseUrl = config.BaseUrl.TrimEnd('/');
            if (!baseUrl.EndsWith("/v1", StringComparison.OrdinalIgnoreCase) && !baseUrl.Contains("/v1/"))
            {
                baseUrl += "/v1";
            }
            string endpoint = $"{baseUrl}/chat/completions";
            string model = string.IsNullOrWhiteSpace(config.ModelName) ? "qwen-3.6" : config.ModelName.Trim();

            var payload = new JObject
            {
                ["model"] = model,
                ["messages"] = new JArray
                {
                    new JObject { ["role"] = "system", ["content"] = "You are a specialized AI assistant that always returns answers in strictly valid JSON format." },
                    new JObject { ["role"] = "user", ["content"] = prompt }
                },
                ["temperature"] = 0.2
            };

            using (var cts = new CancellationTokenSource(TimeSpan.FromSeconds(Math.Max(5, config.TimeoutSeconds))))
            using (var client = new System.Net.Http.HttpClient())
            {
                if (!string.IsNullOrWhiteSpace(config.ApiKey))
                {
                    client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", config.ApiKey);
                }

                var content = new System.Net.Http.StringContent(payload.ToString(), Encoding.UTF8, "application/json");
                var response = await client.PostAsync(endpoint, content, cts.Token);
                string responseBody = await response.Content.ReadAsStringAsync();

                if (response.IsSuccessStatusCode)
                {
                    var root = JObject.Parse(responseBody);
                    string? reply = root["choices"]?[0]?["message"]?["content"]?.ToString();
                    if (!string.IsNullOrWhiteSpace(reply))
                    {
                        // Clean markdown code blocks ```json ... ```
                        reply = Regex.Replace(reply, @"^```[a-z]*\s*", "", RegexOptions.Multiline);
                        reply = Regex.Replace(reply, @"\s*```$", "", RegexOptions.Multiline).Trim();

                        int startIdx = reply.IndexOf('{');
                        int endIdx = reply.LastIndexOf('}');
                        if (startIdx >= 0 && endIdx > startIdx)
                        {
                            reply = reply.Substring(startIdx, endIdx - startIdx + 1);
                        }

                        return JObject.Parse(reply);
                    }
                }
            }

            return null;
        }

        #endregion
    }
}
