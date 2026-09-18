using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using ExcelSupport.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ExcelSupport.Services
{
    public enum TranslationDirection
    {
        JapaneseToVietnamese, // Nhật ➔ Việt
        VietnameseToJapanese, // Việt ➔ Nhật
        EnglishToVietnamese,  // Anh ➔ Việt
        VietnameseToEnglish,  // Việt ➔ Anh
        JapaneseToEnglish,   // Nhật ➔ Anh
        AutoDetect            // Tự động nhận diện
    }

    public enum TranslationTone
    {
        ITSoftware, // Chuyên ngành IT & Phần mềm
        Business,   // Doanh nghiệp / Công sở
        Casual      // Tự nhiên / Đời thường
    }

    public static class AiTranslationService
    {
        private const int DefaultBatchChunkSize = 40;

        public static string GetDirectionDisplayName(TranslationDirection dir)
        {
            switch (dir)
            {
                case TranslationDirection.JapaneseToVietnamese:
                    return "🇯🇵 Tiếng Nhật ➔ 🇻🇳 Tiếng Việt";
                case TranslationDirection.VietnameseToJapanese:
                    return "🇻🇳 Tiếng Việt ➔ 🇯🇵 Tiếng Nhật";
                case TranslationDirection.EnglishToVietnamese:
                    return "🇺🇸 Tiếng Anh ➔ 🇻🇳 Tiếng Việt";
                case TranslationDirection.VietnameseToEnglish:
                    return "🇻🇳 Tiếng Việt ➔ 🇺🇸 Tiếng Anh";
                case TranslationDirection.JapaneseToEnglish:
                    return "🇯🇵 Tiếng Nhật ➔ 🇺🇸 Tiếng Anh";
                case TranslationDirection.AutoDetect:
                default:
                    return "🌐 Tự Động Nhận Diện (Auto-Detect)";
            }
        }

        public static string GetToneDisplayName(TranslationTone tone)
        {
            switch (tone)
            {
                case TranslationTone.ITSoftware:
                    return "💻 IT & Phát Triển Phần Mềm (Hệ thống, CSDL, Thiết kế)";
                case TranslationTone.Business:
                    return "🏢 Công Sở & Doanh Nghiệp (Lịch sự, Kính ngữ/Keigo)";
                case TranslationTone.Casual:
                    return "💬 Giao Tiếp & Tự Nhiên (Đời thường, Dễ hiểu)";
                default:
                    return "💻 IT & Phát Triển Phần Mềm";
            }
        }

        public static async Task<List<CellTextItem>> TranslateBatchAsync(
            List<CellTextItem> items,
            TranslationDirection direction,
            TranslationTone tone,
            bool enableGlossary,
            IEnumerable<GlossaryItem>? customGlossary,
            IProgress<(int processed, int total, string status)>? progress = null,
            CancellationToken cancellationToken = default)
        {
            if (items == null || items.Count == 0) return new List<CellTextItem>();

            var config = AiConfigManager.Current;
            var resultList = items.Select(x => x.Clone()).ToList();

            // Filter out empty items for translation, but keep indexes
            var nonEmptyIndices = new List<int>();
            for (int i = 0; i < resultList.Count; i++)
            {
                if (!string.IsNullOrWhiteSpace(resultList[i].OriginalText))
                {
                    nonEmptyIndices.Add(i);
                }
                else
                {
                    resultList[i].TranslatedText = string.Empty;
                }
            }

            if (nonEmptyIndices.Count == 0) return resultList;

            // Prepare glossary instructions
            string glossarySection = BuildGlossaryInstruction(enableGlossary, customGlossary ?? config.Glossary, direction);

            // Chunk processing
            int totalToProcess = nonEmptyIndices.Count;
            int processedCount = 0;

            for (int chunkStart = 0; chunkStart < totalToProcess; chunkStart += DefaultBatchChunkSize)
            {
                if (cancellationToken.IsCancellationRequested)
                {
                    break;
                }

                int currentChunkLength = Math.Min(DefaultBatchChunkSize, totalToProcess - chunkStart);
                var currentBatchIndices = nonEmptyIndices.Skip(chunkStart).Take(currentChunkLength).ToList();

                var batchPayload = new List<object>();
                for (int bIdx = 0; bIdx < currentBatchIndices.Count; bIdx++)
                {
                    int itemIdx = currentBatchIndices[bIdx];
                    batchPayload.Add(new
                    {
                        id = bIdx,
                        text = resultList[itemIdx].GetTaggedOriginalText()
                    });
                }

                string inputJson = JsonConvert.SerializeObject(batchPayload, Formatting.None);

                // Match glossary terms specifically for this batch chunk
                var matchedTerms = new List<GlossaryItem>();
                var allTerms = (customGlossary ?? config.Glossary)?.Where(g => !string.IsNullOrWhiteSpace(g.Japanese) && !string.IsNullOrWhiteSpace(g.Vietnamese)).ToList();
                if (enableGlossary && allTerms != null && allTerms.Count > 0)
                {
                    var chunkTexts = currentBatchIndices.Select(idx => resultList[idx].OriginalText ?? "").ToList();
                    string combinedChunkText = string.Join("\n", chunkTexts);
                    foreach (var term in allTerms)
                    {
                        if (combinedChunkText.IndexOf(term.Japanese, StringComparison.OrdinalIgnoreCase) >= 0 ||
                            combinedChunkText.IndexOf(term.Vietnamese, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            matchedTerms.Add(term);
                        }
                    }
                }

                var sbUser = new StringBuilder();
                if (matchedTerms.Count > 0)
                {
                    sbUser.AppendLine("[BẢNG THUẬT NGỮ GLOSSARY CẦN ÁP DỤNG CHÍNH XÁC CHO CÁC CÂU NÀY - ƯU TIÊN CAO NHẤT]:");
                    foreach (var m in matchedTerms)
                    {
                        string note = string.IsNullOrWhiteSpace(m.Note) ? "" : $" (Ghi chú: {m.Note})";
                        if (direction == TranslationDirection.JapaneseToVietnamese)
                        {
                            sbUser.AppendLine($"• \"{m.Japanese}\" ➔ BẮT BUỘC DỊCH LÀ: \"{m.Vietnamese}\"{note}");
                        }
                        else if (direction == TranslationDirection.VietnameseToJapanese)
                        {
                            sbUser.AppendLine($"• \"{m.Vietnamese}\" ➔ BẮT BUỘC DỊCH LÀ: \"{m.Japanese}\"{note}");
                        }
                        else
                        {
                            sbUser.AppendLine($"• \"{m.Japanese}\" (Tiếng Nhật) ⇋ \"{m.Vietnamese}\" (Tiếng Việt){note}");
                        }
                    }
                    sbUser.AppendLine("BẮT BUỘC: Khi xuất hiện các từ trên, bạn PHẢI dùng chính xác từ dịch tương ứng trong danh sách, tuyệt đối không dịch khác!");
                    sbUser.AppendLine();
                }

                sbUser.AppendLine("Dịch mảng các đoạn văn bản trong Excel sau đây theo đúng yêu cầu:");
                sbUser.AppendLine(inputJson);

                string systemPrompt = BuildSystemPrompt(direction, tone, glossarySection);
                string userPrompt = sbUser.ToString();

                progress?.Report((processedCount, totalToProcess, $"Đang dịch {processedCount + 1}-{processedCount + currentBatchIndices.Count} / {totalToProcess} ô..."));

                try
                {
                    string aiResponse = await Task.Run(() => OpenAiClientService.SendChatAsync(config, userPrompt, systemPrompt), cancellationToken);

                    var map = ParseTranslationResponse(aiResponse);

                    for (int bIdx = 0; bIdx < currentBatchIndices.Count; bIdx++)
                    {
                        int itemIdx = currentBatchIndices[bIdx];
                        if (map.TryGetValue(bIdx, out string? transText) && !string.IsNullOrWhiteSpace(transText))
                        {
                            var (plainText, runs) = ParseTaggedText(transText.Trim());
                            resultList[itemIdx].TranslatedText = plainText;
                            resultList[itemIdx].TranslatedRuns = runs;
                        }
                        else
                        {
                            // Fallback to original text if missing in response
                            resultList[itemIdx].TranslatedText = resultList[itemIdx].OriginalText;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[AiTranslationService] Chunk error: {ex.Message}");
                    for (int bIdx = 0; bIdx < currentBatchIndices.Count; bIdx++)
                    {
                        int itemIdx = currentBatchIndices[bIdx];
                        if (string.IsNullOrEmpty(resultList[itemIdx].TranslatedText))
                        {
                            resultList[itemIdx].TranslatedText = $"[Lỗi dịch: {ex.Message}]";
                        }
                    }
                }

                processedCount += currentBatchIndices.Count;
                progress?.Report((processedCount, totalToProcess, $"Đã dịch xong {processedCount} / {totalToProcess} ô."));
            }

            return resultList;
        }

        private static string BuildSystemPrompt(TranslationDirection direction, TranslationTone tone, string glossarySection)
        {
            var sb = new StringBuilder();
            sb.AppendLine("Bạn là một chuyên gia biên dịch cao cấp chuyên xử lý tài liệu bảng tính Excel, đặc tả thiết kế hệ thống, cơ sở dữ liệu và giao tiếp doanh nghiệp.");

            // Language direction
            switch (direction)
            {
                case TranslationDirection.JapaneseToVietnamese:
                    sb.AppendLine("Nhiệm vụ: Dịch chính xác từ TIẾNG NHẬT sang TIẾNG VIỆT.");
                    break;
                case TranslationDirection.VietnameseToJapanese:
                    sb.AppendLine("Nhiệm vụ: Dịch chính xác từ TIẾNG VIỆT sang TIẾNG NHẬT.");
                    break;
                case TranslationDirection.EnglishToVietnamese:
                    sb.AppendLine("Nhiệm vụ: Dịch chính xác từ TIẾNG ANH sang TIẾNG VIỆT.");
                    break;
                case TranslationDirection.VietnameseToEnglish:
                    sb.AppendLine("Nhiệm vụ: Dịch chính xác từ TIẾNG VIỆT sang TIẾNG ANH.");
                    break;
                case TranslationDirection.JapaneseToEnglish:
                    sb.AppendLine("Nhiệm vụ: Dịch chính xác từ TIẾNG NHẬT sang TIẾNG ANH.");
                    break;
                case TranslationDirection.AutoDetect:
                default:
                    sb.AppendLine("Nhiệm vụ: Tự động nhận diện ngôn ngữ nguồn từng ô. Nếu nguồn là tiếng Nhật hoặc tiếng Anh thì dịch sang tiếng Việt; nếu nguồn là tiếng Việt thì dịch sang tiếng Nhật chuẩn.");
                    break;
            }

            // Tone
            switch (tone)
            {
                case TranslationTone.ITSoftware:
                    sb.AppendLine("Phong cách văn phong: CHUYÊN NGÀNH CNTT / IT SOFTWARE. Giữ nguyên các từ khóa lập trình, tên bảng CSDL, tên cột, API, mã lỗi, tên biến (như user_id, status, SELECT, WHERE). Sử dụng thuật ngữ IT chuẩn.");
                    break;
                case TranslationTone.Business:
                    sb.AppendLine("Phong cách văn phong: CÔNG SỞ & DOANH NGHIỆP. Sử dụng ngôn từ trang trọng, lịch sự, kính ngữ (Keigo) phù hợp trong báo cáo và trao đổi công việc khách hàng.");
                    break;
                case TranslationTone.Casual:
                    sb.AppendLine("Phong cách văn phong: TỰ NHIÊN / GIAO TIẾP. Dịch thoáng ý, dễ hiểu, giữ đúng sắc thái tự nhiên.");
                    break;
            }

            // General requirements
            sb.AppendLine("Quy tắc xử lý văn bản:");
            sb.AppendLine("- Giữ nguyên các ký tự xuống dòng (\\n), khoảng trắng đặc biệt, số thứ tự (1., 2., •, -), và mã định danh.");
            sb.AppendLine("- Không tự ý thêm bớt nội dung hay chèn ý kiến cá nhân.");
            sb.AppendLine("- Đảm bảo câu văn dịch mượt mà, đúng ngữ pháp của ngôn ngữ đích.");
            sb.AppendLine();
            sb.AppendLine("[QUY TẮC BẢO TOÀN ĐỊNH DẠNG ĐẶC BIỆT - BẮT BUỘC]:");
            sb.AppendLine("- Nếu văn bản nguồn chứa các thẻ định dạng HTML/XML như `<s color=\"...\">...</s>` hoặc `<s>...</s>` (gạch ngang / sửa đổi), `<color hex=\"...\">...</color>` (chữ màu đỏ/màu khác), `<b>...</b>` (in đậm), `<i>...</i>` (in nghiêng), bạn BẮT BUỘC PHẢI GIỮ NGUYÊN các thẻ này và bao bọc đúng phần nội dung dịch tương ứng.");
            sb.AppendLine("- Ví dụ nguồn: `・<s color=\"#FF0000\">画面.実行時パスによりCentral側のファイル出力を行う。</s>` ➔ Bản dịch: `・<s color=\"#FF0000\">Thực hiện output file phía Central theo đường dẫn thời gian chạy màn hình.</s>`");
            sb.AppendLine("- Ví dụ nguồn: `・<color hex=\"#FF0000\">Local → Centralでファイル出力パスのチェックを行う。</color>` ➔ Bản dịch: `・<color hex=\"#FF0000\">Thực hiện kiểm tra đường dẫn output file từ Local → Central.</color>`");
            sb.AppendLine("- Tuyệt đối không xóa thẻ, không làm mất thuộc tính màu `color=\"...\"`, `hex=\"...\"`, không dịch tên thẻ.");
            sb.AppendLine();
            sb.AppendLine("ĐỊNH DẠNG ĐẦU RA BẮT BUỘC:");
            sb.AppendLine("Trả về DUY NHẤT một mảng JSON hợp lệ, không kèm giải thích, không bọc trong markdown hay bất kỳ văn bản nào ngoài JSON:");
            sb.AppendLine("[");
            sb.AppendLine("  { \"id\": 0, \"trans\": \"Nội dung đã dịch\" },");
            sb.AppendLine("  { \"id\": 1, \"trans\": \"Nội dung đã dịch\" }");
            sb.AppendLine("]");

            return sb.ToString();
        }

        private static string BuildGlossaryInstruction(bool enableGlossary, IEnumerable<GlossaryItem>? glossary, TranslationDirection direction)
        {
            if (!enableGlossary || glossary == null) return string.Empty;

            var validTerms = glossary.Where(g => !string.IsNullOrWhiteSpace(g.Japanese) && !string.IsNullOrWhiteSpace(g.Vietnamese)).ToList();
            if (validTerms.Count == 0) return string.Empty;

            var sb = new StringBuilder();
            sb.AppendLine("[QUY TẮC BẮT BUỘC VỀ THUẬT NGỮ CHUYÊN NGÀNH - GLOSSARY]:");
            sb.AppendLine("Nếu trong văn bản nguồn xuất hiện các thuật ngữ dưới đây, BẮT BUỘC PHẢI DÙNG CHÍNH XÁC bản dịch được chỉ định:");

            int idx = 1;
            foreach (var term in validTerms)
            {
                string note = string.IsNullOrWhiteSpace(term.Note) ? string.Empty : $" (Ngữ cảnh: {term.Note})";

                if (direction == TranslationDirection.JapaneseToVietnamese)
                {
                    sb.AppendLine($"{idx}. \"{term.Japanese}\" ➔ BẮT BUỘC dịch: \"{term.Vietnamese}\"{note}");
                }
                else if (direction == TranslationDirection.VietnameseToJapanese)
                {
                    sb.AppendLine($"{idx}. \"{term.Vietnamese}\" ➔ BẮT BUỘC dịch: \"{term.Japanese}\"{note}");
                }
                else
                {
                    sb.AppendLine($"{idx}. Tiếng Nhật: \"{term.Japanese}\" ⇋ Tiếng Việt: \"{term.Vietnamese}\"{note}");
                }
                idx++;
            }

            return sb.ToString();
        }

        private static Dictionary<int, string> ParseTranslationResponse(string response)
        {
            var result = new Dictionary<int, string>();
            if (string.IsNullOrWhiteSpace(response)) return result;

            try
            {
                string cleanJson = response.Trim();

                // Strip markdown code block fences if present
                if (cleanJson.StartsWith("```json", StringComparison.OrdinalIgnoreCase))
                {
                    cleanJson = cleanJson.Substring(7);
                }
                else if (cleanJson.StartsWith("```", StringComparison.OrdinalIgnoreCase))
                {
                    cleanJson = cleanJson.Substring(3);
                }

                if (cleanJson.EndsWith("```"))
                {
                    cleanJson = cleanJson.Substring(0, cleanJson.Length - 3);
                }

                cleanJson = cleanJson.Trim();

                // Find opening bracket
                int startIdx = cleanJson.IndexOf('[');
                int endIdx = cleanJson.LastIndexOf(']');
                if (startIdx >= 0 && endIdx > startIdx)
                {
                    cleanJson = cleanJson.Substring(startIdx, endIdx - startIdx + 1);
                }

                var array = JArray.Parse(cleanJson);
                foreach (var token in array)
                {
                    if (token is JObject obj)
                    {
                        int id = obj.Value<int>("id");
                        string trans = obj.Value<string>("trans") ?? obj.Value<string>("text") ?? string.Empty;
                        result[id] = trans;
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[AiTranslationService] JSON Parse Error: {ex.Message}. Raw: {response}");

                // Fallback: Regex extraction of {"id": X, "trans": "..."}
                try
                {
                    var matches = Regex.Matches(response, @"\{\s*""id""\s*:\s*(\d+)\s*,\s*""(?:trans|text)""\s*:\s*""((?:\\.|[^""\\])*)""\s*\}");
                    foreach (Match m in matches)
                    {
                        if (int.TryParse(m.Groups[1].Value, out int id))
                        {
                            string rawTrans = m.Groups[2].Value;
                            string unescaped = Regex.Unescape(rawTrans);
                            result[id] = unescaped;
                        }
                    }
                }
                catch { }
            }

            return result;
        }

        public static (string plainText, List<TextRunModel>? runs) ParseTaggedText(string taggedText)
        {
            if (string.IsNullOrEmpty(taggedText)) return (string.Empty, null);

            if (!taggedText.Contains("<") || !taggedText.Contains(">"))
            {
                return (taggedText, null);
            }

            var runs = new List<TextRunModel>();
            var sbPlain = new StringBuilder();

            int boldDepth = 0;
            int italicDepth = 0;
            int strikeDepth = 0;
            var colorStack = new Stack<string>();

            int i = 0;
            int len = taggedText.Length;
            TextRunModel? currentRun = null;
            bool hasAnySpecial = false;

            while (i < len)
            {
                if (taggedText[i] == '<')
                {
                    int closeIdx = taggedText.IndexOf('>', i);
                    if (closeIdx > i)
                    {
                        string tagContent = taggedText.Substring(i + 1, closeIdx - i - 1).Trim();
                        if (TryProcessTag(tagContent, ref boldDepth, ref italicDepth, ref strikeDepth, colorStack, out bool isTag))
                        {
                            if (isTag)
                            {
                                i = closeIdx + 1;
                                continue;
                            }
                        }
                    }
                }

                char c = taggedText[i];
                sbPlain.Append(c);

                bool isB = boldDepth > 0;
                bool isI = italicDepth > 0;
                bool isS = strikeDepth > 0;
                string? curColor = null;
                if (colorStack.Count > 0)
                {
                    string topColor = colorStack.Peek();
                    if (!string.IsNullOrEmpty(topColor))
                    {
                        curColor = topColor.StartsWith("s:") ? topColor.Substring(2) : topColor;
                    }
                }

                if (isB || isI || isS || curColor != null)
                {
                    hasAnySpecial = true;
                }

                if (currentRun != null &&
                    currentRun.IsBold == isB &&
                    currentRun.IsItalic == isI &&
                    currentRun.IsStrikethrough == isS &&
                    string.Equals(currentRun.ColorHex, curColor, StringComparison.OrdinalIgnoreCase))
                {
                    currentRun.Text += c;
                }
                else
                {
                    currentRun = new TextRunModel
                    {
                        Text = c.ToString(),
                        IsBold = isB,
                        IsItalic = isI,
                        IsStrikethrough = isS,
                        ColorHex = curColor
                    };
                    runs.Add(currentRun);
                }

                i++;
            }

            if (!hasAnySpecial || runs.Count == 0)
            {
                return (sbPlain.ToString(), null);
            }

            return (sbPlain.ToString(), runs);
        }

        private static bool TryProcessTag(
            string tagContent,
            ref int boldDepth,
            ref int italicDepth,
            ref int strikeDepth,
            Stack<string> colorStack,
            out bool isTag)
        {
            isTag = false;
            if (string.IsNullOrWhiteSpace(tagContent)) return false;

            bool isClosing = tagContent.StartsWith("/");
            string trimmed = isClosing ? tagContent.Substring(1).Trim() : tagContent;

            string tagName = trimmed;
            int spaceIdx = trimmed.IndexOfAny(new[] { ' ', '\t', '\r', '\n' });
            if (spaceIdx > 0)
            {
                tagName = trimmed.Substring(0, spaceIdx);
            }
            tagName = tagName.ToLowerInvariant();

            if (tagName == "b" || tagName == "strong")
            {
                isTag = true;
                if (isClosing) boldDepth = Math.Max(0, boldDepth - 1);
                else boldDepth++;
                return true;
            }
            if (tagName == "i" || tagName == "em")
            {
                isTag = true;
                if (isClosing) italicDepth = Math.Max(0, italicDepth - 1);
                else italicDepth++;
                return true;
            }
            if (tagName == "s" || tagName == "del" || tagName == "strike")
            {
                isTag = true;
                if (isClosing)
                {
                    strikeDepth = Math.Max(0, strikeDepth - 1);
                    if (colorStack.Count > 0 && colorStack.Peek().StartsWith("s:"))
                    {
                        colorStack.Pop();
                    }
                }
                else
                {
                    strikeDepth++;
                    string? color = ExtractColorAttribute(trimmed);
                    if (color != null)
                    {
                        colorStack.Push("s:" + color);
                    }
                }
                return true;
            }
            if (tagName == "color" || tagName == "font" || tagName == "span")
            {
                isTag = true;
                if (isClosing)
                {
                    if (colorStack.Count > 0) colorStack.Pop();
                }
                else
                {
                    string? color = ExtractColorAttribute(trimmed);
                    if (color != null)
                    {
                        colorStack.Push(color);
                    }
                    else
                    {
                        colorStack.Push(colorStack.Count > 0 ? colorStack.Peek() : "");
                    }
                }
                return true;
            }
            if (tagName == "red")
            {
                isTag = true;
                if (isClosing) { if (colorStack.Count > 0) colorStack.Pop(); }
                else colorStack.Push("#FF0000");
                return true;
            }
            if (tagName == "blue")
            {
                isTag = true;
                if (isClosing) { if (colorStack.Count > 0) colorStack.Pop(); }
                else colorStack.Push("#0000FF");
                return true;
            }
            if (tagName == "green")
            {
                isTag = true;
                if (isClosing) { if (colorStack.Count > 0) colorStack.Pop(); }
                else colorStack.Push("#008000");
                return true;
            }

            return false;
        }

        private static string? ExtractColorAttribute(string tag)
        {
            var m = Regex.Match(tag, @"(?:color|hex)\s*=\s*[""']?([^""'\s>]+)[""']?", RegexOptions.IgnoreCase);
            if (m.Success)
            {
                return NormalizeColorHex(m.Groups[1].Value);
            }
            return null;
        }

        private static string NormalizeColorHex(string color)
        {
            if (color.StartsWith("#")) return color;
            if (color.Equals("red", StringComparison.OrdinalIgnoreCase)) return "#FF0000";
            if (color.Equals("blue", StringComparison.OrdinalIgnoreCase)) return "#0000FF";
            if (color.Equals("green", StringComparison.OrdinalIgnoreCase)) return "#008000";
            return color.StartsWith("#") ? color : ("#" + color);
        }
    }
}
