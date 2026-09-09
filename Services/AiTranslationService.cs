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
                        text = resultList[itemIdx].OriginalText
                    });
                }

                string inputJson = JsonConvert.SerializeObject(batchPayload, Formatting.None);

                string systemPrompt = BuildSystemPrompt(direction, tone, glossarySection);
                string userPrompt = $"Dịch mảng các đoạn văn bản trong Excel sau đây theo đúng yêu cầu:\n{inputJson}";

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
                            resultList[itemIdx].TranslatedText = transText.Trim();
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

            // Glossary
            if (!string.IsNullOrWhiteSpace(glossarySection))
            {
                sb.AppendLine();
                sb.AppendLine(glossarySection);
            }

            // Format constraints
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
    }
}
