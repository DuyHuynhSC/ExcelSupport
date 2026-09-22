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

        public static TranslationDirection DetectDirection(string? text)
        {
            if (string.IsNullOrWhiteSpace(text)) return TranslationDirection.JapaneseToVietnamese;

            int japaneseCharScore = 0;
            int vietnameseCharScore = 0;

            const string vietnameseChars = "àáảãạăằắẳẵặâầấẩẫậèéẻẽẹêềếểễệìíỉĩịòóỏõọôồốổỗộơờớởỡợùúủũụưừứửữựỳýỷỹỵđĐÀÁẢÃẠĂẰẮẲẴẶÂẦẤẨẪẬÈÉẺẼẸÊỀẾỂỄỆÌÍỈĨỊÒÓỎÕỌÔỒỐỔỖỘƠỜỚỞỠỢÙÚỦŨỤƯỪỨỬỮỰỲÝỶỸỴ";

            foreach (char c in text!)
            {
                // Hiragana: 0x3040 - 0x309F, Katakana: 0x30A0 - 0x30FF, Halfwidth Katakana: 0xFF65 - 0xFF9F
                if ((c >= 0x3040 && c <= 0x309F) || (c >= 0x30A0 && c <= 0x30FF) || (c >= 0xFF65 && c <= 0xFF9F))
                {
                    japaneseCharScore += 3;
                }
                // CJK Unified Ideographs (Kanji)
                else if (c >= 0x4E00 && c <= 0x9FFF)
                {
                    japaneseCharScore += 2;
                }
                else if (vietnameseChars.IndexOf(c) >= 0)
                {
                    vietnameseCharScore += 3;
                }
            }

            if (vietnameseCharScore > 0 && vietnameseCharScore >= japaneseCharScore)
            {
                return TranslationDirection.VietnameseToJapanese;
            }

            if (japaneseCharScore > 0)
            {
                return TranslationDirection.JapaneseToVietnamese;
            }

            return TranslationDirection.JapaneseToVietnamese;
        }

        public static TranslationDirection DetectDirection(IEnumerable<string>? texts)
        {
            if (texts == null) return TranslationDirection.JapaneseToVietnamese;

            int vnCount = 0;
            int jpCount = 0;

            foreach (var t in texts.Take(50))
            {
                if (string.IsNullOrWhiteSpace(t)) continue;
                var dir = DetectDirection(t);
                if (dir == TranslationDirection.VietnameseToJapanese) vnCount++;
                else if (dir == TranslationDirection.JapaneseToVietnamese) jpCount++;
            }

            return vnCount > jpCount ? TranslationDirection.VietnameseToJapanese : TranslationDirection.JapaneseToVietnamese;
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

                if (direction == TranslationDirection.VietnameseToJapanese)
                {
                    sbUser.AppendLine("Dịch mảng văn bản TIẾNG VIỆT sau đây SANG TIẾNG NHẬT (日本語). Giá trị trường \"trans\" trong JSON BẮT BUỘC phải là tiếng Nhật:");
                }
                else if (direction == TranslationDirection.JapaneseToVietnamese)
                {
                    sbUser.AppendLine("Dịch mảng văn bản TIẾNG NHẬT sau đây SANG TIẾNG VIỆT. Giá trị trường \"trans\" trong JSON BẮT BUỘC phải là tiếng Việt:");
                }
                else if (direction == TranslationDirection.VietnameseToEnglish)
                {
                    sbUser.AppendLine("Dịch mảng văn bản TIẾNG VIỆT sau đây SANG TIẾNG ANH (English). Giá trị trường \"trans\" trong JSON BẮT BUỘC phải là tiếng Anh:");
                }
                else if (direction == TranslationDirection.AutoDetect)
                {
                    sbUser.AppendLine("Nhận diện ngôn ngữ từng ô và dịch: Nếu nguồn là tiếng Việt -> dịch sang tiếng Nhật (日本語); Nếu nguồn là tiếng Nhật hoặc tiếng Anh -> dịch sang tiếng Việt:");
                }
                else
                {
                    sbUser.AppendLine("Dịch mảng các đoạn văn bản trong Excel sau đây theo đúng yêu cầu:");
                }
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
                            if (map.Count == 0 && !string.IsNullOrWhiteSpace(aiResponse))
                            {
                                resultList[itemIdx].TranslatedText = "[Lỗi đọc kết quả phản hồi từ AI]";
                            }
                            else
                            {
                                resultList[itemIdx].TranslatedText = resultList[itemIdx].OriginalText;
                            }
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
                    sb.AppendLine("BẮT BUỘC: Bản dịch ở trường \"trans\" PHẢI LÀ TIẾNG VIỆT. Tuyệt đối không để nguyên tiếng Nhật.");
                    break;
                case TranslationDirection.VietnameseToJapanese:
                    sb.AppendLine("Nhiệm vụ: Dịch chính xác từ TIẾNG VIỆT sang TIẾNG NHẬT (日本語).");
                    sb.AppendLine("BẮT BUỘC: Bản dịch ở trường \"trans\" PHẢI LÀ TIẾNG NHẬT (日本語) CHUẨN XÁC, TỰ NHIÊN. TUYỆT ĐỐI KHÔNG ĐƯỢC GIỮ NGUYÊN TIẾNG VIỆT, KHÔNG TRẢ VỀ TIẾNG VIỆT TRONG KẾT QUẢ \"trans\"!");
                    break;
                case TranslationDirection.EnglishToVietnamese:
                    sb.AppendLine("Nhiệm vụ: Dịch chính xác từ TIẾNG ANH sang TIẾNG VIỆT.");
                    sb.AppendLine("BẮT BUỘC: Bản dịch ở trường \"trans\" PHẢI LÀ TIẾNG VIỆT.");
                    break;
                case TranslationDirection.VietnameseToEnglish:
                    sb.AppendLine("Nhiệm vụ: Dịch chính xác từ TIẾNG VIỆT sang TIẾNG ANH.");
                    sb.AppendLine("BẮT BUỘC: Bản dịch ở trường \"trans\" PHẢI LÀ TIẾNG ANH.");
                    break;
                case TranslationDirection.JapaneseToEnglish:
                    sb.AppendLine("Nhiệm vụ: Dịch chính xác từ TIẾNG NHẬT sang TIẾNG ANH.");
                    sb.AppendLine("BẮT BUỘC: Bản dịch ở trường \"trans\" PHẢI LÀ TIẾNG ANH.");
                    break;
                case TranslationDirection.AutoDetect:
                default:
                    sb.AppendLine("Nhiệm vụ: Tự động nhận diện ngôn ngữ nguồn từng ô:");
                    sb.AppendLine("- Nếu văn bản nguồn là TIẾNG VIỆT: BẮT BUỘC DỊCH SANG TIẾNG NHẬT (日本語). Trường \"trans\" PHẢI LÀ TIẾNG NHẬT!");
                    sb.AppendLine("- Nếu văn bản nguồn là TIẾNG NHẬT hoặc TIẾNG ANH: BẮT BUỘC DỊCH SANG TIẾNG VIỆT. Trường \"trans\" PHẢI LÀ TIẾNG VIỆT!");
                    sb.AppendLine("- Tuyệt đối không trả về cùng một ngôn ngữ với văn bản nguồn.");
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

            if (direction == TranslationDirection.VietnameseToJapanese)
            {
                sb.AppendLine("- Ví dụ nguồn: `・<s color=\"#FF0000\">Thực hiện xuất file phía Central theo đường dẫn thời gian chạy màn hình.</s>` ➔ Bản dịch: `・<s color=\"#FF0000\">画面.実行時パスによりCentral側のファイル出力を行う。</s>`");
                sb.AppendLine("- Ví dụ nguồn: `・<color hex=\"#FF0000\">Local → Central: Thực hiện kiểm tra đường dẫn xuất file.</color>` ➔ Bản dịch: `・<color hex=\"#FF0000\">Local → Centralでファイル出力パスのチェックを行う。</color>`");
            }
            else if (direction == TranslationDirection.VietnameseToEnglish || direction == TranslationDirection.JapaneseToEnglish)
            {
                sb.AppendLine("- Ví dụ nguồn: `・<s color=\"#FF0000\">Kiểm tra đường dẫn file.</s>` ➔ Bản dịch: `・<s color=\"#FF0000\">Check the file path.</s>`");
            }
            else if (direction == TranslationDirection.AutoDetect)
            {
                sb.AppendLine("- Ví dụ nguồn tiếng Nhật: `・<s color=\"#FF0000\">画面.実行時パスによりCentral側のファイル出力を行う。</s>` ➔ Bản dịch: `・<s color=\"#FF0000\">Thực hiện xuất file phía Central theo đường dẫn runtime màn hình.</s>`");
                sb.AppendLine("- Ví dụ nguồn tiếng Việt: `・<color hex=\"#FF0000\">Thực hiện kiểm tra đường dẫn xuất file.</s>` ➔ Bản dịch: `・<color hex=\"#FF0000\">ファイル出力パスのチェックを行う。</s>`");
            }
            else
            {
                sb.AppendLine("- Ví dụ nguồn: `・<s color=\"#FF0000\">画面.実行時パスによりCentral側のファイル出力を行う。</s>` ➔ Bản dịch: `・<s color=\"#FF0000\">Thực hiện output file phía Central theo đường dẫn thời gian chạy màn hình.</s>`");
                sb.AppendLine("- Ví dụ nguồn: `・<color hex=\"#FF0000\">Local → Centralでファイル出力パスのチェックを行う。</color>` ➔ Bản dịch: `・<color hex=\"#FF0000\">Thực hiện kiểm tra đường dẫn output file từ Local → Central.</color>`");
            }

            sb.AppendLine("- Tuyệt đối không xóa thẻ, không làm mất thuộc tính màu `color=\"...\"`, `hex=\"...\"`, không dịch tên thẻ.");
            sb.AppendLine();
            sb.AppendLine("ĐỊNH DẠNG ĐẦU RA BẮT BUỘC:");
            sb.AppendLine("Trả về DUY NHẤT một mảng JSON hợp lệ, không kèm giải thích, không bọc trong markdown hay bất kỳ văn bản nào ngoài JSON:");
            sb.AppendLine("[");
            if (direction == TranslationDirection.VietnameseToJapanese)
            {
                sb.AppendLine("  { \"id\": 0, \"trans\": \"翻訳された日本語テキスト\" },");
                sb.AppendLine("  { \"id\": 1, \"trans\": \"翻訳された日本語テキスト\" }");
            }
            else if (direction == TranslationDirection.VietnameseToEnglish || direction == TranslationDirection.JapaneseToEnglish)
            {
                sb.AppendLine("  { \"id\": 0, \"trans\": \"Translated English text\" },");
                sb.AppendLine("  { \"id\": 1, \"trans\": \"Translated English text\" }");
            }
            else
            {
                sb.AppendLine("  { \"id\": 0, \"trans\": \"Nội dung đã dịch\" },");
                sb.AppendLine("  { \"id\": 1, \"trans\": \"Nội dung đã dịch\" }");
            }
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

            string cleanJson = response.Trim();

            // Extract content from markdown code block if present
            var matchCodeFence = Regex.Match(cleanJson, @"```(?:json)?\s*([\s\S]*?)\s*```", RegexOptions.IgnoreCase);
            if (matchCodeFence.Success)
            {
                cleanJson = matchCodeFence.Groups[1].Value.Trim();
            }

            // Attempt JSON parsing using JToken
            try
            {
                // Find outer bracket or brace
                int arrStart = cleanJson.IndexOf('[');
                int arrEnd = cleanJson.LastIndexOf(']');
                int objStart = cleanJson.IndexOf('{');
                int objEnd = cleanJson.LastIndexOf('}');

                string jsonToParse = cleanJson;
                if (arrStart >= 0 && arrEnd > arrStart && (objStart < 0 || arrStart < objStart))
                {
                    jsonToParse = cleanJson.Substring(arrStart, arrEnd - arrStart + 1);
                }
                else if (objStart >= 0 && objEnd > objStart)
                {
                    jsonToParse = cleanJson.Substring(objStart, objEnd - objStart + 1);
                }

                var token = JToken.Parse(jsonToParse);
                ExtractTokensToMap(token, result);

                if (result.Count > 0) return result;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[AiTranslationService] JToken Parse warning: {ex.Message}. Raw: {response}");
            }

            // Fallback 1: Regex extraction {"id": X, "trans": "..."} or {"trans": "...", "id": X}
            try
            {
                // Matches {"id": 0, "trans": "..."}
                var matches1 = Regex.Matches(response, @"\{\s*""id""\s*:\s*(\d+)\s*,\s*""(?:trans|translation|translated|text|content|result|target|ja|vi)""\s*:\s*""((?:\\.|[^""\\])*)""");
                foreach (Match m in matches1)
                {
                    if (int.TryParse(m.Groups[1].Value, out int id) && !result.ContainsKey(id))
                    {
                        result[id] = Regex.Unescape(m.Groups[2].Value);
                    }
                }

                // Matches {"trans": "...", "id": 0}
                var matches2 = Regex.Matches(response, @"""(?:trans|translation|translated|text|content|result|target|ja|vi)""\s*:\s*""((?:\\.|[^""\\])*)""\s*,\s*""id""\s*:\s*(\d+)");
                foreach (Match m in matches2)
                {
                    if (int.TryParse(m.Groups[2].Value, out int id) && !result.ContainsKey(id))
                    {
                        result[id] = Regex.Unescape(m.Groups[1].Value);
                    }
                }
            }
            catch { }

            // Fallback 2: Single item without id or simple text
            if (result.Count == 0)
            {
                try
                {
                    var singleMatch = Regex.Match(response, @"""(?:trans|translation|translated|ja|vi)""\s*:\s*""((?:\\.|[^""\\])*)""");
                    if (singleMatch.Success)
                    {
                        result[0] = Regex.Unescape(singleMatch.Groups[1].Value);
                    }
                }
                catch { }
            }

            return result;
        }

        private static void ExtractTokensToMap(JToken token, Dictionary<int, string> result)
        {
            if (token is JArray arr)
            {
                for (int i = 0; i < arr.Count; i++)
                {
                    var item = arr[i];
                    if (item is JObject obj)
                    {
                        int id = obj.Value<int?>("id") ?? i;
                        string? trans = ExtractTranslationStringFromObject(obj);
                        if (!string.IsNullOrEmpty(trans))
                        {
                            result[id] = trans!;
                        }
                    }
                    else if (item is JValue val && val.Value != null)
                    {
                        result[i] = val.Value.ToString() ?? string.Empty;
                    }
                }
            }
            else if (token is JObject obj)
            {
                // Check if object wraps a list under 'translations', 'data', 'items', 'results'
                foreach (var propName in new[] { "translations", "items", "results", "data", "list" })
                {
                    if (obj.TryGetValue(propName, StringComparison.OrdinalIgnoreCase, out var innerToken) && innerToken is JArray innerArr)
                    {
                        ExtractTokensToMap(innerArr, result);
                        return;
                    }
                }

                // Single object
                int id = obj.Value<int?>("id") ?? 0;
                string? trans = ExtractTranslationStringFromObject(obj);
                if (!string.IsNullOrEmpty(trans))
                {
                    result[id] = trans!;
                }
            }
        }

        private static string? ExtractTranslationStringFromObject(JObject obj)
        {
            var candidateKeys = new[] { "trans", "translation", "translated", "target", "text", "content", "result", "ja", "vi", "en" };
            foreach (var key in candidateKeys)
            {
                if (obj.TryGetValue(key, StringComparison.OrdinalIgnoreCase, out var val))
                {
                    string s = val.ToString();
                    if (!string.IsNullOrWhiteSpace(s)) return s;
                }
            }

            // Fallback: Pick any non-id property value
            var otherProp = obj.Properties().FirstOrDefault(p => !string.Equals(p.Name, "id", StringComparison.OrdinalIgnoreCase));
            return otherProp?.Value?.ToString();
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
