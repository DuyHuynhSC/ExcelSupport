using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace ExcelSupport.Services
{
    public static class VietnameseToKatakanaConverter
    {
        // Bảng ánh xạ các âm tiết / tên tiếng Việt thông dụng sang Katakana
        private static readonly Dictionary<string, string> SyllableMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            // Họ phổ biến
            { "nguyen", "グエン" },
            { "tran", "チャン" },
            { "le", "レ" },
            { "pham", "ファム" },
            { "hoang", "ホアン" },
            { "huynh", "フイン" },
            { "phan", "ファン" },
            { "vu", "ヴー" },
            { "vo", "ヴォ" },
            { "dang", "ダン" },
            { "bui", "ブイ" },
            { "do", "ドー" },
            { "ho", "ホー" },
            { "ngo", "ゴー" },
            { "duong", "ズオン" },
            { "ly", "リー" },
            { "dinh", "ディン" },
            { "doan", "ドアン" },
            { "truong", "チュオン" },
            { "luong", "ルオン" },
            { "tong", "トン" },
            { "trinh", "チン" },
            { "dao", "ダオ" },
            { "ha", "ハー" },
            { "mai", "マイ" },
            { "cao", "カオ" },
            { "ta", "ター" },
            { "thai", "タイ" },
            { "chu", "チュー" },
            { "luu", "ルー" },

            // Tên đệm & Tên chính phổ biến
            { "van", "ヴァン" },
            { "thi", "ティ" },
            { "anh", "アイン" },
            { "an", "アン" },
            { "bac", "バック" },
            { "bach", "バック" },
            { "bao", "バオ" },
            { "bich", "ビック" },
            { "binh", "ビン" },
            { "cam", "カム" },
            { "canh", "カイン" },
            { "chau", "チャウ" },
            { "chi", "チー" },
            { "chien", "チエン" },
            { "chinh", "チン" },
            { "chung", "チュン" },
            { "cong", "コン" },
            { "cuc", "クック" },
            { "cuong", "クオン" },
            { "dai", "ダイ" },
            { "dan", "ダン" },
            { "dat", "ダット" },
            { "dien", "ディエン" },
            { "diep", "ディエップ" },
            { "diep ", "ディエップ" },
            { "dieu", "ディエウ" },
            { "doanh", "ドアン" },
            { "dong", "ドン" },
            { "duc", "ドゥック" },
            { "dung", "ズン" },
            { "duy", "ズイ" },
            { "duyen", "ズエン" },
            { "giang", "ザン" },
            { "giao", "ザオ" },
            { "hai", "ハイ" },
            { "han", "ハン" },
            { "hanh", "ハイン" },
            { "hao", "ハオ" },
            { "hau", "ハウ" },
            { "hien", "ヒエン" },
            { "hiep", "ヒエップ" },
            { "hieu", "ヒエウ" },
            { "hoa", "ホア" },
            { "hoai", "ホアイ" },
            { "hoan", "ホアン" },
            { "hong", "ホン" },
            { "hop", "ホップ" },
            { "hue", "フエ" },
            { "hung", "フン" },
            { "huong", "フオン" },
            { "huu", "フー" },
            { "huy", "フイ" },
            { "huyen", "フエン" },
            { "kha", "カー" },
            { "khai", "カイ" },
            { "khanh", "カイン" },
            { "khiem", "キエム" },
            { "khoa", "コア" },
            { "khoi", "コイ" },
            { "khuong", "クオン" },
            { "kien", "キエン" },
            { "kiet", "キエット" },
            { "kieu", "キエウ" },
            { "kim", "キム" },
            { "ky", "キー" },
            { "lam", "ラム" },
            { "lan", "ラン" },
            { "lanh", "ライン" },
            { "lap", "ラップ" },
            { "liem", "リエム" },
            { "lien", "リエン" },
            { "linh", "リン" },
            { "loan", "ロアン" },
            { "loc", "ロック" },
            { "loi", "ロイ" },
            { "long", "ロン" },
            { "luan", "ルアン" },
            { "luc", "ルック" },
            { "luat", "ルアット" },
            { "man", "マン" },
            { "manh", "マイン" },
            { "minh", "ミン" },
            { "my", "ミー" },
            { "nam", "ナム" },
            { "nga", "ガー" },
            { "ngan", "ガン" },
            { "nghi", "ギー" },
            { "nghia", "ギア" },
            { "ngoc", "ゴック" },
            { "ngu", "グー" },
            { "nguyet", "グエット" },
            { "nhan", "ニャン" },
            { "nhat", "ニャット" },
            { "nhi", "ニー" },
            { "nhien", "ニエン" },
            { "nhu", "ニュー" },
            { "nhung", "ニュン" },
            { "nu", "ヌー" },
            { "oanh", "オアン" },
            { "phat", "ファット" },
            { "phi", "フィー" },
            { "phong", "フォン" },
            { "phu", "フー" },
            { "phuc", "フック" },
            { "phung", "フン" },
            { "phuoc", "フオック" },
            { "phuong", "フオン" },
            { "quan", "クアン" },
            { "quang", "クアン" },
            { "quoc", "クオック" },
            { "quy", "クイ" },
            { "quyen", "クエン" },
            { "quynh", "クイン" },
            { "sang", "サン" },
            { "sen", "セン" },
            { "sinh", "シン" },
            { "son", "ソン" },
            { "tai", "タイ" },
            { "tam", "タム" },
            { "tan", "タン" },
            { "tao", "タオ" },
            { "thach", "タック" },
            { "thang", "タン" },
            { "thanh", "タイン" },
            { "thao", "タオ" },
            { "thieng", "ティエン" },
            { "thien", "ティエン" },
            { "thinh", "ティン" },
            { "thoa", "トア" },
            { "tho", "トー" },
            { "thong", "トン" },
            { "thu", "トゥー" },
            { "thuan", "トゥアン" },
            { "thuc", "トゥック" },
            { "thung", "トゥン" },
            { "thuy", "トゥイ" },
            { "thuyen", "トゥエン" },
            { "tien", "ティエン" },
            { "tin", "ティン" },
            { "toan", "トアン" },
            { "tong", "トン" },
            { "tra", "チャ" },
            { "trang", "チャン" },
            { "tri", "チー" },
            { "trieu", "チエウ" },
            { "truc", "チュック" },
            { "trung", "チュン" },
            { "tu", "トゥー" },
            { "tuan", "トゥアン" },
            { "tung", "トゥン" },
            { "tuyet", "トゥエット" },
            { "uyen", "ウエン" },
            { "vinh", "ヴィン" },
            { "vuong", "ヴオン" },
            { "xuan", "スアン" },
            { "yen", "イエン" }
        };

        // Bảng quy tắc phụ âm đầu (Consonants)
        private static readonly Dictionary<string, string> ConsonantMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "ngh", "ギ" },
            { "ng", "グ" },
            { "nh", "ニ" },
            { "th", "ト" },
            { "tr", "チ" },
            { "ch", "チ" },
            { "ph", "フ" },
            { "kh", "ク" },
            { "gh", "グ" },
            { "qu", "ク" },
            { "gi", "ジ" },
            { "b", "ブ" },
            { "c", "ク" },
            { "d", "ズ" },
            { "đ", "ド" },
            { "g", "グ" },
            { "h", "ハ" },
            { "k", "ク" },
            { "l", "ル" },
            { "m", "ム" },
            { "n", "ヌ" },
            { "p", "プ" },
            { "r", "ラ" },
            { "s", "サ" },
            { "t", "ト" },
            { "v", "ヴ" },
            { "x", "サ" }
        };

        // Bảng nguyên âm (Vowels)
        private static readonly Dictionary<string, string> VowelMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "a", "ア" },
            { "ai", "アイ" },
            { "ao", "アオ" },
            { "au", "アウ" },
            { "ay", "アイ" },
            { "e", "エ" },
            { "eo", "エオ" },
            { "i", "イ" },
            { "ia", "イア" },
            { "ieu", "イエウ" },
            { "o", "オ" },
            { "oa", "オア" },
            { "oai", "オアイ" },
            { "oay", "オアイ" },
            { "oe", "オエ" },
            { "oi", "オイ" },
            { "oo", "オー" },
            { "u", "ウ" },
            { "ua", "ウア" },
            { "uay", "ウアイ" },
            { "ue", "ウエ" },
            { "ui", "ウイ" },
            { "uo", "ウオ" },
            { "uoi", "ウオイ" },
            { "uou", "ウオウ" },
            { "uy", "ウイ" },
            { "uye", "ウエ" },
            { "uyen", "ウエン" },
            { "uyu", "ウイウ" },
            { "y", "イ" },
            { "ye", "イエ" },
            { "yeu", "イエウ" }
        };

        /// <summary>
        /// Chuyển đổi một tên hoặc chuỗi tiếng Việt thành Katakana
        /// </summary>
        /// <param name="vietnameseText">Chuỗi tiếng Việt (vd: Nguyễn Văn A)</param>
        /// <param name="useMiddleDot">Sử dụng dấu chấm giữa (・) hay dấu cách</param>
        public static string ConvertToKatakana(string? vietnameseText, bool useMiddleDot = true)
        {
            if (string.IsNullOrWhiteSpace(vietnameseText)) return string.Empty;

            var words = vietnameseText!.Trim().Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            if (words.Length == 0) return string.Empty;

            var katakanaWords = new List<string>();
            foreach (var word in words)
            {
                katakanaWords.Add(ConvertWord(word));
            }

            string separator = useMiddleDot ? "・" : " ";
            return string.Join(separator, katakanaWords);
        }

        private static string ConvertWord(string word)
        {
            if (string.IsNullOrWhiteSpace(word)) return string.Empty;

            // Xóa toàn bộ dấu tiếng Việt về dạng không dấu để tra cứu
            string normalized = RemoveDiacritics(word).ToLowerInvariant();

            // 1. Tra cứu trực tiếp trong từ điển âm tiết
            if (SyllableMap.TryGetValue(normalized, out string? kata))
            {
                return kata;
            }

            // 2. Thuật toán phân rã ngữ âm Fallback (Consonant + Vowel + Coda)
            return ConvertPhoneticFallback(normalized);
        }

        private static string ConvertPhoneticFallback(string norm)
        {
            if (string.IsNullOrEmpty(norm)) return string.Empty;

            var sb = new StringBuilder();
            int idx = 0;

            // Bắt phụ âm đầu
            string consonant = string.Empty;
            string[] prefixConsonants = { "ngh", "ng", "nh", "th", "tr", "ch", "ph", "kh", "gh", "qu", "gi", "b", "c", "d", "đ", "g", "h", "k", "l", "m", "n", "p", "r", "s", "t", "v", "x" };
            foreach (var pc in prefixConsonants)
            {
                if (norm.StartsWith(pc, StringComparison.OrdinalIgnoreCase))
                {
                    consonant = pc;
                    idx = pc.Length;
                    break;
                }
            }

            string rest = norm.Substring(idx);
            if (consonant.Length > 0 && ConsonantMap.TryGetValue(consonant, out string? consKata))
            {
                sb.Append(consKata);
            }

            if (rest.Length > 0)
            {
                if (VowelMap.TryGetValue(rest, out string? vowKata))
                {
                    sb.Append(vowKata);
                }
                else
                {
                    // Chuyển từng ký tự
                    foreach (char c in rest)
                    {
                        string cs = c.ToString();
                        if (VowelMap.TryGetValue(cs, out string? ck)) sb.Append(ck);
                        else if (ConsonantMap.TryGetValue(cs, out string? cc)) sb.Append(cc);
                    }
                }
            }

            return sb.Length > 0 ? sb.ToString() : norm.ToUpperInvariant();
        }

        /// <summary>
        /// Xóa bỏ dấu tiếng Việt chuyển thành chữ không dấu (NFC)
        /// </summary>
        public static string RemoveDiacritics(string? text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;

            // Xóa ký tự đ, Đ đặc thù trước khi chuẩn hóa Unicode
            string step1 = text!.Replace("đ", "d").Replace("Đ", "D")
                                .Replace("₫", "d");

            string normalizedString = step1.Normalize(NormalizationForm.FormD);
            var stringBuilder = new StringBuilder();

            foreach (char c in normalizedString)
            {
                var unicodeCategory = CharUnicodeInfo.GetUnicodeCategory(c);
                if (unicodeCategory != UnicodeCategory.NonSpacingMark)
                {
                    stringBuilder.Append(c);
                }
            }

            return stringBuilder.ToString().Normalize(NormalizationForm.FormC);
        }
    }
}
