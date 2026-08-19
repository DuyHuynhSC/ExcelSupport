using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace ExcelSupport.Services
{
    public static class VietnameseToKatakanaConverter
    {
        // Bảng ánh xạ các âm tiết / tên tiếng Việt thông dụng sang Katakana
        private static readonly Dictionary<string, string> SyllableMap;
        private static readonly Dictionary<string, string> ConsonantMap;
        private static readonly Dictionary<string, string> VowelMap;

        static VietnameseToKatakanaConverter()
        {
            SyllableMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            ConsonantMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            VowelMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            try
            {
                // Họ phổ biến
                AddSyllable("nguyen", "グエン");
                AddSyllable("tran", "チャン");
                AddSyllable("le", "レ");
                AddSyllable("pham", "ファム");
                AddSyllable("hoang", "ホアン");
                AddSyllable("huynh", "フイン");
                AddSyllable("phan", "ファン");
                AddSyllable("vu", "ヴー");
                AddSyllable("vo", "ヴォ");
                AddSyllable("dang", "ダン");
                AddSyllable("bui", "ブイ");
                AddSyllable("do", "ドー");
                AddSyllable("ho", "ホー");
                AddSyllable("ngo", "ゴー");
                AddSyllable("duong", "ズオン");
                AddSyllable("ly", "リー");
                AddSyllable("dinh", "ディン");
                AddSyllable("doan", "ドアン");
                AddSyllable("truong", "チュオン");
                AddSyllable("luong", "ルオン");
                AddSyllable("tong", "トン");
                AddSyllable("trinh", "チン");
                AddSyllable("dao", "ダオ");
                AddSyllable("ha", "ハー");
                AddSyllable("mai", "マイ");
                AddSyllable("cao", "カオ");
                AddSyllable("ta", "ター");
                AddSyllable("thai", "タイ");
                AddSyllable("chu", "チュー");
                AddSyllable("luu", "ルー");

                // Tên đệm & Tên chính phổ biến
                AddSyllable("van", "ヴァン");
                AddSyllable("thi", "ティ");
                AddSyllable("anh", "アイン");
                AddSyllable("an", "アン");
                AddSyllable("bac", "バック");
                AddSyllable("bach", "バック");
                AddSyllable("bao", "バオ");
                AddSyllable("bich", "ビック");
                AddSyllable("binh", "ビン");
                AddSyllable("cam", "カム");
                AddSyllable("canh", "カイン");
                AddSyllable("chau", "チャウ");
                AddSyllable("chi", "チー");
                AddSyllable("chien", "チエン");
                AddSyllable("chinh", "チン");
                AddSyllable("chung", "チュン");
                AddSyllable("cong", "コン");
                AddSyllable("cuc", "クック");
                AddSyllable("cuong", "クオン");
                AddSyllable("dai", "ダイ");
                AddSyllable("dan", "ダン");
                AddSyllable("dat", "ダット");
                AddSyllable("dien", "ディエン");
                AddSyllable("diep", "ディエップ");
                AddSyllable("dieu", "ディエウ");
                AddSyllable("doanh", "ドアン");
                AddSyllable("dong", "ドン");
                AddSyllable("duc", "ドゥック");
                AddSyllable("dung", "ズン");
                AddSyllable("duy", "ズイ");
                AddSyllable("duyen", "ズエン");
                AddSyllable("giang", "ザン");
                AddSyllable("giao", "ザオ");
                AddSyllable("hai", "ハイ");
                AddSyllable("han", "ハン");
                AddSyllable("hanh", "ハイン");
                AddSyllable("hao", "ハオ");
                AddSyllable("hau", "ハウ");
                AddSyllable("hien", "ヒエン");
                AddSyllable("hiep", "ヒエップ");
                AddSyllable("hieu", "ヒエウ");
                AddSyllable("hoa", "ホア");
                AddSyllable("hoai", "ホアイ");
                AddSyllable("hoan", "ホアン");
                AddSyllable("hong", "ホン");
                AddSyllable("hop", "ホップ");
                AddSyllable("hue", "フエ");
                AddSyllable("hung", "フン");
                AddSyllable("huong", "フオン");
                AddSyllable("huu", "フー");
                AddSyllable("huy", "フイ");
                AddSyllable("huyen", "フエン");
                AddSyllable("kha", "カー");
                AddSyllable("khai", "カイ");
                AddSyllable("khanh", "カイン");
                AddSyllable("khiem", "キエム");
                AddSyllable("khoa", "コア");
                AddSyllable("khoi", "コイ");
                AddSyllable("khuong", "クオン");
                AddSyllable("kien", "キエン");
                AddSyllable("kiet", "キエット");
                AddSyllable("kieu", "キエウ");
                AddSyllable("kim", "キム");
                AddSyllable("ky", "キー");
                AddSyllable("lam", "ラム");
                AddSyllable("lan", "ラン");
                AddSyllable("lanh", "ライン");
                AddSyllable("lap", "ラップ");
                AddSyllable("liem", "リエム");
                AddSyllable("lien", "リエン");
                AddSyllable("linh", "リン");
                AddSyllable("loan", "ロアン");
                AddSyllable("loc", "ロック");
                AddSyllable("loi", "ロイ");
                AddSyllable("long", "ロン");
                AddSyllable("luan", "ルアン");
                AddSyllable("luc", "ルック");
                AddSyllable("luat", "ルアット");
                AddSyllable("man", "マン");
                AddSyllable("manh", "マイン");
                AddSyllable("minh", "ミン");
                AddSyllable("my", "ミー");
                AddSyllable("nam", "ナム");
                AddSyllable("nga", "ガー");
                AddSyllable("ngan", "ガン");
                AddSyllable("nghi", "ギー");
                AddSyllable("nghia", "ギア");
                AddSyllable("ngoc", "ゴック");
                AddSyllable("ngu", "グー");
                AddSyllable("nguyet", "グエット");
                AddSyllable("nhan", "ニャン");
                AddSyllable("nhat", "ニャット");
                AddSyllable("nhi", "ニー");
                AddSyllable("nhien", "ニエン");
                AddSyllable("nhu", "ニュー");
                AddSyllable("nhung", "ニュン");
                AddSyllable("nu", "ヌー");
                AddSyllable("oanh", "オアン");
                AddSyllable("phat", "ファット");
                AddSyllable("phi", "フィー");
                AddSyllable("phong", "フォン");
                AddSyllable("phu", "フー");
                AddSyllable("phuc", "フック");
                AddSyllable("phung", "フン");
                AddSyllable("phuoc", "フオック");
                AddSyllable("phuong", "フオン");
                AddSyllable("quan", "クアン");
                AddSyllable("quang", "クアン");
                AddSyllable("quoc", "クオック");
                AddSyllable("quy", "クイ");
                AddSyllable("quyen", "クエン");
                AddSyllable("quynh", "クイン");
                AddSyllable("sang", "サン");
                AddSyllable("sen", "セン");
                AddSyllable("sinh", "シン");
                AddSyllable("son", "ソン");
                AddSyllable("tai", "タイ");
                AddSyllable("tam", "タム");
                AddSyllable("tan", "タン");
                AddSyllable("tao", "タオ");
                AddSyllable("thach", "タック");
                AddSyllable("thang", "タン");
                AddSyllable("thanh", "タイン");
                AddSyllable("thao", "タオ");
                AddSyllable("thieng", "ティエン");
                AddSyllable("thien", "ティエン");
                AddSyllable("thinh", "ティン");
                AddSyllable("thoa", "トア");
                AddSyllable("tho", "トー");
                AddSyllable("thong", "トン");
                AddSyllable("thu", "トゥー");
                AddSyllable("thuan", "トゥアン");
                AddSyllable("thuc", "トゥック");
                AddSyllable("thung", "トゥン");
                AddSyllable("thuy", "トゥイ");
                AddSyllable("thuyen", "トゥエン");
                AddSyllable("tien", "ティエン");
                AddSyllable("tin", "ティン");
                AddSyllable("toan", "トアン");
                AddSyllable("tra", "チャ");
                AddSyllable("trang", "チャン");
                AddSyllable("tri", "チー");
                AddSyllable("trieu", "チエウ");
                AddSyllable("truc", "チュック");
                AddSyllable("trung", "チュン");
                AddSyllable("tu", "トゥー");
                AddSyllable("tuan", "トゥアン");
                AddSyllable("tung", "トゥン");
                AddSyllable("tuyet", "トゥエット");
                AddSyllable("uyen", "ウエン");
                AddSyllable("vinh", "ヴィン");
                AddSyllable("vuong", "ヴオン");
                AddSyllable("xuan", "スアン");
                AddSyllable("yen", "イエン");

                // Phụ âm đầu
                AddConsonant("ngh", "ギ");
                AddConsonant("ng", "グ");
                AddConsonant("nh", "ニ");
                AddConsonant("th", "ト");
                AddConsonant("tr", "チ");
                AddConsonant("ch", "チ");
                AddConsonant("ph", "フ");
                AddConsonant("kh", "ク");
                AddConsonant("gh", "グ");
                AddConsonant("qu", "ク");
                AddConsonant("gi", "ジ");
                AddConsonant("b", "ブ");
                AddConsonant("c", "ク");
                AddConsonant("d", "ズ");
                AddConsonant("đ", "ド");
                AddConsonant("g", "グ");
                AddConsonant("h", "ハ");
                AddConsonant("k", "ク");
                AddConsonant("l", "ル");
                AddConsonant("m", "ム");
                AddConsonant("n", "ヌ");
                AddConsonant("p", "プ");
                AddConsonant("r", "ラ");
                AddConsonant("s", "サ");
                AddConsonant("t", "ト");
                AddConsonant("v", "ヴ");
                AddConsonant("x", "サ");

                // Nguyên âm
                AddVowel("a", "ア");
                AddVowel("ai", "アイ");
                AddVowel("ao", "アオ");
                AddVowel("au", "アウ");
                AddVowel("ay", "アイ");
                AddVowel("e", "エ");
                AddVowel("eo", "エオ");
                AddVowel("i", "イ");
                AddVowel("ia", "イア");
                AddVowel("ieu", "イエウ");
                AddVowel("o", "オ");
                AddVowel("oa", "オア");
                AddVowel("oai", "オアイ");
                AddVowel("oay", "オアイ");
                AddVowel("oe", "オエ");
                AddVowel("oi", "オイ");
                AddVowel("oo", "オー");
                AddVowel("u", "ウ");
                AddVowel("ua", "ウア");
                AddVowel("uay", "ウアイ");
                AddVowel("ue", "ウエ");
                AddVowel("ui", "ウイ");
                AddVowel("uo", "ウオ");
                AddVowel("uoi", "ウオイ");
                AddVowel("uou", "ウオウ");
                AddVowel("uy", "ウイ");
                AddVowel("uye", "ウエ");
                AddVowel("uyen", "ウエン");
                AddVowel("uyu", "ウイウ");
                AddVowel("y", "イ");
                AddVowel("ye", "イエ");
                AddVowel("yeu", "イエウ");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"VietnameseToKatakanaConverter static init error: {ex.Message}");
            }
        }

        private static void AddSyllable(string key, string val)
        {
            if (!string.IsNullOrWhiteSpace(key))
            {
                SyllableMap[key.Trim().ToLowerInvariant()] = val;
            }
        }

        private static void AddConsonant(string key, string val)
        {
            if (!string.IsNullOrWhiteSpace(key))
            {
                ConsonantMap[key.Trim().ToLowerInvariant()] = val;
            }
        }

        private static void AddVowel(string key, string val)
        {
            if (!string.IsNullOrWhiteSpace(key))
            {
                VowelMap[key.Trim().ToLowerInvariant()] = val;
            }
        }

        /// <summary>
        /// Chuyển đổi một tên hoặc chuỗi tiếng Việt thành Katakana
        /// </summary>
        /// <param name="vietnameseText">Chuỗi tiếng Việt (vd: Nguyễn Văn Ánh)</param>
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
