using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Office.Interop.Excel;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;

namespace ExcelSupport.Services
{
    public enum ConversionScope
    {
        Selection,
        ActiveSheet,
        ActiveWorkbook
    }

    public class JapaneseConversionOptions
    {
        public bool ToHankaku { get; set; } = true; // true = To Hankaku (Bán giác), false = To Zenkaku (Toàn giác)
        public bool ConvertAlpha { get; set; } = true; // A-Z, a-z
        public bool ConvertNumbers { get; set; } = true; // 0-9
        public bool ConvertKatakana { get; set; } = true; // アイウ ⇋ ｱｲｳ
        public bool ConvertPunctuation { get; set; } = true; // ()[]:;!?,./ etc.
        public bool ConvertSpace { get; set; } = true; // \u3000 ⇋ ' '
        public ConversionScope Scope { get; set; } = ConversionScope.Selection;
    }

    public class JapaneseConversionResult
    {
        public int TotalCellsProcessed { get; set; }
        public int TotalCellsChanged { get; set; }
        public int TotalCharactersChanged { get; set; }
        public TimeSpan Duration { get; set; }
    }

    public static class JapaneseTextConverterService
    {
        #region Mapping Dictionaries & Tables

        // Half-width Katakana to Full-width Katakana mappings (Single characters)
        private static readonly Dictionary<char, char> HankakuToZenkakuKana = new()
        {
            ['ｱ'] = 'ア', ['ｲ'] = 'イ', ['ｳ'] = 'ウ', ['ｴ'] = 'エ', ['ｵ'] = 'オ',
            ['ｶ'] = 'カ', ['ｷ'] = 'キ', ['ｸ'] = 'ク', ['ｹ'] = 'ケ', ['ｺ'] = 'コ',
            ['ｻ'] = 'サ', ['ｼ'] = 'シ', ['ｽ'] = 'ス', ['ｾ'] = 'セ', ['ｿ'] = 'ソ',
            ['ﾀ'] = 'タ', ['ﾁ'] = 'チ', ['ﾂ'] = 'ツ', ['ﾃ'] = 'テ', ['ﾄ'] = 'ト',
            ['ﾅ'] = 'ナ', ['ﾆ'] = 'ニ', ['ﾇ'] = 'ヌ', ['ﾈ'] = 'ネ', ['ﾉ'] = 'ノ',
            ['ﾊ'] = 'ハ', ['ﾋ'] = 'ヒ', ['ﾌ'] = 'フ', ['ﾍ'] = 'ヘ', ['ﾎ'] = 'ホ',
            ['ﾏ'] = 'マ', ['ﾐ'] = 'ミ', ['ﾑ'] = 'ム', ['ﾒ'] = 'メ', ['ﾓ'] = 'モ',
            ['ﾔ'] = 'ヤ', ['ﾕ'] = 'ユ', ['ﾖ'] = 'ヨ',
            ['ﾗ'] = 'ラ', ['ﾘ'] = 'リ', ['ﾙ'] = 'ル', ['ﾚ'] = 'レ', ['ﾛ'] = 'ロ',
            ['ﾜ'] = 'ワ', ['ｦ'] = 'ヲ', ['ﾝ'] = 'ン',
            ['ｧ'] = 'ァ', ['ｨ'] = 'ィ', ['ｩ'] = 'ゥ', ['ｪ'] = 'ェ', ['ｫ'] = 'ォ',
            ['ｬ'] = 'ャ', ['ｭ'] = 'ュ', ['ｮ'] = 'ョ', ['ｯ'] = 'ッ',
            ['ｰ'] = 'ー', ['･'] = '・', ['｢'] = '「', ['｣'] = '」', ['ﾞ'] = '゛', ['ﾟ'] = '゜'
        };

        // Half-width Katakana with Dakuten (Voiced Sound: ﾞ)
        private static readonly Dictionary<string, char> HankakuDakutenToZenkaku = new()
        {
            ["ｶﾞ"] = 'ガ', ["ｷﾞ"] = 'ギ', ["ｸﾞ"] = 'グ', ["ｹﾞ"] = 'ゲ', ["ｺﾞ"] = 'ゴ',
            ["ｻﾞ"] = 'ザ', ["ｼﾞ"] = 'ジ', ["ｽﾞ"] = 'ズ', ["ｾﾞ"] = 'ゼ', ["ｿﾞ"] = 'ゾ',
            ["ﾀﾞ"] = 'ダ', ["ﾁﾞ"] = 'ヂ', ["ﾂﾞ"] = 'ヅ', ["ﾃﾞ"] = 'デ', ["ﾄﾞ"] = 'ド',
            ["ﾊﾞ"] = 'バ', ["ﾋﾞ"] = 'ビ', ["ﾌﾞ"] = 'ブ', ["ﾍﾞ"] = 'ベ', ["ﾎﾞ"] = 'ボ',
            ["ｳﾞ"] = 'ヴ', ["ﾜﾞ"] = 'ヷ', ["ｦﾞ"] = 'ヺ'
        };

        // Half-width Katakana with Handakuten (Semi-Voiced Sound: ﾟ)
        private static readonly Dictionary<string, char> HankakuHandakutenToZenkaku = new()
        {
            ["ﾊﾟ"] = 'パ', ["ﾋﾟ"] = 'ピ', ["ﾌﾟ"] = 'プ', ["ﾍﾟ"] = 'ペ', ["ﾎﾟ"] = 'ポ'
        };

        // Full-width Katakana to Half-width Katakana string mapping
        private static readonly Dictionary<char, string> ZenkakuToHankakuKana = new();

        // Punctuation & Symbols Mapping
        private static readonly Dictionary<char, char> ZenkakuToHankakuSymbols = new()
        {
            ['！'] = '!', ['＂'] = '"', ['＃'] = '#', ['＄'] = '$', ['％'] = '%',
            ['＆'] = '&', ['＇'] = '\'', ['（'] = '(', ['）'] = ')', ['＊'] = '*',
            ['＋'] = '+', ['，'] = ',', ['－'] = '-', ['．'] = '.', ['／'] = '/',
            ['：'] = ':', ['；'] = ';', ['＜'] = '<', ['＝'] = '=', ['＞'] = '>',
            ['？'] = '?', ['＠'] = '@', ['［'] = '[', ['＼'] = '\\', ['］'] = ']',
            ['＾'] = '^', ['＿'] = '_', ['｀'] = '`', ['｛'] = '{', ['｜'] = '|',
            ['｝'] = '}', ['～'] = '~'
        };

        private static readonly Dictionary<char, char> HankakuToZenkakuSymbols = new();

        static JapaneseTextConverterService()
        {
            // Xây dựng bảng chuyển đổi ngược Zenkaku -> Hankaku Katakana
            foreach (var kvp in HankakuToZenkakuKana)
            {
                if (!ZenkakuToHankakuKana.ContainsKey(kvp.Value))
                {
                    ZenkakuToHankakuKana[kvp.Value] = kvp.Key.ToString();
                }
            }

            foreach (var kvp in HankakuDakutenToZenkaku)
            {
                ZenkakuToHankakuKana[kvp.Value] = kvp.Key;
            }

            foreach (var kvp in HankakuHandakutenToZenkaku)
            {
                ZenkakuToHankakuKana[kvp.Value] = kvp.Key;
            }

            // Xây dựng bảng chuyển đổi ngược Hankaku -> Zenkaku Symbols
            foreach (var kvp in ZenkakuToHankakuSymbols)
            {
                HankakuToZenkakuSymbols[kvp.Value] = kvp.Key;
            }
        }

        #endregion

        #region Core String Conversion Methods

        public static string ConvertText(string input, JapaneseConversionOptions options)
        {
            if (string.IsNullOrEmpty(input)) return input;
            return options.ToHankaku ? ConvertToHankaku(input, options) : ConvertToZenkaku(input, options);
        }

        /// <summary>
        /// Chuyển đổi chuỗi sang Bán Giác (Hankaku) theo cấu hình.
        /// </summary>
        public static string ConvertToHankaku(string input, JapaneseConversionOptions options)
        {
            if (string.IsNullOrEmpty(input)) return input;

            var sb = new StringBuilder(input.Length);
            for (int i = 0; i < input.Length; i++)
            {
                char c = input[i];

                // 1. Khoảng trắng Toàn giác \u3000 -> Bán giác ' '
                if (options.ConvertSpace && c == '\u3000')
                {
                    sb.Append(' ');
                    continue;
                }

                // 2. Chữ số Toàn giác ０-９ (0xFF10 - 0xFF19) -> 0-9
                if (options.ConvertNumbers && c >= '０' && c <= '９')
                {
                    sb.Append((char)(c - '０' + '0'));
                    continue;
                }

                // 3. Chữ cái Toàn giác Ａ-Ｚ, ａ-ｚ (0xFF21 - 0xFF5A) -> A-Z, a-z
                if (options.ConvertAlpha)
                {
                    if (c >= 'Ａ' && c <= 'Ｚ')
                    {
                        sb.Append((char)(c - 'Ａ' + 'A'));
                        continue;
                    }
                    if (c >= 'ａ' && c <= 'ｚ')
                    {
                        sb.Append((char)(c - 'ａ' + 'a'));
                        continue;
                    }
                }

                // 4. Ký hiệu / Dấu câu Toàn giác -> Bán giác
                if (options.ConvertPunctuation && ZenkakuToHankakuSymbols.TryGetValue(c, out char halfSymbol))
                {
                    sb.Append(halfSymbol);
                    continue;
                }

                // 5. Katakana Toàn giác -> Bán giác
                if (options.ConvertKatakana && ZenkakuToHankakuKana.TryGetValue(c, out string? halfKana))
                {
                    sb.Append(halfKana);
                    continue;
                }

                sb.Append(c);
            }

            return sb.ToString();
        }

        /// <summary>
        /// Chuyển đổi chuỗi sang Toàn Giác (Zenkaku) theo cấu hình.
        /// </summary>
        public static string ConvertToZenkaku(string input, JapaneseConversionOptions options)
        {
            if (string.IsNullOrEmpty(input)) return input;

            var sb = new StringBuilder(input.Length * 2);
            for (int i = 0; i < input.Length; i++)
            {
                char c = input[i];

                // 1. Khoảng trắng Bán giác ' ' -> Toàn giác \u3000
                if (options.ConvertSpace && c == ' ')
                {
                    sb.Append('\u3000');
                    continue;
                }

                // 2. Chữ số Bán giác 0-9 -> Toàn giác ０-９
                if (options.ConvertNumbers && c >= '0' && c <= '9')
                {
                    sb.Append((char)(c - '0' + '０'));
                    continue;
                }

                // 3. Chữ cái Bán giác A-Z, a-z -> Toàn giác Ａ-Ｚ, ａ-ｚ
                if (options.ConvertAlpha)
                {
                    if (c >= 'A' && c <= 'Z')
                    {
                        sb.Append((char)(c - 'A' + 'Ａ'));
                        continue;
                    }
                    if (c >= 'a' && c <= 'z')
                    {
                        sb.Append((char)(c - 'a' + 'ａ'));
                        continue;
                    }
                }

                // 4. Katakana Bán giác (xử lý ghép 2 ký tự: âm đục ﾞ và âm bán đục ﾟ)
                if (options.ConvertKatakana)
                {
                    if (i + 1 < input.Length)
                    {
                        char nextChar = input[i + 1];
                        if (nextChar == 'ﾞ')
                        {
                            string pair = $"{c}ﾞ";
                            if (HankakuDakutenToZenkaku.TryGetValue(pair, out char zenDakuten))
                            {
                                sb.Append(zenDakuten);
                                i++; // Bỏ qua ký tự ﾞ tiếp theo
                                continue;
                            }
                        }
                        else if (nextChar == 'ﾟ')
                        {
                            string pair = $"{c}ﾟ";
                            if (HankakuHandakutenToZenkaku.TryGetValue(pair, out char zenHandakuten))
                            {
                                sb.Append(zenHandakuten);
                                i++; // Bỏ qua ký tự ﾟ tiếp theo
                                continue;
                            }
                        }
                    }

                    if (HankakuToZenkakuKana.TryGetValue(c, out char zenKana))
                    {
                        sb.Append(zenKana);
                        continue;
                    }
                }

                // 5. Ký hiệu / Dấu câu Bán giác -> Toàn giác
                if (options.ConvertPunctuation && HankakuToZenkakuSymbols.TryGetValue(c, out char zenSymbol))
                {
                    sb.Append(zenSymbol);
                    continue;
                }

                sb.Append(c);
            }

            return sb.ToString();
        }

        #endregion

        #region Excel Range & Sheet Conversion Methods

        public static JapaneseConversionResult ExecuteConversion(ExcelApp app, JapaneseConversionOptions options, Action<string, int>? progressCallback = null)
        {
            var result = new JapaneseConversionResult();
            var startTime = DateTime.Now;

            if (app == null) return result;

            bool prevScreenUpdating = app.ScreenUpdating;
            bool prevDisplayAlerts = app.DisplayAlerts;

            try
            {
                app.ScreenUpdating = false;
                app.DisplayAlerts = false;

                var rangesToProcess = new List<Range>();

                if (options.Scope == ConversionScope.Selection)
                {
                    dynamic sel = app.Selection;
                    if (sel is Range selRange)
                    {
                        rangesToProcess.Add(selRange);
                    }
                }
                else if (options.Scope == ConversionScope.ActiveSheet)
                {
                    if (app.ActiveSheet is Worksheet ws)
                    {
                        Range used = ws.UsedRange;
                        if (used != null) rangesToProcess.Add(used);
                    }
                }
                else if (options.Scope == ConversionScope.ActiveWorkbook)
                {
                    if (app.ActiveWorkbook is Workbook wb)
                    {
                        foreach (Worksheet ws in wb.Worksheets)
                        {
                            Range used = ws.UsedRange;
                            if (used != null) rangesToProcess.Add(used);
                        }
                    }
                }

                int totalRanges = rangesToProcess.Count;
                int processedRanges = 0;

                foreach (var rng in rangesToProcess)
                {
                    processedRanges++;
                    progressCallback?.Invoke($"Đang xử lý vùng {processedRanges}/{totalRanges}...", (int)((double)processedRanges / totalRanges * 100));

                    ProcessRange(rng, options, result);
                }

                result.Duration = DateTime.Now - startTime;
                try
                {
                    string dir = options.ToHankaku ? "Toàn giác ➔ Bán giác" : "Bán giác ➔ Toàn giác";
                    app.StatusBar = $"✨ ExcelSupport: Đã chuyển đổi {dir} cho {result.TotalCellsChanged} ô ({result.TotalCharactersChanged} ký tự)!";
                }
                catch { }

                return result;
            }
            finally
            {
                try { app.ScreenUpdating = prevScreenUpdating; } catch { }
                try { app.DisplayAlerts = prevDisplayAlerts; } catch { }
            }
        }

        private static void ProcessRange(Range rng, JapaneseConversionOptions options, JapaneseConversionResult result)
        {
            if (rng == null) return;

            int rowCount = rng.Rows.Count;
            int colCount = rng.Columns.Count;

            if (rowCount == 1 && colCount == 1)
            {
                result.TotalCellsProcessed++;
                object? val = rng.Value2;
                if (val is string str && !string.IsNullOrEmpty(str))
                {
                    string converted = ConvertText(str, options);
                    if (converted != str)
                    {
                        rng.Value2 = converted;
                        result.TotalCellsChanged++;
                        result.TotalCharactersChanged += Math.Abs(converted.Length - str.Length) + 1;
                    }
                }
                return;
            }

            // Xử lý mảng 2 chiều tốc độ cao cho vùng ô lớn
            object[,] values;
            try
            {
                object raw = rng.Value2;
                if (raw is object[,] arr)
                {
                    values = arr;
                }
                else
                {
                    return;
                }
            }
            catch
            {
                return;
            }

            int rLower = values.GetLowerBound(0);
            int rUpper = values.GetUpperBound(0);
            int cLower = values.GetLowerBound(1);
            int cUpper = values.GetUpperBound(1);

            bool anyChange = false;

            for (int r = rLower; r <= rUpper; r++)
            {
                for (int c = cLower; c <= cUpper; c++)
                {
                    object? cellVal = values[r, c];
                    result.TotalCellsProcessed++;

                    if (cellVal is string str && !string.IsNullOrEmpty(str))
                    {
                        string converted = ConvertText(str, options);
                        if (converted != str)
                        {
                            values[r, c] = converted;
                            anyChange = true;
                            result.TotalCellsChanged++;
                            result.TotalCharactersChanged += Math.Abs(converted.Length - str.Length) + 1;
                        }
                    }
                }
            }

            if (anyChange)
            {
                try
                {
                    rng.Value2 = values;
                }
                catch { }
            }
        }

        #endregion
    }
}
