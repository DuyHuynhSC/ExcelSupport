using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;

namespace ExcelSupport.Services
{
    public class KatakanaCellLocation
    {
        public string WorkbookName { get; set; } = string.Empty;
        public string SheetName { get; set; } = string.Empty;
        public string CellAddress { get; set; } = string.Empty;
        public int Row { get; set; }
        public int Column { get; set; }
        public string OriginalText { get; set; } = string.Empty;
    }

    public class KatakanaVariantItem
    {
        public string Word { get; set; } = string.Empty;
        public int OccurrenceCount { get; set; }
        public bool HasTrailingChouon => Word.EndsWith("ー");
        public List<KatakanaCellLocation> Locations { get; set; } = new();
    }

    public class KatakanaInconsistencyGroup
    {
        public string BaseStem { get; set; } = string.Empty;
        public List<KatakanaVariantItem> Variants { get; set; } = new();
        public int TotalOccurrences => Variants.Sum(v => v.OccurrenceCount);
        public string PreferredWord { get; set; } = string.Empty;
        public string WithChouonWord { get; set; } = string.Empty;
        public string WithoutChouonWord { get; set; } = string.Empty;

        public string DisplaySummary => string.Join(" ⇋ ", Variants.Select(v => $"{v.Word} ({v.OccurrenceCount})"));
    }

    public static class KatakanaValidatorService
    {
        // Regex tìm các cụm từ Katakana liên tiếp (bao gồm cả trường âm 'ー') dài từ 2 ký tự trở lên
        private static readonly Regex KatakanaRegex = new(@"[\u30A1-\u30FA\u30FC]{2,}", RegexOptions.Compiled);

        /// <summary>
        /// Quét toàn bộ bảng tính để phát hiện các từ Katakana viết không đồng nhất (lệch chuẩn trường âm hoặc biến thể).
        /// </summary>
        public static List<KatakanaInconsistencyGroup> ScanInconsistencies(ExcelApp app, ConversionScope scope, Action<string, int>? progressCallback = null)
        {
            var groups = new List<KatakanaInconsistencyGroup>();
            if (app == null) return groups;

            var rawOccurrences = new Dictionary<string, List<KatakanaCellLocation>>(StringComparer.Ordinal);
            var workbooksToScan = new List<Workbook>();

            if (scope == ConversionScope.Selection)
            {
                dynamic sel = app.Selection;
                if (sel is Range selRng)
                {
                    ScanRangeForKatakana(selRng, app.ActiveWorkbook?.Name ?? "", app.ActiveSheet?.Name ?? "", rawOccurrences);
                }
            }
            else if (scope == ConversionScope.ActiveSheet)
            {
                if (app.ActiveSheet is Worksheet ws && ws.UsedRange != null)
                {
                    ScanRangeForKatakana(ws.UsedRange, app.ActiveWorkbook?.Name ?? "", ws.Name, rawOccurrences);
                }
            }
            else if (scope == ConversionScope.ActiveWorkbook)
            {
                if (app.ActiveWorkbook is Workbook wb)
                {
                    int sheetCount = wb.Worksheets.Count;
                    for (int i = 1; i <= sheetCount; i++)
                    {
                        Worksheet ws = wb.Worksheets[i];
                        progressCallback?.Invoke($"Đang quét Sheet {ws.Name} ({i}/{sheetCount})...", (int)((double)i / sheetCount * 100));
                        if (ws.UsedRange != null)
                        {
                            ScanRangeForKatakana(ws.UsedRange, wb.Name, ws.Name, rawOccurrences);
                        }
                    }
                }
            }

            // Gom nhóm các từ Katakana theo gốc chuẩn (Stem)
            var stemMap = new Dictionary<string, List<string>>(StringComparer.Ordinal);

            foreach (var word in rawOccurrences.Keys)
            {
                // Chuẩn hóa gốc: bỏ ký tự trường âm cuối từ nếu có
                string stem = word.TrimEnd('ー');
                if (stem.Length < 2) stem = word;

                if (!stemMap.ContainsKey(stem))
                {
                    stemMap[stem] = new List<string>();
                }
                stemMap[stem].Add(word);
            }

            // Lọc ra các nhóm có từ 2 biến thể khác nhau trở lên
            foreach (var kvp in stemMap)
            {
                var wordList = kvp.Value.Distinct(StringComparer.Ordinal).ToList();
                if (wordList.Count > 1)
                {
                    var group = new KatakanaInconsistencyGroup
                    {
                        BaseStem = kvp.Key
                    };

                    foreach (var w in wordList)
                    {
                        var locs = rawOccurrences[w];
                        group.Variants.Add(new KatakanaVariantItem
                        {
                            Word = w,
                            OccurrenceCount = locs.Count,
                            Locations = locs
                        });
                    }

                    // Sắp xếp các biến thể theo tần suất giảm dần
                    group.Variants = group.Variants.OrderByDescending(v => v.OccurrenceCount).ToList();

                    // Xác định dạng có trường âm và không có trường âm
                    group.WithChouonWord = group.Variants.FirstOrDefault(v => v.HasTrailingChouon)?.Word ?? (group.BaseStem + "ー");
                    group.WithoutChouonWord = group.Variants.FirstOrDefault(v => !v.HasTrailingChouon)?.Word ?? group.BaseStem;

                    // Mặc định chọn từ có tần suất xuất hiện nhiều nhất
                    group.PreferredWord = group.Variants[0].Word;

                    groups.Add(group);
                }
            }

            return groups.OrderByDescending(g => g.TotalOccurrences).ToList();
        }

        private static void ScanRangeForKatakana(Range rng, string wbName, string wsName, Dictionary<string, List<KatakanaCellLocation>> occurrences)
        {
            if (rng == null) return;

            int rowCount = rng.Rows.Count;
            int colCount = rng.Columns.Count;
            int startRow = rng.Row;
            int startCol = rng.Column;

            if (rowCount == 1 && colCount == 1)
            {
                ExtractKatakanaWords(rng.Value2?.ToString(), wbName, wsName, rng.Address, startRow, startCol, occurrences);
                return;
            }

            object[,] values;
            try
            {
                object raw = rng.Value2;
                if (raw is object[,] arr) values = arr;
                else return;
            }
            catch
            {
                return;
            }

            int rLower = values.GetLowerBound(0);
            int rUpper = values.GetUpperBound(0);
            int cLower = values.GetLowerBound(1);
            int cUpper = values.GetUpperBound(1);

            for (int r = rLower; r <= rUpper; r++)
            {
                for (int c = cLower; c <= cUpper; c++)
                {
                    object? cellVal = values[r, c];
                    if (cellVal is string str && !string.IsNullOrEmpty(str))
                    {
                        int actualRow = startRow + (r - rLower);
                        int actualCol = startCol + (c - cLower);
                        string cellAddr = GetCellAddress(actualRow, actualCol);
                        ExtractKatakanaWords(str, wbName, wsName, cellAddr, actualRow, actualCol, occurrences);
                    }
                }
            }
        }

        private static void ExtractKatakanaWords(string? text, string wbName, string wsName, string address, int row, int col, Dictionary<string, List<KatakanaCellLocation>> occurrences)
        {
            if (string.IsNullOrEmpty(text)) return;

            var matches = KatakanaRegex.Matches(text);
            foreach (Match m in matches)
            {
                string word = m.Value;
                if (!occurrences.ContainsKey(word))
                {
                    occurrences[word] = new List<KatakanaCellLocation>();
                }

                occurrences[word].Add(new KatakanaCellLocation
                {
                    WorkbookName = wbName,
                    SheetName = wsName,
                    CellAddress = address,
                    Row = row,
                    Column = col,
                    OriginalText = text ?? string.Empty
                });
            }
        }

        private static string GetCellAddress(int row, int col)
        {
            string colLetter = string.Empty;
            int div = col;
            while (div > 0)
            {
                int mod = (div - 1) % 26;
                colLetter = (char)(65 + mod) + colLetter;
                div = (div - mod) / 26;
            }
            return $"{colLetter}{row}";
        }

        /// <summary>
        /// Chuẩn hóa đồng loạt một từ Katakana cũ thành từ mới đã chọn trên toàn bộ các vị trí đã ghi nhận.
        /// </summary>
        public static int ReplaceKatakanaWord(ExcelApp app, string oldWord, string newWord, List<KatakanaCellLocation> locations)
        {
            if (app == null || string.IsNullOrEmpty(oldWord) || string.IsNullOrEmpty(newWord) || oldWord == newWord || locations == null || locations.Count == 0)
            {
                return 0;
            }

            int replacedCount = 0;
            bool prevScreenUpdating = app.ScreenUpdating;

            try
            {
                app.ScreenUpdating = false;

                // Gom nhóm các vị trí theo Sheet để tối ưu tốc độ ghi
                var sheetGroups = locations.GroupBy(loc => (loc.WorkbookName, loc.SheetName));

                foreach (var sg in sheetGroups)
                {
                    try
                    {
                        Workbook? targetWb = null;
                        if (!string.IsNullOrEmpty(sg.Key.WorkbookName))
                        {
                            try { targetWb = app.Workbooks[sg.Key.WorkbookName]; } catch { }
                        }
                        targetWb ??= app.ActiveWorkbook;
                        if (targetWb == null) continue;

                        Worksheet? ws = null;
                        try { ws = targetWb.Worksheets[sg.Key.SheetName]; } catch { }
                        if (ws == null) continue;

                        foreach (var loc in sg)
                        {
                            try
                            {
                                Range cell = ws.Range[loc.CellAddress];
                                object? val = cell.Value2;
                                if (val is string str && str.Contains(oldWord))
                                {
                                    cell.Value2 = str.Replace(oldWord, newWord);
                                    replacedCount++;
                                }
                            }
                            catch { }
                        }
                    }
                    catch { }
                }

                try
                {
                    app.StatusBar = $"✨ ExcelSupport: Đã chuẩn hóa [{oldWord} ➔ {newWord}] cho {replacedCount} ô!";
                }
                catch { }

                return replacedCount;
            }
            finally
            {
                try { app.ScreenUpdating = prevScreenUpdating; } catch { }
            }
        }
    }
}
