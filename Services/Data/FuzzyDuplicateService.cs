using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using ExcelSupport.Models;
using Microsoft.Office.Interop.Excel;

namespace ExcelSupport.Services
{
    public static class FuzzyDuplicateService
    {
        public static string CleanString(string text, bool ignoreCase, bool ignoreAccent, bool cleanInvisibleSpaces)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;

            string s = text;

            if (cleanInvisibleSpaces)
            {
                // Thay thế khoảng trắng không ngắt dòng (NBSP) và ký tự vô hình
                s = s.Replace('\u00A0', ' ')
                     .Replace('\u200B', ' ')
                     .Replace('\u200C', ' ')
                     .Replace('\u200D', ' ')
                     .Replace('\uFEFF', ' ')
                     .Replace('\t', ' ');
            }

            s = Regex.Replace(s, @"\s+", " ").Trim();

            if (ignoreAccent)
            {
                s = TableMergeService.RemoveDiacritics(s);
            }

            if (ignoreCase)
            {
                s = s.ToLowerInvariant();
            }

            return s;
        }

        public static double CalculateSimilarity(string s1, string s2, FuzzyMatchAlgorithm algo)
        {
            if (string.Equals(s1, s2, StringComparison.OrdinalIgnoreCase)) return 100.0;
            if (string.IsNullOrEmpty(s1) || string.IsNullOrEmpty(s2)) return 0.0;

            switch (algo)
            {
                case FuzzyMatchAlgorithm.Levenshtein:
                    return CalculateLevenshteinSimilarity(s1, s2);

                case FuzzyMatchAlgorithm.JaroWinkler:
                default:
                    return CalculateJaroWinklerSimilarity(s1, s2);
            }
        }

        public static double CalculateLevenshteinSimilarity(string s1, string s2)
        {
            int len1 = s1.Length;
            int len2 = s2.Length;
            if (len1 == 0) return len2 == 0 ? 100.0 : 0.0;
            if (len2 == 0) return 0.0;

            int[,] d = new int[len1 + 1, len2 + 1];

            for (int i = 0; i <= len1; i++) d[i, 0] = i;
            for (int j = 0; j <= len2; j++) d[0, j] = j;

            for (int i = 1; i <= len1; i++)
            {
                for (int j = 1; j <= len2; j++)
                {
                    int cost = (s1[i - 1] == s2[j - 1]) ? 0 : 1;
                    d[i, j] = Math.Min(
                        Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1),
                        d[i - 1, j - 1] + cost);
                }
            }

            int dist = d[len1, len2];
            int maxLen = Math.Max(len1, len2);
            return (1.0 - (double)dist / maxLen) * 100.0;
        }

        public static double CalculateJaroWinklerSimilarity(string s1, string s2)
        {
            int l1 = s1.Length;
            int l2 = s2.Length;
            if (l1 == 0 || l2 == 0) return 0.0;

            int matchDistance = Math.Max(l1, l2) / 2 - 1;
            if (matchDistance < 0) matchDistance = 0;

            bool[] s1Matches = new bool[l1];
            bool[] s2Matches = new bool[l2];

            int matches = 0;
            for (int i = 0; i < l1; i++)
            {
                int start = Math.Max(0, i - matchDistance);
                int end = Math.Min(i + matchDistance + 1, l2);

                for (int j = start; j < end; j++)
                {
                    if (s2Matches[j]) continue;
                    if (s1[i] != s2[j]) continue;

                    s1Matches[i] = true;
                    s2Matches[j] = true;
                    matches++;
                    break;
                }
            }

            if (matches == 0) return 0.0;

            int transpositions = 0;
            int k = 0;
            for (int i = 0; i < l1; i++)
            {
                if (!s1Matches[i]) continue;
                while (!s2Matches[k]) k++;
                if (s1[i] != s2[k]) transpositions++;
                k++;
            }

            double jaro = ((double)matches / l1 + (double)matches / l2 + (double)(matches - transpositions / 2) / matches) / 3.0;

            // Winkler bonus cho tiền tố chung (prefix up to 4 chars)
            int prefix = 0;
            for (int i = 0; i < Math.Min(4, Math.Min(l1, l2)); i++)
            {
                if (s1[i] == s2[i]) prefix++;
                else break;
            }

            double jaroWinkler = jaro + prefix * 0.1 * (1.0 - jaro);
            return Math.Min(100.0, jaroWinkler * 100.0);
        }

        public static List<FuzzyClusterGroup> ScanFuzzyDuplicates(_Worksheet ws, FuzzyScanOptions options)
        {
            var clusters = new List<FuzzyClusterGroup>();
            if (ws == null || options == null) return clusters;

            Range? usedRange = null;
            try
            {
                usedRange = ws.UsedRange;
                if (usedRange == null || usedRange.Rows.Count == 0) return clusters;

                int startRow = usedRange.Row;
                int startCol = usedRange.Column;
                int numRows = usedRange.Rows.Count;
                int numCols = usedRange.Columns.Count;

                int colOffset = options.TargetColumnIndex - startCol + 1;
                if (colOffset < 1 || colOffset > numCols) return clusters;

                object? rawVal = usedRange.Value2;
                if (!(rawVal is object[,] allVals)) return clusters;

                var records = new List<FuzzyRecordItem>();

                for (int r = 1; r <= numRows; r++)
                {
                    int absoluteRow = startRow + r - 1;
                    if (absoluteRow < options.StartRow) continue;

                    object? cellVal = allVals[r, colOffset];
                    if (cellVal == null) continue;

                    string original = cellVal.ToString()?.Trim() ?? string.Empty;
                    if (string.IsNullOrEmpty(original)) continue;

                    string normalized = CleanString(original, options.IgnoreCase, options.IgnoreAccent, options.CleanInvisibleSpaces);
                    if (string.IsNullOrEmpty(normalized)) continue;

                    records.Add(new FuzzyRecordItem
                    {
                        RowIndex = absoluteRow,
                        CellAddress = ws.Cells[absoluteRow, options.TargetColumnIndex].Address[false, false],
                        OriginalText = original,
                        NormalizedText = normalized,
                        SimilarityPercent = 100.0,
                        IsSelected = true
                    });
                }

                if (records.Count < 2) return clusters;

                // Thuật toán gom cụm (Clustering)
                var visited = new bool[records.Count];
                int groupIdCounter = 1;

                for (int i = 0; i < records.Count; i++)
                {
                    if (visited[i]) continue;

                    var groupItems = new List<FuzzyRecordItem> { records[i] };
                    visited[i] = true;

                    for (int j = i + 1; j < records.Count; j++)
                    {
                        if (visited[j]) continue;

                        double sim = CalculateSimilarity(records[i].NormalizedText, records[j].NormalizedText, options.Algorithm);
                        if (sim >= options.SimilarityThreshold)
                        {
                            var item = records[j];
                            item.SimilarityPercent = sim;
                            groupItems.Add(item);
                            visited[j] = true;
                        }
                    }

                    // Chỉ coi là nhóm trùng lặp khi có từ 2 bản ghi trở lên
                    if (groupItems.Count >= 2)
                    {
                        // Chọn giá trị xuất hiện nhiều nhất làm MasterValue
                        string masterVal = groupItems
                            .GroupBy(x => x.OriginalText)
                            .OrderByDescending(g => g.Count())
                            .First().Key;

                        clusters.Add(new FuzzyClusterGroup
                        {
                            GroupId = groupIdCounter++,
                            MasterValue = masterVal,
                            Items = groupItems.OrderByDescending(x => x.SimilarityPercent).ToList()
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ScanFuzzyDuplicates error: {ex.Message}");
            }
            finally
            {
                if (usedRange != null) Marshal.ReleaseComObject(usedRange);
            }

            return clusters;
        }

        public static int StandardizeValues(_Worksheet ws, int targetColIndex, List<FuzzyClusterGroup> groups)
        {
            if (ws == null || groups == null || groups.Count == 0) return 0;

            int count = 0;
            try
            {
                ws.Application.ScreenUpdating = false;

                foreach (var group in groups)
                {
                    foreach (var item in group.Items)
                    {
                        if (item.IsSelected && item.OriginalText != group.MasterValue)
                        {
                            try
                            {
                                Range cell = ws.Cells[item.RowIndex, targetColIndex];
                                cell.Value2 = group.MasterValue;
                                Marshal.ReleaseComObject(cell);
                                count++;
                            }
                            catch { }
                        }
                    }
                }
            }
            finally
            {
                try { ws.Application.ScreenUpdating = true; } catch { }
            }

            return count;
        }

        public static int HighlightClusters(_Worksheet ws, int targetColIndex, List<FuzzyClusterGroup> groups)
        {
            if (ws == null || groups == null || groups.Count == 0) return 0;

            int[] pastelColors = new int[]
            {
                ColorTranslator.ToOle(Color.FromArgb(254, 240, 138)), // Vàng nhạt
                ColorTranslator.ToOle(Color.FromArgb(187, 247, 208)), // Xanh lá nhạt
                ColorTranslator.ToOle(Color.FromArgb(186, 230, 253)), // Xanh biển nhạt
                ColorTranslator.ToOle(Color.FromArgb(233, 213, 255)), // Tím nhạt
                ColorTranslator.ToOle(Color.FromArgb(254, 202, 202)), // Đỏ hồng nhạt
                ColorTranslator.ToOle(Color.FromArgb(254, 215, 170))  // Cam nhạt
            };

            int cellCount = 0;
            try
            {
                ws.Application.ScreenUpdating = false;

                for (int g = 0; g < groups.Count; g++)
                {
                    int colorOle = pastelColors[g % pastelColors.Length];
                    foreach (var item in groups[g].Items)
                    {
                        if (item.IsSelected)
                        {
                            try
                            {
                                Range cell = ws.Cells[item.RowIndex, targetColIndex];
                                cell.Interior.Color = colorOle;
                                Marshal.ReleaseComObject(cell);
                                cellCount++;
                            }
                            catch { }
                        }
                    }
                }
            }
            finally
            {
                try { ws.Application.ScreenUpdating = true; } catch { }
            }

            return cellCount;
        }
    }
}
