using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using ExcelSupport.Models;
using Newtonsoft.Json;

namespace ExcelSupport.Services
{
    public static class GlossaryService
    {
        #region JSON Import / Export

        public static List<GlossaryItem> ImportFromJson(string filePath)
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException("Không tìm thấy file JSON:", filePath);

            string json = File.ReadAllText(filePath, Encoding.UTF8);
            var items = JsonConvert.DeserializeObject<List<GlossaryItem>>(json);

            if (items == null) return new List<GlossaryItem>();

            // Filter out empty items
            return items.Where(i => !string.IsNullOrWhiteSpace(i.Japanese) || !string.IsNullOrWhiteSpace(i.Vietnamese)).ToList();
        }

        public static void ExportToJson(string filePath, IEnumerable<GlossaryItem> items)
        {
            var validItems = items.Where(i => !string.IsNullOrWhiteSpace(i.Japanese) || !string.IsNullOrWhiteSpace(i.Vietnamese)).ToList();
            string json = JsonConvert.SerializeObject(validItems, Formatting.Indented);
            File.WriteAllText(filePath, json, Encoding.UTF8);
        }

        #endregion

        #region CSV Import / Export

        public static List<GlossaryItem> ImportFromCsv(string filePath)
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException("Không tìm thấy file CSV:", filePath);

            var items = new List<GlossaryItem>();
            var lines = File.ReadAllLines(filePath, Encoding.UTF8);

            if (lines.Length == 0) return items;

            bool isFirstLine = true;
            foreach (var rawLine in lines)
            {
                if (string.IsNullOrWhiteSpace(rawLine)) continue;

                var columns = ParseCsvLine(rawLine);
                if (columns.Count == 0) continue;

                // Kiểm tra xem dòng đầu tiên có phải là header ("Japanese", "Tiếng Nhật", ...) không
                if (isFirstLine)
                {
                    isFirstLine = false;
                    string col0 = columns[0].Trim().ToLowerInvariant();
                    string col1 = columns.Count > 1 ? columns[1].Trim().ToLowerInvariant() : string.Empty;

                    if (col0.Contains("japanese") || col0.Contains("tiếng nhật") || col0.Contains("nhật") ||
                        col1.Contains("vietnamese") || col1.Contains("tiếng việt") || col1.Contains("việt"))
                    {
                        continue; // Bỏ qua header
                    }
                }

                string ja = columns.Count > 0 ? columns[0].Trim() : string.Empty;
                string vi = columns.Count > 1 ? columns[1].Trim() : string.Empty;
                string note = columns.Count > 2 ? columns[2].Trim() : string.Empty;

                if (!string.IsNullOrWhiteSpace(ja) || !string.IsNullOrWhiteSpace(vi))
                {
                    items.Add(new GlossaryItem
                    {
                        Japanese = ja,
                        Vietnamese = vi,
                        Note = note
                    });
                }
            }

            return items;
        }

        public static void ExportToCsv(string filePath, IEnumerable<GlossaryItem> items)
        {
            var sb = new StringBuilder();
            // Header
            sb.AppendLine("Japanese,Vietnamese,Note");

            foreach (var item in items)
            {
                if (string.IsNullOrWhiteSpace(item.Japanese) && string.IsNullOrWhiteSpace(item.Vietnamese))
                    continue;

                string ja = EscapeCsvField(item.Japanese ?? string.Empty);
                string vi = EscapeCsvField(item.Vietnamese ?? string.Empty);
                string note = EscapeCsvField(item.Note ?? string.Empty);

                sb.AppendLine($"{ja},{vi},{note}");
            }

            // Ghi file với UTF-8 with BOM để Excel mở không bị lỗi font Tiếng Nhật và Tiếng Việt
            var encodingWithBom = new UTF8Encoding(encoderShouldEmitUTF8Identifier: true);
            File.WriteAllText(filePath, sb.ToString(), encodingWithBom);
        }

        private static List<string> ParseCsvLine(string line)
        {
            var result = new List<string>();
            var current = new StringBuilder();
            bool inQuotes = false;

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];

                if (c == '"')
                {
                    if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                    {
                        current.Append('"');
                        i++; // Bỏ qua dấu nháy kép escape
                    }
                    else
                    {
                        inQuotes = !inQuotes;
                    }
                }
                else if (c == ',' && !inQuotes)
                {
                    result.Add(current.ToString());
                    current.Clear();
                }
                else
                {
                    current.Append(c);
                }
            }

            result.Add(current.ToString());
            return result;
        }

        private static string EscapeCsvField(string field)
        {
            if (string.IsNullOrEmpty(field)) return string.Empty;

            if (field.Contains(",") || field.Contains("\"") || field.Contains("\n") || field.Contains("\r"))
            {
                return "\"" + field.Replace("\"", "\"\"") + "\"";
            }

            return field;
        }

        #endregion
    }
}
