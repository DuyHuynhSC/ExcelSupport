using System;
using System.Collections.Generic;

namespace ExcelSupport.Models
{
    public enum FuzzyMatchAlgorithm
    {
        Levenshtein,   // Khoảng cách biên tập (Levenshtein Distance)
        JaroWinkler,   // Tương đồng Jaro-Winkler (tối ưu cho họ tên & từ ngữ)
        PhoneticVietnamese // Bỏ dấu & chuẩn hóa cấu trúc âm tiết tiếng Việt
    }

    public class FuzzyRecordItem
    {
        public int RowIndex { get; set; }
        public string CellAddress { get; set; } = string.Empty;
        public string OriginalText { get; set; } = string.Empty;
        public string NormalizedText { get; set; } = string.Empty;
        public double SimilarityPercent { get; set; }
        public string SimilarityBadge => $"{SimilarityPercent:0.#}%";
        public bool IsSelected { get; set; } = true;
    }

    public class FuzzyClusterGroup
    {
        public int GroupId { get; set; }
        public string MasterValue { get; set; } = string.Empty;
        public string GroupTitle => $"Nhóm #{GroupId}: \"{MasterValue}\" ({Items.Count} biến thể)";
        public List<FuzzyRecordItem> Items { get; set; } = new List<FuzzyRecordItem>();
        public int Count => Items.Count;
    }

    public class FuzzyScanOptions
    {
        public int TargetColumnIndex { get; set; } = 1;
        public int StartRow { get; set; } = 2; // Bỏ qua tiêu đề
        public double SimilarityThreshold { get; set; } = 80.0; // % tương đồng (e.g. 80%)
        public FuzzyMatchAlgorithm Algorithm { get; set; } = FuzzyMatchAlgorithm.JaroWinkler;
        public bool IgnoreCase { get; set; } = true;
        public bool IgnoreAccent { get; set; } = true; // Bỏ dấu tiếng Việt
        public bool CleanInvisibleSpaces { get; set; } = true; // Dọn dẹp khoảng trắng vô hình NBSP \u00A0, zero-width
    }
}
