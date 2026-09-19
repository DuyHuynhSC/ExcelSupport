using System;
using System.Collections.Generic;

namespace ExcelSupport.Models
{
    public enum FilteredPasteType
    {
        ValuesOnly,     // Chỉ dán giá trị
        Formulas,       // Dán công thức
        FormatsOnly,    // Chỉ dán định dạng
        All             // Dán toàn bộ (Giá trị & Định dạng)
    }

    public class FilteredPasteOptions
    {
        public FilteredPasteType PasteType { get; set; } = FilteredPasteType.ValuesOnly;
        public bool RepeatIfShorter { get; set; } = false; // Lặp lại dữ liệu nếu vùng đích nhiều dòng hơn vùng nguồn
        public bool SkipBlanks { get; set; } = false;      // Bỏ qua các ô trống trong nguồn
        public string SourceAddress { get; set; } = string.Empty;
        public string TargetAddress { get; set; } = string.Empty;
    }

    public class FilteredPasteResult
    {
        public bool Success { get; set; }
        public string Message { get; set; } = string.Empty;
        public int SourceRowCount { get; set; }
        public int TargetVisibleRowCount { get; set; }
        public int RowsPasted { get; set; }
        public int HiddenRowsProtected { get; set; }
    }

    public class VisibleCellBlock
    {
        public int RowIndex { get; set; }
        public int ColIndex { get; set; }
        public object? Value { get; set; }
        public string? Formula { get; set; }
        public string? NumberFormat { get; set; }
    }
}
