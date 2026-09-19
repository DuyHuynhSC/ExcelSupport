using System;
using System.Collections.Generic;

namespace ExcelSupport.Models
{
    public class SheetCellSnapshot
    {
        public int Row { get; set; }
        public int Column { get; set; }
        public string? Formula { get; set; }
        public object? Value { get; set; }
        public string? NumberFormat { get; set; }
    }

    public class SheetSnapshotItem
    {
        public string Id { get; set; } = Guid.NewGuid().ToString("N");
        public DateTime Timestamp { get; set; } = DateTime.Now;
        public string WorkbookName { get; set; } = string.Empty;
        public string SheetName { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public int StartRow { get; set; }
        public int StartColumn { get; set; }
        public int RowCount { get; set; }
        public int ColumnCount { get; set; }
        public bool IsAutoSnapshot { get; set; }
        
        // 2D Array storage for fast bulk COM reading/writing
        public object[,] Values { get; set; } = new object[0, 0];
        public object[,] Formulas { get; set; } = new object[0, 0];
        public object[,] NumberFormats { get; set; } = new object[0, 0];

        // Column widths map
        public Dictionary<int, double> ColumnWidths { get; set; } = new Dictionary<int, double>();

        public string DisplayTimestamp => Timestamp.ToString("HH:mm:ss dd/MM/yyyy");
        public string DisplaySummary => $"{SheetName} ({RowCount:N0} × {ColumnCount:N0} ô) - {(IsAutoSnapshot ? "⚡ Tự động" : "📸 Thủ công")}";
    }

    public enum SnapshotDiffType
    {
        Identical,
        ValueChanged,
        FormulaChanged,
        AddedInSheet,
        MissingInSheet
    }

    public class SnapshotCellDiff
    {
        public string CellAddress { get; set; } = string.Empty;
        public int Row { get; set; }
        public int Column { get; set; }
        public string SnapshotValue { get; set; } = string.Empty;
        public string CurrentValue { get; set; } = string.Empty;
        public string SnapshotFormula { get; set; } = string.Empty;
        public string CurrentFormula { get; set; } = string.Empty;
        public SnapshotDiffType DiffType { get; set; }
        public string DiffTypeDisplay => DiffType switch
        {
            SnapshotDiffType.ValueChanged => "Giá trị thay đổi",
            SnapshotDiffType.FormulaChanged => "Công thức thay đổi",
            SnapshotDiffType.AddedInSheet => "Thêm mới trên Sheet",
            SnapshotDiffType.MissingInSheet => "Đã bị xóa trên Sheet",
            _ => "Trùng khớp"
        };
    }
}
