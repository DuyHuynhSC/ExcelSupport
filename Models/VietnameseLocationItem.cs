using System;

namespace ExcelSupport.Models
{
    public enum VietnameseLocationType
    {
        Cell,
        SheetName,
        Comment
    }

    public class VietnameseLocationItem
    {
        public int Index { get; set; }
        public string WorkbookName { get; set; } = string.Empty;
        public string SheetName { get; set; } = string.Empty;
        public string CellAddress { get; set; } = string.Empty;
        public string TextContent { get; set; } = string.Empty;
        public VietnameseLocationType Type { get; set; } = VietnameseLocationType.Cell;

        public string TypeDescription => Type switch
        {
            VietnameseLocationType.Cell => "Ô tính",
            VietnameseLocationType.SheetName => "Tên Sheet",
            VietnameseLocationType.Comment => "Ghi chú / Comment",
            _ => "Khác"
        };
    }
}
