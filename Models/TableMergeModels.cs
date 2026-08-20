using System;
using System.Collections.Generic;

namespace ExcelSupport.Models
{
    public enum TableJoinType
    {
        LeftJoin,       // Giữ tất cả dòng Bảng 1 + Ghép thông tin từ Bảng 2
        InnerJoin,      // Chỉ lấy các dòng có Mã Khóa xuất hiện ở cả 2 bảng
        FullOuterJoin,  // Lấy tất cả dòng từ cả 2 bảng
        LookupColumn    // Chỉ lấy 1 hoặc vài cột từ Bảng 2 chèn tiếp nối vào Bảng 1
    }

    public enum TableMergeOutputTarget
    {
        InsertAdjacentToTable1, // Chèn các cột ghép trực tiếp vào bên phải Bảng 1
        CreateNewWorksheet      // Trích xuất kết quả gộp thành 1 Sheet mới riêng biệt
    }

    public class MergeColumnItem
    {
        public int ColumnIndex { get; set; }
        public string ColumnLetter { get; set; } = string.Empty;
        public string HeaderText { get; set; } = string.Empty;
        public string DisplayText => $"[{ColumnLetter}] {(!string.IsNullOrEmpty(HeaderText) ? HeaderText : "(Không tên)")}";
        public bool IsSelected { get; set; } = true;
        public string OutputHeaderName { get; set; } = string.Empty;
    }

    public class TableMergeOptions
    {
        public TableJoinType JoinType { get; set; } = TableJoinType.LeftJoin;
        public TableMergeOutputTarget OutputTarget { get; set; } = TableMergeOutputTarget.CreateNewWorksheet;

        // Bảng 1
        public string Table1WorkbookName { get; set; } = string.Empty;
        public string Table1SheetName { get; set; } = string.Empty;
        public int Table1KeyColIndex { get; set; } = 1;
        public int Table1HeaderRow { get; set; } = 1;

        // Bảng 2
        public string Table2WorkbookName { get; set; } = string.Empty;
        public string Table2SheetName { get; set; } = string.Empty;
        public int Table2KeyColIndex { get; set; } = 1;
        public int Table2HeaderRow { get; set; } = 1;

        // Các cột từ Bảng 2 cần ghép
        public List<MergeColumnItem> SelectedColumnsFromTable2 { get; set; } = new List<MergeColumnItem>();

        // Tùy chọn so khớp
        public bool MatchCase { get; set; } = false;
        public bool TrimSpaces { get; set; } = true;
        public bool IgnoreAccent { get; set; } = false; // Bỏ qua dấu tiếng Việt
    }

    public class TableMergeResult
    {
        public bool Success { get; set; }
        public int TotalRowsMerged { get; set; }
        public int MatchedRows { get; set; }
        public int UnmatchedRows { get; set; }
        public string OutputSheetName { get; set; } = string.Empty;
        public string Message { get; set; } = string.Empty;
    }
}
