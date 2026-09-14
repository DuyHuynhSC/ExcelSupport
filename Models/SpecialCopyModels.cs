using System;

namespace ExcelSupport.Models
{
    public enum SpecialCopyDelimiter
    {
        Comma,          // ,
        CommaSpace,     // , 
        Semicolon,      // ;
        Pipe,           // |
        Tab,            // \t
        Space,          //  
        Newline,        // \r\n
        Custom          // Ký tự tùy chỉnh do người dùng nhập
    }

    public enum SpecialCopyJoinMode
    {
        ByRow,          // Nối theo từng dòng (mỗi dòng là 1 chuỗi, ngăn cách bởi newline - chuẩn dòng CSV)
        AllCells,       // Nối tất cả các ô thành 1 chuỗi đơn duy nhất
        ByColumn        // Nối theo từng cột
    }

    public enum SpecialCopyQuoteMode
    {
        None,               // Không bọc ngoặc
        AutoCsv,            // Tự động bọc ngoặc kép "..." nếu ô chứa dấu nối, dấu ngoặc kép hoặc xuống dòng
        AlwaysDoubleQuote,  // Luôn bọc dấu ngoặc kép "..."
        SingleQuoteSql      // Bọc dấu nháy đơn '...' (phục vụ câu lệnh SQL IN ('a', 'b'))
    }

    public class SpecialCopyOptions
    {
        public SpecialCopyDelimiter Delimiter { get; set; } = SpecialCopyDelimiter.Comma;
        public string CustomDelimiter { get; set; } = ", ";
        public SpecialCopyJoinMode JoinMode { get; set; } = SpecialCopyJoinMode.ByRow;
        public SpecialCopyQuoteMode QuoteMode { get; set; } = SpecialCopyQuoteMode.None;
        public bool SkipBlanks { get; set; } = false;
        public bool TrimSpaces { get; set; } = false;
        public bool VisibleOnly { get; set; } = true; // Kế thừa từ FilteredCopyPaste: bỏ qua dòng ẩn

        public string GetDelimiterString()
        {
            return Delimiter switch
            {
                SpecialCopyDelimiter.Comma => ",",
                SpecialCopyDelimiter.CommaSpace => ", ",
                SpecialCopyDelimiter.Semicolon => ";",
                SpecialCopyDelimiter.Pipe => "|",
                SpecialCopyDelimiter.Tab => "\t",
                SpecialCopyDelimiter.Space => " ",
                SpecialCopyDelimiter.Newline => Environment.NewLine,
                SpecialCopyDelimiter.Custom => CustomDelimiter ?? string.Empty,
                _ => ","
            };
        }
    }

    public class SpecialCopyResult
    {
        public bool Success { get; set; }
        public string Message { get; set; } = string.Empty;
        public string ResultText { get; set; } = string.Empty;
        public int TotalCellsProcessed { get; set; }
        public int TotalLines { get; set; }
    }
}
