namespace ExcelSupport.Models
{
    public enum CompareMode
    {
        CellByCell, // So sánh theo tọa độ ô tuyệt đối A1, B2...
        LcsRows,    // Thuật toán LCS theo Dòng (Tự động bắt dòng chèn/xóa)
        LcsColumns, // Thuật toán LCS theo Cột (Tự động bắt cột chèn/xóa)
        Lcs2D,      // Thuật toán LCS 2 Chiều (Cả Dòng và Cột)
        KeyColumn   // So sánh theo Cột Khóa chính (ID, Mã NV...)
    }

    public class CompareOptions
    {
        public CompareMode Mode { get; set; } = CompareMode.LcsRows; // Mặc định LCS theo Dòng thông minh
        public int KeyColumnIndex { get; set; } = 1; // Cột khóa chính (1-based, vd 1 = Cột A)
        public bool IgnoreWhitespace { get; set; } = true;
        public bool CaseInsensitive { get; set; } = false;
        public bool CompareFormulas { get; set; } = false; // So sánh công thức thay vì giá trị
        public bool IgnoreCase { get; set; } = false;
    }
}
