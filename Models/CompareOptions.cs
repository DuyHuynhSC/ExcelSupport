namespace ExcelSupport.Models
{
    public enum CompareMode
    {
        CellByCell, // So sánh theo tọa độ ô A1, B2...
        KeyColumn   // So sánh theo cột khóa chính (ID, Mã NV...)
    }

    public class CompareOptions
    {
        public CompareMode Mode { get; set; } = CompareMode.CellByCell;
        public int KeyColumnIndex { get; set; } = 1; // Cột khóa chính (1-based, vd 1 = Cột A)
        public bool IgnoreWhitespace { get; set; } = true;
        public bool CaseInsensitive { get; set; } = false;
        public bool CompareFormulas { get; set; } = false; // So sánh công thức thay vì giá trị
        public bool IgnoreCase { get; set; } = false;
    }
}
