namespace ExcelSupport.Models
{
    public enum CleaningScope
    {
        SelectedRange,  // Vùng ô đang chọn trên Excel
        ActiveSheet,    // Toàn bộ Sheet hiện tại
        ActiveWorkbook  // Toàn bộ Workbook hiện tại
    }

    public enum TextCaseOption
    {
        None,
        UpperCase,      // IN HOA TOÀN BỘ
        LowerCase,      // in thường toàn bộ
        ProperCase,     // Viết Hoa Đầu Từ (Title Case)
        SentenceCase    // Viết hoa đầu câu
    }

    public enum BlankFillOption
    {
        None,
        CustomValue,
        FillDownFromAbove,
        FillUpFromBelow
    }

    public class DataCleaningOptions
    {
        public CleaningScope Scope { get; set; } = CleaningScope.SelectedRange;

        // 1. Khoảng trắng & ký tự điều khiển
        public bool TrimSpaces { get; set; } = true;                  // Xóa khoảng trắng đầu/cuối
        public bool ReduceMultipleSpaces { get; set; } = true;        // Thu gọn nhiều khoảng trắng liên tiếp
        public bool RemoveNonBreakingSpaces { get; set; } = true;    // Xóa khoảng trắng không ngắt (nbsp / \u00A0)
        public bool RemoveLineBreaks { get; set; } = false;           // Xóa ký tự xuống dòng (\r, \n)
        public bool RemoveUnprintableChars { get; set; } = false;     // Xóa ký tự ẩn không in được (ASCII 0-31)

        // 2. Chữ HOA / thường
        public TextCaseOption CaseOption { get; set; } = TextCaseOption.None;

        // 3. Ngôn ngữ & Ký tự chuyên dụng
        public bool RemoveVietnameseDiacritics { get; set; } = false; // Bỏ dấu tiếng Việt
        public bool ConvertVietnameseToKatakana { get; set; } = false;// Chuyển tên tiếng Việt sang Katakana Nhật Bản
        public bool KatakanaUseMiddleDot { get; set; } = true;        // Dùng dấu chấm giữa (・) cho Katakana
        public bool JapaneseHalfWidthToFullWidth { get; set; } = false;// Hankaku -> Zenkaku
        public bool JapaneseFullWidthToHalfWidth { get; set; } = false;// Zenkaku -> Hankaku
        public bool RemoveDigits { get; set; } = false;               // Xóa chữ số 0-9
        public bool RemoveLetters { get; set; } = false;              // Xóa chữ cái
        public bool RemoveSpecialSymbols { get; set; } = false;       // Chỉ giữ chữ và số

        // 4. Số & Ngày tháng
        public bool ConvertNumbersStoredAsText { get; set; } = false; // Sửa số lưu dạng text thành số thực
        public bool StandardizeDates { get; set; } = false;           // Chuẩn hóa ngày tháng
        public string DateFormat { get; set; } = "yyyy-MM-dd";        // Định dạng ngày đích

        // 5. Ô trống & Mã lỗi
        public BlankFillOption FillBlanks { get; set; } = BlankFillOption.None;
        public string CustomBlankValue { get; set; } = "N/A";
        public bool ReplaceErrorValues { get; set; } = false;
        public string CustomErrorReplacement { get; set; } = string.Empty;
    }
}
