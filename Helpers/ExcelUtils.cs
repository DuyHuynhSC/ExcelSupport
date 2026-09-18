using System;

namespace ExcelSupport.Helpers
{
    /// <summary>
    /// Các hàm tiện ích toán học và chuyển đổi dùng chung cho Excel Range, Column, Row
    /// </summary>
    public static class ExcelUtils
    {
        /// <summary>
        /// Chuyển đổi chỉ số cột số (1-based: 1, 2, ..., 26, 27) sang chữ cái cột Excel (A, B, ..., Z, AA)
        /// </summary>
        public static string ConvertColIndexToLetter(int colIndex)
        {
            if (colIndex <= 0) return string.Empty;

            string colLetter = string.Empty;
            while (colIndex > 0)
            {
                int modulo = (colIndex - 1) % 26;
                colLetter = Convert.ToChar('A' + modulo) + colLetter;
                colIndex = (colIndex - modulo) / 26;
            }
            return colLetter;
        }
    }
}
