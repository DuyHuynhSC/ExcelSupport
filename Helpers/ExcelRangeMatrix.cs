using System;
using Microsoft.Office.Interop.Excel;

namespace ExcelSupport.Helpers
{
    /// <summary>
    /// Deep Module đại diện cho bộ đệm ma trận ô 2D in-memory (0-based) của một Excel Range.
    /// Giấu kín toàn bộ sự phức tạp về:
    /// 1. Xử lý ô đơn lẻ (1x1) vs vùng ô (NxM)
    /// 2. Chuyển đổi 1-based index (Excel COM) sang 0-based index chuẩn C#
    /// 3. Cập nhật ghi ngược (WriteBack) tốc độ cao qua mảng 2 chiều
    /// 4. An toàn giải phóng đối tượng COM trung gian
    /// </summary>
    public sealed class ExcelRangeMatrix
    {
        private readonly object?[,] _data;
        public int RowCount { get; }
        public int ColumnCount { get; }
        public int StartRow { get; }
        public int StartColumn { get; }

        public bool IsEmpty => RowCount == 0 || ColumnCount == 0;

        public object?[,] RawData => _data;

        public object? this[int row, int col]
        {
            get => _data[row, col];
            set => _data[row, col] = value;
        }

        private ExcelRangeMatrix(object?[,] data, int rowCount, int colCount, int startRow, int startCol)
        {
            _data = data;
            RowCount = rowCount;
            ColumnCount = colCount;
            StartRow = startRow;
            StartColumn = startCol;
        }

        public static ExcelRangeMatrix Empty => new ExcelRangeMatrix(new object?[0, 0], 0, 0, 1, 1);

        public static ExcelRangeMatrix Load(Range? range, bool loadFormulas = false)
        {
            if (range == null) return Empty;

            try
            {
                int rowCount = range.Rows.Count;
                int colCount = range.Columns.Count;
                int startRow = range.Row;
                int startCol = range.Column;

                if (rowCount <= 0 || colCount <= 0) return Empty;

                object? raw = loadFormulas ? range.Formula : range.Value2;
                if (raw == null && rowCount == 1 && colCount == 1)
                {
                    return new ExcelRangeMatrix(new object?[1, 1], 1, 1, startRow, startCol);
                }

                object?[,] buffer = new object?[rowCount, colCount];

                if (raw is object[,] arr2D)
                {
                    int rLower = arr2D.GetLowerBound(0);
                    int cLower = arr2D.GetLowerBound(1);

                    for (int r = 0; r < rowCount; r++)
                    {
                        for (int c = 0; c < colCount; c++)
                        {
                            buffer[r, c] = arr2D[rLower + r, cLower + c];
                        }
                    }
                }
                else
                {
                    buffer[0, 0] = raw;
                }

                return new ExcelRangeMatrix(buffer, rowCount, colCount, startRow, startCol);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ExcelRangeMatrix] Load error: {ex.Message}");
                return Empty;
            }
        }

        public bool WriteBack(Range? range, bool writeAsFormulas = false)
        {
            if (range == null || IsEmpty) return false;

            try
            {
                if (RowCount == 1 && ColumnCount == 1)
                {
                    if (writeAsFormulas)
                        range.Formula = _data[0, 0];
                    else
                        range.Value2 = _data[0, 0];
                    return true;
                }

                object[,] comArr = (object[,])Array.CreateInstance(typeof(object), new int[] { RowCount, ColumnCount }, new int[] { 1, 1 });
                for (int r = 0; r < RowCount; r++)
                {
                    for (int c = 0; c < ColumnCount; c++)
                    {
                        comArr[r + 1, c + 1] = _data[r, c]!;
                    }
                }

                if (writeAsFormulas)
                    range.Formula = comArr;
                else
                    range.Value2 = comArr;

                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ExcelRangeMatrix] WriteBack error: {ex.Message}");
                return false;
            }
        }

        public string GetString(int row, int col, bool trim = true)
        {
            if (row < 0 || row >= RowCount || col < 0 || col >= ColumnCount) return string.Empty;
            object? val = _data[row, col];
            if (val == null) return string.Empty;
            string str = val.ToString() ?? string.Empty;
            return trim ? str.Trim() : str;
        }

        public int TransformText(Func<string, string> transformer)
        {
            if (IsEmpty || transformer == null) return 0;

            int changedCount = 0;
            for (int r = 0; r < RowCount; r++)
            {
                for (int c = 0; c < ColumnCount; c++)
                {
                    object? cellVal = _data[r, c];
                    if (cellVal is string str && !string.IsNullOrEmpty(str))
                    {
                        string transformed = transformer(str);
                        if (!string.Equals(str, transformed, StringComparison.Ordinal))
                        {
                            _data[r, c] = transformed;
                            changedCount++;
                        }
                    }
                }
            }
            return changedCount;
        }
    }
}
