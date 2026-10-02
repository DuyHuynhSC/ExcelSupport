using System;
using System.Collections.Generic;
using Newtonsoft.Json;

namespace ExcelSupport.Models
{
    /// <summary>
    /// Các kiểu kẻ viền ô bảng tính chuẩn cho tài liệu kỹ thuật Nhật Bản
    /// </summary>
    public enum PresetBorderStyle
    {
        /// <summary>
        /// Không kẻ viền (xóa viền hiện tại)
        /// </summary>
        None = 0,

        /// <summary>
        /// Kẻ viền mỏng toàn bộ các ô (All Thin Borders) - chuẩn thông dụng nhất
        /// </summary>
        AllThin = 1,

        /// <summary>
        /// Viền ngoài đậm (Medium/Thick), viền trong mỏng (Outline Box)
        /// </summary>
        OutlineMedium = 2,

        /// <summary>
        /// Viền dưới đôi (Double Bottom Border) - thường dùng cho Header hoặc dòng Tổng kết
        /// </summary>
        HeaderBottomDouble = 3,

        /// <summary>
        /// Viền nét chấm đứt (Dotted Borders)
        /// </summary>
        Dotted = 4,

        /// <summary>
        /// Giữ nguyên trạng thái viền ô hiện tại
        /// </summary>
        KeepCurrent = 5
    }

    /// <summary>
    /// Căn lề ngang
    /// </summary>
    public enum PresetHorizontalAlign
    {
        General = 0,
        Left = 1,
        Center = 2,
        Right = 3
    }

    /// <summary>
    /// Căn lề dọc
    /// </summary>
    public enum PresetVerticalAlign
    {
        Top = 0,
        Center = 1,
        Bottom = 2
    }

    /// <summary>
    /// Mẫu định dạng ô tính thiết kế Nhật Bản
    /// </summary>
    public class JapaneseFormatPreset
    {
        public string Id { get; set; } = Guid.NewGuid().ToString();

        /// <summary>
        /// Tên preset hiển thị trên UI
        /// </summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>
        /// Mô tả mục đích sử dụng
        /// </summary>
        public string Description { get; set; } = string.Empty;

        /// <summary>
        /// Đánh dấu mẫu định dạng hệ thống tích hợp sẵn
        /// </summary>
        public bool IsBuiltIn { get; set; }

        // ==================== Font & Typography ====================
        public string FontName { get; set; } = "Tahoma";
        public double FontSize { get; set; } = 10.0;
        public string FontColorHex { get; set; } = "#0F172A";
        public bool IsBold { get; set; }
        public bool IsItalic { get; set; }
        public bool IsUnderline { get; set; }

        // ==================== Fill / Background ====================
        public bool HasFill { get; set; }
        public string FillColorHex { get; set; } = "#FFFFFF";

        // ==================== Alignment ====================
        public PresetHorizontalAlign HorizontalAlign { get; set; } = PresetHorizontalAlign.Left;
        public PresetVerticalAlign VerticalAlign { get; set; } = PresetVerticalAlign.Center;
        public bool WrapText { get; set; }

        // ==================== Borders ====================
        public PresetBorderStyle BorderStyle { get; set; } = PresetBorderStyle.AllThin;
        public string BorderColorHex { get; set; } = "#CBD5E1";

        // ==================== Number Formatting ====================
        public bool ApplyNumberFormat { get; set; }
        public string NumberFormat { get; set; } = string.Empty;

        /// <summary>
        /// Trạng thái đang kích hoạt làm mặc định (dùng cho UI)
        /// </summary>
        [JsonIgnore]
        public bool IsDefault { get; set; }

        public JapaneseFormatPreset Clone()
        {
            return new JapaneseFormatPreset
            {
                Id = this.Id,
                Name = this.Name,
                Description = this.Description,
                IsBuiltIn = this.IsBuiltIn,
                FontName = this.FontName,
                FontSize = this.FontSize,
                FontColorHex = this.FontColorHex,
                IsBold = this.IsBold,
                IsItalic = this.IsItalic,
                IsUnderline = this.IsUnderline,
                HasFill = this.HasFill,
                FillColorHex = this.FillColorHex,
                HorizontalAlign = this.HorizontalAlign,
                VerticalAlign = this.VerticalAlign,
                WrapText = this.WrapText,
                BorderStyle = this.BorderStyle,
                BorderColorHex = this.BorderColorHex,
                ApplyNumberFormat = this.ApplyNumberFormat,
                NumberFormat = this.NumberFormat,
                IsDefault = this.IsDefault
            };
        }

        /// <summary>
        /// Tạo bản sao mới (dùng khi người dùng nhấn nút Nhân bản / Clone)
        /// </summary>
        public JapaneseFormatPreset Duplicate()
        {
            var copy = this.Clone();
            copy.Id = Guid.NewGuid().ToString();
            string cleanName = System.Text.RegularExpressions.Regex.Replace(this.Name ?? string.Empty, @"(\s*\([Cc]opy\))+$", "").Trim();
            copy.Name = (string.IsNullOrWhiteSpace(cleanName) ? "Preset" : cleanName) + " (Copy)";
            copy.IsBuiltIn = false;
            copy.IsDefault = false;
            return copy;
        }

        /// <summary>
        /// Tạo danh sách 8 mẫu định dạng chuẩn phổ biến trong các dự án tài liệu Nhật Bản
        /// </summary>
        public static List<JapaneseFormatPreset> CreateDefaultPresets()
        {
            return new List<JapaneseFormatPreset>
            {
                // 1. Header Bảng TKCT (Navy Blue Header / テーブルヘッダー)
                new JapaneseFormatPreset
                {
                    Id = "builtin_table_header",
                    Name = "Header Bảng TKCT (Navy / White)",
                    Description = "Tiêu đề cột bảng thiết kế chi tiết (Chi tiết màn hình, danh sách trường DB, API)",
                    IsBuiltIn = true,
                    FontName = "Meiryo UI",
                    FontSize = 10.0,
                    FontColorHex = "#FFFFFF",
                    IsBold = true,
                    HasFill = true,
                    FillColorHex = "#1B365D",
                    HorizontalAlign = PresetHorizontalAlign.Center,
                    VerticalAlign = PresetVerticalAlign.Center,
                    WrapText = true,
                    BorderStyle = PresetBorderStyle.AllThin,
                    BorderColorHex = "#CBD5E1"
                },

                // 2. Sub-Header Bảng / Cột Phụ (Ice Blue / サブヘッダー)
                new JapaneseFormatPreset
                {
                    Id = "builtin_sub_header",
                    Name = "Sub-Header Bảng (Ice Blue / Nhạt)",
                    Description = "Nhóm tiêu đề phân cấp 2 hoặc cột phụ phân loại trong bảng",
                    IsBuiltIn = true,
                    FontName = "Meiryo UI",
                    FontSize = 9.5,
                    FontColorHex = "#1E293B",
                    IsBold = true,
                    HasFill = true,
                    FillColorHex = "#D9E1F2",
                    HorizontalAlign = PresetHorizontalAlign.Center,
                    VerticalAlign = PresetVerticalAlign.Center,
                    WrapText = true,
                    BorderStyle = PresetBorderStyle.AllThin,
                    BorderColorHex = "#94A3B8"
                },

                // 3. Mục Bắt Buộc / Cần Nhập (Required Item / 必須項目 [Must])
                new JapaneseFormatPreset
                {
                    Id = "builtin_required_must",
                    Name = "Mục Bắt Buộc [Must] (Red / Peach)",
                    Description = "Đánh dấu trường bắt buộc nhập, Not-Null, hoặc điều kiện tiên quyết trong spec",
                    IsBuiltIn = true,
                    FontName = "Meiryo UI",
                    FontSize = 9.5,
                    FontColorHex = "#9C0006",
                    IsBold = true,
                    HasFill = true,
                    FillColorHex = "#FFC7CE",
                    HorizontalAlign = PresetHorizontalAlign.Center,
                    VerticalAlign = PresetVerticalAlign.Center,
                    WrapText = false,
                    BorderStyle = PresetBorderStyle.AllThin,
                    BorderColorHex = "#FCA5A5"
                },

                // 4. Mã Kỹ Thuật / Tên Vật Lý (Code / Physical Name / 物理名)
                new JapaneseFormatPreset
                {
                    Id = "builtin_code_physical",
                    Name = "Mã Kỹ Thuật / Physical Name (Consolas)",
                    Description = "Tên bảng vật lý (M_USER), tên cột DB (USER_ID), Class Name, API endpoint",
                    IsBuiltIn = true,
                    FontName = "Consolas",
                    FontSize = 9.5,
                    FontColorHex = "#0F172A",
                    IsBold = false,
                    HasFill = true,
                    FillColorHex = "#F1F5F9",
                    HorizontalAlign = PresetHorizontalAlign.Left,
                    VerticalAlign = PresetVerticalAlign.Center,
                    WrapText = false,
                    BorderStyle = PresetBorderStyle.AllThin,
                    BorderColorHex = "#CBD5E1",
                    ApplyNumberFormat = true,
                    NumberFormat = "@"
                },

                // 5. Khối Ghi Chú / Lưu Ý (Note / Warning / 備考・注意)
                new JapaneseFormatPreset
                {
                    Id = "builtin_note_warning",
                    Name = "Ghi Chú / Lưu Ý (Amber / Light Yellow)",
                    Description = "Khối chú thích logic nghiệp vụ, cảnh báo ngoại lệ, hoặc điều kiện biên",
                    IsBuiltIn = true,
                    FontName = "Meiryo UI",
                    FontSize = 9.0,
                    FontColorHex = "#78350F",
                    IsBold = false,
                    HasFill = true,
                    FillColorHex = "#FEF3C7",
                    HorizontalAlign = PresetHorizontalAlign.Left,
                    VerticalAlign = PresetVerticalAlign.Center,
                    WrapText = true,
                    BorderStyle = PresetBorderStyle.AllThin,
                    BorderColorHex = "#FDE68A"
                },

                // 6. Ô Dữ Liệu Nội Dung Chuẩn (Standard Detail / 標準明細セル)
                new JapaneseFormatPreset
                {
                    Id = "builtin_standard_data",
                    Name = "Ô Nội Dung Chuẩn (Meiryo UI 9.5pt)",
                    Description = "Vùng dữ liệu dòng chuẩn trong bảng tài liệu thiết kế",
                    IsBuiltIn = true,
                    FontName = "Meiryo UI",
                    FontSize = 9.5,
                    FontColorHex = "#1E293B",
                    IsBold = false,
                    HasFill = false,
                    FillColorHex = "#FFFFFF",
                    HorizontalAlign = PresetHorizontalAlign.Left,
                    VerticalAlign = PresetVerticalAlign.Center,
                    WrapText = false,
                    BorderStyle = PresetBorderStyle.AllThin,
                    BorderColorHex = "#E2E8F0"
                },

                // 7. Ngày Tháng Chuẩn Nhật (Japanese Date / 日付)
                new JapaneseFormatPreset
                {
                    Id = "builtin_date_standard",
                    Name = "Ngày Tháng Chuẩn Nhật (yyyy/mm/dd)",
                    Description = "Định dạng ngày lập tài liệu, ngày nghiệm thu theo chuẩn YYYY/MM/DD",
                    IsBuiltIn = true,
                    FontName = "Meiryo UI",
                    FontSize = 9.5,
                    FontColorHex = "#1E293B",
                    IsBold = false,
                    HasFill = false,
                    FillColorHex = "#FFFFFF",
                    HorizontalAlign = PresetHorizontalAlign.Center,
                    VerticalAlign = PresetVerticalAlign.Center,
                    WrapText = false,
                    BorderStyle = PresetBorderStyle.AllThin,
                    BorderColorHex = "#E2E8F0",
                    ApplyNumberFormat = true,
                    NumberFormat = "yyyy/mm/dd"
                },

                // 8. Số Tiền / Số Lượng (Numeric / Currency / 金額・数値)
                new JapaneseFormatPreset
                {
                    Id = "builtin_number_currency",
                    Name = "Số Tiền / Định Lượng (#,##0)",
                    Description = "Căn lề phải, định dạng phân cách hàng nghìn cho số lượng, dung lượng, đơn giá",
                    IsBuiltIn = true,
                    FontName = "Meiryo UI",
                    FontSize = 9.5,
                    FontColorHex = "#1E293B",
                    IsBold = false,
                    HasFill = false,
                    FillColorHex = "#FFFFFF",
                    HorizontalAlign = PresetHorizontalAlign.Right,
                    VerticalAlign = PresetVerticalAlign.Center,
                    WrapText = false,
                    BorderStyle = PresetBorderStyle.AllThin,
                    BorderColorHex = "#E2E8F0",
                    ApplyNumberFormat = true,
                    NumberFormat = "#,##0"
                }
            };
        }
    }

    /// <summary>
    /// Cấu hình lưu trữ toàn bộ danh sách presets của người dùng
    /// </summary>
    public class JapanesePresetConfig
    {
        public int Version { get; set; } = 1;
        public string ActivePresetId { get; set; } = "builtin_table_header";
        public List<JapaneseFormatPreset> Presets { get; set; } = new List<JapaneseFormatPreset>();
    }
}
