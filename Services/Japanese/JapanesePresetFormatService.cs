using System;
using System.Drawing;
using System.Runtime.InteropServices;
using ExcelSupport.Helpers;
using ExcelSupport.Models;
using Microsoft.Office.Interop.Excel;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;

namespace ExcelSupport.Services
{
    public class JapaneseFormatResult
    {
        public bool Success { get; set; }
        public string Message { get; set; } = string.Empty;
        public long FormattedCellsCount { get; set; }
        public string PresetName { get; set; } = string.Empty;
    }

    /// <summary>
    /// Service áp dụng nhanh định dạng phong cách tài liệu Nhật Bản lên Range trong Excel
    /// </summary>
    public static class JapanesePresetFormatService
    {
        /// <summary>
        /// Áp dụng Preset đang chọn cho vùng Selection hiện tại trong Excel
        /// </summary>
        public static JapaneseFormatResult ApplyPresetToSelection(ExcelApp? app, JapaneseFormatPreset? preset)
        {
            if (app == null)
            {
                return new JapaneseFormatResult
                {
                    Success = false,
                    Message = LocalizationService.Get("JpFormat_ErrNoExcelApp", "Không tìm thấy tiến trình Microsoft Excel đang hoạt động.")
                };
            }

            if (preset == null)
            {
                return new JapaneseFormatResult
                {
                    Success = false,
                    Message = LocalizationService.Get("JpFormat_ErrNoPreset", "Chưa chọn mẫu định dạng (Preset) để áp dụng.")
                };
            }

            Range? selection = null;
            try
            {
                selection = app.Selection as Range;
            }
            catch { }

            if (selection == null)
            {
                return new JapaneseFormatResult
                {
                    Success = false,
                    Message = LocalizationService.Get("JpFormat_ErrNoSelection", "Vui lòng chọn ít nhất một ô trên bảng tính để áp dụng định dạng.")
                };
            }

            return ApplyPresetToRange(app, selection, preset);
        }

        /// <summary>
        /// Áp dụng định dạng Preset lên đối tượng Range cụ thể (hỗ trợ hoàn tác Ctrl+Z)
        /// </summary>
        public static JapaneseFormatResult ApplyPresetToRange(ExcelApp app, Range targetRange, JapaneseFormatPreset preset)
        {
            if (app == null || targetRange == null || preset == null)
            {
                return new JapaneseFormatResult { Success = false };
            }

            // 1. Sao lưu định dạng hiện tại của vùng chọn để hỗ trợ Ctrl + Z (Undo)
            List<CellFormatSnapshot>? backup = null;
            _Worksheet? ws = null;
            string rangeAddress = string.Empty;
            string? wbName = null;

            try
            {
                ws = targetRange.Worksheet;
                var wb = ws?.Parent as Workbook;
                wbName = wb?.Name;
                rangeAddress = targetRange.Address[true, true];

                long totalCells = 0;
                try { totalCells = targetRange.Cells.CountLarge; }
                catch { totalCells = targetRange.Cells.Count; }

                Range rangeToBackup = targetRange;
                if (totalCells > 5000 && ws?.UsedRange != null)
                {
                    try
                    {
                        var intersect = app.Intersect(targetRange, ws.UsedRange);
                        if (intersect != null)
                        {
                            rangeToBackup = intersect;
                        }
                    }
                    catch { }
                }

                backup = ExcelUndoHelper.CaptureRangeFormatting(rangeToBackup, 5000);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[JapanesePresetFormatService] Backup for undo error: {ex.Message}");
            }

            // 2. Thực hiện định dạng trực tiếp
            var result = ApplyPresetDirect(app, targetRange, preset);

            // 3. Nếu thành công và có snapshot, ghi nhận vào Undo stack
            if (result.Success && backup != null && backup.Count > 0 && ws != null)
            {
                try
                {
                    string actionTitle = LocalizationService.Get("JpFormat_UndoTitle", preset.Name);
                    ExcelUndoHelper.RecordAction(new JapaneseFormatUndoAction
                    {
                        SheetName = ws.Name,
                        WorkbookName = wbName,
                        RangeAddress = rangeAddress,
                        Preset = preset,
                        ActionName = actionTitle,
                        Snapshots = backup
                    });
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[JapanesePresetFormatService] Record undo error: {ex.Message}");
                }
            }

            return result;
        }

        /// <summary>
        /// Áp dụng định dạng Preset trực tiếp lên Range mà không ghi đè Undo stack (dùng cho Redo và thực thi)
        /// </summary>
        public static JapaneseFormatResult ApplyPresetDirect(ExcelApp app, Range targetRange, JapaneseFormatPreset preset)
        {
            bool oldScreenUpdating = app.ScreenUpdating;
            bool oldEnableEvents = app.EnableEvents;
            XlCalculation oldCalculation = app.Calculation;

            long totalCells = 0;

            try
            {
                app.ScreenUpdating = false;
                app.EnableEvents = false;
                if (app.Calculation != XlCalculation.xlCalculationManual)
                {
                    app.Calculation = XlCalculation.xlCalculationManual;
                }

                try
                {
                    totalCells = targetRange.Cells.CountLarge;
                }
                catch
                {
                    totalCells = targetRange.Cells.Count;
                }

                int fontColorOle = ColorTranslator.ToOle(ParseHexColor(preset.FontColorHex, Color.FromArgb(15, 23, 42)));
                int fillColorOle = ColorTranslator.ToOle(ParseHexColor(preset.FillColorHex, Color.White));
                int borderColorOle = ColorTranslator.ToOle(ParseHexColor(preset.BorderColorHex, Color.FromArgb(203, 213, 225)));

                var areas = targetRange.Areas;
                int areaCount = areas.Count;

                for (int i = 1; i <= areaCount; i++)
                {
                    Range area = areas[i];

                    // 1. Áp dụng Font
                    area.Font.Name = string.IsNullOrWhiteSpace(preset.FontName) ? "Meiryo UI" : preset.FontName;
                    area.Font.Size = preset.FontSize > 0 ? preset.FontSize : 10.0;
                    area.Font.Bold = preset.IsBold;
                    area.Font.Italic = preset.IsItalic;
                    area.Font.Underline = preset.IsUnderline ? XlUnderlineStyle.xlUnderlineStyleSingle : XlUnderlineStyle.xlUnderlineStyleNone;
                    area.Font.Color = fontColorOle;

                    // 2. Áp dụng Màu nền (Interior)
                    if (preset.HasFill)
                    {
                        area.Interior.Color = fillColorOle;
                    }
                    else
                    {
                        area.Interior.ColorIndex = XlColorIndex.xlColorIndexNone;
                    }

                    // 3. Áp dụng Căn lề & Tự ngắt dòng (Alignment)
                    switch (preset.HorizontalAlign)
                    {
                        case PresetHorizontalAlign.Left:
                            area.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                            break;
                        case PresetHorizontalAlign.Center:
                            area.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                            break;
                        case PresetHorizontalAlign.Right:
                            area.HorizontalAlignment = XlHAlign.xlHAlignRight;
                            break;
                        default:
                            area.HorizontalAlignment = XlHAlign.xlHAlignGeneral;
                            break;
                    }

                    switch (preset.VerticalAlign)
                    {
                        case PresetVerticalAlign.Top:
                            area.VerticalAlignment = XlVAlign.xlVAlignTop;
                            break;
                        case PresetVerticalAlign.Center:
                            area.VerticalAlignment = XlVAlign.xlVAlignCenter;
                            break;
                        case PresetVerticalAlign.Bottom:
                            area.VerticalAlignment = XlVAlign.xlVAlignBottom;
                            break;
                    }

                    area.WrapText = preset.WrapText;

                    // 4. Áp dụng Kẻ viền (Borders)
                    ApplyBorders(area, preset.BorderStyle, borderColorOle);

                    // 5. Áp dụng Định dạng số (Number Format)
                    if (preset.ApplyNumberFormat && !string.IsNullOrEmpty(preset.NumberFormat))
                    {
                        area.NumberFormat = preset.NumberFormat;
                    }
                }

                return new JapaneseFormatResult
                {
                    Success = true,
                    PresetName = preset.Name,
                    FormattedCellsCount = totalCells,
                    Message = LocalizationService.Get("JpFormat_SuccessMsg", preset.Name, totalCells)
                };
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[JapanesePresetFormatService] Apply error: {ex.Message}");
                return new JapaneseFormatResult
                {
                    Success = false,
                    Message = LocalizationService.Get("JpFormat_ErrorMsg", ex.Message)
                };
            }
            finally
            {
                try
                {
                    app.Calculation = oldCalculation;
                    app.EnableEvents = oldEnableEvents;
                    app.ScreenUpdating = oldScreenUpdating;
                }
                catch { }
            }
        }

        private static void ApplyBorders(Range range, PresetBorderStyle style, int borderColorOle)
        {
            switch (style)
            {
                case PresetBorderStyle.None:
                    range.Borders.LineStyle = XlLineStyle.xlLineStyleNone;
                    break;

                case PresetBorderStyle.AllThin:
                    range.Borders.LineStyle = XlLineStyle.xlContinuous;
                    range.Borders.Weight = XlBorderWeight.xlThin;
                    range.Borders.Color = borderColorOle;
                    break;

                case PresetBorderStyle.OutlineMedium:
                    // Viền trong mỏng
                    range.Borders.LineStyle = XlLineStyle.xlContinuous;
                    range.Borders.Weight = XlBorderWeight.xlThin;
                    range.Borders.Color = borderColorOle;

                    // Viền ngoài đậm
                    try
                    {
                        range.BorderAround2(
                            LineStyle: XlLineStyle.xlContinuous,
                            Weight: XlBorderWeight.xlMedium,
                            ColorIndex: XlColorIndex.xlColorIndexAutomatic,
                            Color: borderColorOle);
                    }
                    catch
                    {
                        range.Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlMedium;
                        range.Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlMedium;
                        range.Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlMedium;
                        range.Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlMedium;
                    }
                    break;

                case PresetBorderStyle.HeaderBottomDouble:
                    range.Borders.LineStyle = XlLineStyle.xlContinuous;
                    range.Borders.Weight = XlBorderWeight.xlThin;
                    range.Borders.Color = borderColorOle;

                    // Viền đáy đôi
                    range.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
                    range.Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThick;
                    range.Borders[XlBordersIndex.xlEdgeBottom].Color = borderColorOle;
                    break;

                case PresetBorderStyle.Dotted:
                    range.Borders.LineStyle = XlLineStyle.xlDot;
                    range.Borders.Weight = XlBorderWeight.xlHairline;
                    range.Borders.Color = borderColorOle;
                    break;

                case PresetBorderStyle.KeepCurrent:
                default:
                    // Không thay đổi viền
                    break;
            }
        }

        public static Color ParseHexColor(string? hex, Color fallback)
        {
            if (string.IsNullOrWhiteSpace(hex)) return fallback;

            try
            {
                string clean = hex!.Trim().TrimStart('#');
                if (clean.Length == 6)
                {
                    int r = Convert.ToInt32(clean.Substring(0, 2), 16);
                    int g = Convert.ToInt32(clean.Substring(2, 2), 16);
                    int b = Convert.ToInt32(clean.Substring(4, 2), 16);
                    return Color.FromArgb(r, g, b);
                }
                if (clean.Length == 8)
                {
                    int a = Convert.ToInt32(clean.Substring(0, 2), 16);
                    int r = Convert.ToInt32(clean.Substring(2, 2), 16);
                    int g = Convert.ToInt32(clean.Substring(4, 2), 16);
                    int b = Convert.ToInt32(clean.Substring(6, 2), 16);
                    return Color.FromArgb(a, r, g, b);
                }
                if (clean.Length == 3)
                {
                    int r = Convert.ToInt32(new string(clean[0], 2), 16);
                    int g = Convert.ToInt32(new string(clean[1], 2), 16);
                    int b = Convert.ToInt32(new string(clean[2], 2), 16);
                    return Color.FromArgb(r, g, b);
                }
            }
            catch { }

            return fallback;
        }
    }
}
