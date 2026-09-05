using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using ExcelDna.Integration.CustomUI;
using ExcelSupport.Host;
using ExcelSupport.ViewModels;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using VmSortOrder = ExcelSupport.ViewModels.SortOrder;

namespace ExcelSupport.Ribbon
{
    [ComVisible(true)]
    [Guid("A7B5F8C1-6E3D-4A2B-9F1C-8D4E2A1B3C5D")]
    public class RibbonController : ExcelRibbon
    {
        public static RibbonController? Instance { get; private set; }

        private IRibbonUI? _ribbon;

        public override string GetCustomUI(string ribbonId)
        {
            Instance = this;
            try
            {
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                using var stream = assembly.GetManifestResourceStream("ExcelSupport.Ribbon.CustomRibbon.xml");
                if (stream != null)
                {
                    using var reader = new System.IO.StreamReader(stream, System.Text.Encoding.UTF8);
                    return reader.ReadToEnd();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading CustomRibbon.xml resource: {ex.Message}");
            }

            return "<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' />";
        }

        public override object? LoadImage(string imageId)
        {
            if (imageId == "ribbon_settings_icon")
            {
                return CreateRibbonSettingsBitmap();
            }
            if (imageId == "navigator_icon")
            {
                return CreateNavigatorBitmap();
            }
            if (imageId == "compare_icon")
            {
                return CreateCompareBitmap();
            }
            if (imageId == "oracle_compare_icon")
            {
                return CreateOracleCompareBitmap();
            }
            if (imageId == "oracle_query_icon")
            {
                return CreateOracleQuickQueryBitmap();
            }
            if (imageId == "doctor_formula_icon")
            {
                return CreateDoctorFormulaBitmap();
            }
            if (imageId == "snapshot_rollback_icon")
            {
                return CreateSnapshotRollbackBitmap();
            }
            if (imageId == "cleaner_icon")
            {
                return CreateCleanerBitmap();
            }
            if (imageId == "duplicate_icon")
            {
                return CreateDuplicateBitmap();
            }
            if (imageId == "link_icon")
            {
                return CreateLinkBitmap();
            }
            if (imageId == "filter_icon")
            {
                return CreateFilterBitmap();
            }
            if (imageId == "blank_cleaner_icon")
            {
                return CreateBlankCleanerBitmap();
            }
            if (imageId == "find_replace_icon")
            {
                return CreateFindReplaceBitmap();
            }
            if (imageId == "ai_formula_icon")
            {
                return CreateAiFormulaBitmap();
            }
            if (imageId == "merge_icon")
            {
                return CreateMergeBitmap();
            }
            if (imageId == "table_merge_icon")
            {
                return CreateTableMergeBitmap();
            }
            if (imageId == "fuzzy_duplicate_icon")
            {
                return CreateFuzzyDuplicateBitmap();
            }
            if (imageId == "file_converter_icon")
            {
                return CreateFileConverterBitmap();
            }
            if (imageId == "page_counter_icon")
            {
                return CreatePageCounterBitmap();
            }
            if (imageId == "filtered_paste_icon")
            {
                return CreateFilteredPasteBitmap();
            }
            if (imageId == "manual_icon")
            {
                return CreateManualBitmap();
            }
            if (imageId == "globe_lang_icon")
            {
                return CreateGlobeLangBitmap();
            }
            if (imageId == "flag_vi_icon")
            {
                return CreateFlagViBitmap();
            }
            if (imageId == "flag_en_icon")
            {
                return CreateFlagEnBitmap();
            }
            if (imageId == "flag_ja_icon")
            {
                return CreateFlagJaBitmap();
            }
            if (imageId == "zenkaku_icon")
            {
                return CreateZenkakuBitmap();
            }
            if (imageId == "katakana_icon")
            {
                return CreateKatakanaBitmap();
            }
            if (imageId == "markdown_icon")
            {
                return CreateMarkdownBitmap();
            }
            if (imageId == "ruler_icon")
            {
                return CreateRulerBitmap();
            }
            if (imageId == "ruler_settings_icon")
            {
                return CreateRulerSettingsBitmap();
            }
            if (imageId == "hud_icon")
            {
                return CreateHudBitmap();
            }
            if (imageId == "color_yellow_icon")
            {
                return CreateColorSwatchBitmap(Color.FromArgb(253, 224, 71), Color.FromArgb(234, 179, 8));
            }
            if (imageId == "color_sky_icon")
            {
                return CreateColorSwatchBitmap(Color.FromArgb(56, 189, 248), Color.FromArgb(2, 132, 199));
            }
            if (imageId == "color_emerald_icon")
            {
                return CreateColorSwatchBitmap(Color.FromArgb(74, 222, 128), Color.FromArgb(22, 163, 74));
            }
            if (imageId == "color_orange_icon")
            {
                return CreateColorSwatchBitmap(Color.FromArgb(251, 146, 60), Color.FromArgb(234, 88, 12));
            }
            if (imageId == "color_purple_icon")
            {
                return CreateColorSwatchBitmap(Color.FromArgb(192, 132, 252), Color.FromArgb(147, 51, 234));
            }
            if (imageId == "color_pink_icon")
            {
                return CreateColorSwatchBitmap(Color.FromArgb(244, 114, 182), Color.FromArgb(219, 39, 119));
            }
            if (imageId == "color_gray_icon")
            {
                return CreateColorSwatchBitmap(Color.FromArgb(148, 163, 184), Color.FromArgb(100, 116, 139));
            }
            if (imageId == "mode_both_icon")
            {
                return CreateModeIcon("Both");
            }
            if (imageId == "mode_row_icon")
            {
                return CreateModeIcon("Row");
            }
            if (imageId == "mode_col_icon")
            {
                return CreateModeIcon("Col");
            }
            return base.LoadImage(imageId);
        }

        private Bitmap CreateRibbonSettingsBitmap()
        {
            var bmp = new Bitmap(32, 32);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.Clear(Color.Transparent);

                // Background soft slate circle
                using (var bgBrush = new SolidBrush(Color.FromArgb(241, 245, 249)))
                {
                    g.FillEllipse(bgBrush, 2, 2, 28, 28);
                }

                // Slate border
                using (var pen = new Pen(Color.FromArgb(148, 163, 184), 1.5f))
                {
                    g.DrawEllipse(pen, 2, 2, 28, 28);
                }

                // Gear center circle
                using (var gearBrush = new SolidBrush(Color.FromArgb(71, 85, 105)))
                {
                    g.FillEllipse(gearBrush, 9, 9, 14, 14);
                }

                // Center hole
                using (var holeBrush = new SolidBrush(Color.FromArgb(241, 245, 249)))
                {
                    g.FillEllipse(holeBrush, 13, 13, 6, 6);
                }

                // Gear teeth (8 teeth around)
                using (var toothPen = new Pen(Color.FromArgb(71, 85, 105), 3f))
                {
                    toothPen.StartCap = LineCap.Round;
                    toothPen.EndCap = LineCap.Round;
                    // Horizontal & Vertical teeth
                    g.DrawLine(toothPen, 16, 5, 16, 9);
                    g.DrawLine(toothPen, 16, 23, 16, 27);
                    g.DrawLine(toothPen, 5, 16, 9, 16);
                    g.DrawLine(toothPen, 23, 16, 27, 16);
                    // Diagonal teeth
                    g.DrawLine(toothPen, 8, 8, 11, 11);
                    g.DrawLine(toothPen, 21, 21, 24, 24);
                    g.DrawLine(toothPen, 24, 8, 21, 11);
                    g.DrawLine(toothPen, 8, 24, 11, 21);
                }

                // Sparkle / Accent dot (Indigo #6366F1)
                using (var accentBrush = new SolidBrush(Color.FromArgb(99, 102, 241)))
                {
                    g.FillEllipse(accentBrush, 20, 4, 7, 7);
                }
            }
            return bmp;
        }

        private Bitmap CreateCompareBitmap()
        {
            var bmp = new Bitmap(32, 32);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.Clear(Color.Transparent);

                // 1. Sheet A (Phía sau bên trái - Màu Xanh Dương #2563EB)
                using (var brushA = new LinearGradientBrush(new Rectangle(2, 3, 17, 22),
                    Color.FromArgb(37, 99, 235), Color.FromArgb(29, 78, 216), 45f))
                {
                    FillRoundedRectangle(g, brushA, new Rectangle(2, 3, 17, 22), 3);
                }
                using (var penA = new Pen(Color.FromArgb(30, 64, 175), 1f))
                {
                    DrawRoundedRectangle(g, penA, new Rectangle(2, 3, 17, 22), 3);
                }

                // Header Sheet A
                using (var headBrush = new SolidBrush(Color.FromArgb(220, 255, 255, 255)))
                {
                    FillRoundedRectangle(g, headBrush, new Rectangle(4, 5, 13, 4), 1);
                }
                // Các dòng Sheet A
                using (var lineBrush = new SolidBrush(Color.FromArgb(160, 255, 255, 255)))
                {
                    g.FillRectangle(lineBrush, 4, 11, 13, 2);
                    g.FillRectangle(lineBrush, 4, 15, 13, 2);
                    g.FillRectangle(lineBrush, 4, 19, 13, 2);
                }

                // 2. Sheet B (Phía trước bên phải - Màu Xanh Lục Excel #107C41)
                using (var brushB = new LinearGradientBrush(new Rectangle(13, 7, 17, 22),
                    Color.FromArgb(16, 124, 65), Color.FromArgb(10, 85, 42), 45f))
                {
                    FillRoundedRectangle(g, brushB, new Rectangle(13, 7, 17, 22), 3);
                }
                using (var penB = new Pen(Color.FromArgb(6, 60, 28), 1f))
                {
                    DrawRoundedRectangle(g, penB, new Rectangle(13, 7, 17, 22), 3);
                }

                // Header Sheet B
                using (var headBrushB = new SolidBrush(Color.FromArgb(220, 255, 255, 255)))
                {
                    FillRoundedRectangle(g, headBrushB, new Rectangle(15, 9, 13, 4), 1);
                }
                // Các dòng Sheet B (với 1 ô màu Cam nổi bật thể hiện sự khác biệt!)
                using (var lineBrushB = new SolidBrush(Color.FromArgb(160, 255, 255, 255)))
                {
                    g.FillRectangle(lineBrushB, 15, 15, 13, 2);
                    // Dòng khác biệt màu cam
                    using (var diffBrush = new SolidBrush(Color.FromArgb(245, 158, 11)))
                    {
                        g.FillRectangle(diffBrush, 15, 19, 13, 2);
                    }
                    g.FillRectangle(lineBrushB, 15, 23, 13, 2);
                }

                // 3. Biểu tượng mũi tên so sánh 2 chiều ⇋ hình tròn ở góc dưới bên trái
                using (var circleBrush = new SolidBrush(Color.FromArgb(245, 158, 11))) // Amber Badge
                {
                    g.FillEllipse(circleBrush, 1, 17, 14, 14);
                }
                using (var circlePen = new Pen(Color.FromArgb(255, 255, 255), 1.2f))
                {
                    g.DrawEllipse(circlePen, 1, 17, 14, 14);
                }
                // Mũi tên ⇋ màu trắng bên trong hình tròn
                using (var arrowPen = new Pen(Color.White, 1.5f))
                {
                    // Mũi tên trên sang phải: ->
                    g.DrawLine(arrowPen, 3, 22, 11, 22);
                    g.DrawLine(arrowPen, 9, 20, 12, 22);
                    // Mũi tên dưới sang trái: <-
                    g.DrawLine(arrowPen, 4, 26, 12, 26);
                    g.DrawLine(arrowPen, 6, 28, 3, 26);
                }
            }
            return bmp;
        }

        private Bitmap CreateOracleCompareBitmap()
        {
            var bmp = new Bitmap(32, 32);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.Clear(Color.Transparent);

                // 1. Database A (Bên trái - Màu Đỏ Cam thương hiệu Oracle #EA580C / #C2410C)
                using (var brushDbA = new LinearGradientBrush(new Rectangle(2, 4, 13, 22),
                    Color.FromArgb(234, 88, 12), Color.FromArgb(180, 50, 8), 45f))
                {
                    FillRoundedRectangle(g, brushDbA, new Rectangle(2, 4, 13, 22), 2);
                }
                using (var penDbA = new Pen(Color.FromArgb(154, 52, 18), 1f))
                {
                    DrawRoundedRectangle(g, penDbA, new Rectangle(2, 4, 13, 22), 2);
                }
                // Các vạch đĩa Database A
                using (var diskPenA = new Pen(Color.FromArgb(254, 215, 170), 1f))
                {
                    g.DrawLine(diskPenA, 4, 10, 13, 10);
                    g.DrawLine(diskPenA, 4, 16, 13, 16);
                    g.DrawLine(diskPenA, 4, 21, 13, 21);
                }

                // 2. Database B (Bên phải - Màu Xanh Đậm / Xanh Ngọc #0284C7 / #0369A1)
                using (var brushDbB = new LinearGradientBrush(new Rectangle(17, 4, 13, 22),
                    Color.FromArgb(2, 132, 199), Color.FromArgb(3, 105, 161), 45f))
                {
                    FillRoundedRectangle(g, brushDbB, new Rectangle(17, 4, 13, 22), 2);
                }
                using (var penDbB = new Pen(Color.FromArgb(7, 89, 133), 1f))
                {
                    DrawRoundedRectangle(g, penDbB, new Rectangle(17, 4, 13, 22), 2);
                }
                // Các vạch đĩa Database B
                using (var diskPenB = new Pen(Color.FromArgb(186, 230, 253), 1f))
                {
                    g.DrawLine(diskPenB, 19, 10, 28, 10);
                    g.DrawLine(diskPenB, 19, 16, 28, 16);
                    g.DrawLine(diskPenB, 19, 21, 28, 21);
                }

                // 3. Vòng tròn Badge Mũi tên so sánh ở giữa phía dưới
                using (var badgeBrush = new SolidBrush(Color.FromArgb(245, 158, 11))) // Amber
                {
                    g.FillEllipse(badgeBrush, 9, 17, 14, 14);
                }
                using (var badgePen = new Pen(Color.White, 1.2f))
                {
                    g.DrawEllipse(badgePen, 9, 17, 14, 14);
                }

                // Mũi tên 2 chiều ⇋
                using (var arrowPen = new Pen(Color.White, 1.5f))
                {
                    // Mũi tên trên: ->
                    g.DrawLine(arrowPen, 11, 22, 19, 22);
                    g.DrawLine(arrowPen, 17, 20, 20, 22);
                    // Mũi tên dưới: <-
                    g.DrawLine(arrowPen, 13, 26, 21, 26);
                    g.DrawLine(arrowPen, 15, 28, 12, 26);
                }
            }
            return bmp;
        }

        private Bitmap CreateOracleQuickQueryBitmap()
        {
            var bmp = new Bitmap(32, 32);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.Clear(Color.Transparent);

                // 1. Database Block (Oracle Red-Orange #EA580C)
                using (var brushDb = new LinearGradientBrush(new Rectangle(4, 5, 24, 22),
                    Color.FromArgb(234, 88, 12), Color.FromArgb(180, 50, 8), 45f))
                {
                    FillRoundedRectangle(g, brushDb, new Rectangle(4, 5, 24, 22), 3);
                }
                using (var penDb = new Pen(Color.FromArgb(154, 52, 18), 1f))
                {
                    DrawRoundedRectangle(g, penDb, new Rectangle(4, 5, 24, 22), 3);
                }

                // Các vạch đĩa Database
                using (var diskPen = new Pen(Color.FromArgb(254, 215, 170), 1.2f))
                {
                    g.DrawLine(diskPen, 7, 11, 25, 11);
                    g.DrawLine(diskPen, 7, 17, 25, 17);
                    g.DrawLine(diskPen, 7, 23, 25, 23);
                }

                // 2. Tia sét / Flash Query màu vàng sáng (#F59E0B / #FEF08A)
                Point[] lightning = new Point[]
                {
                    new Point(18, 3),
                    new Point(11, 15),
                    new Point(16, 15),
                    new Point(13, 27),
                    new Point(23, 13),
                    new Point(18, 13)
                };

                using (var shadowBrush = new SolidBrush(Color.FromArgb(100, 0, 0, 0)))
                {
                    Point[] shadowPoints = lightning.Select(p => new Point(p.X + 1, p.Y + 1)).ToArray();
                    g.FillPolygon(shadowBrush, shadowPoints);
                }

                using (var boltBrush = new LinearGradientBrush(new Rectangle(11, 3, 12, 24),
                    Color.FromArgb(254, 240, 138), Color.FromArgb(245, 158, 11), 90f))
                {
                    g.FillPolygon(boltBrush, lightning);
                }
                using (var boltPen = new Pen(Color.FromArgb(180, 83, 9), 1f))
                {
                    g.DrawPolygon(boltPen, lightning);
                }
            }
            return bmp;
        }

        private Bitmap CreateDoctorFormulaBitmap()
        {
            var bmp = new Bitmap(32, 32);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.Clear(Color.Transparent);

                // 1. Nền thẻ Bác sĩ công thức (Gradient Xanh Dương Y Tế #0284C7 / #0369A1)
                using (var brush = new LinearGradientBrush(new Rectangle(2, 2, 28, 28),
                    Color.FromArgb(2, 132, 199), Color.FromArgb(3, 105, 161), 45f))
                {
                    FillRoundedRectangle(g, brush, new Rectangle(2, 2, 28, 28), 4);
                }
                using (var borderPen = new Pen(Color.FromArgb(7, 89, 133), 1f))
                {
                    DrawRoundedRectangle(g, borderPen, new Rectangle(2, 2, 28, 28), 4);
                }

                // 2. Chữ "fx" màu trắng sáng
                using (var font = new System.Drawing.Font("Georgia", 11f, System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))
                using (var textBrush = new SolidBrush(Color.White))
                {
                    g.DrawString("fx", font, textBrush, new PointF(4, 5));
                }

                // 3. Huy hiệu Chữ Thập Y Tế / Cứu Hộ ở góc dưới bên phải (Vòng tròn Đỏ/Trắng)
                using (var badgeBrush = new SolidBrush(Color.FromArgb(239, 68, 68))) // Red 500
                {
                    g.FillEllipse(badgeBrush, 16, 16, 14, 14);
                }
                using (var badgePen = new Pen(Color.White, 1.2f))
                {
                    g.DrawEllipse(badgePen, 16, 16, 14, 14);
                }

                // Chữ thập y tế màu trắng bên trong
                using (var crossBrush = new SolidBrush(Color.White))
                {
                    g.FillRectangle(crossBrush, 21, 18, 4, 10);
                    g.FillRectangle(crossBrush, 18, 21, 10, 4);
                }
            }
            return bmp;
        }

        private Bitmap CreateSnapshotRollbackBitmap()
        {
            var bmp = new Bitmap(32, 32);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.Clear(Color.Transparent);

                // 1. Khung Sheet tài liệu nền (Slate Dark #0F172A / #1E293B)
                using (var brush = new LinearGradientBrush(new Rectangle(2, 2, 28, 28),
                    Color.FromArgb(30, 41, 59), Color.FromArgb(15, 23, 42), 45f))
                {
                    FillRoundedRectangle(g, brush, new Rectangle(2, 2, 28, 28), 4);
                }
                using (var borderPen = new Pen(Color.FromArgb(51, 65, 85), 1f))
                {
                    DrawRoundedRectangle(g, borderPen, new Rectangle(2, 2, 28, 28), 4);
                }

                // 2. Các vạch bảng tính mờ màu xanh nhạt
                using (var gridPen = new Pen(Color.FromArgb(100, 148, 163, 184), 1f))
                {
                    g.DrawLine(gridPen, 6, 8, 26, 8);
                    g.DrawLine(gridPen, 6, 13, 26, 13);
                    g.DrawLine(gridPen, 6, 18, 14, 18);
                    g.DrawLine(gridPen, 6, 23, 14, 23);
                }

                // 3. Biểu tượng Đồng Hồ / Mũi Tên Rollback Xoay Ngược màu Xanh Lục Ngọc #10B981
                using (var circleBrush = new SolidBrush(Color.FromArgb(16, 185, 129)))
                {
                    g.FillEllipse(circleBrush, 15, 15, 15, 15);
                }
                using (var circlePen = new Pen(Color.White, 1.2f))
                {
                    g.DrawEllipse(circlePen, 15, 15, 15, 15);
                }

                // Mũi tên quay ngược (Rewind / Undo arrow)
                using (var arrowPen = new Pen(Color.White, 1.5f))
                {
                    // Vòng cung kim đồng hồ
                    g.DrawArc(arrowPen, 18, 18, 9, 9, 45, 260);
                    // Đầu mũi tên quay ngược
                    g.DrawLine(arrowPen, 21, 17, 18, 20);
                    g.DrawLine(arrowPen, 21, 23, 18, 20);
                }
            }
            return bmp;
        }

        private Bitmap CreateNavigatorBitmap()
        {
            // Tạo icon độ phân giải cao 32x32 với khử răng cưa chuẩn Office
            var bmp = new Bitmap(32, 32);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.Clear(Color.Transparent);

                // Khung nền ứng dụng Excel (Màu xanh thương hiệu Excel #107C41)
                using (var brush = new LinearGradientBrush(new Rectangle(1, 1, 30, 30),
                    Color.FromArgb(16, 124, 65), Color.FromArgb(10, 85, 42), 45f))
                {
                    FillRoundedRectangle(g, brush, new Rectangle(1, 1, 30, 30), 4);
                }

                // Viền khung
                using (var borderPen = new Pen(Color.FromArgb(6, 60, 28), 1f))
                {
                    DrawRoundedRectangle(g, borderPen, new Rectangle(1, 1, 30, 30), 4);
                }

                // Vùng 1: Header Workbook ở phía trên (Màu trắng mờ)
                using (var headerBrush = new SolidBrush(Color.FromArgb(230, 255, 255, 255)))
                {
                    FillRoundedRectangle(g, headerBrush, new Rectangle(4, 4, 24, 7), 2);
                }

                // Biểu tượng thư mục nhỏ ở Header
                using (var docBrush = new SolidBrush(Color.FromArgb(16, 124, 65)))
                {
                    g.FillRectangle(docBrush, 6, 6, 4, 3);
                    g.FillRectangle(docBrush, 12, 6, 13, 3);
                }

                // Vùng 2: Các Sheet nằm ở phía dưới (3 dòng thẻ Sheet trực quan)
                int[] yOffsets = { 13, 18, 23 };
                foreach (int y in yOffsets)
                {
                    // Nền từng Sheet
                    using (var sheetBrush = new SolidBrush(Color.FromArgb(240, 255, 255, 255)))
                    {
                        FillRoundedRectangle(g, sheetBrush, new Rectangle(4, y, 24, 4), 1);
                    }
                    // Dấu chấm đầu dòng của Sheet
                    using (var dotBrush = new SolidBrush(Color.FromArgb(16, 124, 65)))
                    {
                        g.FillRectangle(dotBrush, 6, y + 1, 2, 2);
                        g.FillRectangle(dotBrush, 9, y + 1, 16, 2);
                    }
                }
            }
            return bmp;
        }

        private static void FillRoundedRectangle(Graphics g, Brush brush, Rectangle bounds, int cornerRadius)
        {
            using (var path = GetRoundedPath(bounds, cornerRadius))
            {
                g.FillPath(brush, path);
            }
        }

        private static void DrawRoundedRectangle(Graphics g, Pen pen, Rectangle bounds, int cornerRadius)
        {
            using (var path = GetRoundedPath(bounds, cornerRadius))
            {
                g.DrawPath(pen, path);
            }
        }

        private static GraphicsPath GetRoundedPath(Rectangle bounds, int radius)
        {
            var path = new GraphicsPath();
            int diameter = radius * 2;
            var arc = new Rectangle(bounds.Location, new Size(diameter, diameter));

            path.AddArc(arc, 180, 90);
            arc.X = bounds.Right - diameter;
            path.AddArc(arc, 270, 90);
            arc.Y = bounds.Bottom - diameter;
            path.AddArc(arc, 0, 90);
            arc.X = bounds.Left;
            path.AddArc(arc, 90, 90);
            path.CloseFigure();
            return path;
        }

        public void OnRibbonLoaded(IRibbonUI ribbon)
        {
            _ribbon = ribbon;
            Instance = this;

            TaskPaneRegistry.VisibilityChanged += isVisible =>
            {
                _ribbon?.InvalidateControl("btnToggleTaskPane");
            };

            Services.LocalizationService.LanguageChanged += lang =>
            {
                _ribbon?.Invalidate();
            };
        }

        public string GetRibbonLabel(IRibbonControl control)
        {
            return Services.LocalizationService.GetLabel(control.Id);
        }

        public string GetRibbonScreentip(IRibbonControl control)
        {
            return Services.LocalizationService.GetScreentip(control.Id);
        }

        public string GetRibbonSupertip(IRibbonControl control)
        {
            return Services.LocalizationService.GetSupertip(control.Id);
        }

        public void OnSelectLanguage(IRibbonControl control)
        {
            string tag = control.Tag ?? "vi";
            var lang = tag switch
            {
                "en" => Services.AppLanguage.English,
                "ja" => Services.AppLanguage.Japanese,
                _ => Services.AppLanguage.Vietnamese
            };
            Services.LocalizationService.CurrentLanguage = lang;
            _ribbon?.Invalidate();
        }

        public bool GetControlVisible(IRibbonControl control)
        {
            return Services.RibbonVisibilityService.IsControlVisible(control.Id);
        }

        public void OnCustomizeRibbon(IRibbonControl control)
        {
            try
            {
                Views.RibbonCustomizeDialog.ShowWindow();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Không thể mở Tùy chỉnh Ribbon:\n{ex.Message}", "Lỗi giao diện", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void InvalidateRibbon()
        {
            _ribbon?.Invalidate();
        }

        public void InvalidateControl(string controlId)
        {
            _ribbon?.InvalidateControl(controlId);
        }

        public bool GetTaskPanePressed(IRibbonControl control)
        {
            return TaskPaneRegistry.IsTaskPaneVisible;
        }

        public void OnToggleTaskPane(IRibbonControl control, bool pressed)
        {
            if (AddInEvents.MainViewModel != null)
            {
                TaskPaneRegistry.ToggleTaskPane(AddInEvents.MainViewModel, pressed);
            }
        }

        public void OnRefreshTree(IRibbonControl control)
        {
            AddInEvents.Instance?.RefreshWorkbookTreePublic();
        }

        public void OnCreateTOC(IRibbonControl control)
        {
            ExcelDna.Integration.ExcelAsyncUtil.QueueAsMacro(() =>
            {
                AddInEvents.Instance?.CreateTableOfContents(null);
            });
        }

        public void OnSplitSheets(IRibbonControl control)
        {
            if (AddInEvents.MainViewModel?.SelectedWorkbook != null)
            {
                var dlg = new Views.SheetToolsDialog(AddInEvents.MainViewModel.SelectedWorkbook, 0, AddInEvents.MainViewModel.IsDarkTheme);
                dlg.ShowDialog();
            }
        }

        public void OnMergeSheets(IRibbonControl control)
        {
            if (AddInEvents.MainViewModel?.SelectedWorkbook != null)
            {
                var dlg = new Views.SheetToolsDialog(AddInEvents.MainViewModel.SelectedWorkbook, 1, AddInEvents.MainViewModel.IsDarkTheme);
                dlg.ShowDialog();
            }
        }

        public void OnSortAZ(IRibbonControl control)
        {
            if (AddInEvents.MainViewModel != null)
            {
                AddInEvents.MainViewModel.WorkbookSortOrder = VmSortOrder.Ascending;
                AddInEvents.MainViewModel.SheetSortOrder = VmSortOrder.Ascending;
            }
        }

        public void OnSortZA(IRibbonControl control)
        {
            if (AddInEvents.MainViewModel != null)
            {
                AddInEvents.MainViewModel.WorkbookSortOrder = VmSortOrder.Descending;
                AddInEvents.MainViewModel.SheetSortOrder = VmSortOrder.Descending;
            }
        }

        public void OnCloseCurrentWorkbook(IRibbonControl control)
        {
            if (AddInEvents.MainViewModel?.SelectedWorkbook != null)
            {
                AddInEvents.MainViewModel.CloseWorkbookCommand.Execute(AddInEvents.MainViewModel.SelectedWorkbook.WorkbookName);
            }
        }

        public void OnCompareWorkbooks(IRibbonControl control)
        {
            Views.WorkbookCompareDialog.ShowWindow(AddInEvents.MainViewModel?.SelectedWorkbook?.WorkbookName, AddInEvents.MainViewModel?.IsDarkTheme ?? false);
        }

        public void OnCheckVietnamese(IRibbonControl control)
        {
            Views.VietnameseCheckDialog.ShowWindow(AddInEvents.MainViewModel?.IsDarkTheme ?? false);
        }

        public void OnExternalLinksManager(IRibbonControl control)
        {
            Views.ExternalLinksManagerDialog.ShowWindow(AddInEvents.MainViewModel?.IsDarkTheme ?? false);
        }

        public void OnOracleTableCompare(IRibbonControl control)
        {
            Views.OracleTableCompareDialog.ShowWindow(AddInEvents.MainViewModel?.IsDarkTheme ?? false);
        }

        public void OnOracleQuickQuery(IRibbonControl control)
        {
            Views.OracleQuickQueryDialog.ShowWindow(AddInEvents.MainViewModel?.IsDarkTheme ?? false);
        }

        public void OnAdvancedFilter(IRibbonControl control)
        {
            Views.AdvancedFilterDialog.ShowWindow(AddInEvents.MainViewModel?.IsDarkTheme ?? false);
        }

        public void OnDataCleaner(IRibbonControl control)
        {
            Views.DataCleaningWizardDialog.ShowWindow(AddInEvents.MainViewModel?.IsDarkTheme ?? false);
        }

        public void OnDuplicateFinder(IRibbonControl control)
        {
            Views.DuplicateFinderDialog.ShowWindow(AddInEvents.MainViewModel?.IsDarkTheme ?? false);
        }

        public void OnBatchBlankCleaner(IRibbonControl control)
        {
            Views.BatchCleanerAndMergeDialog.ShowWindow(0, AddInEvents.MainViewModel?.IsDarkTheme ?? false);
        }

        public void OnBatchFindReplace(IRibbonControl control)
        {
            Views.BatchFindReplaceDialog.ShowWindow(AddInEvents.MainViewModel?.IsDarkTheme ?? false);
        }

        public void OnVisualTableMerge(IRibbonControl control)
        {
            Views.VisualTableMergeDialog.ShowWindow(AddInEvents.MainViewModel?.IsDarkTheme ?? false);
        }

        public void OnFuzzyDuplicate(IRibbonControl control)
        {
            Views.FuzzyDuplicateDialog.ShowWindow(AddInEvents.MainViewModel?.IsDarkTheme ?? false);
        }

        public void OnBatchFileConverter(IRibbonControl control)
        {
            Views.BatchFileConverterDialog.ShowWindow(AddInEvents.MainViewModel?.IsDarkTheme ?? false);
        }

        public void OnDesignPageCounter(IRibbonControl control)
        {
            Views.DesignPageCounterDialog.ShowWindow(AddInEvents.MainViewModel?.IsDarkTheme ?? false);
        }

        public void OnJapaneseConvert(IRibbonControl control)
        {
            var app = AddInEvents.Instance?.ExcelAppInstance ?? (ExcelApp)ExcelDna.Integration.ExcelDnaUtil.Application;
            Views.JapaneseTextConverterDialog.ShowWindow(app);
        }

        public void OnKatakanaCheck(IRibbonControl control)
        {
            var app = AddInEvents.Instance?.ExcelAppInstance ?? (ExcelApp)ExcelDna.Integration.ExcelDnaUtil.Application;
            Views.KatakanaValidatorDialog.ShowWindow(app);
        }

        public void OnExportMarkdown(IRibbonControl control)
        {
            var app = AddInEvents.Instance?.ExcelAppInstance ?? (ExcelApp)ExcelDna.Integration.ExcelDnaUtil.Application;
            Views.TableExportDialog.ShowWindow(app);
        }

        public void OnSafeMergeConsolidate(IRibbonControl control)
        {
            Views.BatchCleanerAndMergeDialog.ShowWindow(1, AddInEvents.MainViewModel?.IsDarkTheme ?? false);
        }

        public void OnFilteredCopyPasteWizard(IRibbonControl control)
        {
            Views.FilteredCopyPasteDialog.ShowWindow(AddInEvents.MainViewModel?.IsDarkTheme ?? false);
        }

        public void OnCopyVisibleOnly(IRibbonControl control)
        {
            var app = AddInEvents.Instance?.ExcelAppInstance;
            var result = Services.FilteredCopyPasteService.CopyVisibleCells(app);
            if (!result.Success)
            {
                System.Windows.MessageBox.Show(result.Message, "Thông Báo", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
            }
        }

        public void OnPasteToVisibleOnly(IRibbonControl control)
        {
            var app = AddInEvents.Instance?.ExcelAppInstance;
            var result = Services.FilteredCopyPasteService.PasteToVisibleCells(app);
            if (result.Success)
            {
                System.Windows.MessageBox.Show(result.Message, "Dán Dữ Liệu Thành Công", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
            }
            else
            {
                System.Windows.MessageBox.Show(result.Message, "Thông Báo", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
            }
        }

        public void OnAiFormula(IRibbonControl control)
        {
            if (AddInEvents.MainViewModel != null)
            {
                TaskPaneRegistry.ToggleTaskPane(AddInEvents.MainViewModel, true);
                AddInEvents.MainViewModel.SelectedTabIndex = 1; // Tab AI Assistant
                if (AddInEvents.MainViewModel.AiAssistant != null)
                {
                    AddInEvents.MainViewModel.AiAssistant.SelectedSubTab = 1; // Sub-Tab Sinh Công Thức
                }
            }
        }

        public void OnAiFormulaDoctor(IRibbonControl control)
        {
            Views.AiFormulaDoctorDialog.ShowWindow(AddInEvents.MainViewModel?.IsDarkTheme ?? false);
        }

        public void OnSnapshotRollback(IRibbonControl control)
        {
            Views.SheetSnapshotDialog.ShowWindow(AddInEvents.MainViewModel?.IsDarkTheme ?? false);
        }

        public void OnUserManual(IRibbonControl control)
        {
            Views.UserManualDialog.ShowWindow(AddInEvents.MainViewModel?.IsDarkTheme ?? false);
        }

        public void OnToggleGridRuler(IRibbonControl control, bool pressed)
        {
            var app = AddInEvents.Instance?.ExcelAppInstance;
            Services.GridRulerService.Toggle(app);
            _ribbon?.InvalidateControl("btnToggleGridRuler");
        }

        public bool GetGridRulerPressed(IRibbonControl control)
        {
            return Services.GridRulerService.IsEnabled;
        }

        public void OnSelectRulerColor(IRibbonControl control)
        {
            string colorKey = control.Tag ?? "Yellow";
            var app = AddInEvents.Instance?.ExcelAppInstance;
            Services.GridRulerService.SetColor(colorKey, app);
            if (!Services.GridRulerService.IsEnabled)
            {
                Services.GridRulerService.Toggle(app);
                _ribbon?.InvalidateControl("btnToggleGridRuler");
            }
        }

        public void OnSelectRulerMode(IRibbonControl control)
        {
            string modeTag = control.Tag ?? "Both";
            var mode = Services.GridRulerMode.BothRowAndCol;
            if (modeTag == "Row") mode = Services.GridRulerMode.RowOnly;
            else if (modeTag == "Col") mode = Services.GridRulerMode.ColOnly;

            var app = AddInEvents.Instance?.ExcelAppInstance;
            Services.GridRulerService.SetMode(mode, app);
            if (!Services.GridRulerService.IsEnabled)
            {
                Services.GridRulerService.Toggle(app);
                _ribbon?.InvalidateControl("btnToggleGridRuler");
            }
        }

        public void OnToggleRulerHud(IRibbonControl control)
        {
            var isDark = AddInEvents.MainViewModel?.IsDarkTheme ?? true;
            Views.RulerHudWindow.ForceOpenHud(isDark);
        }

        private Bitmap CreateCleanerBitmap()
        {
            var bmp = new Bitmap(32, 32);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.Clear(Color.Transparent);

                // Nền bảng tính xanh lục Excel
                using (var brush = new LinearGradientBrush(new Rectangle(2, 2, 28, 28),
                    Color.FromArgb(16, 124, 65), Color.FromArgb(10, 85, 42), 45f))
                {
                    FillRoundedRectangle(g, brush, new Rectangle(2, 2, 28, 28), 4);
                }

                // Viền
                using (var pen = new Pen(Color.FromArgb(6, 60, 28), 1f))
                {
                    DrawRoundedRectangle(g, pen, new Rectangle(2, 2, 28, 28), 4);
                }

                // Các ô lưới trắng mờ
                using (var gridBrush = new SolidBrush(Color.FromArgb(200, 255, 255, 255)))
                {
                    g.FillRectangle(gridBrush, 5, 5, 6, 4);
                    g.FillRectangle(gridBrush, 13, 5, 6, 4);
                    g.FillRectangle(gridBrush, 5, 11, 6, 4);
                    g.FillRectangle(gridBrush, 13, 11, 6, 4);
                }

                // Cây chổi thần / Tia sáng lấp lánh (Sparkle & Magic Wand)
                using (var wandBrush = new SolidBrush(Color.FromArgb(245, 158, 11)))
                using (var wandPen = new Pen(Color.FromArgb(254, 240, 138), 2f))
                {
                    // Thân đũa thần
                    g.DrawLine(wandPen, 10, 26, 24, 12);
                    // Đầu đũa ngôi sao / lấp lánh
                    g.FillEllipse(wandBrush, 22, 10, 6, 6);
                }

                // Tia lấp lánh nhỏ
                using (var starBrush = new SolidBrush(Color.White))
                {
                    g.FillEllipse(starBrush, 24, 6, 3, 3);
                    g.FillEllipse(starBrush, 27, 13, 3, 3);
                    g.FillEllipse(starBrush, 18, 11, 2, 2);
                }
            }
            return bmp;
        }

        private Bitmap CreateDuplicateBitmap()
        {
            var bmp = new Bitmap(32, 32);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.Clear(Color.Transparent);

                // Trang 1 phía sau (Màu tím Pastel)
                using (var card1 = new LinearGradientBrush(new Rectangle(8, 2, 20, 22),
                    Color.FromArgb(139, 92, 246), Color.FromArgb(109, 40, 217), 45f))
                {
                    FillRoundedRectangle(g, card1, new Rectangle(8, 2, 20, 22), 3);
                }

                // Trang 2 phía trước (Màu xanh dương đậm)
                using (var card2 = new LinearGradientBrush(new Rectangle(3, 8, 20, 22),
                    Color.FromArgb(37, 99, 235), Color.FromArgb(29, 78, 216), 45f))
                {
                    FillRoundedRectangle(g, card2, new Rectangle(3, 8, 20, 22), 3);
                }

                // Viền trắng cho trang trước
                using (var pen = new Pen(Color.White, 1f))
                {
                    DrawRoundedRectangle(g, pen, new Rectangle(3, 8, 20, 22), 3);
                }

                // Các dòng nội dung trên trang trước
                using (var lineBrush = new SolidBrush(Color.FromArgb(200, 255, 255, 255)))
                {
                    g.FillRectangle(lineBrush, 6, 12, 14, 2);
                    g.FillRectangle(lineBrush, 6, 16, 14, 2);
                    g.FillRectangle(lineBrush, 6, 20, 10, 2);
                }

                // Huy hiệu kính lúp & số 2 (Duplicate Badge)
                using (var badgeBrush = new SolidBrush(Color.FromArgb(245, 158, 11)))
                using (var badgePen = new Pen(Color.White, 1.5f))
                {
                    g.FillEllipse(badgeBrush, 17, 17, 13, 13);
                    g.DrawEllipse(badgePen, 17, 17, 13, 13);
                }

                // Dấu "=" màu trắng bên trong huy hiệu thể hiện trùng khớp
                using (var equalPen = new Pen(Color.White, 2f))
                {
                    g.DrawLine(equalPen, 21, 22, 26, 22);
                    g.DrawLine(equalPen, 21, 25, 26, 25);
                }
            }
            return bmp;
        }

        private Bitmap CreateLinkBitmap()
        {
            var bmp = new Bitmap(32, 32);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.Clear(Color.Transparent);

                // Nền gradient xanh dương đậm chuyên nghiệp
                using (var brush = new LinearGradientBrush(new Rectangle(2, 2, 28, 28),
                    Color.FromArgb(2, 132, 199), Color.FromArgb(3, 105, 161), 45f))
                {
                    FillRoundedRectangle(g, brush, new Rectangle(2, 2, 28, 28), 4);
                }

                // Viền ngoài
                using (var pen = new Pen(Color.FromArgb(2, 84, 130), 1f))
                {
                    DrawRoundedRectangle(g, pen, new Rectangle(2, 2, 28, 28), 4);
                }

                // Vẽ 2 mắt xích liên kết (Chain Links) xoay nghiêng 45 độ
                using (var linkPen = new Pen(Color.White, 2.5f))
                {
                    linkPen.StartCap = LineCap.Round;
                    linkPen.EndCap = LineCap.Round;

                    // Mắt xích 1: góc trên bên trái
                    g.DrawArc(linkPen, 6, 8, 10, 10, 135, 180);
                    g.DrawLine(linkPen, 9, 8, 14, 13);
                    g.DrawLine(linkPen, 6, 12, 11, 17);

                    // Mắt xích 2: góc dưới bên phải
                    g.DrawArc(linkPen, 15, 13, 10, 10, -45, 180);
                    g.DrawLine(linkPen, 20, 14, 15, 19);
                    g.DrawLine(linkPen, 24, 18, 19, 23);
                }

                // Huy hiệu tia chớp cam (Fix/Break Link Badge)
                using (var badgeBrush = new SolidBrush(Color.FromArgb(245, 158, 11)))
                using (var badgePen = new Pen(Color.White, 1.2f))
                {
                    g.FillEllipse(badgeBrush, 17, 17, 13, 13);
                    g.DrawEllipse(badgePen, 17, 17, 13, 13);
                }

                // Biểu tượng tia sét bên trong huy hiệu
                using (var boltBrush = new SolidBrush(Color.White))
                {
                    var pts = new PointF[]
                    {
                        new PointF(25f, 19f),
                        new PointF(21f, 24f),
                        new PointF(24f, 24f),
                        new PointF(22f, 28f),
                        new PointF(27f, 23f),
                        new PointF(24f, 23f)
                    };
                    g.FillPolygon(boltBrush, pts);
                }
            }
            return bmp;
        }

        private Bitmap CreateFilterBitmap()
        {
            var bmp = new Bitmap(32, 32);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.Clear(Color.Transparent);

                // Nền gradient tím - chàm hiện đại
                using (var brush = new LinearGradientBrush(new Rectangle(2, 2, 28, 28),
                    Color.FromArgb(99, 102, 241), Color.FromArgb(67, 56, 202), 45f))
                {
                    FillRoundedRectangle(g, brush, new Rectangle(2, 2, 28, 28), 4);
                }

                // Viền ngoài
                using (var pen = new Pen(Color.FromArgb(49, 46, 129), 1f))
                {
                    DrawRoundedRectangle(g, pen, new Rectangle(2, 2, 28, 28), 4);
                }

                // Vẽ Phễu lọc (Funnel) màu trắng
                using (var funnelBrush = new SolidBrush(Color.White))
                using (var funnelPen = new Pen(Color.White, 1.5f))
                {
                    funnelPen.LineJoin = LineJoin.Round;

                    var funnelPts = new PointF[]
                    {
                        new PointF(6f, 7f),
                        new PointF(26f, 7f),
                        new PointF(18f, 17f),
                        new PointF(18f, 24f),
                        new PointF(14f, 26f),
                        new PointF(14f, 17f)
                    };
                    g.FillPolygon(funnelBrush, funnelPts);
                }

                // Các vạch dòng dữ liệu bên trên phễu
                using (var linePen = new Pen(Color.FromArgb(200, 255, 255, 255), 1.5f))
                {
                    g.DrawLine(linePen, 9, 10, 23, 10);
                    g.DrawLine(linePen, 11, 13, 21, 13);
                }

                // Huy hiệu tia chớp cam / Smart Filter Badge
                using (var badgeBrush = new SolidBrush(Color.FromArgb(245, 158, 11)))
                using (var badgePen = new Pen(Color.White, 1.2f))
                {
                    g.FillEllipse(badgeBrush, 17, 17, 13, 13);
                    g.DrawEllipse(badgePen, 17, 17, 13, 13);
                }

                // Biểu tượng tia chớp
                using (var boltBrush = new SolidBrush(Color.White))
                {
                    var pts = new PointF[]
                    {
                        new PointF(25f, 19f),
                        new PointF(21f, 24f),
                        new PointF(24f, 24f),
                        new PointF(22f, 28f),
                        new PointF(27f, 23f),
                        new PointF(24f, 23f)
                    };
                    g.FillPolygon(boltBrush, pts);
                }
            }
            return bmp;
        }

        private Bitmap CreateRulerBitmap()
        {
            var bmp = new Bitmap(32, 32);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.Clear(Color.Transparent);

                // Nền gradient xanh ngọc (Teal Gradient)
                using (var brush = new LinearGradientBrush(new Rectangle(2, 2, 28, 28),
                    Color.FromArgb(13, 148, 136), Color.FromArgb(15, 118, 110), 45f))
                {
                    FillRoundedRectangle(g, brush, new Rectangle(2, 2, 28, 28), 4);
                }

                // Viền ngoài
                using (var pen = new Pen(Color.FromArgb(17, 94, 89), 1f))
                {
                    DrawRoundedRectangle(g, pen, new Rectangle(2, 2, 28, 28), 4);
                }

                // Vẽ lưới ô bảng tính mờ
                using (var gridPen = new Pen(Color.FromArgb(90, 255, 255, 255), 1f))
                {
                    g.DrawRectangle(gridPen, 5, 5, 22, 22);
                    g.DrawLine(gridPen, 16, 5, 16, 27);
                    g.DrawLine(gridPen, 5, 16, 27, 16);
                }

                // Dải thước ngang (Row Ruler) màu vàng bán trong suốt
                using (var hBrush = new SolidBrush(Color.FromArgb(200, 253, 224, 71)))
                {
                    g.FillRectangle(hBrush, 4, 13, 24, 6);
                }

                // Dải thước dọc (Column Ruler) màu vàng bán trong suốt
                using (var vBrush = new SolidBrush(Color.FromArgb(200, 253, 224, 71)))
                {
                    g.FillRectangle(vBrush, 13, 4, 6, 24);
                }

                // Ô giao điểm (Active Cell Focus) viền đỏ cam rực rỡ
                using (var focusPen = new Pen(Color.FromArgb(239, 68, 68), 2f))
                {
                    g.DrawRectangle(focusPen, 13, 13, 6, 6);
                }

                // Điểm tâm chữ thập màu trắng
                using (var centerBrush = new SolidBrush(Color.White))
                {
                    g.FillEllipse(centerBrush, 14, 14, 4, 4);
                }
            }
            return bmp;
        }

        private Bitmap CreateRulerSettingsBitmap()
        {
            var bmp = new Bitmap(32, 32);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.Clear(Color.Transparent);

                // Nền gradient tím - hồng cao cấp
                using (var brush = new LinearGradientBrush(new Rectangle(2, 2, 28, 28),
                    Color.FromArgb(168, 85, 247), Color.FromArgb(236, 72, 153), 45f))
                {
                    FillRoundedRectangle(g, brush, new Rectangle(2, 2, 28, 28), 4);
                }

                using (var pen = new Pen(Color.FromArgb(147, 51, 234), 1f))
                {
                    DrawRoundedRectangle(g, pen, new Rectangle(2, 2, 28, 28), 4);
                }

                // Bảng màu tròn (Color Palette Shape)
                using (var paletteBrush = new SolidBrush(Color.White))
                {
                    g.FillEllipse(paletteBrush, 5, 5, 22, 22);
                }

                // Các chấm màu trên bảng màu
                using (var yBrush = new SolidBrush(Color.FromArgb(234, 179, 8))) // Vàng
                using (var bBrush = new SolidBrush(Color.FromArgb(2, 132, 199))) // Xanh dương
                using (var gBrush = new SolidBrush(Color.FromArgb(22, 163, 74))) // Xanh lá
                using (var rBrush = new SolidBrush(Color.FromArgb(239, 68, 68))) // Đỏ
                {
                    g.FillEllipse(yBrush, 9, 8, 5, 5);
                    g.FillEllipse(bBrush, 17, 8, 5, 5);
                    g.FillEllipse(gBrush, 9, 16, 5, 5);
                    g.FillEllipse(rBrush, 17, 16, 5, 5);
                }
            }
            return bmp;
        }

        private Bitmap CreateColorSwatchBitmap(Color fill, Color border)
        {
            var bmp = new Bitmap(16, 16);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.Clear(Color.Transparent);

                using (var b = new SolidBrush(fill))
                {
                    g.FillEllipse(b, 1, 1, 13, 13);
                }
                using (var p = new Pen(border, 1.5f))
                {
                    g.DrawEllipse(p, 1, 1, 13, 13);
                }
            }
            return bmp;
        }

        private Bitmap CreateModeIcon(string mode)
        {
            var bmp = new Bitmap(16, 16);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.Clear(Color.Transparent);

                using (var bgBrush = new SolidBrush(Color.FromArgb(241, 245, 249)))
                using (var borderPen = new Pen(Color.FromArgb(203, 213, 225), 1f))
                {
                    g.FillRectangle(bgBrush, 0, 0, 15, 15);
                    g.DrawRectangle(borderPen, 0, 0, 15, 15);
                }

                using (var hBrush = new SolidBrush(Color.FromArgb(253, 224, 71)))
                using (var linePen = new Pen(Color.FromArgb(234, 179, 8), 1.5f))
                {
                    if (mode == "Both" || mode == "Row")
                    {
                        g.FillRectangle(hBrush, 1, 6, 14, 4);
                        g.DrawLine(linePen, 1, 8, 14, 8);
                    }
                    if (mode == "Both" || mode == "Col")
                    {
                        g.FillRectangle(hBrush, 6, 1, 4, 14);
                        g.DrawLine(linePen, 8, 1, 8, 14);
                    }
                }
            }
            return bmp;
        }

        private Bitmap CreateBlankCleanerBitmap()
        {
            var bmp = new Bitmap(32, 32);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.Clear(Color.Transparent);

                // Nền gradient đỏ cam (Crimson - Red)
                using (var brush = new LinearGradientBrush(new Rectangle(2, 2, 28, 28),
                    Color.FromArgb(239, 68, 68), Color.FromArgb(185, 28, 28), 45f))
                {
                    FillRoundedRectangle(g, brush, new Rectangle(2, 2, 28, 28), 4);
                }

                using (var pen = new Pen(Color.FromArgb(153, 27, 27), 1f))
                {
                    DrawRoundedRectangle(g, pen, new Rectangle(2, 2, 28, 28), 4);
                }

                // Dòng 1 (có dữ liệu - màu trắng)
                using (var rowBrush = new SolidBrush(Color.White))
                {
                    g.FillRectangle(rowBrush, 6, 7, 20, 3);
                }

                // Dòng 2 (dòng trống bị đứt đoạn / mờ)
                using (var dashPen = new Pen(Color.FromArgb(254, 202, 202), 1.5f))
                {
                    dashPen.DashPattern = new float[] { 2, 2 };
                    g.DrawLine(dashPen, 6, 14, 26, 14);
                }

                // Dòng 3 (có dữ liệu - màu trắng)
                using (var rowBrush = new SolidBrush(Color.White))
                {
                    g.FillRectangle(rowBrush, 6, 19, 20, 3);
                }

                // Huy hiệu dấu trừ tròn đỏ nổi bật
                using (var badgeBrush = new SolidBrush(Color.FromArgb(220, 38, 38)))
                using (var badgePen = new Pen(Color.White, 1.2f))
                {
                    g.FillEllipse(badgeBrush, 17, 17, 13, 13);
                    g.DrawEllipse(badgePen, 17, 17, 13, 13);
                }

                // Biểu tượng dấu trừ (-) màu trắng
                using (var whitePen = new Pen(Color.White, 2f))
                {
                    g.DrawLine(whitePen, 20, 23, 27, 23);
                }
            }
            return bmp;
        }

        private Bitmap CreateMergeBitmap()
        {
            var bmp = new Bitmap(32, 32);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.Clear(Color.Transparent);

                // Nền gradient tím chàm (Indigo Gradient)
                using (var brush = new LinearGradientBrush(new Rectangle(2, 2, 28, 28),
                    Color.FromArgb(99, 102, 241), Color.FromArgb(67, 56, 202), 45f))
                {
                    FillRoundedRectangle(g, brush, new Rectangle(2, 2, 28, 28), 4);
                }

                using (var pen = new Pen(Color.FromArgb(55, 48, 163), 1f))
                {
                    DrawRoundedRectangle(g, pen, new Rectangle(2, 2, 28, 28), 4);
                }

                // 2 ô ghép lại thành 1 ô lớn ở giữa
                using (var cellBrush = new SolidBrush(Color.FromArgb(224, 231, 255)))
                using (var cellPen = new Pen(Color.White, 1.2f))
                {
                    g.FillRectangle(cellBrush, 5, 8, 22, 16);
                    g.DrawRectangle(cellPen, 5, 8, 22, 16);
                }

                // Hai mũi tên hướng vào nhau (--> <--) màu tím đậm
                using (var arrowPen = new Pen(Color.FromArgb(79, 70, 229), 2f))
                {
                    // Mũi tên trái ->
                    g.DrawLine(arrowPen, 8, 16, 13, 16);
                    g.DrawLine(arrowPen, 11, 13, 14, 16);
                    g.DrawLine(arrowPen, 11, 19, 14, 16);

                    // Mũi tên phải <-
                    g.DrawLine(arrowPen, 24, 16, 19, 16);
                    g.DrawLine(arrowPen, 21, 13, 18, 16);
                    g.DrawLine(arrowPen, 21, 19, 18, 16);
                }
            }
            return bmp;
        }

        private Bitmap CreateFindReplaceBitmap()
        {
            var bmp = new Bitmap(32, 32);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.Clear(Color.Transparent);

                // Nền gradient xanh dương đậm (#0284C7 -> #0369A1)
                using (var brush = new LinearGradientBrush(new Rectangle(2, 2, 28, 28),
                    Color.FromArgb(2, 132, 199), Color.FromArgb(3, 105, 161), 45f))
                {
                    FillRoundedRectangle(g, brush, new Rectangle(2, 2, 28, 28), 4);
                }

                using (var pen = new Pen(Color.FromArgb(2, 100, 150), 1f))
                {
                    DrawRoundedRectangle(g, pen, new Rectangle(2, 2, 28, 28), 4);
                }

                // Vẽ kính lúp (Search)
                using (var glassPen = new Pen(Color.White, 2f))
                {
                    g.DrawEllipse(glassPen, 6, 6, 10, 10);
                    g.DrawLine(glassPen, 14, 14, 19, 19);
                }

                // Vẽ 2 mũi tên xoay tròn thay thế (Refresh / Replace arrows)
                using (var arrowBrush = new SolidBrush(Color.FromArgb(254, 240, 138)))
                using (var arrowPen = new Pen(Color.FromArgb(254, 240, 138), 1.8f))
                {
                    // Vòng tròn mũi tên
                    g.DrawArc(arrowPen, 16, 14, 11, 11, 30, 240);
                    // Đầu mũi tên
                    PointF[] arrowHead = new PointF[]
                    {
                        new PointF(26, 13),
                        new PointF(29, 19),
                        new PointF(23, 18)
                    };
                    g.FillPolygon(arrowBrush, arrowHead);
                }
            }
            return bmp;
        }

        private Bitmap CreateAiFormulaBitmap()
        {
            var bmp = new Bitmap(32, 32);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.Clear(Color.Transparent);

                // Nền gradient Tím vũ trụ & Xanh AI (#7C3AED -> #4F46E5)
                using (var brush = new LinearGradientBrush(new Rectangle(2, 2, 28, 28),
                    Color.FromArgb(124, 58, 237), Color.FromArgb(79, 70, 229), 45f))
                {
                    FillRoundedRectangle(g, brush, new Rectangle(2, 2, 28, 28), 4);
                }

                using (var pen = new Pen(Color.FromArgb(67, 56, 202), 1f))
                {
                    DrawRoundedRectangle(g, pen, new Rectangle(2, 2, 28, 28), 4);
                }

                // Chữ fx màu vàng sáng
                using (var font = new System.Drawing.Font("Georgia", 11, System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))
                using (var fxBrush = new SolidBrush(Color.FromArgb(254, 240, 138)))
                {
                    g.DrawString("fx", font, fxBrush, 4, 7);
                }

                // Ngôi sao lấp lánh AI (Sparkle Star)
                using (var starBrush = new SolidBrush(Color.White))
                {
                    PointF[] star = new PointF[]
                    {
                        new PointF(22, 6),
                        new PointF(23.5f, 10.5f),
                        new PointF(28, 12),
                        new PointF(23.5f, 13.5f),
                        new PointF(22, 18),
                        new PointF(20.5f, 13.5f),
                        new PointF(16, 12),
                        new PointF(20.5f, 10.5f)
                    };
                    g.FillPolygon(starBrush, star);
                }

                // Ngôi sao nhỏ phụ
                using (var starSmall = new SolidBrush(Color.FromArgb(224, 231, 255)))
                {
                    g.FillEllipse(starSmall, 24, 21, 4, 4);
                }
            }
            return bmp;
        }

        private Bitmap CreateHudBitmap()
        {
            var bmp = new Bitmap(32, 32);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.Clear(Color.Transparent);

                // Nền gradient Dark Slate (#0F172A -> #1E293B)
                using (var brush = new LinearGradientBrush(new Rectangle(2, 2, 28, 28),
                    Color.FromArgb(15, 23, 42), Color.FromArgb(30, 41, 59), 45f))
                {
                    FillRoundedRectangle(g, brush, new Rectangle(2, 2, 28, 28), 4);
                }

                using (var pen = new Pen(Color.FromArgb(56, 189, 248), 1.2f))
                {
                    DrawRoundedRectangle(g, pen, new Rectangle(2, 2, 28, 28), 4);
                }

                // Cột biểu đồ thống kê màu xanh ngọc + xanh biển + vàng
                using (var bar1 = new SolidBrush(Color.FromArgb(56, 189, 248)))
                using (var bar2 = new SolidBrush(Color.FromArgb(74, 222, 128)))
                using (var bar3 = new SolidBrush(Color.FromArgb(250, 204, 21)))
                {
                    g.FillRectangle(bar1, 6, 16, 4, 9);
                    g.FillRectangle(bar2, 13, 10, 4, 15);
                    g.FillRectangle(bar3, 20, 6, 4, 19);
                }

                // Đường chỉ số thống kê màu trắng nối các đỉnh cột
                using (var linePen = new Pen(Color.White, 1.5f))
                {
                    g.DrawLine(linePen, 8, 16, 15, 10);
                    g.DrawLine(linePen, 15, 10, 22, 6);
                }
            }
            return bmp;
        }

        private static Bitmap CreateTableMergeBitmap()
        {
            var bmp = new Bitmap(32, 32);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);

                // Bảng trái (Màu xanh biển)
                using (var b1 = new SolidBrush(Color.FromArgb(37, 99, 235)))
                {
                    g.FillRectangle(b1, 3, 5, 11, 22);
                }

                // Bảng phải (Màu xanh lá)
                using (var b2 = new SolidBrush(Color.FromArgb(16, 124, 65)))
                {
                    g.FillRectangle(b2, 18, 5, 11, 22);
                }

                // Mũi tên ghép nối ở giữa màu cam
                using (var arrowBrush = new SolidBrush(Color.FromArgb(249, 115, 22)))
                {
                    PointF[] points = {
                        new PointF(12, 12),
                        new PointF(20, 16),
                        new PointF(12, 20)
                    };
                    g.FillPolygon(arrowBrush, points);
                }
            }
            return bmp;
        }

        private static Bitmap CreateFuzzyDuplicateBitmap()
        {
            var bmp = new Bitmap(32, 32);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);

                // Nền vòng tròn tím
                using (var bg = new SolidBrush(Color.FromArgb(147, 51, 234)))
                {
                    g.FillEllipse(bg, 3, 3, 26, 26);
                }

                // Dấu sóng xấp xỉ ~ (Fuzzy) màu trắng
                using (var font = new System.Drawing.Font("Arial", 16, System.Drawing.FontStyle.Bold))
                using (var textBrush = new SolidBrush(Color.White))
                {
                    g.DrawString("≈", font, textBrush, 6, 2);
                }
            }
            return bmp;
        }

        private static Bitmap CreateFileConverterBitmap()
        {
            var bmp = new Bitmap(32, 32);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);

                // Folder màu cam/vàng
                using (var folderBrush = new SolidBrush(Color.FromArgb(245, 158, 11)))
                {
                    g.FillRectangle(folderBrush, 3, 6, 26, 20);
                }

                // Huy hiệu chuyển đổi xanh ngọc ở góc dưới
                using (var badge = new SolidBrush(Color.FromArgb(14, 165, 233)))
                {
                    g.FillEllipse(badge, 14, 14, 16, 16);
                }

                using (var p = new Pen(Color.White, 2f))
                {
                    g.DrawArc(p, 17, 17, 10, 10, 0, 270);
                }
            }
            return bmp;
        }

        private static Bitmap CreatePageCounterBitmap()
        {
            var bmp = new Bitmap(32, 32);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);

                // Nền trang tài liệu
                using (var pageBrush = new SolidBrush(Color.FromArgb(241, 245, 249)))
                using (var borderPen = new Pen(Color.FromArgb(37, 99, 235), 1.5f))
                {
                    g.FillRectangle(pageBrush, 4, 3, 20, 26);
                    g.DrawRectangle(borderPen, 4, 3, 20, 26);
                }

                // Các dòng text mô phỏng
                using (var lineBrush = new SolidBrush(Color.FromArgb(148, 163, 184)))
                {
                    g.FillRectangle(lineBrush, 7, 7, 14, 2);
                    g.FillRectangle(lineBrush, 7, 11, 14, 2);
                    g.FillRectangle(lineBrush, 7, 15, 10, 2);
                }

                // Huy hiệu đếm số trang xanh lá ở góc phải
                using (var badge = new SolidBrush(Color.FromArgb(22, 163, 74)))
                {
                    g.FillEllipse(badge, 14, 14, 16, 16);
                }

                using (var font = new System.Drawing.Font("Arial", 8.5f, System.Drawing.FontStyle.Bold))
                using (var textBrush = new SolidBrush(Color.White))
                {
                    g.DrawString("12", font, textBrush, 14.5f, 15.5f);
                }
            }
            return bmp;
        }

        private static Bitmap CreateManualBitmap()
        {
            var bmp = new Bitmap(32, 32);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);

                // Cuốn sách mở màu xanh dương đậm + xanh ngọc
                using (var bookLeft = new SolidBrush(Color.FromArgb(37, 99, 235)))
                using (var bookRight = new SolidBrush(Color.FromArgb(14, 165, 233)))
                {
                    g.FillRectangle(bookLeft, 4, 6, 11, 20);
                    g.FillRectangle(bookRight, 17, 6, 11, 20);
                }

                // Gáy sách ở giữa
                using (var spinePen = new Pen(Color.White, 2f))
                {
                    g.DrawLine(spinePen, 15, 5, 15, 27);
                }

                // Các dòng chữ mô phỏng màu trắng
                using (var linePen = new Pen(Color.White, 1.2f))
                {
                    g.DrawLine(linePen, 6, 10, 13, 10);
                    g.DrawLine(linePen, 6, 14, 13, 14);
                    g.DrawLine(linePen, 6, 18, 13, 18);

                    g.DrawLine(linePen, 19, 10, 26, 10);
                    g.DrawLine(linePen, 19, 14, 26, 14);
                    g.DrawLine(linePen, 19, 18, 26, 18);
                }
            }
            return bmp;
        }

        private static Bitmap CreateFilteredPasteBitmap()
        {
            var bmp = new Bitmap(32, 32);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.Clear(Color.Transparent);

                // Bìa Clipboard màu xanh lục đậm #107C41
                using (var brush = new LinearGradientBrush(new Rectangle(3, 3, 26, 26),
                    Color.FromArgb(16, 124, 65), Color.FromArgb(10, 85, 42), 45f))
                {
                    FillRoundedRectangle(g, brush, new Rectangle(3, 3, 26, 26), 4);
                }

                // Tờ giấy trắng trên clipboard
                using (var paperBrush = new SolidBrush(Color.White))
                {
                    FillRoundedRectangle(g, paperBrush, new Rectangle(6, 7, 20, 20), 2);
                }

                // Kẹp giấy màu vàng cam ở trên
                using (var clipBrush = new SolidBrush(Color.FromArgb(245, 158, 11)))
                {
                    FillRoundedRectangle(g, clipBrush, new Rectangle(11, 2, 10, 5), 2);
                }

                // Dòng kẻ dữ liệu xám nhạt
                using (var lineBrush = new SolidBrush(Color.FromArgb(203, 213, 225)))
                {
                    g.FillRectangle(lineBrush, 8, 10, 12, 2);
                    g.FillRectangle(lineBrush, 8, 14, 10, 2);
                    g.FillRectangle(lineBrush, 8, 18, 8, 2);
                }

                // Biểu tượng phễu lọc màu xanh lam ở góc dưới bên phải
                using (var funnelBrush = new SolidBrush(Color.FromArgb(37, 99, 235)))
                {
                    var pts = new PointF[]
                    {
                        new PointF(16f, 15f),
                        new PointF(29f, 15f),
                        new PointF(24f, 21f),
                        new PointF(24f, 29f),
                        new PointF(21f, 27f),
                        new PointF(21f, 21f)
                    };
                    g.FillPolygon(funnelBrush, pts);
                }

                // Viền phễu trắng sắc nét
                using (var funnelPen = new Pen(Color.White, 1.2f))
                {
                    var pts = new PointF[]
                    {
                        new PointF(16f, 15f),
                        new PointF(29f, 15f),
                        new PointF(24f, 21f),
                        new PointF(24f, 29f),
                        new PointF(21f, 27f),
                        new PointF(21f, 21f)
                    };
                    g.DrawPolygon(funnelPen, pts);
                }
            }
            return bmp;
        }

        private static Bitmap CreateGlobeLangBitmap()
        {
            var bmp = new Bitmap(32, 32);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.Clear(Color.Transparent);

                // Quả địa cầu tròn màu xanh dương biển
                using (var globeBrush = new LinearGradientBrush(new Rectangle(2, 2, 28, 28),
                    Color.FromArgb(14, 165, 233), Color.FromArgb(2, 132, 199), 45f))
                {
                    g.FillEllipse(globeBrush, 2, 2, 28, 28);
                }

                // Viền ngoài địa cầu
                using (var borderPen = new Pen(Color.FromArgb(3, 105, 161), 1.2f))
                {
                    g.DrawEllipse(borderPen, 2, 2, 28, 28);
                }

                // Các đường kinh tuyến & vĩ tuyến màu trắng mờ
                using (var linePen = new Pen(Color.FromArgb(220, 255, 255, 255), 1.2f))
                {
                    // Xích đạo ngang
                    g.DrawLine(linePen, 2, 16, 30, 16);
                    // Vĩ tuyến trên và dưới
                    g.DrawArc(linePen, 4, 7, 24, 8, 0, 180);
                    g.DrawArc(linePen, 4, 17, 24, 8, 180, 180);
                    // Kinh tuyến đứng
                    g.DrawLine(linePen, 16, 2, 16, 30);
                    // Kinh tuyến cong
                    g.DrawEllipse(linePen, 9, 2, 14, 28);
                }

                // Huy hiệu bong bóng hội thoại ở góc dưới bên phải
                using (var badgeBrush = new SolidBrush(Color.FromArgb(245, 158, 11)))
                using (var badgePen = new Pen(Color.White, 1.2f))
                {
                    g.FillEllipse(badgeBrush, 16, 16, 14, 14);
                    g.DrawEllipse(badgePen, 16, 16, 14, 14);
                }

                // Chữ "A" nhỏ bên trong huy hiệu
                using (var font = new System.Drawing.Font("Segoe UI", 7.5f, FontStyle.Bold))
                using (var textBrush = new SolidBrush(Color.White))
                {
                    var sf = new StringFormat
                    {
                        Alignment = StringAlignment.Center,
                        LineAlignment = StringAlignment.Center
                    };
                    g.DrawString("A", font, textBrush, new RectangleF(16, 16, 14, 14), sf);
                }
            }
            return bmp;
        }

        private static Bitmap CreateFlagViBitmap()
        {
            var bmp = new Bitmap(24, 18);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.Clear(Color.Transparent);

                // Nền đỏ cờ Việt Nam
                using (var redBrush = new SolidBrush(Color.FromArgb(218, 37, 29)))
                {
                    FillRoundedRectangle(g, redBrush, new Rectangle(0, 0, 24, 18), 3);
                }
                using (var pen = new Pen(Color.FromArgb(180, 20, 20), 0.8f))
                {
                    DrawRoundedRectangle(g, pen, new Rectangle(0, 0, 23, 17), 3);
                }

                // Ngôi sao vàng 5 cánh ở giữa
                float cx = 12f, cy = 9f, rOuter = 5.5f, rInner = 2.2f;
                var pts = new PointF[10];
                for (int i = 0; i < 10; i++)
                {
                    double angle = -Math.PI / 2 + i * Math.PI / 5;
                    float r = (i % 2 == 0) ? rOuter : rInner;
                    pts[i] = new PointF(cx + (float)(r * Math.Cos(angle)), cy + (float)(r * Math.Sin(angle)));
                }

                using (var starBrush = new SolidBrush(Color.FromArgb(255, 255, 0)))
                {
                    g.FillPolygon(starBrush, pts);
                }
            }
            return bmp;
        }

        private static Bitmap CreateFlagEnBitmap()
        {
            var bmp = new Bitmap(24, 18);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.Clear(Color.Transparent);

                // Nền xanh Navy
                using (var blueBrush = new SolidBrush(Color.FromArgb(1, 33, 105)))
                {
                    FillRoundedRectangle(g, blueBrush, new Rectangle(0, 0, 24, 18), 3);
                }

                // Chữ thập chéo trắng
                using (var whiteDiagPen = new Pen(Color.White, 3f))
                {
                    g.DrawLine(whiteDiagPen, 0, 0, 24, 18);
                    g.DrawLine(whiteDiagPen, 0, 18, 24, 0);
                }
                // Chữ thập chéo đỏ
                using (var redDiagPen = new Pen(Color.FromArgb(200, 16, 46), 1.2f))
                {
                    g.DrawLine(redDiagPen, 0, 0, 24, 18);
                    g.DrawLine(redDiagPen, 0, 18, 24, 0);
                }

                // Chữ thập thẳng trắng
                using (var whiteCross = new SolidBrush(Color.White))
                {
                    g.FillRectangle(whiteCross, 9, 0, 6, 18);
                    g.FillRectangle(whiteCross, 0, 6, 24, 6);
                }
                // Chữ thập thẳng đỏ
                using (var redCross = new SolidBrush(Color.FromArgb(200, 16, 46)))
                {
                    g.FillRectangle(redCross, 10, 0, 4, 18);
                    g.FillRectangle(redCross, 0, 7, 24, 4);
                }

                using (var pen = new Pen(Color.FromArgb(150, 150, 150), 0.8f))
                {
                    DrawRoundedRectangle(g, pen, new Rectangle(0, 0, 23, 17), 3);
                }
            }
            return bmp;
        }

        private static Bitmap CreateFlagJaBitmap()
        {
            var bmp = new Bitmap(24, 18);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.Clear(Color.Transparent);

                // Nền trắng cờ Nhật Bản
                using (var whiteBrush = new SolidBrush(Color.White))
                {
                    FillRoundedRectangle(g, whiteBrush, new Rectangle(0, 0, 24, 18), 3);
                }
                using (var pen = new Pen(Color.FromArgb(203, 213, 225), 0.8f))
                {
                    DrawRoundedRectangle(g, pen, new Rectangle(0, 0, 23, 17), 3);
                }

                // Vòng tròn đỏ mặt trời ở giữa
                using (var sunBrush = new SolidBrush(Color.FromArgb(188, 0, 45)))
                {
                    g.FillEllipse(sunBrush, 7, 4, 10, 10);
                }
            }
            return bmp;
        }

        private static Bitmap CreateZenkakuBitmap()
        {
            var bmp = new Bitmap(32, 32);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.Clear(Color.Transparent);

                // Nền gradient xanh dương đậm
                using (var bgBrush = new LinearGradientBrush(new Rectangle(2, 2, 28, 28),
                    Color.FromArgb(30, 58, 138), Color.FromArgb(37, 99, 235), 45f))
                {
                    FillRoundedRectangle(g, bgBrush, new Rectangle(2, 2, 28, 28), 5);
                }
                using (var borderPen = new Pen(Color.FromArgb(29, 78, 216), 1.2f))
                {
                    DrawRoundedRectangle(g, borderPen, new Rectangle(2, 2, 28, 28), 5);
                }

                // Chữ "全" (Zenkaku) màu trắng bên trái
                using (var font = new System.Drawing.Font("Meiryo", 8.5f, FontStyle.Bold))
                using (var textBrush = new SolidBrush(Color.White))
                {
                    g.DrawString("全", font, textBrush, 3, 4);
                }

                // Mũi tên 2 chiều ⇋ màu vàng cam ở giữa
                using (var arrowPen = new Pen(Color.FromArgb(245, 158, 11), 1.5f))
                {
                    g.DrawLine(arrowPen, 7, 22, 24, 22);
                    g.DrawLine(arrowPen, 21, 20, 24, 22);
                    g.DrawLine(arrowPen, 21, 24, 24, 22);
                    g.DrawLine(arrowPen, 10, 20, 7, 22);
                    g.DrawLine(arrowPen, 10, 24, 7, 22);
                }

                // Chữ "半" (Hankaku) màu xanh ngọc bên phải
                using (var font = new System.Drawing.Font("Meiryo", 8.5f, FontStyle.Bold))
                using (var textBrush = new SolidBrush(Color.FromArgb(125, 211, 252)))
                {
                    g.DrawString("半", font, textBrush, 16, 4);
                }
            }
            return bmp;
        }

        private static Bitmap CreateKatakanaBitmap()
        {
            var bmp = new Bitmap(32, 32);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.Clear(Color.Transparent);

                // Nền tím Slate Indigo
                using (var bgBrush = new LinearGradientBrush(new Rectangle(2, 2, 28, 28),
                    Color.FromArgb(79, 70, 229), Color.FromArgb(67, 56, 202), 45f))
                {
                    FillRoundedRectangle(g, bgBrush, new Rectangle(2, 2, 28, 28), 5);
                }
                using (var borderPen = new Pen(Color.FromArgb(55, 48, 163), 1.2f))
                {
                    DrawRoundedRectangle(g, borderPen, new Rectangle(2, 2, 28, 28), 5);
                }

                // Chữ Katakana "ア" lớn màu trắng
                using (var font = new System.Drawing.Font("Meiryo", 12f, FontStyle.Bold))
                using (var textBrush = new SolidBrush(Color.White))
                {
                    var sf = new StringFormat
                    {
                        Alignment = StringAlignment.Center,
                        LineAlignment = StringAlignment.Center
                    };
                    g.DrawString("ア", font, textBrush, new RectangleF(2, 1, 28, 20), sf);
                }

                // Dấu kiểm checkmark xanh lục ngọc ở góc dưới
                using (var checkBg = new SolidBrush(Color.FromArgb(16, 185, 129)))
                using (var checkPen = new Pen(Color.White, 1.2f))
                {
                    g.FillEllipse(checkBg, 16, 16, 13, 13);
                    g.DrawEllipse(checkPen, 16, 16, 13, 13);
                }
                using (var checkIconPen = new Pen(Color.White, 1.6f))
                {
                    g.DrawLine(checkIconPen, 19, 23, 22, 26);
                    g.DrawLine(checkIconPen, 22, 26, 26, 19);
                }
            }
            return bmp;
        }

        private static Bitmap CreateMarkdownBitmap()
        {
            var bmp = new Bitmap(32, 32);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.Clear(Color.Transparent);

                // Nền Dark Slate #0F172A
                using (var bgBrush = new LinearGradientBrush(new Rectangle(2, 2, 28, 28),
                    Color.FromArgb(30, 41, 59), Color.FromArgb(15, 23, 42), 45f))
                {
                    FillRoundedRectangle(g, bgBrush, new Rectangle(2, 2, 28, 28), 5);
                }
                using (var borderPen = new Pen(Color.FromArgb(51, 65, 85), 1.2f))
                {
                    DrawRoundedRectangle(g, borderPen, new Rectangle(2, 2, 28, 28), 5);
                }

                // Chữ M phong cách Markdown
                using (var font = new System.Drawing.Font("Segoe UI", 11f, FontStyle.Bold))
                using (var textBrush = new SolidBrush(Color.White))
                {
                    g.DrawString("M", font, textBrush, 4, 3);
                }

                // Mũi tên xuống màu Sky Blue #38BDF8
                using (var arrowPen = new Pen(Color.FromArgb(56, 189, 248), 1.8f))
                {
                    g.DrawLine(arrowPen, 23, 6, 23, 16);
                    g.DrawLine(arrowPen, 20, 13, 23, 16);
                    g.DrawLine(arrowPen, 26, 13, 23, 16);
                }

                // Các vạch bảng Markdown ở đáy
                using (var gridPen = new Pen(Color.FromArgb(148, 163, 184), 1f))
                {
                    g.DrawLine(gridPen, 6, 22, 26, 22);
                    g.DrawLine(gridPen, 16, 20, 16, 25);
                }
            }
            return bmp;
        }
    }
}
