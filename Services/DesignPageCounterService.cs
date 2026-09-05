using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;

namespace ExcelSupport.Services
{
    #region Data Models & Enums

    public enum PageCounterMode
    {
        UserHighlightedColor, // Thuật toán đếm theo màu ô người dùng đã tô (Mặc định)
        AutoDiffTemplate,     // Thuật toán tự động so sánh với Template & Tô màu đối chiếu
        PrintBreakGrid        // Thuật toán ngắt trang in Excel
    }

    public enum PageStatus
    {
        WorkPage,       // Trang thiết kế thực tế (có nội dung mới/sửa)
        TemplatePage,   // Trang thuần template (không thay đổi)
        NewPage,        // Trang mới hoàn toàn (thuộc sheet mới hoặc mở rộng)
        BlankPage       // Trang trống không có dữ liệu
    }

    public enum SheetStatus
    {
        NewSheet,       // Sheet mới hoàn toàn không có trong template
        ModifiedSheet,  // Sheet có chỉnh sửa/thêm trang
        TemplateOnly,   // Sheet hoàn toàn trùng khớp template
        SkippedSheet    // Sheet bị bỏ qua theo cấu hình
    }

    public class PageCounterOptions
    {
        public PageCounterMode Mode { get; set; } = PageCounterMode.UserHighlightedColor;
        public bool IgnoreCoverAndHistory { get; set; } = true;
        public bool IgnoreBlankPages { get; set; } = true;
        public bool CountShapesAndPictures { get; set; } = true;
        public int MinChangedCellsThreshold { get; set; } = 2;
        public int CharactersPerPage { get; set; } = 600; // Định mức ký tự / trang (mặc định 600 cho Tiếng Nhật, 1200 cho Tiếng Việt/Anh)
        public double ShapePageFactor { get; set; } = 0.5; // Mỗi sơ đồ/hình vẽ lớn quy đổi = 0.5 trang
        public bool HighlightChangedCells { get; set; } = true; // Tạo bản sao và tô màu ô thay đổi (cho mode AutoDiff)
        public string HighlightColorHex { get; set; } = "ANY"; // "ANY" = Bất kỳ màu nào, hoặc mã Hex "#FEF08A"
        public bool MatchAnyHighlightColor { get; set; } = true; // Khác màu trắng và không màu
        public HashSet<string> ExcludedSheetNames { get; set; } = new(StringComparer.OrdinalIgnoreCase);
    }

    public class PageDetailItem
    {
        public int PageNumber { get; set; }
        public string RangeAddress { get; set; } = string.Empty;
        public int StartRow { get; set; }
        public int EndRow { get; set; }
        public int StartCol { get; set; }
        public int EndCol { get; set; }
        public PageStatus Status { get; set; } = PageStatus.TemplatePage;
        public int ChangedCellsCount { get; set; }
        public int AddedShapesCount { get; set; }
        public int TotalNonEmptyCells { get; set; }
        public string Description { get; set; } = string.Empty;

        public bool IsWorkPage => Status == PageStatus.WorkPage || Status == PageStatus.NewPage;
    }

    public class SheetPageCounterResult
    {
        public string SheetName { get; set; } = string.Empty;
        public SheetStatus Status { get; set; } = SheetStatus.TemplateOnly;
        public int TotalPrintPages { get; set; }
        public int TemplatePagesCount { get; set; }
        public double WorkPagesCount { get; set; }
        public int BlankPagesCount { get; set; }
        public int TotalChangedCells { get; set; }
        public int TotalChangedCharacters { get; set; }
        public int TotalAddedShapes { get; set; }
        public double WorkPercent => TotalPrintPages > 0 ? Math.Min(100.0, Math.Round(WorkPagesCount / TotalPrintPages * 100.0, 1)) : 0;
        public List<PageDetailItem> Pages { get; set; } = new();

        public string SummaryText
        {
            get
            {
                if (Status == SheetStatus.SkippedSheet)
                    return LocalizationService.Get("PageCounter_StatusSkipped");
                if (Status == SheetStatus.NewSheet)
                    return LocalizationService.Get("PageCounter_StatusNewSheet", WorkPagesCount, TotalChangedCharacters);
                if (Status == SheetStatus.TemplateOnly)
                    return LocalizationService.Get("PageCounter_StatusTemplateOnly", TemplatePagesCount);
                return LocalizationService.Get("PageCounter_StatusModified", WorkPagesCount, TotalChangedCharacters, TotalAddedShapes);
            }
        }
    }

    public class WorkbookPageCounterResult
    {
        public string TargetWorkbookName { get; set; } = string.Empty;
        public string TargetWorkbookPath { get; set; } = string.Empty;
        public string TemplateWorkbookName { get; set; } = string.Empty;
        public string TemplateWorkbookPath { get; set; } = string.Empty;
        public string? HighlightedClonedWorkbookPath { get; set; }
        public DateTime AnalyzedAt { get; set; } = DateTime.Now;

        public int TotalTargetPrintPages { get; set; }
        public int TotalTemplatePages { get; set; }
        public double TotalWorkPages { get; set; }
        public int TotalBlankPages { get; set; }
        public int TotalChangedCells { get; set; }
        public int TotalChangedCharacters { get; set; }
        public int TotalAddedShapes { get; set; }
        public int TotalNewSheetsCount { get; set; }
        public int TotalModifiedSheetsCount { get; set; }
        public int TotalUnchangedSheetsCount { get; set; }

        public double OverallWorkPercent => TotalTargetPrintPages > 0
            ? Math.Min(100.0, Math.Round(TotalWorkPages / TotalTargetPrintPages * 100.0, 1))
            : 0;

        public List<SheetPageCounterResult> SheetResults { get; set; } = new();
    }

    #endregion

    public static class DesignPageCounterService
    {
        private static readonly string[] CoverAndHistoryPatterns = new[]
        {
            "cover", "bìa", "bia", "history", "lịch sử", "lich su", "履歴", "改訂", "表紙", "目次", "toc", "guide", "hướng dẫn", "huong dan"
        };

        /// <summary>
        /// Phân tích và đếm số trang thiết kế thực tế giữa Target Workbook và Template Workbook.
        /// </summary>
        public static WorkbookPageCounterResult AnalyzePages(
            ExcelApp app,
            string targetWbNameOrPath,
            string templateWbNameOrPath,
            PageCounterOptions options,
            Action<string, int>? progressCallback = null)
        {
            if (app == null) throw new ArgumentNullException(nameof(app));

            Workbook? targetWb = null;
            Workbook? templateWb = null;
            Workbook? clonedWb = null;
            string? tempClonePath = null;

            bool targetOpenedHere = false;
            bool templateOpenedHere = false;

            var result = new WorkbookPageCounterResult
            {
                TargetWorkbookName = Path.GetFileName(targetWbNameOrPath),
                TargetWorkbookPath = targetWbNameOrPath,
                TemplateWorkbookName = string.IsNullOrWhiteSpace(templateWbNameOrPath) ? "(Không dùng template)" : Path.GetFileName(templateWbNameOrPath),
                TemplateWorkbookPath = templateWbNameOrPath
            };

            bool prevScreenUpdating = app.ScreenUpdating;
            bool prevDisplayAlerts = app.DisplayAlerts;

            try
            {
                app.ScreenUpdating = false;
                app.DisplayAlerts = false;

                // 1. Resolve Target Workbook
                progressCallback?.Invoke("Đang mở và tải tài liệu thiết kế...", 5);
                targetWb = FindOrOpenWorkbook(app, targetWbNameOrPath, out targetOpenedHere);
                if (targetWb == null)
                    throw new FileNotFoundException($"Không thể mở file thiết kế: {targetWbNameOrPath}");

                // 2. Tạo bản sao tạm thời (Clone) nếu người dùng chọn tính năng tô màu (Highlight)
                if (options.HighlightChangedCells)
                {
                    try
                    {
                        string tempDir = Path.Combine(Path.GetTempPath(), "ExcelSupport_DesignPages");
                        if (!Directory.Exists(tempDir)) Directory.CreateDirectory(tempDir);
                        string baseName = Path.GetFileNameWithoutExtension(targetWb.Name);
                        string ext = Path.GetExtension(targetWb.FullName);
                        if (string.IsNullOrEmpty(ext)) ext = ".xlsx";
                        tempClonePath = Path.Combine(tempDir, $"Evidence_{baseName}_{DateTime.Now:yyyyMMdd_HHmmss}{ext}");

                        if (File.Exists(targetWb.FullName))
                        {
                            File.Copy(targetWb.FullName, tempClonePath, true);
                        }
                        else
                        {
                            targetWb.SaveCopyAs(tempClonePath);
                        }

                        if (File.Exists(tempClonePath))
                        {
                            clonedWb = app.Workbooks.Open(tempClonePath);
                        }
                    }
                    catch { }
                }

                // 3. Resolve Template Workbook (nếu có)
                if (!string.IsNullOrWhiteSpace(templateWbNameOrPath))
                {
                    progressCallback?.Invoke("Đang mở và tải template gốc...", 15);
                    templateWb = FindOrOpenWorkbook(app, templateWbNameOrPath, out templateOpenedHere);
                }

                // 4. Xây dựng danh sách Sheets của Template để so khớp nhanh
                var templateSheetsDict = new Dictionary<string, Worksheet>(StringComparer.OrdinalIgnoreCase);
                if (templateWb != null)
                {
                    foreach (Worksheet ws in templateWb.Worksheets)
                    {
                        templateSheetsDict[ws.Name] = ws;
                    }
                }

                var clonedSheetsDict = new Dictionary<string, Worksheet>(StringComparer.OrdinalIgnoreCase);
                if (clonedWb != null)
                {
                    foreach (Worksheet ws in clonedWb.Worksheets)
                    {
                        clonedSheetsDict[ws.Name] = ws;
                    }
                }

                int sheetCount = targetWb.Worksheets.Count;
                int currentSheetIndex = 0;

                // 5. Lặp qua từng Sheet trong Target Workbook
                foreach (Worksheet targetWs in targetWb.Worksheets)
                {
                    currentSheetIndex++;
                    int progress = 20 + (int)((double)currentSheetIndex / Math.Max(1, sheetCount) * 70);
                    progressCallback?.Invoke($"Đang phân tích và tô màu sheet: {targetWs.Name} ({currentSheetIndex}/{sheetCount})...", progress);

                    bool isExplicitlyExcluded = options.ExcludedSheetNames != null && options.ExcludedSheetNames.Contains(targetWs.Name);
                    bool isCoverOrHistory = IsCoverOrHistorySheet(targetWs.Name);
                    if (isExplicitlyExcluded || (options.IgnoreCoverAndHistory && isCoverOrHistory))
                    {
                        var skippedResult = new SheetPageCounterResult
                        {
                            SheetName = targetWs.Name,
                            Status = SheetStatus.SkippedSheet,
                            TotalPrintPages = 1,
                            TemplatePagesCount = 1,
                            WorkPagesCount = 0
                        };
                        result.SheetResults.Add(skippedResult);
                        result.TotalTargetPrintPages += 1;
                        result.TotalTemplatePages += 1;
                        result.TotalUnchangedSheetsCount++;
                        continue;
                    }

                    templateSheetsDict.TryGetValue(targetWs.Name, out Worksheet? tplWs);
                    clonedSheetsDict.TryGetValue(targetWs.Name, out Worksheet? clonedWs);

                    SheetPageCounterResult sheetResult;
                    if (options.Mode == PageCounterMode.UserHighlightedColor)
                    {
                        sheetResult = AnalyzeWorksheetByHighlightedColor(app, targetWs, options);
                    }
                    else if (options.Mode == PageCounterMode.AutoDiffTemplate)
                    {
                        sheetResult = AnalyzeWorksheetByCharacterAndHighlight(app, targetWs, tplWs, clonedWs, options);
                    }
                    else
                    {
                        sheetResult = AnalyzeWorksheetPages(app, targetWs, tplWs, options);
                    }

                    result.SheetResults.Add(sheetResult);

                    result.TotalTargetPrintPages += sheetResult.TotalPrintPages;
                    result.TotalTemplatePages += sheetResult.TemplatePagesCount;
                    result.TotalWorkPages += sheetResult.WorkPagesCount;
                    result.TotalBlankPages += sheetResult.BlankPagesCount;
                    result.TotalChangedCells += sheetResult.TotalChangedCells;
                    result.TotalChangedCharacters += sheetResult.TotalChangedCharacters;
                    result.TotalAddedShapes += sheetResult.TotalAddedShapes;

                    if (sheetResult.Status == SheetStatus.NewSheet)
                        result.TotalNewSheetsCount++;
                    else if (sheetResult.Status == SheetStatus.ModifiedSheet)
                        result.TotalModifiedSheetsCount++;
                    else if (sheetResult.Status == SheetStatus.TemplateOnly)
                        result.TotalUnchangedSheetsCount++;
                }

                // Lưu lại file cloned workbook nếu có
                if (clonedWb != null)
                {
                    try
                    {
                        clonedWb.Save();
                        clonedWb.Close(false);
                        Marshal.ReleaseComObject(clonedWb);
                        clonedWb = null;
                        result.HighlightedClonedWorkbookPath = tempClonePath;
                    }
                    catch { }
                }

                progressCallback?.Invoke("Hoàn tất phân tích!", 100);
                return result;
            }
            finally
            {
                app.ScreenUpdating = prevScreenUpdating;
                app.DisplayAlerts = prevDisplayAlerts;

                if (clonedWb != null)
                {
                    try { clonedWb.Close(false); Marshal.ReleaseComObject(clonedWb); } catch { }
                }

                // Dọn dẹp các workbook nếu được mở tạm thời
                if (targetOpenedHere && targetWb != null)
                {
                    try { targetWb.Close(false); Marshal.ReleaseComObject(targetWb); } catch { }
                }
                if (templateOpenedHere && templateWb != null)
                {
                    try { templateWb.Close(false); Marshal.ReleaseComObject(templateWb); } catch { }
                }
            }
        }

        public static string CurrentHighlightColorHex { get; set; } = "#FEF08A";

        /// <summary>
        /// Tô màu vùng ô đang được chọn trong Excel theo màu chỉ định (phím tắt Ctrl + Shift + H).
        /// </summary>
        public static bool HighlightSelection(ExcelApp? excelApp = null, string? hexColor = null)
        {
            try
            {
                var app = excelApp ?? AddInEvents.Instance?.ExcelAppInstance ?? (ExcelApp)ExcelDna.Integration.ExcelDnaUtil.Application;
                if (app == null) return false;

                dynamic selection = app.Selection;
                if (selection == null) return false;

                string targetHex = hexColor ?? CurrentHighlightColorHex;
                if (string.IsNullOrEmpty(targetHex) || targetHex.Equals("ANY", StringComparison.OrdinalIgnoreCase))
                {
                    targetHex = "#FEF08A"; // Default to vivid yellow pastel
                }

                int r = Convert.ToInt32(targetHex.Substring(1, 2), 16);
                int g = Convert.ToInt32(targetHex.Substring(3, 2), 16);
                int b = Convert.ToInt32(targetHex.Substring(5, 2), 16);
                int oleColor = r | (g << 8) | (b << 16);

                if (selection is Range rng)
                {
                    rng.Interior.Color = oleColor;
                }
                else
                {
                    try
                    {
                        selection.Interior.Color = oleColor;
                    }
                    catch { }
                }

                try
                {
                    app.StatusBar = $"🎨 ExcelSupport: Đã tô màu đánh dấu thiết kế [{targetHex}] (Ctrl + Shift + H)";
                }
                catch { }

                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Xóa màu đánh dấu của các ô đang được chọn trong Excel (phím tắt Ctrl + Shift + Alt + H).
        /// </summary>
        public static bool ClearHighlightSelection(ExcelApp? excelApp = null)
        {
            try
            {
                var app = excelApp ?? AddInEvents.Instance?.ExcelAppInstance ?? (ExcelApp)ExcelDna.Integration.ExcelDnaUtil.Application;
                if (app == null) return false;

                dynamic selection = app.Selection;
                if (selection == null) return false;

                if (selection is Range rng)
                {
                    rng.Interior.ColorIndex = -4142; // xlColorIndexNone
                }
                else
                {
                    try
                    {
                        selection.Interior.ColorIndex = -4142;
                    }
                    catch { }
                }

                try
                {
                    app.StatusBar = "🧹 ExcelSupport: Đã xóa màu đánh dấu vùng chọn (Ctrl + Shift + Alt + H)";
                }
                catch { }

                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Tạo và mở một bản sao mới của file thiết kế để người dùng tự do tô màu các ô trước khi đếm.
        /// </summary>
        public static string? CreateAndOpenNewCopyForHighlighting(ExcelApp app, string targetWbNameOrPath)
        {
            if (app == null || string.IsNullOrWhiteSpace(targetWbNameOrPath)) return null;

            try
            {
                string tempDir = Path.Combine(Path.GetTempPath(), "ExcelSupport_DesignPages");
                if (!Directory.Exists(tempDir)) Directory.CreateDirectory(tempDir);

                string baseName = Path.GetFileNameWithoutExtension(targetWbNameOrPath);
                string ext = Path.GetExtension(targetWbNameOrPath);
                if (string.IsNullOrEmpty(ext)) ext = ".xlsx";

                string newCopyPath = Path.Combine(tempDir, $"New_Design_{baseName}_{DateTime.Now:yyyyMMdd_HHmmss}{ext}");

                bool openedHere = false;
                Workbook? targetWb = FindOrOpenWorkbook(app, targetWbNameOrPath, out openedHere);
                if (targetWb == null) return null;

                try
                {
                    if (File.Exists(targetWb.FullName))
                    {
                        File.Copy(targetWb.FullName, newCopyPath, true);
                    }
                    else
                    {
                        targetWb.SaveCopyAs(newCopyPath);
                    }

                    app.Visible = true;
                    Workbook newWb = app.Workbooks.Open(newCopyPath);
                    newWb.Activate();
                    return newCopyPath;
                }
                finally
                {
                    if (openedHere && targetWb != null)
                    {
                        try { targetWb.Close(false); Marshal.ReleaseComObject(targetWb); } catch { }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DesignPageCounterService] CreateAndOpenNewCopy error: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Mở workbook đã được tô màu highlight (Evidence) trong Excel để người dùng xem trực tiếp.
        /// </summary>
        public static bool OpenEvidenceWorkbook(ExcelApp app, string filePath)
        {
            if (app == null || string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath)) return false;
            try
            {
                app.Visible = true;
                var wb = app.Workbooks.Open(filePath);
                wb.Activate();
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Phân tích số trang thiết kế dựa trên màu sắc mà người dùng đã tự tô vào các ô.
        /// </summary>
        private static SheetPageCounterResult AnalyzeWorksheetByHighlightedColor(
            ExcelApp app,
            Worksheet targetWs,
            PageCounterOptions options)
        {
            var sheetResult = new SheetPageCounterResult
            {
                SheetName = targetWs.Name
            };

            // 1. Lấy UsedRange của Target
            Range? targetUsed = null;
            int targetStartRow = 1;
            int targetStartCol = 1;
            int targetRowCount = 1;
            int targetColCount = 1;

            try
            {
                targetUsed = targetWs.UsedRange;
                if (targetUsed != null)
                {
                    targetStartRow = targetUsed.Row;
                    targetStartCol = targetUsed.Column;
                    targetRowCount = targetUsed.Rows.Count;
                    targetColCount = targetUsed.Columns.Count;
                }
            }
            catch { }

            object?[,] targetVals = Extract2DArray(targetUsed?.Value2, targetRowCount, targetColCount);
            object?[,] targetFormulas = Extract2DArray(targetUsed?.Formula, targetRowCount, targetColCount);

            Color targetColor = ColorTranslator.FromHtml(options.HighlightColorHex.Equals("ANY", StringComparison.OrdinalIgnoreCase) ? "#FEF08A" : options.HighlightColorHex);
            bool matchAnyColor = options.MatchAnyHighlightColor || options.HighlightColorHex.Equals("ANY", StringComparison.OrdinalIgnoreCase);

            var highlightedCells = new List<(int row, int col, int charCount)>();
            int totalChars = 0;
            int totalCells = 0;

            for (int r = 1; r <= targetRowCount; r++)
            {
                int actualRow = targetStartRow + r - 1;
                for (int c = 1; c <= targetColCount; c++)
                {
                    int actualCol = targetStartCol + c - 1;

                    object? tVal = targetVals[r, c];
                    object? tForm = targetFormulas[r, c];

                    string tStr = tVal?.ToString()?.Trim() ?? string.Empty;
                    string tFormStr = tForm?.ToString()?.Trim() ?? string.Empty;

                    if (string.IsNullOrEmpty(tStr) && string.IsNullOrEmpty(tFormStr))
                        continue;

                    // Kiểm tra xem ô có được tô màu hay không
                    Range? cell = null;
                    bool isHighlighted = false;
                    try
                    {
                        cell = targetWs.Cells[actualRow, actualCol] as Range;
                        if (cell != null)
                        {
                            object interiorColor = cell.Interior.Color;
                            object interiorColorIndex = cell.Interior.ColorIndex;
                            isHighlighted = IsColorMatch(interiorColor, interiorColorIndex, targetColor, matchAnyColor);
                        }
                    }
                    catch { }
                    finally
                    {
                        if (cell != null) Marshal.ReleaseComObject(cell);
                    }

                    if (isHighlighted)
                    {
                        int len = tStr.Length > 0 ? tStr.Length : tFormStr.Length;
                        totalCells++;
                        totalChars += len;
                        highlightedCells.Add((actualRow, actualCol, len));
                    }
                }
            }

            // 2. Đếm các hình vẽ / sơ đồ hợp lệ (kích thước lớn)
            int addedShapes = 0;
            if (options.CountShapesAndPictures)
            {
                addedShapes = CountMeaningfulDesignShapes(targetWs);
            }

            sheetResult.TotalChangedCells = totalCells;
            sheetResult.TotalChangedCharacters = totalChars;
            sheetResult.TotalAddedShapes = addedShapes;

            // 3. Quy đổi ra số trang thiết kế
            double charPages = (double)totalChars / Math.Max(1, options.CharactersPerPage);
            double shapePages = addedShapes * options.ShapePageFactor;
            double calculatedWorkPages = Math.Round(charPages + shapePages, 1);

            sheetResult.TotalPrintPages = Math.Max(1, (int)Math.Ceiling(calculatedWorkPages > 0 ? calculatedWorkPages : 1));
            sheetResult.WorkPagesCount = calculatedWorkPages;
            sheetResult.TemplatePagesCount = 0;

            if (totalCells == 0 && addedShapes == 0)
            {
                sheetResult.Status = SheetStatus.TemplateOnly;
                sheetResult.WorkPagesCount = 0;
                sheetResult.TemplatePagesCount = sheetResult.TotalPrintPages;
            }
            else
            {
                sheetResult.Status = SheetStatus.ModifiedSheet;
            }

            // 4. Xây dựng danh sách trang chi tiết
            var targetPageRanges = GetPrintPageRanges(app, targetWs);
            if (targetPageRanges.Count > 0)
            {
                sheetResult.TotalPrintPages = targetPageRanges.Count;
                sheetResult.TemplatePagesCount = totalCells == 0 && addedShapes == 0 ? targetPageRanges.Count : 0;
                int pageIndex = 1;
                foreach (var pr in targetPageRanges)
                {
                    int pageHighlightedCount = highlightedCells.Count(c => c.row >= pr.StartRow && c.row <= pr.EndRow && c.col >= pr.StartCol && c.col <= pr.EndCol);
                    int pageChars = highlightedCells.Where(c => c.row >= pr.StartRow && c.row <= pr.EndRow && c.col >= pr.StartCol && c.col <= pr.EndCol).Sum(c => c.charCount);
                    bool isWorkPage = pageHighlightedCount > 0;
                    sheetResult.Pages.Add(new PageDetailItem
                    {
                        PageNumber = pageIndex++,
                        RangeAddress = pr.Address,
                        StartRow = pr.StartRow,
                        EndRow = pr.EndRow,
                        StartCol = pr.StartCol,
                        EndCol = pr.EndCol,
                        Status = isWorkPage ? PageStatus.WorkPage : PageStatus.TemplatePage,
                        ChangedCellsCount = pageHighlightedCount,
                        Description = isWorkPage ? $"{pageHighlightedCount} ô tô màu ({pageChars} ký tự)" : "Không có ô tô màu"
                    });
                }
            }
            else
            {
                string addr = targetUsed != null ? targetUsed.Address : "A1";
                sheetResult.Pages.Add(new PageDetailItem
                {
                    PageNumber = 1,
                    RangeAddress = addr,
                    StartRow = targetStartRow,
                    EndRow = targetStartRow + targetRowCount - 1,
                    StartCol = targetStartCol,
                    EndCol = targetStartCol + targetColCount - 1,
                    Status = totalCells > 0 ? PageStatus.WorkPage : PageStatus.TemplatePage,
                    ChangedCellsCount = totalCells,
                    Description = totalCells > 0 ? $"{totalCells} ô tô màu ({totalChars} ký tự)" : "Không có ô tô màu"
                });
            }

            return sheetResult;
        }

        /// <summary>
        /// Phân tích số trang thiết kế bằng thuật toán tự động so sánh với Template & tô màu trực quan các ô thay đổi.
        /// </summary>
        private static SheetPageCounterResult AnalyzeWorksheetByCharacterAndHighlight(
            ExcelApp app,
            Worksheet targetWs,
            Worksheet? templateWs,
            Worksheet? clonedWs,
            PageCounterOptions options)
        {
            var sheetResult = new SheetPageCounterResult
            {
                SheetName = targetWs.Name
            };

            // 1. Tổng số trang in thực tế của sheet để làm cơ sở hiển thị
            var targetPageRanges = GetPrintPageRanges(app, targetWs);
            sheetResult.TotalPrintPages = Math.Max(1, targetPageRanges.Count);

            // 2. Lấy UsedRange của Target và Template
            Range? targetUsed = null;
            int targetStartRow = 1;
            int targetStartCol = 1;
            int targetRowCount = 1;
            int targetColCount = 1;

            try
            {
                targetUsed = targetWs.UsedRange;
                if (targetUsed != null)
                {
                    targetStartRow = targetUsed.Row;
                    targetStartCol = targetUsed.Column;
                    targetRowCount = targetUsed.Rows.Count;
                    targetColCount = targetUsed.Columns.Count;
                }
            }
            catch { }

            object?[,] targetVals = Extract2DArray(targetUsed?.Value2, targetRowCount, targetColCount);
            object?[,] targetFormulas = Extract2DArray(targetUsed?.Formula, targetRowCount, targetColCount);

            object?[,]? tplVals = null;
            object?[,]? tplFormulas = null;
            int tplStartRow = 1, tplStartCol = 1;
            int tplRowCount = 0, tplColCount = 0;

            if (templateWs != null)
            {
                try
                {
                    Range tplUsed = templateWs.UsedRange;
                    if (tplUsed != null)
                    {
                        tplStartRow = tplUsed.Row;
                        tplStartCol = tplUsed.Column;
                        tplRowCount = tplUsed.Rows.Count;
                        tplColCount = tplUsed.Columns.Count;
                        tplVals = Extract2DArray(tplUsed.Value2, tplRowCount, tplColCount);
                        tplFormulas = Extract2DArray(tplUsed.Formula, tplRowCount, tplColCount);
                    }
                }
                catch { }
            }

            var changedCells = new List<(int row, int col)>();
            int totalChars = 0;
            int totalCells = 0;

            for (int r = 1; r <= targetRowCount; r++)
            {
                int actualRow = targetStartRow + r - 1;
                for (int c = 1; c <= targetColCount; c++)
                {
                    int actualCol = targetStartCol + c - 1;

                    object? tVal = targetVals[r, c];
                    object? tForm = targetFormulas[r, c];

                    string tStr = tVal?.ToString()?.Trim() ?? string.Empty;
                    string tFormStr = tForm?.ToString()?.Trim() ?? string.Empty;

                    if (string.IsNullOrEmpty(tStr) && string.IsNullOrEmpty(tFormStr))
                        continue;

                    // Đối chiếu với ô tương ứng trên template
                    string tplStr = string.Empty;
                    string tplFormStr = string.Empty;

                    if (tplVals != null)
                    {
                        int tplR = actualRow - tplStartRow + 1;
                        int tplC = actualCol - tplStartCol + 1;
                        if (tplR >= 1 && tplR <= tplRowCount && tplC >= 1 && tplC <= tplColCount)
                        {
                            object? tpV = tplVals[tplR, tplC];
                            object? tpF = tplFormulas != null ? tplFormulas[tplR, tplC] : null;
                            tplStr = tpV?.ToString()?.Trim() ?? string.Empty;
                            tplFormStr = tpF?.ToString()?.Trim() ?? string.Empty;
                        }
                    }

                    bool isDiff = !string.Equals(tStr, tplStr, StringComparison.Ordinal) ||
                                  !string.Equals(tFormStr, tplFormStr, StringComparison.Ordinal);

                    if (isDiff)
                    {
                        totalCells++;
                        totalChars += tStr.Length;
                        changedCells.Add((actualRow, actualCol));
                    }
                }
            }

            // 3. Đếm số sơ đồ / hình vẽ hợp lệ mới thêm
            int addedShapes = 0;
            if (options.CountShapesAndPictures)
            {
                int targetShapes = CountMeaningfulDesignShapes(targetWs);
                int tplShapes = 0;
                if (templateWs != null)
                {
                    tplShapes = CountMeaningfulDesignShapes(templateWs);
                }
                addedShapes = Math.Max(0, targetShapes - tplShapes);
            }

            // 4. Tô màu các ô thay đổi trên bản sao (Cloned Worksheet)
            if (clonedWs != null && changedCells.Count > 0 && options.HighlightChangedCells)
            {
                try
                {
                    Color highlightColor = ColorTranslator.FromHtml(options.HighlightColorHex.Equals("ANY", StringComparison.OrdinalIgnoreCase) ? "#FEF08A" : options.HighlightColorHex);
                    ApplyHighlightToCells(clonedWs, changedCells, highlightColor);
                }
                catch { }
            }

            sheetResult.TotalChangedCells = totalCells;
            sheetResult.TotalChangedCharacters = totalChars;
            sheetResult.TotalAddedShapes = addedShapes;

            // 5. Quy đổi ra số trang thiết kế
            double charPages = (double)totalChars / Math.Max(1, options.CharactersPerPage);
            double shapePages = addedShapes * options.ShapePageFactor;
            double calculatedWorkPages = Math.Round(charPages + shapePages, 1);

            if (totalCells == 0 && addedShapes == 0)
            {
                sheetResult.Status = SheetStatus.TemplateOnly;
                sheetResult.WorkPagesCount = 0;
                sheetResult.TemplatePagesCount = sheetResult.TotalPrintPages;
            }
            else if (templateWs == null)
            {
                sheetResult.Status = SheetStatus.NewSheet;
                sheetResult.WorkPagesCount = Math.Max(1.0, calculatedWorkPages);
                sheetResult.TemplatePagesCount = 0;
            }
            else
            {
                sheetResult.Status = SheetStatus.ModifiedSheet;
                sheetResult.WorkPagesCount = calculatedWorkPages;
                sheetResult.TemplatePagesCount = Math.Max(0, (int)Math.Ceiling(sheetResult.TotalPrintPages - calculatedWorkPages));
            }

            // 6. Xây dựng danh sách trang chi tiết
            int pageIndex = 1;
            foreach (var pr in targetPageRanges)
            {
                int pageChangedCells = changedCells.Count(c => c.row >= pr.StartRow && c.row <= pr.EndRow && c.col >= pr.StartCol && c.col <= pr.EndCol);
                bool isWorkPage = pageChangedCells > 0;
                sheetResult.Pages.Add(new PageDetailItem
                {
                    PageNumber = pageIndex++,
                    RangeAddress = pr.Address,
                    StartRow = pr.StartRow,
                    EndRow = pr.EndRow,
                    StartCol = pr.StartCol,
                    EndCol = pr.EndCol,
                    Status = isWorkPage ? (templateWs == null ? PageStatus.NewPage : PageStatus.WorkPage) : PageStatus.TemplatePage,
                    ChangedCellsCount = pageChangedCells,
                    Description = isWorkPage ? $"{pageChangedCells} ô thay đổi (Đã tô màu)" : "Nguyên bản template"
                });
            }

            return sheetResult;
        }

        /// <summary>
        /// Tô màu nền hàng loạt cho danh sách các ô bằng cách gom nhóm địa chỉ Range.
        /// </summary>
        private static void ApplyHighlightToCells(Worksheet ws, List<(int row, int col)> cells, Color color)
        {
            if (cells == null || cells.Count == 0) return;
            int oleColor = ColorTranslator.ToOle(color);

            var addressChunks = new List<string>();
            var currentChunk = new List<string>();

            foreach (var (r, c) in cells)
            {
                string addr = GetExcelCellAddress(r, c);
                currentChunk.Add(addr);
                if (currentChunk.Count >= 25)
                {
                    addressChunks.Add(string.Join(",", currentChunk));
                    currentChunk.Clear();
                }
            }
            if (currentChunk.Count > 0)
            {
                addressChunks.Add(string.Join(",", currentChunk));
            }

            foreach (var chunk in addressChunks)
            {
                try
                {
                    Range rng = ws.Range[chunk];
                    rng.Interior.Color = oleColor;
                    Marshal.ReleaseComObject(rng);
                }
                catch { }
            }
        }

        private static string GetExcelCellAddress(int row, int col)
        {
            int dividend = col;
            string colName = string.Empty;
            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                colName = Convert.ToChar(65 + modulo) + colName;
                dividend = (dividend - modulo) / 26;
            }
            return $"{colName}{row}";
        }

        /// <summary>
        /// Phân tích danh sách trang in của một Worksheet cụ thể.
        /// </summary>
        private static SheetPageCounterResult AnalyzeWorksheetPages(
            ExcelApp app,
            Worksheet targetWs,
            Worksheet? templateWs,
            PageCounterOptions options)
        {
            var sheetResult = new SheetPageCounterResult
            {
                SheetName = targetWs.Name
            };

            // 1. Phân rã trang in của Target Sheet
            var targetPageRanges = GetPrintPageRanges(app, targetWs);
            if (targetPageRanges.Count == 0)
            {
                sheetResult.Status = (templateWs == null) ? SheetStatus.NewSheet : SheetStatus.TemplateOnly;
                sheetResult.TotalPrintPages = 1;
                sheetResult.TemplatePagesCount = (templateWs == null) ? 0 : 1;
                sheetResult.WorkPagesCount = 0;
                sheetResult.Pages.Add(new PageDetailItem
                {
                    PageNumber = 1,
                    RangeAddress = "A1",
                    StartRow = 1,
                    EndRow = 1,
                    StartCol = 1,
                    EndCol = 1,
                    Status = PageStatus.BlankPage,
                    Description = "Trang trống"
                });
                return sheetResult;
            }

            sheetResult.TotalPrintPages = targetPageRanges.Count;

            // 2. Nếu không có template tương ứng -> Sheet mới hoàn toàn
            if (templateWs == null)
            {
                sheetResult.Status = SheetStatus.NewSheet;
                int pageNo = 1;
                foreach (var pr in targetPageRanges)
                {
                    int nonEmptyCells = CountNonEmptyCells(targetWs, pr);
                    bool isBlank = nonEmptyCells == 0;
                    if (isBlank && options.IgnoreBlankPages)
                    {
                        sheetResult.BlankPagesCount++;
                        sheetResult.Pages.Add(new PageDetailItem
                        {
                            PageNumber = pageNo++,
                            RangeAddress = pr.Address,
                            StartRow = pr.StartRow,
                            EndRow = pr.EndRow,
                            StartCol = pr.StartCol,
                            EndCol = pr.EndCol,
                            Status = PageStatus.BlankPage,
                            Description = "Trang trống"
                        });
                    }
                    else
                    {
                        sheetResult.WorkPagesCount++;
                        sheetResult.Pages.Add(new PageDetailItem
                        {
                            PageNumber = pageNo++,
                            RangeAddress = pr.Address,
                            StartRow = pr.StartRow,
                            EndRow = pr.EndRow,
                            StartCol = pr.StartCol,
                            EndCol = pr.EndCol,
                            Status = PageStatus.NewPage,
                            TotalNonEmptyCells = nonEmptyCells,
                            Description = $"Sheet mới ({nonEmptyCells} ô có dữ liệu)"
                        });
                    }
                }
                return sheetResult;
            }

            // 3. Có template đối chiếu -> So sánh từng trang
            int pageIndex = 1;
            int totalModifiedPages = 0;

            foreach (var pr in targetPageRanges)
            {
                var pageDetail = ComparePageRange(targetWs, templateWs, pr, pageIndex, options);
                sheetResult.Pages.Add(pageDetail);
                pageIndex++;

                sheetResult.TotalChangedCells += pageDetail.ChangedCellsCount;
                sheetResult.TotalAddedShapes += pageDetail.AddedShapesCount;

                if (pageDetail.Status == PageStatus.BlankPage)
                {
                    sheetResult.BlankPagesCount++;
                }
                else if (pageDetail.IsWorkPage)
                {
                    sheetResult.WorkPagesCount++;
                    totalModifiedPages++;
                }
                else
                {
                    sheetResult.TemplatePagesCount++;
                }
            }

            if (totalModifiedPages > 0)
            {
                sheetResult.Status = SheetStatus.ModifiedSheet;
            }
            else
            {
                sheetResult.Status = SheetStatus.TemplateOnly;
            }

            return sheetResult;
        }

        /// <summary>
        /// So sánh nội dung một trang in giữa Target Sheet và Template Sheet.
        /// </summary>
        private static PageDetailItem ComparePageRange(
            Worksheet targetWs,
            Worksheet templateWs,
            PageRect pr,
            int pageNumber,
            PageCounterOptions options)
        {
            var item = new PageDetailItem
            {
                PageNumber = pageNumber,
                RangeAddress = pr.Address,
                StartRow = pr.StartRow,
                EndRow = pr.EndRow,
                StartCol = pr.StartCol,
                EndCol = pr.EndCol
            };

            Range? targetRange = null;
            Range? tplRange = null;

            try
            {
                int rowCount = pr.EndRow - pr.StartRow + 1;
                int colCount = pr.EndCol - pr.StartCol + 1;

                targetRange = targetWs.Range[targetWs.Cells[pr.StartRow, pr.StartCol], targetWs.Cells[pr.EndRow, pr.EndCol]];
                int targetNonEmpty = 0;
                int changedCount = 0;

                object?[,] targetVals = Extract2DArray(targetRange.Value2, rowCount, colCount);
                object?[,] targetFormulas = Extract2DArray(targetRange.Formula, rowCount, colCount);

                // Lấy range tương ứng trên template
                int tplMaxRow = 0;
                int tplMaxCol = 0;
                try
                {
                    Range tplUsed = templateWs.UsedRange;
                    if (tplUsed != null)
                    {
                        tplMaxRow = tplUsed.Row + tplUsed.Rows.Count - 1;
                        tplMaxCol = tplUsed.Column + tplUsed.Columns.Count - 1;
                    }
                }
                catch { }

                object?[,]? tplVals = null;
                object?[,]? tplFormulas = null;
                int tplRowCount = 0;
                int tplColCount = 0;

                if (pr.StartRow <= tplMaxRow && pr.StartCol <= tplMaxCol)
                {
                    try
                    {
                        int endR = Math.Min(pr.EndRow, tplMaxRow);
                        int endC = Math.Min(pr.EndCol, tplMaxCol);
                        tplRowCount = endR - pr.StartRow + 1;
                        tplColCount = endC - pr.StartCol + 1;
                        tplRange = templateWs.Range[templateWs.Cells[pr.StartRow, pr.StartCol], templateWs.Cells[endR, endC]];
                        tplVals = Extract2DArray(tplRange.Value2, tplRowCount, tplColCount);
                        tplFormulas = Extract2DArray(tplRange.Formula, tplRowCount, tplColCount);
                    }
                    catch { }
                }

                for (int r = 1; r <= rowCount; r++)
                {
                    for (int c = 1; c <= colCount; c++)
                    {
                        object? tVal = targetVals[r, c];
                        object? tForm = targetFormulas[r, c];

                        string tStr = tVal?.ToString()?.Trim() ?? string.Empty;
                        string tFormStr = tForm?.ToString()?.Trim() ?? string.Empty;

                        if (!string.IsNullOrEmpty(tStr) || !string.IsNullOrEmpty(tFormStr))
                        {
                            targetNonEmpty++;
                        }

                        object? tplVal = (tplVals != null && r <= tplRowCount && c <= tplColCount) ? tplVals[r, c] : null;
                        object? tplForm = (tplFormulas != null && r <= tplRowCount && c <= tplColCount) ? tplFormulas[r, c] : null;

                        string tplStr = tplVal?.ToString()?.Trim() ?? string.Empty;
                        string tplFormStr = tplForm?.ToString()?.Trim() ?? string.Empty;

                        bool isDiff = !string.Equals(tStr, tplStr, StringComparison.Ordinal) ||
                                      !string.Equals(tFormStr, tplFormStr, StringComparison.Ordinal);

                        if (isDiff)
                        {
                            changedCount++;
                        }
                    }
                }

                item.TotalNonEmptyCells = targetNonEmpty;
                item.ChangedCellsCount = changedCount;

                // Kiểm tra Shapes / Pictures mới
                if (options.CountShapesAndPictures)
                {
                    int targetShapes = CountShapesInRange(targetWs, pr);
                    int tplShapes = CountShapesInRange(templateWs, pr);
                    item.AddedShapesCount = Math.Max(0, targetShapes - tplShapes);
                }

                // Đánh giá trạng thái trang
                if (targetNonEmpty == 0 && options.IgnoreBlankPages)
                {
                    item.Status = PageStatus.BlankPage;
                    item.Description = "Trang trống";
                }
                else if (changedCount >= options.MinChangedCellsThreshold || item.AddedShapesCount > 0)
                {
                    item.Status = PageStatus.WorkPage;
                    List<string> details = new();
                    if (changedCount > 0) details.Add($"{changedCount} ô sửa đổi/thêm mới");
                    if (item.AddedShapesCount > 0) details.Add($"{item.AddedShapesCount} hình vẽ/ảnh mới");
                    item.Description = string.Join(", ", details);
                }
                else
                {
                    item.Status = PageStatus.TemplatePage;
                    item.Description = "Nguyên bản Template";
                }
            }
            catch (Exception ex)
            {
                item.Status = PageStatus.WorkPage;
                item.Description = $"Lỗi đọc: {ex.Message}";
            }
            finally
            {
                if (targetRange != null) Marshal.ReleaseComObject(targetRange);
                if (tplRange != null) Marshal.ReleaseComObject(tplRange);
            }

            return item;
        }

        #region Helper: Page Rect & Grid Calculation

        public class PageRect
        {
            public int StartRow { get; set; }
            public int EndRow { get; set; }
            public int StartCol { get; set; }
            public int EndCol { get; set; }
            public string Address => $"R{StartRow}C{StartCol}:R{EndRow}C{EndCol}";
        }

        /// <summary>
        /// Tính toán danh sách các hình chữ nhật (Range) tương ứng với từng trang in.
        /// </summary>
        private static List<PageRect> GetPrintPageRanges(ExcelApp app, Worksheet ws)
        {
            var pages = new List<PageRect>();

            Range? usedRange = null;
            try
            {
                usedRange = ws.UsedRange;
                if (usedRange == null || usedRange.Rows.Count == 0 || usedRange.Columns.Count == 0)
                    return pages;

                int startRow = usedRange.Row;
                int endRow = startRow + usedRange.Rows.Count - 1;
                int startCol = usedRange.Column;
                int endCol = startCol + usedRange.Columns.Count - 1;

                // Thu thập các mốc ngắt dòng (HPageBreaks)
                var rowBreaks = new List<int> { startRow };
                try
                {
                    foreach (HPageBreak hpb in ws.HPageBreaks)
                    {
                        int r = hpb.Location.Row;
                        if (r > startRow && r <= endRow && !rowBreaks.Contains(r))
                        {
                            rowBreaks.Add(r);
                        }
                    }
                }
                catch { }
                rowBreaks.Sort();

                // Thu thập các mốc ngắt cột (VPageBreaks)
                var colBreaks = new List<int> { startCol };
                try
                {
                    foreach (VPageBreak vpb in ws.VPageBreaks)
                    {
                        int c = vpb.Location.Column;
                        if (c > startCol && c <= endCol && !colBreaks.Contains(c))
                        {
                            colBreaks.Add(c);
                        }
                    }
                }
                catch { }
                colBreaks.Sort();

                // Thứ tự in: DownThenOver hay OverThenDown
                XlOrder printOrder = XlOrder.xlDownThenOver;
                try
                {
                    printOrder = (XlOrder)ws.PageSetup.Order;
                }
                catch { }

                if (printOrder == XlOrder.xlDownThenOver)
                {
                    for (int c = 0; c < colBreaks.Count; c++)
                    {
                        int sc = colBreaks[c];
                        int ec = (c + 1 < colBreaks.Count) ? colBreaks[c + 1] - 1 : endCol;

                        for (int r = 0; r < rowBreaks.Count; r++)
                        {
                            int sr = rowBreaks[r];
                            int er = (r + 1 < rowBreaks.Count) ? rowBreaks[r + 1] - 1 : endRow;

                            pages.Add(new PageRect { StartRow = sr, EndRow = er, StartCol = sc, EndCol = ec });
                        }
                    }
                }
                else
                {
                    for (int r = 0; r < rowBreaks.Count; r++)
                    {
                        int sr = rowBreaks[r];
                        int er = (r + 1 < rowBreaks.Count) ? rowBreaks[r + 1] - 1 : endRow;

                        for (int c = 0; c < colBreaks.Count; c++)
                        {
                            int sc = colBreaks[c];
                            int ec = (c + 1 < colBreaks.Count) ? colBreaks[c + 1] - 1 : endCol;

                            pages.Add(new PageRect { StartRow = sr, EndRow = er, StartCol = sc, EndCol = ec });
                        }
                    }
                }
            }
            catch { }
            finally
            {
                if (usedRange != null) Marshal.ReleaseComObject(usedRange);
            }

            return pages;
        }

        private static int CountNonEmptyCells(Worksheet ws, PageRect pr)
        {
            Range? range = null;
            try
            {
                int rowCount = pr.EndRow - pr.StartRow + 1;
                int colCount = pr.EndCol - pr.StartCol + 1;
                range = ws.Range[ws.Cells[pr.StartRow, pr.StartCol], ws.Cells[pr.EndRow, pr.EndCol]];
                object?[,] arr = Extract2DArray(range.Value2, rowCount, colCount);
                int count = 0;
                for (int r = 1; r <= rowCount; r++)
                {
                    for (int c = 1; c <= colCount; c++)
                    {
                        if (arr[r, c] != null && !string.IsNullOrWhiteSpace(arr[r, c]?.ToString()))
                        {
                            count++;
                        }
                    }
                }
                return count;
            }
            catch
            {
                return 0;
            }
            finally
            {
                if (range != null) Marshal.ReleaseComObject(range);
            }
        }

        /// <summary>
        /// Đếm các đối tượng hình ảnh/sơ đồ thiết kế có ý nghĩa thực tế (loại trừ drop-down, button, comment, icon nhỏ).
        /// </summary>
        public static int CountMeaningfulDesignShapes(Worksheet ws)
        {
            if (ws == null) return 0;
            int count = 0;
            try
            {
                foreach (Shape shape in ws.Shapes)
                {
                    try
                    {
                        // 1. Phải đang hiển thị (Visible)
                        if (shape.Visible != Microsoft.Office.Core.MsoTriState.msoTrue)
                            continue;

                        // 2. Lọc loại Shape: Loại bỏ form control, comment, line viền nhỏ, drop-down arrows
                        var type = shape.Type;
                        if (type == Microsoft.Office.Core.MsoShapeType.msoComment ||
                            type == Microsoft.Office.Core.MsoShapeType.msoFormControl ||
                            type == Microsoft.Office.Core.MsoShapeType.msoOLEControlObject ||
                            type == Microsoft.Office.Core.MsoShapeType.msoLine ||
                            type == Microsoft.Office.Core.MsoShapeType.msoFreeform)
                        {
                            continue;
                        }

                        // 3. Lọc theo kích thước: Phải đủ lớn để là một sơ đồ thiết kế / ảnh chụp màn hình / chart
                        float width = shape.Width;
                        float height = shape.Height;

                        // Bỏ qua các icon, bullet points, mũi tên nhỏ, viền khung (< 60x45 pt)
                        if (width < 60 || height < 45)
                            continue;

                        // Diện tích tối thiểu: >= 3,600 pt^2
                        if (width * height < 3600)
                            continue;

                        count++;
                    }
                    catch { }
                }
            }
            catch { }
            return count;
        }

        private static int CountShapesInRange(Worksheet ws, PageRect pr)
        {
            return CountMeaningfulShapesInRange(ws, pr);
        }

        public static int CountMeaningfulShapesInRange(Worksheet ws, PageRect pr)
        {
            int count = 0;
            try
            {
                foreach (Shape shape in ws.Shapes)
                {
                    try
                    {
                        if (shape.Visible != Microsoft.Office.Core.MsoTriState.msoTrue)
                            continue;

                        var type = shape.Type;
                        if (type == Microsoft.Office.Core.MsoShapeType.msoComment ||
                            type == Microsoft.Office.Core.MsoShapeType.msoFormControl ||
                            type == Microsoft.Office.Core.MsoShapeType.msoOLEControlObject ||
                            type == Microsoft.Office.Core.MsoShapeType.msoLine ||
                            type == Microsoft.Office.Core.MsoShapeType.msoFreeform)
                        {
                            continue;
                        }

                        float width = shape.Width;
                        float height = shape.Height;
                        if (width < 60 || height < 45 || width * height < 3600)
                            continue;

                        int topRow = shape.TopLeftCell.Row;
                        int leftCol = shape.TopLeftCell.Column;
                        if (topRow >= pr.StartRow && topRow <= pr.EndRow &&
                            leftCol >= pr.StartCol && leftCol <= pr.EndCol)
                        {
                            count++;
                        }
                    }
                    catch { }
                }
            }
            catch { }
            return count;
        }

        /// <summary>
        /// Kiểm tra xem màu nền của ô có khớp với màu chỉ định hay không.
        /// </summary>
        public static bool IsColorMatch(object interiorColor, object interiorColorIndex, Color targetColor, bool matchAnyHighlight)
        {
            if (interiorColor == null && interiorColorIndex == null) return false;

            int colorIndex = 0;
            if (interiorColorIndex != null && int.TryParse(interiorColorIndex.ToString(), out int cIdx))
            {
                colorIndex = cIdx;
            }

            // xlNone = -4142 (Không tô màu)
            if (colorIndex == -4142) return false;

            long oleColor = 0;
            if (interiorColor != null)
            {
                if (interiorColor is double d) oleColor = (long)d;
                else if (interiorColor is int i) oleColor = i;
                else if (interiorColor is long l) oleColor = l;
                else long.TryParse(interiorColor.ToString(), out oleColor);
            }

            // Không màu hoặc màu trắng: 16777215 (0xFFFFFF), hoặc ColorIndex == 2 khi ole == 0
            if (oleColor == 16777215 || (colorIndex == 2 && oleColor == 0) || (oleColor == 0 && colorIndex <= 0))
            {
                return false;
            }

            // Nếu người dùng chọn bất kỳ màu nào: Bất kỳ màu nào khác trắng và không màu đều khớp!
            if (matchAnyHighlight)
            {
                return oleColor > 0 || (colorIndex > 0 && colorIndex != 2);
            }

            // So khớp màu cụ thể
            int r = (int)(oleColor & 0xFF);
            int g = (int)((oleColor >> 8) & 0xFF);
            int b = (int)((oleColor >> 16) & 0xFF);

            // Bỏ qua màu trắng / xám cực nhạt gần trắng (RGB >= 250)
            if (r >= 250 && g >= 250 && b >= 250) return false;

            int dr = Math.Abs(r - targetColor.R);
            int dg = Math.Abs(g - targetColor.G);
            int db = Math.Abs(b - targetColor.B);

            // Khoảng cách RGB nhỏ
            if (dr + dg + db <= 120) return true;

            // Kiểm tra theo họ màu (Yellow, Green, Cyan, Orange, Pink)
            return IsInSameColorFamily(r, g, b, targetColor);
        }

        private static bool IsInSameColorFamily(int r, int g, int b, Color targetColor)
        {
            // Họ màu Vàng / Hổ phách (Yellow / Amber)
            if (targetColor.R > 200 && targetColor.G > 180 && targetColor.B < 180)
            {
                return r > 160 && g > 140 && b < 190 && (r >= b || g >= b);
            }
            // Họ màu Xanh ngọc / Xanh dương (Cyan / Light Blue)
            if (targetColor.B > 180 && targetColor.G > 140)
            {
                return b > 150 && (b >= r || g >= r);
            }
            // Họ màu Xanh lá (Green)
            if (targetColor.G > 160 && targetColor.R < 220)
            {
                return g > 140 && g >= r && g >= b;
            }
            // Họ màu Cam / Đào (Orange / Peach)
            if (targetColor.R > 200 && targetColor.G > 100 && targetColor.B < 180)
            {
                return r > 170 && g > 90 && b < 180 && r > b;
            }
            // Họ màu Hồng / Tím (Pink / Purple / Violet / Magenta)
            if (targetColor.R > 160 && targetColor.B > 160)
            {
                return (r > 130 && b > 130) || (b > 130 && r >= g) || (r > 130 && b >= g);
            }

            return false;
        }

        private static object?[,] Extract2DArray(object? rawVal, int rowCount = 1, int colCount = 1)
        {
            int rows = Math.Max(1, rowCount);
            int cols = Math.Max(1, colCount);
            var result = new object?[rows + 1, cols + 1];

            if (rawVal is object[,] arr)
            {
                int lowerR = arr.GetLowerBound(0);
                int upperR = arr.GetUpperBound(0);
                int lowerC = arr.GetLowerBound(1);
                int upperC = arr.GetUpperBound(1);

                for (int r = lowerR; r <= upperR; r++)
                {
                    int targetR = r - lowerR + 1;
                    if (targetR > rows) break;
                    for (int c = lowerC; c <= upperC; c++)
                    {
                        int targetC = c - lowerC + 1;
                        if (targetC > cols) break;
                        result[targetR, targetC] = arr[r, c];
                    }
                }
            }
            else if (rawVal != null)
            {
                result[1, 1] = rawVal;
            }

            return result;
        }

        public static bool IsCoverOrHistorySheet(string sheetName)
        {
            if (string.IsNullOrWhiteSpace(sheetName)) return false;
            string clean = sheetName.Trim().ToLowerInvariant();
            return CoverAndHistoryPatterns.Any(p => clean.Contains(p));
        }

        public static Workbook? FindOrOpenWorkbook(ExcelApp app, string nameOrPath, out bool openedHere)
        {
            openedHere = false;
            if (string.IsNullOrWhiteSpace(nameOrPath)) return null;

            // 1. Kiểm tra trong danh sách đang mở
            foreach (Workbook wb in app.Workbooks)
            {
                if (string.Equals(wb.Name, nameOrPath, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(wb.FullName, nameOrPath, StringComparison.OrdinalIgnoreCase))
                {
                    return wb;
                }
            }

            // 2. Mở file nếu tồn tại đường dẫn
            if (File.Exists(nameOrPath))
            {
                openedHere = true;
                return app.Workbooks.Open(nameOrPath, ReadOnly: true);
            }

            return null;
        }

        #endregion

        #region Report Export to Excel

        /// <summary>
        /// Xuất kết quả thống kê ra một Sheet mới trong Workbook đang chọn.
        /// </summary>
        public static bool ExportReportToExcel(ExcelApp app, Workbook targetWb, WorkbookPageCounterResult result)
        {
            if (app == null || targetWb == null || result == null) return false;

            bool prevScreenUpdating = app.ScreenUpdating;
            bool prevDisplayAlerts = app.DisplayAlerts;

            try
            {
                app.ScreenUpdating = false;
                app.DisplayAlerts = false;

                string sheetName = "ThongKe_TrangThietKe";
                int counter = 1;
                while (SheetExists(targetWb, sheetName))
                {
                    sheetName = $"ThongKe_TrangThietKe_{counter++}";
                }

                Worksheet reportWs = targetWb.Worksheets.Add(After: targetWb.Worksheets[targetWb.Worksheets.Count]);
                reportWs.Name = sheetName;

                // --- 1. TIÊU ĐỀ BÁO CÁO ---
                reportWs.Cells[1, 1].Value2 = "BÁO CÁO THỐNG KÊ SỐ TRANG THIẾT KẾ THỰC TẾ (DESIGN PAGE REPORT)";
                reportWs.Cells[1, 1].Font.Bold = true;
                reportWs.Cells[1, 1].Font.Size = 14;
                reportWs.Cells[1, 1].Font.Color = ColorTranslator.ToOle(Color.FromArgb(37, 99, 235));

                reportWs.Cells[2, 1].Value2 = $"File thiết kế: {result.TargetWorkbookName} | Template đối chiếu: {result.TemplateWorkbookName}";
                reportWs.Cells[2, 1].Font.Size = 10;
                reportWs.Cells[2, 1].Font.Color = ColorTranslator.ToOle(Color.FromArgb(100, 116, 139));

                reportWs.Cells[3, 1].Value2 = $"Thời gian thống kê: {result.AnalyzedAt:yyyy-MM-dd HH:mm:ss}";
                reportWs.Cells[3, 1].Font.Size = 10;
                reportWs.Cells[3, 1].Font.Color = ColorTranslator.ToOle(Color.FromArgb(100, 116, 139));

                // --- 2. THẺ CHỈ SỐ KPI TỔNG HỢP ---
                int kpiRow = 5;
                reportWs.Cells[kpiRow, 1].Value2 = "TỔNG KÝ TỰ MỚI / SỬA";
                reportWs.Cells[kpiRow + 1, 1].Value2 = $"{result.TotalChangedCharacters:N0} ký tự";
                FormatKpiCard(reportWs, kpiRow, 1, Color.FromArgb(239, 246, 255), Color.FromArgb(37, 99, 235));

                reportWs.Cells[kpiRow, 3].Value2 = "SƠ ĐỒ / ẢNH MỚI";
                reportWs.Cells[kpiRow + 1, 3].Value2 = $"{result.TotalAddedShapes} hình";
                FormatKpiCard(reportWs, kpiRow, 3, Color.FromArgb(241, 245, 249), Color.FromArgb(100, 116, 139));

                reportWs.Cells[kpiRow, 5].Value2 = "SỐ TRANG THIẾT KẾ QUY ĐỔI";
                reportWs.Cells[kpiRow + 1, 5].Value2 = $"{result.TotalWorkPages:F1} trang";
                FormatKpiCard(reportWs, kpiRow, 5, Color.FromArgb(240, 253, 244), Color.FromArgb(22, 163, 74));

                reportWs.Cells[kpiRow, 7].Value2 = "TỶ LỆ LÀM MỚI / SỬA ĐỔI";
                reportWs.Cells[kpiRow + 1, 7].Value2 = $"{result.OverallWorkPercent}%";
                FormatKpiCard(reportWs, kpiRow, 7, Color.FromArgb(254, 242, 242), Color.FromArgb(220, 38, 38));

                // --- 3. BẢNG CHI TIẾT TỪNG SHEET ---
                int tableHeaderRow = 8;
                string[] headers = new[]
                {
                    "STT", "Tên Sheet", "Trạng thái", "Tổng trang in", "Trang Template", "Trang Thiết kế", "% Thiết kế", "Ký tự mới/sửa", "Hình vẽ mới", "Số ô sửa đổi", "Chi tiết các trang"
                };

                for (int col = 0; col < headers.Length; col++)
                {
                    var cell = reportWs.Cells[tableHeaderRow, col + 1];
                    cell.Value2 = headers[col];
                    cell.Font.Bold = true;
                    cell.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(37, 99, 235));
                    cell.Font.Color = ColorTranslator.ToOle(Color.White);
                    cell.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    cell.VerticalAlignment = XlVAlign.xlVAlignCenter;
                }

                int curRow = tableHeaderRow + 1;
                int stt = 1;

                foreach (var s in result.SheetResults)
                {
                    reportWs.Cells[curRow, 1].Value2 = stt++;
                    reportWs.Cells[curRow, 2].Value2 = s.SheetName;
                    reportWs.Cells[curRow, 3].Value2 = s.SummaryText;
                    reportWs.Cells[curRow, 4].Value2 = s.TotalPrintPages;
                    reportWs.Cells[curRow, 5].Value2 = s.TemplatePagesCount;
                    reportWs.Cells[curRow, 6].Value2 = s.WorkPagesCount;
                    reportWs.Cells[curRow, 7].Value2 = $"{s.WorkPercent}%";
                    reportWs.Cells[curRow, 8].Value2 = s.TotalChangedCharacters;
                    reportWs.Cells[curRow, 9].Value2 = s.TotalAddedShapes;
                    reportWs.Cells[curRow, 10].Value2 = s.TotalChangedCells;

                    // Chi tiết từng trang
                    var pageSummaries = s.Pages
                        .Where(p => p.IsWorkPage)
                        .Select(p => $"Trang {p.PageNumber} ({p.Description})");
                    reportWs.Cells[curRow, 11].Value2 = string.Join("; ", pageSummaries);

                    if (s.WorkPagesCount > 0)
                    {
                        reportWs.Cells[curRow, 6].Font.Bold = true;
                        reportWs.Cells[curRow, 6].Font.Color = ColorTranslator.ToOle(Color.FromArgb(22, 163, 74));
                    }

                    // Kẻ viền dòng
                    Range rowRange = reportWs.Range[reportWs.Cells[curRow, 1], reportWs.Cells[curRow, headers.Length]];
                    rowRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                    rowRange.Borders.Color = ColorTranslator.ToOle(Color.FromArgb(226, 232, 240));

                    curRow++;
                }

                // Dòng tổng cộng
                reportWs.Cells[curRow, 1].Value2 = "TỔNG CỘNG";
                reportWs.Cells[curRow, 1].Font.Bold = true;
                reportWs.Cells[curRow, 4].Value2 = result.TotalTargetPrintPages;
                reportWs.Cells[curRow, 4].Font.Bold = true;
                reportWs.Cells[curRow, 5].Value2 = result.TotalTemplatePages;
                reportWs.Cells[curRow, 5].Font.Bold = true;
                reportWs.Cells[curRow, 6].Value2 = result.TotalWorkPages;
                reportWs.Cells[curRow, 6].Font.Bold = true;
                reportWs.Cells[curRow, 6].Font.Color = ColorTranslator.ToOle(Color.FromArgb(22, 163, 74));
                reportWs.Cells[curRow, 7].Value2 = $"{result.OverallWorkPercent}%";
                reportWs.Cells[curRow, 7].Font.Bold = true;
                reportWs.Cells[curRow, 8].Value2 = result.TotalChangedCharacters;
                reportWs.Cells[curRow, 8].Font.Bold = true;
                reportWs.Cells[curRow, 9].Value2 = result.TotalAddedShapes;
                reportWs.Cells[curRow, 9].Font.Bold = true;
                reportWs.Cells[curRow, 10].Value2 = result.TotalChangedCells;
                reportWs.Cells[curRow, 10].Font.Bold = true;

                Range totalRange = reportWs.Range[reportWs.Cells[curRow, 1], reportWs.Cells[curRow, headers.Length]];
                totalRange.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(241, 245, 249));
                totalRange.Borders.LineStyle = XlLineStyle.xlContinuous;

                // Thông tin file Evidence nếu có
                if (!string.IsNullOrEmpty(result.HighlightedClonedWorkbookPath) && File.Exists(result.HighlightedClonedWorkbookPath))
                {
                    curRow += 2;
                    reportWs.Cells[curRow, 1].Value2 = "🎨 File bản sao đã tô màu đối chiếu (Evidence):";
                    reportWs.Cells[curRow, 1].Font.Bold = true;
                    reportWs.Cells[curRow, 1].Font.Color = ColorTranslator.ToOle(Color.FromArgb(124, 58, 237));
                    reportWs.Cells[curRow + 1, 1].Value2 = result.HighlightedClonedWorkbookPath;
                    reportWs.Cells[curRow + 1, 1].Font.Italic = true;
                }

                // --- 4. TỰ ĐỘNG VẼ BIỂU ĐỒ TRỰC QUAN (CHARTS) ---
                try
                {
                    // 4.1 Bảng dữ liệu phụ cho Biểu đồ tròn (Pie Chart) đặt tại cột M, N
                    int chartDataRow = 8;
                    reportWs.Cells[chartDataRow, 13].Value2 = "Phân loại trang";
                    reportWs.Cells[chartDataRow, 14].Value2 = "Số trang";
                    reportWs.Cells[chartDataRow + 1, 13].Value2 = "Trang Thiết kế mới";
                    reportWs.Cells[chartDataRow + 1, 14].Value2 = Math.Max(0.0, result.TotalWorkPages);
                    reportWs.Cells[chartDataRow + 2, 13].Value2 = "Trang Template gốc";
                    reportWs.Cells[chartDataRow + 2, 14].Value2 = Math.Max(0.0, (double)result.TotalTemplatePages);

                    Range pieDataRange = reportWs.Range[reportWs.Cells[chartDataRow, 13], reportWs.Cells[chartDataRow + 2, 14]];
                    
                    ChartObjects chartObjs = (ChartObjects)reportWs.ChartObjects();
                    ChartObject pieChartObj = chartObjs.Add(540, 20, 320, 200);
                    pieChartObj.Chart.ChartType = XlChartType.xlDoughnut;
                    pieChartObj.Chart.SetSourceData(pieDataRange);
                    pieChartObj.Chart.HasTitle = true;
                    pieChartObj.Chart.ChartTitle.Text = "Tỷ Lệ Thiết Kế vs Template";

                    // 4.2 Biểu đồ cột phân bổ khối lượng từng Sheet (nếu có nhiều hơn 1 sheet)
                    if (result.SheetResults.Count > 1)
                    {
                        int barDataStartRow = curRow + 4;
                        reportWs.Cells[barDataStartRow, 1].Value2 = "Tên Sheet";
                        reportWs.Cells[barDataStartRow, 2].Value2 = "Trang Thiết kế";
                        
                        int rIdx = barDataStartRow + 1;
                        foreach (var s in result.SheetResults.Take(20))
                        {
                            reportWs.Cells[rIdx, 1].Value2 = s.SheetName;
                            reportWs.Cells[rIdx, 2].Value2 = s.WorkPagesCount;
                            rIdx++;
                        }

                        Range barDataRange = reportWs.Range[reportWs.Cells[barDataStartRow, 1], reportWs.Cells[rIdx - 1, 2]];
                        ChartObject barChartObj = chartObjs.Add(20, curRow + 4, 500, 220);
                        barChartObj.Chart.ChartType = XlChartType.xlColumnClustered;
                        barChartObj.Chart.SetSourceData(barDataRange);
                        barChartObj.Chart.HasTitle = true;
                        barChartObj.Chart.ChartTitle.Text = "Số Trang Thiết Kế Từng Sheet";
                        try { barChartObj.Chart.Legend?.Delete(); } catch { }
                    }
                }
                catch (Exception exChart)
                {
                    System.Diagnostics.Debug.WriteLine($"Error creating report charts: {exChart.Message}");
                }

                // Tự căn chỉnh độ rộng cột
                reportWs.Columns.AutoFit();
                reportWs.Activate();
                return true;
            }
            catch
            {
                return false;
            }
            finally
            {
                app.ScreenUpdating = prevScreenUpdating;
                app.DisplayAlerts = prevDisplayAlerts;
            }
        }

        private static void FormatKpiCard(Worksheet ws, int row, int col, Color bgColor, Color textColor)
        {
            Range titleCell = ws.Cells[row, col];
            Range valCell = ws.Cells[row + 1, col];

            Range card = ws.Range[ws.Cells[row, col], ws.Cells[row + 1, col + 1]];
            card.Merge();
            card.Interior.Color = ColorTranslator.ToOle(bgColor);
            card.Borders.LineStyle = XlLineStyle.xlContinuous;
            card.Borders.Color = ColorTranslator.ToOle(Color.FromArgb(203, 213, 225));

            titleCell.Font.Size = 10;
            titleCell.Font.Bold = true;
            titleCell.Font.Color = ColorTranslator.ToOle(textColor);
            titleCell.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            titleCell.VerticalAlignment = XlVAlign.xlVAlignCenter;
        }

        private static bool SheetExists(Workbook wb, string sheetName)
        {
            foreach (Worksheet s in wb.Worksheets)
            {
                if (string.Equals(s.Name, sheetName, StringComparison.OrdinalIgnoreCase)) return true;
            }
            return false;
        }

        #endregion
    }
}
