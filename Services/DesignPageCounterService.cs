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
        public bool IgnoreCoverAndHistory { get; set; } = true;
        public bool IgnoreBlankPages { get; set; } = true;
        public bool CountShapesAndPictures { get; set; } = true;
        public int MinChangedCellsThreshold { get; set; } = 2; // Tối thiểu N ô sửa đổi mới tính là trang làm việc
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
        public int WorkPagesCount { get; set; }
        public int BlankPagesCount { get; set; }
        public int TotalChangedCells { get; set; }
        public int TotalAddedShapes { get; set; }
        public double WorkPercent => TotalPrintPages > 0 ? Math.Round((double)WorkPagesCount / TotalPrintPages * 100.0, 1) : 0;
        public List<PageDetailItem> Pages { get; set; } = new();

        public string SummaryText
        {
            get
            {
                if (Status == SheetStatus.NewSheet)
                    return $"Sheet mới (+{WorkPagesCount} trang)";
                if (Status == SheetStatus.TemplateOnly)
                    return $"Nguyên bản Template ({TemplatePagesCount} trang)";
                return $"Thiết kế {WorkPagesCount}/{TotalPrintPages} trang ({WorkPercent}%)";
            }
        }
    }

    public class WorkbookPageCounterResult
    {
        public string TargetWorkbookName { get; set; } = string.Empty;
        public string TargetWorkbookPath { get; set; } = string.Empty;
        public string TemplateWorkbookName { get; set; } = string.Empty;
        public string TemplateWorkbookPath { get; set; } = string.Empty;
        public DateTime AnalyzedAt { get; set; } = DateTime.Now;

        public int TotalTargetPrintPages { get; set; }
        public int TotalTemplatePages { get; set; }
        public int TotalWorkPages { get; set; }
        public int TotalBlankPages { get; set; }
        public int TotalNewSheetsCount { get; set; }
        public int TotalModifiedSheetsCount { get; set; }
        public int TotalUnchangedSheetsCount { get; set; }

        public double OverallWorkPercent => TotalTargetPrintPages > 0
            ? Math.Round((double)TotalWorkPages / TotalTargetPrintPages * 100.0, 1)
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

                // 2. Resolve Template Workbook (nếu có)
                if (!string.IsNullOrWhiteSpace(templateWbNameOrPath))
                {
                    progressCallback?.Invoke("Đang mở và tải template gốc...", 15);
                    templateWb = FindOrOpenWorkbook(app, templateWbNameOrPath, out templateOpenedHere);
                }

                // 3. Xây dựng danh sách Sheets của Template để so khớp nhanh
                var templateSheetsDict = new Dictionary<string, Worksheet>(StringComparer.OrdinalIgnoreCase);
                if (templateWb != null)
                {
                    foreach (Worksheet ws in templateWb.Worksheets)
                    {
                        templateSheetsDict[ws.Name] = ws;
                    }
                }

                int sheetCount = targetWb.Worksheets.Count;
                int currentSheetIndex = 0;

                // 4. Lặp qua từng Sheet trong Target Workbook
                foreach (Worksheet targetWs in targetWb.Worksheets)
                {
                    currentSheetIndex++;
                    int progress = 20 + (int)((double)currentSheetIndex / Math.Max(1, sheetCount) * 70);
                    progressCallback?.Invoke($"Đang phân tích trang in sheet: {targetWs.Name} ({currentSheetIndex}/{sheetCount})...", progress);

                    bool isCoverOrHistory = IsCoverOrHistorySheet(targetWs.Name);
                    if (options.IgnoreCoverAndHistory && isCoverOrHistory)
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

                    var sheetResult = AnalyzeWorksheetPages(app, targetWs, tplWs, options);
                    result.SheetResults.Add(sheetResult);

                    result.TotalTargetPrintPages += sheetResult.TotalPrintPages;
                    result.TotalTemplatePages += sheetResult.TemplatePagesCount;
                    result.TotalWorkPages += sheetResult.WorkPagesCount;
                    result.TotalBlankPages += sheetResult.BlankPagesCount;

                    if (sheetResult.Status == SheetStatus.NewSheet)
                        result.TotalNewSheetsCount++;
                    else if (sheetResult.Status == SheetStatus.ModifiedSheet)
                        result.TotalModifiedSheetsCount++;
                    else if (sheetResult.Status == SheetStatus.TemplateOnly)
                        result.TotalUnchangedSheetsCount++;
                }

                progressCallback?.Invoke("Hoàn tất phân tích!", 100);
                return result;
            }
            finally
            {
                app.ScreenUpdating = prevScreenUpdating;
                app.DisplayAlerts = prevDisplayAlerts;

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
                targetRange = targetWs.Range[targetWs.Cells[pr.StartRow, pr.StartCol], targetWs.Cells[pr.EndRow, pr.EndCol]];
                int targetNonEmpty = 0;
                int changedCount = 0;

                object[,] targetVals = Extract2DArray(targetRange.Value2);
                object[,] targetFormulas = Extract2DArray(targetRange.Formula);

                // Lấy range tương ứng trên template
                int tplMaxRow = 0;
                int tplMaxCol = 0;
                try
                {
                    tplMaxRow = templateWs.UsedRange.Row + templateWs.UsedRange.Rows.Count - 1;
                    tplMaxCol = templateWs.UsedRange.Column + templateWs.UsedRange.Columns.Count - 1;
                }
                catch { }

                object[,]? tplVals = null;
                object[,]? tplFormulas = null;

                if (pr.StartRow <= tplMaxRow && pr.StartCol <= tplMaxCol)
                {
                    try
                    {
                        int endR = Math.Min(pr.EndRow, tplMaxRow);
                        int endC = Math.Min(pr.EndCol, tplMaxCol);
                        tplRange = templateWs.Range[templateWs.Cells[pr.StartRow, pr.StartCol], templateWs.Cells[endR, endC]];
                        tplVals = Extract2DArray(tplRange.Value2);
                        tplFormulas = Extract2DArray(tplRange.Formula);
                    }
                    catch { }
                }

                int rowCount = pr.EndRow - pr.StartRow + 1;
                int colCount = pr.EndCol - pr.StartCol + 1;

                for (int r = 1; r <= rowCount; r++)
                {
                    for (int c = 1; c <= colCount; c++)
                    {
                        object? tVal = (r <= targetVals.GetLength(0) && c <= targetVals.GetLength(1)) ? targetVals[r, c] : null;
                        object? tForm = (r <= targetFormulas.GetLength(0) && c <= targetFormulas.GetLength(1)) ? targetFormulas[r, c] : null;

                        string tStr = tVal?.ToString()?.Trim() ?? string.Empty;
                        string tFormStr = tForm?.ToString()?.Trim() ?? string.Empty;

                        if (!string.IsNullOrEmpty(tStr) || !string.IsNullOrEmpty(tFormStr))
                        {
                            targetNonEmpty++;
                        }

                        object? tplVal = (tplVals != null && r <= tplVals.GetLength(0) && c <= tplVals.GetLength(1)) ? tplVals[r, c] : null;
                        object? tplForm = (tplFormulas != null && r <= tplFormulas.GetLength(0) && c <= tplFormulas.GetLength(1)) ? tplFormulas[r, c] : null;

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
                range = ws.Range[ws.Cells[pr.StartRow, pr.StartCol], ws.Cells[pr.EndRow, pr.EndCol]];
                object[,] arr = Extract2DArray(range.Value2);
                int count = 0;
                for (int r = 1; r <= arr.GetLength(0); r++)
                {
                    for (int c = 1; c <= arr.GetLength(1); c++)
                    {
                        if (arr[r, c] != null && !string.IsNullOrWhiteSpace(arr[r, c].ToString()))
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

        private static int CountShapesInRange(Worksheet ws, PageRect pr)
        {
            int count = 0;
            try
            {
                foreach (Shape shape in ws.Shapes)
                {
                    try
                    {
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

        private static object[,] Extract2DArray(object rawVal)
        {
            if (rawVal is object[,] arr)
            {
                return arr;
            }
            else if (rawVal != null)
            {
                object[,] single = new object[2, 2];
                single[1, 1] = rawVal;
                return single;
            }
            return new object[1, 1];
        }

        private static bool IsCoverOrHistorySheet(string sheetName)
        {
            if (string.IsNullOrWhiteSpace(sheetName)) return false;
            string clean = sheetName.Trim().ToLowerInvariant();
            return CoverAndHistoryPatterns.Any(p => clean.Contains(p));
        }

        private static Workbook? FindOrOpenWorkbook(ExcelApp app, string nameOrPath, out bool openedHere)
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
                reportWs.Cells[kpiRow, 1].Value2 = "TỔNG SỐ TRANG IN";
                reportWs.Cells[kpiRow + 1, 1].Value2 = result.TotalTargetPrintPages;
                FormatKpiCard(reportWs, kpiRow, 1, Color.FromArgb(239, 246, 255), Color.FromArgb(37, 99, 235));

                reportWs.Cells[kpiRow, 3].Value2 = "TRANG TEMPLATE GỐC (LOẠI TRỪ)";
                reportWs.Cells[kpiRow + 1, 3].Value2 = result.TotalTemplatePages;
                FormatKpiCard(reportWs, kpiRow, 3, Color.FromArgb(241, 245, 249), Color.FromArgb(100, 116, 139));

                reportWs.Cells[kpiRow, 5].Value2 = "SỐ TRANG THIẾT KẾ THỰC TẾ";
                reportWs.Cells[kpiRow + 1, 5].Value2 = result.TotalWorkPages;
                FormatKpiCard(reportWs, kpiRow, 5, Color.FromArgb(240, 253, 244), Color.FromArgb(22, 163, 74));

                reportWs.Cells[kpiRow, 7].Value2 = "TỶ LỆ LÀM MỚI / CHỈNH SỬA";
                reportWs.Cells[kpiRow + 1, 7].Value2 = $"{result.OverallWorkPercent}%";
                FormatKpiCard(reportWs, kpiRow, 7, Color.FromArgb(254, 242, 242), Color.FromArgb(220, 38, 38));

                // --- 3. BẢNG CHI TIẾT TỪNG SHEET ---
                int tableHeaderRow = 8;
                string[] headers = new[]
                {
                    "STT", "Tên Sheet", "Trạng thái", "Tổng số trang", "Trang Template", "Trang Thiết kế", "% Thiết kế", "Số ô sửa đổi", "Hình vẽ mới", "Chi tiết các trang"
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
                    reportWs.Cells[curRow, 8].Value2 = s.TotalChangedCells;
                    reportWs.Cells[curRow, 9].Value2 = s.TotalAddedShapes;

                    // Chi tiết từng trang
                    var pageSummaries = s.Pages
                        .Where(p => p.IsWorkPage)
                        .Select(p => $"Trang {p.PageNumber} ({p.Description})");
                    reportWs.Cells[curRow, 10].Value2 = string.Join("; ", pageSummaries);

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

                Range totalRange = reportWs.Range[reportWs.Cells[curRow, 1], reportWs.Cells[curRow, headers.Length]];
                totalRange.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(241, 245, 249));
                totalRange.Borders.LineStyle = XlLineStyle.xlContinuous;

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
