using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using ExcelSupport.Models;
using Microsoft.Office.Interop.Excel;

namespace ExcelSupport.Services
{
    public class ExternalLinksScanResult
    {
        public List<ExternalSourceItem> Sources { get; set; } = new List<ExternalSourceItem>();
        public List<BrokenFormulaCellItem> FormulaCells { get; set; } = new List<BrokenFormulaCellItem>();
        public List<ExternalNamedRangeItem> NamedRanges { get; set; } = new List<ExternalNamedRangeItem>();
        public int TotalBrokenLinksCount { get; set; }
    }

    public static class ExternalLinksService
    {
        private static readonly Regex ExternalRefRegex = new Regex(@"\[([^\]]+)\]", RegexOptions.Compiled);

        /// <summary>
        /// Quét toàn bộ Workbook để tìm các liên kết ngoài, ô công thức chứa link và Named Range hỏng
        /// </summary>
        public static ExternalLinksScanResult ScanWorkbook(Workbook wb)
        {
            var result = new ExternalLinksScanResult();
            if (wb == null) return result;

            var sourceCountMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            // 1. Quét LinkSources từ Workbook Object Model
            try
            {
                object rawLinks = wb.LinkSources(XlLink.xlExcelLinks);
                if (rawLinks is Array linksArray)
                {
                    foreach (object item in linksArray)
                    {
                        if (item is string linkPath && !string.IsNullOrWhiteSpace(linkPath))
                        {
                            bool exists = CheckFileExists(linkPath);
                            string fileName = Path.GetFileName(linkPath);

                            var srcItem = new ExternalSourceItem
                            {
                                SourcePath = linkPath,
                                FileName = string.IsNullOrEmpty(fileName) ? linkPath : fileName,
                                Exists = exists,
                                StatusDisplay = exists ? "⚠️ File tồn tại trên máy" : "❌ File không tồn tại (Broken)",
                                FormulaCount = 0
                            };
                            result.Sources.Add(srcItem);
                            sourceCountMap[linkPath] = 0;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"LinkSources scan error: {ex.Message}");
            }

            // 2. Quét chi tiết các Sheet và các ô chứa công thức (Sử dụng 2D Array cực nhanh)
            Sheets? sheets = null;
            try
            {
                sheets = wb.Worksheets;
                int sheetCount = sheets.Count;

                for (int i = 1; i <= sheetCount; i++)
                {
                    _Worksheet? ws = null;
                    Range? usedRange = null;

                    try
                    {
                        ws = sheets[i] as _Worksheet;
                        if (ws == null) continue;

                        string sheetName = ws.Name;
                        usedRange = ws.UsedRange;
                        if (usedRange == null) continue;

                        int rowCount = usedRange.Rows.Count;
                        int colCount = usedRange.Columns.Count;
                        int startRow = usedRange.Row;
                        int startCol = usedRange.Column;

                        if (rowCount == 1 && colCount == 1)
                        {
                            // 1 ô duy nhất
                            object? formObj = usedRange.Formula;
                            string? formStr = formObj?.ToString();
                            if (!string.IsNullOrEmpty(formStr) && formStr!.StartsWith("=") && formStr.Contains("["))
                            {
                                ProcessFormulaCell(ws, usedRange, formStr, sheetName, startRow, startCol, result, sourceCountMap);
                            }
                        }
                        else
                        {
                            // Lấy mảng 2D công thức và giá trị trong 1 COM call
                            object[,] formulas = (object[,])usedRange.Formula;
                            object[,]? values = null;
                            try { values = (object[,])usedRange.Value2; } catch { }

                            for (int r = 1; r <= rowCount; r++)
                            {
                                for (int c = 1; c <= colCount; c++)
                                {
                                    object? fObj = formulas[r, c];
                                    if (fObj is string fStr && fStr.StartsWith("=") && fStr.Contains("["))
                                    {
                                        int curRow = startRow + r - 1;
                                        int curCol = startCol + c - 1;
                                        string curVal = string.Empty;

                                        if (values != null)
                                        {
                                            object? vObj = values[r, c];
                                            curVal = vObj?.ToString() ?? string.Empty;
                                        }

                                        ExtractAndRecordCell(sheetName, curRow, curCol, fStr, curVal, result, sourceCountMap);
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Sheet {i} scan error: {ex.Message}");
                    }
                    finally
                    {
                        if (usedRange != null) Marshal.ReleaseComObject(usedRange);
                        if (ws != null) Marshal.ReleaseComObject(ws);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Worksheets scan error: {ex.Message}");
            }
            finally
            {
                if (sheets != null) Marshal.ReleaseComObject(sheets);
            }

            // 3. Quét Defined Names (Tên vùng)
            Names? names = null;
            try
            {
                names = wb.Names;
                if (names != null)
                {
                    int nameCount = names.Count;
                    for (int i = 1; i <= nameCount; i++)
                    {
                        Name? nameObj = null;
                        try
                        {
                            nameObj = names.Item(i);
                            if (nameObj != null)
                            {
                                string nameText = nameObj.Name;
                                string refersTo = nameObj.RefersTo?.ToString() ?? string.Empty;

                                if (refersTo.Contains("[") || refersTo.Contains("#REF!"))
                                {
                                    bool isBroken = refersTo.Contains("#REF!");
                                    if (!isBroken)
                                    {
                                        var match = ExternalRefRegex.Match(refersTo);
                                        if (match.Success)
                                        {
                                            string extPath = match.Groups[1].Value;
                                            isBroken = !CheckFileExists(extPath);
                                        }
                                    }

                                    result.NamedRanges.Add(new ExternalNamedRangeItem
                                    {
                                        Name = nameText,
                                        Scope = "Workbook",
                                        RefersTo = refersTo,
                                        IsBroken = isBroken,
                                        IsSelected = true
                                    });
                                }
                            }
                        }
                        catch { }
                        finally
                        {
                            if (nameObj != null) Marshal.ReleaseComObject(nameObj);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Names scan error: {ex.Message}");
            }
            finally
            {
                if (names != null) Marshal.ReleaseComObject(names);
            }

            // 4. Cập nhật số lượng ô vào Sources
            foreach (var src in result.Sources)
            {
                if (sourceCountMap.TryGetValue(src.SourcePath, out int count))
                {
                    src.FormulaCount = count;
                }
                else
                {
                    // Đếm theo tên file
                    int matches = 0;
                    foreach (var cell in result.FormulaCells)
                    {
                        if (cell.Formula.IndexOf(src.FileName, StringComparison.OrdinalIgnoreCase) >= 0 ||
                            cell.Formula.IndexOf(src.SourcePath, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            matches++;
                        }
                    }
                    src.FormulaCount = matches;
                }
            }

            // Đếm tổng số link bị hỏng
            int brokenSources = 0;
            foreach (var s in result.Sources) if (!s.Exists) brokenSources++;
            int brokenCells = 0;
            foreach (var c in result.FormulaCells) if (c.IsBroken) brokenCells++;
            result.TotalBrokenLinksCount = brokenSources + brokenCells;

            return result;
        }

        private static void ProcessFormulaCell(_Worksheet ws, Range cell, string formula, string sheetName, int row, int col, ExternalLinksScanResult result, Dictionary<string, int> sourceCountMap)
        {
            string curVal = string.Empty;
            try { curVal = cell.Value2?.ToString() ?? string.Empty; } catch { }
            ExtractAndRecordCell(sheetName, row, col, formula, curVal, result, sourceCountMap);
        }

        private static void ExtractAndRecordCell(string sheetName, int row, int col, string formula, string curVal, ExternalLinksScanResult result, Dictionary<string, int> sourceCountMap)
        {
            string extSource = string.Empty;
            var match = ExternalRefRegex.Match(formula);
            if (match.Success)
            {
                extSource = match.Groups[1].Value;
            }

            bool isBroken = formula.Contains("#REF!") || curVal.Contains("#REF!");
            if (!string.IsNullOrEmpty(extSource))
            {
                bool exists = CheckFileExists(extSource);
                if (!exists) isBroken = true;

                if (sourceCountMap.ContainsKey(extSource))
                {
                    sourceCountMap[extSource]++;
                }
                else
                {
                    // Kiểm tra xem có nằm trong source list nào không
                    bool matched = false;
                    foreach (var key in sourceCountMap.Keys)
                    {
                        if (key.EndsWith(extSource, StringComparison.OrdinalIgnoreCase) ||
                            extSource.EndsWith(Path.GetFileName(key), StringComparison.OrdinalIgnoreCase))
                        {
                            sourceCountMap[key]++;
                            matched = true;
                            break;
                        }
                    }
                    if (!matched)
                    {
                        sourceCountMap[extSource] = 1;
                    }
                }
            }

            string cellAddr = GetCellAddress(row, col);

            result.FormulaCells.Add(new BrokenFormulaCellItem
            {
                SheetName = sheetName,
                CellAddress = cellAddr,
                Row = row,
                Column = col,
                Formula = formula,
                CurrentValue = curVal,
                ExternalSource = extSource,
                IsBroken = isBroken,
                IsSelected = true
            });
        }

        private static string GetCellAddress(int row, int col)
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

        private static bool CheckFileExists(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return false;
            try
            {
                // Nếu là đường dẫn đầy đủ
                if (Path.IsPathRooted(path) || path.StartsWith("\\\\"))
                {
                    return File.Exists(path);
                }
                return false;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Bẻ gãy một liên kết ngoài cụ thể (Break Link)
        /// </summary>
        public static bool BreakSpecificLink(Workbook wb, string linkPath)
        {
            if (wb == null || string.IsNullOrWhiteSpace(linkPath)) return false;
            try
            {
                wb.BreakLink(linkPath, XlLinkType.xlLinkTypeExcelLinks);
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"BreakSpecificLink error: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Bẻ gãy toàn bộ liên kết ngoài trong Workbook
        /// </summary>
        public static int BreakAllWorkbookLinks(Workbook wb)
        {
            if (wb == null) return 0;
            int count = 0;

            try
            {
                object rawLinks = wb.LinkSources(XlLink.xlExcelLinks);
                if (rawLinks is Array linksArray)
                {
                    foreach (object item in linksArray)
                    {
                        if (item is string linkPath && !string.IsNullOrWhiteSpace(linkPath))
                        {
                            try
                            {
                                wb.BreakLink(linkPath, XlLinkType.xlLinkTypeExcelLinks);
                                count++;
                            }
                            catch { }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"BreakAllWorkbookLinks error: {ex.Message}");
            }

            return count;
        }

        /// <summary>
        /// Chuyển các ô công thức đã chọn thành giá trị tĩnh (Freeze values)
        /// </summary>
        public static int ConvertCellsToValues(Workbook wb, IEnumerable<BrokenFormulaCellItem> items)
        {
            if (wb == null || items == null) return 0;
            int converted = 0;

            Sheets? sheets = null;
            try
            {
                sheets = wb.Worksheets;
                foreach (var item in items)
                {
                    if (!item.IsSelected) continue;

                    _Worksheet? ws = null;
                    Range? cell = null;
                    try
                    {
                        ws = sheets[item.SheetName] as _Worksheet;
                        if (ws != null)
                        {
                            cell = ws.Cells[item.Row, item.Column] as Range;
                            if (cell != null)
                            {
                                object? val = cell.Value2;
                                cell.Value2 = val;
                                converted++;
                            }
                        }
                    }
                    catch { }
                    finally
                    {
                        if (cell != null) Marshal.ReleaseComObject(cell);
                        if (ws != null) Marshal.ReleaseComObject(ws);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ConvertCellsToValues error: {ex.Message}");
            }
            finally
            {
                if (sheets != null) Marshal.ReleaseComObject(sheets);
            }

            return converted;
        }

        /// <summary>
        /// Đổi nguồn liên kết sang file mới (Change Source Link)
        /// </summary>
        public static bool ChangeLinkSource(Workbook wb, string oldSource, string newSource)
        {
            if (wb == null || string.IsNullOrWhiteSpace(oldSource) || string.IsNullOrWhiteSpace(newSource)) return false;
            try
            {
                wb.ChangeLink(oldSource, newSource, XlLinkType.xlLinkTypeExcelLinks);
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ChangeLinkSource error: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Tô màu đánh dấu các ô chứa link ngoài trên Excel
        /// </summary>
        public static int HighlightCellsOnExcel(Workbook wb, IEnumerable<BrokenFormulaCellItem> items)
        {
            if (wb == null || items == null) return 0;
            int highlighted = 0;

            Sheets? sheets = null;
            try
            {
                sheets = wb.Worksheets;
                foreach (var item in items)
                {
                    if (!item.IsSelected) continue;

                    _Worksheet? ws = null;
                    Range? cell = null;
                    try
                    {
                        ws = sheets[item.SheetName] as _Worksheet;
                        if (ws != null)
                        {
                            cell = ws.Cells[item.Row, item.Column] as Range;
                            if (cell != null)
                            {
                                // Màu vàng nhạt cho link ngoài, màu đỏ hồng cho link bị hỏng
                                if (item.IsBroken)
                                {
                                    cell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(254, 202, 202)); // #FECACA
                                }
                                else
                                {
                                    cell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(254, 240, 138)); // #FEF08A
                                }
                                highlighted++;
                            }
                        }
                    }
                    catch { }
                    finally
                    {
                        if (cell != null) Marshal.ReleaseComObject(cell);
                        if (ws != null) Marshal.ReleaseComObject(ws);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"HighlightCellsOnExcel error: {ex.Message}");
            }
            finally
            {
                if (sheets != null) Marshal.ReleaseComObject(sheets);
            }

            return highlighted;
        }

        /// <summary>
        /// Xóa các Named Ranges đã chọn
        /// </summary>
        public static int DeleteNamedRanges(Workbook wb, IEnumerable<ExternalNamedRangeItem> names)
        {
            if (wb == null || names == null) return 0;
            int deleted = 0;

            Names? wbNames = null;
            try
            {
                wbNames = wb.Names;
                if (wbNames != null)
                {
                    foreach (var item in names)
                    {
                        if (!item.IsSelected) continue;
                        try
                        {
                            Name? targetName = wbNames.Item(item.Name);
                            if (targetName != null)
                            {
                                targetName.Delete();
                                Marshal.ReleaseComObject(targetName);
                                deleted++;
                            }
                        }
                        catch { }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"DeleteNamedRanges error: {ex.Message}");
            }
            finally
            {
                if (wbNames != null) Marshal.ReleaseComObject(wbNames);
            }

            return deleted;
        }

        /// <summary>
        /// Tạo Sheet báo cáo chi tiết liên kết ngoài
        /// </summary>
        public static bool ExportReportToSheet(Workbook wb, ExternalLinksScanResult scanResult)
        {
            if (wb == null || scanResult == null) return false;

            _Worksheet? reportSheet = null;
            try
            {
                string baseName = "BaoCao_LinkNgoai";
                string sheetName = baseName;
                int suffix = 1;

                // Kiểm tra tên sheet trùng lặp
                while (true)
                {
                    bool exists = false;
                    foreach (_Worksheet s in wb.Worksheets)
                    {
                        if (string.Equals(s.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                        {
                            exists = true;
                            Marshal.ReleaseComObject(s);
                            break;
                        }
                        Marshal.ReleaseComObject(s);
                    }
                    if (!exists) break;
                    sheetName = $"{baseName}_{suffix++}";
                }

                reportSheet = wb.Worksheets.Add() as _Worksheet;
                if (reportSheet == null) return false;
                reportSheet.Name = sheetName;

                // Tiêu đề báo cáo
                reportSheet.Cells[1, 1] = "BÁO CÁO TỔNG HỢP LIÊN KẾT NGOÀI (EXTERNAL LINKS AUDIT)";
                Range titleRange = reportSheet.Range["A1", "F1"];
                titleRange.Font.Bold = true;
                titleRange.Font.Size = 14;
                titleRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(16, 124, 65));
                Marshal.ReleaseComObject(titleRange);

                reportSheet.Cells[2, 1] = $"Thời gian quét: {DateTime.Now:dd/MM/yyyy HH:mm:ss} | Tổng file ngoài: {scanResult.Sources.Count} | Tổng ô chứa link: {scanResult.FormulaCells.Count}";
                reportSheet.Range["A2"].Font.Italic = true;

                int row = 4;

                // Phần 1: Danh sách File Nguồn
                reportSheet.Cells[row, 1] = "1. DANH SÁCH FILE LIÊN KẾT NGOÀI (EXTERNAL SOURCES)";
                reportSheet.Range[$"A{row}"].Font.Bold = true;
                row++;

                reportSheet.Cells[row, 1] = "Tên File";
                reportSheet.Cells[row, 2] = "Đường Dẫn Đầy Đủ";
                reportSheet.Cells[row, 3] = "Trạng Thái";
                reportSheet.Cells[row, 4] = "Số Ô Công Thức";
                Range header1 = reportSheet.Range[$"A{row}", $"D{row}"];
                header1.Font.Bold = true;
                header1.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(226, 232, 240));
                Marshal.ReleaseComObject(header1);
                row++;

                foreach (var src in scanResult.Sources)
                {
                    reportSheet.Cells[row, 1] = src.FileName;
                    reportSheet.Cells[row, 2] = src.SourcePath;
                    reportSheet.Cells[row, 3] = src.StatusDisplay;
                    reportSheet.Cells[row, 4] = src.FormulaCount;
                    row++;
                }

                row += 2;

                // Phần 2: Chi tiết các ô công thức
                reportSheet.Cells[row, 1] = "2. CHI TIẾT CÁC Ô CÔNG THỨC THAM CHIẾU FILE NGOÀI";
                reportSheet.Range[$"A{row}"].Font.Bold = true;
                row++;

                reportSheet.Cells[row, 1] = "Tên Sheet";
                reportSheet.Cells[row, 2] = "Địa Chỉ Ô";
                reportSheet.Cells[row, 3] = "Công Thức Hiện Tại";
                reportSheet.Cells[row, 4] = "Giá Trị Tính Toán";
                reportSheet.Cells[row, 5] = "File Nguồn Tham Chiếu";
                reportSheet.Cells[row, 6] = "Trạng Thái";
                Range header2 = reportSheet.Range[$"A{row}", $"F{row}"];
                header2.Font.Bold = true;
                header2.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(226, 232, 240));
                Marshal.ReleaseComObject(header2);
                row++;

                foreach (var cell in scanResult.FormulaCells)
                {
                    reportSheet.Cells[row, 1] = cell.SheetName;
                    reportSheet.Cells[row, 2] = cell.CellAddress;
                    reportSheet.Cells[row, 3] = "'" + cell.Formula; // Thêm ' để hiển thị dạng text công thức
                    reportSheet.Cells[row, 4] = cell.CurrentValue;
                    reportSheet.Cells[row, 5] = cell.ExternalSource;
                    reportSheet.Cells[row, 6] = cell.IsBroken ? "❌ Lỗi / File không tồn tại" : "⚠️ Link ngoài";
                    row++;
                }

                // Tự động căn chỉnh độ rộng cột
                reportSheet.Columns.AutoFit();
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ExportReportToSheet error: {ex.Message}");
                return false;
            }
            finally
            {
                if (reportSheet != null) Marshal.ReleaseComObject(reportSheet);
            }
        }
    }
}
