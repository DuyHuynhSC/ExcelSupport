using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using ExcelSupport.Models;
using Microsoft.Office.Interop.Excel;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;

namespace ExcelSupport.Services
{
    public static class BatchFileConverterService
    {
        public static BatchConvertResult ExecuteBatchConversion(
            ExcelApp app, 
            BatchConvertOptions options, 
            Action<int, int, string>? progressCallback = null)
        {
            var result = new BatchConvertResult();
            if (app == null || options == null || options.InputFiles.Count == 0)
            {
                result.Success = false;
                result.Message = "Không có file nào trong danh sách cần xử lý.";
                return result;
            }

            if (string.IsNullOrEmpty(options.OutputDirectory))
            {
                result.Success = false;
                result.Message = "Vui lòng chọn thư mục lưu kết quả.";
                return result;
            }

            if (!Directory.Exists(options.OutputDirectory))
            {
                try { Directory.CreateDirectory(options.OutputDirectory); }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.Message = $"Không thể tạo thư mục lưu kết quả: {ex.Message}";
                    return result;
                }
            }

            int success = 0;
            int fail = 0;
            int total = options.InputFiles.Count;

            bool prevAlerts = app.DisplayAlerts;
            bool prevScreen = app.ScreenUpdating;

            try
            {
                app.DisplayAlerts = false;
                // KHÔNG tắt ScreenUpdating cố định ở cấp độ batch vì việc đóng mở SDI window
                // khi ScreenUpdating = false sẽ làm đóng băng giao diện và gây treo tiến trình EXCEL.EXE.

                switch (options.Mode)
                {
                    case BatchConvertMode.ConvertFormat:
                        for (int i = 0; i < total; i++)
                        {
                            string file = options.InputFiles[i];
                            progressCallback?.Invoke(i + 1, total, Path.GetFileName(file));
                            if (ConvertSingleFile(app, file, options))
                            {
                                success++;
                            }
                            else
                            {
                                fail++;
                            }
                        }
                        result.Message = $"Đã chuyển đổi định dạng thành công {success:N0}/{total:N0} tập tin sang {options.TargetFormat}!";
                        break;

                    case BatchConvertMode.SplitSheetsToFiles:
                        for (int i = 0; i < total; i++)
                        {
                            string file = options.InputFiles[i];
                            progressCallback?.Invoke(i + 1, total, Path.GetFileName(file));
                            int sheetsSplit = SplitWorkbookSheets(app, file, options.OutputDirectory, options.OverwriteExisting);
                            if (sheetsSplit > 0)
                            {
                                success++;
                            }
                            else
                            {
                                fail++;
                            }
                        }
                        result.Message = $"Đã tách thành công các Sheet từ {success:N0}/{total:N0} tập tin thành các file riêng!";
                        break;

                    case BatchConvertMode.MergeFilesToOne:
                        string outMergedPath = Path.Combine(options.OutputDirectory, options.MergedFileName);
                        progressCallback?.Invoke(1, 1, "Đang gộp tất cả các file...");
                        if (MergeMultipleFilesToOne(app, options.InputFiles, outMergedPath, options.OverwriteExisting))
                        {
                            success = total;
                            result.Message = $"Đã gộp thành công {total:N0} tập tin vào file duy nhất: '{Path.GetFileName(outMergedPath)}'!";
                        }
                        else
                        {
                            fail = total;
                            result.Message = "Không thể gộp các tập tin. Vui lòng kiểm tra quyền ghi và định dạng file.";
                        }
                        break;
                }

                result.Success = (success > 0);
                result.TotalFiles = total;
                result.SuccessCount = success;
                result.FailCount = fail;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"Lỗi xử lý file hàng loạt: {ex.Message}";
            }
            finally
            {
                try
                {
                    app.DisplayAlerts = prevAlerts;
                    app.ScreenUpdating = prevScreen;
                }
                catch { }

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            return result;
        }

        private static bool ConvertSingleFile(ExcelApp app, string inputPath, BatchConvertOptions options)
        {
            if (!File.Exists(inputPath)) return false;

            Workbooks? workbooks = null;
            Workbook? wb = null;
            bool wasAlreadyOpen = false;

            try
            {
                workbooks = app.Workbooks;
                int wbCount = workbooks.Count;

                // 1. Kiểm tra xem file có đang mở sẵn trong Excel không
                for (int i = 1; i <= wbCount; i++)
                {
                    Workbook? openWb = null;
                    try
                    {
                        openWb = workbooks[i];
                        if (string.Equals(openWb.FullName, inputPath, StringComparison.OrdinalIgnoreCase) ||
                            string.Equals(openWb.Name, Path.GetFileName(inputPath), StringComparison.OrdinalIgnoreCase))
                        {
                            wb = openWb;
                            wasAlreadyOpen = true;
                            break;
                        }
                    }
                    catch { }
                    finally
                    {
                        if (openWb != null && wb != openWb)
                        {
                            Marshal.ReleaseComObject(openWb);
                        }
                    }
                }

                // 2. Nếu chưa mở, mở file ở chế độ ReadOnly
                if (wb == null)
                {
                    wb = workbooks.Open(inputPath, ReadOnly: true, UpdateLinks: 0);
                    wasAlreadyOpen = false;
                }

                if (wb == null) return false;

                bool success = false;
                if (options.TargetFormat == ExcelOutputFormat.Markdown)
                {
                    success = ConvertWorkbookToMarkdown(app, wb, inputPath, options.OutputDirectory, options);
                }
                else
                {
                    string baseName = Path.GetFileNameWithoutExtension(inputPath);
                    string ext = GetExtensionForFormat(options.TargetFormat);
                    string outPath = Path.Combine(options.OutputDirectory, baseName + ext);

                    if (File.Exists(outPath))
                    {
                        if (!options.OverwriteExisting) return true;
                        try { File.Delete(outPath); } catch { }
                    }

                    if (options.TargetFormat == ExcelOutputFormat.PDF)
                    {
                        wb.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, outPath);
                    }
                    else
                    {
                        XlFileFormat xlFormat = GetXlFileFormat(options.TargetFormat);
                        wb.SaveAs(outPath, xlFormat);
                    }

                    success = true;
                }

                return success;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ConvertSingleFile error ({inputPath}): {ex.Message}");
                return false;
            }
            finally
            {
                if (wb != null)
                {
                    // Chỉ đóng workbook nếu file đó do converter tự mở từ ổ đĩa.
                    // Nếu file vốn đang được người dùng mở sẵn trong Excel (wasAlreadyOpen = true),
                    // KHÔNG ĐƯỢC đóng, vì sẽ làm mất file của người dùng và để lại cửa sổ Excel xám xịt trống rỗng!
                    if (!wasAlreadyOpen)
                    {
                        try
                        {
                            bool currentScreen = app.ScreenUpdating;
                            if (!currentScreen) app.ScreenUpdating = true;

                            wb.Close(SaveChanges: false);

                            if (!currentScreen) app.ScreenUpdating = false;
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Error closing wb ({inputPath}): {ex.Message}");
                        }
                    }

                    Marshal.ReleaseComObject(wb);
                    wb = null;
                }

                if (workbooks != null)
                {
                    Marshal.ReleaseComObject(workbooks);
                    workbooks = null;
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private static int SplitWorkbookSheets(ExcelApp app, string inputPath, string outputDir, bool overwrite)
        {
            if (!File.Exists(inputPath)) return 0;

            Workbooks? workbooks = null;
            Workbook? wb = null;
            Sheets? sheets = null;
            int splitCount = 0;
            bool wasAlreadyOpen = false;

            try
            {
                workbooks = app.Workbooks;
                int wbCount = workbooks.Count;

                for (int i = 1; i <= wbCount; i++)
                {
                    Workbook? openWb = null;
                    try
                    {
                        openWb = workbooks[i];
                        if (string.Equals(openWb.FullName, inputPath, StringComparison.OrdinalIgnoreCase) ||
                            string.Equals(openWb.Name, Path.GetFileName(inputPath), StringComparison.OrdinalIgnoreCase))
                        {
                            wb = openWb;
                            wasAlreadyOpen = true;
                            break;
                        }
                    }
                    catch { }
                    finally
                    {
                        if (openWb != null && wb != openWb)
                        {
                            Marshal.ReleaseComObject(openWb);
                        }
                    }
                }

                if (wb == null)
                {
                    wb = workbooks.Open(inputPath, ReadOnly: true, UpdateLinks: 0);
                    wasAlreadyOpen = false;
                }

                if (wb == null) return 0;

                string baseName = Path.GetFileNameWithoutExtension(inputPath);
                sheets = wb.Worksheets;
                int count = sheets.Count;

                for (int i = 1; i <= count; i++)
                {
                    _Worksheet? ws = null;
                    Workbook? newWb = null;
                    try
                    {
                        ws = (_Worksheet)sheets[i];
                        string safeSheetName = SanitizeFileName(ws.Name);
                        string outPath = Path.Combine(outputDir, $"{baseName}_{safeSheetName}.xlsx");

                        if (File.Exists(outPath))
                        {
                            if (!overwrite) continue;
                            try { File.Delete(outPath); } catch { }
                        }

                        // Copy sheet sang 1 workbook mới
                        ws.Copy();
                        newWb = app.ActiveWorkbook;
                        if (newWb != null)
                        {
                            newWb.SaveAs(outPath, XlFileFormat.xlOpenXMLWorkbook);
                            newWb.Close(SaveChanges: false);
                            splitCount++;
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Split sheet error: {ex.Message}");
                    }
                    finally
                    {
                        if (newWb != null) Marshal.ReleaseComObject(newWb);
                        if (ws != null) Marshal.ReleaseComObject(ws);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"SplitWorkbookSheets error: {ex.Message}");
            }
            finally
            {
                if (sheets != null)
                {
                    Marshal.ReleaseComObject(sheets);
                    sheets = null;
                }

                if (wb != null)
                {
                    if (!wasAlreadyOpen)
                    {
                        try
                        {
                            bool currentScreen = app.ScreenUpdating;
                            if (!currentScreen) app.ScreenUpdating = true;

                            wb.Close(SaveChanges: false);

                            if (!currentScreen) app.ScreenUpdating = false;
                        }
                        catch { }
                    }
                    Marshal.ReleaseComObject(wb);
                    wb = null;
                }

                if (workbooks != null)
                {
                    Marshal.ReleaseComObject(workbooks);
                    workbooks = null;
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            return splitCount;
        }

        private static bool MergeMultipleFilesToOne(ExcelApp app, List<string> inputFiles, string outMergedPath, bool overwrite)
        {
            if (inputFiles == null || inputFiles.Count == 0) return false;

            Workbooks? workbooks = null;
            Workbook? masterWb = null;
            Sheets? masterSheets = null;
            try
            {
                if (File.Exists(outMergedPath))
                {
                    if (!overwrite) return false;
                    try { File.Delete(outMergedPath); } catch { }
                }

                workbooks = app.Workbooks;
                masterWb = workbooks.Add();
                if (masterWb == null) return false;

                masterSheets = masterWb.Worksheets;
                var initialSheets = new List<_Worksheet>();
                int initCount = masterSheets.Count;
                for (int i = 1; i <= initCount; i++)
                {
                    initialSheets.Add((_Worksheet)masterSheets[i]);
                }

                var usedSheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                foreach (var file in inputFiles)
                {
                    if (!File.Exists(file)) continue;

                    Workbook? srcWb = null;
                    Sheets? srcSheets = null;
                    bool srcWasAlreadyOpen = false;
                    try
                    {
                        int openCount = workbooks.Count;
                        for (int k = 1; k <= openCount; k++)
                        {
                            Workbook? owb = null;
                            try
                            {
                                owb = workbooks[k];
                                if (string.Equals(owb.FullName, file, StringComparison.OrdinalIgnoreCase) ||
                                    string.Equals(owb.Name, Path.GetFileName(file), StringComparison.OrdinalIgnoreCase))
                                {
                                    srcWb = owb;
                                    srcWasAlreadyOpen = true;
                                    break;
                                }
                            }
                            catch { }
                            finally
                            {
                                if (owb != null && srcWb != owb) Marshal.ReleaseComObject(owb);
                            }
                        }

                        if (srcWb == null)
                        {
                            srcWb = workbooks.Open(file, ReadOnly: true, UpdateLinks: 0);
                            srcWasAlreadyOpen = false;
                        }

                        if (srcWb == null) continue;

                        string srcFileName = Path.GetFileNameWithoutExtension(file);
                        srcSheets = srcWb.Worksheets;
                        int srcCount = srcSheets.Count;

                        for (int i = 1; i <= srcCount; i++)
                        {
                            _Worksheet? srcWs = null;
                            _Worksheet? lastSheet = null;
                            _Worksheet? newCopiedSheet = null;
                            try
                            {
                                srcWs = (_Worksheet)srcSheets[i];
                                string candidateName = (srcCount == 1)
                                    ? srcFileName
                                    : $"{srcFileName}_{srcWs.Name}";

                                candidateName = SanitizeFileName(candidateName);
                                if (candidateName.Length > 31) candidateName = candidateName.Substring(0, 31);

                                string uniqueName = candidateName;
                                int counter = 1;
                                while (usedSheetNames.Contains(uniqueName))
                                {
                                    string suffix = $"_{counter++}";
                                    int maxLen = 31 - suffix.Length;
                                    uniqueName = (candidateName.Length > maxLen ? candidateName.Substring(0, maxLen) : candidateName) + suffix;
                                }
                                usedSheetNames.Add(uniqueName);

                                // Copy vào sau sheet cuối của masterWb
                                lastSheet = (_Worksheet)masterSheets[masterSheets.Count];
                                srcWs.Copy(After: lastSheet);
                                newCopiedSheet = (_Worksheet)masterSheets[masterSheets.Count];
                                newCopiedSheet.Name = uniqueName;
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine($"Copy sheet in merge error: {ex.Message}");
                            }
                            finally
                            {
                                if (lastSheet != null) Marshal.ReleaseComObject(lastSheet);
                                if (newCopiedSheet != null) Marshal.ReleaseComObject(newCopiedSheet);
                                if (srcWs != null) Marshal.ReleaseComObject(srcWs);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Merge file error ({file}): {ex.Message}");
                    }
                    finally
                    {
                        if (srcSheets != null)
                        {
                            Marshal.ReleaseComObject(srcSheets);
                            srcSheets = null;
                        }

                        if (srcWb != null)
                        {
                            if (!srcWasAlreadyOpen)
                            {
                                try { srcWb.Close(SaveChanges: false); } catch { }
                            }
                            Marshal.ReleaseComObject(srcWb);
                            srcWb = null;
                        }
                    }
                }

                // Xóa các sheet mặc định ban đầu của masterWb
                foreach (var initSheet in initialSheets)
                {
                    try { initSheet.Delete(); } catch { }
                    Marshal.ReleaseComObject(initSheet);
                }
                initialSheets.Clear();

                masterWb.SaveAs(outMergedPath, XlFileFormat.xlOpenXMLWorkbook);
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"MergeMultipleFilesToOne error: {ex.Message}");
                return false;
            }
            finally
            {
                if (masterSheets != null)
                {
                    Marshal.ReleaseComObject(masterSheets);
                    masterSheets = null;
                }

                if (masterWb != null)
                {
                    try
                    {
                        bool currentScreen = app.ScreenUpdating;
                        if (!currentScreen) app.ScreenUpdating = true;

                        masterWb.Close(SaveChanges: false);

                        if (!currentScreen) app.ScreenUpdating = false;
                    }
                    catch { }
                    Marshal.ReleaseComObject(masterWb);
                    masterWb = null;
                }

                if (workbooks != null)
                {
                    Marshal.ReleaseComObject(workbooks);
                    workbooks = null;
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private static string GetExtensionForFormat(ExcelOutputFormat format)
        {
            return format switch
            {
                ExcelOutputFormat.XLSX => ".xlsx",
                ExcelOutputFormat.XLS => ".xls",
                ExcelOutputFormat.XLSB => ".xlsb",
                ExcelOutputFormat.XLSM => ".xlsm",
                ExcelOutputFormat.CSV => ".csv",
                ExcelOutputFormat.PDF => ".pdf",
                ExcelOutputFormat.Markdown => ".md",
                _ => ".xlsx"
            };
        }

        private static XlFileFormat GetXlFileFormat(ExcelOutputFormat format)
        {
            return format switch
            {
                ExcelOutputFormat.XLSX => XlFileFormat.xlOpenXMLWorkbook,
                ExcelOutputFormat.XLS => XlFileFormat.xlExcel8,
                ExcelOutputFormat.XLSB => XlFileFormat.xlExcel12,
                ExcelOutputFormat.XLSM => XlFileFormat.xlOpenXMLWorkbookMacroEnabled,
                ExcelOutputFormat.CSV => XlFileFormat.xlCSV,
                _ => XlFileFormat.xlOpenXMLWorkbook
            };
        }

        private static string SanitizeFileName(string name)
        {
            if (string.IsNullOrEmpty(name)) return "Sheet";
            char[] invalidChars = Path.GetInvalidFileNameChars();
            char[] excelInvalid = new char[] { '\\', '/', '?', '*', '[', ']', ':' };
            var allInvalid = new HashSet<char>(invalidChars.Concat(excelInvalid));

            var sb = new System.Text.StringBuilder();
            foreach (char c in name)
            {
                sb.Append(allInvalid.Contains(c) ? '_' : c);
            }
            return sb.ToString();
        }

        private static bool ConvertWorkbookToMarkdown(ExcelApp app, Workbook wb, string inputPath, string outputDir, BatchConvertOptions options)
        {
            string baseName = Path.GetFileNameWithoutExtension(inputPath);
            var eligibleSheets = new List<_Worksheet>();
            Sheets? sheets = null;

            try
            {
                sheets = wb.Worksheets;
                int sheetCount = sheets.Count;

                for (int i = 1; i <= sheetCount; i++)
                {
                    _Worksheet? ws = null;
                    try
                    {
                        ws = (_Worksheet)sheets[i];
                        if (ShouldProcessSheet(ws.Name, options.SheetFilterMode, options.SheetFilterPatterns))
                        {
                            eligibleSheets.Add(ws);
                            ws = null; // Giữ lại để giải phóng sau trong finally
                        }
                    }
                    finally
                    {
                        if (ws != null) Marshal.ReleaseComObject(ws);
                    }
                }

                if (eligibleSheets.Count == 0)
                {
                    System.Diagnostics.Debug.WriteLine($"Không có sheet nào khớp bộ lọc trong file: {inputPath}");
                    return false;
                }

                if (options.MarkdownMode == MarkdownSheetExportMode.SeparateFilePerSheet)
                {
                    int convertedCount = 0;
                    foreach (var ws in eligibleSheets)
                    {
                        try
                        {
                            string safeSheetName = SanitizeFileName(ws.Name);
                            string outPath = Path.Combine(outputDir, $"{baseName}_{safeSheetName}.md");

                            if (File.Exists(outPath))
                            {
                                if (!options.OverwriteExisting)
                                {
                                    convertedCount++;
                                    continue;
                                }
                                try { File.Delete(outPath); } catch { }
                            }

                            var sb = new System.Text.StringBuilder();
                            sb.AppendLine($"# {ws.Name}");
                            sb.AppendLine();

                            string tableMd = ConvertWorksheetToMarkdownTable(ws, options.StartRow, options.ConvertLineBreaksToBr);
                            if (!string.IsNullOrWhiteSpace(tableMd))
                            {
                                sb.AppendLine(tableMd);
                            }
                            else
                            {
                                sb.AppendLine("*Sheet trống không có dữ liệu.*");
                            }

                            File.WriteAllText(outPath, sb.ToString(), new System.Text.UTF8Encoding(false));
                            convertedCount++;
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Convert sheet to MD error ({ws.Name}): {ex.Message}");
                        }
                    }
                    return convertedCount > 0;
                }
                else
                {
                    // Gộp tất cả Sheet vào 1 file duy nhất
                    string outPath = Path.Combine(outputDir, $"{baseName}.md");
                    if (File.Exists(outPath))
                    {
                        if (!options.OverwriteExisting) return true;
                        try { File.Delete(outPath); } catch { }
                    }

                    var sb = new System.Text.StringBuilder();
                    sb.AppendLine($"# {baseName}");
                    sb.AppendLine();

                    if (options.IncludeMarkdownToc && eligibleSheets.Count > 1)
                    {
                        sb.AppendLine("## Mục Lục");
                        sb.AppendLine();
                        foreach (var ws in eligibleSheets)
                        {
                            string slug = GenerateSlug(ws.Name);
                            sb.AppendLine($"- [{ws.Name}](#{slug})");
                        }
                        sb.AppendLine();
                        sb.AppendLine("---");
                        sb.AppendLine();
                    }

                    for (int i = 0; i < eligibleSheets.Count; i++)
                    {
                        var ws = eligibleSheets[i];
                        sb.AppendLine($"## {ws.Name}");
                        sb.AppendLine();

                        string tableMd = ConvertWorksheetToMarkdownTable(ws, options.StartRow, options.ConvertLineBreaksToBr);
                        if (!string.IsNullOrWhiteSpace(tableMd))
                        {
                            sb.AppendLine(tableMd);
                        }
                        else
                        {
                            sb.AppendLine("*Sheet trống không có dữ liệu.*");
                        }

                        if (i < eligibleSheets.Count - 1)
                        {
                            sb.AppendLine();
                            sb.AppendLine("---");
                            sb.AppendLine();
                        }
                    }

                    File.WriteAllText(outPath, sb.ToString(), new System.Text.UTF8Encoding(false));
                    return true;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ConvertWorkbookToMarkdown error: {ex.Message}");
                return false;
            }
            finally
            {
                foreach (var ws in eligibleSheets)
                {
                    Marshal.ReleaseComObject(ws);
                }
                eligibleSheets.Clear();

                if (sheets != null)
                {
                    Marshal.ReleaseComObject(sheets);
                    sheets = null;
                }
            }
        }

        private static string ConvertWorksheetToMarkdownTable(_Worksheet ws, int startRow, bool convertLineBreaksToBr)
        {
            Range? usedRange = null;
            Range? rowsRange = null;
            Range? colsRange = null;
            Range? cellTopLeft = null;
            Range? cellBottomRight = null;
            Range? targetRange = null;
            try
            {
                usedRange = ws.UsedRange;
                if (usedRange == null) return string.Empty;

                rowsRange = usedRange.Rows;
                colsRange = usedRange.Columns;

                int usedStartRow = usedRange.Row;
                int usedStartCol = usedRange.Column;
                int usedRowCount = rowsRange.Count;
                int usedColCount = colsRange.Count;

                if (usedRowCount == 0 || usedColCount == 0) return string.Empty;

                int reqStartRow = Math.Max(1, startRow);
                int lastRow = usedStartRow + usedRowCount - 1;
                if (reqStartRow > lastRow)
                {
                    return string.Empty;
                }

                int effectiveStartRow = Math.Max(reqStartRow, usedStartRow);
                int rowsToExtract = lastRow - effectiveStartRow + 1;
                int colsToExtract = usedColCount;

                if (rowsToExtract <= 0 || colsToExtract <= 0) return string.Empty;

                cellTopLeft = (Range)ws.Cells[effectiveStartRow, usedStartCol];
                cellBottomRight = (Range)ws.Cells[effectiveStartRow + rowsToExtract - 1, usedStartCol + colsToExtract - 1];
                targetRange = ws.Range[cellTopLeft, cellBottomRight];

                object[,] matrix;
                if (rowsToExtract == 1 && colsToExtract == 1)
                {
                    object? singleVal = targetRange.Value2;
                    if (singleVal == null || string.IsNullOrWhiteSpace(singleVal.ToString())) return string.Empty;
                    matrix = new object[2, 2];
                    matrix[1, 1] = singleVal;
                }
                else
                {
                    object? rawData = null;
                    try
                    {
                        rawData = targetRange.Value;
                    }
                    catch
                    {
                        rawData = targetRange.Value2;
                    }

                    if (rawData is object[,] arr)
                    {
                        matrix = arr;
                    }
                    else
                    {
                        return string.Empty;
                    }
                }

                int rMin = matrix.GetLowerBound(0);
                int rMax = matrix.GetUpperBound(0);
                int cMin = matrix.GetLowerBound(1);
                int cMax = matrix.GetUpperBound(1);

                bool hasAnyData = false;
                for (int r = rMin; r <= rMax; r++)
                {
                    for (int c = cMin; c <= cMax; c++)
                    {
                        if (matrix[r, c] != null && !string.IsNullOrWhiteSpace(matrix[r, c].ToString()))
                        {
                            hasAnyData = true;
                            break;
                        }
                    }
                    if (hasAnyData) break;
                }
                if (!hasAnyData) return string.Empty;

                var sb = new System.Text.StringBuilder();

                // 1. Dòng Header (dòng rMin)
                sb.Append("|");
                for (int c = cMin; c <= cMax; c++)
                {
                    string headerVal = FormatCellMarkdown(matrix[rMin, c], convertLineBreaksToBr);
                    if (string.IsNullOrWhiteSpace(headerVal))
                    {
                        int actualCol = usedStartCol + (c - cMin);
                        headerVal = $"Cột {ConvertColIndexToLetter(actualCol)}";
                    }
                    sb.Append($" {headerVal} |");
                }
                sb.AppendLine();

                // 2. Dòng căn lề (Separator Row)
                sb.Append("|");
                for (int c = cMin; c <= cMax; c++)
                {
                    int numCount = 0;
                    int textCount = 0;
                    for (int r = rMin + 1; r <= rMax; r++)
                    {
                        object? v = matrix[r, c];
                        if (v == null) continue;
                        if (v is double || v is int || v is long || v is decimal || v is float)
                            numCount++;
                        else if (!string.IsNullOrWhiteSpace(v.ToString()))
                            textCount++;
                    }

                    if (numCount > 0 && numCount >= textCount)
                    {
                        sb.Append(" ---: |");
                    }
                    else
                    {
                        sb.Append(" :--- |");
                    }
                }
                sb.AppendLine();

                // 3. Các dòng dữ liệu (từ rMin + 1 đến rMax)
                for (int r = rMin + 1; r <= rMax; r++)
                {
                    sb.Append("|");
                    for (int c = cMin; c <= cMax; c++)
                    {
                        string cellVal = FormatCellMarkdown(matrix[r, c], convertLineBreaksToBr);
                        sb.Append($" {cellVal} |");
                    }
                    sb.AppendLine();
                }

                return sb.ToString().TrimEnd();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ConvertWorksheetToMarkdownTable error: {ex.Message}");
                return string.Empty;
            }
            finally
            {
                if (cellTopLeft != null) Marshal.ReleaseComObject(cellTopLeft);
                if (cellBottomRight != null) Marshal.ReleaseComObject(cellBottomRight);
                if (targetRange != null) Marshal.ReleaseComObject(targetRange);
                if (rowsRange != null) Marshal.ReleaseComObject(rowsRange);
                if (colsRange != null) Marshal.ReleaseComObject(colsRange);
                if (usedRange != null) Marshal.ReleaseComObject(usedRange);
            }
        }

        private static string FormatCellMarkdown(object? val, bool convertLineBreaksToBr)
        {
            if (val == null) return string.Empty;

            string text;
            if (val is DateTime dt)
            {
                text = (dt.TimeOfDay.TotalSeconds == 0)
                    ? dt.ToString("yyyy-MM-dd")
                    : dt.ToString("yyyy-MM-dd HH:mm:ss");
            }
            else if (val is double d)
            {
                if (d % 1 == 0 && d >= -1e15 && d <= 1e15)
                {
                    text = ((long)d).ToString();
                }
                else
                {
                    text = d.ToString("G15", System.Globalization.CultureInfo.InvariantCulture);
                }
            }
            else if (val is bool b)
            {
                text = b ? "TRUE" : "FALSE";
            }
            else
            {
                text = val.ToString() ?? string.Empty;
            }

            if (string.IsNullOrEmpty(text)) return string.Empty;

            text = text.Replace("|", "\\|");

            if (convertLineBreaksToBr)
            {
                text = text.Replace("\r\n", "<br>")
                           .Replace("\n", "<br>")
                           .Replace("\r", "<br>");
            }
            else
            {
                text = text.Replace("\r\n", " ")
                           .Replace("\n", " ")
                           .Replace("\r", " ");
            }

            return text.Trim();
        }

        private static string ConvertColIndexToLetter(int colIndex)
        {
            int div = colIndex;
            string colLetter = string.Empty;
            while (div > 0)
            {
                int mod = (div - 1) % 26;
                colLetter = (char)(65 + mod) + colLetter;
                div = (div - mod) / 26;
            }
            return colLetter;
        }

        private static string GenerateSlug(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return "sheet";
            var sb = new System.Text.StringBuilder();
            foreach (char c in text.Trim().ToLowerInvariant())
            {
                if (char.IsLetterOrDigit(c))
                {
                    sb.Append(c);
                }
                else if (c == ' ' || c == '-' || c == '_')
                {
                    sb.Append('-');
                }
            }
            string slug = sb.ToString();
            while (slug.Contains("--")) slug = slug.Replace("--", "-");
            return slug.Trim('-');
        }

        private static bool ShouldProcessSheet(string sheetName, MarkdownSheetFilterMode mode, string patterns)
        {
            if (mode == MarkdownSheetFilterMode.All || string.IsNullOrWhiteSpace(patterns))
            {
                return true;
            }

            var tokens = patterns.Split(new[] { ',', ';', '|' }, StringSplitOptions.RemoveEmptyEntries)
                                 .Select(t => t.Trim())
                                 .Where(t => !string.IsNullOrEmpty(t))
                                 .ToList();

            if (tokens.Count == 0) return true;

            bool matches = tokens.Any(token =>
            {
                if (token.Contains("*") || token.Contains("?"))
                {
                    string regexPattern = "^" + System.Text.RegularExpressions.Regex.Escape(token)
                                                    .Replace("\\*", ".*")
                                                    .Replace("\\?", ".") + "$";
                    return System.Text.RegularExpressions.Regex.IsMatch(sheetName, regexPattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                }
                return string.Equals(sheetName, token, StringComparison.OrdinalIgnoreCase);
            });

            return mode == MarkdownSheetFilterMode.IncludeOnly ? matches : !matches;
        }
    }
}
