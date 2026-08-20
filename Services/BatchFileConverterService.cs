using System;
using System.Collections.Generic;
using System.IO;
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
                app.ScreenUpdating = false;

                switch (options.Mode)
                {
                    case BatchConvertMode.ConvertFormat:
                        for (int i = 0; i < total; i++)
                        {
                            string file = options.InputFiles[i];
                            progressCallback?.Invoke(i + 1, total, Path.GetFileName(file));
                            if (ConvertSingleFile(app, file, options.OutputDirectory, options.TargetFormat, options.OverwriteExisting))
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
            }

            return result;
        }

        private static bool ConvertSingleFile(ExcelApp app, string inputPath, string outputDir, ExcelOutputFormat format, bool overwrite)
        {
            if (!File.Exists(inputPath)) return false;

            Workbook? wb = null;
            try
            {
                wb = app.Workbooks.Open(inputPath, ReadOnly: true, UpdateLinks: 0);
                if (wb == null) return false;

                string baseName = Path.GetFileNameWithoutExtension(inputPath);
                string ext = GetExtensionForFormat(format);
                string outPath = Path.Combine(outputDir, baseName + ext);

                if (File.Exists(outPath))
                {
                    if (!overwrite) return true;
                    try { File.Delete(outPath); } catch { }
                }

                if (format == ExcelOutputFormat.PDF)
                {
                    wb.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, outPath);
                }
                else
                {
                    XlFileFormat xlFormat = GetXlFileFormat(format);
                    wb.SaveAs(outPath, xlFormat);
                }

                return true;
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
                    try { wb.Close(SaveChanges: false); } catch { }
                    Marshal.ReleaseComObject(wb);
                }
            }
        }

        private static int SplitWorkbookSheets(ExcelApp app, string inputPath, string outputDir, bool overwrite)
        {
            if (!File.Exists(inputPath)) return 0;

            Workbook? wb = null;
            int splitCount = 0;

            try
            {
                wb = app.Workbooks.Open(inputPath, ReadOnly: true, UpdateLinks: 0);
                if (wb == null) return 0;

                string baseName = Path.GetFileNameWithoutExtension(inputPath);

                foreach (_Worksheet ws in wb.Worksheets)
                {
                    Workbook? newWb = null;
                    try
                    {
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
                        Marshal.ReleaseComObject(ws);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"SplitWorkbookSheets error: {ex.Message}");
            }
            finally
            {
                if (wb != null)
                {
                    try { wb.Close(SaveChanges: false); } catch { }
                    Marshal.ReleaseComObject(wb);
                }
            }

            return splitCount;
        }

        private static bool MergeMultipleFilesToOne(ExcelApp app, List<string> inputFiles, string outMergedPath, bool overwrite)
        {
            if (inputFiles == null || inputFiles.Count == 0) return false;

            Workbook? masterWb = null;
            try
            {
                if (File.Exists(outMergedPath))
                {
                    if (!overwrite) return false;
                    try { File.Delete(outMergedPath); } catch { }
                }

                masterWb = app.Workbooks.Add();
                if (masterWb == null) return false;

                var initialSheets = new List<_Worksheet>();
                foreach (_Worksheet s in masterWb.Worksheets) initialSheets.Add(s);

                var usedSheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                foreach (var file in inputFiles)
                {
                    if (!File.Exists(file)) continue;

                    Workbook? srcWb = null;
                    try
                    {
                        srcWb = app.Workbooks.Open(file, ReadOnly: true, UpdateLinks: 0);
                        if (srcWb == null) continue;

                        string srcFileName = Path.GetFileNameWithoutExtension(file);

                        foreach (_Worksheet srcWs in srcWb.Worksheets)
                        {
                            try
                            {
                                string candidateName = (srcWb.Worksheets.Count == 1)
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
                                _Worksheet lastSheet = (_Worksheet)masterWb.Worksheets[masterWb.Worksheets.Count];
                                srcWs.Copy(After: lastSheet);
                                _Worksheet newCopiedSheet = (_Worksheet)masterWb.Worksheets[masterWb.Worksheets.Count];
                                newCopiedSheet.Name = uniqueName;

                                Marshal.ReleaseComObject(lastSheet);
                                Marshal.ReleaseComObject(newCopiedSheet);
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine($"Copy sheet in merge error: {ex.Message}");
                            }
                            finally
                            {
                                Marshal.ReleaseComObject(srcWs);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Merge file error ({file}): {ex.Message}");
                    }
                    finally
                    {
                        if (srcWb != null)
                        {
                            try { srcWb.Close(SaveChanges: false); } catch { }
                            Marshal.ReleaseComObject(srcWb);
                        }
                    }
                }

                // Xóa các sheet mặc định ban đầu của masterWb
                foreach (var initSheet in initialSheets)
                {
                    try { initSheet.Delete(); } catch { }
                    Marshal.ReleaseComObject(initSheet);
                }

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
                if (masterWb != null)
                {
                    try { masterWb.Close(SaveChanges: false); } catch { }
                    Marshal.ReleaseComObject(masterWb);
                }
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
    }
}
