using System;
using System.Collections.Generic;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using WpfMessageBox = System.Windows.MessageBox;

namespace ExcelSupport
{
    public partial class AddInEvents
    {
        #region Excel COM Helper Methods for Rename, Split, and Merge Sheets

        public bool RenameWorksheet(string wbName, string oldSheetName, string newSheetName)
        {
            if (_excelApp == null || string.IsNullOrWhiteSpace(newSheetName)) return false;

            // Kiểm tra quy chuẩn đặt tên sheet Excel: tối đa 31 ký tự, không chứa \ / ? * [ ] :
            string cleanName = newSheetName.Trim();
            if (cleanName.Length > 31)
            {
                WpfMessageBox.Show("Tên Sheet không được vượt quá 31 ký tự.", "Đổi Tên Sheet",
                                   System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                return false;
            }

            char[] invalidChars = { '\\', '/', '?', '*', '[', ']', ':' };
            if (cleanName.IndexOfAny(invalidChars) >= 0)
            {
                WpfMessageBox.Show("Tên Sheet không được chứa các ký tự đặc biệt: \\ / ? * [ ] :", "Đổi Tên Sheet",
                                   System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                return false;
            }

            try
            {
                dynamic app = _excelApp;
                dynamic? targetWb = null;
                if (!string.IsNullOrEmpty(wbName))
                {
                    try { targetWb = app.Workbooks[wbName]; } catch { }
                }
                if (targetWb == null)
                {
                    try { targetWb = app.ActiveWorkbook; } catch { }
                }

                if (targetWb == null) return false;

                dynamic ws = targetWb.Worksheets[oldSheetName];
                if (ws != null)
                {
                    ws.Name = cleanName;
                    RefreshWorkbookTree();
                    return true;
                }
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Không thể đổi tên sheet:\n{ex.Message}", "Đổi Tên Sheet",
                                   System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
            }
            return false;
        }

        public bool BatchRenameWorksheets(string wbName, string prefix, string suffix, string findText, string replaceText)
        {
            if (_excelApp == null) return false;

            _isBatchProcessing = true;
            dynamic app = _excelApp;

            try
            {
                try { app.EnableEvents = false; } catch { }
                try { app.ScreenUpdating = false; } catch { }

                dynamic? targetWb = null;
                if (!string.IsNullOrEmpty(wbName))
                {
                    try { targetWb = app.Workbooks[wbName]; } catch { }
                }
                if (targetWb == null)
                {
                    try { targetWb = app.ActiveWorkbook; } catch { }
                }

                if (targetWb == null) return false;

                int count = targetWb.Sheets.Count;
                int renamedCount = 0;

                for (int i = 1; i <= count; i++)
                {
                    try
                    {
                        dynamic ws = targetWb.Sheets[i];
                        string currentName = ws.Name;
                        string newName = currentName;

                        if (!string.IsNullOrEmpty(findText))
                        {
                            newName = newName.Replace(findText, replaceText ?? string.Empty);
                        }

                        if (!string.IsNullOrEmpty(prefix))
                        {
                            newName = prefix + newName;
                        }

                        if (!string.IsNullOrEmpty(suffix))
                        {
                            newName = newName + suffix;
                        }

                        if (newName.Length > 31)
                        {
                            newName = newName.Substring(0, 31);
                        }

                        if (newName != currentName)
                        {
                            ws.Name = newName;
                            renamedCount++;
                        }
                    }
                    catch { }
                }

                _isBatchProcessing = false;
                try { app.EnableEvents = true; } catch { }
                try { app.ScreenUpdating = true; } catch { }

                RefreshWorkbookTree();
                WpfMessageBox.Show($"✅ Đã đổi tên thành công cho {renamedCount} sheet!", "Đổi Tên Hàng Loạt",
                                   System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
                return true;
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Lỗi đổi tên hàng loạt:\n{ex.Message}", "Đổi Tên Hàng Loạt",
                                   System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                return false;
            }
            finally
            {
                _isBatchProcessing = false;
                if (_excelApp != null)
                {
                    try { _excelApp.EnableEvents = true; } catch { }
                    try { _excelApp.ScreenUpdating = true; } catch { }
                }
            }
        }

        public bool SplitWorksheetsToFiles(string wbName, List<string>? sheetNames, string outputFolder, bool keepOriginalSheets = true)
        {
            if (_excelApp == null || string.IsNullOrWhiteSpace(outputFolder)) return false;

            if (!System.IO.Directory.Exists(outputFolder))
            {
                try { System.IO.Directory.CreateDirectory(outputFolder); }
                catch (Exception ex)
                {
                    WpfMessageBox.Show($"Không thể tạo thư mục lưu:\n{ex.Message}", "Tách Sheet",
                                       System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                    return false;
                }
            }

            _isBatchProcessing = true;
            dynamic app = _excelApp;

            try
            {
                try { app.EnableEvents = false; } catch { }
                try { app.DisplayAlerts = false; } catch { }
                try { app.ScreenUpdating = false; } catch { }

                dynamic? targetWb = null;
                if (!string.IsNullOrEmpty(wbName))
                {
                    try { targetWb = app.Workbooks[wbName]; } catch { }
                }
                if (targetWb == null)
                {
                    try { targetWb = app.ActiveWorkbook; } catch { }
                }

                if (targetWb == null) return false;

                int totalSheets = targetWb.Sheets.Count;
                int exportedCount = 0;
                var exportedSheetNames = new List<string>();

                for (int i = 1; i <= totalSheets; i++)
                {
                    try
                    {
                        dynamic ws = targetWb.Sheets[i];
                        string sheetName = ws.Name;

                        if (sheetNames != null && sheetNames.Count > 0 && !sheetNames.Contains(sheetName))
                        {
                            continue;
                        }

                        // Sao chép sheet sang một Workbook mới hoàn toàn
                        ws.Copy();
                        dynamic newWb = app.ActiveWorkbook;

                        if (newWb != null)
                        {
                            // Chuẩn hóa tên file xuất
                            string cleanSheetName = string.Join("_", sheetName.Split(System.IO.Path.GetInvalidFileNameChars()));
                            string filePath = System.IO.Path.Combine(outputFolder, $"{cleanSheetName}.xlsx");

                            // Lưu file và đóng lại
                            newWb.SaveAs(filePath, 51); // 51 = xlOpenXMLWorkbook (.xlsx)
                            newWb.Close(false);
                            exportedCount++;
                            exportedSheetNames.Add(sheetName);
                        }
                    }
                    catch (Exception exSheet)
                    {
                        System.Diagnostics.Debug.WriteLine($"Split sheet error: {exSheet.Message}");
                    }
                }

                // Nếu người dùng chọn xóa sheet gốc sau khi tách
                if (!keepOriginalSheets && exportedSheetNames.Count > 0)
                {
                    try
                    {
                        // Excel bắt buộc phải có ít nhất 1 sheet hiển thị
                        if (targetWb.Sheets.Count <= exportedSheetNames.Count)
                        {
                            dynamic emptyWs = targetWb.Worksheets.Add();
                            emptyWs.Name = "Sheet1";
                        }

                        foreach (var sName in exportedSheetNames)
                        {
                            try
                            {
                                dynamic wsDel = targetWb.Sheets[sName];
                                if (wsDel != null)
                                {
                                    wsDel.Delete();
                                }
                            }
                            catch { }
                        }
                    }
                    catch (Exception exDel)
                    {
                        System.Diagnostics.Debug.WriteLine($"Delete sheet error: {exDel.Message}");
                    }
                }

                // Kích hoạt lại Workbook ban đầu
                try { targetWb.Activate(); } catch { }

                _isBatchProcessing = false;
                try { app.EnableEvents = true; } catch { }
                try { app.DisplayAlerts = true; } catch { }
                try { app.ScreenUpdating = true; } catch { }

                RefreshWorkbookTree();

                string msg = keepOriginalSheets
                    ? $"✅ Đã tách và lưu thành công {exportedCount} file Excel (.xlsx) vào thư mục:\n{outputFolder}"
                    : $"✅ Đã tách thành công {exportedCount} file Excel (.xlsx) và xóa {exportedSheetNames.Count} sheet tương ứng khỏi file hiện tại!\nThư mục lưu: {outputFolder}";

                WpfMessageBox.Show(msg, "Tách Sheet Thành Công", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
                return true;
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Lỗi khi tách sheet:\n{ex.Message}", "Tách Sheet",
                                   System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                return false;
            }
            finally
            {
                _isBatchProcessing = false;
                if (_excelApp != null)
                {
                    try { _excelApp.EnableEvents = true; } catch { }
                    try { _excelApp.DisplayAlerts = true; } catch { }
                    try { _excelApp.ScreenUpdating = true; } catch { }
                }
            }
        }

        public bool ConsolidateSheetsData(string wbName, List<string>? sheetNames, bool hasHeaderRow = true)
        {
            if (_excelApp == null) return false;

            _isBatchProcessing = true;
            dynamic app = _excelApp;

            try
            {
                try { app.EnableEvents = false; } catch { }
                try { app.DisplayAlerts = false; } catch { }
                try { app.ScreenUpdating = false; } catch { }

                dynamic? targetWb = null;
                if (!string.IsNullOrEmpty(wbName))
                {
                    try { targetWb = app.Workbooks[wbName]; } catch { }
                }
                if (targetWb == null)
                {
                    try { targetWb = app.ActiveWorkbook; } catch { }
                }

                if (targetWb == null) return false;

                string resultSheetName = "Tong_Hop";
                dynamic? wsSummary = null;

                try { wsSummary = targetWb.Worksheets[resultSheetName]; } catch { }

                if (wsSummary == null)
                {
                    // Thêm Sheet Tong_Hop vào vị trí cuối cùng
                    dynamic lastSheet = targetWb.Sheets[targetWb.Sheets.Count];
                    wsSummary = targetWb.Worksheets.Add(After: lastSheet);
                    wsSummary.Name = resultSheetName;
                }
                else
                {
                    wsSummary.Cells.Clear();
                }

                int totalSheets = targetWb.Sheets.Count;
                int currentDestRow = 1;
                bool isFirstSheet = true;
                int sheetsMerged = 0;

                for (int i = 1; i <= totalSheets; i++)
                {
                    try
                    {
                        dynamic ws = targetWb.Sheets[i];
                        string sheetName = ws.Name;

                        if (sheetName == resultSheetName || (sheetNames != null && sheetNames.Count > 0 && !sheetNames.Contains(sheetName)))
                        {
                            continue;
                        }

                        dynamic usedRange = ws.UsedRange;
                        int firstRow = usedRange.Row;
                        int rowsCount = usedRange.Rows.Count;
                        int firstCol = usedRange.Column;
                        int colsCount = usedRange.Columns.Count;

                        if (rowsCount <= 0 || colsCount <= 0) continue;

                        int lastRow = firstRow + rowsCount - 1;
                        int lastCol = firstCol + colsCount - 1;

                        int startRow = firstRow;
                        if (!isFirstSheet && hasHeaderRow)
                        {
                            if (rowsCount > 1)
                            {
                                startRow = firstRow + 1; // Bỏ qua dòng tiêu đề nếu sheet có nhiều hơn 1 dòng
                            }
                            else
                            {
                                startRow = firstRow; // Nếu sheet chỉ có 1 dòng thì vẫn lấy dòng đó
                            }
                        }

                        if (startRow <= lastRow)
                        {
                            string startColLetter = GetExcelColumnLetter(firstCol);
                            string endColLetter = GetExcelColumnLetter(lastCol);
                            string rangeAddress = $"{startColLetter}{startRow}:{endColLetter}{lastRow}";

                            dynamic sourceRange = ws.Range[rangeAddress];
                            dynamic destCell = wsSummary.Range[$"A{currentDestRow}"];

                            sourceRange.Copy(destCell);
                            try { app.CutCopyMode = false; } catch { }

                            int rowsCopied = (lastRow - startRow + 1);
                            currentDestRow += rowsCopied;
                            sheetsMerged++;
                            isFirstSheet = false;
                        }
                    }
                    catch (Exception exSheet)
                    {
                        System.Diagnostics.Debug.WriteLine($"Consolidate sheet error: {exSheet.Message}");
                    }
                }

                // Kích hoạt Sheet Tổng Hợp
                try { wsSummary.Activate(); } catch { }
                try { wsSummary.Range["A:Z"].EntireColumn.AutoFit(); } catch { }

                _isBatchProcessing = false;
                try { app.EnableEvents = true; } catch { }
                try { app.DisplayAlerts = true; } catch { }
                try { app.ScreenUpdating = true; } catch { }

                RefreshWorkbookTree();

                WpfMessageBox.Show($"✅ Đã gộp thành công dữ liệu từ {sheetsMerged} sheet vào sheet [{resultSheetName}] ({currentDestRow - 1} dòng dữ liệu)!",
                                   "Gộp Sheet Thành Công", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
                return true;
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Lỗi gộp dữ liệu các sheet:\n{ex.Message}", "Gộp Sheet",
                                   System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                return false;
            }
            finally
            {
                _isBatchProcessing = false;
                if (_excelApp != null)
                {
                    try { _excelApp.EnableEvents = true; } catch { }
                    try { _excelApp.DisplayAlerts = true; } catch { }
                    try { _excelApp.ScreenUpdating = true; } catch { }
                }
            }
        }

        private static string GetExcelColumnLetter(int colIndex)
        {
            int div = colIndex;
            string colLetter = string.Empty;
            while (div > 0)
            {
                int mod = (div - 1) % 26;
                colLetter = (char)(65 + mod) + colLetter;
                div = (div - mod) / 26;
            }
            return string.IsNullOrEmpty(colLetter) ? "A" : colLetter;
        }

        public bool ImportSheetsFromExternalFiles(string targetWbName, string[] filePaths)
        {
            if (_excelApp == null || filePaths == null || filePaths.Length == 0) return false;

            _isBatchProcessing = true;
            dynamic app = _excelApp;

            try
            {
                try { app.EnableEvents = false; } catch { }
                try { app.DisplayAlerts = false; } catch { }
                try { app.ScreenUpdating = false; } catch { }

                dynamic? targetWb = null;
                if (!string.IsNullOrEmpty(targetWbName))
                {
                    try { targetWb = app.Workbooks[targetWbName]; } catch { }
                }
                if (targetWb == null)
                {
                    try { targetWb = app.ActiveWorkbook; } catch { }
                }

                if (targetWb == null) return false;

                int importedCount = 0;

                foreach (var path in filePaths)
                {
                    if (!System.IO.File.Exists(path)) continue;

                    try
                    {
                        dynamic sourceWb = app.Workbooks.Open(path, ReadOnly: true);
                        if (sourceWb != null)
                        {
                            int sourceSheetCount = sourceWb.Sheets.Count;
                            for (int i = 1; i <= sourceSheetCount; i++)
                            {
                                dynamic ws = sourceWb.Sheets[i];
                                dynamic lastSheet = targetWb.Sheets[targetWb.Sheets.Count];
                                ws.Copy(After: lastSheet);
                                importedCount++;
                            }
                            sourceWb.Close(false);
                        }
                    }
                    catch (Exception exFile)
                    {
                        System.Diagnostics.Debug.WriteLine($"Import file error: {exFile.Message}");
                    }
                }

                _isBatchProcessing = false;
                try { app.EnableEvents = true; } catch { }
                try { app.DisplayAlerts = true; } catch { }
                try { app.ScreenUpdating = true; } catch { }

                RefreshWorkbookTree();
                WpfMessageBox.Show($"✅ Đã nhập thành công {importedCount} sheet vào [{targetWb.Name}]!",
                                   "Nhập File Thành Công", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
                return true;
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Lỗi nhập sheet từ file:\n{ex.Message}", "Nhập File",
                                   System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                return false;
            }
            finally
            {
                _isBatchProcessing = false;
                if (_excelApp != null)
                {
                    try { _excelApp.EnableEvents = true; } catch { }
                    try { _excelApp.ScreenUpdating = true; } catch { }
                    try { _excelApp.DisplayAlerts = true; } catch { }
                }
            }
        }

        #endregion
    }
}
