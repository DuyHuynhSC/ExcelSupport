using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Office.Interop.Excel;
using ExcelSupport.Models;
using ExcelSupport.Services;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using WpfMessageBox = System.Windows.MessageBox;

namespace ExcelSupport
{
    public partial class AddInEvents
    {
        #region Excel COM Helper Methods for AI Assistant & Translation

        public class ActiveCellInfo
        {
            public string WorkbookName { get; set; } = string.Empty;
            public string SheetName { get; set; } = string.Empty;
            public string CellAddress { get; set; } = string.Empty;
            public string Formula { get; set; } = string.Empty;
            public string Value { get; set; } = string.Empty;
            public string ErrorText { get; set; } = string.Empty;
            public bool HasError => !string.IsNullOrEmpty(ErrorText);
        }

        public ActiveCellInfo? GetActiveCellInfo()
        {
            if (_excelApp == null) return null;

            try
            {
                Range? cell = null;
                _Worksheet? ws = null;
                Workbook? wb = null;
                try
                {
                    cell = _excelApp.ActiveCell;
                    if (cell == null) return null;

                    ws = cell.Worksheet;
                    wb = ws.Parent as Workbook;

                    var info = new ActiveCellInfo
                    {
                        WorkbookName = wb?.Name ?? string.Empty,
                        SheetName = ws.Name,
                        CellAddress = cell.Address[false, false],
                        Formula = cell.Formula?.ToString() ?? string.Empty,
                        Value = cell.Text?.ToString() ?? string.Empty
                    };

                    // Kiểm tra nếu ô đang chứa mã lỗi Excel (#N/A, #VALUE!, #REF!, #DIV/0!, ...)
                    string valStr = info.Value.Trim();
                    if (valStr.StartsWith("#") && (valStr.EndsWith("!") || valStr.EndsWith("?")))
                    {
                        info.ErrorText = valStr;
                    }

                    return info;
                }
                finally
                {
                    if (cell != null) Marshal.ReleaseComObject(cell);
                    if (ws != null) Marshal.ReleaseComObject(ws);
                    if (wb != null) Marshal.ReleaseComObject(wb);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"GetActiveCellInfo error: {ex.Message}");
                return null;
            }
        }

        public bool InsertFormulaToActiveCell(string formula)
        {
            if (_excelApp == null || string.IsNullOrWhiteSpace(formula)) return false;

            try
            {
                string cleanFormula = formula.Trim();
                if (!cleanFormula.StartsWith("=") && !cleanFormula.StartsWith("@"))
                {
                    cleanFormula = "=" + cleanFormula;
                }

                Range? cell = null;
                try
                {
                    cell = _excelApp.ActiveCell;
                    if (cell != null)
                    {
                        cell.Formula = cleanFormula;
                        return true;
                    }
                }
                finally
                {
                    if (cell != null) Marshal.ReleaseComObject(cell);
                }
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Không thể chèn công thức vào ô Excel:\n{ex.Message}", "Trợ Lý AI",
                                   System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
            }
            return false;
        }

        public string GetActiveSheetContextSummary()
        {
            if (_excelApp == null) return string.Empty;

            _Worksheet? ws = null;
            Range? usedRange = null;
            try
            {
                ws = _excelApp.ActiveSheet as _Worksheet;
                if (ws == null) return string.Empty;

                var sb = new StringBuilder();
                sb.AppendLine($"Tên Sheet: {ws.Name}");

                usedRange = ws.UsedRange;
                if (usedRange != null && usedRange.Rows.Count > 0 && usedRange.Columns.Count > 0)
                {
                    int startCol = usedRange.Column;
                    int totalCols = Math.Min(usedRange.Columns.Count, 30);
                    int numRows = Math.Min(usedRange.Rows.Count, 4);

                    object? rawVal = usedRange.Value2;
                    if (rawVal is object[,] allVals)
                    {
                        sb.AppendLine("Cấu trúc các cột dữ liệu trên Sheet:");
                        for (int c = 1; c <= totalCols; c++)
                        {
                            int colIndex = startCol + c - 1;
                            string colLetter = ConvertColIndexToLetter(colIndex);
                            string header = allVals[1, c]?.ToString()?.Trim() ?? string.Empty;

                            var samples = new List<string>();
                            for (int r = 2; r <= numRows; r++)
                            {
                                string sample = allVals[r, c]?.ToString()?.Trim() ?? string.Empty;
                                if (!string.IsNullOrEmpty(sample) && samples.Count < 2) samples.Add(sample);
                            }

                            string sampleStr = samples.Count > 0 ? $" (Giá trị mẫu: {string.Join(", ", samples)})" : "";
                            sb.AppendLine($"• Cột {colLetter}: {(!string.IsNullOrEmpty(header) ? header : "(Không tên)")}{sampleStr}");
                        }
                    }
                }

                Range? activeCell = _excelApp.ActiveCell;
                if (activeCell != null)
                {
                    sb.AppendLine($"Tọa độ ô đang chọn: {activeCell.Address[false, false]}");
                    Marshal.ReleaseComObject(activeCell);
                }

                return sb.ToString();
            }
            catch
            {
                return string.Empty;
            }
            finally
            {
                if (usedRange != null) Marshal.ReleaseComObject(usedRange);
                if (ws != null) Marshal.ReleaseComObject(ws);
            }
        }

        private static string ConvertColIndexToLetter(int colIndex)
        {
            string colLetter = string.Empty;
            while (colIndex > 0)
            {
                int modulo = (colIndex - 1) % 26;
                colLetter = Convert.ToChar('A' + modulo) + colLetter;
                colIndex = (colIndex - modulo) / 26;
            }
            return colLetter;
        }

        public List<CellTextItem> GetSelectedCellsText(int maxCells = 500)
        {
            var list = new List<CellTextItem>();
            if (_excelApp == null) return list;

            try
            {
                Range? selection = null;
                try
                {
                    selection = _excelApp.Selection as Range;
                    if (selection == null) return list;

                    int count = 0;
                    foreach (Range cell in selection.Cells)
                    {
                        try
                        {
                            string text = cell.Text?.ToString() ?? string.Empty;
                            if (!string.IsNullOrWhiteSpace(text))
                            {
                                list.Add(new CellTextItem
                                {
                                    Row = cell.Row,
                                    Column = cell.Column,
                                    Address = cell.Address[false, false],
                                    OriginalText = text.Trim()
                                });

                                count++;
                                if (count >= maxCells) break;
                            }
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(cell);
                        }
                    }
                }
                finally
                {
                    if (selection != null) Marshal.ReleaseComObject(selection);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"GetSelectedCellsText error: {ex.Message}");
            }

            return list;
        }

        public bool WriteTranslatedCells(List<CellTextItem> items, bool writeToAdjacentColumn)
        {
            if (_excelApp == null || items == null || items.Count == 0) return false;

            try
            {
                _Worksheet? ws = null;
                try
                {
                    ws = _excelApp.ActiveSheet as _Worksheet;
                    if (ws == null) return false;

                    _excelApp.ScreenUpdating = false;

                    var backupList = new List<TranslationUndoHelper.CellBackupItem>();

                    foreach (var item in items)
                    {
                        if (string.IsNullOrEmpty(item.TranslatedText)) continue;

                        int targetCol = writeToAdjacentColumn ? (item.Column + 1) : item.Column;
                        Range? cell = null;
                        try
                        {
                            cell = ws.Cells[item.Row, targetCol] as Range;
                            if (cell != null)
                            {
                                object? oldVal = cell.Value2;
                                cell.Value2 = item.TranslatedText;

                                backupList.Add(new TranslationUndoHelper.CellBackupItem
                                {
                                    Row = item.Row,
                                    Column = targetCol,
                                    OldValue = oldVal,
                                    NewValue = item.TranslatedText
                                });
                            }
                        }
                        finally
                        {
                            if (cell != null) Marshal.ReleaseComObject(cell);
                        }
                    }

                    if (backupList.Count > 0)
                    {
                        TranslationUndoHelper.RecordAndApply(ws, backupList, "Dịch Thuật AI");
                    }

                    return true;
                }
                finally
                {
                    _excelApp.ScreenUpdating = true;
                    if (ws != null) Marshal.ReleaseComObject(ws);
                }
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Lỗi ghi dữ liệu dịch vào bảng tính:\n{ex.Message}", "Dịch Thuật AI",
                                   System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                return false;
            }
        }

        public bool CreateTableOfContents(string? workbookName, bool addBackLinkToSheets = true)
        {
            if (_excelApp == null) return false;

            try
            {
                dynamic app = _excelApp;
                dynamic? targetWb = null;
                if (!string.IsNullOrEmpty(workbookName))
                {
                    try { targetWb = app.Workbooks[workbookName]; } catch { }
                }
                if (targetWb == null)
                {
                    try { targetWb = app.ActiveWorkbook; } catch { }
                }

                if (targetWb == null)
                {
                    WpfMessageBox.Show("Không tìm thấy Workbook đang mở.", "Tạo Mục Lục",
                                       System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                    return false;
                }

                try { app.ScreenUpdating = false; } catch { }

                string indexSheetName = "Mục Lục";
                dynamic? wsIndex = null;

                // Tìm xem đã có sheet Mục Lục chưa
                try
                {
                    wsIndex = targetWb.Worksheets[indexSheetName];
                }
                catch { }

                if (wsIndex == null)
                {
                    try
                    {
                        dynamic firstSheet = targetWb.Sheets[1];
                        wsIndex = targetWb.Worksheets.Add(firstSheet);
                    }
                    catch
                    {
                        wsIndex = targetWb.Worksheets.Add();
                    }

                    if (wsIndex != null)
                    {
                        wsIndex.Name = indexSheetName;
                    }
                }
                else
                {
                    // Di chuyển lên đầu tiên
                    try
                    {
                        dynamic firstSheet = targetWb.Sheets[1];
                        wsIndex.Move(firstSheet);
                    }
                    catch { }

                    try
                    {
                        if ((int)wsIndex.Visible != (int)XlSheetVisibility.xlSheetVisible)
                        {
                            wsIndex.Visible = (int)XlSheetVisibility.xlSheetVisible;
                        }
                    }
                    catch { }

                    try
                    {
                        wsIndex.Cells.Clear();
                    }
                    catch { }
                }

                if (wsIndex == null) return false;

                // 1. Tiêu đề lớn
                dynamic titleRange = wsIndex.Range["A1:E1"];
                titleRange.Merge();
                titleRange.Value2 = $"📋 BẢNG MỤC LỤC CÁC SHEET - {targetWb.Name}";
                titleRange.Font.Size = 14;
                titleRange.Font.Bold = true;
                titleRange.Font.Name = "Segoe UI";
                titleRange.Font.Color = ColorTranslator.ToOle(Color.White);
                titleRange.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(16, 124, 65)); // Office Dark Green
                titleRange.HorizontalAlignment = -4108; // xlCenter
                titleRange.VerticalAlignment = -4108; // xlCenter
                titleRange.RowHeight = 36;

                // 2. Tiêu đề các cột
                wsIndex.Cells[3, 1] = "STT";
                wsIndex.Cells[3, 2] = "Tên Sheet (Click để mở)";
                wsIndex.Cells[3, 3] = "Trạng Thái";
                wsIndex.Cells[3, 4] = "Màu Tab";
                wsIndex.Cells[3, 5] = "Ghi Chú";

                dynamic headerRange = wsIndex.Range["A3:E3"];
                headerRange.Font.Bold = true;
                headerRange.Font.Size = 11;
                headerRange.Font.Name = "Segoe UI";
                headerRange.Font.Color = ColorTranslator.ToOle(Color.White);
                headerRange.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(30, 41, 59)); // Slate 800
                headerRange.HorizontalAlignment = -4108; // xlCenter
                headerRange.VerticalAlignment = -4108; // xlCenter
                headerRange.RowHeight = 26;

                int currentRow = 4;
                int stt = 1;
                int sheetCount = targetWb.Sheets.Count;

                for (int i = 1; i <= sheetCount; i++)
                {
                    try
                    {
                        dynamic ws = targetWb.Sheets[i];
                        string wsName = ws.Name;
                        if (string.Equals(wsName, indexSheetName, StringComparison.OrdinalIgnoreCase))
                        {
                            continue; // Bỏ qua chính sheet Mục Lục
                        }

                        // Cột 1: STT
                        dynamic sttCell = wsIndex.Cells[currentRow, 1];
                        sttCell.Value2 = stt;
                        sttCell.HorizontalAlignment = -4108; // xlCenter

                        // Cột 2: Tên Sheet + Hyperlink
                        dynamic linkCell = wsIndex.Cells[currentRow, 2];
                        wsIndex.Hyperlinks.Add(
                            linkCell,
                            "",
                            $"'{wsName}'!A1",
                            $"Chuyển đến sheet [{wsName}]",
                            wsName
                        );
                        linkCell.Font.Name = "Segoe UI";
                        linkCell.Font.Size = 11;

                        // Cột 3: Trạng thái Ẩn/Hiện
                        dynamic statusCell = wsIndex.Cells[currentRow, 3];
                        int vis = (int)ws.Visible;
                        if (vis == (int)XlSheetVisibility.xlSheetVisible)
                        {
                            statusCell.Value2 = "Hiển thị";
                            statusCell.Font.Color = ColorTranslator.ToOle(Color.FromArgb(22, 101, 52));
                        }
                        else if (vis == (int)XlSheetVisibility.xlSheetHidden)
                        {
                            statusCell.Value2 = "Bị ẩn (Hidden)";
                            statusCell.Font.Color = ColorTranslator.ToOle(Color.FromArgb(180, 83, 9));
                        }
                        else
                        {
                            statusCell.Value2 = "Ẩn sâu (Very Hidden)";
                            statusCell.Font.Color = ColorTranslator.ToOle(Color.FromArgb(220, 38, 38));
                        }
                        statusCell.HorizontalAlignment = -4108;

                        // Cột 4: Màu Tab
                        dynamic colorCell = wsIndex.Cells[currentRow, 4];
                        try
                        {
                            object rawColor = ws.Tab.Color;
                            int colorIndex = -4142;
                            try { colorIndex = (int)ws.Tab.ColorIndex; } catch { }

                            if (colorIndex != -4142 && rawColor != null && !(rawColor is bool))
                            {
                                colorCell.Interior.Color = Convert.ToInt32(rawColor);
                                colorCell.Value2 = "   ";
                            }
                            else
                            {
                                colorCell.Value2 = "(Mặc định)";
                                colorCell.Font.Color = ColorTranslator.ToOle(Color.FromArgb(148, 163, 184));
                                colorCell.HorizontalAlignment = -4108;
                            }
                        }
                        catch
                        {
                            colorCell.Value2 = "-";
                            colorCell.HorizontalAlignment = -4108;
                        }

                        // Cột 5: Ghi chú
                        dynamic noteCell = wsIndex.Cells[currentRow, 5];
                        noteCell.Value2 = "";

                        stt++;
                        currentRow++;
                    }
                    catch (Exception exSheet)
                    {
                        System.Diagnostics.Debug.WriteLine($"Sheet index error: {exSheet.Message}");
                    }
                }

                // Định dạng toàn bộ bảng dữ liệu
                if (currentRow > 4)
                {
                    dynamic dataTable = wsIndex.Range[$"A3:E{currentRow - 1}"];
                    dataTable.Borders.LineStyle = 1; // xlContinuous
                    dataTable.Borders.Color = ColorTranslator.ToOle(Color.FromArgb(203, 213, 225));
                    dataTable.Font.Name = "Segoe UI";
                }

                // Tự động căn chỉnh độ rộng cột
                try
                {
                    wsIndex.Range["A:E"].EntireColumn.AutoFit();
                    wsIndex.Range["B:B"].ColumnWidth = Math.Max(25.0, Convert.ToDouble(wsIndex.Range["B:B"].ColumnWidth) + 5.0);
                    wsIndex.Range["E:E"].ColumnWidth = 20.0;
                }
                catch { }

                // Kích hoạt Sheet Mục Lục
                try
                {
                    wsIndex.Activate();
                }
                catch { }

                // Đổi màu Tab cho Sheet Mục Lục thành xanh lá nổi bật
                try
                {
                    wsIndex.Tab.Color = ColorTranslator.ToOle(Color.FromArgb(16, 124, 65));
                }
                catch { }

                RefreshWorkbookTree();

                WpfMessageBox.Show($"✅ Đã tạo thành công Bảng Mục Lục cho {stt - 1} sheet trong [{targetWb.Name}]!",
                                   "Tạo Mục Lục Sheet", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
                return true;
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Lỗi tạo bảng mục lục:\n{ex.Message}", "Tạo Mục Lục",
                                   System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                return false;
            }
            finally
            {
                if (_excelApp != null)
                {
                    try { _excelApp.ScreenUpdating = true; } catch { }
                }
            }
        }

        #endregion
    }
}
