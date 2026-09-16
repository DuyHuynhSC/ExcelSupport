using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelSupport.Host;
using ExcelSupport.Models;
using ExcelSupport.Views;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using WpfMessageBox = System.Windows.MessageBox;
using MessageBoxButton = System.Windows.MessageBoxButton;
using MessageBoxImage = System.Windows.MessageBoxImage;
using MessageBoxResult = System.Windows.MessageBoxResult;

namespace ExcelSupport.Services
{
    public static class ProjectDocumentLauncherService
    {
        // Regex nhận diện version: v1.0, ver2.1, rev03, _20260915, v1, etc.
        private static readonly Regex VersionRegex = new Regex(
            @"(?:[_\-\s]|^)(?:v|ver|version|rev|revision)?\.?\s*(\d+(?:\.\d+)*(?:[a-zA-Z\-_]\w*)?|20\d{6}(?:_\d+)?)",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex DateVersionRegex = new Regex(
            @"(?:[_\-\s])(20\d{2}[-_]?\d{2}[-_]?\d{2})",
            RegexOptions.Compiled);

        public static MessageBoxResult ShowExcelMessageBox(
            string text,
            string caption,
            MessageBoxButton buttons = MessageBoxButton.OK,
            MessageBoxImage icon = MessageBoxImage.Information)
        {
            try
            {
                return WpfMessageBox.Show(text, caption, buttons, icon);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ShowExcelMessageBox] WpfMessageBox error: {ex.Message}");
                try
                {
                    var formsButtons = buttons switch
                    {
                        MessageBoxButton.YesNo => System.Windows.Forms.MessageBoxButtons.YesNo,
                        MessageBoxButton.YesNoCancel => System.Windows.Forms.MessageBoxButtons.YesNoCancel,
                        _ => System.Windows.Forms.MessageBoxButtons.OK
                    };
                    var formsIcon = icon switch
                    {
                        MessageBoxImage.Error => System.Windows.Forms.MessageBoxIcon.Error,
                        MessageBoxImage.Warning => System.Windows.Forms.MessageBoxIcon.Warning,
                        MessageBoxImage.Question => System.Windows.Forms.MessageBoxIcon.Question,
                        _ => System.Windows.Forms.MessageBoxIcon.Information
                    };
                    var dr = System.Windows.Forms.MessageBox.Show(text, caption, formsButtons, formsIcon);
                    return dr == System.Windows.Forms.DialogResult.Yes ? MessageBoxResult.Yes :
                           dr == System.Windows.Forms.DialogResult.No ? MessageBoxResult.No : MessageBoxResult.OK;
                }
                catch (Exception exForms)
                {
                    Debug.WriteLine($"[ShowExcelMessageBox] FormsMessageBox error: {exForms.Message}");
                    return MessageBoxResult.None;
                }
            }
        }

        /// <summary>
        /// <summary>
        /// Trích xuất danh sách tất cả các từ khóa không trùng lặp từ vùng ô đang chọn (Selection)
        /// </summary>
        public static List<string> ExtractKeywordsFromSelection(ExcelApp? app)
        {
            var keywords = new List<string>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            if (app == null) return keywords;

            void AddText(object? val)
            {
                if (val == null) return;
                string raw = val.ToString()?.Trim() ?? string.Empty;
                if (string.IsNullOrWhiteSpace(raw)) return;

                var lines = raw.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var line in lines)
                {
                    string trimmed = line.Trim();
                    if (!string.IsNullOrWhiteSpace(trimmed) && seen.Add(trimmed))
                    {
                        keywords.Add(trimmed);
                    }
                }
            }

            try
            {
                var selection = app.Selection as Microsoft.Office.Interop.Excel.Range;
                if (selection != null)
                {
                    int totalCells = 0;
                    foreach (Microsoft.Office.Interop.Excel.Range area in selection.Areas)
                    {
                        int rowCount = 1;
                        int colCount = 1;
                        try { rowCount = area.Rows.Count; } catch { }
                        try { colCount = area.Columns.Count; } catch { }

                        object? rawValues = null;
                        try { rawValues = area.Value2; } catch { }
                        if (rawValues == null)
                        {
                            try { rawValues = area.Value; } catch { }
                        }

                        if (rawValues is object[,] valArray)
                        {
                            for (int r = 1; r <= rowCount && totalCells < 500; r++)
                            {
                                for (int c = 1; c <= colCount && totalCells < 500; c++)
                                {
                                    AddText(valArray[r, c]);
                                    totalCells++;
                                }
                            }
                        }
                        else if (rawValues != null)
                        {
                            AddText(rawValues);
                            totalCells++;
                        }

                        if (totalCells >= 500) break;
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ExtractKeywordsFromSelection] Error: {ex.Message}");
            }

            // Fallback nếu Selection không trích xuất được từ khóa nào
            if (keywords.Count == 0)
            {
                try
                {
                    if (app.ActiveCell != null)
                    {
                        object? cellVal = app.ActiveCell.Value2 ?? app.ActiveCell.Value;
                        AddText(cellVal);
                    }
                }
                catch { }
            }

            return keywords;
        }

        /// <summary>
        /// Kích hoạt tìm kiếm và mở tài liệu từ ô hoặc vùng ô đang chọn trên Excel (chuẩn hóa trả về LaunchResult như Special Copy)
        /// </summary>
        public static LaunchResult LaunchFromSelection(ExcelApp? app, SpecDocumentType docType, bool? forceReadOnly = null)
        {
            try
            {
                if (app == null)
                {
                    app = AddInEvents.Instance?.ExcelAppInstance ?? (ExcelApp)ExcelDna.Integration.ExcelDnaUtil.Application;
                }

                var profile = ProjectProfileManager.GetActiveProfile();
                if (profile == null)
                {
                    return LaunchResult.Fail(
                        LocalizationService.Get("SpecLauncher_NoProfileMsg"));
                }

                // Kiểm tra thư mục tương ứng
                if (docType == SpecDocumentType.DetailedDesign || docType == SpecDocumentType.BasicDesign)
                {
                    string folderVi = profile.GetTargetFolder(docType, "vi");
                    string folderJa = profile.GetTargetFolder(docType, "ja");
                    bool hasVi = !string.IsNullOrWhiteSpace(folderVi) && Directory.Exists(folderVi);
                    bool hasJa = !string.IsNullOrWhiteSpace(folderJa) && Directory.Exists(folderJa);

                    if (!hasVi && !hasJa)
                    {
                        string docTypeName = GetDocTypeName(docType);
                        string pathDesc = $"VN: {folderVi}\nJP: {folderJa}";
                        return LaunchResult.Fail(
                            LocalizationService.Get("SpecLauncher_FolderNotFoundMsg", docTypeName, profile.Name, pathDesc));
                    }
                }
                else
                {
                    string targetFolder = profile.GetTargetFolder(docType);
                    if (string.IsNullOrWhiteSpace(targetFolder) || !Directory.Exists(targetFolder))
                    {
                        string docTypeName = GetDocTypeName(docType);
                        return LaunchResult.Fail(
                            LocalizationService.Get("SpecLauncher_FolderNotFoundMsg", docTypeName, profile.Name, targetFolder));
                    }
                }

                // Lấy danh sách từ khóa từ vùng chọn Excel
                var keywords = ExtractKeywordsFromSelection(app);
                bool isReadOnly = forceReadOnly ?? profile.OpenReadOnlyDefault;

                // Nếu không có từ khóa nào (người dùng chọn ô trống)
                if (keywords.Count == 0)
                {
                    return LaunchResult.Fail(
                        LocalizationService.Get("SpecLauncher_EmptySelectionMsg"));
                }

                // Tìm kiếm file
                var results = SearchFiles(profile, docType, keywords);

                if (results.Count == 0)
                {
                    string docTypeName = GetDocTypeName(docType);
                    string kwDisplay = string.Join(", ", keywords.Take(5));
                    if (keywords.Count > 5) kwDisplay += $" (+{keywords.Count - 5})";

                    string folderDesc = (docType == SpecDocumentType.TestSpec)
                        ? profile.GetTargetFolder(docType)
                        : $"{profile.GetTargetFolder(docType, "vi")} | {profile.GetTargetFolder(docType, "ja")}";

                    return LaunchResult.Fail(
                        LocalizationService.Get("SpecLauncher_NoFilesFoundMsg", docTypeName, kwDisplay, folderDesc));
                }

                // 1. Đối với Chỉ Thị Test (TestSpec) -> Mở trực tiếp toàn bộ các file tìm thấy ngay lập tức mà không hiển thị form kết quả
                if (docType == SpecDocumentType.TestSpec)
                {
                    var filePaths = results.Select(r => r.FilePath).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
                    OpenMultipleFilesAsync(filePaths, isReadOnly, app);
                    return LaunchResult.Ok();
                }

                // 2. Nếu là TKCT / TKCB VÀ chỉ chọn 1 ô duy nhất VÀ chỉ tìm thấy đúng 1 file -> Mở trực tiếp ngay lập tức!
                if (keywords.Count == 1 && results.Count == 1)
                {
                    OpenMultipleFilesAsync(new[] { results[0].FilePath }, isReadOnly, app);
                    return LaunchResult.Ok();
                }

                // 3. Với TKCT / TKCB khi có nhiều file hoặc chọn vùng nhiều ô:
                // Mở hộp thoại hiển thị 2 Tab (Tiếng Việt & Tiếng Nhật) để người dùng xem và chọn!
                string keywordSummary = string.Join(", ", keywords.Take(4));
                if (keywords.Count > 4) keywordSummary += $" (+{keywords.Count - 4})";

                SpecVersionSelectorDialog.ShowWindow(
                    profile, 
                    docType, 
                    keywordSummary, 
                    results, 
                    isReadOnly, 
                    AddInEvents.MainViewModel?.IsDarkTheme ?? false,
                    keywords);

                return LaunchResult.Ok();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[LaunchFromSelection] Error: {ex}");
                return LaunchResult.Fail(
                    LocalizationService.Get("SpecLauncher_GeneralError", ex.Message));
            }
        }

        /// <summary>
        /// Tìm kiếm đệ quy toàn bộ file trong thư mục chỉ định của profile khớp với 1 từ khóa
        /// </summary>
        public static List<SpecSearchResultItem> SearchFiles(ProjectProfile profile, SpecDocumentType docType, string keyword)
        {
            var kwList = new List<string>();
            if (!string.IsNullOrWhiteSpace(keyword))
            {
                kwList.Add(keyword.Trim());
            }
            return SearchFiles(profile, docType, kwList);
        }

        /// <summary>
        /// Tìm kiếm đệ quy toàn bộ file trong thư mục chỉ định của profile khớp với danh sách các từ khóa (vùng chọn)
        /// </summary>
        public static List<SpecSearchResultItem> SearchFiles(ProjectProfile profile, SpecDocumentType docType, IEnumerable<string> keywords)
        {
            var list = new List<SpecSearchResultItem>();

            // Chuẩn bị danh sách định dạng hỗ trợ
            var allowedExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (!string.IsNullOrWhiteSpace(profile.FileExtensions))
            {
                var parts = profile.FileExtensions.Split(new[] { ';', ',', '|' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var part in parts)
                {
                    string ext = part.Trim();
                    if (!ext.StartsWith(".")) ext = "." + ext;
                    allowedExtensions.Add(ext);
                }
            }

            if (allowedExtensions.Count == 0)
            {
                allowedExtensions.Add(".xlsx");
                allowedExtensions.Add(".xlsm");
                allowedExtensions.Add(".xls");
                allowedExtensions.Add(".docx");
                allowedExtensions.Add(".pdf");
                allowedExtensions.Add(".pptx");
            }

            // Chuẩn hóa danh sách từ khóa tìm kiếm
            var kwList = keywords?
                .Where(k => !string.IsNullOrWhiteSpace(k))
                .Select(k => k.Trim())
                .ToList() ?? new List<string>();

            var kwPairs = kwList.Select(k => new KeywordPair
            {
                Raw = k,
                NoSep = k.Replace("_", "").Replace("-", "").Replace(" ", "")
            }).ToList();

            bool searchSubfolders = !profile.DoNotSearchSubfolders;

            if (docType == SpecDocumentType.DetailedDesign || docType == SpecDocumentType.BasicDesign)
            {
                string folderVi = profile.GetTargetFolder(docType, "vi");
                if (!string.IsNullOrWhiteSpace(folderVi) && Directory.Exists(folderVi))
                {
                    list.AddRange(ScanFolder(folderVi, profile, docType, "vi", kwPairs, allowedExtensions, searchSubfolders));
                }

                string folderJa = profile.GetTargetFolder(docType, "ja");
                if (!string.IsNullOrWhiteSpace(folderJa) && Directory.Exists(folderJa))
                {
                    list.AddRange(ScanFolder(folderJa, profile, docType, "ja", kwPairs, allowedExtensions, searchSubfolders));
                }
            }
            else
            {
                string targetFolder = profile.GetTargetFolder(docType);
                if (!string.IsNullOrWhiteSpace(targetFolder) && Directory.Exists(targetFolder))
                {
                    list.AddRange(ScanFolder(targetFolder, profile, docType, "vi", kwPairs, allowedExtensions, searchSubfolders));
                }
            }

            return list;
        }

        private static List<SpecSearchResultItem> ScanFolder(
            string targetFolder,
            ProjectProfile profile,
            SpecDocumentType docType,
            string language,
            List<KeywordPair> kwPairs,
            HashSet<string> allowedExtensions,
            bool searchSubfolders)
        {
            var list = new List<SpecSearchResultItem>();
            var addedPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            var allFiles = SafeEnumerateFiles(targetFolder, searchSubfolders);

            foreach (var filePath in allFiles)
            {
                string ext = Path.GetExtension(filePath);
                if (!allowedExtensions.Contains(ext)) continue;
                if (addedPaths.Contains(filePath)) continue;

                string fileName = Path.GetFileName(filePath);
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(filePath);

                bool isMatch = false;

                if (kwPairs.Count == 0)
                {
                    isMatch = true;
                }
                else
                {
                    string fileNameNoSep = fileNameWithoutExt.Replace("_", "").Replace("-", "").Replace(" ", "");

                    foreach (var kw in kwPairs)
                    {
                        if (fileNameWithoutExt.IndexOf(kw.Raw, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            isMatch = true;
                            break;
                        }

                        if (kw.NoSep.Length > 2 && fileNameNoSep.IndexOf(kw.NoSep, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            isMatch = true;
                            break;
                        }
                    }
                }

                if (isMatch)
                {
                    try
                    {
                        var fi = new FileInfo(filePath);
                        string relativeDir = string.Empty;
                        try
                        {
                            if (filePath.StartsWith(targetFolder, StringComparison.OrdinalIgnoreCase))
                            {
                                relativeDir = Path.GetDirectoryName(filePath.Substring(targetFolder.Length).TrimStart('\\', '/')) ?? string.Empty;
                            }
                        }
                        catch { }

                        // Đối với Chỉ Thị Test (TestSpec), tài liệu không phân version -> không trích xuất version
                        string detectedVer = (docType != SpecDocumentType.TestSpec)
                            ? ExtractVersion(fileNameWithoutExt)
                            : string.Empty;

                        list.Add(new SpecSearchResultItem
                        {
                            FilePath = filePath,
                            FileName = fileName,
                            DirectoryPath = Path.GetDirectoryName(filePath) ?? string.Empty,
                            RelativeDirectory = string.IsNullOrWhiteSpace(relativeDir) ? "." : relativeDir,
                            DocType = docType,
                            Language = language,
                            DetectedVersion = detectedVer,
                            LastModified = fi.LastWriteTime,
                            FileSizeBytes = fi.Length
                        });

                        addedPaths.Add(filePath);
                    }
                    catch { }
                }
            }

            // Sắp xếp kết quả:
            if (docType == SpecDocumentType.TestSpec)
            {
                list.Sort((a, b) => string.Compare(a.FileName, b.FileName, StringComparison.OrdinalIgnoreCase));
                for (int i = 0; i < list.Count; i++)
                {
                    list[i].IsLatest = false;
                    list[i].ItemIndex = i + 1;
                }
            }
            else
            {
                list.Sort((a, b) =>
                {
                    int verCompare = CompareVersions(b.DetectedVersion, a.DetectedVersion);
                    if (verCompare != 0) return verCompare;
                    return b.LastModified.CompareTo(a.LastModified);
                });

                for (int i = 0; i < list.Count; i++)
                {
                    list[i].IsLatest = (i == 0);
                    list[i].ItemIndex = i + 1;
                }
            }

            return list;
        }

        private class KeywordPair
        {
            public string Raw { get; set; } = string.Empty;
            public string NoSep { get; set; } = string.Empty;
        }

        /// <summary>
        /// Mở tài liệu bằng Excel (nếu là workbook) hoặc ứng dụng hệ thống mặc định
        /// </summary>
        public static bool OpenFile(string filePath, bool isReadOnly, ExcelApp? app = null)
        {
            if (!File.Exists(filePath))
            {
                System.Windows.MessageBox.Show(
                    LocalizationService.Get("SpecLauncher_FileNotFound", filePath),
                    "Thông Báo",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return false;
            }

            OpenMultipleFilesAsync(new[] { filePath }, isReadOnly, app);
            return true;
        }

        /// <summary>
        /// Mở an toàn danh sách một hoặc nhiều file tài liệu trong Excel (hoặc ứng dụng ngoài) mà không làm đơ hay crash Excel
        /// </summary>
        public static void OpenMultipleFilesAsync(IEnumerable<string> filePaths, bool isReadOnly, ExcelApp? app = null)
        {
            var pathList = filePaths?.Where(p => !string.IsNullOrWhiteSpace(p) && File.Exists(p)).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            if (pathList == null || pathList.Count == 0) return;

            // Tách các file Excel và file định dạng khác (PDF, Word, PPTX...)
            var excelFiles = new List<string>();
            var nonExcelFiles = new List<string>();

            foreach (var path in pathList)
            {
                string ext = Path.GetExtension(path).ToLowerInvariant();
                if (ext == ".xlsx" || ext == ".xlsm" || ext == ".xls" || ext == ".xlsb" || ext == ".csv")
                {
                    excelFiles.Add(path);
                }
                else
                {
                    nonExcelFiles.Add(path);
                }
            }

            // Mở các file non-Excel bằng ứng dụng mặc định
            foreach (var nonExcel in nonExcelFiles)
            {
                try
                {
                    Process.Start(new ProcessStartInfo(nonExcel) { UseShellExecute = true });
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[OpenMultipleFilesAsync] Non-Excel open error: {ex.Message}");
                }
            }

            if (excelFiles.Count == 0) return;

            // Chạy mở file Excel trên luồng Macro an toàn của Excel-DNA
            ExcelDna.Integration.ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    app ??= AddInEvents.Instance?.ExcelAppInstance ?? (ExcelApp)ExcelDna.Integration.ExcelDnaUtil.Application;
                    if (app == null) return;

                    if (AddInEvents.Instance != null)
                    {
                        AddInEvents.Instance.IsBatchProcessing = true;
                    }

                    var failedFiles = new List<string>();

                    try
                    {
                        foreach (var filePath in excelFiles)
                        {
                            try
                            {
                                string targetFileName = Path.GetFileName(filePath);
                                bool alreadyOpen = false;

                                try
                                {
                                    foreach (Microsoft.Office.Interop.Excel.Workbook wb in app.Workbooks)
                                    {
                                        try
                                        {
                                            if (string.Equals(wb.FullName, filePath, StringComparison.OrdinalIgnoreCase) ||
                                                string.Equals(wb.Name, targetFileName, StringComparison.OrdinalIgnoreCase))
                                            {
                                                alreadyOpen = true;
                                                wb.Activate();
                                                break;
                                            }
                                        }
                                        catch { }
                                    }
                                }
                                catch { }

                                if (!alreadyOpen)
                                {
                                    try
                                    {
                                        // Sử dụng tham số chuẩn: tắt update links và bỏ qua prompt ReadOnlyRecommended
                                        app.Workbooks.Open(
                                            Filename: filePath,
                                            UpdateLinks: false,
                                            ReadOnly: isReadOnly,
                                            IgnoreReadOnlyRecommended: true,
                                            AddToMru: true);
                                    }
                                    catch (Exception exCom)
                                    {
                                        Debug.WriteLine($"[OpenMultipleFilesAsync] COM Open failed for {filePath}: {exCom.Message}, falling back to Process.Start");
                                        // Fallback: Dùng Shell Execute để Excel tự mở file như khi người dùng double click
                                        Process.Start(new ProcessStartInfo(filePath) { UseShellExecute = true });
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Debug.WriteLine($"[OpenMultipleFilesAsync] Error opening {filePath}: {ex.Message}");
                                failedFiles.Add($"{Path.GetFileName(filePath)} ({ex.Message})");
                            }
                        }
                    }
                    finally
                    {
                        if (AddInEvents.Instance != null)
                        {
                            AddInEvents.Instance.IsBatchProcessing = false;
                            try { AddInEvents.Instance.QueueRefresh(); } catch { }
                        }
                    }

                    if (failedFiles.Count > 0)
                    {
                        System.Windows.MessageBox.Show(
                            LocalizationService.Get("SpecLauncher_SomeFilesOpenError", string.Join("\n", failedFiles)),
                            "Thông Báo",
                            System.Windows.MessageBoxButton.OK,
                            System.Windows.MessageBoxImage.Warning);
                    }
                }
                catch (Exception exOverall)
                {
                    Debug.WriteLine($"[OpenMultipleFilesAsync] Global macro error: {exOverall.Message}");
                    System.Windows.MessageBox.Show(
                        exOverall.Message,
                        "Thông Báo",
                        System.Windows.MessageBoxButton.OK,
                        System.Windows.MessageBoxImage.Warning);
                }
            });
        }

        /// <summary>
        /// Mở thư mục chứa file trong Windows Explorer và chọn file đó
        /// </summary>
        public static void OpenContainingFolder(string filePathOrDirectory)
        {
            try
            {
                if (File.Exists(filePathOrDirectory))
                {
                    Process.Start("explorer.exe", $"/select,\"{filePathOrDirectory}\"");
                }
                else if (Directory.Exists(filePathOrDirectory))
                {
                    Process.Start("explorer.exe", $"\"{filePathOrDirectory}\"");
                }
            }
            catch { }
        }

        public static string GetDocTypeName(SpecDocumentType docType)
        {
            return docType switch
            {
                SpecDocumentType.DetailedDesign => LocalizationService.Get("SpecLauncher_DocType_TKCT", "Thiết Kế Chi Tiết (TKCT / Detailed Design)"),
                SpecDocumentType.BasicDesign => LocalizationService.Get("SpecLauncher_DocType_TKCB", "Thiết Kế Cơ Bản (TKCB / Basic Design)"),
                SpecDocumentType.TestSpec => LocalizationService.Get("SpecLauncher_DocType_TestSpec", "Chỉ Thị Test (Test Spec)"),
                _ => "Tài Liệu Thiết Kế"
            };
        }

        public static string GetDocTypeShortBadge(SpecDocumentType docType)
        {
            return docType switch
            {
                SpecDocumentType.DetailedDesign => "TKCT",
                SpecDocumentType.BasicDesign => "TKCB",
                SpecDocumentType.TestSpec => "TEST",
                _ => "SPEC"
            };
        }

        private static string ExtractVersion(string text)
        {
            // Kiểm tra date format: _20260915
            var dateMatch = DateVersionRegex.Match(text);
            if (dateMatch.Success)
            {
                return dateMatch.Groups[1].Value;
            }

            // Kiểm tra version thông thường: _v1.0, _ver2.0, rev01
            var match = VersionRegex.Match(text);
            if (match.Success)
            {
                return match.Value.Trim('_', '-', ' ');
            }

            return string.Empty;
        }

        private static int CompareVersions(string verA, string verB)
        {
            if (string.IsNullOrEmpty(verA) && string.IsNullOrEmpty(verB)) return 0;
            if (string.IsNullOrEmpty(verA)) return -1;
            if (string.IsNullOrEmpty(verB)) return 1;

            // Làm sạch tiền tố chữ: v1.0 -> 1.0
            string cleanA = Regex.Replace(verA, @"^[a-zA-Z]+", "").Trim('.', '-', '_');
            string cleanB = Regex.Replace(verB, @"^[a-zA-Z]+", "").Trim('.', '-', '_');

            if (Version.TryParse(cleanA, out var parsedA) && Version.TryParse(cleanB, out var parsedB))
            {
                return parsedA.CompareTo(parsedB);
            }

            if (long.TryParse(cleanA, out var numA) && long.TryParse(cleanB, out var numB))
            {
                return numA.CompareTo(numB);
            }

            return string.Compare(cleanA, cleanB, StringComparison.OrdinalIgnoreCase);
        }

        private static IEnumerable<string> SafeEnumerateFiles(string rootDirectory, bool searchSubfolders = true)
        {
            var stack = new Stack<string>();
            stack.Push(rootDirectory);

            while (stack.Count > 0)
            {
                string dir = stack.Pop();

                string[]? files = null;
                try
                {
                    files = Directory.GetFiles(dir);
                }
                catch { }

                if (files != null)
                {
                    foreach (var file in files)
                    {
                        yield return file;
                    }
                }

                if (!searchSubfolders)
                {
                    continue;
                }

                string[]? subDirs = null;
                try
                {
                    subDirs = Directory.GetDirectories(dir);
                }
                catch { }

                if (subDirs != null)
                {
                    foreach (var subDir in subDirs)
                    {
                        // Bỏ qua các thư mục ẩn / hệ thống / git / node_modules / bin / obj
                        string name = Path.GetFileName(subDir);
                        if (name.StartsWith(".") || 
                            name.Equals("node_modules", StringComparison.OrdinalIgnoreCase) ||
                            name.Equals("$Recycle.Bin", StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        stack.Push(subDir);
                    }
                }
            }
        }
    }
}
