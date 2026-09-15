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

        /// <summary>
        /// Kích hoạt tìm kiếm và mở tài liệu từ ô đang chọn trên Excel
        /// </summary>
        public static void LaunchFromSelection(ExcelApp? app, SpecDocumentType docType, bool? forceReadOnly = null)
        {
            if (app == null)
            {
                app = AddInEvents.Instance?.ExcelAppInstance ?? (ExcelApp)ExcelDna.Integration.ExcelDnaUtil.Application;
            }

            var profile = ProjectProfileManager.GetActiveProfile();
            if (profile == null)
            {
                var ask = WpfMessageBox.Show(
                    LocalizationService.Get("SpecLauncher_NoProfilePrompt", 
                        "Chưa có Profile dự án nào được thiết lập. Bạn có muốn mở Cài Đặt Profile Dự Án ngay bây giờ?"),
                    LocalizationService.Get("SpecLauncher_WindowTitle", "Project Document & Quick Spec Launcher"),
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question);

                if (ask == MessageBoxResult.Yes)
                {
                    ProjectProfileSettingsDialog.ShowWindow(AddInEvents.MainViewModel?.IsDarkTheme ?? false);
                }
                return;
            }

            string targetFolder = profile.GetTargetFolder(docType);
            if (string.IsNullOrWhiteSpace(targetFolder) || !Directory.Exists(targetFolder))
            {
                string docTypeName = GetDocTypeName(docType);
                var ask = WpfMessageBox.Show(
                    string.Format(LocalizationService.Get("SpecLauncher_FolderNotFoundPrompt",
                        "Thư mục tài liệu {0} của dự án '{1}' chưa được thiết lập hoặc không tồn tại:\n{2}\n\nBạn có muốn mở Cài Đặt Profile Dự Án để cập nhật đường dẫn?"),
                        docTypeName, profile.Name, targetFolder),
                    LocalizationService.Get("SpecLauncher_WindowTitle", "Project Document & Quick Spec Launcher"),
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning);

                if (ask == MessageBoxResult.Yes)
                {
                    ProjectProfileSettingsDialog.ShowWindow(AddInEvents.MainViewModel?.IsDarkTheme ?? false);
                }
                return;
            }

            // Lấy từ khóa từ ô đang chọn
            string keyword = string.Empty;
            try
            {
                if (app.ActiveCell != null)
                {
                    object cellVal = app.ActiveCell.Text ?? app.ActiveCell.Value;
                    if (cellVal != null)
                    {
                        keyword = cellVal.ToString()?.Trim() ?? string.Empty;
                    }
                }
            }
            catch { }

            bool isReadOnly = forceReadOnly ?? profile.OpenReadOnlyDefault;

            // Nếu ô không có text, mở hộp thoại chọn version với ô tìm kiếm để người dùng tự nhập từ khóa
            if (string.IsNullOrWhiteSpace(keyword))
            {
                SpecVersionSelectorDialog.ShowWindow(profile, docType, string.Empty, new List<SpecSearchResultItem>(), isReadOnly, AddInEvents.MainViewModel?.IsDarkTheme ?? false);
                return;
            }

            // Tìm kiếm file đệ quy
            var results = SearchFiles(profile, docType, keyword);

            if (results.Count == 0)
            {
                string docTypeName = GetDocTypeName(docType);
                var promptResult = WpfMessageBox.Show(
                    string.Format(LocalizationService.Get("SpecLauncher_NoFilesFoundPrompt",
                        "Không tìm thấy tài liệu {0} nào chứa từ khóa: '{1}'\n\nThư mục quét:\n{2}\n\nBạn có muốn mở thư mục này trong Windows Explorer để kiểm tra không?"),
                        docTypeName, keyword, targetFolder),
                    LocalizationService.Get("SpecLauncher_WindowTitle", "Project Document & Quick Spec Launcher"),
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Information);

                if (promptResult == MessageBoxResult.Yes)
                {
                    OpenContainingFolder(targetFolder);
                }
                return;
            }

            // Nếu chỉ tìm thấy đúng 1 file duy nhất -> Mở trực tiếp ngay lập tức!
            if (results.Count == 1)
            {
                OpenFile(results[0].FilePath, isReadOnly, app);
                return;
            }

            // Nếu tìm thấy nhiều file hoặc nhiều version -> Mở hộp thoại Version Selector
            SpecVersionSelectorDialog.ShowWindow(profile, docType, keyword, results, isReadOnly, AddInEvents.MainViewModel?.IsDarkTheme ?? false);
        }

        /// <summary>
        /// Tìm kiếm đệ quy toàn bộ file trong thư mục chỉ định của profile khớp với từ khóa
        /// </summary>
        public static List<SpecSearchResultItem> SearchFiles(ProjectProfile profile, SpecDocumentType docType, string keyword)
        {
            var list = new List<SpecSearchResultItem>();
            string targetFolder = profile.GetTargetFolder(docType);

            if (string.IsNullOrWhiteSpace(targetFolder) || !Directory.Exists(targetFolder))
            {
                return list;
            }

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

            // Chuẩn hóa từ khóa tìm kiếm (bỏ khoảng trắng, kiểm tra biến thể _ và -)
            string normalizedKeyword = keyword.Trim();
            string keywordNoSep = normalizedKeyword.Replace("_", "").Replace("-", "").Replace(" ", "");

            // Quét đệ quy thư mục an toàn (bỏ qua folder không có quyền truy cập)
            var allFiles = SafeEnumerateFiles(targetFolder);

            foreach (var filePath in allFiles)
            {
                string ext = Path.GetExtension(filePath);
                if (!allowedExtensions.Contains(ext)) continue;

                string fileName = Path.GetFileName(filePath);
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(filePath);

                // Khớp Contains Match không phân biệt hoa thường
                bool isMatch = fileNameWithoutExt.IndexOf(normalizedKeyword, StringComparison.OrdinalIgnoreCase) >= 0;

                if (!isMatch && keywordNoSep.Length > 2)
                {
                    string fileNameNoSep = fileNameWithoutExt.Replace("_", "").Replace("-", "").Replace(" ", "");
                    isMatch = fileNameNoSep.IndexOf(keywordNoSep, StringComparison.OrdinalIgnoreCase) >= 0;
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

                        string detectedVer = ExtractVersion(fileNameWithoutExt);

                        list.Add(new SpecSearchResultItem
                        {
                            FilePath = filePath,
                            FileName = fileName,
                            DirectoryPath = Path.GetDirectoryName(filePath) ?? string.Empty,
                            RelativeDirectory = string.IsNullOrWhiteSpace(relativeDir) ? "." : relativeDir,
                            DocType = docType,
                            DetectedVersion = detectedVer,
                            LastModified = fi.LastWriteTime,
                            FileSizeBytes = fi.Length
                        });
                    }
                    catch { }
                }
            }

            // Sắp xếp ưu tiên bản mới nhất:
            // 1. Phân tích version/date kết hợp ngày sửa đổi LastWriteTime giảm dần
            list.Sort((a, b) =>
            {
                int verCompare = CompareVersions(b.DetectedVersion, a.DetectedVersion);
                if (verCompare != 0) return verCompare;
                return b.LastModified.CompareTo(a.LastModified);
            });

            // Đánh dấu IsLatest và gán ItemIndex (1-9) cho phím tắt nhanh
            for (int i = 0; i < list.Count; i++)
            {
                list[i].IsLatest = (i == 0);
                list[i].ItemIndex = i + 1;
            }

            return list;
        }

        /// <summary>
        /// Mở tài liệu bằng Excel (nếu là workbook) hoặc ứng dụng hệ thống mặc định
        /// </summary>
        public static bool OpenFile(string filePath, bool isReadOnly, ExcelApp? app = null)
        {
            if (!File.Exists(filePath))
            {
                WpfMessageBox.Show(
                    string.Format(LocalizationService.Get("SpecLauncher_FileNotFound", "File tài liệu không tồn tại trên đĩa:\n{0}"), filePath),
                    LocalizationService.Get("SpecLauncher_WindowTitle", "Project Document & Quick Spec Launcher"),
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                return false;
            }

            string ext = Path.GetExtension(filePath).ToLowerInvariant();
            bool isExcel = ext == ".xlsx" || ext == ".xlsm" || ext == ".xls" || ext == ".xlsb" || ext == ".csv";

            try
            {
                if (isExcel)
                {
                    if (app == null)
                    {
                        app = AddInEvents.Instance?.ExcelAppInstance ?? (ExcelApp)ExcelDna.Integration.ExcelDnaUtil.Application;
                    }

                    // Kiểm tra xem file đã được mở trong Excel hiện tại chưa
                    foreach (Microsoft.Office.Interop.Excel.Workbook wb in app.Workbooks)
                    {
                        if (string.Equals(wb.FullName, filePath, StringComparison.OrdinalIgnoreCase))
                        {
                            wb.Activate();
                            app.Visible = true;
                            return true;
                        }
                    }

                    // Mở file trong Excel
                    app.Workbooks.Open(filePath, ReadOnly: isReadOnly);
                    app.Visible = true;
                    return true;
                }
                else
                {
                    // Mở bằng ứng dụng mặc định của Windows
                    var psi = new ProcessStartInfo(filePath)
                    {
                        UseShellExecute = true
                    };
                    Process.Start(psi);
                    return true;
                }
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show(
                    string.Format(LocalizationService.Get("SpecLauncher_ErrorOpeningFile", "Không thể mở file tài liệu:\n{0}\n\nLỗi: {1}"), filePath, ex.Message),
                    LocalizationService.Get("SpecLauncher_WindowTitle", "Project Document & Quick Spec Launcher"),
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                return false;
            }
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

        private static IEnumerable<string> SafeEnumerateFiles(string rootDirectory)
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
