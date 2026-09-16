using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelSupport.Models
{
    /// <summary>
    /// Phân loại tài liệu thiết kế trong dự án phần mềm
    /// </summary>
    public enum SpecDocumentType
    {
        /// <summary>
        /// Thiết Kế Chi Tiết (Detailed Design - 詳細設計書 / TKCT)
        /// </summary>
        DetailedDesign = 0,

        /// <summary>
        /// Thiết Kế Cơ Bản (Basic Design - 基本設計書 / TKCB)
        /// </summary>
        BasicDesign = 1,

        /// <summary>
        /// Chỉ Thị Test (Test Specification - 単体/結合テスト仕様書)
        /// </summary>
        TestSpec = 2
    }

    /// <summary>
    /// Cấu hình Profile cho một dự án phần mềm
    /// </summary>
    public class ProjectProfile
    {
        public string Id { get; set; } = Guid.NewGuid().ToString();

        /// <summary>
        /// Tên dự án (Ví dụ: Dự Án ERP Khách Hàng Sakura)
        /// </summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>
        /// Thư mục gốc của dự án (Root directory)
        /// </summary>
        public string RootFolder { get; set; } = string.Empty;

        /// <summary>
        /// Trạng thái đang kích hoạt (chỉ dùng cho hiển thị UI)
        /// </summary>
        [Newtonsoft.Json.JsonIgnore]
        public bool IsActive { get; set; }

        /// <summary>
        /// Thư mục chứa tài liệu Thiết Kế Chi Tiết (TKCT) - Tiếng Việt.
        /// </summary>
        public string DetailedDesignFolderVi { get; set; } = string.Empty;

        /// <summary>
        /// Thư mục chứa tài liệu Thiết Kế Chi Tiết (TKCT) - Tiếng Nhật.
        /// </summary>
        public string DetailedDesignFolderJa { get; set; } = string.Empty;

        /// <summary>
        /// Thư mục chứa tài liệu Thiết Kế Cơ Bản (TKCB) - Tiếng Việt.
        /// </summary>
        public string BasicDesignFolderVi { get; set; } = string.Empty;

        /// <summary>
        /// Thư mục chứa tài liệu Thiết Kế Cơ Bản (TKCB) - Tiếng Nhật.
        /// </summary>
        public string BasicDesignFolderJa { get; set; } = string.Empty;

        /// <summary>
        /// Tương thích ngược cấu hình cũ: DetailedDesignFolder map với DetailedDesignFolderVi
        /// </summary>
        public string DetailedDesignFolder
        {
            get => !string.IsNullOrEmpty(DetailedDesignFolderVi) ? DetailedDesignFolderVi : string.Empty;
            set
            {
                if (string.IsNullOrEmpty(DetailedDesignFolderVi) && !string.IsNullOrEmpty(value))
                {
                    DetailedDesignFolderVi = value;
                }
            }
        }

        /// <summary>
        /// Tương thích ngược cấu hình cũ: BasicDesignFolder map với BasicDesignFolderVi
        /// </summary>
        public string BasicDesignFolder
        {
            get => !string.IsNullOrEmpty(BasicDesignFolderVi) ? BasicDesignFolderVi : string.Empty;
            set
            {
                if (string.IsNullOrEmpty(BasicDesignFolderVi) && !string.IsNullOrEmpty(value))
                {
                    BasicDesignFolderVi = value;
                }
            }
        }

        /// <summary>
        /// Thư mục chứa tài liệu Chỉ Thị Test (Test Spec).
        /// Có thể là đường dẫn tuyệt đối hoặc tương đối so với RootFolder.
        /// </summary>
        public string TestSpecFolder { get; set; } = string.Empty;

        /// <summary>
        /// Nếu là true, không tìm kiếm đệ quy vào các thư mục con (chỉ tìm trong thư mục cấp 1).
        /// Mặc định là false (tìm kiếm toàn bộ thư mục con).
        /// </summary>
        public bool DoNotSearchSubfolders { get; set; } = false;

        /// <summary>
        /// Mặc định mở tài liệu ở chế độ chỉ đọc (Read-Only) để tránh vô tình chỉnh sửa tài liệu gốc
        /// </summary>
        public bool OpenReadOnlyDefault { get; set; } = false;

        /// <summary>
        /// Danh sách các phần mở rộng file được hỗ trợ (cách nhau bởi dấu chấm phẩy)
        /// Ví dụ: .xlsx;.xlsm;.xls;.docx;.pdf;.pptx
        /// </summary>
        public string FileExtensions { get; set; } = ".xlsx;.xlsm;.xls;.docx;.pdf;.pptx";

        /// <summary>
        /// Thời gian cập nhật gần nhất
        /// </summary>
        public DateTime LastModified { get; set; } = DateTime.Now;

        /// <summary>
        /// Lấy đường dẫn tuyệt đối thực tế của thư mục theo loại tài liệu và ngôn ngữ
        /// </summary>
        public string GetTargetFolder(SpecDocumentType docType, string? language = "vi")
        {
            bool isJa = string.Equals(language, "ja", StringComparison.OrdinalIgnoreCase);

            string folder = docType switch
            {
                SpecDocumentType.DetailedDesign => isJa ? DetailedDesignFolderJa : DetailedDesignFolderVi,
                SpecDocumentType.BasicDesign => isJa ? BasicDesignFolderJa : BasicDesignFolderVi,
                SpecDocumentType.TestSpec => TestSpecFolder,
                _ => string.Empty
            };

            if (string.IsNullOrWhiteSpace(folder))
            {
                return string.Empty;
            }

            // Nếu là đường dẫn tuyệt đối đã có drive letter hoặc root UNC
            if (Path.IsPathRooted(folder))
            {
                return folder;
            }

            // Nếu là đường dẫn tương đối, ghép với RootFolder
            if (!string.IsNullOrWhiteSpace(RootFolder))
            {
                return Path.Combine(RootFolder, folder);
            }

            return folder;
        }

        public ProjectProfile Clone()
        {
            return new ProjectProfile
            {
                Id = Guid.NewGuid().ToString(),
                Name = $"{Name} (Bản sao)",
                RootFolder = RootFolder,
                DetailedDesignFolderVi = DetailedDesignFolderVi,
                DetailedDesignFolderJa = DetailedDesignFolderJa,
                BasicDesignFolderVi = BasicDesignFolderVi,
                BasicDesignFolderJa = BasicDesignFolderJa,
                TestSpecFolder = TestSpecFolder,
                DoNotSearchSubfolders = DoNotSearchSubfolders,
                OpenReadOnlyDefault = OpenReadOnlyDefault,
                FileExtensions = FileExtensions,
                LastModified = DateTime.Now
            };
        }
    }

    /// <summary>
    /// File cấu hình toàn cục lưu trữ trong %APPDATA%\ExcelSupport\project_profiles.json
    /// </summary>
    public class ProjectProfilesConfig
    {
        public string? ActiveProfileId { get; set; }
        public List<ProjectProfile> Profiles { get; set; } = new List<ProjectProfile>();
    }

    /// <summary>
    /// Đại diện cho 1 kết quả file tài liệu tìm thấy
    /// </summary>
    public class SpecSearchResultItem
    {
        public string FilePath { get; set; } = string.Empty;
        public string FileName { get; set; } = string.Empty;
        public string DirectoryPath { get; set; } = string.Empty;
        public string RelativeDirectory { get; set; } = string.Empty;
        public SpecDocumentType DocType { get; set; }
        public string Language { get; set; } = "vi"; // "vi" hoặc "ja"
        public string DetectedVersion { get; set; } = string.Empty;
        public DateTime LastModified { get; set; }
        public long FileSizeBytes { get; set; }
        public bool IsLatest { get; set; }
        public int ItemIndex { get; set; }
        public bool HasVersion => !string.IsNullOrWhiteSpace(DetectedVersion);

        public string FormattedSize
        {
            get
            {
                if (FileSizeBytes < 1024) return $"{FileSizeBytes} B";
                if (FileSizeBytes < 1024 * 1024) return $"{FileSizeBytes / 1024.0:F1} KB";
                return $"{FileSizeBytes / (1024.0 * 1024.0):F2} MB";
            }
        }

        public string Extension => Path.GetExtension(FilePath).ToLowerInvariant();

        public bool IsExcelFile => Extension == ".xlsx" || Extension == ".xlsm" || Extension == ".xls";
    }

    /// <summary>
    /// Kết quả thực thi khởi chạy mở tài liệu thiết kế
    /// </summary>
    public class LaunchResult
    {
        public bool Success { get; set; } = true;
        public string Message { get; set; } = string.Empty;

        public static LaunchResult Ok() => new LaunchResult { Success = true };
        public static LaunchResult Fail(string message) => new LaunchResult { Success = false, Message = message };
    }
}
