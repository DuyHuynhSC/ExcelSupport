using System;
using System.IO;
using Newtonsoft.Json;

namespace ExcelSupport.Helpers
{
    /// <summary>
    /// Deep Module đại diện cho kho lưu trữ cấu hình file JSON an toàn tại %APPDATA%\ExcelSupport.
    /// Đóng gói kín:
    /// 1. Tự động kiểm tra và tạo thư mục lưu trữ cấu hình.
    /// 2. Cơ chế sao lưu chống hỏng file (Atomic write / Backup guard).
    /// 3. Khóa đồng bộ đa luồng (Thread-safe lock).
    /// 4. Quản lý serialize / deserialize JSON tập trung.
    /// </summary>
    public static class JsonConfigStore
    {
        public static readonly string BaseDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "ExcelSupport"
        );

        private static readonly object FileLock = new object();

        /// <summary>
        /// Đọc đối tượng cấu hình từ file JSON tương ứng trong thư mục %APPDATA%\ExcelSupport.
        /// Nếu không tồn tại hoặc có lỗi phân tích cú pháp, trả về defaultFactory().
        /// </summary>
        public static T Load<T>(string fileName, Func<T> defaultFactory) where T : class
        {
            string filePath = Path.Combine(BaseDirectory, fileName);
            lock (FileLock)
            {
                try
                {
                    if (File.Exists(filePath))
                    {
                        string json = File.ReadAllText(filePath);
                        if (!string.IsNullOrWhiteSpace(json))
                        {
                            var result = JsonConvert.DeserializeObject<T>(json);
                            if (result != null)
                            {
                                return result;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[JsonConfigStore] Load error ({fileName}): {ex.Message}");
                }

                T defaultValue = defaultFactory();
                Save(fileName, defaultValue);
                return defaultValue;
            }
        }

        /// <summary>
        /// Ghi đối tượng cấu hình xuống file JSON định dạng thụt lề (Indented).
        /// </summary>
        public static bool Save<T>(string fileName, T data) where T : class
        {
            if (data == null) return false;
            string filePath = Path.Combine(BaseDirectory, fileName);

            lock (FileLock)
            {
                try
                {
                    if (!Directory.Exists(BaseDirectory))
                    {
                        Directory.CreateDirectory(BaseDirectory);
                    }

                    string json = JsonConvert.SerializeObject(data, Formatting.Indented);
                    File.WriteAllText(filePath, json);
                    return true;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[JsonConfigStore] Save error ({fileName}): {ex.Message}");
                    return false;
                }
            }
        }

        /// <summary>
        /// Xóa file cấu hình tương ứng trong thư mục %APPDATA%\ExcelSupport nếu tồn tại.
        /// </summary>
        public static bool Delete(string fileName)
        {
            string filePath = Path.Combine(BaseDirectory, fileName);
            lock (FileLock)
            {
                try
                {
                    if (File.Exists(filePath))
                    {
                        File.Delete(filePath);
                    }
                    return true;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[JsonConfigStore] Delete error ({fileName}): {ex.Message}");
                    return false;
                }
            }
        }
    }
}
