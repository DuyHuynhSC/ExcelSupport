using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace ExcelSupport.Helpers
{
    /// <summary>
    /// Tiện ích quản lý vòng đời và tự động giải phóng các đối tượng COM Interop theo phạm vi using,
    /// loại bỏ hoàn toàn nguy cơ rò rỉ bộ nhớ hoặc treo tiến trình EXCEL.EXE ngầm.
    /// </summary>
    public sealed class ComScope : IDisposable
    {
        private readonly List<object> _trackedObjects = new();
        private bool _disposed;

        /// <summary>
        /// Đăng ký theo dõi một đối tượng COM để tự động giải phóng khi ra khỏi khối using.
        /// </summary>
        public T? Track<T>(T? comObject) where T : class
        {
            if (comObject != null && Marshal.IsComObject(comObject))
            {
                lock (_trackedObjects)
                {
                    _trackedObjects.Add(comObject);
                }
            }
            return comObject;
        }

        /// <summary>
        /// Giải phóng toàn bộ các đối tượng COM đã được theo dõi theo thứ tự ngược lại (LIFO).
        /// </summary>
        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;

            lock (_trackedObjects)
            {
                for (int i = _trackedObjects.Count - 1; i >= 0; i--)
                {
                    ComHelper.Release(_trackedObjects[i]);
                }
                _trackedObjects.Clear();
            }
        }
    }

    /// <summary>
    /// Các phương thức tiện ích giải phóng đối tượng COM an toàn tuyệt đối.
    /// </summary>
    public static class ComHelper
    {
        /// <summary>
        /// Giải phóng một đối tượng COM và gán biến về null.
        /// </summary>
        public static void Release<T>(ref T? comObject) where T : class
        {
            if (comObject != null)
            {
                Release((object)comObject);
                comObject = null;
            }
        }

        /// <summary>
        /// Giải phóng một đối tượng COM mà không làm phát sinh ngoại lệ.
        /// </summary>
        public static void Release(object? comObject)
        {
            if (comObject != null && Marshal.IsComObject(comObject))
            {
                try
                {
                    Marshal.ReleaseComObject(comObject);
                }
                catch
                {
                    // Nuốt lỗi COM exception nếu đối tượng đã bị hủy trước đó
                }
            }
        }

        /// <summary>
        /// Giải phóng hàng loạt nhiều đối tượng COM.
        /// </summary>
        public static void SafeRelease(params object?[] comObjects)
        {
            if (comObjects == null) return;
            foreach (var obj in comObjects)
            {
                Release(obj);
            }
        }
    }
}
