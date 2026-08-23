using System;
using Microsoft.Win32;

namespace ExcelSupport.Host
{
    public static class AppSettings
    {
        private const string RegKeyPath = @"Software\ExcelSupport\Settings";
        private const string TaskPaneOpenKey = "IsTaskPaneOpen";

        /// <summary>
        /// Trạng thái TaskPane được lưu trong Windows Registry (HKCU) để duy trì giữa các phiên Excel
        /// </summary>
        public static bool IsTaskPaneAutoOpen
        {
            get
            {
                try
                {
                    using (var key = Registry.CurrentUser.OpenSubKey(RegKeyPath))
                    {
                        if (key != null)
                        {
                            var val = key.GetValue(TaskPaneOpenKey);
                            if (val is int intVal) return intVal == 1;
                            if (val is string strVal && bool.TryParse(strVal, out bool b)) return b;
                        }
                    }
                }
                catch { }
                return false;
            }
            set
            {
                try
                {
                    using (var key = Registry.CurrentUser.CreateSubKey(RegKeyPath))
                    {
                        key?.SetValue(TaskPaneOpenKey, value ? 1 : 0, RegistryValueKind.DWord);
                    }
                }
                catch { }
            }
        }

        private const string LanguageKey = "AppLanguage";

        /// <summary>
        /// Ngôn ngữ giao diện (vi = Tiếng Việt, en = English, ja = 日本語)
        /// </summary>
        public static string CurrentLanguage
        {
            get
            {
                try
                {
                    using (var key = Registry.CurrentUser.OpenSubKey(RegKeyPath))
                    {
                        if (key != null)
                        {
                            var val = key.GetValue(LanguageKey) as string;
                            if (!string.IsNullOrEmpty(val)) return val!;
                        }
                    }
                }
                catch { }
                return "vi"; // Mặc định là Tiếng Việt
            }
            set
            {
                try
                {
                    using (var key = Registry.CurrentUser.CreateSubKey(RegKeyPath))
                    {
                        key?.SetValue(LanguageKey, value ?? "vi", RegistryValueKind.String);
                    }
                }
                catch { }
            }
        }
    }
}
