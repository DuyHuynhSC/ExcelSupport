using System;
using Microsoft.Win32;

namespace ExcelSupport.Host
{
    public static class AppSettings
    {
        private const string RegKeyPath = @"Software\ExcelSupport\Settings";
        private const string TaskPaneOpenKey = "IsTaskPaneOpen";
        private const string LanguageKey = "AppLanguage";
        private const string QuickActionBarKey = "IsQuickActionBarEnabled";

        /// <summary>
        /// Trạng thái TaskPane được lưu trong Windows Registry (HKCU) để duy trì giữa các phiên Excel
        /// </summary>
        public static bool IsTaskPaneAutoOpen
        {
            get => GetBool(TaskPaneOpenKey, false);
            set => SetBool(TaskPaneOpenKey, value);
        }

        public static string CurrentLanguage
        {
            get => GetString(LanguageKey, "vi");
            set => SetString(LanguageKey, value ?? "vi");
        }

        public static bool IsQuickActionBarEnabled
        {
            get => GetBool(QuickActionBarKey, true);
            set => SetBool(QuickActionBarKey, value);
        }

        private static bool GetBool(string keyName, bool defaultValue)
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(RegKeyPath);
                var val = key?.GetValue(keyName);
                if (val is int intVal) return intVal == 1;
                if (val is string strVal && bool.TryParse(strVal, out bool b)) return b;
            }
            catch { }
            return defaultValue;
        }

        private static void SetBool(string keyName, bool value)
        {
            try
            {
                using var key = Registry.CurrentUser.CreateSubKey(RegKeyPath);
                key?.SetValue(keyName, value ? 1 : 0, RegistryValueKind.DWord);
            }
            catch { }
        }

        private static string GetString(string keyName, string defaultValue)
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(RegKeyPath);
                if (key?.GetValue(keyName) is string val && !string.IsNullOrEmpty(val)) return val;
            }
            catch { }
            return defaultValue;
        }

        private static void SetString(string keyName, string value)
        {
            try
            {
                using var key = Registry.CurrentUser.CreateSubKey(RegKeyPath);
                key?.SetValue(keyName, value, RegistryValueKind.String);
            }
            catch { }
        }
    }
}
