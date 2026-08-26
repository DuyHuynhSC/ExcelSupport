using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Newtonsoft.Json;

namespace ExcelSupport.Services
{
    public class RibbonControlMetadata
    {
        public string ControlId { get; set; } = "";
        public string GroupId { get; set; } = "";
        public string GroupNameKey { get; set; } = "";
        public string NameKey { get; set; } = "";
        public string IconEmoji { get; set; } = "🔹";
        public bool IsVisible { get; set; } = true;
    }

    public static class RibbonVisibilityService
    {
        private static readonly string ConfigDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "ExcelSupport"
        );
        private static readonly string ConfigFilePath = Path.Combine(ConfigDirectory, "ribbon_visibility.json");

        private static Dictionary<string, bool>? _cachedVisibility;
        private static readonly object _lock = new object();

        public static List<RibbonControlMetadata> GetAllControlsMetadata()
        {
            var visibilityMap = GetVisibilityMap();

            var list = new List<RibbonControlMetadata>
            {
                // Group 4: Xử Lý Dữ Liệu (grpDataTools)
                new() { ControlId = "btnAdvancedFilter", GroupId = "grpDataTools", GroupNameKey = "grpDataTools", NameKey = "btnAdvancedFilter", IconEmoji = "⚡" },
                new() { ControlId = "splitFilteredCopyPaste", GroupId = "grpDataTools", GroupNameKey = "grpDataTools", NameKey = "btnFilteredCopyPasteWizard", IconEmoji = "📋" },
                new() { ControlId = "btnDataCleaner", GroupId = "grpDataTools", GroupNameKey = "grpDataTools", NameKey = "btnDataCleaner", IconEmoji = "🧹" },
                new() { ControlId = "btnDuplicateFinder", GroupId = "grpDataTools", GroupNameKey = "grpDataTools", NameKey = "btnDuplicateFinder", IconEmoji = "📑" },
                new() { ControlId = "btnBatchBlankCleaner", GroupId = "grpDataTools", GroupNameKey = "grpDataTools", NameKey = "btnBatchBlankCleaner", IconEmoji = "⛔" },
                new() { ControlId = "btnBatchFindReplace", GroupId = "grpDataTools", GroupNameKey = "grpDataTools", NameKey = "btnBatchFindReplace", IconEmoji = "🔍" },
                new() { ControlId = "btnVisualTableMerge", GroupId = "grpDataTools", GroupNameKey = "grpDataTools", NameKey = "btnVisualTableMerge", IconEmoji = "📊" },
                new() { ControlId = "btnFuzzyDuplicate", GroupId = "grpDataTools", GroupNameKey = "grpDataTools", NameKey = "btnFuzzyDuplicate", IconEmoji = "🔄" },
                new() { ControlId = "btnSafeMergeConsolidate", GroupId = "grpDataTools", GroupNameKey = "grpDataTools", NameKey = "btnSafeMergeConsolidate", IconEmoji = "🔀" },

                // Group 3: Kiểm Tra & Đối Soát (grpAuditTools)
                new() { ControlId = "btnCompareWorkbooks", GroupId = "grpAuditTools", GroupNameKey = "grpAuditTools", NameKey = "btnCompareWorkbooks", IconEmoji = "⚖️" },
                new() { ControlId = "btnCheckVietnamese", GroupId = "grpAuditTools", GroupNameKey = "grpAuditTools", NameKey = "btnCheckVietnamese", IconEmoji = "🇻🇳" },
                new() { ControlId = "btnExternalLinks", GroupId = "grpAuditTools", GroupNameKey = "grpAuditTools", NameKey = "btnExternalLinks", IconEmoji = "🔗" },
                new() { ControlId = "btnOracleTableCompare", GroupId = "grpAuditTools", GroupNameKey = "grpAuditTools", NameKey = "btnOracleTableCompare", IconEmoji = "🗄️" },
                new() { ControlId = "btnOracleQuickQuery", GroupId = "grpAuditTools", GroupNameKey = "grpAuditTools", NameKey = "btnOracleQuickQuery", IconEmoji = "⚡" },

                // Group 2: Thao Tác Nhanh (grpQuickTools)
                new() { ControlId = "btnCreateTOC", GroupId = "grpQuickTools", GroupNameKey = "grpQuickTools", NameKey = "btnCreateTOC", IconEmoji = "📑" },
                new() { ControlId = "btnSplitSheets", GroupId = "grpQuickTools", GroupNameKey = "grpQuickTools", NameKey = "btnSplitSheets", IconEmoji = "📤" },
                new() { ControlId = "btnMergeSheets", GroupId = "grpQuickTools", GroupNameKey = "grpQuickTools", NameKey = "btnMergeSheets", IconEmoji = "📥" },
                new() { ControlId = "btnQuickSortAZ", GroupId = "grpQuickTools", GroupNameKey = "grpQuickTools", NameKey = "btnQuickSortAZ", IconEmoji = "🔤" },
                new() { ControlId = "btnQuickSortZA", GroupId = "grpQuickTools", GroupNameKey = "grpQuickTools", NameKey = "btnQuickSortZA", IconEmoji = "🔤" },
                new() { ControlId = "btnCloseCurrentWb", GroupId = "grpQuickTools", GroupNameKey = "grpQuickTools", NameKey = "btnCloseCurrentWb", IconEmoji = "❌" },

                // Group 5: Quản Trị Tập Tin (grpFileTools)
                new() { ControlId = "btnBatchFileConverter", GroupId = "grpFileTools", GroupNameKey = "grpFileTools", NameKey = "btnBatchFileConverter", IconEmoji = "🔄" },
                new() { ControlId = "btnDesignPageCounter", GroupId = "grpFileTools", GroupNameKey = "grpFileTools", NameKey = "btnDesignPageCounter", IconEmoji = "📑" },

                // Group 7: Trợ Lý AI & Năng Suất (grpAiTools)
                new() { ControlId = "btnAiFormula", GroupId = "grpAiTools", GroupNameKey = "grpAiTools", NameKey = "btnAiFormula", IconEmoji = "✨" },
                new() { ControlId = "btnAiFormulaDoctor", GroupId = "grpAiTools", GroupNameKey = "grpAiTools", NameKey = "btnAiFormulaDoctor", IconEmoji = "🩺" },
                new() { ControlId = "btnSnapshotRollback", GroupId = "grpAiTools", GroupNameKey = "grpAiTools", NameKey = "btnSnapshotRollback", IconEmoji = "📸" },
                new() { ControlId = "btnUserManual", GroupId = "grpAiTools", GroupNameKey = "grpAiTools", NameKey = "btnUserManual", IconEmoji = "📖" }
            };

            foreach (var item in list)
            {
                if (visibilityMap.TryGetValue(item.ControlId, out bool isVis))
                {
                    item.IsVisible = isVis;
                }
                else
                {
                    item.IsVisible = true;
                }
            }

            return list;
        }

        public static Dictionary<string, bool> GetVisibilityMap()
        {
            lock (_lock)
            {
                if (_cachedVisibility != null)
                {
                    return new Dictionary<string, bool>(_cachedVisibility);
                }

                try
                {
                    if (File.Exists(ConfigFilePath))
                    {
                        string json = File.ReadAllText(ConfigFilePath);
                        var map = JsonConvert.DeserializeObject<Dictionary<string, bool>>(json);
                        if (map != null)
                        {
                            _cachedVisibility = map;
                            return new Dictionary<string, bool>(_cachedVisibility);
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[RibbonVisibilityService] Load error: {ex.Message}");
                }

                _cachedVisibility = new Dictionary<string, bool>();
                return new Dictionary<string, bool>(_cachedVisibility);
            }
        }

        public static bool IsControlVisible(string controlId)
        {
            if (string.IsNullOrEmpty(controlId)) return true;

            var map = GetVisibilityMap();
            if (map.TryGetValue(controlId, out bool isVis))
            {
                return isVis;
            }
            return true; // Mặc định hiện
        }

        public static bool SaveVisibilityMap(Dictionary<string, bool> newMap)
        {
            lock (_lock)
            {
                try
                {
                    if (!Directory.Exists(ConfigDirectory))
                    {
                        Directory.CreateDirectory(ConfigDirectory);
                    }

                    string json = JsonConvert.SerializeObject(newMap, Formatting.Indented);
                    File.WriteAllText(ConfigFilePath, json);
                    _cachedVisibility = new Dictionary<string, bool>(newMap);
                    return true;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[RibbonVisibilityService] Save error: {ex.Message}");
                    return false;
                }
            }
        }

        public static bool ResetToDefault()
        {
            lock (_lock)
            {
                try
                {
                    if (File.Exists(ConfigFilePath))
                    {
                        File.Delete(ConfigFilePath);
                    }
                    _cachedVisibility = new Dictionary<string, bool>();
                    return true;
                }
                catch
                {
                    return false;
                }
            }
        }
    }
}
