using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelSupport.Models;
using Newtonsoft.Json;

namespace ExcelSupport.Services
{
    /// <summary>
    /// Quản lý cấu hình danh sách các mẫu định dạng Japanese Preset lưu trữ trong %APPDATA%\ExcelSupport\format_presets.json
    /// </summary>
    public static class JapanesePresetFormatManager
    {
        private static readonly string ConfigDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "ExcelSupport"
        );

        private static readonly string ConfigFilePath = Path.Combine(ConfigDirectory, "format_presets.json");

        private static JapanesePresetConfig? _currentConfig;
        private static readonly object SyncLock = new object();

        public static event Action? ActivePresetChanged;
        public static event Action? PresetsUpdated;

        public static JapanesePresetConfig CurrentConfig
        {
            get
            {
                lock (SyncLock)
                {
                    if (_currentConfig == null)
                    {
                        _currentConfig = LoadConfig();
                    }
                    return _currentConfig;
                }
            }
        }

        public static JapaneseFormatPreset? GetActivePreset()
        {
            var config = CurrentConfig;
            if (string.IsNullOrWhiteSpace(config.ActivePresetId))
            {
                return config.Presets.FirstOrDefault();
            }

            var active = config.Presets.FirstOrDefault(p => p.Id == config.ActivePresetId);
            return active ?? config.Presets.FirstOrDefault();
        }

        public static JapaneseFormatPreset? GetPresetById(string id)
        {
            var config = CurrentConfig;
            return config.Presets.FirstOrDefault(p => p.Id == id);
        }

        public static void SetActivePreset(string presetId)
        {
            lock (SyncLock)
            {
                var config = CurrentConfig;
                if (config.Presets.Any(p => p.Id == presetId))
                {
                    config.ActivePresetId = presetId;
                    SaveConfig(config);
                    ActivePresetChanged?.Invoke();
                }
            }
        }

        public static void SaveAllPresets(List<JapaneseFormatPreset> presets, string? activePresetId = null)
        {
            lock (SyncLock)
            {
                var config = CurrentConfig;
                config.Presets = new List<JapaneseFormatPreset>(presets);
                if (!string.IsNullOrEmpty(activePresetId))
                {
                    config.ActivePresetId = activePresetId!;
                }
                else if (!config.Presets.Any(p => p.Id == config.ActivePresetId))
                {
                    config.ActivePresetId = config.Presets.FirstOrDefault()?.Id ?? string.Empty;
                }

                SaveConfig(config);
                PresetsUpdated?.Invoke();
                ActivePresetChanged?.Invoke();
            }
        }

        public static void AddPreset(JapaneseFormatPreset preset)
        {
            lock (SyncLock)
            {
                var config = CurrentConfig;
                config.Presets.Add(preset);
                SaveConfig(config);
                PresetsUpdated?.Invoke();
            }
        }

        public static void UpdatePreset(JapaneseFormatPreset preset)
        {
            lock (SyncLock)
            {
                var config = CurrentConfig;
                var existing = config.Presets.FirstOrDefault(p => p.Id == preset.Id);
                if (existing != null)
                {
                    int index = config.Presets.IndexOf(existing);
                    config.Presets[index] = preset;
                    SaveConfig(config);
                    PresetsUpdated?.Invoke();
                }
            }
        }

        public static bool DeletePreset(string presetId)
        {
            lock (SyncLock)
            {
                var config = CurrentConfig;
                var existing = config.Presets.FirstOrDefault(p => p.Id == presetId);
                if (existing == null) return false;

                config.Presets.Remove(existing);
                if (config.ActivePresetId == presetId)
                {
                    config.ActivePresetId = config.Presets.FirstOrDefault()?.Id ?? string.Empty;
                }

                SaveConfig(config);
                PresetsUpdated?.Invoke();
                ActivePresetChanged?.Invoke();
                return true;
            }
        }

        public static void ResetToDefaults()
        {
            lock (SyncLock)
            {
                var config = new JapanesePresetConfig
                {
                    Version = 1,
                    ActivePresetId = "builtin_table_header",
                    Presets = JapaneseFormatPreset.CreateDefaultPresets()
                };

                SaveConfig(config);
                _currentConfig = config;
                PresetsUpdated?.Invoke();
                ActivePresetChanged?.Invoke();
            }
        }

        public static void ExportToFile(string destinationPath)
        {
            var config = CurrentConfig;
            var json = JsonConvert.SerializeObject(config, Formatting.Indented);
            File.WriteAllText(destinationPath, json);
        }

        public static bool ImportFromFile(string sourcePath, bool appendOnly = false)
        {
            if (!File.Exists(sourcePath)) return false;

            try
            {
                var json = File.ReadAllText(sourcePath);
                var imported = JsonConvert.DeserializeObject<JapanesePresetConfig>(json);
                if (imported == null || imported.Presets == null || imported.Presets.Count == 0)
                {
                    return false;
                }

                lock (SyncLock)
                {
                    var config = CurrentConfig;
                    if (appendOnly)
                    {
                        foreach (var p in imported.Presets)
                        {
                            p.Id = Guid.NewGuid().ToString(); // gán ID mới tránh trùng
                            config.Presets.Add(p);
                        }
                    }
                    else
                    {
                        config.Presets = imported.Presets;
                        if (!string.IsNullOrEmpty(imported.ActivePresetId))
                        {
                            config.ActivePresetId = imported.ActivePresetId;
                        }
                    }

                    SaveConfig(config);
                    PresetsUpdated?.Invoke();
                    ActivePresetChanged?.Invoke();
                }
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[JapanesePresetFormatManager] Import error: {ex.Message}");
                return false;
            }
        }

        private static JapanesePresetConfig LoadConfig()
        {
            try
            {
                if (File.Exists(ConfigFilePath))
                {
                    var json = File.ReadAllText(ConfigFilePath);
                    var cfg = JsonConvert.DeserializeObject<JapanesePresetConfig>(json);
                    if (cfg != null && cfg.Presets != null && cfg.Presets.Count > 0)
                    {
                        return cfg;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[JapanesePresetFormatManager] Load error: {ex.Message}");
            }

            // Mặc định tạo mới nếu chưa tồn tại hoặc file lỗi
            var defaultCfg = new JapanesePresetConfig
            {
                Version = 1,
                ActivePresetId = "builtin_table_header",
                Presets = JapaneseFormatPreset.CreateDefaultPresets()
            };

            SaveConfig(defaultCfg);
            return defaultCfg;
        }

        private static void SaveConfig(JapanesePresetConfig config)
        {
            try
            {
                if (!Directory.Exists(ConfigDirectory))
                {
                    Directory.CreateDirectory(ConfigDirectory);
                }

                var json = JsonConvert.SerializeObject(config, Formatting.Indented);
                File.WriteAllText(ConfigFilePath, json);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[JapanesePresetFormatManager] Save error: {ex.Message}");
            }
        }
    }
}
