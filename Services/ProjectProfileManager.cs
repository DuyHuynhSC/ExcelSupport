using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelSupport.Models;
using Newtonsoft.Json;

namespace ExcelSupport.Services
{
    /// <summary>
    /// Quản lý cấu hình danh sách các Profile dự án lưu trữ trong %APPDATA%\ExcelSupport\project_profiles.json
    /// </summary>
    public static class ProjectProfileManager
    {
        private static readonly string ConfigDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "ExcelSupport"
        );

        private static readonly string ConfigFilePath = Path.Combine(ConfigDirectory, "project_profiles.json");

        private static ProjectProfilesConfig? _currentConfig;
        private static readonly object SyncLock = new object();

        public static event Action? ActiveProfileChanged;
        public static event Action? ProfilesUpdated;

        public static ProjectProfilesConfig CurrentConfig
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

        public static ProjectProfile? GetActiveProfile()
        {
            var config = CurrentConfig;
            if (string.IsNullOrWhiteSpace(config.ActiveProfileId))
            {
                return config.Profiles.FirstOrDefault();
            }

            var active = config.Profiles.FirstOrDefault(p => p.Id == config.ActiveProfileId);
            return active ?? config.Profiles.FirstOrDefault();
        }

        public static void SetActiveProfile(string profileId)
        {
            lock (SyncLock)
            {
                var config = CurrentConfig;
                if (config.Profiles.Any(p => p.Id == profileId))
                {
                    config.ActiveProfileId = profileId;
                    SaveConfig(config);
                }
            }

            ActiveProfileChanged?.Invoke();
            ProfilesUpdated?.Invoke();
        }

        public static void SaveProfile(ProjectProfile profile)
        {
            lock (SyncLock)
            {
                var config = CurrentConfig;
                var existingIndex = config.Profiles.FindIndex(p => p.Id == profile.Id);

                profile.LastModified = DateTime.Now;

                if (existingIndex >= 0)
                {
                    config.Profiles[existingIndex] = profile;
                }
                else
                {
                    config.Profiles.Add(profile);
                }

                // Nếu chưa có active profile, chọn profile này làm active
                if (string.IsNullOrWhiteSpace(config.ActiveProfileId) || config.Profiles.Count == 1)
                {
                    config.ActiveProfileId = profile.Id;
                }

                SaveConfig(config);
            }

            ProfilesUpdated?.Invoke();
        }

        public static bool DeleteProfile(string profileId)
        {
            bool removed = false;
            lock (SyncLock)
            {
                var config = CurrentConfig;
                var profileToRemove = config.Profiles.FirstOrDefault(p => p.Id == profileId);
                if (profileToRemove != null)
                {
                    config.Profiles.Remove(profileToRemove);
                    removed = true;

                    if (config.ActiveProfileId == profileId)
                    {
                        config.ActiveProfileId = config.Profiles.FirstOrDefault()?.Id;
                        ActiveProfileChanged?.Invoke();
                    }

                    SaveConfig(config);
                }
            }

            if (removed)
            {
                ProfilesUpdated?.Invoke();
            }

            return removed;
        }

        public static bool ValidateFolderExists(string? path)
        {
            if (string.IsNullOrWhiteSpace(path)) return false;
            try
            {
                return Directory.Exists(path);
            }
            catch
            {
                return false;
            }
        }

        private static ProjectProfilesConfig LoadConfig()
        {
            try
            {
                if (!Directory.Exists(ConfigDirectory))
                {
                    Directory.CreateDirectory(ConfigDirectory);
                }

                if (File.Exists(ConfigFilePath))
                {
                    string json = File.ReadAllText(ConfigFilePath);
                    var config = JsonConvert.DeserializeObject<ProjectProfilesConfig>(json);
                    if (config != null && config.Profiles != null && config.Profiles.Count > 0)
                    {
                        return config;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ProjectProfileManager] Error loading config: {ex.Message}");
            }

            // Tạo cấu hình mặc định ban đầu
            var defaultConfig = new ProjectProfilesConfig
            {
                Profiles = new List<ProjectProfile>
                {
                    new ProjectProfile
                    {
                        Name = "Dự Án Mẫu (Sample Project)",
                        RootFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                        DetailedDesignFolder = "Detailed_Design",
                        BasicDesignFolder = "Basic_Design",
                        TestSpecFolder = "Test_Specification",
                        OpenReadOnlyDefault = false,
                        FileExtensions = ".xlsx;.xlsm;.xls;.docx;.pdf;.pptx"
                    }
                }
            };
            defaultConfig.ActiveProfileId = defaultConfig.Profiles[0].Id;

            SaveConfig(defaultConfig);
            return defaultConfig;
        }

        private static void SaveConfig(ProjectProfilesConfig config)
        {
            try
            {
                if (!Directory.Exists(ConfigDirectory))
                {
                    Directory.CreateDirectory(ConfigDirectory);
                }

                string json = JsonConvert.SerializeObject(config, Formatting.Indented);
                File.WriteAllText(ConfigFilePath, json);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ProjectProfileManager] Error saving config: {ex.Message}");
            }
        }
    }
}
