using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelSupport.Helpers;
using ExcelSupport.Models;
using Newtonsoft.Json;

namespace ExcelSupport.Services
{
    /// <summary>
    /// Quản lý cấu hình danh sách các Profile dự án lưu trữ trong %APPDATA%\ExcelSupport\project_profiles.json
    /// </summary>
    public static class ProjectProfileManager
    {
        private const string ConfigFileName = "project_profiles.json";

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

        public static void SaveAllProfiles(List<ProjectProfile> profiles, string? activeProfileId)
        {
            lock (SyncLock)
            {
                var config = CurrentConfig;
                config.Profiles = new List<ProjectProfile>(profiles);
                config.ActiveProfileId = !string.IsNullOrWhiteSpace(activeProfileId)
                    ? activeProfileId
                    : config.Profiles.FirstOrDefault()?.Id;
                SaveConfig(config);
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
            var config = JsonConfigStore.Load(ConfigFileName, () =>
            {
                var def = new ProjectProfilesConfig
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
                def.ActiveProfileId = def.Profiles[0].Id;
                return def;
            });

            if (config.Profiles == null || config.Profiles.Count == 0)
            {
                config.Profiles = new List<ProjectProfile>
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
                };
                config.ActiveProfileId = config.Profiles[0].Id;
            }

            return config;
        }

        private static void SaveConfig(ProjectProfilesConfig config)
        {
            JsonConfigStore.Save(ConfigFileName, config);
        }
    }
}
