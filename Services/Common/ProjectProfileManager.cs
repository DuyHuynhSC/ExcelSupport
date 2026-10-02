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
                    foreach (var p in config.Profiles)
                    {
                        p.IsActive = (p.Id == profileId);
                        if (p.DatabaseConnection != null)
                        {
                            p.DatabaseConnection.IsDefault = (p.Id == profileId);
                        }
                    }
                    SaveConfig(config);
                }
            }

            ActiveProfileChanged?.Invoke();
            ProfilesUpdated?.Invoke();
            OracleConnectionManager.RaiseProfilesChanged();
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

                foreach (var p in config.Profiles)
                {
                    p.IsActive = (p.Id == config.ActiveProfileId);
                    if (p.DatabaseConnection != null)
                    {
                        p.DatabaseConnection.Id = p.Id;
                        p.DatabaseConnection.Name = p.Name;
                        p.DatabaseConnection.IsDefault = (p.Id == config.ActiveProfileId);
                    }
                }

                SaveConfig(config);
            }

            ActiveProfileChanged?.Invoke();
            ProfilesUpdated?.Invoke();
            OracleConnectionManager.RaiseProfilesChanged();
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

                foreach (var p in config.Profiles)
                {
                    p.IsActive = (p.Id == config.ActiveProfileId);
                    if (p.DatabaseConnection != null)
                    {
                        p.DatabaseConnection.Id = p.Id;
                        p.DatabaseConnection.Name = p.Name;
                        p.DatabaseConnection.IsDefault = (p.Id == config.ActiveProfileId);
                    }
                }

                SaveConfig(config);
            }

            ProfilesUpdated?.Invoke();
            OracleConnectionManager.RaiseProfilesChanged();
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
                        // Ensure database connection is initialized for all profiles
                        int pIdx = 1;
                        foreach (var p in config.Profiles)
                        {
                            if (string.IsNullOrWhiteSpace(p.Name))
                            {
                                p.Name = $"Dự Án {pIdx}";
                            }
                            pIdx++;

                            if (p.DatabaseConnection == null)
                            {
                                p.DatabaseConnection = new OracleConnectionProfile
                                {
                                    Id = p.Id,
                                    Name = p.Name,
                                    Host = "localhost",
                                    Port = 1521,
                                    ServiceNameOrSid = "ORCL",
                                    ServiceType = OracleServiceNameType.ServiceName
                                };
                            }
                            else
                            {
                                p.DatabaseConnection.Id = p.Id;
                                p.DatabaseConnection.Name = p.Name;
                            }
                            p.DatabaseConnection.EnsureDefaultUsers();
                            p.IsActive = (p.Id == config.ActiveProfileId);
                            p.DatabaseConnection.IsDefault = (p.Id == config.ActiveProfileId);
                        }

                        TryMigrateLegacyOracleConnections(config);
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
            defaultConfig.Profiles[0].IsActive = true;
            defaultConfig.Profiles[0].DatabaseConnection.Id = defaultConfig.Profiles[0].Id;
            defaultConfig.Profiles[0].DatabaseConnection.Name = defaultConfig.Profiles[0].Name;
            defaultConfig.Profiles[0].DatabaseConnection.IsDefault = true;

            TryMigrateLegacyOracleConnections(defaultConfig);

            SaveConfig(defaultConfig);
            return defaultConfig;
        }

        private static void TryMigrateLegacyOracleConnections(ProjectProfilesConfig config)
        {
            try
            {
                string oracleJsonPath = Path.Combine(ConfigDirectory, "oracle_connections.json");
                if (File.Exists(oracleJsonPath) && config.Profiles.Count > 0)
                {
                    string json = File.ReadAllText(oracleJsonPath);
                    var legacyList = JsonConvert.DeserializeObject<List<OracleConnectionProfile>>(json);
                    if (legacyList != null && legacyList.Count > 0)
                    {
                        var firstProj = config.Profiles[0];
                        // If first project's DB is untouched/default, import legacy settings
                        if (firstProj.DatabaseConnection.Users.Count <= 1 &&
                            (firstProj.DatabaseConnection.Users.Count == 0 || firstProj.DatabaseConnection.Users[0].Username == "XXX_USR1" || firstProj.DatabaseConnection.Users[0].Username == ""))
                        {
                            var legacy = legacyList[0];
                            firstProj.DatabaseConnection.Host = legacy.Host;
                            firstProj.DatabaseConnection.Port = legacy.Port;
                            firstProj.DatabaseConnection.ServiceNameOrSid = legacy.ServiceNameOrSid;
                            firstProj.DatabaseConnection.ServiceType = legacy.ServiceType;
                            firstProj.DatabaseConnection.Users = legacy.Users;
                            firstProj.DatabaseConnection.SelectedUserId = legacy.SelectedUserId;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ProjectProfileManager] Migration error: {ex.Message}");
            }
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
