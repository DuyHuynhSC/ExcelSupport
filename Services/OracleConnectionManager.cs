using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelSupport.Models;
using Newtonsoft.Json;

namespace ExcelSupport.Services
{
    public static class OracleConnectionManager
    {
        private static readonly string ConfigDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "ExcelSupport"
        );

        private static readonly string ConfigFilePath = Path.Combine(ConfigDirectory, "oracle_connections.json");

        private static List<OracleConnectionProfile>? _profiles;
        private static readonly object SyncLock = new object();

        public static event Action? ProfilesChanged;

        public static List<OracleConnectionProfile> GetProfiles()
        {
            lock (SyncLock)
            {
                if (_profiles == null)
                {
                    _profiles = Load();
                }
                return _profiles;
            }
        }

        public static List<OracleConnectionProfile> Load()
        {
            try
            {
                if (File.Exists(ConfigFilePath))
                {
                    string json = File.ReadAllText(ConfigFilePath);
                    var list = JsonConvert.DeserializeObject<List<OracleConnectionProfile>>(json);
                    if (list != null && list.Count > 0)
                    {
                        return list;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[OracleConnectionManager] Load error: {ex.Message}");
            }

            // Default sample profiles if none exist
            var defaults = new List<OracleConnectionProfile>
            {
                new OracleConnectionProfile
                {
                    Name = "Localhost ORCL (Default)",
                    Host = "localhost",
                    Port = 1521,
                    ServiceNameOrSid = "ORCL",
                    ServiceType = OracleServiceNameType.ServiceName,
                    Username = "SYSTEM",
                    Password = ""
                },
                new OracleConnectionProfile
                {
                    Name = "Dev / UAT Environment",
                    Host = "192.168.1.100",
                    Port = 1521,
                    ServiceNameOrSid = "DEVDB",
                    ServiceType = OracleServiceNameType.ServiceName,
                    Username = "APP_USER",
                    Password = ""
                }
            };

            Save(defaults);
            return defaults;
        }

        public static bool Save(List<OracleConnectionProfile> profiles)
        {
            lock (SyncLock)
            {
                try
                {
                    if (!Directory.Exists(ConfigDirectory))
                    {
                        Directory.CreateDirectory(ConfigDirectory);
                    }

                    string json = JsonConvert.SerializeObject(profiles, Formatting.Indented);
                    File.WriteAllText(ConfigFilePath, json);
                    _profiles = profiles;

                    ProfilesChanged?.Invoke();
                    return true;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[OracleConnectionManager] Save error: {ex.Message}");
                    return false;
                }
            }
        }

        public static bool AddOrUpdateProfile(OracleConnectionProfile profile)
        {
            var list = GetProfiles().ToList();
            int idx = list.FindIndex(p => p.Id == profile.Id);
            if (idx >= 0)
            {
                list[idx] = profile;
            }
            else
            {
                list.Add(profile);
            }
            return Save(list);
        }

        public static bool DeleteProfile(string profileId)
        {
            var list = GetProfiles().ToList();
            int removed = list.RemoveAll(p => p.Id == profileId);
            if (removed > 0)
            {
                return Save(list);
            }
            return false;
        }

        public static OracleConnectionProfile? GetDefaultProfile()
        {
            var list = GetProfiles();
            if (list.Count == 0) return null;
            return list.FirstOrDefault(p => p.IsDefault) ?? list.FirstOrDefault();
        }

        public static bool SetDefaultProfile(string profileId)
        {
            var list = GetProfiles().ToList();
            bool found = false;
            foreach (var p in list)
            {
                if (p.Id == profileId)
                {
                    p.IsDefault = true;
                    found = true;
                }
                else
                {
                    p.IsDefault = false;
                }
            }
            if (found)
            {
                return Save(list);
            }
            return false;
        }

        #region Last Compare Session History

        private static readonly string LastSessionFilePath = Path.Combine(ConfigDirectory, "oracle_last_compare.json");

        public static OracleLastCompareSession? GetLastSession()
        {
            try
            {
                if (File.Exists(LastSessionFilePath))
                {
                    string json = File.ReadAllText(LastSessionFilePath);
                    return JsonConvert.DeserializeObject<OracleLastCompareSession>(json);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[OracleConnectionManager] GetLastSession error: {ex.Message}");
            }
            return null;
        }

        public static bool SaveLastSession(OracleLastCompareSession session)
        {
            try
            {
                if (!Directory.Exists(ConfigDirectory))
                {
                    Directory.CreateDirectory(ConfigDirectory);
                }
                string json = JsonConvert.SerializeObject(session, Formatting.Indented);
                File.WriteAllText(LastSessionFilePath, json);
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[OracleConnectionManager] SaveLastSession error: {ex.Message}");
                return false;
            }
        }

        #endregion

        #region Query History

        private static readonly string QueryHistoryFilePath = Path.Combine(ConfigDirectory, "oracle_query_history.json");

        public static List<OracleQueryHistoryItem> GetQueryHistory()
        {
            try
            {
                if (File.Exists(QueryHistoryFilePath))
                {
                    string json = File.ReadAllText(QueryHistoryFilePath);
                    var list = JsonConvert.DeserializeObject<List<OracleQueryHistoryItem>>(json);
                    if (list != null) return list;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[OracleConnectionManager] GetQueryHistory error: {ex.Message}");
            }
            return new List<OracleQueryHistoryItem>();
        }

        public static bool AddQueryHistory(string sql, int rowCount, string? profileName)
        {
            if (string.IsNullOrWhiteSpace(sql)) return false;

            try
            {
                var history = GetQueryHistory();

                // Remove existing identical SQL to bring it to top
                string cleanSql = sql.Trim();
                history.RemoveAll(h => string.Equals(h.Sql?.Trim(), cleanSql, StringComparison.OrdinalIgnoreCase));

                history.Insert(0, new OracleQueryHistoryItem
                {
                    Sql = cleanSql,
                    ExecutedAt = DateTime.Now,
                    RowCount = rowCount,
                    ProfileName = profileName
                });

                // Keep up to 30 recent queries
                if (history.Count > 30)
                {
                    history = history.Take(30).ToList();
                }

                if (!Directory.Exists(ConfigDirectory))
                {
                    Directory.CreateDirectory(ConfigDirectory);
                }
                string json = JsonConvert.SerializeObject(history, Formatting.Indented);
                File.WriteAllText(QueryHistoryFilePath, json);
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[OracleConnectionManager] AddQueryHistory error: {ex.Message}");
                return false;
            }
        }

        public static bool ClearQueryHistory()
        {
            try
            {
                if (File.Exists(QueryHistoryFilePath))
                {
                    File.Delete(QueryHistoryFilePath);
                }
                return true;
            }
            catch { return false; }
        }

        #endregion
    }
}
