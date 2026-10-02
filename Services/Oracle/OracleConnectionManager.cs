using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelSupport.Helpers;
using ExcelSupport.Models;
using Newtonsoft.Json;

namespace ExcelSupport.Services
{
    public static class OracleConnectionManager
    {
        private const string ConfigFileName = "oracle_connections.json";

        private static List<OracleConnectionProfile>? _profiles;
        private static readonly object SyncLock = new object();

        public static event Action? ProfilesChanged;

        static OracleConnectionManager()
        {
            ProjectProfileManager.ActiveProfileChanged += RaiseProfilesChanged;
            ProjectProfileManager.ProfilesUpdated += RaiseProfilesChanged;
        }

        public static void RaiseProfilesChanged()
        {
            lock (SyncLock)
            {
                _profiles = null;
            }
            ProfilesChanged?.Invoke();
        }

        public static List<OracleConnectionProfile> GetProfiles()
        {
            lock (SyncLock)
            {
                var projectProfiles = ProjectProfileManager.CurrentConfig.Profiles;
                if (projectProfiles != null && projectProfiles.Count > 0)
                {
                    string? activeId = ProjectProfileManager.CurrentConfig.ActiveProfileId;
                    var list = new List<OracleConnectionProfile>();
                    foreach (var p in projectProfiles)
                    {
                        var db = p.DatabaseConnection;
                        db.Id = p.Id;
                        db.Name = p.Name;
                        db.IsDefault = (p.Id == activeId);
                        list.Add(db);
                    }
                    _profiles = list;
                    return list;
                }

                if (_profiles == null)
                {
                    _profiles = Load();
                }
                return _profiles;
            }
        }

        public static List<OracleConnectionProfile> Load()
        {
            var list = JsonConfigStore.Load(ConfigFileName, () =>
            {
                var p1 = new OracleConnectionProfile
                {
                    Name = "Localhost ORCL (Default)",
                    Host = "localhost",
                    Port = 1521,
                    ServiceNameOrSid = "ORCL",
                    ServiceType = OracleServiceNameType.ServiceName,
                    IsDefault = true
                };
                p1.EnsureDefaultUsers();

                var p2 = new OracleConnectionProfile
                {
                    Name = "Dev / UAT Environment",
                    Host = "192.168.1.100",
                    Port = 1521,
                    ServiceNameOrSid = "DEVDB",
                    ServiceType = OracleServiceNameType.ServiceName,
                    IsDefault = false
                };
                p2.EnsureDefaultUsers();

                return new List<OracleConnectionProfile> { p1, p2 };
            });

            foreach (var p in list)
            {
                p.EnsureDefaultUsers();
            }
            return list;
        }

        public static bool Save(List<OracleConnectionProfile> profiles)
        {
            lock (SyncLock)
            {
                bool success = JsonConfigStore.Save(ConfigFileName, profiles);
                if (success)
                {
                    var currentProjects = ProjectProfileManager.CurrentConfig.Profiles;
                    if (currentProjects != null && currentProjects.Count > 0)
                    {
                        foreach (var p in profiles)
                        {
                            var match = currentProjects.FirstOrDefault(cp => cp.Id == p.Id);
                            if (match != null)
                            {
                                match.DatabaseConnection = p;
                            }
                        }
                        ProjectProfileManager.SaveAllProfiles(currentProjects, ProjectProfileManager.CurrentConfig.ActiveProfileId);
                    }

                    _profiles = profiles;
                    ProfilesChanged?.Invoke();
                }
                return success;
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
            var activeProject = ProjectProfileManager.GetActiveProfile();
            if (activeProject != null)
            {
                var db = activeProject.DatabaseConnection;
                db.Id = activeProject.Id;
                db.Name = activeProject.Name;
                db.IsDefault = true;
                return db;
            }
            var list = GetProfiles();
            if (list.Count == 0) return null;
            return list.FirstOrDefault(p => p.IsDefault) ?? list.FirstOrDefault();
        }

        public static bool SetDefaultProfile(string profileId)
        {
            ProjectProfileManager.SetActiveProfile(profileId);
            return true;
        }

        #region Last Compare Session History

        private const string LastSessionFileName = "oracle_last_compare.json";

        public static OracleLastCompareSession? GetLastSession()
        {
            return JsonConfigStore.Load<OracleLastCompareSession>(LastSessionFileName, () => null!);
        }

        public static bool SaveLastSession(OracleLastCompareSession session)
        {
            return JsonConfigStore.Save(LastSessionFileName, session);
        }

        #endregion

        #region Query History

        private const string QueryHistoryFileName = "oracle_query_history.json";

        public static List<OracleQueryHistoryItem> GetQueryHistory()
        {
            return JsonConfigStore.Load(QueryHistoryFileName, () => new List<OracleQueryHistoryItem>());
        }

        public static bool AddQueryHistory(string sql, int rowCount, string? profileName)
        {
            if (string.IsNullOrWhiteSpace(sql)) return false;

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

            return JsonConfigStore.Save(QueryHistoryFileName, history);
        }

        public static bool ClearQueryHistory()
        {
            return JsonConfigStore.Delete(QueryHistoryFileName);
        }

        #endregion
    }
}
