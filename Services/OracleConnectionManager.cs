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
    }
}
