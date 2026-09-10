using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelSupport.Models;
using Newtonsoft.Json;

namespace ExcelSupport.Services
{
    public static class AiConfigManager
    {
        private static readonly string ConfigDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "ExcelSupport"
        );

        private static readonly string ConfigFilePath = Path.Combine(ConfigDirectory, "ai_config.json");

        private static AiConfig? _currentConfig;
        private static readonly object SyncLock = new object();

        public static event Action? ProfilesChanged;

        public static AiConfig Current
        {
            get
            {
                lock (SyncLock)
                {
                    if (_currentConfig == null)
                    {
                        _currentConfig = Load();
                    }
                    return _currentConfig;
                }
            }
        }

        public static List<AiConnectionProfile> GetDefaultPresetProfiles()
        {
            return new List<AiConnectionProfile>
            {
                new AiConnectionProfile
                {
                    Name = "Localhost Qwen / vLLM",
                    BaseUrl = "http://localhost:8000/v1",
                    ApiKey = "",
                    ModelName = "qwen-3.6",
                    TimeoutSeconds = 30,
                    Temperature = 0.3,
                    MaxTokens = 2048,
                    IsDefault = true
                },
                new AiConnectionProfile
                {
                    Name = "OpenAI (Official)",
                    BaseUrl = "https://api.openai.com/v1",
                    ApiKey = "",
                    ModelName = "gpt-4o-mini",
                    TimeoutSeconds = 60,
                    Temperature = 0.3,
                    MaxTokens = 2048,
                    IsDefault = false
                },
                new AiConnectionProfile
                {
                    Name = "DeepSeek API",
                    BaseUrl = "https://api.deepseek.com/v1",
                    ApiKey = "",
                    ModelName = "deepseek-chat",
                    TimeoutSeconds = 60,
                    Temperature = 0.3,
                    MaxTokens = 2048,
                    IsDefault = false
                },
                new AiConnectionProfile
                {
                    Name = "Ollama Localhost",
                    BaseUrl = "http://localhost:11434/v1",
                    ApiKey = "",
                    ModelName = "llama3.2",
                    TimeoutSeconds = 60,
                    Temperature = 0.3,
                    MaxTokens = 2048,
                    IsDefault = false
                },
                new AiConnectionProfile
                {
                    Name = "OpenRouter",
                    BaseUrl = "https://openrouter.ai/api/v1",
                    ApiKey = "",
                    ModelName = "openai/gpt-4o-mini",
                    TimeoutSeconds = 60,
                    Temperature = 0.3,
                    MaxTokens = 2048,
                    IsDefault = false
                }
            };
        }

        public static AiConfig Load()
        {
            try
            {
                if (File.Exists(ConfigFilePath))
                {
                    string json = File.ReadAllText(ConfigFilePath);
                    var config = JsonConvert.DeserializeObject<AiConfig>(json);
                    if (config != null)
                    {
                        EnsureProfiles(config);
                        return config;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[AiConfigManager] Load error: {ex.Message}");
            }

            var defaultConfig = new AiConfig();
            EnsureProfiles(defaultConfig);
            Save(defaultConfig);
            return defaultConfig;
        }

        private static void EnsureProfiles(AiConfig config)
        {
            if (config.Profiles == null || config.Profiles.Count == 0)
            {
                var defaults = GetDefaultPresetProfiles();

                // If user already had custom BaseUrl/ModelName/ApiKey in legacy config, preserve it as default profile
                if (!string.IsNullOrWhiteSpace(config.BaseUrl) && config.BaseUrl != "http://localhost:8000/v1")
                {
                    defaults[0].Name = "Custom Server (Mặc định)";
                    defaults[0].BaseUrl = config.BaseUrl;
                    defaults[0].ApiKey = config.ApiKey ?? "";
                    defaults[0].ModelName = string.IsNullOrWhiteSpace(config.ModelName) ? "qwen-3.6" : config.ModelName;
                    defaults[0].TimeoutSeconds = config.TimeoutSeconds > 0 ? config.TimeoutSeconds : 30;
                    defaults[0].Temperature = config.Temperature;
                }
                config.Profiles = defaults;
            }

            // Ensure exactly one default profile
            var defaultProf = config.Profiles.FirstOrDefault(p => p.IsDefault);
            if (defaultProf == null && config.Profiles.Count > 0)
            {
                config.Profiles[0].IsDefault = true;
                defaultProf = config.Profiles[0];
            }

            // Sync active profile
            var activeProf = config.Profiles.FirstOrDefault(p => p.Id == config.ActiveProfileId) ?? defaultProf;
            if (activeProf != null)
            {
                config.ActiveProfileId = activeProf.Id;
                config.BaseUrl = activeProf.BaseUrl;
                config.ApiKey = activeProf.ApiKey;
                config.ModelName = activeProf.ModelName;
                config.TimeoutSeconds = activeProf.TimeoutSeconds;
                config.Temperature = activeProf.Temperature;
                config.MaxTokens = activeProf.MaxTokens;
            }
        }

        public static bool Save(AiConfig config)
        {
            lock (SyncLock)
            {
                try
                {
                    if (!Directory.Exists(ConfigDirectory))
                    {
                        Directory.CreateDirectory(ConfigDirectory);
                    }

                    EnsureProfiles(config);

                    string json = JsonConvert.SerializeObject(config, Formatting.Indented);
                    File.WriteAllText(ConfigFilePath, json);
                    _currentConfig = config;

                    ProfilesChanged?.Invoke();
                    return true;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[AiConfigManager] Save error: {ex.Message}");
                    return false;
                }
            }
        }

        public static List<AiConnectionProfile> GetProfiles()
        {
            return Current.Profiles ?? new List<AiConnectionProfile>();
        }

        public static AiConnectionProfile? GetDefaultProfile()
        {
            var list = GetProfiles();
            return list.FirstOrDefault(p => p.IsDefault) ?? list.FirstOrDefault();
        }

        public static AiConnectionProfile? GetActiveProfile()
        {
            var config = Current;
            var list = config.Profiles ?? new List<AiConnectionProfile>();
            if (!string.IsNullOrEmpty(config.ActiveProfileId))
            {
                var match = list.FirstOrDefault(p => p.Id == config.ActiveProfileId);
                if (match != null) return match;
            }
            return GetDefaultProfile();
        }

        public static bool SetActiveProfile(string profileId)
        {
            var config = Current;
            var target = config.Profiles.FirstOrDefault(p => p.Id == profileId);
            if (target != null)
            {
                config.ActiveProfileId = target.Id;
                config.BaseUrl = target.BaseUrl;
                config.ApiKey = target.ApiKey;
                config.ModelName = target.ModelName;
                config.TimeoutSeconds = target.TimeoutSeconds;
                config.Temperature = target.Temperature;
                config.MaxTokens = target.MaxTokens;
                return Save(config);
            }
            return false;
        }

        public static bool SetDefaultProfile(string profileId)
        {
            var config = Current;
            bool found = false;
            foreach (var p in config.Profiles)
            {
                if (p.Id == profileId)
                {
                    p.IsDefault = true;
                    found = true;
                    config.ActiveProfileId = p.Id;
                    config.BaseUrl = p.BaseUrl;
                    config.ApiKey = p.ApiKey;
                    config.ModelName = p.ModelName;
                    config.TimeoutSeconds = p.TimeoutSeconds;
                    config.Temperature = p.Temperature;
                    config.MaxTokens = p.MaxTokens;
                }
                else
                {
                    p.IsDefault = false;
                }
            }

            if (found)
            {
                return Save(config);
            }
            return false;
        }

        public static bool AddOrUpdateProfile(AiConnectionProfile profile)
        {
            var config = Current;
            int idx = config.Profiles.FindIndex(p => p.Id == profile.Id);
            if (idx >= 0)
            {
                config.Profiles[idx] = profile;
            }
            else
            {
                config.Profiles.Add(profile);
            }
            return Save(config);
        }

        public static bool DeleteProfile(string profileId)
        {
            var config = Current;
            int removed = config.Profiles.RemoveAll(p => p.Id == profileId);
            if (removed > 0)
            {
                if (config.Profiles.Count == 0)
                {
                    config.Profiles = GetDefaultPresetProfiles();
                }
                else if (!config.Profiles.Any(p => p.IsDefault))
                {
                    config.Profiles[0].IsDefault = true;
                }
                return Save(config);
            }
            return false;
        }
    }
}
