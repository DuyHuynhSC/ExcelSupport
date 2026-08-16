using System;
using System.IO;
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

        public static AiConfig Current
        {
            get
            {
                if (_currentConfig == null)
                {
                    _currentConfig = Load();
                }
                return _currentConfig;
            }
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
                        return config;
                    }
                }
            }
            catch
            {
                // Fallback to default
            }

            return new AiConfig();
        }

        public static bool Save(AiConfig config)
        {
            try
            {
                if (!Directory.Exists(ConfigDirectory))
                {
                    Directory.CreateDirectory(ConfigDirectory);
                }

                string json = JsonConvert.SerializeObject(config, Formatting.Indented);
                File.WriteAllText(ConfigFilePath, json);
                _currentConfig = config;
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
