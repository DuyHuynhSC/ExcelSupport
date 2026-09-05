using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using ExcelSupport.Models;
using Newtonsoft.Json;

namespace ExcelSupport.Services
{
    public enum AppLanguage
    {
        Vietnamese,
        English,
        Japanese
    }

    public class LocalizationService : INotifyPropertyChanged
    {
        private static LocalizationService? _instance;
        public static LocalizationService Instance => _instance ??= new LocalizationService();

        private static AppLanguage _currentLanguage = LoadSavedLanguage();

        private static AppLanguage LoadSavedLanguage()
        {
            try
            {
                string saved = Host.AppSettings.CurrentLanguage;
                return saved?.ToLowerInvariant() switch
                {
                    "en" => AppLanguage.English,
                    "ja" => AppLanguage.Japanese,
                    _ => AppLanguage.Vietnamese
                };
            }
            catch
            {
                return AppLanguage.Vietnamese;
            }
        }

        public static AppLanguage CurrentLanguage
        {
            get => _currentLanguage;
            set
            {
                if (_currentLanguage != value)
                {
                    _currentLanguage = value;

                    try
                    {
                        Host.AppSettings.CurrentLanguage = value switch
                        {
                            AppLanguage.English => "en",
                            AppLanguage.Japanese => "ja",
                            _ => "vi"
                        };
                    }
                    catch { }

                    LanguageChanged?.Invoke(_currentLanguage);
                    Instance.OnPropertyChanged(string.Empty);
                }
            }
        }

        public static event Action<AppLanguage>? LanguageChanged;

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public string this[string key] => Get(key);

        public static string Get(string key, params object[] args)
        {
            if (string.IsNullOrEmpty(key)) return string.Empty;

            EnsureLoaded();

            if (_translations.TryGetValue(CurrentLanguage, out var dict) && dict.TryGetValue(key, out var text))
            {
                return args.Length > 0 ? string.Format(text, args) : text;
            }

            // Fallback sang tiếng Việt nếu key chưa được dịch ở ngôn ngữ hiện hành
            if (CurrentLanguage != AppLanguage.Vietnamese &&
                _translations.TryGetValue(AppLanguage.Vietnamese, out var viDict) &&
                viDict.TryGetValue(key, out var viText))
            {
                return args.Length > 0 ? string.Format(viText, args) : viText;
            }

            return key;
        }

        public static string GetLabel(string controlId) => Get(controlId);
        public static string GetScreentip(string controlId) => Get(controlId + "_Tip") != (controlId + "_Tip") ? Get(controlId + "_Tip") : Get(controlId);
        public static string GetSupertip(string controlId) => Get(controlId + "_SuperTip") != (controlId + "_SuperTip") ? Get(controlId + "_SuperTip") : Get(controlId);

        public static string GetOperatorDescription(FilterOperator op)
        {
            string key = "Op_" + op.ToString();
            string trans = Get(key);
            return trans != key ? trans : op.ToString();
        }

        private static readonly Dictionary<AppLanguage, Dictionary<string, string>> _translations = new();
        private static bool _isLoaded;
        private static readonly object _loadLock = new();

        private static void EnsureLoaded()
        {
            if (_isLoaded) return;
            lock (_loadLock)
            {
                if (_isLoaded) return;

                LoadLanguageFromResource(AppLanguage.Vietnamese, "ExcelSupport.Languages.vi.json");
                LoadLanguageFromResource(AppLanguage.English, "ExcelSupport.Languages.en.json");
                LoadLanguageFromResource(AppLanguage.Japanese, "ExcelSupport.Languages.ja.json");

                _isLoaded = true;
            }
        }

        private static void LoadLanguageFromResource(AppLanguage lang, string resourceName)
        {
            try
            {
                var asm = Assembly.GetExecutingAssembly();
                using var stream = asm.GetManifestResourceStream(resourceName);
                if (stream != null)
                {
                    using var reader = new StreamReader(stream, System.Text.Encoding.UTF8);
                    string json = reader.ReadToEnd();
                    var dict = JsonConvert.DeserializeObject<Dictionary<string, string>>(json);
                    if (dict != null)
                    {
                        _translations[lang] = new Dictionary<string, string>(dict, StringComparer.Ordinal);
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading language resource '{resourceName}': {ex.Message}");
            }

            if (!_translations.ContainsKey(lang))
            {
                _translations[lang] = new Dictionary<string, string>(StringComparer.Ordinal);
            }
        }
    }
}
