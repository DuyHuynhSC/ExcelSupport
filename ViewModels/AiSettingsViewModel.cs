using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Input;
using ExcelSupport.Models;
using ExcelSupport.Services;
using WpfMessageBox = System.Windows.MessageBox;
using WpfMessageBoxButton = System.Windows.MessageBoxButton;
using WpfMessageBoxImage = System.Windows.MessageBoxImage;
using WpfMessageBoxResult = System.Windows.MessageBoxResult;

namespace ExcelSupport.ViewModels
{
    public class AiSettingsViewModel : ViewModelBase
    {
        private ObservableCollection<AiConnectionProfile> _profiles = new ObservableCollection<AiConnectionProfile>();
        private AiConnectionProfile? _selectedProfile;

        private bool _isTesting;
        private string _testStatusMessage = string.Empty;
        private bool? _isTestSuccess;
        private string _latencyText = string.Empty;
        private string _saveNotification = string.Empty;

        public ObservableCollection<AiConnectionProfile> Profiles
        {
            get => _profiles;
            set => SetProperty(ref _profiles, value);
        }

        public AiConnectionProfile? SelectedProfile
        {
            get => _selectedProfile;
            set
            {
                if (SetProperty(ref _selectedProfile, value))
                {
                    OnPropertyChanged(nameof(HasSelectedProfile));
                    OnPropertyChanged(nameof(ProfileName));
                    OnPropertyChanged(nameof(BaseUrl));
                    OnPropertyChanged(nameof(ApiKey));
                    OnPropertyChanged(nameof(ModelName));
                    OnPropertyChanged(nameof(TimeoutSeconds));
                    OnPropertyChanged(nameof(Temperature));
                    OnPropertyChanged(nameof(MaxTokens));

                    TestStatusMessage = string.Empty;
                    IsTestSuccess = null;
                    SaveNotification = string.Empty;
                }
            }
        }

        public bool HasSelectedProfile => SelectedProfile != null;

        public string ProfileName
        {
            get => SelectedProfile?.Name ?? string.Empty;
            set
            {
                if (SelectedProfile != null && SelectedProfile.Name != value)
                {
                    SelectedProfile.Name = value;
                    OnPropertyChanged();
                }
            }
        }

        public string BaseUrl
        {
            get => SelectedProfile?.BaseUrl ?? string.Empty;
            set
            {
                if (SelectedProfile != null && SelectedProfile.BaseUrl != value)
                {
                    SelectedProfile.BaseUrl = value;
                    OnPropertyChanged();
                }
            }
        }

        public string ApiKey
        {
            get => SelectedProfile?.ApiKey ?? string.Empty;
            set
            {
                if (SelectedProfile != null && SelectedProfile.ApiKey != value)
                {
                    SelectedProfile.ApiKey = value;
                    OnPropertyChanged();
                }
            }
        }

        public string ModelName
        {
            get => SelectedProfile?.ModelName ?? string.Empty;
            set
            {
                if (SelectedProfile != null && SelectedProfile.ModelName != value)
                {
                    SelectedProfile.ModelName = value;
                    OnPropertyChanged();
                }
            }
        }

        public int TimeoutSeconds
        {
            get => SelectedProfile?.TimeoutSeconds ?? 30;
            set
            {
                if (SelectedProfile != null && SelectedProfile.TimeoutSeconds != value)
                {
                    SelectedProfile.TimeoutSeconds = value;
                    OnPropertyChanged();
                }
            }
        }

        public double Temperature
        {
            get => SelectedProfile?.Temperature ?? 0.3;
            set
            {
                if (SelectedProfile != null && Math.Abs(SelectedProfile.Temperature - value) > 0.001)
                {
                    SelectedProfile.Temperature = value;
                    OnPropertyChanged();
                }
            }
        }

        public int MaxTokens
        {
            get => SelectedProfile?.MaxTokens ?? 2048;
            set
            {
                if (SelectedProfile != null && SelectedProfile.MaxTokens != value)
                {
                    SelectedProfile.MaxTokens = value;
                    OnPropertyChanged();
                }
            }
        }

        public bool IsTesting
        {
            get => _isTesting;
            private set => SetProperty(ref _isTesting, value);
        }

        public string TestStatusMessage
        {
            get => _testStatusMessage;
            private set
            {
                if (SetProperty(ref _testStatusMessage, value))
                {
                    OnPropertyChanged(nameof(HasTestStatus));
                }
            }
        }

        public bool HasTestStatus => !string.IsNullOrWhiteSpace(TestStatusMessage);

        public bool? IsTestSuccess
        {
            get => _isTestSuccess;
            private set => SetProperty(ref _isTestSuccess, value);
        }

        public string LatencyText
        {
            get => _latencyText;
            private set => SetProperty(ref _latencyText, value);
        }

        public string SaveNotification
        {
            get => _saveNotification;
            private set => SetProperty(ref _saveNotification, value);
        }

        public ICommand TestConnectionCommand { get; }
        public ICommand SaveSettingsCommand { get; }
        public ICommand ResetDefaultsCommand { get; }
        public ICommand AddProfileCommand { get; }
        public ICommand CloneProfileCommand { get; }
        public ICommand DeleteProfileCommand { get; }
        public ICommand SetDefaultProfileCommand { get; }
        public ICommand ApplyPresetCommand { get; }

        public AiSettingsViewModel()
        {
            TestConnectionCommand = new RelayCommand(async _ => await ExecuteTestConnectionAsync(), _ => !IsTesting && HasSelectedProfile);
            SaveSettingsCommand = new RelayCommand(_ => ExecuteSaveSettings());
            ResetDefaultsCommand = new RelayCommand(_ => ExecuteResetDefaults());
            AddProfileCommand = new RelayCommand(_ => ExecuteAddProfile());
            CloneProfileCommand = new RelayCommand(_ => ExecuteCloneProfile(), _ => SelectedProfile != null);
            DeleteProfileCommand = new RelayCommand(_ => ExecuteDeleteProfile(), _ => SelectedProfile != null);
            SetDefaultProfileCommand = new RelayCommand(_ => ExecuteSetDefaultProfile(), _ => SelectedProfile != null);
            ApplyPresetCommand = new RelayCommand(preset => ExecuteApplyPreset(preset?.ToString()), _ => SelectedProfile != null);

            ReloadProfiles();
        }

        public void ReloadProfiles()
        {
            var list = AiConfigManager.GetProfiles();

            _profiles.Clear();
            foreach (var p in list)
            {
                _profiles.Add(p);
            }

            var active = AiConfigManager.GetActiveProfile();
            if (active != null)
            {
                var match = _profiles.FirstOrDefault(p => p.Id == active.Id);
                SelectedProfile = match ?? _profiles.FirstOrDefault();
            }
            else
            {
                SelectedProfile = _profiles.FirstOrDefault();
            }
        }

        private void ExecuteAddProfile()
        {
            var newProfile = new AiConnectionProfile
            {
                Name = $"AI Provider {_profiles.Count + 1}",
                BaseUrl = "http://localhost:8000/v1",
                ApiKey = "",
                ModelName = "qwen-3.6",
                TimeoutSeconds = 30,
                Temperature = 0.3,
                MaxTokens = 2048,
                IsDefault = _profiles.Count == 0
            };

            _profiles.Add(newProfile);
            SelectedProfile = newProfile;
            SaveNotification = "Đã tạo cấu hình mới. Hãy nhập thông số và bấm Lưu Cấu Hình.";
        }

        private void ExecuteCloneProfile()
        {
            if (SelectedProfile == null) return;

            var clone = SelectedProfile.Clone();
            _profiles.Add(clone);
            SelectedProfile = clone;
            SaveNotification = $"Đã nhân bản '{clone.Name}'.";
        }

        private void ExecuteDeleteProfile()
        {
            if (SelectedProfile == null) return;

            string name = SelectedProfile.Name;
            var confirm = WpfMessageBox.Show(
                string.Format(LocalizationService.Get("AiSet_DeleteConfirm"), name),
                LocalizationService.Get("AiSet_DeleteTitle"),
                WpfMessageBoxButton.YesNo,
                WpfMessageBoxImage.Question);

            if (confirm == WpfMessageBoxResult.Yes)
            {
                string idToDelete = SelectedProfile.Id;
                _profiles.Remove(SelectedProfile);

                if (_profiles.Count == 0)
                {
                    ExecuteAddProfile();
                }
                else
                {
                    if (!_profiles.Any(p => p.IsDefault))
                    {
                        _profiles[0].IsDefault = true;
                    }
                    SelectedProfile = _profiles.FirstOrDefault();
                }

                ExecuteSaveSettings();
            }
        }

        private void ExecuteSetDefaultProfile()
        {
            if (SelectedProfile == null) return;

            foreach (var p in _profiles)
            {
                p.IsDefault = (p.Id == SelectedProfile.Id);
            }

            SaveNotification = $"⭐ Đã đặt '{SelectedProfile.Name}' làm cấu hình mặc định.";
            ExecuteSaveSettings();
        }

        private void ExecuteApplyPreset(string? preset)
        {
            if (string.IsNullOrWhiteSpace(preset)) return;

            switch (preset!.ToLowerInvariant())
            {
                case "openai":
                    ProfileName = "OpenAI (Official)";
                    BaseUrl = "https://api.openai.com/v1";
                    ModelName = "gpt-4o-mini";
                    TimeoutSeconds = 60;
                    Temperature = 0.3;
                    break;

                case "deepseek":
                    ProfileName = "DeepSeek API";
                    BaseUrl = "https://api.deepseek.com/v1";
                    ModelName = "deepseek-chat";
                    TimeoutSeconds = 60;
                    Temperature = 0.3;
                    break;

                case "ollama":
                    ProfileName = "Ollama Localhost";
                    BaseUrl = "http://localhost:11434/v1";
                    ModelName = "llama3.2";
                    TimeoutSeconds = 60;
                    Temperature = 0.3;
                    break;

                case "localhost":
                case "qwen":
                    ProfileName = "Localhost Qwen / vLLM";
                    BaseUrl = "http://localhost:8000/v1";
                    ModelName = "qwen-3.6";
                    TimeoutSeconds = 30;
                    Temperature = 0.3;
                    break;

                case "openrouter":
                    ProfileName = "OpenRouter";
                    BaseUrl = "https://openrouter.ai/api/v1";
                    ModelName = "openai/gpt-4o-mini";
                    TimeoutSeconds = 60;
                    Temperature = 0.3;
                    break;
            }

            SaveNotification = $"Đã nạp mẫu '{ProfileName}'.";
        }

        private async Task ExecuteTestConnectionAsync()
        {
            IsTesting = true;
            IsTestSuccess = null;
            TestStatusMessage = "Đang kết nối tới máy chủ AI...";
            LatencyText = string.Empty;
            SaveNotification = string.Empty;

            var testConfig = new AiConfig
            {
                BaseUrl = BaseUrl?.Trim() ?? string.Empty,
                ApiKey = ApiKey?.Trim() ?? string.Empty,
                ModelName = string.IsNullOrWhiteSpace(ModelName) ? "qwen-3.6" : ModelName.Trim(),
                TimeoutSeconds = TimeoutSeconds <= 0 ? 30 : TimeoutSeconds,
                Temperature = Temperature
            };

            var result = await Task.Run(() => OpenAiClientService.TestConnectionAsync(testConfig));

            IsTesting = false;
            IsTestSuccess = result.IsSuccess;
            TestStatusMessage = result.Message;
            if (result.IsSuccess)
            {
                LatencyText = $"⚡ {result.LatencyMs} ms";
            }
        }

        private void ExecuteSaveSettings()
        {
            if (SelectedProfile != null)
            {
                SelectedProfile.Name = string.IsNullOrWhiteSpace(ProfileName) ? "AI Provider" : ProfileName.Trim();
                SelectedProfile.BaseUrl = BaseUrl?.Trim() ?? string.Empty;
                SelectedProfile.ApiKey = ApiKey?.Trim() ?? string.Empty;
                SelectedProfile.ModelName = string.IsNullOrWhiteSpace(ModelName) ? "qwen-3.6" : ModelName.Trim();
                SelectedProfile.TimeoutSeconds = TimeoutSeconds <= 0 ? 30 : TimeoutSeconds;
                SelectedProfile.Temperature = Temperature;
                SelectedProfile.MaxTokens = MaxTokens <= 0 ? 2048 : MaxTokens;
            }

            var config = AiConfigManager.Current;
            config.Profiles = _profiles.ToList();

            if (SelectedProfile != null)
            {
                config.ActiveProfileId = SelectedProfile.Id;
            }

            bool ok = AiConfigManager.Save(config);
            if (ok)
            {
                SaveNotification = LocalizationService.Get("AiSet_SavedSuccess");
            }
            else
            {
                SaveNotification = LocalizationService.Get("AiSet_SaveFailed");
            }
        }

        private void ExecuteResetDefaults()
        {
            var confirm = WpfMessageBox.Show(
                "Bạn có chắc chắn muốn khôi phục danh sách các nhà cung cấp AI về mặc định không?",
                "Khôi phục mặc định",
                WpfMessageBoxButton.YesNo,
                WpfMessageBoxImage.Question);

            if (confirm == WpfMessageBoxResult.Yes)
            {
                var defaults = AiConfigManager.GetDefaultPresetProfiles();
                _profiles.Clear();
                foreach (var p in defaults)
                {
                    _profiles.Add(p);
                }

                SelectedProfile = _profiles.FirstOrDefault();
                ExecuteSaveSettings();
                SaveNotification = "Đã khôi phục danh sách cấu hình mặc định.";
            }
        }
    }
}
