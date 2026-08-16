using System;
using System.Threading.Tasks;
using System.Windows.Input;
using ExcelSupport.Services;

namespace ExcelSupport.ViewModels
{
    public class AiSettingsViewModel : ViewModelBase
    {
        private string _baseUrl;
        private string _apiKey;
        private string _modelName;
        private int _timeoutSeconds;
        private double _temperature;

        private bool _isTesting;
        private string _testStatusMessage = string.Empty;
        private bool? _isTestSuccess;
        private string _latencyText = string.Empty;
        private string _saveNotification = string.Empty;

        public string BaseUrl
        {
            get => _baseUrl;
            set => SetProperty(ref _baseUrl, value);
        }

        public string ApiKey
        {
            get => _apiKey;
            set => SetProperty(ref _apiKey, value);
        }

        public string ModelName
        {
            get => _modelName;
            set => SetProperty(ref _modelName, value);
        }

        public int TimeoutSeconds
        {
            get => _timeoutSeconds;
            set => SetProperty(ref _timeoutSeconds, value);
        }

        public double Temperature
        {
            get => _temperature;
            set => SetProperty(ref _temperature, value);
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

        public AiSettingsViewModel()
        {
            var config = AiConfigManager.Current;
            _baseUrl = config.BaseUrl;
            _apiKey = config.ApiKey;
            _modelName = config.ModelName;
            _timeoutSeconds = config.TimeoutSeconds;
            _temperature = config.Temperature;

            TestConnectionCommand = new RelayCommand(async _ => await ExecuteTestConnectionAsync(), _ => !IsTesting);
            SaveSettingsCommand = new RelayCommand(_ => ExecuteSaveSettings());
            ResetDefaultsCommand = new RelayCommand(_ => ExecuteResetDefaults());
        }

        private async Task ExecuteTestConnectionAsync()
        {
            IsTesting = true;
            IsTestSuccess = null;
            TestStatusMessage = "Đang kết nối tới máy chủ AI nội bộ...";
            LatencyText = string.Empty;
            SaveNotification = string.Empty;

            var testConfig = new AiConfig
            {
                BaseUrl = BaseUrl,
                ApiKey = ApiKey,
                ModelName = ModelName,
                TimeoutSeconds = TimeoutSeconds,
                Temperature = Temperature
            };

            var result = await Task.Run(() => OpenAiClientService.TestConnectionAsync(testConfig));

            IsTesting = false;
            IsTestSuccess = result.IsSuccess;
            TestStatusMessage = result.Message;
            if (result.IsSuccess)
            {
                LatencyText = $"⚡ Độ trễ: {result.LatencyMs} ms";
            }
        }

        private void ExecuteSaveSettings()
        {
            var config = new AiConfig
            {
                BaseUrl = BaseUrl?.Trim() ?? string.Empty,
                ApiKey = ApiKey?.Trim() ?? string.Empty,
                ModelName = string.IsNullOrWhiteSpace(ModelName) ? "qwen-3.6" : ModelName.Trim(),
                TimeoutSeconds = TimeoutSeconds <= 0 ? 30 : TimeoutSeconds,
                Temperature = Temperature
            };

            bool ok = AiConfigManager.Save(config);
            if (ok)
            {
                SaveNotification = "✅ Đã lưu cấu hình AI thành công!";
            }
            else
            {
                SaveNotification = "❌ Lưu cấu hình thất bại.";
            }
        }

        private void ExecuteResetDefaults()
        {
            var defaultConfig = new AiConfig();
            BaseUrl = defaultConfig.BaseUrl;
            ApiKey = defaultConfig.ApiKey;
            ModelName = defaultConfig.ModelName;
            TimeoutSeconds = defaultConfig.TimeoutSeconds;
            Temperature = defaultConfig.Temperature;
            TestStatusMessage = string.Empty;
            IsTestSuccess = null;
            SaveNotification = "Đã khôi phục cài đặt mặc định.";
        }
    }
}
