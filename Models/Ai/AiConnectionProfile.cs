using System;
using System.ComponentModel;
using Newtonsoft.Json;

namespace ExcelSupport.Models
{
    public class AiConnectionProfile : ObservableModel
    {
        private string _id = Guid.NewGuid().ToString();
        private string _name = "OpenAI / Local AI";
        private string _baseUrl = "http://localhost:8000/v1";
        private string _apiKey = string.Empty;
        private string _modelName = "qwen-3.6";
        private int _timeoutSeconds = 30;
        private double _temperature = 0.3;
        private int _maxTokens = 2048;
        private bool _isDefault = false;

        public string Id
        {
            get => _id;
            set { _id = value; OnPropertyChanged(nameof(Id)); }
        }

        public string Name
        {
            get => _name;
            set { _name = value; OnPropertyChanged(nameof(Name)); }
        }

        public string BaseUrl
        {
            get => _baseUrl;
            set 
            { 
                _baseUrl = value; 
                OnPropertyChanged(nameof(BaseUrl)); 
                OnPropertyChanged(nameof(DisplaySummary)); 
            }
        }

        public string ApiKey
        {
            get => _apiKey;
            set { _apiKey = value; OnPropertyChanged(nameof(ApiKey)); }
        }

        public string ModelName
        {
            get => _modelName;
            set 
            { 
                _modelName = value; 
                OnPropertyChanged(nameof(ModelName)); 
                OnPropertyChanged(nameof(DisplaySummary)); 
            }
        }

        public int TimeoutSeconds
        {
            get => _timeoutSeconds;
            set { _timeoutSeconds = value; OnPropertyChanged(nameof(TimeoutSeconds)); }
        }

        public double Temperature
        {
            get => _temperature;
            set { _temperature = value; OnPropertyChanged(nameof(Temperature)); }
        }

        public int MaxTokens
        {
            get => _maxTokens;
            set { _maxTokens = value; OnPropertyChanged(nameof(MaxTokens)); }
        }

        public bool IsDefault
        {
            get => _isDefault;
            set
            {
                _isDefault = value;
                OnPropertyChanged(nameof(IsDefault));
                OnPropertyChanged(nameof(DefaultBadge));
            }
        }

        [JsonIgnore]
        public string DefaultBadge => IsDefault ? "⭐ (Mặc định)" : "";

        [JsonIgnore]
        public string DisplaySummary
        {
            get
            {
                string m = string.IsNullOrWhiteSpace(ModelName) ? "Default Model" : ModelName;
                string u = string.IsNullOrWhiteSpace(BaseUrl) ? "" : BaseUrl;
                return $"{m} • {u}";
            }
        }

        public AiConnectionProfile Clone()
        {
            return new AiConnectionProfile
            {
                Id = Guid.NewGuid().ToString(),
                Name = $"{Name} (Bản sao)",
                BaseUrl = BaseUrl,
                ApiKey = ApiKey,
                ModelName = ModelName,
                TimeoutSeconds = TimeoutSeconds,
                Temperature = Temperature,
                MaxTokens = MaxTokens,
                IsDefault = false
            };
        }
    }
}
