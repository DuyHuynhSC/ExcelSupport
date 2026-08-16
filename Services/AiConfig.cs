using System;

namespace ExcelSupport.Services
{
    public class AiConfig
    {
        public string BaseUrl { get; set; } = "http://localhost:8000/v1";
        public string ApiKey { get; set; } = string.Empty;
        public string ModelName { get; set; } = "qwen-3.6";
        public double Temperature { get; set; } = 0.3;
        public int MaxTokens { get; set; } = 2048;
        public int TimeoutSeconds { get; set; } = 30;
        public bool IsDarkTheme { get; set; } = false;
    }
}
