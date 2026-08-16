using System;
using System.Diagnostics;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ExcelSupport.Services
{
    public class TestConnectionResult
    {
        public bool IsSuccess { get; set; }
        public string Message { get; set; } = string.Empty;
        public long LatencyMs { get; set; }
        public string? ReplySample { get; set; }
    }

    public static class OpenAiClientService
    {
        private static readonly HttpClient HttpClient = new HttpClient();

        static OpenAiClientService()
        {
            try
            {
                System.Net.ServicePointManager.SecurityProtocol |= 
                    System.Net.SecurityProtocolType.Tls12 | 
                    System.Net.SecurityProtocolType.Tls11 | 
                    (System.Net.SecurityProtocolType)12288 /* Tls13 */ | 
                    System.Net.SecurityProtocolType.Tls;
            }
            catch
            {
                // ignore
            }
        }

        public static async Task<TestConnectionResult> TestConnectionAsync(AiConfig config)
        {
            var stopwatch = Stopwatch.StartNew();
            try
            {
                string baseUrl = NormalizeBaseUrl(config.BaseUrl);
                string endpoint = $"{baseUrl}/chat/completions";
                string model = string.IsNullOrWhiteSpace(config.ModelName) ? "qwen-3.6" : config.ModelName.Trim();

                var payload = new JObject
                {
                    ["model"] = model,
                    ["messages"] = new JArray
                    {
                        new JObject { ["role"] = "system", ["content"] = "You are a helpful AI assistant." },
                        new JObject { ["role"] = "user", ["content"] = "Ping! Reply with 'OK' only." }
                    }
                };

                // Use max_completion_tokens for o1/o3 or max_tokens
                if (IsReasoningModel(model))
                {
                    payload["max_completion_tokens"] = 50;
                }
                else
                {
                    payload["max_tokens"] = 30;
                    payload["temperature"] = 0.1;
                }

                using (var cts = new CancellationTokenSource(TimeSpan.FromSeconds(Math.Max(5, config.TimeoutSeconds))))
                {
                    using (var response = await SendWithFallbackAsync(endpoint, config.ApiKey, payload, cts.Token))
                    {
                        stopwatch.Stop();
                        string responseBody = await response.Content.ReadAsStringAsync();

                        if (!response.IsSuccessStatusCode)
                        {
                            string errDetail = ParseErrorMessage(responseBody) ?? response.ReasonPhrase ?? "Lỗi không xác định";
                            return new TestConnectionResult
                            {
                                IsSuccess = false,
                                Message = $"HTTP {(int)response.StatusCode}: {errDetail}",
                                LatencyMs = stopwatch.ElapsedMilliseconds
                            };
                        }

                        string reply = ExtractAssistantReply(responseBody) ?? "OK";
                        return new TestConnectionResult
                        {
                            IsSuccess = true,
                            Message = $"Kết nối thành công tới model [{model}]",
                            LatencyMs = stopwatch.ElapsedMilliseconds,
                            ReplySample = reply.Trim()
                        };
                    }
                }
            }
            catch (TaskCanceledException)
            {
                stopwatch.Stop();
                return new TestConnectionResult
                {
                    IsSuccess = false,
                    Message = $"Hết thời gian chờ (Timeout sau {config.TimeoutSeconds}s). Vui lòng kiểm tra lại URL hoặc kết nối mạng.",
                    LatencyMs = stopwatch.ElapsedMilliseconds
                };
            }
            catch (Exception ex)
            {
                stopwatch.Stop();
                string detail = ex.InnerException != null ? $"{ex.Message} ({ex.InnerException.Message})" : ex.Message;
                return new TestConnectionResult
                {
                    IsSuccess = false,
                    Message = $"Lỗi kết nối: {detail}",
                    LatencyMs = stopwatch.ElapsedMilliseconds
                };
            }
        }

        public static async Task<string> SendChatAsync(AiConfig config, string userPrompt, string? systemPrompt = null)
        {
            string baseUrl = NormalizeBaseUrl(config.BaseUrl);
            string endpoint = $"{baseUrl}/chat/completions";
            string model = string.IsNullOrWhiteSpace(config.ModelName) ? "qwen-3.6" : config.ModelName.Trim();

            var messagesArray = new JArray();
            if (!string.IsNullOrWhiteSpace(systemPrompt))
            {
                messagesArray.Add(new JObject { ["role"] = "system", ["content"] = systemPrompt });
            }
            messagesArray.Add(new JObject { ["role"] = "user", ["content"] = userPrompt });

            var payload = new JObject
            {
                ["model"] = model,
                ["messages"] = messagesArray
            };

            if (IsReasoningModel(model))
            {
                payload["max_completion_tokens"] = config.MaxTokens;
            }
            else
            {
                payload["max_tokens"] = config.MaxTokens;
                payload["temperature"] = config.Temperature;
            }

            using (var cts = new CancellationTokenSource(TimeSpan.FromSeconds(Math.Max(10, config.TimeoutSeconds))))
            {
                using (var response = await SendWithFallbackAsync(endpoint, config.ApiKey, payload, cts.Token))
                {
                    string responseBody = await response.Content.ReadAsStringAsync();

                    if (!response.IsSuccessStatusCode)
                    {
                        string errDetail = ParseErrorMessage(responseBody) ?? response.ReasonPhrase ?? "Lỗi API";
                        throw new InvalidOperationException($"Lỗi từ máy chủ AI (HTTP {(int)response.StatusCode}): {errDetail}");
                    }

                    return ExtractAssistantReply(responseBody) ?? string.Empty;
                }
            }
        }

        private static async Task<HttpResponseMessage> SendWithFallbackAsync(string endpoint, string? apiKey, JObject payload, CancellationToken token)
        {
            var response = await SendSingleRequestAsync(endpoint, apiKey, payload, token);
            if (!response.IsSuccessStatusCode)
            {
                string responseBody = await response.Content.ReadAsStringAsync();
                string err = ParseErrorMessage(responseBody) ?? "";

                bool needRetry = false;

                // Auto fallback between max_tokens and max_completion_tokens
                if (err.IndexOf("max_completion_tokens", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    err.IndexOf("max_tokens", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    if (payload.TryGetValue("max_tokens", out var val))
                    {
                        payload.Remove("max_tokens");
                        payload["max_completion_tokens"] = val;
                        needRetry = true;
                    }
                    else if (payload.TryGetValue("max_completion_tokens", out var compVal))
                    {
                        payload.Remove("max_completion_tokens");
                        payload["max_tokens"] = compVal;
                        needRetry = true;
                    }
                }

                // Auto fallback if temperature is unsupported (e.g. OpenAI o1/o3 reasoning models)
                if (err.IndexOf("temperature", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    if (payload.TryGetValue("temperature", out var tempVal))
                    {
                        payload.Remove("temperature");
                        needRetry = true;
                    }
                }

                if (needRetry)
                {
                    response.Dispose();
                    response = await SendSingleRequestAsync(endpoint, apiKey, payload, token);
                }
            }
            return response;
        }

        private static async Task<HttpResponseMessage> SendSingleRequestAsync(string endpoint, string? apiKey, JObject payload, CancellationToken token)
        {
            var request = new HttpRequestMessage(HttpMethod.Post, endpoint);
            string jsonContent = payload.ToString(Formatting.None);
            request.Content = new StringContent(jsonContent, Encoding.UTF8, "application/json");

            if (!string.IsNullOrWhiteSpace(apiKey))
            {
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", apiKey!.Trim());
            }

            return await HttpClient.SendAsync(request, token);
        }

        private static bool IsReasoningModel(string model)
        {
            if (string.IsNullOrWhiteSpace(model)) return false;
            string m = model.ToLowerInvariant();
            return m.StartsWith("o1") || m.StartsWith("o3");
        }

        private static string NormalizeBaseUrl(string url)
        {
            if (string.IsNullOrWhiteSpace(url)) return "http://localhost:8000/v1";
            string trimmed = url.Trim().TrimEnd('/');
            
            if (trimmed.EndsWith("/chat/completions", StringComparison.OrdinalIgnoreCase))
            {
                trimmed = trimmed.Substring(0, trimmed.Length - "/chat/completions".Length).TrimEnd('/');
            }
            if (trimmed.EndsWith("/chat", StringComparison.OrdinalIgnoreCase))
            {
                trimmed = trimmed.Substring(0, trimmed.Length - "/chat".Length).TrimEnd('/');
            }

            if (!trimmed.EndsWith("/v1", StringComparison.OrdinalIgnoreCase))
            {
                trimmed += "/v1";
            }
            return trimmed;
        }

        private static string? ExtractAssistantReply(string json)
        {
            try
            {
                var obj = JObject.Parse(json);
                var choices = obj["choices"] as JArray;
                if (choices != null && choices.Count > 0)
                {
                    var message = choices[0]["message"];
                    if (message != null)
                    {
                        return message["content"]?.ToString();
                    }
                }
            }
            catch
            {
                // ignore
            }
            return null;
        }

        private static string? ParseErrorMessage(string json)
        {
            try
            {
                var obj = JObject.Parse(json);
                var error = obj["error"];
                if (error != null)
                {
                    if (error is JValue) return error.ToString();
                    return error["message"]?.ToString() ?? error.ToString();
                }
            }
            catch
            {
                // ignore
            }
            return null;
        }
    }
}
