using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OSEMAddIn.Models;

namespace OSEMAddIn.Services
{
    internal sealed class LlmExtractionService
    {
        private static readonly HttpClient _httpClient = new HttpClient();

        public async Task<Dictionary<string, string>> ExtractAsync(PromptDefinition prompt, DashboardTemplate template, MailItem mailItem, LlmModelConfiguration configuration)
        {
            if (prompt is null) throw new ArgumentNullException(nameof(prompt));
            if (template is null) throw new ArgumentNullException(nameof(template));
            if (mailItem is null) throw new ArgumentNullException(nameof(mailItem));
            if (configuration is null) throw new ArgumentNullException(nameof(configuration));

            // 1. Prepare Prompt
            string promptBody = prompt.Body;
            
            // Replace {{DASHBOARD_JSON}}
            var schema = BuildDashboardJson(template);
            promptBody = promptBody.Replace("{{DASHBOARD_JSON}}", schema.ToString(Formatting.Indented));
            
            // Replace Mail Variables
            promptBody = promptBody.Replace("{{MailSubject}}", mailItem.Subject ?? "");
            promptBody = promptBody.Replace("{{MailSender}}", mailItem.SenderName ?? "");
            
            string mailBody = mailItem.Body ?? "";
            
            // Handle {{MAIL_BODY_LATEST}}
            if (promptBody.Contains("{{MAIL_BODY_LATEST}}"))
            {
                string latestBody = ExtractLatestBody(mailBody);
                promptBody = promptBody.Replace("{{MAIL_BODY_LATEST}}", latestBody);
            }

            // Handle {{MAIL_BODY}}
            if (promptBody.Contains("{{MAIL_BODY}}"))
            {
                promptBody = promptBody.Replace("{{MAIL_BODY}}", mailBody);
            }
            // Automatic appending of mail body removed as per requirement.
            // Users must explicitly include {{MAIL_BODY}} or {{MAIL_BODY_LATEST}} in their prompts.

            // 2. Call LLM
            string jsonResponse;
            try 
            {
                if (string.Equals(configuration.Provider, "Ollama", StringComparison.OrdinalIgnoreCase))
                {
                    jsonResponse = await CallOllamaAsync(promptBody, configuration);
                }
                else if (string.Equals(configuration.Provider, "OpenAI", StringComparison.OrdinalIgnoreCase) || 
                         string.Equals(configuration.Provider, "Custom", StringComparison.OrdinalIgnoreCase))
                {
                    jsonResponse = await CallOpenAiAsync(promptBody, configuration);
                }
                else
                {
                    // Default to Ollama if unknown
                    jsonResponse = await CallOllamaAsync(promptBody, configuration);
                }
            }
            catch (System.Exception ex)
            {
                DebugLogger.Log($"LLM Call failed: {ex}");
                throw;
            }

            // 3. Parse JSON
            return ParseLlmResponse(jsonResponse);
        }

        private async Task<string> CallOllamaAsync(string prompt, LlmModelConfiguration config)
        {
            var endpoint = config.ApiEndpoint;
            if (string.IsNullOrWhiteSpace(endpoint))
            {
                endpoint = "http://localhost:11434";
            }
            endpoint = endpoint!.TrimEnd('/');
            
            // Use /api/generate for simple completion
            var url = $"{endpoint}/api/generate";
            
            var requestObj = new
            {
                model = string.IsNullOrWhiteSpace(config.ModelName) ? "llama3" : config.ModelName,
                prompt = prompt,
                stream = false,
                format = "json" // Force JSON mode if supported by model/ollama version
            };

            var json = JsonConvert.SerializeObject(requestObj);
            var content = new StringContent(json, Encoding.UTF8, "application/json");

            var response = await _httpClient.PostAsync(url, content);
            response.EnsureSuccessStatusCode();

            var responseString = await response.Content.ReadAsStringAsync();
            var responseJson = JObject.Parse(responseString);
            
            return responseJson["response"]?.ToString() ?? string.Empty;
        }

        private async Task<string> CallOpenAiAsync(string prompt, LlmModelConfiguration config)
        {
            // Ensure TLS 1.2 is enabled (crucial for VSTO/.NET Framework connecting to modern APIs)
            try
            {
                System.Net.ServicePointManager.SecurityProtocol |= System.Net.SecurityProtocolType.Tls12;
            }
            catch
            {
                // Ignore if already set or not supported (unlikely on modern Windows)
            }

            var endpoint = config.ApiEndpoint;
            if (string.IsNullOrWhiteSpace(endpoint))
            {
                endpoint = "https://api.openai.com/v1";
            }
            endpoint = endpoint!.TrimEnd('/');
            
            // Smart URL handling: avoid double appending /chat/completions
            string url;
            if (endpoint.EndsWith("/chat/completions", StringComparison.OrdinalIgnoreCase))
            {
                url = endpoint;
            }
            else
            {
                url = $"{endpoint}/chat/completions";
            }

            var requestObj = new
            {
                model = string.IsNullOrWhiteSpace(config.ModelName) ? "gpt-3.5-turbo" : config.ModelName,
                messages = new[]
                {
                    new { role = "system", content = "You are a helpful assistant that extracts data from emails into JSON format." },
                    new { role = "user", content = prompt }
                }
                // response_format removed for compatibility
            };

            var json = JsonConvert.SerializeObject(requestObj, new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore });
            var request = new HttpRequestMessage(HttpMethod.Post, url);
            request.Content = new StringContent(json, Encoding.UTF8, "application/json");
            
            if (!string.IsNullOrWhiteSpace(config.ApiKey))
            {
                request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", config.ApiKey);
            }

            var response = await _httpClient.SendAsync(request);
            
            if (!response.IsSuccessStatusCode)
            {
                var errorContent = await response.Content.ReadAsStringAsync();
                throw new HttpRequestException($"API Request failed with status {response.StatusCode}: {errorContent}");
            }

            var responseString = await response.Content.ReadAsStringAsync();
            var responseJson = JObject.Parse(responseString);
            
            return responseJson["choices"]?[0]?["message"]?["content"]?.ToString() ?? string.Empty;
        }

        private Dictionary<string, string> ParseLlmResponse(string json)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (string.IsNullOrWhiteSpace(json)) return result;

            try
            {
                // Try to find JSON object in the text if it's mixed with other text
                var startIndex = json.IndexOf('{');
                var endIndex = json.LastIndexOf('}');
                
                if (startIndex >= 0 && endIndex > startIndex)
                {
                    json = json.Substring(startIndex, endIndex - startIndex + 1);
                }

                var jObject = JObject.Parse(json);
                foreach (var prop in jObject.Properties())
                {
                    result[prop.Name] = prop.Value.ToString();
                }
            }
            catch (System.Exception ex)
            {
                DebugLogger.Log($"Failed to parse LLM JSON response: {ex}");
                // Fallback: maybe try to extract key-value pairs with regex if JSON fails?
                // For now, just return empty or partial
            }

            return result;
        }

        public JObject BuildDashboardJson(DashboardTemplate template)
        {
            var obj = new JObject();
            foreach (var field in template.Fields)
            {
                obj[field] = string.Empty;
            }

            return obj;
        }

        private string ExtractLatestBody(string fullBody)
        {
            if (string.IsNullOrWhiteSpace(fullBody)) return string.Empty;

            // Common separators for email history
            string[] separators = new[]
            {
                "-----Original Message-----",
                "________________________________",
                "From:",
                "Sent:",
                "To:",
                "Subject:"
            };

            using (var reader = new System.IO.StringReader(fullBody))
            {
                var sb = new StringBuilder();
                string? line;
                while ((line = reader.ReadLine()) != null)
                {
                    var trimmed = line.Trim();
                    // Check for separators
                    // Simple heuristic: if a line starts with "From:" or contains "Original Message", stop.
                    // Note: This is a basic implementation. Email parsing is complex.
                    
                    if (trimmed.StartsWith("-----Original Message-----") || 
                        trimmed.StartsWith("________________________________"))
                    {
                        break;
                    }

                    // "From:" check is tricky because it might be part of normal text.
                    // Usually "From:" followed by "Sent:" in next lines indicates a header.
                    // For now, let's stick to the explicit separators or just return full body if not found.
                    // A slightly better check for Outlook style:
                    if (trimmed.StartsWith("From:") && (fullBody.Contains("Sent:") || fullBody.Contains("To:")))
                    {
                         // This is risky without looking ahead, but acceptable for a simple "Latest" feature.
                         // Let's try to be a bit more conservative: only break if we see a block of headers.
                         // For this simple version, we will just stop at the explicit separators.
                    }

                    sb.AppendLine(line);
                }
                return sb.ToString().Trim();
            }
        }
    }
}
