using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace ResumeBuilder.Services;

public class OpenAiService
{
    private readonly HttpClient _httpClient;
    private readonly string _apiKey;

    public OpenAiService(string apiKey)
    {
        _apiKey = apiKey;
        _httpClient = new HttpClient
        {
            Timeout = TimeSpan.FromSeconds(90)
        };
        _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _apiKey);
    }

    public async Task<string> CreateChatCompletionAsync(string systemMessage, string userPrompt, string model, double temperature, double topP)
    {
        var payload = new
        {
            model,
            temperature,
            top_p = topP,
            messages = new[]
            {
                new { role = "system", content = systemMessage },
                new { role = "user", content = userPrompt }
            }
        };

        var json = JsonSerializer.Serialize(payload);
        using var content = new StringContent(json, Encoding.UTF8, "application/json");
        using var response = await _httpClient.PostAsync("https://api.openai.com/v1/chat/completions", content);
        var responseText = await response.Content.ReadAsStringAsync();

        if (!response.IsSuccessStatusCode)
        {
            throw new Exception($"OpenAI API error: {response.StatusCode} - {responseText}");
        }

        using var doc = JsonDocument.Parse(responseText);
        var root = doc.RootElement;
        if (!root.TryGetProperty("choices", out var choices) || choices.GetArrayLength() == 0)
        {
            throw new Exception("OpenAI API returned no choices.");
        }

        var contentNode = choices[0].GetProperty("message").GetProperty("content");
        return contentNode.GetString() ?? string.Empty;
    }

    public static string ExtractJson(string content)
    {
        if (string.IsNullOrWhiteSpace(content))
        {
            return string.Empty;
        }

        var cleaned = content.Trim();
        if (cleaned.StartsWith("```"))
        {
            cleaned = cleaned.Trim('`').Trim();
            var firstLineEnd = cleaned.IndexOf('\n');
            if (firstLineEnd >= 0)
            {
                cleaned = cleaned[(firstLineEnd + 1)..].Trim();
            }
        }

        var startIdx = cleaned.IndexOf('{');
        var altIdx = cleaned.IndexOf('[');
        if (startIdx < 0 || (altIdx >= 0 && altIdx < startIdx))
        {
            startIdx = altIdx;
        }

        if (startIdx > 0)
        {
            cleaned = cleaned[startIdx..].Trim();
        }

        return cleaned.Replace("```", string.Empty).Trim();
    }

    public static async Task<(bool IsValid, string ErrorMessage)> ValidateApiKeyAsync(string apiKey)
    {
        if (string.IsNullOrWhiteSpace(apiKey))
        {
            return (false, "OpenAI API key is required.");
        }

        using var client = new HttpClient
        {
            Timeout = TimeSpan.FromSeconds(20)
        };
        client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", apiKey.Trim());

        try
        {
            using var response = await client.GetAsync("https://api.openai.com/v1/models");
            var body = await response.Content.ReadAsStringAsync();

            if (response.IsSuccessStatusCode)
            {
                return (true, string.Empty);
            }

            if (response.StatusCode == System.Net.HttpStatusCode.Unauthorized ||
                response.StatusCode == System.Net.HttpStatusCode.Forbidden)
            {
                return (false, "Invalid OpenAI API key.");
            }

            var preview = body.Length > 220 ? body[..220] + "..." : body;
            return (false, $"Could not validate key ({response.StatusCode}). {preview}");
        }
        catch (Exception ex)
        {
            return (false, $"Failed to validate key: {ex.Message}");
        }
    }
}
