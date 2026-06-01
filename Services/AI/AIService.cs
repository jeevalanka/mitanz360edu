using System.Text;
using System.Text.Json;

namespace MITANZ360Edu.Web.Services.AI;

public class AIService
{
    private readonly HttpClient _http;
    private readonly IConfiguration _config;

    public AIService(HttpClient http, IConfiguration config)
    {
        _http = http;
        _config = config;
    }

    public async Task<string> GenerateTextAsync(string prompt)
    {
        var endpoint = _config["OpenAI:Endpoint"];
        var key = _config["OpenAI:Key"];
        var deployment = _config["OpenAI:Deployment"];

        if (string.IsNullOrWhiteSpace(endpoint) ||
            string.IsNullOrWhiteSpace(key) ||
            string.IsNullOrWhiteSpace(deployment))
        {
            return "⚠ Azure OpenAI configuration missing.";
        }

        var url =
            $"{endpoint}/openai/deployments/{deployment}/chat/completions?api-version=2024-02-15-preview";

        var request = new
        {
            messages = new[]
            {
                new { role = "system", content = "You are an academic evaluation engine." },
                new { role = "user", content = prompt }
            },
            temperature = 0.3,
            max_tokens = 1200
        };

        var json = JsonSerializer.Serialize(request);

        var httpRequest = new HttpRequestMessage(HttpMethod.Post, url);
        httpRequest.Headers.Add("api-key", key);
        httpRequest.Content = new StringContent(json, Encoding.UTF8, "application/json");

        var response = await _http.SendAsync(httpRequest);

        if (!response.IsSuccessStatusCode)
        {
            var error = await response.Content.ReadAsStringAsync();
            return $"⚠ AI request failed: {response.StatusCode}\n{error}";
        }

        var content = await response.Content.ReadAsStringAsync();

        using var doc = JsonDocument.Parse(content);

        var result = doc.RootElement
            .GetProperty("choices")[0]
            .GetProperty("message")
            .GetProperty("content")
            .GetString();

        return result ?? string.Empty;
    }
}