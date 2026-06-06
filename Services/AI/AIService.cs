using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace MITANZ360Edu.Web.Services.AI;

public class AIResultDto
{
    public string Summary { get; set; } = "";
    public string Html { get; set; } = "";
    public string Tags { get; set; } = "";
    public decimal Score { get; set; }
    public string UpdatedJson { get; set; } = "";
    public byte[] DocxBytes { get; set; } = Array.Empty<byte>();
}

public partial class AIService
{
    private readonly HttpClient _http;
    private readonly IConfiguration _config;

    public AIService(HttpClient http, IConfiguration config)
    {
        _http = http;
        _config = config;
    }

    // ✅ EXISTING METHOD (KEEP)
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
                new { role = "system", content = "You are an AI course designer." },
                new { role = "user", content = prompt }
            },
            temperature = 0.3,
            max_tokens = 1500
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

        return doc.RootElement
            .GetProperty("choices")[0]
            .GetProperty("message")
            .GetProperty("content")
            .GetString() ?? "";
    }

    // =====================================================
    // ✅ NEW: COURSE AI ENGINE (THIS FIXES YOUR ERROR)
    // =====================================================
    public async Task<AIResultDto> GenerateCourseAsync(JsonObject json)
    {
        var prompt = BuildCoursePrompt(json);

        var raw = await GenerateTextAsync(prompt);

        return ParseAIResponse(raw, json);
    }

    // ✅ PROMPT BUILDER
    private string BuildCoursePrompt(JsonObject json)
    {
        return $@"
You are an expert academic course designer.

Generate output in STRICT JSON format:

{{
  ""summary"": ""short plain text summary"",
  ""html"": ""HTML formatted course report"",
  ""tags"": ""comma separated keywords"",
  ""score"": 0
}}

INPUT:
{json.ToJsonString()}
";
    }

    // ✅ PARSER
    private AIResultDto ParseAIResponse(string response, JsonObject input)
    {
        try
        {
            var json = JsonNode.Parse(response);

            return new AIResultDto
            {
                Summary = json?["summary"]?.ToString() ?? "",
                Html = json?["html"]?.ToString() ?? "",
                Tags = json?["tags"]?.ToString() ?? "",
                Score = decimal.TryParse(json?["score"]?.ToString(), out var s) ? s : 0,
                UpdatedJson = input.ToJsonString(),
                DocxBytes = GenerateDocx(input)
            };
        }
        catch
        {
            return new AIResultDto
            {
                Summary = "⚠ AI parsing failed",
                Html = "<p>Error parsing AI output</p>",
                Tags = "",
                Score = 0,
                UpdatedJson = input.ToJsonString()
            };
        }
    }

    // ✅ DOCX GENERATOR
    private byte[] GenerateDocx(JsonObject json)
    {
        using var ms = new MemoryStream();

        using var doc =
            DocumentFormat.OpenXml.Packaging.WordprocessingDocument
            .Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);

        var main = doc.AddMainDocumentPart();
        main.Document = new DocumentFormat.OpenXml.Wordprocessing.Document(
            new DocumentFormat.OpenXml.Wordprocessing.Body(
                new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                    new DocumentFormat.OpenXml.Wordprocessing.Run(
                        new DocumentFormat.OpenXml.Wordprocessing.Text(
                            json.ToJsonString()))
                )
            )
        );

        return ms.ToArray();
    }
}