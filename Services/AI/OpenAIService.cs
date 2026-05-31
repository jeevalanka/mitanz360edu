using System.Diagnostics;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace MITANZ360Edu.Web.Services.AI;

public sealed class OpenAIService : IAiProvider
{
    private readonly HttpClient _httpClient;

    private readonly IConfiguration _configuration;

    private readonly ILogger<OpenAIService> _logger;

    public AiProviderType ProviderType =>
        AiProviderType.OpenAI;

    public OpenAIService(
        HttpClient httpClient,
        IConfiguration configuration,
        ILogger<OpenAIService> logger)
    {
        _httpClient = httpClient;

        _configuration = configuration;

        _logger = logger;
    }

    public async Task<AiWorkflowResult> ExecuteAsync(
        AiWorkflowRequest request,
        AiModelConfig modelConfig,
        CancellationToken cancellationToken = default)
    {
        var stopwatch =
            Stopwatch.StartNew();

        try
        {
            _logger.LogInformation(
                "OpenAI execution started.");

            var apiKey =
                _configuration["OpenAI:Key"];

            var endpoint =
                _configuration["OpenAI:Endpoint"];

            if (string.IsNullOrWhiteSpace(apiKey))
            {
                throw new InvalidOperationException(
                    "OpenAI API key not configured.");
            }

            if (string.IsNullOrWhiteSpace(endpoint))
            {
                throw new InvalidOperationException(
                    "OpenAI endpoint not configured.");
            }

            var payload = new
            {
                model = modelConfig.Model,

                messages = new object[]
                {
                    new
                    {
                        role = "system",

                        content =
                            request.SystemPrompt ??
                            GetSystemPrompt(
                                request.OutputMode)
                    },

                    new
                    {
                        role = "user",

                        content = request.Prompt
                    }
                },

                temperature = request.Temperature,

                max_tokens = request.MaxTokens
            };

            var json =
                JsonSerializer.Serialize(payload);

            using var message =
                new HttpRequestMessage(
                    HttpMethod.Post,
                    endpoint);

            message.Headers.Authorization =
                new AuthenticationHeaderValue(
                    "Bearer",
                    apiKey);

            message.Content =
                new StringContent(
                    json,
                    Encoding.UTF8,
                    "application/json");

            var response =
                await _httpClient.SendAsync(
                    message,
                    cancellationToken);

            var rawResponse =
                await response.Content.ReadAsStringAsync(
                    cancellationToken);

            stopwatch.Stop();

            if (!response.IsSuccessStatusCode)
            {
                _logger.LogError(
                    "OpenAI request failed. StatusCode: {StatusCode}",
                    response.StatusCode);

                return new AiWorkflowResult
                {
                    Success = false,

                    Provider = ProviderType,

                    Model = modelConfig.Model,

                    RawResponse = rawResponse,

                    Errors = new List<AiError>
                    {
                        new AiError
                        {
                            Code = "OPENAI_ERROR",
                            Message = "OpenAI request failed"
                        }
                    },
                    DurationMs =
                        stopwatch.ElapsedMilliseconds
                };
            }

            _logger.LogInformation(
                "OpenAI execution completed successfully.");

            return new AiWorkflowResult
            {
                Success = true,

                Provider = ProviderType,

                Model = modelConfig.Model,

                RawResponse = rawResponse,

                DurationMs =
                    stopwatch.ElapsedMilliseconds
            };
        }
        catch (Exception ex)
        {
            stopwatch.Stop();

            _logger.LogError(
                ex,
                "OpenAI execution failed.");

            return new AiWorkflowResult
            {
                Success = false,

                Provider = ProviderType,

                Model = modelConfig.Model,

                Errors = new List<AiError>
                {
                    new AiError
                    {
                        Code = "OPENROUTER_ERROR",
                        Message = "some error"
                    }
                },

                DurationMs =
                    stopwatch.ElapsedMilliseconds
            };
        }
    }

    private static string GetSystemPrompt(
        AiOutputMode outputMode)
    {
        return outputMode switch
        {
            AiOutputMode.Text =>
                """
                You are a professional AI assistant for MITANZ360Edu.

                Respond ONLY in clean readable plain text.

                DO NOT return:
                - JSON
                - HTML
                - Markdown
                - code blocks

                Use human-friendly educational explanations.
                """,

            AiOutputMode.Html =>
                """
                Return clean Bootstrap-compatible HTML only.

                Do not return markdown.
                """,

            AiOutputMode.Markdown =>
                """
                Return professional markdown formatting.
                """,

            AiOutputMode.JsonOnly =>
                """
                Return ONLY valid JSON.

                Do not return explanations.
                """,

            AiOutputMode.FileUpdate =>
                """
                Return updated file content only.
                """,

            AiOutputMode.GeneratedTemplate =>
                """
                Generate professional HTML template output.
                """,

            _ =>
                """
                Respond in clean readable text.
                """
        };
    }
}