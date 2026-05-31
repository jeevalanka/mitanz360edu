using System.Diagnostics;
using System.Text;
using System.Text.Json;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace MITANZ360Edu.Web.Services.AI;

public sealed class AzureOpenAIService : IAiProvider
{
    private readonly HttpClient _httpClient;

    private readonly IConfiguration _configuration;

    private readonly ILogger<AzureOpenAIService> _logger;

    public AiProviderType ProviderType =>
        AiProviderType.AzureOpenAI;

    public AzureOpenAIService(
        HttpClient httpClient,
        IConfiguration configuration,
        ILogger<AzureOpenAIService> logger)
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
                "Azure OpenAI execution started. OutputMode: {OutputMode}",
                request.OutputMode);

            var endpoint =
                _configuration["OpenAI:Endpoint"];

            var deployment =
                _configuration["OpenAI:Deployment"];

            var key =
                _configuration["OpenAI:Key"];

            if (string.IsNullOrWhiteSpace(endpoint))
            {
                throw new InvalidOperationException(
                    "Azure OpenAI endpoint not configured.");
            }

            if (string.IsNullOrWhiteSpace(deployment))
            {
                throw new InvalidOperationException(
                    "Azure OpenAI deployment not configured.");
            }

            if (string.IsNullOrWhiteSpace(key))
            {
                throw new InvalidOperationException(
                    "Azure OpenAI API key not configured.");
            }

            var url =
                $"{endpoint}openai/deployments/{deployment}/chat/completions?api-version=2024-02-15-preview";

            var payload = new
            {
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

            using var requestMessage =
                new HttpRequestMessage(
                    HttpMethod.Post,
                    url);

            requestMessage.Headers.Add(
                "api-key",
                key);

            requestMessage.Content =
                new StringContent(
                    json,
                    Encoding.UTF8,
                    "application/json");

            var response =
                await _httpClient.SendAsync(
                    requestMessage,
                    cancellationToken);

            var rawResponse =
                await response.Content.ReadAsStringAsync(
                    cancellationToken);

            stopwatch.Stop();

            if (!response.IsSuccessStatusCode)
            {
                _logger.LogError(
                    "Azure OpenAI request failed. StatusCode: {StatusCode}",
                    response.StatusCode);

                return new AiWorkflowResult
                {
                    Success = false,

                    Provider = ProviderType,

                    Model = deployment,

                    RawResponse = rawResponse,

                    Errors = new List<AiError>
                    {
                        new AiError
                        {
                            Code = "PROVIDER_NOT_FOUND",
                            Message = $"Provider not registered: {modelConfig.Provider}"
                        }
                    },

                    DurationMs =
                        stopwatch.ElapsedMilliseconds
                };
            }

            _logger.LogInformation(
                "Azure OpenAI execution completed successfully.");

            return new AiWorkflowResult
            {
                Success = true,

                Provider = ProviderType,

                Model = deployment,

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
                "Azure OpenAI execution failed.");

            return new AiWorkflowResult
            {
                Success = false,

                Provider = ProviderType,

                Model = modelConfig.Model,

                Errors = new List<AiError>
                    {
                        new AiError
                        {
                            Code = "EXCEPTION",
                            Message = ex.Message
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

                Use proper Bootstrap cards, tables, alerts, and headings.
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
                Generate professional enterprise HTML template output.
                """,

            _ =>
                """
                Respond in clean readable text.
                """
        };
    }
}