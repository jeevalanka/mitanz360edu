using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using MITANZ360Edu.Web.Models.Workflow;

namespace MITANZ360Edu.Web.Services.Workflow.Plugins;

/// <summary>
/// AI workflow plugin.
/// </summary>
public sealed class AiPlugin : WorkflowPluginBase
{
    private readonly IConfiguration _configuration;

    public AiPlugin(
        IConfiguration configuration,
        ILogger<AiPlugin> logger)
        : base(logger)
    {
        _configuration = configuration;
    }

    public override string Type => "ai";

    public override async Task ExecuteAsync(
        WorkflowContext context,
        WorkflowStep step,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(context);
        ArgumentNullException.ThrowIfNull(step);

        var provider = GetSetting<string>(step, "provider") ?? "AzureOpenAI";
        var model = GetSetting<string>(step, "model") ?? string.Empty;
        var prompt = GetSetting<string>(step, "prompt") ?? string.Empty;

        // TODO:
        // Read Azure OpenAI settings from appsettings.json
        // Build AI request
        // Submit prompt
        // Store response

        var result = new
        {
            Provider = provider,
            Model = model,
            Prompt = prompt,
            Status = "Pending",
            ExecutedOn = DateTime.UtcNow
        };

        context.Set(step.Output, result);

        await Task.CompletedTask;
    }
}