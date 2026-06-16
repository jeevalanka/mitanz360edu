using Microsoft.Extensions.Logging;
using MITANZ360Edu.Web.Models.Workflow;

namespace MITANZ360Edu.Web.Services.Workflow.Plugins;

/// <summary>
/// Reads files from the local file system.
/// </summary>
public sealed class FilePlugin : WorkflowPluginBase
{
    public FilePlugin(ILogger<FilePlugin> logger)
        : base(logger)
    {
    }

    public override string Type => "file";

    public override async Task ExecuteAsync(
        WorkflowContext context,
        WorkflowStep step,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(context);
        ArgumentNullException.ThrowIfNull(step);

        var path = GetSetting<string>(step, "path");

        if (string.IsNullOrWhiteSpace(path))
            throw new InvalidOperationException("File path is required.");

        if (!File.Exists(path))
            throw new FileNotFoundException(path);

        var extension = Path.GetExtension(path).ToLowerInvariant();

        object result = extension switch
        {
            ".txt" or ".json" or ".xml" or ".csv" or ".md"
                => await File.ReadAllTextAsync(path, cancellationToken),

            _ => await File.ReadAllBytesAsync(path, cancellationToken)
        };

        context.Set(step.Output, result);
    }
}