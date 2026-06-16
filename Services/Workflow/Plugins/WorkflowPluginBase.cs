using MITANZ360Edu.Web.Models.Workflow;
using Microsoft.Extensions.Logging;

namespace MITANZ360Edu.Web.Services.Workflow.Plugins;

/// <summary>
/// Base class for workflow plugins.
/// </summary>
public abstract class WorkflowPluginBase : IWorkflowPlugin
{
    protected WorkflowPluginBase(ILogger logger)
    {
        Logger = logger;
    }

    /// <summary>
    /// Logger instance.
    /// </summary>
    protected ILogger Logger { get; }

    /// <summary>
    /// Plugin type.
    /// </summary>
    public abstract string Type { get; }

    /// <summary>
    /// Execute the plugin.
    /// </summary>
    public abstract Task ExecuteAsync(
        WorkflowContext context,
        WorkflowStep step,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Gets a setting value.
    /// </summary>
    protected T? GetSetting<T>(WorkflowStep step, string key)
    {
        if (step.Settings.TryGetValue(key, out var value) &&
            value is T result)
        {
            return result;
        }

        return default;
    }
}