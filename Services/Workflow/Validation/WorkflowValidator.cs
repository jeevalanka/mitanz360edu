using MITANZ360Edu.Web.Models.Workflow;

namespace MITANZ360Edu.Web.Services.Workflow.Validation;

/// <summary>
/// Validates workflow definitions before saving or execution.
/// </summary>
public sealed class WorkflowValidator
{
    public IReadOnlyList<string> Validate(WorkflowDefinition workflow)
    {
        ArgumentNullException.ThrowIfNull(workflow);

        var errors = new List<string>();

        if (string.IsNullOrWhiteSpace(workflow.Name))
            errors.Add("Workflow name is required.");

        foreach (var step in workflow.DataSources)
        {
            if (string.IsNullOrWhiteSpace(step.Name))
                errors.Add("A data source name is required.");

            if (string.IsNullOrWhiteSpace(step.Type))
                errors.Add($"'{step.Name}' type is required.");

            if (string.IsNullOrWhiteSpace(step.Output))
                errors.Add($"'{step.Name}' output is required.");

            if (step.Type.Equals("scrape", StringComparison.OrdinalIgnoreCase))
            {
                if (!step.Settings.ContainsKey("url"))
                    errors.Add($"'{step.Name}' requires a URL.");
            }

            if (step.Type.Equals("http", StringComparison.OrdinalIgnoreCase))
            {
                if (!step.Settings.ContainsKey("url"))
                    errors.Add($"'{step.Name}' requires a URL.");
            }

            if (step.Type.Equals("sql", StringComparison.OrdinalIgnoreCase))
            {
                if (!step.Settings.ContainsKey("query"))
                    errors.Add($"'{step.Name}' requires a SQL query.");
            }

            if (step.Type.Equals("file", StringComparison.OrdinalIgnoreCase))
            {
                if (!step.Settings.ContainsKey("path"))
                    errors.Add($"'{step.Name}' requires a file path.");
            }
        }

        foreach (var ai in workflow.AiSteps)
        {
            if (string.IsNullOrWhiteSpace(ai.Name))
                errors.Add("AI step name is required.");

            if (string.IsNullOrWhiteSpace(ai.Provider))
                errors.Add($"'{ai.Name}' provider is required.");

            if (string.IsNullOrWhiteSpace(ai.Prompt))
                errors.Add($"'{ai.Name}' prompt is required.");

            if (string.IsNullOrWhiteSpace(ai.Output))
                errors.Add($"'{ai.Name}' output is required.");
        }

        return errors;
    }
}