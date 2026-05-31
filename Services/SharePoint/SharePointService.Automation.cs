using System.Text.Json;

namespace MITANZ360Edu.Web.Services;

/// <summary>
/// SharePointService.Automation module
/// ------------------------------------------------------------
/// Uses existing SharePointService helpers:
/// - GetListIdByNameAsync(listName, ct)
/// - UpsertItemAsync(listId, keyField, keyValue, fields, ct)
///
/// REQUIRED SHAREPOINT LISTS:
/// - AutomationRuns
/// - AutomationStepRuns
/// - AutomationAuditLog
///
/// REQUIRED COLUMNS (internal names recommended):
/// AutomationRuns:
/// - Title, RunId, WorkflowId, WorkflowName, Status, TriggeredByUpn, TriggerType,
///   StartedUtc, CompletedUtc, CorrelationId, InputJson, Error
///
/// AutomationStepRuns:
/// - Title, RunId, StepId, StepType, StepName, Attempt, Status,
///   StartedUtc, CompletedUtc, DurationMs, OutputJson, Error, CorrelationId
///
/// AutomationAuditLog:
/// - Title, RunId, WorkflowId, EventType, Message, CreatedUtc, CorrelationId
/// </summary>
public partial class SharePointService
{
    private const string RunsList = "AutomationRuns";
    private const string StepRunsList = "AutomationStepRuns";
    private const string AuditList = "AutomationAuditLog";

    public async Task CreateAutomationRunAsync(
        string runId,
        string workflowId,
        string workflowName,
        string status,
        string triggeredByUpn,
        string triggerType,
        DateTime startedUtc,
        string correlationId,
        string? inputJson,
        CancellationToken ct)
    {
        var listId = await GetListIdByNameAsync(RunsList, ct).ConfigureAwait(false);

        var fields = new Dictionary<string, object?>
        {
            ["Title"] = $"{workflowName} - {runId}",
            ["RunId"] = runId,
            ["WorkflowId"] = workflowId,
            ["WorkflowName"] = workflowName,
            ["Status"] = status,
            ["TriggeredByUpn"] = triggeredByUpn,
            ["TriggerType"] = triggerType,
            ["StartedUtc"] = startedUtc.ToString("o"),
            ["CorrelationId"] = correlationId,
            ["InputJson"] = Trunc(inputJson, 10000)
        };

        // Upsert by RunId (no duplicates)
        await UpsertItemAsync(listId, "RunId", runId, fields, ct).ConfigureAwait(false);
    }

    public async Task UpdateAutomationRunStatusAsync(
        string runId,
        string status,
        DateTime? completedUtc,
        string? error,
        CancellationToken ct)
    {
        var listId = await GetListIdByNameAsync(RunsList, ct).ConfigureAwait(false);

        var fields = new Dictionary<string, object?>
        {
            ["Status"] = status,
            ["CompletedUtc"] = completedUtc?.ToString("o"),
            ["Error"] = Trunc(error, 2000)
        };

        await UpsertItemAsync(listId, "RunId", runId, fields, ct).ConfigureAwait(false);
    }

    public async Task CreateAutomationStepRunAsync(
        string runId,
        string stepKey,
        string stepId,
        string stepType,
        string? stepName,
        int attempt,
        string status,
        DateTime startedUtc,
        string correlationId,
        CancellationToken ct)
    {
        var listId = await GetListIdByNameAsync(StepRunsList, ct).ConfigureAwait(false);

        var fields = new Dictionary<string, object?>
        {
            ["Title"] = stepKey,           // unique key
            ["RunId"] = runId,
            ["StepId"] = stepId,
            ["StepType"] = stepType,
            ["StepName"] = stepName,
            ["Attempt"] = attempt,
            ["Status"] = status,
            ["StartedUtc"] = startedUtc.ToString("o"),
            ["CorrelationId"] = correlationId
        };

        // Upsert by Title (stepKey)
        await UpsertItemAsync(listId, "Title", stepKey, fields, ct).ConfigureAwait(false);
    }

    public async Task CompleteAutomationStepRunAsync(
        string runId,
        string stepKey,
        string status,
        DateTime completedUtc,
        long durationMs,
        string? outputJson,
        string? error,
        CancellationToken ct)
    {
        var listId = await GetListIdByNameAsync(StepRunsList, ct).ConfigureAwait(false);

        var fields = new Dictionary<string, object?>
        {
            ["RunId"] = runId,
            ["Status"] = status,
            ["CompletedUtc"] = completedUtc.ToString("o"),
            ["DurationMs"] = durationMs,
            ["OutputJson"] = Trunc(outputJson, 10000),
            ["Error"] = Trunc(error, 2000)
        };

        await UpsertItemAsync(listId, "Title", stepKey, fields, ct).ConfigureAwait(false);
    }

    public async Task AppendAutomationAuditAsync(
        string runId,
        string workflowId,
        string eventType,
        string message,
        string correlationId,
        CancellationToken ct)
    {
        var listId = await GetListIdByNameAsync(AuditList, ct).ConfigureAwait(false);

        var auditKey = $"{runId}:{DateTime.UtcNow:yyyyMMddHHmmssfff}:{Guid.NewGuid():N}";

        var fields = new Dictionary<string, object?>
        {
            ["Title"] = auditKey,
            ["RunId"] = runId,
            ["WorkflowId"] = workflowId,
            ["EventType"] = eventType,
            ["Message"] = Trunc(message, 2000),
            ["CreatedUtc"] = DateTime.UtcNow.ToString("o"),
            ["CorrelationId"] = correlationId
        };

        await UpsertItemAsync(listId, "Title", auditKey, fields, ct).ConfigureAwait(false);
    }

    /// <summary>
    /// Generic automation upsert to a target list.
    /// IMPORTANT: listName must be whitelisted by AutomationService before calling this.
    /// </summary>
    public async Task UpsertAutomationTargetItemAsync(
        string listName,
        string keyField,
        string keyValue,
        Dictionary<string, object?> fields,
        CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(listName)) throw new ArgumentException("listName required");
        if (string.IsNullOrWhiteSpace(keyField)) throw new ArgumentException("keyField required");
        if (string.IsNullOrWhiteSpace(keyValue)) throw new ArgumentException("keyValue required");

        var listId = await GetListIdByNameAsync(listName, ct).ConfigureAwait(false);
        await UpsertItemAsync(listId, keyField, keyValue, fields, ct).ConfigureAwait(false);
    }

    private static string? Trunc(string? s, int max)
        => string.IsNullOrWhiteSpace(s) ? s : (s.Length <= max ? s : s[..max] + "...(truncated)");
}