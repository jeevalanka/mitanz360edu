using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Extensions.Logging;
using System.Text.Json.Nodes;

using MITANZ360Edu.Web.Services;
using MITANZ360Edu.Web.Services.AI;

namespace MITANZ360Edu.Web.Services.Automation;

public partial class AutomationService
{
    private readonly AIService _aiService;
    private readonly SharePointService _sharePointService;
    private readonly ILogger<AutomationService> _logger;

    public AutomationService(
        AIService aiService,
        SharePointService sharePointService,
        ILogger<AutomationService> logger)
    {
        _aiService = aiService;
        _sharePointService = sharePointService;
        _logger = logger;
    }

    // =========================================================
    // ✅ EXISTING FUNCTION (UNCHANGED)
    // =========================================================

    public async Task<string> GenerateAIFeedbackAsync(
        Dictionary<string, string> metadata,
        string fileContent)
    {
        if (string.IsNullOrWhiteSpace(fileContent))
            return "⚠ No content available for AI evaluation.";

        if (metadata == null || metadata.Count == 0)
            return "⚠ Metadata is missing.";

        var relevance = ValidateContentRelevance(metadata, fileContent);

        if (!relevance.IsRelevant)
            return BuildRelevanceFailureMessage(relevance);

        var prompt = BuildEvaluationPrompt(
            GetSafe(metadata, "CourseTitle"),
            GetSafe(metadata, "CourseDescription"),
            GetSafe(metadata, "CourseLearningOutcomes"),
            fileContent,
            relevance.Score);

        var aiResult = await _aiService.GenerateTextAsync(prompt);

        return string.IsNullOrWhiteSpace(aiResult)
            ? "⚠ AI returned empty response."
            : aiResult.Trim();
    }

    // =========================================================
    // ✅ NEW: COURSE BUILDER AI WORKFLOW
    // =========================================================

    public async Task RunCourseWorkflowAsync(
        AIRepositoryItem item,
        JsonObject json,
        CancellationToken ct = default)
    {
        try
        {
            _logger.LogInformation("🚀 AI Course Workflow started: {Id}", item.Id);

            // ✅ STEP 1 — MARK RUNNING
            item.Status = "Running";
            await _sharePointService.UpdateAIRepositoryItemAsync(item, ct);

            // ✅ STEP 2 — CALL AI SERVICE
            var result = await _aiService.GenerateCourseAsync(json);

            if (result == null)
                throw new Exception("AI returned null result");

            // ✅ STEP 3 — UPDATE FIELDS
            item.Summary = result.Summary;
            item.HtmlReport = result.Html;
            item.Tags = result.Tags;
            item.Score = (int)Math.Clamp(result.Score, 0, 100);
            item.Metadata = result.UpdatedJson;
            item.Status = "Completed";

            await _sharePointService.UpdateAIRepositoryItemAsync(item, ct);

            // ✅ STEP 4 — UPLOAD DOCX
            if (result.DocxBytes?.Length > 0)
            {
                await _sharePointService.UploadCourseDocAsync(
                    item.Id,
                    result.DocxBytes,
                    $"{item.Title}.docx");
            }

            // ✅ STEP 5 — AUDIT ✅ FIXED HERE
            await _sharePointService.WriteAuditPublicAsync(
                "AI Workflow Completed",
                "AIRepository",
                item.Id.ToString(),
                ct);

            _logger.LogInformation("✅ AI Course Workflow completed: {Id}", item.Id);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "❌ Workflow failed: {Id}", item.Id);

            item.Status = "Failed";

            try
            {
                await _sharePointService.UpdateAIRepositoryItemAsync(item, ct);
            }
            catch { }

            throw;
        }
    }
    // =========================================================
    // ✅ RELEVANCE VALIDATION (UNCHANGED)
    // =========================================================

    private RelevanceResult ValidateContentRelevance(
        Dictionary<string, string> metadata,
        string content)
    {
        var result = new RelevanceResult();

        metadata.TryGetValue("CourseTitle", out var title);
        metadata.TryGetValue("CourseDescription", out var description);
        metadata.TryGetValue("CourseLearningOutcomes", out var outcomes);

        var courseText = $"{title} {description} {outcomes}".ToLower();

        var keywords = ExtractKeywords(courseText);

        int matchCount = keywords.Count(k =>
            content.Contains(k, StringComparison.OrdinalIgnoreCase));

        result.Score = keywords.Count == 0
            ? 0
            : (int)((double)matchCount / keywords.Count * 100);

        result.IsRelevant = result.Score >= 50;

        return result;
    }
    private List<string> ExtractKeywords(string text)
    {
        return Regex
            .Split(text, @"\W+")
            .Where(w => w.Length > 4)
            .Distinct()
            .ToList();
    }

    // =========================================================
    // ✅ HELPERS
    // =========================================================

    private string BuildEvaluationPrompt(
        string courseTitle,
        string courseDescription,
        string learningOutcomes,
        string fileContent,
        double relevanceScore)
    {
        var sb = new StringBuilder();

        sb.AppendLine("Evaluate the document academically:");
        sb.AppendLine($"Course: {courseTitle}");
        sb.AppendLine($"Relevance: {relevanceScore}%");

        sb.AppendLine(fileContent);

        return sb.ToString();
    }
    private string BuildRelevanceFailureMessage(RelevanceResult relevance)
    {
        return $"⚠ RELEVANCE FAILED ({relevance.Score}%)";
    }
    private string GetSafe(Dictionary<string, string> metadata, string key)
    {
        return metadata.TryGetValue(key, out var value)
            ? value ?? ""
            : "";
    }
    public async Task RunCourseFromMetadataAsync(string itemId, CancellationToken ct = default)
    {
        try
        {
            _logger.LogInformation("🚀 Metadata-based AI workflow started: {Id}", itemId);

            // ✅ ONLY convert for list updates
            var spItemId = int.Parse(itemId);

            // ✅ STEP 1 — READ METADATA (string ✅)
            var metadataJson = await _sharePointService.GetMetadataAsync(itemId);

            if (string.IsNullOrWhiteSpace(metadataJson))
                throw new Exception("Metadata is empty");

            // ✅ STEP 2 — PARSE
            var json = JsonNode.Parse(metadataJson)?.AsObject();

            if (json == null)
                throw new Exception("Invalid metadata JSON");

            // ✅ STEP 3 — RUNNING (int ✅)
            await _sharePointService.UpdateAIResultFieldsAsync(
                spItemId,
                0,
                "Running",
                "",
                "",
                ""
            );

            // ✅ STEP 4 — AI
            var result = await _aiService.GenerateCourseAsync(json);

            if (result == null)
                throw new Exception("AI returned null");

            var score = (int)Math.Round(result.Score);

            // ✅ STEP 5 — UPDATE LIST (int ✅)
            await _sharePointService.UpdateAIResultFieldsAsync(
                spItemId,
                score,
                "Completed",
                result.Html,
                result.Summary,
                result.Tags
            );

            // ✅ STEP 6 — UPDATE METADATA (string ✅)
            if (!string.IsNullOrWhiteSpace(result.UpdatedJson))
            {
                await _sharePointService.UpdateMetadataAsync(itemId, result.UpdatedJson);
            }

            // ✅ STEP 7 — UPLOAD DOC (string ✅)
            if (result.DocxBytes?.Length > 0)
            {
                await _sharePointService.UploadCourseDocAsync(
                    itemId,
                    result.DocxBytes,
                    $"AI-{itemId}.docx"
                );
            }

            // ✅ ✅ FIX: CREATE VALID WORKFLOW OBJECT
            var workflowResult = new AiWorkflowResult
            {
                // ✅ SAFE: store full AI result as JSON
                Data = JsonSerializer.Serialize(result)
            };

            // ✅ STEP 8 — SAVE FILES (string ✅)
            await _sharePointService.SaveAIResultAsync(
                "CourseBuilder",
                itemId,
                workflowResult,
                result.Html,
                result.Summary,
                ct
            );

            _logger.LogInformation("✅ Metadata AI workflow completed: {Id}", itemId);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "❌ Metadata workflow failed: {Id}", itemId);

            if (int.TryParse(itemId, out var spItemId))
            {
                await _sharePointService.UpdateAIResultFieldsAsync(
                    spItemId,
                    0,
                    "Failed",
                    "",
                    ex.Message,
                    "Error"
                );
            }

            throw;
        }
    }
``    private AiWorkflowResult MapToWorkflowResult(AIResultDto dto)
    {
        return new AiWorkflowResult
        {
            Summary = dto.Summary,
            Html = dto.Html,
            Tags = dto.Tags,
            Score = dto.Score,
            // add more fields if your model requires
        };
    }

}