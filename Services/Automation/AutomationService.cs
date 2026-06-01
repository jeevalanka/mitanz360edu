using System.Text;
using MITANZ360Edu.Web.Services.AI;

namespace MITANZ360Edu.Web.Services.Automation;

public partial class AutomationService
{
    private readonly AIService _aiService;

    public AutomationService(AIService aiService)
    {
        _aiService = aiService;
    }

    /// <summary>
    /// Main entry point for AI evaluation
    /// </summary>
    public async Task<string> GenerateAIFeedbackAsync(
        Dictionary<string, string> metadata,
        string fileContent)
    {
        // ✅ STEP 1 — Validate inputs
        if (string.IsNullOrWhiteSpace(fileContent))
        {
            return "⚠ No content available for AI evaluation.";
        }

        if (metadata == null || metadata.Count == 0)
        {
            return "⚠ Metadata is missing. Cannot evaluate document.";
        }

        // ✅ STEP 2 — RELEVANCE VALIDATION (CRITICAL)
        var relevance = ValidateContentRelevance(metadata, fileContent);

        if (!relevance.IsRelevant)
        {
            return BuildRelevanceFailureMessage(relevance);
        }

        // ✅ STEP 3 — CLEAN METADATA
        var courseTitle = GetSafe(metadata, "CourseTitle");
        var courseDescription = GetSafe(metadata, "CourseDescription");
        var learningOutcomes = GetSafe(metadata, "CourseLearningOutcomes");

        // ✅ STEP 4 — BUILD PROMPT
        var prompt = BuildEvaluationPrompt(
            courseTitle,
            courseDescription,
            learningOutcomes,
            fileContent,
            relevance.Score);

        // ✅ STEP 5 — CALL AI SERVICE
        var aiResult = await _aiService.GenerateTextAsync(prompt);

        if (string.IsNullOrWhiteSpace(aiResult))
        {
            return "⚠ AI returned empty response.";
        }

        // ✅ STEP 6 — RETURN OUTPUT
        return aiResult.Trim();
    }

    // =========================================================
    // ✅ PRIVATE HELPERS
    // =========================================================

    private string BuildEvaluationPrompt(
        string courseTitle,
        string courseDescription,
        string learningOutcomes,
        string fileContent,
        double relevanceScore)
    {
        var sb = new StringBuilder();

        sb.AppendLine("You are an academic evaluation engine.");
        sb.AppendLine("Evaluate the document against the provided course.");
        sb.AppendLine();

        sb.AppendLine("=== COURSE INFORMATION ===");
        sb.AppendLine($"Course: {courseTitle}");
        sb.AppendLine($"Description: {courseDescription}");
        sb.AppendLine("Learning Outcomes:");
        sb.AppendLine(learningOutcomes);
        sb.AppendLine();

        sb.AppendLine("=== DOCUMENT CONTENT ===");
        sb.AppendLine(fileContent);
        sb.AppendLine();

        sb.AppendLine("=== INSTRUCTIONS ===");
        sb.AppendLine("- Provide an alignment score (0–100%)");
        sb.AppendLine("- Evaluate coverage of learning outcomes");
        sb.AppendLine("- Identify strengths and weaknesses");
        sb.AppendLine("- Provide actionable recommendations");
        sb.AppendLine("- Keep academic tone");
        sb.AppendLine("- Be structured in sections");

        sb.AppendLine();
        sb.AppendLine($"(Relevance Score Pre-Check: {relevanceScore:0}%)");

        return sb.ToString();
    }

    private string BuildRelevanceFailureMessage(RelevanceResult relevance)
    {
        return $@"
        ⚠ RELEVANCE VALIDATION FAILED

        Score: {relevance.Score:0}%

        Reason:
        {relevance.Reason}

        Matched Keywords:
        - {string.Join("\n- ", relevance.MatchedKeywords.Take(5))}

        Missing Keywords:
        - {string.Join("\n- ", relevance.MissingKeywords.Take(10))}

        👉 ACTION REQUIRED:
        - Select the correct course
        - OR upload a document relevant to the selected course
        ";
    }

    private string GetSafe(Dictionary<string, string> metadata, string key)
    {
        return metadata.TryGetValue(key, out var value)
            ? value ?? string.Empty
            : string.Empty;
    }
}