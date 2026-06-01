using System.Text.RegularExpressions;

namespace MITANZ360Edu.Web.Services.Automation;

public partial class AutomationService
{
    public class RelevanceResult
    {
        public bool IsRelevant { get; set; }
        public double Score { get; set; }
        public string Reason { get; set; } = string.Empty;
        public List<string> MatchedKeywords { get; set; } = new();
        public List<string> MissingKeywords { get; set; } = new();
    }

    public RelevanceResult ValidateContentRelevance(
        Dictionary<string, string> metadata,
        string fileContent)
    {
        var result = new RelevanceResult();

        if (metadata == null || metadata.Count == 0)
        {
            result.IsRelevant = false;
            result.Score = 0;
            result.Reason = "Metadata is empty.";
            return result;
        }

        // ✅ Extract course context
        metadata.TryGetValue("CourseTitle", out var title);
        metadata.TryGetValue("CourseDescription", out var desc);
        metadata.TryGetValue("CourseLearningOutcomes", out var lo);

        var courseText = $"{title} {desc} {lo}".ToLowerInvariant();

        if (string.IsNullOrWhiteSpace(courseText))
        {
            result.IsRelevant = false;
            result.Score = 0;
            result.Reason = "Course metadata is invalid.";
            return result;
        }

        var contentText = (fileContent ?? string.Empty).ToLowerInvariant();

        // ✅ Extract keywords
        var keywords = ExtractKeywords(courseText);

        if (keywords.Count == 0)
        {
            result.IsRelevant = false;
            result.Score = 0;
            result.Reason = "No meaningful keywords extracted.";
            return result;
        }

        int matchCount = 0;

        foreach (var keyword in keywords)
        {
            if (contentText.Contains(keyword))
            {
                matchCount++;
                result.MatchedKeywords.Add(keyword);
            }
            else
            {
                result.MissingKeywords.Add(keyword);
            }
        }

        // ✅ Score calculation
        result.Score = (double)matchCount / keywords.Count * 100;

        // ✅ Decision threshold
        result.IsRelevant = result.Score >= 30;

        // ✅ Reason
        result.Reason = result.IsRelevant
            ? "Content aligns with course context."
            : "Content does not match course context.";

        return result;
    }

    // ✅ Keyword extractor
    private List<string> ExtractKeywords(string text)
    {
        var words = Regex.Split(text, @"\W+")
            .Where(w => w.Length > 4)
            .Select(w => w.Trim().ToLowerInvariant())
            .Distinct()
            .Take(25)
            .ToList();

        return words;
    }
}