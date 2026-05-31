using System.Text;
using System.Text.Json;
using MITANZ360Edu.Web.Services.AI;

namespace MITANZ360Edu.Web.Services.Automation
{
    public sealed class AutomationService
    {
        private readonly AiWorkflowEngine _workflowEngine;

        public AutomationService(AiWorkflowEngine workflowEngine)
        {
            _workflowEngine = workflowEngine;
        }

        /// <summary>
        /// Executes AI Feedback (TEXT ONLY)
        /// </summary>
        public async Task<string> GenerateAIFeedbackAsync(
            Dictionary<string, string> metadata,
            string fileContent,
            CancellationToken cancellationToken = default)
        {
            var promptBuilder = new StringBuilder();

            promptBuilder.AppendLine("You are an Academic Quality Assurance Reviewer.");
            promptBuilder.AppendLine();
            promptBuilder.AppendLine("You will receive:");
            promptBuilder.AppendLine("1. Course Metadata");
            promptBuilder.AppendLine("2. Uploaded Document Content");
            promptBuilder.AppendLine();

            promptBuilder.AppendLine("Your task is to evaluate the document against the metadata.");
            promptBuilder.AppendLine();

            promptBuilder.AppendLine("COURSE METADATA:");
            foreach (var item in metadata)
            {
                promptBuilder.AppendLine($"{item.Key}: {item.Value}");
            }

            promptBuilder.AppendLine();
            promptBuilder.AppendLine("DOCUMENT CONTENT:");
            promptBuilder.AppendLine(fileContent);

            promptBuilder.AppendLine();
            promptBuilder.AppendLine("Evaluate the following:");
            promptBuilder.AppendLine("1. Alignment with Course Description");
            promptBuilder.AppendLine("2. Coverage of Learning Outcomes");
            promptBuilder.AppendLine("3. Missing topics");
            promptBuilder.AppendLine("4. Weak content areas");
            promptBuilder.AppendLine("5. Content quality and academic suitability");

            promptBuilder.AppendLine();
            promptBuilder.AppendLine("Return EXACTLY in this format:");
            promptBuilder.AppendLine();

            promptBuilder.AppendLine("DOCUMENT EVALUATION REPORT");
            promptBuilder.AppendLine();
            promptBuilder.AppendLine("Course: <Course Title>");
            promptBuilder.AppendLine();
            promptBuilder.AppendLine("Alignment Score: <Percentage>");
            promptBuilder.AppendLine();
            promptBuilder.AppendLine("Course Description Assessment: <Assessment>");
            promptBuilder.AppendLine();
            promptBuilder.AppendLine("Learning Outcomes Covered:");
            promptBuilder.AppendLine("* Item 1");
            promptBuilder.AppendLine("* Item 2");
            promptBuilder.AppendLine();
            promptBuilder.AppendLine("Learning Outcomes Partially Covered:");
            promptBuilder.AppendLine("* Item 1");
            promptBuilder.AppendLine("* Item 2");
            promptBuilder.AppendLine();
            promptBuilder.AppendLine("Learning Outcomes Missing:");
            promptBuilder.AppendLine("* Item 1");
            promptBuilder.AppendLine("* Item 2");
            promptBuilder.AppendLine();
            promptBuilder.AppendLine("Content Strengths:");
            promptBuilder.AppendLine("* Item 1");
            promptBuilder.AppendLine("* Item 2");
            promptBuilder.AppendLine();
            promptBuilder.AppendLine("Content Weaknesses:");
            promptBuilder.AppendLine("* Item 1");
            promptBuilder.AppendLine("* Item 2");
            promptBuilder.AppendLine();
            promptBuilder.AppendLine("Recommendations:");
            promptBuilder.AppendLine("* Item 1");
            promptBuilder.AppendLine("* Item 2");
            promptBuilder.AppendLine();
            promptBuilder.AppendLine("Final Assessment:");
            promptBuilder.AppendLine("<Summary>");
            promptBuilder.AppendLine();
            promptBuilder.AppendLine("Important:");
            promptBuilder.AppendLine("- Do NOT return JSON");
            promptBuilder.AppendLine("- Do NOT return HTML");
            promptBuilder.AppendLine("- Only plain text");
            promptBuilder.AppendLine("- Follow format strictly");

            var request = new AiWorkflowRequest
            {
                TaskType = AiTaskType.DocumentAnalysis, // ensure valid enum
                Prompt = promptBuilder.ToString(),
                StrictJsonResponse = false, // ✅ IMPORTANT
                Temperature = 0.2
            };

            // ✅ 2. Call AI
            var aiResult = await _workflowEngine.ExecuteAsync(request, cancellationToken);

            if (!aiResult.Success)
                return "AI processing failed.";

            // ✅ 3. Extract TEXT ONLY
            return ExtractTextOnly(aiResult);
        }

        /// <summary>
        /// Extract final TEXT from AI response
        /// </summary>
        private string ExtractTextOnly(AiWorkflowResult result)
        {
            // ✅ 1. Safety checks
            if (result == null || !result.Success)
                return "AI processing failed.";

            if (string.IsNullOrWhiteSpace(result.RawResponse))
                return "AI returned empty response.";

            try
            {
                using var doc = JsonDocument.Parse(result.RawResponse);

                // ✅ 2. SAFE NAVIGATION (no crashes)
                if (!doc.RootElement.TryGetProperty("choices", out var choices) ||
                    choices.ValueKind != JsonValueKind.Array ||
                    choices.GetArrayLength() == 0)
                {
                    return result.RawResponse; // fallback
                }

                var firstChoice = choices[0];

                if (!firstChoice.TryGetProperty("message", out var message))
                {
                    return result.RawResponse;
                }

                if (!message.TryGetProperty("content", out var contentElement))
                {
                    return result.RawResponse;
                }

                var content = contentElement.GetString();

                if (string.IsNullOrWhiteSpace(content))
                    return "AI returned empty content.";

                // ✅ 3. Clean formatting
                return content
                    .Replace("\\r", "")
                    .Replace("\\n", Environment.NewLine)
                    .Trim();
            }
            catch
            {
                // ✅ 4. NEVER break UI — always show something
                return result.RawResponse;
            }
        }
    }
}