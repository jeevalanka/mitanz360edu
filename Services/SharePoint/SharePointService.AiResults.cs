using Microsoft.Graph;
using MITANZ360Edu.Web.Services.AI;
using System.Text;
using System.Text.Json;

namespace MITANZ360Edu.Web.Services
{
    public partial class SharePointService
    {
        public async Task SaveAIResultAsync(
            string workflowId,
            string targetEntityId,
            AiWorkflowResult result,
            string? renderedHtml,
            string? summaryText,
            CancellationToken cancellationToken)
        {
            if (string.IsNullOrWhiteSpace(workflowId))
                throw new ArgumentException(nameof(workflowId));

            if (string.IsNullOrWhiteSpace(targetEntityId))
                throw new ArgumentException(nameof(targetEntityId));

            if (result == null)
                throw new ArgumentNullException(nameof(result));

            var executionId = targetEntityId;

            var folderPath = $"AI-Results/{workflowId}/{executionId}";

            var json = JsonSerializer.Serialize(
                    result,
                    new JsonSerializerOptions
                    {
                        WriteIndented = true
                    });

            await UploadTextFileAsync( $"{folderPath}/result.json", json, cancellationToken);

            if (!string.IsNullOrWhiteSpace(renderedHtml))
            {
                await UploadTextFileAsync(
                    $"{folderPath}/result.html",
                    renderedHtml,
                    cancellationToken);
            }

            if (!string.IsNullOrWhiteSpace(summaryText))
            {
                await UploadTextFileAsync(
                    $"{folderPath}/summary.txt",
                    summaryText,
                    cancellationToken);
            }

            _logger.LogInformation(
                "AI artifacts saved | Workflow={WorkflowId} | Execution={ExecutionId}",
                workflowId,
                executionId);
        }

        private async Task UploadTextFileAsync(
            string path,
            string content,
            CancellationToken cancellationToken)
        {
            var driveId =
                _configuration["SharePoint:Libraries:LMS:DriveId"]
                ?? throw new InvalidOperationException(
                    "SharePoint LMS DriveId is missing.");

            using var stream =
                new MemoryStream(
                    Encoding.UTF8.GetBytes(content));

            await _graphClient
                .Drives[driveId]
                .Root
                .ItemWithPath(path)
                .Content
                .PutAsync(
                    stream,
                    cancellationToken: cancellationToken);
        }
    }
}