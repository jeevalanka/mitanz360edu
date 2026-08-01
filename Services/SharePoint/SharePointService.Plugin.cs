using Microsoft.Graph;
using System.Text;

namespace MITANZ360Edu.Web.Services;

public partial class SharePointService
{
    public async Task<string> UploadWorkflowTextAsync(
        string fileName,
        string content)
    {
        var driveId =
            await GetLmsDriveIdAsync();

        var userFolder =
            CurrentUserName
                .Replace("@", "_")
                .Replace(".", "_");

        using var stream =
            new MemoryStream(
                Encoding.UTF8.GetBytes(content));

        var uploaded =
            await _graphClient
                .Drives[driveId]
                .Root
                .ItemWithPath(
                    $"AI-Generated-Reports/{userFolder}/{fileName}")
                .Content
                .PutAsync(stream);

        return uploaded?.WebUrl ?? "";
    }
}