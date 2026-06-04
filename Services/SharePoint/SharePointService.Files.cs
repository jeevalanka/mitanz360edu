using Microsoft.Graph;
using System.Text;

namespace MITANZ360Edu.Web.Services;

public partial class SharePointService
{
    // =====================================================
    // UPLOAD Course Doc
    // =====================================================
            public async Task UploadCourseDocAsync(
                int itemId,
                byte[] fileBytes,
                string fileName)
            {
                var driveId = await GetLmsDriveIdAsync();

                using var stream = new MemoryStream(fileBytes);

                await _graphClient
                    .Drives[driveId]
                    .Root
                    .ItemWithPath($"AIRepository/{fileName}")
                    .Content
                    .PutAsync(stream);
            }


    // =====================================================
    // DOWNLOAD FILE BY DRIVE ITEM ID
    // =====================================================
    public async Task<Stream> DownloadFileAsync(
        string driveItemId,
        CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(driveItemId))
            throw new ArgumentNullException(nameof(driveItemId));

        var driveId = await GetLmsDriveIdAsync();

        _logger.LogInformation(
            "Downloading file | DriveId={DriveId} DriveItemId={DriveItemId}",
            driveId,
            driveItemId);

        return await ExecuteWithRetryAsync(
            async token =>
            {
                var stream =
                    await _graphClient
                        .Drives[driveId]
                        .Items[driveItemId]
                        .Content
                        .GetAsync(
                            requestConfig => { },
                            token);

                if (stream == null)
                    throw new InvalidOperationException(
                        "Downloaded file stream was null.");

                return stream;
            },
            "DownloadFileAsync",
            ct);
    }

    public async Task<Stream?> DownloadFileAsync(string driveId, string itemId)
    {
        try
        {
            var stream = await _graphClient
                .Drives[driveId]
                .Items[itemId]
                .Content
                .GetAsync();

            return stream;
        }
        catch
        {
            return null;
        }
    }

    // =====================================================
    // DOWNLOAD FILE BY FILE NAME (LMS LIB ROOT)
    // =====================================================
    public async Task<Stream> DownloadFileByNameAsync(
        string fileName,
        CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(fileName))
            throw new ArgumentNullException(nameof(fileName));

        var driveId = await GetLmsDriveIdAsync();

        _logger.LogInformation(
            "Resolving DriveItem by name | DriveId={DriveId} File={File}",
            driveId,
            fileName);

        var driveItem =
            await _graphClient
                .Drives[driveId]
                .Root
                .ItemWithPath(fileName)
                .GetAsync(
                    requestConfig => { },
                    ct);

        if (driveItem == null || string.IsNullOrWhiteSpace(driveItem.Id))
            throw new InvalidOperationException(
                $"Drive item not found for file '{fileName}'.");

        _logger.LogInformation(
            "Resolved DriveItemId={DriveItemId} for file={File}",
            driveItem.Id,
            fileName);

        return await DownloadFileAsync(driveItem.Id, ct);
    }

    // =====================================================
    // ✅ READ AI HTML REPORT (AI/LMS/{ContentId}/result.html)
    // =====================================================
    public async Task<string?> GetAIResultHtmlAsync(
        string workflowId,
        string executionId,
        CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(workflowId))
            return null;

        if (string.IsNullOrWhiteSpace(executionId))
            return null;

        var driveId =
            await GetLmsDriveIdAsync();

        var path = $"AI-Results/{workflowId}/{executionId}/result.html";

        _logger.LogInformation(
            "Reading AI HTML report | DriveId={DriveId} Path={Path}",
            driveId,
            path);

        return await ExecuteWithRetryAsync(
            async token =>
            {
                try
                {
                    using var stream =
                        await _graphClient
                            .Drives[driveId]
                            .Root
                            .ItemWithPath(path)
                            .Content
                            .GetAsync(
                                requestConfig => { },
                                token);

                    if (stream == null)
                        return null;

                    using var reader =
                        new StreamReader(
                            stream,
                            Encoding.UTF8);

                    return await reader.ReadToEndAsync();
                }
                catch (ServiceException ex)
                {
                    _logger.LogWarning(
                        ex,
                        "AI HTML report not found | Workflow={WorkflowId} | Execution={ExecutionId}",
                        workflowId,
                        executionId);

                    return null;
                }
            },
            "GetAIResultHtmlAsync",
            ct);
    }
    // =====================================================
    // ✅ READ AI JSON RESULT (AI/LMS/{ContentId}/result.json)
    // =====================================================
    public async Task<string?> GetAIResultJsonAsync(
        string contentId,
        CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(contentId))
            return null;

        var driveId = await GetLmsDriveIdAsync();
        var path = $"AI/LMS/{contentId}/result.json";

        _logger.LogInformation(
            "Reading AI JSON result | DriveId={DriveId} Path={Path}",
            driveId,
            path);

        return await ExecuteWithRetryAsync(
            async token =>
            {
                try
                {
                    using var stream =
                        await _graphClient
                            .Drives[driveId]
                            .Root
                            .ItemWithPath(path)
                            .Content
                            .GetAsync(
                                requestConfig => { },
                                token);

                    if (stream == null)
                        return null;

                    using var reader =
                        new StreamReader(stream, Encoding.UTF8);

                    return await reader.ReadToEndAsync();
                }
                catch (ServiceException ex)
                {
                    _logger.LogWarning(
                        ex,
                        "AI JSON result not found | ContentId={ContentId}",
                        contentId);

                    return null;
                }
            },
            "GetAIResultJsonAsync",
            ct);
    }



    /// <summary>
    /// Extract raw text content from SharePoint file
    /// </summary>
    public async Task<string> ExtractDocumentContentAsync(
        string driveId,
        string itemId)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(driveId) ||
                string.IsNullOrWhiteSpace(itemId))
            {
                return string.Empty;
            }

            var stream = await _graphClient
                .Drives[driveId]
                .Items[itemId]
                .Content
                .GetAsync();

            if (stream == null)
            {
                return string.Empty;
            }

            using var reader = new StreamReader(stream, Encoding.UTF8);
            return await reader.ReadToEndAsync();
        }
        catch
        {
            // ✅ never break UI
            return string.Empty;
        }
    }

}
