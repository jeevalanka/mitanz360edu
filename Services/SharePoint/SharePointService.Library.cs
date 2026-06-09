using Microsoft.Graph;
using Microsoft.Graph.Models;

namespace MITANZ360Edu.Web.Services;

/// <summary>
/// ============================================================
/// ✅ SHAREPOINT LIBRARY OPERATIONS
/// ============================================================
///
/// Responsibilities:
/// - Library browsing
/// - Folder navigation
/// - File upload
/// - Rename
/// - Delete
/// - Download
/// - LMS metadata extraction
/// - SharePoint Drive CRUD
///
/// ============================================================
/// </summary>
public class LibraryItem
{
    // =====================================================
    // GRAPH / SHAREPOINT IDENTIFIERS
    // =====================================================

    public string Id { get; set; } = string.Empty;
    public string DriveId { get; set; } = string.Empty;
    public string ParentFolderId { get; set; } = string.Empty;

    // =====================================================
    // FILE / FOLDER INFO
    // =====================================================

    public string Name { get; set; } = string.Empty;
    public bool IsFolder { get; set; }
    public long Size { get; set; }
    public string MimeType { get; set; } = string.Empty;
    public string WebUrl { get; set; } = string.Empty;
    public string DownloadUrl { get; set; } = string.Empty;
    public DateTimeOffset? LastModified { get; set; }

    // =====================================================
    // LMS METADATA
    // =====================================================

    public string? Title { get; set; }
    public string ContentTypeCode { get; set; } = string.Empty;
    public string Source { get; set; } = string.Empty;
    public string ContentId { get; set; } = string.Empty;
    public string CourseId { get; set; } = string.Empty;
    public string CourseModelCode { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public bool IsPublished { get; set; }
    public bool IsArchived { get; set; }
    public string MetadataJson { get; set; } = string.Empty;
    public string? AiFeed { get; set; }
}

public partial class SharePointService
{
    // =====================================================
    // ROOT LIBRARY
    // =====================================================

    public async Task<List<LibraryItem>>
        GetLibraryRootAsync()
        {
            try
            {
                var driveId =
                    await GetLmsDriveIdAsync();

                return await GetLibraryFolderChildrenAsync(
                    driveId,
                    "root");
            }
            catch (Exception ex)
            {
                _logger.LogError(
                    ex,
                    "Failed loading LMS library root.");

                throw;
            }
        }

    // =====================================================
    // GET FOLDER CHILDREN
    // =====================================================

    public async Task<List<LibraryItem>>
        GetLibraryFolderChildrenAsync(
            string driveId,
            string folderId)
    {
        try
        {
            // =============================================
            // VALIDATION
            // =============================================

            if (string.IsNullOrWhiteSpace(driveId))
            {
                throw new InvalidOperationException(
                    "DriveId is required.");
            }

            if (string.IsNullOrWhiteSpace(folderId))
            {
                throw new InvalidOperationException(
                    "FolderId is required.");
            }

            // =============================================
            // LOGGING
            // =============================================

            _logger.LogInformation(
                "Loading folder items. DriveId={DriveId}, FolderId={FolderId}",
                driveId,
                folderId);

            // =============================================
            // GRAPH QUERY
            // =============================================

            var response =
                await _graphClient
                    .Drives[driveId]
                    .Items[folderId]
                    .Children
                    .GetAsync(req =>
                    {
                        req.QueryParameters.Expand =
                        [
                            "listItem($expand=fields)"
                        ];
                    });

            var results =
                new List<LibraryItem>();

            // =============================================
            // EMPTY RESULT
            // =============================================

            if (response?.Value == null)
            {
                return results;
            }

            // =============================================
            // MAP RESULTS
            // =============================================

            foreach (var item in response.Value)
            {
                var fields =
                    item.ListItem?
                        .Fields?
                        .AdditionalData;

                results.Add(new LibraryItem
                {
                    // =====================================
                    // CORE
                    // =====================================

                    Id = item.Id ?? string.Empty,

                    Name = item.Name ?? "Unnamed",

                    DriveId = driveId,

                    ParentFolderId =
                        item.ParentReference?.Id
                        ?? string.Empty,

                    WebUrl =
                        item.WebUrl
                        ?? string.Empty,

                    DownloadUrl =
                        item.AdditionalData != null &&
                        item.AdditionalData.ContainsKey(
                            "@microsoft.graph.downloadUrl")
                            ? item.AdditionalData[
                                "@microsoft.graph.downloadUrl"]
                                ?.ToString() ?? string.Empty
                            : string.Empty,

                    IsFolder =
                        item.Folder != null,

                    Size =
                        item.Size ?? 0,

                    MimeType =
                        item.File?.MimeType
                        ?? string.Empty,

                    LastModified =
                        item.LastModifiedDateTime,

                    // =====================================
                    // LMS METADATA
                    // =====================================

                    Title =
                        GetField(fields, "Title"),

                    Description =
                        GetField(fields, "_ExtendedDescription"),

                    ContentId =
                        GetField(fields, "COR_ContentId"),

                    CourseId =
                        GetField(fields, "COR_CourseId"),

                    CourseModelCode =
                            GetField(fields, "COR_CourseModel"),

                    ContentTypeCode =
                        GetField(fields, "COR_ContentType"),

                    IsPublished =
                        GetBoolField(
                            fields,
                            "REL_IsPublished"),

                    IsArchived =
                        GetBoolField(
                            fields,
                            "ARC_IsArchived"),

                    Source =
                        GetField(fields, "GOV_Source"),

                    MetadataJson =
                        GetField(fields, "EXT_Metadata"),

                    AiFeed =
                        GetField(fields, "AI_Feed")
                });
            }

            // =============================================
            // SORT
            // =============================================

            return results
                .OrderByDescending(x => x.IsFolder)
                .ThenBy(x => x.Name)
                .ToList();
        }
        catch (Exception ex)
        {
            _logger.LogError(
                ex,
                "Failed loading folder children.");

            throw;
        }
    }

    // =====================================================
    // CREATE FOLDER
    // =====================================================

    public async Task CreateFolderAsync(
        string driveId,
        string parentFolderId,
        string folderName)
    {
        try
        {
            EnsureUserHasPermission(
                "Library.CreateFolder");

            if (string.IsNullOrWhiteSpace(folderName))
            {
                throw new InvalidOperationException(
                    "Folder name is required.");
            }

            var folder = new DriveItem
            {
                Name = folderName,

                Folder = new Folder(),

                AdditionalData =
                    new Dictionary<string, object>
                    {
                        {
                            "@microsoft.graph.conflictBehavior",
                            "rename"
                        }
                    }
            };

            await _graphClient
                .Drives[driveId]
                .Items[parentFolderId]
                .Children
                .PostAsync(folder);

            _logger.LogInformation(
                "Folder created successfully: {FolderName}",
                folderName);
        }
        catch (Exception ex)
        {
            _logger.LogError(
                ex,
                "Folder creation failed.");

            throw;
        }
    }

    // =====================================================
    // UPLOAD FILE
    // =====================================================

    public async Task<DriveItem> UploadFileAsync(
        string driveId,
        string folderId,
        string fileName,
        Stream stream)
    {
        try
        {
            EnsureUserHasPermission(
                "Library.Upload");

            if (string.IsNullOrWhiteSpace(driveId))
            {
                throw new InvalidOperationException(
                    "DriveId is required.");
            }

            if (string.IsNullOrWhiteSpace(folderId))
            {
                throw new InvalidOperationException(
                    "FolderId is required.");
            }

            if (string.IsNullOrWhiteSpace(fileName))
            {
                throw new InvalidOperationException(
                    "FileName is required.");
            }

            if (stream == null)
            {
                throw new InvalidOperationException(
                    "Upload stream is null.");
            }

            _logger.LogInformation(
                "Uploading file '{FileName}' to folder '{FolderId}'",
                fileName,
                folderId);

            var uploaded =
                await _graphClient
                    .Drives[driveId]
                    .Items[folderId]
                    .ItemWithPath(fileName)
                    .Content
                    .PutAsync(stream);

            if (uploaded == null)
            {
                throw new InvalidOperationException(
                    "SharePoint upload returned null.");
            }

            _logger.LogInformation(
                "File uploaded successfully. ItemId={ItemId}",
                uploaded.Id);

            return uploaded;
        }
        catch (Exception ex)
        {
            _logger.LogError(
                ex,
                "File upload failed.");

            throw;
        }
    }

    // =====================================================
    // UPDATE LIBRARY ITEM
    // =====================================================

    public async Task UpdateLibraryItemAsync(
        string driveId,
        string itemId,
        LibraryItem model)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(driveId))
            {
                throw new InvalidOperationException(
                    "DriveId is required.");
            }

            if (string.IsNullOrWhiteSpace(itemId))
            {
                throw new InvalidOperationException(
                    "ItemId is required.");
            }

            _logger.LogInformation(
                "Updating LMS metadata for item {ItemId}",
                itemId);

            var fieldValues = new FieldValueSet
            {
                AdditionalData =
                    new Dictionary<string, object>
                    {
                        ["Title"] =
                            model.Title ?? string.Empty,

                        ["_ExtendedDescription"] =
                            model.Description ?? string.Empty,

                        ["COR_ContentId"] =
                            model.ContentId ?? string.Empty,

                        ["COR_CourseId"] =
                            model.CourseId ?? string.Empty,

                        ["COR_CourseModel"] = model.CourseModelCode ?? string.Empty,

                        ["COR_ContentType"] =
                            model.ContentTypeCode ?? string.Empty,

                        ["COR_ContentTitle"] =
                            model.Title ?? string.Empty,

                        ["REL_IsPublished"] =
                            model.IsPublished,

                        ["ARC_IsArchived"] =
                            model.IsArchived,

                        ["GOV_Source"] =
                            model.Source ?? string.Empty,

                        ["EXT_Metadata"] =
                            model.MetadataJson ?? string.Empty,

                        ["AI_Feed"] =
                            model.AiFeed ?? string.Empty
                    }
            };

            var listItem =
                await _graphClient
                    .Drives[driveId]
                    .Items[itemId]
                    .ListItem
                    .GetAsync();

            if (listItem?.Id == null)
            {
                throw new Exception(
                    "SharePoint ListItem not found.");
            }

            if (string.IsNullOrWhiteSpace(
                LmsLibraryListId))
            {
                throw new Exception(
                    "LMS ListId configuration missing.");
            }

            await _graphClient
                .Sites[SiteId]
                .Lists[LmsLibraryListId]
                .Items[listItem.Id]
                .Fields
                .PatchAsync(fieldValues);

            _logger.LogInformation(
                "Successfully updated LMS metadata for item {ItemId}",
                itemId);
        }
        catch (Exception ex)
        {
            _logger.LogError(
                ex,
                "Error updating SharePoint library item {ItemId}",
                itemId);

            throw;
        }
    }

    // =====================================================
    // GET ITEM
    // =====================================================

    public async Task<LibraryItem?>
        GetItemAsync(
            string driveId,
            string itemId)
    {
        try
        {
            var item =
                await _graphClient
                    .Drives[driveId]
                    .Items[itemId]
                    .GetAsync(req =>
                    {
                        req.QueryParameters.Expand =
                        [
                            "listItem($expand=fields)"
                        ];
                    });

            if (item == null)
            {
                return null;
            }

            var fields =
                item.ListItem?
                    .Fields?
                    .AdditionalData;

            return new LibraryItem
            {
                Id = item.Id ?? string.Empty,

                Name = item.Name ?? "Unnamed",

                DriveId = driveId,

                ParentFolderId =
                    item.ParentReference?.Id
                    ?? string.Empty,

                WebUrl =
                    item.WebUrl
                    ?? string.Empty,

                DownloadUrl =
                    item.AdditionalData != null &&
                    item.AdditionalData.ContainsKey(
                        "@microsoft.graph.downloadUrl")
                        ? item.AdditionalData[
                            "@microsoft.graph.downloadUrl"]
                            ?.ToString() ?? string.Empty
                        : string.Empty,

                IsFolder =
                    item.Folder != null,

                Size =
                    item.Size ?? 0,

                MimeType =
                    item.File?.MimeType
                    ?? string.Empty,

                LastModified =
                    item.LastModifiedDateTime,

                Title =
                    GetField(fields, "Title"),

                Description =
                    GetField(fields, "_ExtendedDescription"),

                ContentId =
                    GetField(fields, "COR_ContentId"),

                CourseId =
                    GetField(fields, "COR_CourseId"),

                CourseModelCode =
                    GetField(fields, "COR_CourseModel"),

                ContentTypeCode =
                    GetField(fields, "COR_ContentType"),

                IsPublished =
                    GetBoolField(
                        fields,
                        "REL_IsPublished"),

                IsArchived =
                    GetBoolField(
                        fields,
                        "ARC_IsArchived"),

                Source =
                    GetField(fields, "GOV_Source"),

                MetadataJson =
                    GetField(fields, "EXT_Metadata"),

                AiFeed =
                    GetField(fields, "AI_Feed")
            };
        }
        catch (Exception ex)
        {
            _logger.LogError(
                ex,
                "Failed loading item.");

            throw;
        }
    }

    // =====================================================
    // GET ITEM BY ID (RECURSIVE)
    // =====================================================

    public async Task<LibraryItem?> GetLibraryItemByIdAsync(
        string itemId)
    {
        try
        {
            var rootItems =
                await GetLibraryRootAsync();

            return await FindLibraryItemRecursiveAsync(
                rootItems,
                itemId);
        }
        catch (Exception ex)
        {
            _logger.LogError(
                ex,
                "Error loading LMS library item by ID: {ItemId}",
                itemId);

            return null;
        }
    }

    private async Task<LibraryItem?> FindLibraryItemRecursiveAsync(
        IEnumerable<LibraryItem> items,
        string itemId)
    {
        foreach (var item in items)
        {
            if (item.Id == itemId)
            {
                return item;
            }

            if (item.IsFolder)
            {
                var children =
                    await GetLibraryFolderChildrenAsync(
                        item.DriveId,
                        item.Id);

                var found =
                    await FindLibraryItemRecursiveAsync(
                        children,
                        itemId);

                if (found != null)
                {
                    return found;
                }
            }
        }

        return null;
    }

    // =====================================================
    // ROOT CONTEXT
    // =====================================================

    public class LibraryRootContext
    {
        public string DriveId { get; set; } = string.Empty;

        public string FolderId { get; set; } = "root";

        public List<LibraryItem> Items { get; set; } = [];
    }

    public async Task<LibraryRootContext>
        GetRootContextAsync()
    {
        var driveId =
            await GetLmsDriveIdAsync();

        var items =
            await GetLibraryFolderChildrenAsync(
                driveId,
                "root");

        return new LibraryRootContext
        {
            DriveId = driveId,
            FolderId = "root",
            Items = items
        };
    }

    // =====================================================
    // DELETE ITEM
    // =====================================================

    public async Task DeleteItemAsync(
        string driveId,
        string itemId)
    {
        try
        {
            EnsureUserHasPermission(
                "Library.Delete");

            await _graphClient
                .Drives[driveId]
                .Items[itemId]
                .DeleteAsync();

            _logger.LogInformation(
                "Item deleted successfully.");
        }
        catch (Exception ex)
        {
            _logger.LogError(
                ex,
                "Delete failed.");

            throw;
        }
    }


}