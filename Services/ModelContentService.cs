using Microsoft.Graph;
using Microsoft.Graph.Models;
using MITANZ360Edu.Web.Components.Pages.ModelContent;
using MITANZ360Edu.Web.Models;

namespace MITANZ360Edu.Web.Services;

public class ModelContentService
{
    private readonly GraphServiceClient _graphClient;
    private readonly IConfiguration _configuration;
    private readonly ILogger<ModelContentService> _logger;

    private string SiteId =>
        _configuration["SharePoint:SiteId"]
        ?? throw new InvalidOperationException(
            "SharePoint:SiteId missing.");

    private string ListId =>
        _configuration["SharePoint:Lists:ModelContent"]
        ?? throw new InvalidOperationException(
            "SharePoint:Lists:ModelContent missing.");

    public ModelContentService(
        GraphServiceClient graphClient,
        IConfiguration configuration,
        ILogger<ModelContentService> logger)
    {
        _graphClient = graphClient;
        _configuration = configuration;
        _logger = logger;
    }

    // =====================================================
    // GET ALL
    // =====================================================

    public async Task<List<ModelContentModel>>
        GetAllAsync(
            CancellationToken cancellationToken = default)
    {
        var results =
            new List<ModelContentModel>();

        try
        {
            var response =
                await _graphClient
                    .Sites[SiteId]
                    .Lists[ListId]
                    .Items
                    .GetAsync(config =>
                    {
                        config.QueryParameters.Expand =
                            ["fields"];

                        config.QueryParameters.Top = 500;
                    },
                    cancellationToken);

            if (response?.Value == null)
            {
                return results;
            }

            foreach (var item in response.Value)
            {
                if (item.Fields?.AdditionalData == null)
                {
                    continue;
                }

                var f =
                    item.Fields.AdditionalData;

                results.Add(
                    new ModelContentModel
                    {
                        // =====================================================
                        // SYSTEM
                        // =====================================================

                        Id =
                            int.TryParse(
                                item.Id,
                                out var id)
                                    ? id
                                    : 0,

                        // =====================================================
                        // BASIC
                        // =====================================================

                        Title =
                            GetString(f, "Title"),

                        CourseId =
                            GetString(f, "CourseId"),

                        ModuleNo =
                            GetString(f, "ModuleNo"),

                        // =====================================================
                        // CONTENT
                        // =====================================================

                        ContentBody =
                            GetString(f, "ContentBody"),

                        LearningOutcomes =
                            GetString(f, "LearningOutcomes"),

                        // =====================================================
                        // TYPE
                        // =====================================================
                        ContentTypeCode = GetString(f, "ContentType0") ?? "",

                        // =====================================================
                        // RESOURCES
                        // =====================================================

                        FileUrl =
                            GetString(f, "FileUrl"),

                        ResourceUrl =
                            GetString(f, "ResourceUrl"),

                        ParentId =
                            GetString(f, "ParentId"),

                        // =====================================================
                        // FLAGS
                        // =====================================================

                        IsPublished =
                            GetBool(
                                f,
                                "IsPublished"),

                        IsMandatory =
                            GetBool(
                                f,
                                "IsMandatory"),

                        IsVisibleToStudents =
                            GetBool(
                                f,
                                "IsVisibleToStudents"),

                        // =====================================================
                        // METADATA
                        // =====================================================

                        DurationMinutes =
                            GetDecimal(
                                f,
                                "DurationMinutes"),

                        Order =
                            GetInt(
                                f,
                                "Order"),

                        // =====================================================
                        // AI
                        // =====================================================

                        IsAiGenerated =
                            GetBool(
                                f,
                                "IsAiGenerated"),

                        AiSummary =
                            GetString(
                                f,
                                "AiSummary"),

                        AiTags =
                            GetString(
                                f,
                                "AiTags"),

                        AiMetadata =
                            GetString(
                                f,
                                "AiMetadata"),

                        // =====================================================
                        // AUDIT
                        // =====================================================

                        CreatedBy =
                            GetString(
                                f,
                                "CreatedBy"),

                        ModifiedBy =
                            GetString(
                                f,
                                "ModifiedBy"),

                        PublishedBy =
                            GetString(
                                f,
                                "PublishedBy")
                    });
            }

            return results
                .OrderBy(x => x.Order)
                .ThenBy(x => x.Title)
                .ToList();
        }
        catch (Exception ex)
        {
            _logger.LogError(
                ex,
                "Failed loading ModelContent.");

            return [];
        }
    }

    // =====================================================
    // CREATE
    // =====================================================

    public async Task<bool>
        CreateAsync(
            ModelContentModel model,
            CancellationToken cancellationToken = default)
    {
        try
        {
            model.Normalize();

            _logger.LogInformation(
                "Creating ModelContent with ContentType: {ContentType}",
                GetContentTypeValue(model));

            var item =
                new ListItem
                {
                    Fields =
                        new FieldValueSet
                        {
                            AdditionalData =
                                BuildFieldDictionary(model)
                        }
                };

            await _graphClient
                .Sites[SiteId]
                .Lists[ListId]
                .Items
                .PostAsync(
                    item,
                    cancellationToken:
                    cancellationToken);

            _logger.LogInformation(
                "ModelContent created successfully.");

            return true;
        }
        catch (Exception ex)
        {
            _logger.LogError(
                ex,
                "Create failed.");

            return false;
        }
    }

    // =====================================================
    // UPDATE
    // =====================================================

    public async Task<bool>
        UpdateAsync(
            ModelContentModel model,
            CancellationToken cancellationToken = default)
    {
        try
        {
            model.Normalize();

            _logger.LogInformation(
                "Updating ModelContent Id: {Id} with ContentType: {ContentType}",
                model.Id,
                GetContentTypeValue(model));

            await _graphClient
                .Sites[SiteId]
                .Lists[ListId]
                .Items[model.Id.ToString()]
                .Fields
                .PatchAsync(
                    new FieldValueSet
                    {
                        AdditionalData =
                            BuildFieldDictionary(model)
                    },
                    cancellationToken:
                    cancellationToken);

            _logger.LogInformation(
                "ModelContent updated successfully. Id: {Id}",
                model.Id);

            return true;
        }
        catch (Exception ex)
        {
            _logger.LogError(
                ex,
                "Update failed. Id: {Id}",
                model.Id);

            return false;
        }
    }

    // =====================================================
    // DELETE
    // =====================================================

    public async Task<bool>
        DeleteAsync(
            int id,
            CancellationToken cancellationToken = default)
    {
        try
        {
            await _graphClient
                .Sites[SiteId]
                .Lists[ListId]
                .Items[id.ToString()]
                .DeleteAsync(
                    cancellationToken:
                    cancellationToken);

            _logger.LogInformation(
                "ModelContent deleted. Id: {Id}",
                id);

            return true;
        }
        catch (Exception ex)
        {
            _logger.LogError(
                ex,
                "Delete failed. Id: {Id}",
                id);

            return false;
        }
    }

    // =====================================================
    // FIELD BUILDER
    // =====================================================

    private Dictionary<string, object>
        BuildFieldDictionary(
            ModelContentModel model)
    {
        return new Dictionary<string, object>
        {
            // =====================================================
            // BASIC
            // =====================================================

            ["Title"] =
                model.Title,

            ["CourseId"] =
                model.CourseId,

            ["ModuleNo"] =
                model.ModuleNo,

            // =====================================================
            // SHAREPOINT INTERNAL FIELD
            // =====================================================

            ["ContentType0"] =
                GetContentTypeValue(model),

            // =====================================================
            // CONTENT
            // =====================================================

            ["ContentBody"] =
                model.ContentBody,

            ["LearningOutcomes"] =
                model.LearningOutcomes,

            // =====================================================
            // RESOURCES
            // =====================================================

            ["FileUrl"] =
                model.FileUrl,

            ["ResourceUrl"] =
                model.ResourceUrl,

            ["ParentId"] =
                model.ParentId,

            // =====================================================
            // METADATA
            // =====================================================

            ["DurationMinutes"] =
                model.DurationMinutes ?? 0,

            ["Order"] =
                model.Order ?? 0,

            // =====================================================
            // FLAGS
            // =====================================================

            ["IsPublished"] =
                model.IsPublished,

            ["IsMandatory"] =
                model.IsMandatory,

            ["IsVisibleToStudents"] =
                model.IsVisibleToStudents,

            // =====================================================
            // AI FIELDS
            // =====================================================

            ["IsAiGenerated"] =
                model.IsAiGenerated,

            ["AiSummary"] =
                model.AiSummary,

            ["AiTags"] =
                model.AiTags,

            ["AiMetadata"] =
                model.AiMetadata,

            ["AiScore"] =
                model.AiScore
        };
    }

    // =====================================================
    // CONTENT TYPE SAFE MAPPING
    // =====================================================

    private static string GetContentTypeValue(ModelContentModel model)
    {
        return string.IsNullOrWhiteSpace(model.ContentTypeCode)
            ? "N/A"
            : model.ContentTypeCode;
    }


    // =====================================================
    // SAFE HELPERS
    // =====================================================

    private static string
        GetString(
            IDictionary<string, object> fields,
            string key)
    {
        return fields.ContainsKey(key)
            ? fields[key]?.ToString()
                ?? string.Empty
            : string.Empty;
    }

    private static bool
        GetBool(
            IDictionary<string, object> fields,
            string key)
    {
        return fields.ContainsKey(key)
               &&
               bool.TryParse(
                   fields[key]?.ToString(),
                   out var value)
               &&
               value;
    }

    private static decimal?
        GetDecimal(
            IDictionary<string, object> fields,
            string key)
    {
        return fields.ContainsKey(key)
               &&
               decimal.TryParse(
                   fields[key]?.ToString(),
                   out var value)
            ? value
            : null;
    }

    private static int?
        GetInt(
            IDictionary<string, object> fields,
            string key)
    {
        return fields.ContainsKey(key)
               &&
               int.TryParse(
                   fields[key]?.ToString(),
                   out var value)
            ? value
            : null;
    }
}