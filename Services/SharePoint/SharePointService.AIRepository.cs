using Microsoft.Graph.Models;

namespace MITANZ360Edu.Web.Services;

public class AIRepositoryItem
{
    public int Id { get; set; }
    public string Title { get; set; } = string.Empty;
    public string EntityType { get; set; } = string.Empty;
    public decimal Score { get; set; }
    public string Status { get; set; } = "Complete";
    public string HtmlReport { get; set; } = string.Empty;
    public string Summary { get; set; } = string.Empty;
    public string Tags { get; set; } = string.Empty;
    public string Metadata { get; set; } = string.Empty;
    public DateTime Created { get; set; }
}

public partial class SharePointService
{
    private const string AIRepositoryListName = "AIRepository";

    // =====================================================
    // 📄 GET ITEMS
    // =====================================================

    public async Task<List<AIRepositoryItem>>
        GetAIRepositoryItemsAsync(
        string? status = null,
        string? entityType = null,
        decimal? minScore = null,
        string? search = null,
        string? orderBy = null,
        CancellationToken ct = default)
    {
        var listId =
            await GetListIdByTitleAsync(
                AIRepositoryListName,
                ct);

        var results = new List<AIRepositoryItem>();

        // ✅ BUILD FILTER
        var filter = BuildAIRepositoryFilter(
            status,
            entityType,
            minScore,
            search);

        // =============================================
        // ✅ FIRST PAGE
        // =============================================
        var response =
            await ExecuteWithRetryAsync(
                token => _graphClient
                    .Sites[SiteId]
                    .Lists[listId]
                    .Items
                    .GetAsync(
                        requestConfiguration: req =>
                        {
                            req.QueryParameters.Expand =
                                new[] { "fields" };

                            req.QueryParameters.Top = 100;

                            // ✅ FILTER
                            if (!string.IsNullOrWhiteSpace(filter))
                            {
                                req.QueryParameters.Filter = filter;
                            }

                            // ✅ SELECT (reduce payload)
                            req.QueryParameters.Select = new[]
                            {
                            "id",
                            "fields"
                            };

                            // ✅ SORT
                            if (!string.IsNullOrWhiteSpace(orderBy))
                            {
                                req.QueryParameters.Orderby = new[] { orderBy };
                            }
                        },
                        cancellationToken: token),
                "GetAIRepositoryItems",
                ct);

        // ✅ MAP FIRST PAGE
        if (response?.Value != null)
        {
            MapItems(response.Value, results);
        }

        var nextLink = response?.OdataNextLink;

        // =============================================
        // ✅ PAGINATION LOOP
        // =============================================
        while (!string.IsNullOrWhiteSpace(nextLink))
        {
            var nextPage =
                await ExecuteWithRetryAsync(
                    token => new Microsoft.Graph.Sites
                        .Item.Lists.Item.Items.ItemsRequestBuilder(
                            nextLink,
                            _graphClient.RequestAdapter)
                        .GetAsync(
                            cancellationToken: token),
                    "GetAIRepositoryItems_NextPage",
                    ct);

            if (nextPage?.Value != null)
            {
                MapItems(nextPage.Value, results);
            }

            nextLink = nextPage?.OdataNextLink;
        }

        return results;
    }

    // =====================================================
    // ➕ CREATE
    // =====================================================

    public async Task CreateAIRepositoryItemAsync(
        AIRepositoryItem model,
        CancellationToken ct = default)
    {
        EnforceAdminOrTrainer();

        var listId =
            await GetListIdByTitleAsync(
                AIRepositoryListName,
                ct);

        await ExecuteWithRetryAsync(
            token => _graphClient
                .Sites[SiteId]
                .Lists[listId]
                .Items
                .PostAsync(
                    new ListItem
                    {
                        Fields = new FieldValueSet
                        {
                            AdditionalData = new Dictionary<string, object>
                            {
                                ["Title"] = model.Title,
                                ["EntityType"] = model.EntityType,
                                ["Score"] = model.Score,
                                ["Status"] = model.Status,
                                ["Summary"] = model.Summary,
                                ["Tags"] = model.Tags,
                                ["Metadata"] = model.Metadata,
                                ["HtmlReport"] = model.HtmlReport
                            }
                        }
                    },
                    cancellationToken: token),
            "CreateAIRepositoryItem",   // ✅ REQUIRED
            ct);

        await WriteAuditAsync(
            "CREATE",
            "AIRepository",
            model.Title,
            ct);
    }

    // =====================================================
    // ✏ UPDATE
    // =====================================================

    public async Task UpdateAIRepositoryItemAsync(
        AIRepositoryItem model,
        CancellationToken ct = default)
    {
        EnforceAdminOrTrainer();

        var listId =
            await GetListIdByTitleAsync(
                AIRepositoryListName,
                ct);

        await ExecuteWithRetryAsync(
            token => _graphClient
                .Sites[SiteId]
                .Lists[listId]
                .Items[model.Id.ToString()]
                .Fields
                .PatchAsync(
                    new FieldValueSet
                    {
                        AdditionalData = new Dictionary<string, object>
                        {
                            ["Title"] = model.Title,
                            ["EntityType"] = model.EntityType,
                            ["Score"] = model.Score,
                            ["Status"] = model.Status,
                            ["Summary"] = model.Summary,
                            ["Tags"] = model.Tags,
                            ["Metadata"] = model.Metadata,
                            ["HtmlReport"] = model.HtmlReport
                        }
                    },
                    cancellationToken: token),
            "UpdateAIRepositoryItem",
            ct);
        await WriteAuditAsync(
            "UPDATE",
            "AIRepository",
            model.Id.ToString(),
            ct);
    }

    // =====================================================
    // ❌ DELETE
    // =====================================================

    public async Task DeleteAIRepositoryItemAsync(
        int id,
        CancellationToken ct = default)
    {
        EnforceAdmin();

        var listId =
            await GetListIdByTitleAsync(
                AIRepositoryListName,
                ct);

        await ExecuteWithRetryAsync(
            async token =>
                await _graphClient
                    .Sites[SiteId]
                    .Lists[listId]
                    .Items[id.ToString()]
                    .DeleteAsync( cancellationToken: token),
            "DeleteAIRepositoryItem",
            ct);

        await WriteAuditAsync(
            "DELETE",
            "AIRepository",
            id.ToString(),
            ct);
    }

    // =====================================================
    // 📄 LIST ITEM PAGINATION (GRAPH SAFE)
    // =====================================================

    protected async Task<List<ListItem>> GetPagedListItemsAsync(
        string listId,
        Action<
            Microsoft.Kiota.Abstractions.RequestConfiguration<
                Microsoft.Graph.Sites.Item.Lists.Item.Items.ItemsRequestBuilder
                    .ItemsRequestBuilderGetQueryParameters>>? config,
        string operation,
        CancellationToken ct)
    {
        var allItems = new List<ListItem>();

        var requestBuilder = _graphClient
            .Sites[SiteId]
            .Lists[listId]
            .Items;

        // ✅ FIRST PAGE
        var response = await ExecuteWithRetryAsync(
            token => requestBuilder.GetAsync(
                requestConfiguration: config,
                cancellationToken: token),
            operation,
            ct);

        if (response?.Value != null)
            allItems.AddRange(response.Value);

        var nextLink = response?.OdataNextLink;

        // ✅ LOOP NEXT PAGES
        while (!string.IsNullOrWhiteSpace(nextLink))
        {
            var nextPage = await ExecuteWithRetryAsync(
                token => new Microsoft.Graph.Sites.Item.Lists.Item.Items
                    .ItemsRequestBuilder(nextLink, _graphClient.RequestAdapter)
                    .GetAsync(cancellationToken: token),
                operation + "_NextPage",
                ct);

            if (nextPage?.Value != null)
                allItems.AddRange(nextPage.Value);

            nextLink = nextPage?.OdataNextLink;
        }

        return allItems;
    }
}
