using Microsoft.Graph;
using Microsoft.Graph.Models;
using MITANZ360Edu.Web.Services.AI;
using System.Text.Json;
using System.Threading;

namespace MITANZ360Edu.Web.Services
{
    public class AIRepositoryItem
    {
        public int Id { get; set; }
        public string Title { get; set; } = "";
        public string Metadata { get; set; } = "";
        public string Summary { get; set; } = "";
        public string HtmlReport { get; set; } = "";
        public string Status { get; set; } = "";
        public string Tags { get; set; } = "";
        public decimal Score { get; set; }
        public string EntityType { get; set; } = "";
    }

    public partial class SharePointService
    {
        // =====================================================
        // ✅ SIMPLE GET (FOR UI)
        // =====================================================
        public async Task<List<AIRepositoryItem>> GetAIRepositoryItemsAsync(string? orderBy)
        {
            return await GetAIRepositoryItemsAsync(null, orderBy, null, null, CancellationToken.None);
        }

        // =====================================================
        // ✅ OVERLOAD FOR TEST
        // =====================================================
        public async Task<List<AIRepositoryItem>> GetAIRepositoryItemsAsync(
            string? filter,
            string? orderBy,
            decimal? top,
            string? select,
            string? expand)
        {
            return await GetAIRepositoryItemsAsync(
                filter,
                orderBy,
                top,
                select,
                CancellationToken.None
            );
        }

        // =====================================================
        // ✅ MAIN GET
        // =====================================================
        public async Task<List<AIRepositoryItem>> GetAIRepositoryItemsAsync(
            string? filter,
            string? orderBy,
            decimal? top,
            string? select,
            CancellationToken cancellationToken)
        {
            var listId = _configuration["SharePoint:Lists:AIRepository"];

            if (string.IsNullOrWhiteSpace(listId))
                throw new InvalidOperationException("AIRepository ListId missing.");

            var topInt = top.HasValue ? (int)top.Value : 50;

            var response = await _graphClient
                .Sites[SiteId]
                .Lists[listId]
                .Items
                .GetAsync(
                    requestConfiguration =>
                    {
                        requestConfiguration.QueryParameters.Top = topInt;
                    },
                    cancellationToken
                );

            var items = new List<AIRepositoryItem>();

            foreach (var item in response?.Value ?? Enumerable.Empty<ListItem>())
            {
                var fields = item.Fields?.AdditionalData;

                items.Add(new AIRepositoryItem
                {
                    Id = int.Parse(item.Id),
                    Title = GetField(fields, "Title"),
                    Metadata = GetField(fields, "Metadata"),
                    Summary = GetField(fields, "Summary"),
                    HtmlReport = GetField(fields, "HtmlReport"),
                    Status = GetField(fields, "Status"),
                    Tags = GetField(fields, "Tags"),
                    Score = decimal.TryParse(GetField(fields, "Score"), out var s) ? s : 0,
                    EntityType = GetField(fields, "EntityType")
                });
            }

            return items;
        }

        // =====================================================
        // ✅ CREATE
        // =====================================================
        public async Task CreateAIRepositoryItemAsync(AIRepositoryItem item)
        {
            var listId = _configuration["SharePoint:Lists:AIRepository"];

            var fields = new Dictionary<string, object>
            {
                { "Title", item.Title },
                { "Metadata", item.Metadata },
                { "Summary", item.Summary },
                { "HtmlReport", item.HtmlReport },
                { "Status", item.Status },
                { "Tags", item.Tags },
                { "Score", item.Score },
                { "EntityType", item.EntityType }
            };

            await _graphClient
                .Sites[SiteId]
                .Lists[listId]
                .Items
                .PostAsync(new ListItem
                {
                    Fields = new FieldValueSet { AdditionalData = fields }
                });
        }

        // =====================================================
        // ✅ UPDATE (MAIN)
        // =====================================================
        public async Task UpdateAIRepositoryItemAsync(AIRepositoryItem item)
        {
            await UpdateAIRepositoryItemAsync(item, CancellationToken.None);
        }

        // ✅ OVERLOAD (automation uses this)
        public async Task UpdateAIRepositoryItemAsync(
            AIRepositoryItem item,
            CancellationToken cancellationToken)
        {
            var listId = _configuration["SharePoint:Lists:AIRepository"];

            var fields = new Dictionary<string, object>
            {
                { "Title", item.Title },
                { "Metadata", item.Metadata },
                { "Summary", item.Summary },
                { "HtmlReport", item.HtmlReport },
                { "Status", item.Status },
                { "Tags", item.Tags },
                { "Score", item.Score },
                { "EntityType", item.EntityType }
            };

            await _graphClient
                .Sites[SiteId]
                .Lists[listId]
                .Items[item.Id.ToString()]
                .Fields
                .PatchAsync(
                new FieldValueSet { AdditionalData = fields },
                null,                      // ✅ REQUIRED placeholder
                cancellationToken          // ✅ NOW correct position
            );
        }

        // =====================================================
        // ✅ DELETE
        // =====================================================
        public async Task DeleteAIRepositoryItemAsync(int itemId)
        {
            var listId = _configuration["SharePoint:Lists:AIRepository"];

            await _graphClient
                .Sites[SiteId]
                .Lists[listId]
                .Items[itemId.ToString()]
                .DeleteAsync();
        }

        // =====================================================
        // ✅ AI SAVE
        // =====================================================
        public async Task UpdateAIFieldsAndAttachFileAsync(
            int itemId,
            AIService.AIResult aiResult)
        {
            var listId = _configuration["SharePoint:Lists:AIRepository"];

            var fields = new Dictionary<string, object>
            {
                { "Summary", aiResult.SummaryText ?? "" },
                { "Metadata", aiResult.Json ?? "" },
                { "HtmlReport", aiResult.HtmlContent ?? "" }
            };

            await _graphClient
                .Sites[SiteId]
                .Lists[listId]
                .Items[itemId.ToString()]
                .Fields
                .PatchAsync(new FieldValueSet
                {
                    AdditionalData = fields
                });
        }
    }
}