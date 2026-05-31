using Microsoft.Graph.Models;
using System.Collections.Concurrent;

namespace MITANZ360Edu.Web.Services;

public partial class SharePointService
{
    // ======================================================
    // IN-MEMORY CACHE (THREAD SAFE)
    // ======================================================
    private static readonly ConcurrentDictionary<string, CacheEntry> _choiceCache = new();
    private static readonly TimeSpan CacheTtl = TimeSpan.FromMinutes(30);

    private sealed class CacheEntry
    {
        public IReadOnlyList<string> Values { get; init; } = Array.Empty<string>();
        public DateTime CachedAtUtc { get; init; }
    }

    // ======================================================
    // CHOICES : GET VALUES FOR LIST COLUMN (CACHED)
    // ======================================================
    public async Task<IReadOnlyList<string>> GetChoiceValuesAsync(string listName, string columnName)
    {
        if (string.IsNullOrWhiteSpace(listName))
            throw new ArgumentException("List name is required", nameof(listName));

        if (string.IsNullOrWhiteSpace(columnName))
            throw new ArgumentException("Column name is required", nameof(columnName));

        var cacheKey = $"{listName}\n{columnName}".ToLowerInvariant();

        if (_choiceCache.TryGetValue(cacheKey, out var cached)
            && DateTime.UtcNow - cached.CachedAtUtc < CacheTtl)
        {
            return cached.Values;
        }

        try
        {
            var columns = await _graphClient
                .Sites[SiteId]
                .Lists[listName]
                .Columns
                .GetAsync();

            var values = columns?.Value?
                .FirstOrDefault(c => c.Name.Equals(columnName, StringComparison.OrdinalIgnoreCase))
                ?.Choice?.Choices?
                .ToList()
                ?? new List<string>();

            _choiceCache[cacheKey] = new CacheEntry
            {
                Values = values,
                CachedAtUtc = DateTime.UtcNow
            };

            return values;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to load choices for list {List} column {Column}", listName, columnName);

            if (cached != null)
                return cached.Values;

            return Array.Empty<string>();
        }
    }

    // ======================================================
    // ADMIN : CLEAR CACHE
    // ======================================================
    public void ClearChoiceCache()
    {
        _choiceCache.Clear();
    }
}