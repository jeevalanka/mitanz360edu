using Microsoft.Graph.Models;
using System.Collections.Generic;
using System.Linq;


namespace MITANZ360Edu.Web.Services;

public class CountryLookup
{
    public string Id { get; set; } = string.Empty;

    public string Title { get; set; } = string.Empty;

    public string ISOCode { get; set; } = string.Empty;

    public string Currency { get; set; } = string.Empty;

    public string TimeZone { get; set; } = string.Empty;

    public bool IsActive { get; set; }
}

public partial class SharePointService
{
    public async Task<List<CountryLookup>>
    GetCountriesAsync()
    {
        var results = new List<CountryLookup>();

        var response =
            await _graphClient
                .Sites[SiteId]
                .Lists["Countries"]
                .Items
                .GetAsync(req =>
                {
                    req.QueryParameters.Expand =
                    [
                        "fields"
                    ];
                });

        if (response?.Value == null)
            return results;

        foreach (var item in response.Value)
        {
            var fields = item.Fields?.AdditionalData;

            results.Add(new CountryLookup
            {
                Id = item.Id ?? string.Empty,
                Title = GetField(fields, "Title"),
                ISOCode = GetField(fields, "ISOCode"),
                Currency = GetField(fields, "Currency"),
                TimeZone = GetField(fields, "TimeZone"),
                IsActive = GetBoolField(fields, "IsActive")
            });
        }

        return results
            .Where(x => x.IsActive)
            .OrderBy(x => x.Title)
            .ToList();
    }
}