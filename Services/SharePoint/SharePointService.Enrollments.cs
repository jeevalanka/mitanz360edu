using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions;
using MITANZ360Edu.Web.Models;

namespace MITANZ360Edu.Web.Services;

public partial class SharePointService
{
    private const string EnrollmentsListDisplayName = "Enrollments";
    private string? _enrollmentsListId;

    // ======================================================
    // INTERNAL: RESOLVE LIST ID
    // ======================================================
    private async Task<string> GetEnrollmentsListIdAsync(CancellationToken ct = default)
    {
        if (!string.IsNullOrWhiteSpace(_enrollmentsListId))
            return _enrollmentsListId!;

        var safe = EscapeODataString(EnrollmentsListDisplayName);

        var lists = await _graphClient
            .Sites[SiteId]
            .Lists
            .GetAsync(r =>
            {
                r.QueryParameters.Filter = $"displayName eq '{safe}'";
                r.QueryParameters.Select = new[] { "id", "displayName" };
                r.QueryParameters.Top = 5;
            }, ct);

        var list = lists?.Value?.FirstOrDefault();

        if (list == null || string.IsNullOrWhiteSpace(list.Id))
            throw new InvalidOperationException($"SharePoint list '{EnrollmentsListDisplayName}' not found.");

        _enrollmentsListId = list.Id!;
        return _enrollmentsListId!;
    }

    // ======================================================
    // COUNT BY COURSE (PAGINATED)
    // ======================================================
    public async Task<int> GetEnrollmentCountAsync(int courseId)
    {
        EnforceAdminOrTrainer();

        if (courseId <= 0)
            return 0;

        var listId = await GetEnrollmentsListIdAsync();
        var total = 0;

        var response = await _graphClient
            .Sites[SiteId]
            .Lists[listId]
            .Items
            .GetAsync(r =>
            {
                r.QueryParameters.Filter = $"fields/CourseLookupId eq {courseId}";
                r.QueryParameters.Top = 200;
            });

        while (response != null)
        {
            total += response.Value?.Count ?? 0;

            if (string.IsNullOrWhiteSpace(response.OdataNextLink))
                break;

            var req = new RequestInformation
            {
                HttpMethod = Method.GET,
                UrlTemplate = response.OdataNextLink
            };

            response = await _graphClient.RequestAdapter.SendAsync(
                req,
                ListItemCollectionResponse.CreateFromDiscriminatorValue,
                errorMapping: null,
                cancellationToken: CancellationToken.None);
        }

        return total;
    }

    // ======================================================
    // GET BY COURSE
    // ======================================================
    public async Task<List<EnrollmentModel>> GetEnrollmentsByCourseAsync(int courseId)
    {
        EnforceAdminOrTrainer();

        var listId = await GetEnrollmentsListIdAsync();
        var results = new List<EnrollmentModel>();

        var response = await _graphClient
            .Sites[SiteId]
            .Lists[listId]
            .Items
            .GetAsync(r =>
            {
                r.QueryParameters.Filter = $"fields/CourseLookupId eq {courseId}";
                r.QueryParameters.Expand = new[] { "fields" };
                r.QueryParameters.Top = 200;
            });

        foreach (var item in response?.Value ?? Enumerable.Empty<ListItem>())
        {
            if (item.Fields != null)
                results.Add(MapEnrollment(item));
        }

        return results;
    }

    // ======================================================
    // CREATE
    // ======================================================
    public async Task<string> CreateEnrollmentAsync(EnrollmentModel model)
    {
        EnforceAdmin();

        model.Status = "PENDING";
        model.PaymentStatus = "PENDING";
        model.EnrollmentDate = DateTime.UtcNow;

        var code = await GenerateEnrollmentCodeAsync();
        var listId = await GetEnrollmentsListIdAsync();

        var created = await _graphClient
            .Sites[SiteId]
            .Lists[listId]
            .Items
            .PostAsync(new ListItem
            {
                Fields = new FieldValueSet
                {
                    AdditionalData = new Dictionary<string, object?>
                    {
                        ["Title"] = code,
                        ["EnrollmentCode"] = code,
                        ["CourseLookupId"] = model.CourseId,
                        ["StudentLookupId"] = model.StudentId,
                        ["Status"] = model.Status,
                        ["PaymentStatus"] = model.PaymentStatus,
                        ["EnrollmentDate"] = model.EnrollmentDate,
                        ["StartDate"] = model.StartDate,
                        ["EndDate"] = model.EndDate,
                        ["Notes"] = BuildNote("Admin", "Enrollment created")
                    }
                }
            });

        return created!.Id!;
    }

    // ======================================================
    // UPDATE
    // ======================================================
    public async Task UpdateEnrollmentAsync(EnrollmentModel model)
    {
        EnforceAdmin();

        if (string.IsNullOrWhiteSpace(model.Id))
            throw new InvalidOperationException("Invalid Enrollment ID.");

        var listId = await GetEnrollmentsListIdAsync();

        await _graphClient
            .Sites[SiteId]
            .Lists[listId]
            .Items[model.Id]
            .Fields
            .PatchAsync(new FieldValueSet
            {
                AdditionalData = new Dictionary<string, object?>
                {
                    ["StartDate"] = model.StartDate,
                    ["EndDate"] = model.EndDate,
                    ["Status"] = model.Status,
                    ["PaymentStatus"] = model.PaymentStatus,
                    ["Notes"] = BuildNote("Admin", "Enrollment updated")
                }
            });
    }

    // ======================================================
    // DELETE
    // ======================================================
    public async Task DeleteEnrollmentAsync(string id)
    {
        EnforceAdmin();

        if (string.IsNullOrWhiteSpace(id))
            throw new InvalidOperationException("Invalid Enrollment ID.");

        var listId = await GetEnrollmentsListIdAsync();

        await _graphClient
            .Sites[SiteId]
            .Lists[listId]
            .Items[id]
            .DeleteAsync();
    }

    // ======================================================
    // ACADEMIC APPROVAL
    // ======================================================
    public async Task ApproveEnrollmentAsync(string id, string comment)
    {
        EnforceAcademicAuthority();

        var listId = await GetEnrollmentsListIdAsync();

        await _graphClient
            .Sites[SiteId]
            .Lists[listId]
            .Items[id]
            .Fields
            .PatchAsync(new FieldValueSet
            {
                AdditionalData = new Dictionary<string, object?>
                {
                    ["Status"] = "APPROVED",
                    ["ApprovedBy"] = CurrentUserUpn,
                    ["Notes"] = BuildNote("Academic", comment)
                }
            });
    }

    // ======================================================
    // FINANCE APPROVAL
    // ======================================================
    public async Task ApproveEnrollmentPaymentAsync(string id, string comment)
    {
        EnforceFinance();

        var listId = await GetEnrollmentsListIdAsync();

        await _graphClient
            .Sites[SiteId]
            .Lists[listId]
            .Items[id]
            .Fields
            .PatchAsync(new FieldValueSet
            {
                AdditionalData = new Dictionary<string, object?>
                {
                    ["PaymentStatus"] = "PAID",
                    ["Status"] = "ACTIVE",
                    ["PaymentApprovedBy"] = CurrentUserUpn,
                    ["Notes"] = BuildNote("Finance", comment)
                }
            });
    }

    // ======================================================
    // MAPPER
    // ======================================================
    private EnrollmentModel MapEnrollment(ListItem item)
    {
        var f = item.Fields!;

        return new EnrollmentModel
        {
            Id = item.Id,
            EnrollmentCode = GetString(f, "EnrollmentCode"),
            Title = GetString(f, "Title"),
            CourseId = GetInt(f, "CourseLookupId") ?? 0,
            StudentId = GetInt(f, "StudentLookupId") ?? 0,
            Status = Safe(GetString(f, "Status"), "PENDING"),
            PaymentStatus = Safe(GetString(f, "PaymentStatus"), "PENDING"),
            EnrollmentDate = GetDate(f, "EnrollmentDate") ?? DateTime.MinValue,
            StartDate = GetDate(f, "StartDate"),
            EndDate = GetDate(f, "EndDate"),
            Notes = GetStringNullable(f, "Notes"),
            ApprovedBy = GetUserDisplayName(f, "ApprovedBy"),
            PaymentApprovedBy = GetUserDisplayName(f, "PaymentApprovedBy")
        };
    }

    // ======================================================
    // HELPER
    // ======================================================
    private static string Safe(string? value, string fallback)
        => string.IsNullOrWhiteSpace(value) ? fallback : value;
}