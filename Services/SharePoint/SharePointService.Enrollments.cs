using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions;
using Microsoft.Kiota.Abstractions.Serialization;
using MITANZ360Edu.Web.Models;

namespace MITANZ360Edu.Web.Services;

public partial class SharePointService
{
    private const string EnrollmentsListDisplayName = "Enrollments";
    private string? _enrollmentsListId;

    // ======================================================
    // INTERNAL : RESOLVE LIST ID (FILTERED, NO FULL LOAD)
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
                r.QueryParameters.Select = ["id", "displayName"];
                r.QueryParameters.Top = 5;
            }, ct);

        var list = lists?.Value?.FirstOrDefault();

        if (list == null || string.IsNullOrWhiteSpace(list.Id))
            throw new InvalidOperationException($"SharePoint list '{EnrollmentsListDisplayName}' not found.");

        _enrollmentsListId = list.Id!;
        return _enrollmentsListId!;
    }

    // ======================================================
    // ✅ ENROLLMENTS : COUNT BY COURSE (GRAPH‑SAFE, PAGINATED)
    // ======================================================
    public async Task<int> GetEnrollmentCountAsync(int courseId)
    {
        EnforceAdminOrTrainer();

        if (courseId <= 0)
            return 0;

        var listId = await GetEnrollmentsListIdAsync().ConfigureAwait(false);
        var total = 0;

        var response = await _graphClient
            .Sites[SiteId]
            .Lists[listId]
            .Items
            .GetAsync(r =>
            {
                r.QueryParameters.Filter = $"fields/CourseLookupId eq {courseId}";
                r.QueryParameters.Expand = new[] { "fields" };
                r.QueryParameters.Top = 200;
                r.Headers.Add("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly");
            });

        while (response != null)
        {
            if (response.Value != null)
                total += response.Value.Count;

            if (string.IsNullOrWhiteSpace(response.OdataNextLink))
                break;

            // ✅ Kiota correct SendAsync overload: (requestInfo, factory, errorMapping, cancellationToken)
            var req = new RequestInformation
            {
                HttpMethod = Method.GET,
                UrlTemplate = response.OdataNextLink
            };

            response = await _graphClient.RequestAdapter.SendAsync(
                req,
                ListItemCollectionResponse.CreateFromDiscriminatorValue,
                errorMapping: null,
                cancellationToken: CancellationToken.None
            ).ConfigureAwait(false);
        }

        return total;
    }

    // ======================================================
    // ENROLLMENTS : GET BY COURSE
    // ======================================================
    public async Task<List<EnrollmentModel>> GetEnrollmentsByCourseAsync(int courseId)
    {
        EnforceAdminOrTrainer();

        var listId = await GetEnrollmentsListIdAsync().ConfigureAwait(false);
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
                r.Headers.Add("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly");
            });

        if (response?.Value == null)
            return results;

        foreach (var item in response.Value)
        {
            if (item.Fields == null)
                continue;

            results.Add(MapEnrollment(item));
        }

        return results;
    }

    // ======================================================
    // ENROLLMENTS : CREATE (ADMIN → PENDING)
    // ======================================================
    public async Task<string> CreateEnrollmentAsync(EnrollmentModel model)
    {
        EnforceAdmin();

        model.Status = EnrollmentStatus.Pending;
        model.PaymentStatus = PaymentStatus.Pending;
        model.EnrollmentDate = DateTime.UtcNow;

        var enrollmentCode = await GenerateEnrollmentCodeAsync().ConfigureAwait(false);
        var listId = await GetEnrollmentsListIdAsync().ConfigureAwait(false);

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
                        ["Title"] = enrollmentCode,
                        ["EnrollmentCode"] = enrollmentCode,
                        ["CourseLookupId"] = model.CourseId,
                        ["StudentLookupId"] = model.StudentId,
                        ["Status"] = model.Status.ToString(),
                        ["PaymentStatus"] = model.PaymentStatus.ToString(),
                        ["EnrollmentDate"] = model.EnrollmentDate,
                        ["StartDate"] = model.StartDate,
                        ["EndDate"] = model.EndDate,
                        ["Notes"] = BuildNote("Admin", "Enrollment created and submitted for approval")
                    }
                }
            });

        return created!.Id!;
    }

    // ======================================================
    // ENROLLMENTS : UPDATE (ADMIN)
    // ======================================================
    public async Task UpdateEnrollmentAsync(EnrollmentModel model)
    {
        EnforceAdmin();

        if (string.IsNullOrWhiteSpace(model.Id))
            throw new InvalidOperationException("Invalid Enrollment ID.");

        var listId = await GetEnrollmentsListIdAsync().ConfigureAwait(false);

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
                    ["Notes"] = BuildNote("Admin", "Enrollment updated")
                }
            });
    }

    // ======================================================
    // ENROLLMENTS : DELETE (ADMIN)
    // ======================================================
    public async Task DeleteEnrollmentAsync(string enrollmentId)
    {
        EnforceAdmin();

        if (string.IsNullOrWhiteSpace(enrollmentId))
            throw new InvalidOperationException("Invalid Enrollment ID.");

        var listId = await GetEnrollmentsListIdAsync().ConfigureAwait(false);

        await _graphClient
            .Sites[SiteId]
            .Lists[listId]
            .Items[enrollmentId]
            .DeleteAsync();
    }

    // ======================================================
    // ENROLLMENTS : ACADEMIC APPROVAL
    // ======================================================
    public async Task ApproveEnrollmentAsync(string enrollmentId, string comment)
    {
        EnforceAcademicAuthority();

        var listId = await GetEnrollmentsListIdAsync().ConfigureAwait(false);

        await _graphClient
            .Sites[SiteId]
            .Lists[listId]
            .Items[enrollmentId]
            .Fields
            .PatchAsync(new FieldValueSet
            {
                AdditionalData = new Dictionary<string, object?>
                {
                    ["Status"] = EnrollmentStatus.Approved.ToString(),
                    ["ApprovedBy"] = CurrentUserUpn,
                    ["Notes"] = BuildNote("Academic", comment)
                }
            });
    }

    // ======================================================
    // ENROLLMENTS : FINANCE APPROVAL → ACTIVE
    // ======================================================
    public async Task ApproveEnrollmentPaymentAsync(string enrollmentId, string comment)
    {
        EnforceFinance();

        var listId = await GetEnrollmentsListIdAsync().ConfigureAwait(false);

        await _graphClient
            .Sites[SiteId]
            .Lists[listId]
            .Items[enrollmentId]
            .Fields
            .PatchAsync(new FieldValueSet
            {
                AdditionalData = new Dictionary<string, object?>
                {
                    ["PaymentStatus"] = PaymentStatus.Paid.ToString(),
                    ["Status"] = EnrollmentStatus.Active.ToString(),
                    ["PaymentApprovedBy"] = CurrentUserUpn,
                    ["Notes"] = BuildNote("Finance", comment)
                }
            });
    }

    // ======================================================
    // INTERNAL : ENROLLMENT MAPPER (kept safe)
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
            Status = ParseStatus(GetString(f, "Status")),
            PaymentStatus = ParsePaymentStatus(GetString(f, "PaymentStatus")),
            EnrollmentDate = GetDate(f, "EnrollmentDate") ?? DateTime.MinValue,
            StartDate = GetDate(f, "StartDate"),
            EndDate = GetDate(f, "EndDate"),
            Notes = GetStringNullable(f, "Notes"),
            ApprovedBy = GetUserDisplayName(f, "ApprovedBy"),
            PaymentApprovedBy = GetUserDisplayName(f, "PaymentApprovedBy"),
            Created = GetDate(f, "Created") ?? DateTime.MinValue,
            Modified = GetDate(f, "Modified") ?? DateTime.MinValue,
            CreatedBy = GetUserDisplayName(f, "Author"),
            ModifiedBy = GetUserDisplayName(f, "Editor")
        };
    }

    private static EnrollmentStatus ParseStatus(string s)
        => Enum.TryParse<EnrollmentStatus>(s, ignoreCase: true, out var v) ? v : EnrollmentStatus.Pending;

    private static PaymentStatus ParsePaymentStatus(string s)
        => Enum.TryParse<PaymentStatus>(s, ignoreCase: true, out var v) ? v : PaymentStatus.Pending;
}