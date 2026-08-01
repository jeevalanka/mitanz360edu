using Microsoft.Graph.Models;
using MITANZ360Edu.Web.Models;

namespace MITANZ360Edu.Web.Services;

public partial class SharePointService
{
    private const string FieldDurationHours = "DurationHrs";
    private const string FieldCreditValue = "CreditValue";

    private string? _coursesListId;

    // ======================================================
    // LIST ID
    // ======================================================
    private async Task<string> GetCoursesListIdAsync(CancellationToken ct = default)
    {
        if (!string.IsNullOrWhiteSpace(_coursesListId))
            return _coursesListId!;

        if (string.IsNullOrWhiteSpace(CoursesListName))
            throw new InvalidOperationException("Courses List ID missing.");

        _coursesListId = CoursesListName;
        return _coursesListId!;
    }

    // ======================================================
    // ✅ SEARCH COURSES (STANDARDIZED)
    // ======================================================
    public async Task<List<CourseModel>> SearchCoursesAsync(
        string search,
        int top = 50,
        CancellationToken ct = default)
    {
        if (!IsAuthenticated())
            return new List<CourseModel>();

        var listId = await GetCoursesListIdAsync(ct);
        var safe = (search ?? "").Replace("'", "''");

        var result = await ExecuteWithRetryAsync(
            async token =>
                await _graphClient
                    .Sites[SiteId]                      // ✅ FIXED
                    .Lists[listId]
                    .Items
                    .GetAsync(cfg =>
                    {
                        cfg.QueryParameters.Top = top;

                        cfg.QueryParameters.Filter =
                            $"contains(fields/Title,'{safe}') or contains(fields/CourseCode,'{safe}')";

                        cfg.QueryParameters.Select = new[] { "id", "fields" };
                        cfg.QueryParameters.Expand = new[] { "fields" };
                    }, token),
            "SearchCourses",
            ct);

        return result?.Value?.Select(MapCourse).ToList()
               ?? new List<CourseModel>();
    }

    // ======================================================
    // ✅ GET ALL COURSES (PAGINATED)
    // ======================================================
    public async Task<IReadOnlyList<CourseModel>> GetCoursesAsync()
    {
        if (!IsAuthenticated())
            return Array.Empty<CourseModel>();

        var results = new List<CourseModel>();
        var listId = await GetCoursesListIdAsync();

        var response = await _graphClient
            .Sites[SiteId]                              // ✅ FIXED
            .Lists[listId]
            .Items
            .GetAsync(r =>
            {
                r.QueryParameters.Expand = new[] { "fields" };
                r.QueryParameters.Select = new[] { "id", "fields" };
                r.QueryParameters.Top = 100;
            });

        while (response?.Value != null)
        {
            foreach (var item in response.Value)
            {
                if (item.Fields == null)
                    continue;

                results.Add(MapCourse(item));
            }

            if (string.IsNullOrWhiteSpace(response.OdataNextLink))
                break;

            response = await _graphClient
                .Sites[SiteId]                          // ✅ FIXED
                .Lists[listId]
                .Items
                .WithUrl(response.OdataNextLink)
                .GetAsync();
        }

        return results;
    }

    // ======================================================
    // ✅ GET COURSE BY ID
    // ======================================================
    public async Task<CourseModel?> GetCourseByIdAsync(int courseId)
    {
        if (!IsAuthenticated() || courseId <= 0)
            return null;

        var listId = await GetCoursesListIdAsync();

        var item = await _graphClient
            .Sites[SiteId]                              // ✅ FIXED
            .Lists[listId]
            .Items[courseId.ToString()]
            .GetAsync(r =>
            {
                r.QueryParameters.Expand = new[] { "fields" };
                r.QueryParameters.Select = new[] { "id", "fields" };
            });

        return item?.Fields != null ? MapCourse(item) : null;
    }

    // ======================================================
    // ✅ CREATE COURSE
    // ======================================================
    public async Task<int> CreateCourseAsync(CourseModel model)
    {
        EnforceAdmin();

        if (string.IsNullOrWhiteSpace(model.CourseCode))
            throw new Exception("CourseCode required");

        var listId = await GetCoursesListIdAsync();

        var folderId = await CreateCourseFolderStructureAsync(model.CourseCode);
        var fields = BuildCourseFields(model, folderId);

        var created = await _graphClient
            .Sites[SiteId]                              // ✅ FIXED
            .Lists[listId]
            .Items
            .PostAsync(new ListItem
            {
                Fields = fields
            });

        return int.Parse(created!.Id!);
    }

    // ======================================================
    // ✅ UPDATE COURSE
    // ======================================================
    public async Task UpdateCourseAsync(CourseModel model)
    {
        EnforceAdmin();

        if (model.Id <= 0)
            throw new Exception("Invalid ID");

        var listId = await GetCoursesListIdAsync();

        await _graphClient
            .Sites[SiteId]                              // ✅ FIXED
            .Lists[listId]
            .Items[model.Id.ToString()]
            .Fields
            .PatchAsync(BuildCourseFields(model));
    }

    // ======================================================
    // ✅ DELETE COURSE
    // ======================================================
    public async Task DeleteCourseAsync(int courseId)
    {
        EnforceAdmin();

        var listId = await GetCoursesListIdAsync();

        await _graphClient
            .Sites[SiteId]                              // ✅ FIXED
            .Lists[listId]
            .Items[courseId.ToString()]
            .DeleteAsync();
    }

    // ======================================================
    // ✅ BUILD FIELDS
    // ======================================================
    private static FieldValueSet BuildCourseFields(
        CourseModel model,
        string? folderId = null)
    {
        var data = new Dictionary<string, object?>
        {
            ["Title"] = model.Title,
            ["CourseCode"] = model.CourseCode,
            ["CourseCategory"] = model.CourseCategory,
            ["CourseContentType"] = model.CourseContentType,
            ["CourseType"] = model.CourseType,
            ["Level"] = model.Level,
            ["Language"] = model.Language,
            ["CourseStatus"] = model.CourseStatus,
            ["ApprovalStatus"] = model.ApprovalStatus,
            ["Archived"] = model.Archived,
            ["CourseVersion"] = model.CourseVersion,
            ["EffectiveFrom"] = model.EffectiveFrom,
            ["DeliveryMode"] = model.DeliveryMode,
            ["IsSelfPaced"] = model.IsSelfPaced,
            ["EnrollmentOpen"] = model.EnrollmentOpen,
            ["EnrollmentType"] = model.EnrollmentType,
            [FieldDurationHours] = model.DurationMinutes,
            [FieldCreditValue] = model.CreditValue,
            ["Description"] = model.Description,
            ["LearningOutcomes"] = model.LearningOutcomes,
            ["CertificateIssued"] = model.CertificateIssued,
            ["ImageUrl"] = model.ImageUrl
        };

        if (!string.IsNullOrWhiteSpace(folderId))
        {
            data["FolderId"] = folderId;
        }

        return new FieldValueSet
        {
            AdditionalData = data
        };
    }

    // ======================================================
    // ✅ MAPPER
    // ======================================================
    private CourseModel MapCourse(ListItem item)
    {
        var f = item.Fields!;

        return new CourseModel
        {
            Id = int.TryParse(item.Id, out var id) ? id : 0,
            Title = GetString(f, "Title"),
            CourseCode = GetString(f, "CourseCode"),
            FolderId = GetStringNullable(f, "FolderId"),
            CourseCategory = GetString(f, "CourseCategory"),
            CourseContentType = GetString(f, "CourseContentType"),
            CourseType = GetString(f, "CourseType"),
            Level = GetString(f, "Level"),
            Language = GetString(f, "Language"),
            CourseStatus = GetString(f, "CourseStatus"),
            ApprovalStatus = GetString(f, "ApprovalStatus"),
            Archived = GetBool(f, "Archived"),
            CourseVersion = GetString(f, "CourseVersion"),
            EffectiveFrom = GetDateTime(f, "EffectiveFrom"),
            DeliveryMode = GetString(f, "DeliveryMode"),
            IsSelfPaced = GetBool(f, "IsSelfPaced"),
            EnrollmentOpen = GetBool(f, "EnrollmentOpen"),
            EnrollmentType = GetString(f, "EnrollmentType"),
            DurationMinutes = GetDecimal(f, FieldDurationHours),
            CreditValue = GetDecimal(f, FieldCreditValue),
            Description = GetStringNullable(f, "Description"),
            LearningOutcomes = GetStringNullable(f, "LearningOutcomes"),
            CertificateIssued = GetBool(f, "CertificateIssued"),
            ImageUrl = GetStringNullable(f, "ImageUrl"),
            AiFeed = GetStringNullable(f, "AiSummary"),
        };
    }

    private static decimal? GetDecimal(FieldValueSet fields, string key)
    {
        if (fields.AdditionalData == null)
            return null;

        if (!fields.AdditionalData.TryGetValue(key, out var value))
            return null;

        return decimal.TryParse(value?.ToString(), out var result)
            ? result
            : null;
    }
}