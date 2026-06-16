using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;
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
    // GET ALL COURSES
    // ======================================================
    public async Task<IReadOnlyList<CourseModel>> GetCoursesAsync()
    {
        if (!IsAuthenticated())
            return Array.Empty<CourseModel>();

        var results = new List<CourseModel>();
        var listId = await GetCoursesListIdAsync();

        var response = await _graphClient
            .Sites[SiteId]
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
                .Sites[SiteId]
                .Lists[listId]
                .Items
                .WithUrl(response.OdataNextLink)
                .GetAsync();
        }

        return results;
    }

    // ======================================================
    // GET COURSE BY ID
    // ======================================================
    public async Task<CourseModel?> GetCourseByIdAsync(int courseId)
    {
        if (!IsAuthenticated() || courseId <= 0)
            return null;

        var listId = await GetCoursesListIdAsync();

        var item = await _graphClient
            .Sites[SiteId]
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
    // CREATE COURSE
    // ======================================================
    public async Task<int> CreateCourseAsync(CourseModel model)
    {
        EnforceAdmin();

        if (string.IsNullOrWhiteSpace(model.CourseCode))
            throw new Exception("CourseCode required");

        var listId = await GetCoursesListIdAsync();

        // ✅ CREATE FOLDER STRUCTURE
        var folderId = await CreateCourseFolderStructureAsync(model.CourseCode);

        var fields = BuildCourseFields(model, folderId);

        var created = await _graphClient
            .Sites[SiteId]
            .Lists[listId]
            .Items
            .PostAsync(new ListItem
            {
                Fields = fields
            });

        return int.Parse(created!.Id!);
    }

    // ======================================================
    // UPDATE COURSE
    // ======================================================
    public async Task UpdateCourseAsync(CourseModel model)
    {
        EnforceAdmin();

        if (model.Id <= 0)
            throw new Exception("Invalid ID");

        var listId = await GetCoursesListIdAsync();

        var fields = BuildCourseFields(model);

        await _graphClient
            .Sites[SiteId]
            .Lists[listId]
            .Items[model.Id.ToString()]
            .Fields
            .PatchAsync(fields);
    }

    // ======================================================
    // DELETE COURSE
    // ======================================================
    public async Task DeleteCourseAsync(int courseId)
    {
        EnforceAdmin();

        var listId = await GetCoursesListIdAsync();

        await _graphClient
            .Sites[SiteId]
            .Lists[listId]
            .Items[courseId.ToString()]
            .DeleteAsync();
    }

    // ======================================================
    // BUILD FIELDS
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
            data["FolderId"] = folderId; // ✅ IMPORTANT
        }

        return new FieldValueSet
        {
            AdditionalData = data
        };
    }

    // ======================================================
    // ✅ ✅ FIXED MAPPER
    // ======================================================
    private CourseModel MapCourse(ListItem item)
    {
        var f = item.Fields!;

        return new CourseModel
        {
            Id = int.TryParse(item.Id, out var id) ? id : 0,

            Title = GetString(f, "Title"),
            CourseCode = GetString(f, "CourseCode"),

            // ✅ ✅ THIS LINE FIXES YOUR PROBLEM
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

    // ======================================================
    // HELPERS
    // ======================================================
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

    public async Task<CourseModel?> GetCourseByFolderIdAsync(string folderId)
    {
        if (string.IsNullOrWhiteSpace(folderId))
            return null;

        var listId = await GetCoursesListIdAsync();

        var request = _graphClient
            .Sites[SiteId]
            .Lists[listId]
            .Items;

        var response = await request.GetAsync(r =>
        {
            r.QueryParameters.Expand = new[] { "fields" };
            r.QueryParameters.Select = new[] { "id", "fields" };
            r.QueryParameters.Top = 100;
        });

        while (response?.Value != null)
        {
            foreach (var item in response.Value)
            {
                if (item.Fields?.AdditionalData == null)
                    continue;

                var fields = item.Fields.AdditionalData;

                // ✅ STEP 1: Get FolderId
                if (!fields.TryGetValue("FolderId", out var val))
                    continue;

                var spFolderId = val?.ToString()?.Trim().Trim('"');

                if (!string.Equals(spFolderId, folderId.Trim(), StringComparison.OrdinalIgnoreCase))
                    continue;

                // ✅ ✅ STEP 2: VERY IMPORTANT FILTER (REAL COURSE ONLY)
                // Ignore invalid/duplicate records

                if (!fields.TryGetValue("CourseCode", out var courseCodeVal))
                    continue;

                var courseCode = courseCodeVal?.ToString();

                if (string.IsNullOrWhiteSpace(courseCode))
                    continue;

                // ✅ OPTIONAL: FILTER OUT "Test-Folder" or known wrong patterns
                if (fields.TryGetValue("Title", out var titleVal))
                {
                    var title = titleVal?.ToString();

                    if (!string.IsNullOrWhiteSpace(title) &&
                        title.StartsWith("Test-Folder", StringComparison.OrdinalIgnoreCase))
                    {
                        continue; // ❌ skip wrong one
                    }
                }

                // ✅ ✅ FINAL MATCH
                return MapCourse(item);
            }

            if (string.IsNullOrWhiteSpace(response.OdataNextLink))
                break;

            response = await request
                .WithUrl(response.OdataNextLink)
                .GetAsync();
        }

        return null;
    }
}