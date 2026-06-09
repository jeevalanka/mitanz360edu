using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;
using MITANZ360Edu.Web.Models;

namespace MITANZ360Edu.Web.Services;

public partial class SharePointService
{
    // ======================================================
    // SHAREPOINT INTERNAL FIELD NAMES
    // ======================================================

    private const string FieldDurationHours = "DurationHrs";
    private const string FieldCreditValue = "CreditValue";

    private string? _coursesListId;

    // ======================================================
    // INTERNAL : RESOLVE COURSES LIST ID
    // ======================================================

    private async Task<string> GetCoursesListIdAsync(
        CancellationToken ct = default)
    {
        if (!string.IsNullOrWhiteSpace(_coursesListId))
        {
            return _coursesListId!;
        }

        if (string.IsNullOrWhiteSpace(CoursesListName))
        {
            throw new InvalidOperationException(
                "SharePoint Courses List ID configuration is missing.");
        }

        // IMPORTANT:
        // Config already contains LIST ID.
        _coursesListId = CoursesListName;

        _logger.LogDebug(
            "Using SharePoint Courses List ID: {ListId}",
            _coursesListId);

        return _coursesListId!;
    }

    // ======================================================
    // COURSES : GET ALL
    // ======================================================

    public async Task<IReadOnlyList<CourseModel>> GetCoursesAsync()
    {
        if (!IsAuthenticated())
        {
            _logger.LogWarning("Unauthorized access attempt to GetCoursesAsync.");

            return Array.Empty<CourseModel>();
        }

        var results = new List<CourseModel>();

        try
        {
            var listId = await GetCoursesListIdAsync();

            _logger.LogDebug(
                "Loading courses from SharePoint List ID: {ListId}",
                listId);

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
                    {
                        continue;
                    }

                    var course = MapCourse(item);

                    results.Add(course);

                    _logger.LogDebug("COURSE LOADED: {Title} | Duration: {Duration} | Credits: {Credits}",
                        course.Title,
                        course.DurationMinutes,
                        course.CreditValue);
                }

                if (string.IsNullOrWhiteSpace(response.OdataNextLink))
                {
                    break;
                }

                response = await _graphClient
                    .Sites[SiteId]
                    .Lists[listId]
                    .Items
                    .WithUrl(response.OdataNextLink)
                    .GetAsync();
            }

            _logger.LogDebug(
                "Total courses loaded: {Count}",
                results.Count);
        }
        catch (ODataError ex)
        {
            _logger.LogError(
                "Microsoft Graph OData error while loading courses: {Message}",
                ex.Error?.Message);

            throw;
        }
        catch (Exception ex)
        {
            _logger.LogError(
                ex,
                "Unexpected error while loading courses.");

            throw;
        }

        return results;
    }

    // ======================================================
    // COURSES : GET BY ID
    // ======================================================

    public async Task<CourseModel?> GetCourseByIdAsync(
        int courseId)
    {
        if (!IsAuthenticated() || courseId <= 0)
        {
            return null;
        }

        try
        {
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

            return item?.Fields != null
                ? MapCourse(item)
                : null;
        }
        catch (Exception ex)
        {
            _logger.LogError(
                ex,
                "Error fetching course by id: {Id}",
                courseId);

            return null;
        }
    }

    // ======================================================
    // COURSES : CREATE
    // ======================================================

    public async Task<int> CreateCourseAsync(
        CourseModel model)
    {
        EnforceAdmin();

        if (string.IsNullOrWhiteSpace(model.Title))
        {
            throw new InvalidOperationException(
                "Course title is required.");
        }

        var listId = await GetCoursesListIdAsync();

        var fields = BuildCourseFields(model);

        fields.AdditionalData["Archived"] = false;

        var created = await _graphClient
            .Sites[SiteId]
            .Lists[listId]
            .Items
            .PostAsync(new ListItem
            {
                Fields = fields
            });

        _logger.LogDebug(
            "Course created successfully: {Title}",
            model.Title);

        return int.Parse(created!.Id!);
    }

    // ======================================================
    // COURSES : UPDATE
    // ======================================================

    public async Task UpdateCourseAsync(
        CourseModel model)
    {
        EnforceAdmin();

        if (model.Id <= 0)
        {
            throw new InvalidOperationException(
                "Invalid Course ID.");
        }

        var listId = await GetCoursesListIdAsync();

        var fields = BuildCourseFields(model);

        await _graphClient
            .Sites[SiteId]
            .Lists[listId]
            .Items[model.Id.ToString()]
            .Fields
            .PatchAsync(fields);

        _logger.LogDebug(
            "Course updated successfully: {Id}",
            model.Id);
    }

    // ======================================================
    // COURSES : DELETE
    // ======================================================

    public async Task DeleteCourseAsync(
        int courseId)
    {
        EnforceAdmin();

        if (courseId <= 0)
        {
            throw new InvalidOperationException(
                "Invalid Course ID.");
        }

        var listId = await GetCoursesListIdAsync();

        await _graphClient
            .Sites[SiteId]
            .Lists[listId]
            .Items[courseId.ToString()]
            .DeleteAsync();

        _logger.LogDebug(
            "Course deleted successfully: {Id}",
            courseId);
    }

    // ======================================================
    // FIELD BUILDER
    // ======================================================

    private static FieldValueSet BuildCourseFields(
        CourseModel model)
    {
        return new FieldValueSet
        {
            AdditionalData = new Dictionary<string, object?>
            {
                // CORE
                ["Title"] = model.Title,
                ["CourseCode"] = model.CourseCode,

                // CLASSIFICATION
                ["CourseCategory"] = model.CourseCategory,
                ["CourseContentType"] = model.CourseContentType,
                ["CourseType"] = model.CourseType,
                ["Level"] = model.Level,
                ["Language"] = model.Language,

                // STATUS
                ["CourseStatus"] = model.CourseStatus,
                ["ApprovalStatus"] = model.ApprovalStatus,
                ["Archived"] = model.Archived,
                ["CourseVersion"] = model.CourseVersion,
                ["EffectiveFrom"] = model.EffectiveFrom,

                // DELIVERY
                ["DeliveryMode"] = model.DeliveryMode,
                ["IsSelfPaced"] = model.IsSelfPaced,
                ["EnrollmentOpen"] = model.EnrollmentOpen,
                ["EnrollmentType"] = model.EnrollmentType,

                // NUMBER FIELDS
                [FieldDurationHours] = model.DurationMinutes,
                [FieldCreditValue] = model.CreditValue,

                // CONTENT
                ["Description"] = model.Description,
                ["LearningOutcomes"] = model.LearningOutcomes,

                // CERTIFICATION
                ["CertificateIssued"] = model.CertificateIssued,

                // IMAGE
                ["ImageUrl"] = model.ImageUrl
            }
        };
    }

    // ======================================================
    // MAPPER
    // ======================================================

    private CourseModel MapCourse(
        ListItem item)
    {
        var f = item.Fields!;

        return new CourseModel
        {
            // SYSTEM
            Id = int.TryParse(item.Id, out var id)
                ? id
                : 0,

            Created = GetDateTime(f, "Created"),
            Modified = GetDateTime(f, "Modified"),

            CreatedBy =
                GetStringNullable(f, "AuthorLookupValue"),

            ModifiedBy =
                GetStringNullable(f, "EditorLookupValue"),

            CreatedByEmail =
                GetStringNullable(f, "AuthorEmail"),

            ModifiedByEmail =
                GetStringNullable(f, "EditorEmail"),

            // CORE
            Title = GetString(f, "Title"),
            CourseCode = GetString(f, "CourseCode"),

            // CLASSIFICATION
            CourseCategory =
                GetString(f, "CourseCategory"),

            CourseContentType =
                GetString(f, "CourseContentType"),

            CourseType =
                GetString(f, "CourseType"),

            Level =
                GetString(f, "Level"),

            Language =
                GetString(f, "Language"),

            // STATUS
            CourseStatus =
                GetString(f, "CourseStatus"),

            ApprovalStatus =
                GetString(f, "ApprovalStatus"),

            Archived =
                GetBool(f, "Archived"),

            CourseVersion =
                GetString(f, "CourseVersion"),

            EffectiveFrom =
                GetDateTime(f, "EffectiveFrom"),

            // OWNER
            CourseOwnerDisplayName =
                GetStringNullable(
                    f,
                    "CourseOwnerLookupValue"),

            CourseOwnerEmail =
                GetStringNullable(
                    f,
                    "CourseOwnerEmail"),

            // DELIVERY
            DeliveryMode =
                GetString(f, "DeliveryMode"),

            IsSelfPaced =
                GetBool(f, "IsSelfPaced"),

            EnrollmentOpen =
                GetBool(f, "EnrollmentOpen"),

            EnrollmentType =
                GetString(f, "EnrollmentType"),

            // NUMBERS
            DurationMinutes =
                GetDecimal(f, FieldDurationHours),

            CreditValue =
                GetDecimal(f, FieldCreditValue),

            // CONTENT
            Description =
                GetStringNullable(f, "Description"),

            LearningOutcomes =
                GetStringNullable(
                    f,
                    "LearningOutcomes"),

            // CERTIFICATION
            CertificateIssued =
                GetBool(f, "CertificateIssued"),

            // IMAGE
            ImageUrl =
                GetStringNullable(f, "ImageUrl"),

            // ✅ ✅ AI FEED (THIS WAS MISSING)
            AiFeed = GetStringNullable(f, "AiSummary"),
        };
    }
    // ======================================================
    // HELPER : DECIMAL
    // ======================================================

    private static decimal? GetDecimal(
        FieldValueSet fields,
        string key)
    {
        if (fields.AdditionalData == null)
        {
            return null;
        }

        if (!fields.AdditionalData.TryGetValue(
                key,
                out var value))
        {
            return null;
        }

        if (value == null)
        {
            return null;
        }

        return decimal.TryParse(
            value.ToString(),
            out var result)
            ? result
            : null;
    }
}