using Microsoft.Graph.Models;
using MITANZ360Edu.Web.Models;
using System.Collections.Concurrent;
using System.ComponentModel.DataAnnotations;
using System.Security.Claims;

using Enrollment = MITANZ360Edu.Web.Models.Enrollment;

namespace MITANZ360Edu.Web.Services;

// ============================================================
// ✅ ENROLLED COURSE (used by CourseSelector / LibraryTree / DocumentViewer)
// ============================================================
public sealed record EnrolledCourse(string CourseCode, string CourseName, string Status);

public partial class SharePointService : IEnrollmentService
{
    private string? _enrollmentListId;

    private async Task<string> GetEnrollmentListIdAsync()
    {
        if (!string.IsNullOrWhiteSpace(_enrollmentListId))
            return _enrollmentListId!;

        var id = _configuration["SharePoint:Lists:Enrollments"];

        if (string.IsNullOrWhiteSpace(id))
            throw new InvalidOperationException("Missing SharePoint:Lists:Enrollments");

        _enrollmentListId = id;
        return _enrollmentListId!;
    }

    public async Task<List<Enrollment>> GetEnrollmentsAsync()
    {
        if (!IsAuthenticated()) return new();

        var listId = await GetEnrollmentListIdAsync();
        var results = new List<Enrollment>();

        var response = await _graphClient
            .Sites[SiteId]
            .Lists[listId]
            .Items
            .GetAsync(cfg =>
            {
                cfg.QueryParameters.Top = 100;
                cfg.QueryParameters.Select = new[] { "id", "fields" };
                cfg.QueryParameters.Expand = new[] { "fields" };
            });

        while (response?.Value != null)
        {
            foreach (var item in response.Value)
            {
                if (item.Fields != null)
                    results.Add(MapEnrollment(item));
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

    public async Task<Enrollment?> GetEnrollmentAsync(int id)
    {
        if (!IsAuthenticated() || id <= 0) return null;

        var listId = await GetEnrollmentListIdAsync();

        var item = await _graphClient
            .Sites[SiteId]
            .Lists[listId]
            .Items[id.ToString()]
            .GetAsync(cfg =>
            {
                cfg.QueryParameters.Select = new[] { "id", "fields" };
                cfg.QueryParameters.Expand = new[] { "fields" };
            });

        return item?.Fields != null ? MapEnrollment(item) : null;
    }

    public async Task<Enrollment?> GetEnrollmentByEnrollmentIdAsync(string enrollmentId)
    {
        try
        {
            if (!IsAuthenticated() || string.IsNullOrWhiteSpace(enrollmentId))
                return null;

            var listId = await GetEnrollmentListIdAsync();

            var result = await _graphClient
                .Sites[SiteId]
                .Lists[listId]
                .Items
                .GetAsync(cfg =>
                {
                    cfg.QueryParameters.Top = 1;
                    cfg.QueryParameters.Select = new[] { "id", "fields" };
                    cfg.QueryParameters.Expand = new[] { "fields" };
                    cfg.QueryParameters.Filter =
                        $"fields/{EnrollmentSP.EnrollmentId} eq '{EscapeODataString(enrollmentId)}'";
                });

            var item = result?.Value?.FirstOrDefault();

            return item?.Fields != null ? MapEnrollment(item) : null;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "GetEnrollmentByEnrollmentIdAsync failed");
            throw;
        }
    }

    public async Task<int> CreateEnrollmentAsync(Enrollment e)
    {
        var listId = await GetEnrollmentListIdAsync();

        var fields = new Dictionary<string, object>
        {
            [EnrollmentSP.Title] = e.Title,
            [EnrollmentSP.EnrollmentId] = e.EnrollmentId,
            [EnrollmentSP.EnrollmentNo] = e.EnrollmentNo,
            [EnrollmentSP.Status] = e.Status
        };

        AddIfNotEmpty(fields, EnrollmentSP.StudentId, e.StudentId);
        AddIfNotEmpty(fields, EnrollmentSP.StudentName, e.StudentName);
        AddIfNotEmpty(fields, EnrollmentSP.GroupId, e.GroupId);
        AddIfNotEmpty(fields, EnrollmentSP.GroupName, e.GroupName);
        AddIfNotEmpty(fields, EnrollmentSP.CourseCode, e.CourseCode);
        AddIfNotEmpty(fields, EnrollmentSP.CourseName, e.CourseName);
        AddIfNotEmpty(fields, EnrollmentSP.IntakeName, e.IntakeName);
        AddIfNotEmpty(fields, EnrollmentSP.CampusName, e.CampusName);

        if (e.EnrollmentDate.HasValue)
            fields[EnrollmentSP.EnrollmentDate] = e.EnrollmentDate.Value;

        var item = new ListItem
        {
            Fields = new FieldValueSet
            {
                AdditionalData = fields.ToDictionary(k => k.Key, v => (object)v.Value)
            }
        };

        var result = await _graphClient
            .Sites[SiteId]
            .Lists[listId]
            .Items
            .PostAsync(item);

        return int.Parse(result.Id);
    }

    public async Task<bool> UpdateEnrollmentAsync(Enrollment e)
    {
        EnforceAdmin();

        if (e == null || e.Id <= 0)
            throw new Exception("Invalid Enrollment");

        e.Validate();

        var listId = await GetEnrollmentListIdAsync();

        await _graphClient
            .Sites[SiteId]
            .Lists[listId]
            .Items[e.Id.ToString()]
            .Fields
            .PatchAsync(BuildEnrollmentFields(e));

        return true;
    }

    public async Task<bool> DeleteEnrollmentAsync(int id)
    {
        EnforceAdmin();

        if (id <= 0)
            throw new Exception("Invalid Id");

        var listId = await GetEnrollmentListIdAsync();

        await _graphClient
            .Sites[SiteId]
            .Lists[listId]
            .Items[id.ToString()]
            .DeleteAsync();

        return true;
    }

    public async Task<List<Enrollment>> SearchEnrollmentsAsync(string searchText)
    {
        if (!IsAuthenticated()) return new();

        var listId = await GetEnrollmentListIdAsync();

        var safe = EscapeODataString(searchText ?? "");

        var result = await _graphClient
            .Sites[SiteId]
            .Lists[listId]
            .Items
            .GetAsync(cfg =>
            {
                cfg.QueryParameters.Top = 50;
                cfg.QueryParameters.Select = new[] { "id", "fields" };
                cfg.QueryParameters.Expand = new[] { "fields" };

                if (!string.IsNullOrWhiteSpace(safe))
                {
                    cfg.QueryParameters.Filter =
                        $"contains(fields/{EnrollmentSP.EnrollmentNo},'{safe}') or " +
                        $"contains(fields/{EnrollmentSP.StudentName},'{safe}') or " +
                        $"contains(fields/{EnrollmentSP.GroupName},'{safe}')";
                }
            });

        return result?.Value?.Select(MapEnrollment).ToList() ?? new();
    }

    public async Task<string> GenerateTempEnrollmentNumberAsync()
    {
        await Task.Yield();
        return $"ENR-{DateTime.UtcNow:yyyyMMdd}-{Random.Shared.Next(1000, 9999)}";
    }

    private static FieldValueSet BuildEnrollmentFields(Enrollment e)
    {
        return new FieldValueSet
        {
            AdditionalData = new Dictionary<string, object?>
            {
                [EnrollmentSP.Title] = e.Title,
                [EnrollmentSP.EnrollmentNo] = e.EnrollmentNo,
                [EnrollmentSP.Status] = e.Status,
                [EnrollmentSP.StudentId] = e.StudentId,
                [EnrollmentSP.StudentName] = e.StudentName,

                [EnrollmentSP.GroupId] = e.GroupId,
                [EnrollmentSP.GroupName] = e.GroupName,

                // ✅ 🔥 ADD THIS (MISSING)
                [EnrollmentSP.CourseCode] = e.CourseCode,
                [EnrollmentSP.CourseName] = e.CourseName,

                [EnrollmentSP.IntakeName] = e.IntakeName,
                [EnrollmentSP.CampusName] = e.CampusName,

                [EnrollmentSP.EnrollmentDate] = e.EnrollmentDate,

                [EnrollmentSP.Notes] = e.Notes,
                [EnrollmentSP.Metadata] = e.Metadata
            }
        };
    }
    private Enrollment MapEnrollment(ListItem item)
    {
        var f = item.Fields?.AdditionalData;

        return new Enrollment
        {
            Id = int.TryParse(item.Id, out var id) ? id : 0,

            Title = GetField(f, EnrollmentSP.Title),

            EnrollmentId = GetField(f, EnrollmentSP.EnrollmentId),
            EnrollmentNo = GetField(f, EnrollmentSP.EnrollmentNo),
            Status = GetField(f, EnrollmentSP.Status),

            StudentId = GetField(f, EnrollmentSP.StudentId),
            StudentNumber = GetField(f, EnrollmentSP.StudentNumber),
            StudentName = GetField(f, EnrollmentSP.StudentName),
            Email = GetField(f, EnrollmentSP.Email),
            Mobile = GetField(f, EnrollmentSP.Mobile),

            GroupId = GetField(f, EnrollmentSP.GroupId),
            GroupCode = GetField(f, EnrollmentSP.GroupCode),
            GroupName = GetField(f, EnrollmentSP.GroupName),

            CourseId = GetField(f, EnrollmentSP.CourseId),
            CourseCode = GetField(f, EnrollmentSP.CourseCode),
            CourseName = GetField(f, EnrollmentSP.CourseName),

            IntakeId = GetField(f, EnrollmentSP.IntakeId),
            IntakeName = GetField(f, EnrollmentSP.IntakeName),

            CampusId = GetField(f, EnrollmentSP.CampusId),
            CampusName = GetField(f, EnrollmentSP.CampusName),

            EnrollmentDate = GetDate(f, EnrollmentSP.EnrollmentDate),
            StartDate = GetDate(f, EnrollmentSP.StartDate),
            EndDate = GetDate(f, EnrollmentSP.EndDate),
            CompletionDate = GetDate(f, EnrollmentSP.CompletionDate),

            AttendanceTarget = GetDouble(f, EnrollmentSP.AttendanceTarget),
            AttendancePercentage = GetDouble(f, EnrollmentSP.AttendancePercentage),

            AIEnabled = GetBoolField(f, EnrollmentSP.AIEnabled),
            LMSEnabled = GetBoolField(f, EnrollmentSP.LMSEnabled),
            SMSEnabled = GetBoolField(f, EnrollmentSP.SMSEnabled),
            FinanceEnabled = GetBoolField(f, EnrollmentSP.FinanceEnabled),
            TeamsEnabled = GetBoolField(f, EnrollmentSP.TeamsEnabled),
            ZoomEnabled = GetBoolField(f, EnrollmentSP.ZoomEnabled),

            SharePointFolderCreated = GetBoolField(f, EnrollmentSP.SharePointFolderCreated),
            CertificateIssued = GetBoolField(f, EnrollmentSP.CertificateIssued),

            Progress = GetDouble(f, EnrollmentSP.Progress),

            Notes = GetField(f, EnrollmentSP.Notes),
            Metadata = GetField(f, EnrollmentSP.Metadata),

            Created = GetDate(f, EnrollmentSP.Created),
            Modified = GetDate(f, EnrollmentSP.Modified)
        };
    }
    private static DateTime? GetDate(IDictionary<string, object>? f, string key)
    {
        if (f == null || !f.TryGetValue(key, out var v) || v == null)
            return null;

        return DateTime.TryParse(v.ToString(), out var dt) ? dt : null;
    }

    private static double GetDouble(IDictionary<string, object>? fields, string key)
    {
        if (fields == null)
            return 0;

        if (!fields.TryGetValue(key, out var value) || value == null)
            return 0;

        return double.TryParse(value.ToString(), out var result) ? result : 0;
    }

    /// <summary>
    /// Returns enrollments for a single student (server-side OData filter,
    /// not a client-side Where() over the full enrollment list).
    /// </summary>
    public async Task<List<Enrollment>> GetEnrollmentsByStudentIdAsync(string studentId)
    {
        if (!IsAuthenticated() || string.IsNullOrWhiteSpace(studentId))
            return new();

        var listId = await GetEnrollmentListIdAsync();

        var result = await _graphClient
            .Sites[SiteId]
            .Lists[listId]
            .Items
            .GetAsync(cfg =>
            {
                cfg.QueryParameters.Top = 100;
                cfg.QueryParameters.Select = new[] { "id", "fields" };
                cfg.QueryParameters.Expand = new[] { "fields" };
                cfg.QueryParameters.Filter =
                    $"fields/{EnrollmentSP.StudentId} eq '{EscapeODataString(studentId)}'";
            });

        return result?.Value?.Select(MapEnrollment).ToList() ?? new();
    }

    // ================= COURSE ACCESS (used by CourseSelector, LibraryTree, DocumentViewer) =================

    private static readonly string[] StaffRoles = { "Tutor", "Trainer", "Admin", "SysAdmin" };
    private static readonly string[] ActiveEnrollmentStatuses = { "Active", "Enrolled", "Current" };
    private static readonly ConcurrentDictionary<string, (List<EnrolledCourse> Courses, DateTime ExpiresAt)> _courseAccessCache = new();

    /// <summary>True if the signed-in user is Tutor/Trainer/Admin/SysAdmin.</summary>
    public static bool IsStaff(ClaimsPrincipal user)
    {
        var roles = user.Claims
            .Where(c => c.Type == "roles" || c.Type == ClaimTypes.Role)
            .Select(c => c.Value);

        return roles.Any(r => StaffRoles.Contains(r, StringComparer.OrdinalIgnoreCase));
    }

    /// <summary>
    /// Courses visible in the selector: the student's own active enrollments,
    /// or — for staff — every course folder in the library. Cached 5 minutes
    /// per user so CourseSelector/LibraryTree/StatusBar don't each trigger
    /// their own Graph round trip on the same page load.
    /// </summary>
    public async Task<List<EnrolledCourse>> GetVisibleCoursesAsync(ClaimsPrincipal user)
    {
        var userId = ResolveUserId(user);
        if (string.IsNullOrWhiteSpace(userId))
            return new();

        if (_courseAccessCache.TryGetValue(userId, out var cached) && cached.ExpiresAt > DateTime.UtcNow)
            return cached.Courses;

        List<EnrolledCourse> result;

        if (IsStaff(user))
        {
            var folders = await GetCourseFoldersAsync();

            result = folders
                .Select(f => new EnrolledCourse(
                    f.Name,
                    string.IsNullOrWhiteSpace(f.Title) ? f.Name : f.Title,
                    "Staff"))
                .OrderBy(c => c.CourseCode)
                .ToList();
        }
        else
        {
            var student = await GetStudentByUserIdAsync(userId);

            if (student == null)
            {
                _logger.LogWarning("No StudentProfile found for user {UserId}", userId);
                result = new();
            }
            else
            {
                var enrollments = await GetEnrollmentsByStudentIdAsync(student.Id.ToString());

                result = enrollments
                    .Where(e => !string.IsNullOrWhiteSpace(e.CourseCode))
                    .Where(e => ActiveEnrollmentStatuses.Contains(e.Status, StringComparer.OrdinalIgnoreCase))
                    .Select(e => new EnrolledCourse(e.CourseCode, e.CourseName ?? e.CourseCode, e.Status))
                    .DistinctBy(c => c.CourseCode)
                    .OrderBy(c => c.CourseName)
                    .ToList();
            }
        }

        _courseAccessCache[userId] = (result, DateTime.UtcNow.AddMinutes(5));
        return result;
    }

    /// <summary>
    /// Server-side gate: throws if the signed-in user has no right to see
    /// this course. Staff always pass. Students pass only if they have an
    /// active enrollment whose CourseCode matches. Call this before any
    /// Graph read/write scoped to a specific course folder — never rely on
    /// the Course Selector's client-side state alone.
    /// </summary>
    public async Task EnsureCanAccessCourseAsync(ClaimsPrincipal user, string courseCode)
    {
        if (string.IsNullOrWhiteSpace(courseCode))
            throw new ArgumentException("courseCode is required", nameof(courseCode));

        if (IsStaff(user))
            return;

        var courses = await GetVisibleCoursesAsync(user);

        if (!courses.Any(c => string.Equals(c.CourseCode, courseCode, StringComparison.OrdinalIgnoreCase)))
            throw new UnauthorizedAccessException($"User is not enrolled in course '{courseCode}'.");
    }

    /// <summary>Clears the cached course list for a user — call after enrollment changes.</summary>
    public void InvalidateCourseAccessCache(ClaimsPrincipal user)
    {
        var userId = ResolveUserId(user);
        if (!string.IsNullOrWhiteSpace(userId))
            _courseAccessCache.TryRemove(userId, out _);
    }

    private static string? ResolveUserId(ClaimsPrincipal user) =>
        user.FindFirst("preferred_username")?.Value
        ?? user.FindFirst("email")?.Value
        ?? user.FindFirst(ClaimTypes.Email)?.Value
        ?? user.Identity?.Name;
}
