using Microsoft.Graph;
using Microsoft.Graph.Models;
using MITANZ360Edu.Web.Models;
using System.Text.Json;

namespace MITANZ360Edu.Web.Services;

public partial class SharePointService : IStudentService
{
    // =====================================================
    // ✅ LIST ID
    // =====================================================
    private string GetStudentListId()
    {
        var id = _configuration["SharePoint:Lists:Student_Profiles"];

        if (string.IsNullOrWhiteSpace(id))
            throw new Exception("Student_Profiles List ID missing");

        return id;
    }

    // =====================================================
    // ✅ RETRY (SAFE EXECUTION ONLY)
    // =====================================================
    private async Task<T> ExecuteWithRetryAsync<T>(Func<Task<T>> action)
    {
        for (int i = 1; i <= 5; i++)
        {
            try
            {
                return await action();
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Retry {Attempt} failed", i);

                if (i == 5) throw;

                await Task.Delay(500 * i);
            }
        }

        throw new Exception("Retry failed");
    }
    private async Task ExecuteWithRetryAsync(Func<Task> action)
    {
        await ExecuteWithRetryAsync(async () =>
        {
            await action();
            return true;
        });
    }

    // =====================================================
    // ✅ SEARCH
    // =====================================================
    public async Task<List<StudentProfile>> SearchStudentsAsync(string search)
    {
        var all = await GetStudentsAsync();

        return all.Where(s =>
            (!string.IsNullOrEmpty(s.FirstName) && s.FirstName.Contains(search, StringComparison.OrdinalIgnoreCase)) ||
            (!string.IsNullOrEmpty(s.LastName) && s.LastName.Contains(search, StringComparison.OrdinalIgnoreCase)) ||
            (!string.IsNullOrEmpty(s.StudentNumber) && s.StudentNumber.Contains(search, StringComparison.OrdinalIgnoreCase)) ||
            (!string.IsNullOrEmpty(s.Email) && s.Email.Contains(search, StringComparison.OrdinalIgnoreCase))
        ).ToList();
    }
    // =====================================================
    // ✅ GET ALL ✅ FIXED
    // =====================================================
    public async Task<List<StudentProfile>> GetStudentsAsync()
    {
        var listId = GetStudentListId();

        var result = await ExecuteWithRetryAsync(() =>
            _graphClient
                .Sites[SitePath]
                .Lists[listId]
                .Items
                .GetAsync(cfg =>
                {
                    cfg.QueryParameters.Top = 50;
                    cfg.QueryParameters.Select = new[] { "id", "fields" };
                    cfg.QueryParameters.Expand = new[] { "fields" };
                })
        );

        return result?.Value?.Select(MapStudent).ToList()
            ?? new List<StudentProfile>();
    }

    // =====================================================
    // ✅ GET BY ID ✅ FIXED
    // =====================================================
    public async Task<StudentProfile?> GetStudentAsync(int id)
    {
        var listId = GetStudentListId();

        var request = _graphClient
            .Sites[SitePath]
            .Lists[listId]
            .Items[id.ToString()];

        var item = await ExecuteWithRetryAsync(() =>
            request.GetAsync(cfg =>
            {
                cfg.QueryParameters.Select = new[] { "id", "fields" };
                cfg.QueryParameters.Expand = new[] { "fields" };
            })
        );

        return item == null ? null : MapStudent(item);
    }

    // =====================================================
    // ✅ GET BY USER ID ✅ FIXED
    // =====================================================
    public async Task<StudentProfile?> GetStudentByUserIdAsync(string userId)
    {
        var listId = GetStudentListId();

        var request = _graphClient
            .Sites[SitePath]
            .Lists[listId]
            .Items;

        var result = await ExecuteWithRetryAsync(() =>
            request.GetAsync(cfg =>
            {
                cfg.QueryParameters.Filter =
                    $"fields/{SP.UserId} eq '{userId.Replace("'", "''")}'";

                cfg.QueryParameters.Top = 1;
                cfg.QueryParameters.Select = new[] { "id", "fields" };
                cfg.QueryParameters.Expand = new[] { "fields" };
            })
        );

        var item = result?.Value?.FirstOrDefault();

        return item == null ? null : MapStudent(item);
    }

    // =====================================================
    // ✅ CREATE
    // =====================================================
    public async Task<int> CreateStudentAsync(StudentProfile student)
    {
        var listId = GetStudentListId();

        var request = _graphClient
            .Sites[SitePath]
            .Lists[listId]
            .Items;

        var created = await ExecuteWithRetryAsync(() =>
            request.PostAsync(new ListItem
            {
                Fields = new FieldValueSet
                {
                    AdditionalData = BuildFields(student)
                }
            })
        );

        return int.TryParse(created?.Id, out var id) ? id : 0;
    }

    // =====================================================
    // ✅ UPDATE
    // =====================================================
    public async Task<bool> UpdateStudentAsync(StudentProfile student)
    {
        var listId = GetStudentListId();

        var request = _graphClient
            .Sites[SitePath]
            .Lists[listId]
            .Items[student.Id.ToString()]
            .Fields;

        await ExecuteWithRetryAsync(() =>
            request.PatchAsync(new FieldValueSet
            {
                AdditionalData = BuildFields(student)
            })
        );

        return true;
    }

    // =====================================================
    // ✅ DELETE
    // =====================================================
    public async Task<bool> DeleteStudentAsync(int id)
    {
        var listId = GetStudentListId();

        var request = _graphClient
            .Sites[SitePath]
            .Lists[listId]
            .Items[id.ToString()];

        await ExecuteWithRetryAsync(() =>
            request.DeleteAsync()
        );

        return true;
    }

    // =====================================================
    // ✅ MAP
    // =====================================================
    private StudentProfile MapStudent(ListItem item)
    {
        var f = item.Fields?.AdditionalData;

        return new StudentProfile
        {
            Id = int.TryParse(item.Id, out var id) ? id : 0,

            StudentNumber = GetFieldSafe(f, SP.StudentNumber),
            UserId = GetFieldSafe(f, SP.UserId),

            FirstName = GetFieldSafe(f, SP.FirstName),
            LastName = GetFieldSafe(f, SP.LastName),
            PreferredName = GetFieldSafe(f, SP.PreferredName),

            Email = GetFieldSafe(f, SP.Email),
            Phone = GetFieldSafe(f, SP.Phone),
            Address = GetFieldSafe(f, SP.Address),

            NIC = GetFieldSafe(f, SP.NIC),
            Passport = GetFieldSafe(f, SP.Passport),
            Gender = GetFieldSafe(f, SP.Gender),
            DOB = GetFieldSafe(f, SP.DOB),

            Country = GetFieldSafe(f, SP.Country),
            Faculty = GetFieldSafe(f, SP.Faculty),

            AcademicStatus = GetFieldSafe(f, SP.AcademicStatus),
            StudyMode = GetFieldSafe(f, SP.StudyMode),

            IsActive = GetBoolSafe(f, SP.IsActive)
        };
    }

    // =====================================================
    // ✅ HELPERS
    // =====================================================
    private static string GetFieldSafe(IDictionary<string, object>? fields, string key)
    {
        if (fields == null) return "";
        if (!fields.TryGetValue(key, out var value) || value == null) return "";
        if (value is JsonElement json) return json.ToString();
        return value.ToString() ?? "";
    }
    private static bool GetBoolSafe(IDictionary<string, object>? fields, string key)
    {
        if (fields == null) return false;
        if (!fields.TryGetValue(key, out var value) || value == null) return false;
        if (value is bool b) return b;
        if (value is JsonElement json) return json.GetBoolean();
        return bool.TryParse(value.ToString(), out var result) && result;
    }

    // =====================================================
    // ✅ BUILD STUDENT FIELDS
    // =====================================================
    private Dictionary<string, object> BuildFields(StudentProfile s)
    {
        return new Dictionary<string, object>
        {
            [SP.Title] = s.Title,
            [SP.StudentNumber] = s.StudentNumber,
            [SP.UserId] = s.UserId,

            [SP.Email] = s.Email,
            [SP.FirstName] = s.FirstName,
            [SP.LastName] = s.LastName,
            [SP.PreferredName] = s.PreferredName,

            [SP.NIC] = s.NIC,
            [SP.Passport] = s.Passport,
            [SP.Gender] = s.Gender,
            [SP.DOB] = s.DOB,

            [SP.Phone] = s.Phone,
            [SP.Address] = s.Address,

            [SP.Country] = s.Country,
            [SP.Faculty] = s.Faculty,
            [SP.AcademicStatus] = s.AcademicStatus,
            [SP.StudyMode] = s.StudyMode,

            [SP.IsActive] = s.IsActive
        };
    }

    //🔷 Add method: Generate Temporary Student Number
    public async Task<string> GenerateTempStudentNumberAsync()
    {
        var prefix = "TMP";
        var date = DateTime.UtcNow.ToString("yyyyMMdd");

        var random = new Random();
        var number = random.Next(1000, 9999);

        return $"{prefix}-{date}-{number}";
    }

}