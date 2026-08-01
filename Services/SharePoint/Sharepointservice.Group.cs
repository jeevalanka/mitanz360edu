using System.ComponentModel.DataAnnotations;
using Microsoft.Graph.Models;
using MITANZ360Edu.Web.Models;

using Group = MITANZ360Edu.Web.Models.Group;

namespace MITANZ360Edu.Web.Services;

public partial class SharePointService : IGroupService
{
    private string? _groupsListId;

    private async Task<string> GetGroupsListIdAsync(CancellationToken ct = default)
    {
        if (!string.IsNullOrWhiteSpace(_groupsListId))
            return _groupsListId!;

        var id = _configuration["SharePoint:Lists:Groups"];

        if (string.IsNullOrWhiteSpace(id))
            throw new InvalidOperationException("Missing SharePoint:Lists:Groups");

        _groupsListId = id;
        return _groupsListId!;
    }

    public async Task<List<Group>> GetGroupsAsync()
    {
        if (!IsAuthenticated()) return new();

        var listId = await GetGroupsListIdAsync();
        var results = new List<Group>();

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
                    results.Add(MapGroup(item));
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

    public async Task<Group?> GetGroupAsync(int id)
    {
        if (!IsAuthenticated() || id <= 0) return null;

        var listId = await GetGroupsListIdAsync();

        var item = await _graphClient
            .Sites[SiteId]
            .Lists[listId]
            .Items[id.ToString()]
            .GetAsync(cfg =>
            {
                cfg.QueryParameters.Select = new[] { "id", "fields" };
                cfg.QueryParameters.Expand = new[] { "fields" };
            });

        return item?.Fields != null ? MapGroup(item) : null;
    }

    public async Task<int> CreateGroupAsync(Group group)
    {
        var listId = await GetGroupsListIdAsync();

        var fields = new Dictionary<string, object>
        {
            [GroupSP.Title] = group.Title,
            [GroupSP.GroupId] = group.GroupId,
            [GroupSP.GroupCode] = group.GroupCode,
            [GroupSP.GroupName] = group.GroupName,
            [GroupSP.IsActive] = group.IsActive
        };

        AddIfNotEmpty(fields, GroupSP.Status, group.Status);
        AddIfNotEmpty(fields, GroupSP.CourseName, group.CourseName);
        AddIfNotEmpty(fields, GroupSP.IntakeName, group.IntakeName);
        AddIfNotEmpty(fields, GroupSP.CampusName, group.CampusName);

        if (group.MaxStudents.HasValue)
            fields[GroupSP.MaxStudents] = group.MaxStudents.Value;

        if (group.CurrentStudents.HasValue)
            fields[GroupSP.CurrentStudents] = group.CurrentStudents.Value;

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

    public async Task<bool> UpdateGroupAsync(Group group)
    {
        EnforceAdmin();

        if (group == null || group.Id <= 0)
            throw new Exception("Invalid Group.");

        group.Normalize();
        ValidateGroup(group);

        var listId = await GetGroupsListIdAsync();

        await _graphClient
            .Sites[SiteId]
            .Lists[listId]
            .Items[group.Id.ToString()]
            .Fields
            .PatchAsync(BuildGroupFields(group));

        return true;
    }

    public async Task<bool> DeleteGroupAsync(int id)
    {
        EnforceAdmin();

        if (id <= 0)
            throw new Exception("Invalid Id");

        var listId = await GetGroupsListIdAsync();

        await _graphClient
            .Sites[SiteId]
            .Lists[listId]
            .Items[id.ToString()]
            .DeleteAsync();

        return true;
    }

    public async Task<List<Group>> SearchGroupsAsync(string searchText)
    {
        if (!IsAuthenticated()) return new();

        var listId = await GetGroupsListIdAsync();

        var result = await _graphClient
            .Sites[SiteId]
            .Lists[listId]
            .Items
            .GetAsync(cfg =>
            {
                cfg.QueryParameters.Top = 50;
                cfg.QueryParameters.Select = new[] { "id", "fields" };
                cfg.QueryParameters.Expand = new[] { "fields" };
            });

        return result?.Value?.Select(MapGroup).ToList() ?? new();
    }

    public async Task<string> GenerateTempGroupNumberAsync()
    {
        await Task.Yield();
        return $"GRP-{DateTime.UtcNow:yyyyMMdd}-{Random.Shared.Next(1000, 9999)}";
    }

    private static void AddIfNotEmpty(Dictionary<string, object> dict, string key, string value)
    {
        if (!string.IsNullOrWhiteSpace(value))
            dict[key] = value;
    }

    private static void ValidateGroup(Group group)
    {
        var context = new ValidationContext(group);
        var results = new List<ValidationResult>();

        if (!Validator.TryValidateObject(group, context, results, true))
        {
            var msg = string.Join("; ", results.Select(r => r.ErrorMessage));
            throw new ValidationException(msg);
        }
    }

    private static FieldValueSet BuildGroupFields(Group g)
    {
        return new FieldValueSet
        {
            AdditionalData = new Dictionary<string, object?>
            {
                [GroupSP.Title] = g.Title,
                [GroupSP.GroupId] = g.GroupId,
                [GroupSP.GroupCode] = g.GroupCode,
                [GroupSP.GroupName] = g.GroupName,
                [GroupSP.Status] = g.Status,
                [GroupSP.CourseName] = g.CourseName,
                [GroupSP.IntakeName] = g.IntakeName,
                [GroupSP.CampusName] = g.CampusName,
                [GroupSP.MaxStudents] = g.MaxStudents,
                [GroupSP.CurrentStudents] = g.CurrentStudents,
                [GroupSP.IsActive] = g.IsActive
            }
        };
    }

    private static int? GetGroupInt(IDictionary<string, object>? fields, string key)
    {
        if (fields == null) return null;

        if (!fields.TryGetValue(key, out var value) || value == null)
            return null;

        return int.TryParse(value.ToString(), out var i) ? i : null;
    }

    private static DateTime? GetGroupDateTime(IDictionary<string, object>? fields, string key)
    {
        if (fields == null) return null;

        if (!fields.TryGetValue(key, out var value) || value == null)
            return null;

        return DateTime.TryParse(value.ToString(), out var dt) ? dt : null;
    }

    private Group MapGroup(ListItem item)
    {
        var f = item.Fields?.AdditionalData;

        return new Group
        {
            Id = int.TryParse(item.Id, out var id) ? id : 0,

            Title = GetField(f, GroupSP.Title),

            GroupId = GetField(f, GroupSP.GroupId),
            GroupCode = GetField(f, GroupSP.GroupCode),
            GroupName = GetField(f, GroupSP.GroupName),
            Status = GetField(f, GroupSP.Status),

            CourseCode = GetField(f, GroupSP.CourseCode),
            CourseName = GetField(f, GroupSP.CourseName),

            IntakeId = GetField(f, GroupSP.IntakeId),
            IntakeName = GetField(f, GroupSP.IntakeName),

            CampusId = GetField(f, GroupSP.CampusId),
            CampusName = GetField(f, GroupSP.CampusName),

            DeliveryMode = GetField(f, GroupSP.DeliveryMode),
            StudyMode = GetField(f, GroupSP.StudyMode),

            MaxStudents = GetGroupInt(f, GroupSP.MaxStudents),
            CurrentStudents = GetGroupInt(f, GroupSP.CurrentStudents),

            StartDate = GetGroupDateTime(f, GroupSP.StartDate),
            EndDate = GetGroupDateTime(f, GroupSP.EndDate),
            OrientationDate = GetGroupDateTime(f, GroupSP.OrientationDate),

            TutorId = GetField(f, GroupSP.TutorId),
            TutorName = GetField(f, GroupSP.TutorName),

            AssistantTutorId = GetField(f, GroupSP.AssistantTutorId),
            AssistantTutorName = GetField(f, GroupSP.AssistantTutorName),

            AcademicManagerId = GetField(f, GroupSP.AcademicManagerId),
            AcademicManagerName = GetField(f, GroupSP.AcademicManagerName),

            AdminId = GetField(f, GroupSP.AdminId),
            AdminName = GetField(f, GroupSP.AdminName),

            TeamsName = GetField(f, GroupSP.TeamsName),
            TeamsUrl = GetField(f, GroupSP.TeamsUrl),
            TeamsMeetingId = GetField(f, GroupSP.TeamsMeetingId),
            TeamsChannel = GetField(f, GroupSP.TeamsChannel),

            ZoomUrl = GetField(f, GroupSP.ZoomUrl),
            ZoomMeetingId = GetField(f, GroupSP.ZoomMeetingId),
            ZoomPasscode = GetField(f, GroupSP.ZoomPasscode),

            CalendarUrl = GetField(f, GroupSP.CalendarUrl),

            TimetableJson = GetField(f, GroupSP.TimetableJson),

            CountryCode = GetField(f, GroupSP.CountryCode),
            TimeZone = GetField(f, GroupSP.TimeZone),
            Language = GetField(f, GroupSP.Language),

            Metadata = GetField(f, GroupSP.Metadata),

            IsActive = GetBoolField(f, GroupSP.IsActive)
        };
    }
    public async Task<Group?> GetGroupByGroupIdAsync(string groupId)
    {
        try
        {
            if (!IsAuthenticated() || string.IsNullOrWhiteSpace(groupId))
                return null;

            var listId = await GetGroupsListIdAsync();

            var result = await _graphClient
                .Sites[SiteId]
                .Lists[listId]
                .Items
                .GetAsync(cfg =>
                {
                    cfg.QueryParameters.Top = 1;
                    cfg.QueryParameters.Select = new[] { "id", "fields" };
                    cfg.QueryParameters.Expand = new[] { "fields" };
                    cfg.QueryParameters.Filter = $"fields/{GroupSP.GroupId} eq '{EscapeODataString(groupId)}'";
                });

            var item = result?.Value?.FirstOrDefault();

            return item?.Fields != null ? MapGroup(item) : null;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "GetGroupByGroupIdAsync failed for {GroupId}", groupId);
            throw;
        }
    }
}