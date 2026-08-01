using System.ComponentModel.DataAnnotations;

namespace MITANZ360Edu.Web.Models
{
    public interface IGroupService
    {
        Task<List<Group>> GetGroupsAsync();
        Task<Group?> GetGroupAsync(int id);
        Task<Group?> GetGroupByGroupIdAsync(string groupId);
        Task<int> CreateGroupAsync(Group group);
        Task<bool> UpdateGroupAsync(Group group);
        Task<bool> DeleteGroupAsync(int id);
        Task<List<Group>> SearchGroupsAsync(string searchText);
        Task<string> GenerateTempGroupNumberAsync();
    }

    public static class GroupSP
    {
        // SharePoint System
        public const string Title = "Title";

        // Group
        public const string GroupId = "field_1";
        public const string GroupCode = "field_2";
        public const string GroupName = "field_3";
        public const string Status = "field_4";

        // Course
        public const string CourseCode = "field_6";
        public const string CourseName = "field_7";

        // Intake
        public const string IntakeId = "field_8";
        public const string IntakeName = "field_9";

        // Campus
        public const string CampusId = "field_10";
        public const string CampusName = "field_11";

        // Delivery
        public const string DeliveryMode = "field_12";
        public const string StudyMode = "field_13";

        // Capacity
        public const string MaxStudents = "field_14";
        public const string CurrentStudents = "field_15";

        // Dates
        public const string StartDate = "field_16";
        public const string EndDate = "field_17";
        public const string OrientationDate = "field_18";

        // Staff
        public const string TutorId = "field_19";
        public const string TutorName = "field_20";
        public const string AssistantTutorId = "field_21";
        public const string AssistantTutorName = "field_22";
        public const string AcademicManagerId = "field_23";
        public const string AcademicManagerName = "field_24";
        public const string AdminId = "field_25";
        public const string AdminName = "field_26";

        // Microsoft Teams
        public const string TeamsName = "field_27";
        public const string TeamsUrl = "field_28";
        public const string TeamsMeetingId = "field_29";
        public const string TeamsChannel = "field_30";

        // Zoom
        public const string ZoomUrl = "field_31";
        public const string ZoomMeetingId = "field_32";
        public const string ZoomPasscode = "field_33";

        // Calendar
        public const string CalendarUrl = "field_34";

        // Timetable
        public const string TimetableJson = "field_35";

        // Regional
        public const string CountryCode = "field_36";
        public const string TimeZone = "field_37";
        public const string Language = "field_38";

        // Status
        public const string IsActive = "field_39";

        // Metadata
        public const string Metadata = "field_40";

        // SharePoint Readonly
        public const string Created = "Created";
        public const string Modified = "Modified";
        public const string Author = "Author";
        public const string Editor = "Editor";
    }

    public class Group
    {
        public int Id { get; set; }

        public string Title { get; set; } = string.Empty;

        // Group
        [Required, StringLength(50)]
        public string GroupId { get; set; } = string.Empty;

        [Required, StringLength(50)]
        public string GroupCode { get; set; } = string.Empty;

        [Required, StringLength(200)]
        public string GroupName { get; set; } = string.Empty;

        [StringLength(50)]
        public string Status { get; set; } = string.Empty;

        // Course
        public string CourseCode { get; set; } = string.Empty;
        public string CourseName { get; set; } = string.Empty;

        // Intake
        public string IntakeId { get; set; } = string.Empty;
        public string IntakeName { get; set; } = string.Empty;

        // Campus
        public string CampusId { get; set; } = string.Empty;
        public string CampusName { get; set; } = string.Empty;

        // Delivery
        public string DeliveryMode { get; set; } = string.Empty;
        public string StudyMode { get; set; } = string.Empty;

        // Capacity
        public int? MaxStudents { get; set; }
        public int? CurrentStudents { get; set; }

        // Dates
        public DateTime? StartDate { get; set; }
        public DateTime? EndDate { get; set; }
        public DateTime? OrientationDate { get; set; }

        // Staff
        public string TutorId { get; set; } = string.Empty;
        public string TutorName { get; set; } = string.Empty;

        public string AssistantTutorId { get; set; } = string.Empty;
        public string AssistantTutorName { get; set; } = string.Empty;

        public string AcademicManagerId { get; set; } = string.Empty;
        public string AcademicManagerName { get; set; } = string.Empty;

        public string AdminId { get; set; } = string.Empty;
        public string AdminName { get; set; } = string.Empty;

        // Microsoft Teams
        public string TeamsName { get; set; } = string.Empty;
        public string TeamsUrl { get; set; } = string.Empty;
        public string TeamsMeetingId { get; set; } = string.Empty;
        public string TeamsChannel { get; set; } = string.Empty;

        // Zoom
        public string ZoomUrl { get; set; } = string.Empty;
        public string ZoomMeetingId { get; set; } = string.Empty;
        public string ZoomPasscode { get; set; } = string.Empty;

        // Calendar
        public string CalendarUrl { get; set; } = string.Empty;

        // Timetable
        public string TimetableJson { get; set; } = string.Empty;

        // Regional
        public string CountryCode { get; set; } = string.Empty;
        public string TimeZone { get; set; } = string.Empty;
        public string Language { get; set; } = string.Empty;

        // Status
        public bool IsActive { get; set; } = true;

        // Metadata
        public string Metadata { get; set; } = string.Empty;

        // SharePoint Readonly
        public DateTime? Created { get; set; }
        public DateTime? Modified { get; set; }
        public string? Author { get; set; }
        public string? Editor { get; set; }

        // Helper Properties
        public string DisplayName => $"{GroupCode} - {GroupName}";

        public int AvailableSeats =>
            Math.Max(0, (MaxStudents ?? 0) - (CurrentStudents ?? 0));

        // Helper Methods
        public Group Clone()
        {
            return (Group)MemberwiseClone();
        }

        public void Normalize()
        {
            Title = Title.Trim();

            GroupId = GroupId.Trim();
            GroupCode = GroupCode.Trim();
            GroupName = GroupName.Trim();
            Status = Status.Trim();

            CourseCode = CourseCode.Trim();
            CourseName = CourseName.Trim();

            IntakeId = IntakeId.Trim();
            IntakeName = IntakeName.Trim();

            CampusId = CampusId.Trim();
            CampusName = CampusName.Trim();

            DeliveryMode = DeliveryMode.Trim();
            StudyMode = StudyMode.Trim();

            TutorId = TutorId.Trim();
            TutorName = TutorName.Trim();

            AssistantTutorId = AssistantTutorId.Trim();
            AssistantTutorName = AssistantTutorName.Trim();

            AcademicManagerId = AcademicManagerId.Trim();
            AcademicManagerName = AcademicManagerName.Trim();

            AdminId = AdminId.Trim();
            AdminName = AdminName.Trim();

            TeamsName = TeamsName.Trim();
            TeamsUrl = TeamsUrl.Trim();
            TeamsMeetingId = TeamsMeetingId.Trim();
            TeamsChannel = TeamsChannel.Trim();

            ZoomUrl = ZoomUrl.Trim();
            ZoomMeetingId = ZoomMeetingId.Trim();
            ZoomPasscode = ZoomPasscode.Trim();

            CalendarUrl = CalendarUrl.Trim();

            TimetableJson = TimetableJson.Trim();

            CountryCode = CountryCode.Trim();
            TimeZone = TimeZone.Trim();
            Language = Language.Trim();

            Metadata = Metadata.Trim();
        }
    }
}