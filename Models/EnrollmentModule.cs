using System.ComponentModel.DataAnnotations;

namespace MITANZ360Edu.Web.Models
{
    public interface IEnrollmentService
    {
        Task<List<Enrollment>> GetEnrollmentsAsync();
        Task<Enrollment?> GetEnrollmentAsync(int id);
        Task<Enrollment?> GetEnrollmentByEnrollmentIdAsync(string enrollmentId);
        Task<int> CreateEnrollmentAsync(Enrollment enrollment);
        Task<bool> UpdateEnrollmentAsync(Enrollment enrollment);
        Task<bool> DeleteEnrollmentAsync(int id);
        Task<List<Enrollment>> SearchEnrollmentsAsync(string searchText);

        Task<string> GenerateTempEnrollmentNumberAsync();
    }

    public static class EnrollmentSP
    {
        // SharePoint System
        public const string Title = "Title";

        public const string EnrollmentId = "field_1";
        public const string EnrollmentNo = "field_2";
        public const string Status = "field_3";

        public const string StudentId = "field_4";
        public const string StudentNumber = "field_5";
        public const string StudentName = "field_6";
        public const string Email = "field_7";
        public const string Mobile = "field_8";

        public const string GroupId = "field_9";
        public const string GroupCode = "field_10";
        public const string GroupName = "field_11";

        public const string CourseId = "field_12";
        public const string CourseCode = "field_13";
        public const string CourseName = "field_14";

        public const string IntakeId = "field_15";
        public const string IntakeName = "field_16";

        public const string CampusId = "field_17";
        public const string CampusName = "field_18";

        public const string EnrollmentDate = "field_19";
        public const string StartDate = "field_20";
        public const string EndDate = "field_21";
        public const string CompletionDate = "field_22";

        public const string AttendanceTarget = "field_23";
        public const string AttendancePercentage = "field_24";

        public const string AIEnabled = "field_25";
        public const string LMSEnabled = "field_26";
        public const string SMSEnabled = "field_27";
        public const string FinanceEnabled = "field_28";
        public const string TeamsEnabled = "field_29";
        public const string ZoomEnabled = "field_30";

        public const string SharePointFolderCreated = "field_31";
        public const string CertificateIssued = "field_32";

        public const string Progress = "field_33";
        public const string Notes = "field_34";
        public const string Metadata = "field_35";

        // SharePoint Readonly
        public const string Created = "Created";
        public const string Modified = "Modified";
        public const string Author = "Author";
        public const string Editor = "Editor";
    }

    public class Enrollment
    {
        // ==========================================
        // PRIMARY
        // ==========================================

        public int Id { get; set; }

        public string Title { get; set; } = "";

        // ==========================================
        // ENROLLMENT
        // ==========================================

        public string EnrollmentId { get; set; } = "";

        [Required]
        public string EnrollmentNo { get; set; } = "";

        public string Status { get; set; } = "Pending";

        // ==========================================
        // STUDENT
        // ==========================================

        public string StudentId { get; set; } = "";

        public string StudentNumber { get; set; } = "";

        public string StudentName { get; set; } = "";

        public string Email { get; set; } = "";

        public string Mobile { get; set; } = "";

        // ==========================================
        // GROUP
        // ==========================================

        public string GroupId { get; set; } = "";

        public string GroupCode { get; set; } = "";

        public string GroupName { get; set; } = "";

        // ==========================================
        // COURSE
        // ==========================================

        public string CourseId { get; set; } = "";

        public string CourseCode { get; set; } = "";

        public string CourseName { get; set; } = "";

        // ==========================================
        // INTAKE
        // ==========================================

        public string IntakeId { get; set; } = "";

        public string IntakeName { get; set; } = "";

        // ==========================================
        // CAMPUS
        // ==========================================

        public string CampusId { get; set; } = "";

        public string CampusName { get; set; } = "";

        // ==========================================
        // DATES
        // ==========================================

        public DateTime? EnrollmentDate { get; set; }

        public DateTime? StartDate { get; set; }

        public DateTime? EndDate { get; set; }

        public DateTime? CompletionDate { get; set; }

        // ==========================================
        // ATTENDANCE
        // ==========================================

        public double AttendanceTarget { get; set; }

        public double AttendancePercentage { get; set; }

        // ==========================================
        // FEATURES
        // ==========================================

        public bool AIEnabled { get; set; }

        public bool LMSEnabled { get; set; }

        public bool SMSEnabled { get; set; }

        public bool FinanceEnabled { get; set; }

        public bool TeamsEnabled { get; set; }

        public bool ZoomEnabled { get; set; }

        // ==========================================
        // LMS
        // ==========================================

        public bool SharePointFolderCreated { get; set; }

        public bool CertificateIssued { get; set; }

        public double Progress { get; set; }

        public string Notes { get; set; } = "";

        public string Metadata { get; set; } = "";

        // ==========================================
        // SHAREPOINT
        // ==========================================

        public DateTime? Created { get; set; }

        public DateTime? Modified { get; set; }

        public string CreatedBy { get; set; } = "";

        public string ModifiedBy { get; set; } = "";

        // ==========================================
        // HELPERS
        // ==========================================

        public void Validate()
        {
            if (string.IsNullOrWhiteSpace(EnrollmentNo))
                throw new Exception("Enrollment Number is required.");

            if (string.IsNullOrWhiteSpace(StudentId))
                throw new Exception("Student is required.");

            if (string.IsNullOrWhiteSpace(CourseId))
                throw new Exception("Course is required.");
        }
    }
}