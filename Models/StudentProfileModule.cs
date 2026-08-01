using System.ComponentModel.DataAnnotations;
using System.Text.Json.Serialization;

namespace MITANZ360Edu.Web.Models
{
    public interface IStudentService
    {
        Task<List<StudentProfile>> GetStudentsAsync();
        Task<StudentProfile?> GetStudentAsync(int id);
        Task<StudentProfile?> GetStudentByUserIdAsync(string userId);
        Task<int> CreateStudentAsync(StudentProfile student);
        Task<bool> UpdateStudentAsync(StudentProfile student);
        Task<bool> DeleteStudentAsync(int id);
        Task<List<StudentProfile>> SearchStudentsAsync(string searchText);

        // ✅ ADD THIS LINE (your missing method)
        Task<string> GenerateTempStudentNumberAsync();
    }

    public static class SP
    {
        public const string Title = "Title";                     // Required SharePoint title
        public const string StudentNumber = "StudentNumber";     // Unique student number

        public const string UserId = "field_1";                  // Application user ID
        public const string Email = "field_2";                   // Primary email address

        public const string FirstName = "field_3";               // Student first name
        public const string LastName = "field_4";                // Student last name
        public const string PreferredName = "field_5";           // Preferred display name

        public const string NIC = "field_6";                     // National ID number
        public const string Passport = "field_7";                // Passport number
        public const string DOB = "field_8";                     // Date of birth
        public const string Gender = "field_9";                  // Student gender

        public const string Phone = "field_10";                  // Primary contact number
        public const string Address = "field_11";                // Residential address

        public const string Country = "field_12";                // Country of residence
        public const string Faculty = "field_13";                // Faculty or department
        public const string AcademicStatus = "field_14";         // Student academic status
        public const string StudyMode = "field_15";              // Study mode

        public const string Photo = "field_16";                  // Profile photo
        public const string IsActive = "field_17";               // Active status
        public const string CreatedDate = "field_18";            // Record created date

        public const string Created = "Created";                 // SharePoint created date
        public const string Modified = "Modified";               // SharePoint modified date
        public const string Author = "Author";                   // SharePoint item creator
        public const string Editor = "Editor";                   // SharePoint last editor
    }
    public class StudentProfile
    {
        // ================================
        // ✅ PRIMARY
        // ================================
        public int Id { get; set; }

        // SharePoint default column (MANDATORY)
        public string Title { get; set; } = "";

        // ================================
        // ✅ CORE
        // ================================
        public string UserId { get; set; } = "";

        [Required]
        public string StudentNumber { get; set; } = "";

        // ================================
        // ✅ NAME
        // ================================
        [Required]
        public string FirstName { get; set; } = "";

        [Required]
        public string LastName { get; set; } = "";

        public string PreferredName { get; set; } = "";

        // ================================
        // ✅ CONTACT
        // ================================
        [Required, EmailAddress]
        public string Email { get; set; } = "";

        public string Phone { get; set; } = "";

        public string Address { get; set; } = "";

        // ================================
        // ✅ IDENTITY
        // ================================
        public string NIC { get; set; } = "";

        [JsonPropertyName("Passport")]
        public string Passport { get; set; } = "";

        public string Gender { get; set; } = "";

        // ✅ SharePoint TEXT column (IMPORTANT)
        public string DOB { get; set; } = "";

        // ✅ Helper (SAFE – UI only)
        [JsonIgnore]
        public DateTime? DOB_Date
        {
            get => DateTime.TryParse(DOB, out var d) ? d : null;
            set => DOB = value?.ToString("yyyy-MM-dd") ?? "";
        }

        // ================================
        // ✅ ACADEMIC
        // ================================
        public string Country { get; set; } = "";
        public string Faculty { get; set; } = "";
        public string AcademicStatus { get; set; } = "";
        public string StudyMode { get; set; } = "";

        // ================================
        // ✅ MEDIA
        // ================================
        public string Photo { get; set; } = "";

        // ================================
        // ✅ SYSTEM
        // ================================
        public bool IsActive { get; set; } = true;

        public string CreatedDate { get; set; } = "";

        // SharePoint system timestamps
        public DateTime? Created { get; set; }
        public DateTime? Modified { get; set; }

        // SharePoint Person fields (stored as display only)
        [JsonPropertyName("Created By")]
        public string CreatedBy { get; set; } = "";

        [JsonPropertyName("Modified By")]
        public string ModifiedBy { get; set; } = "";

        // ================================
        // ✅ HELPERS (SAFE)
        // ================================

        [JsonIgnore]
        public string FullName =>
            $"{FirstName} {LastName}".Trim();

        public StudentProfile Clone()
        {
            return (StudentProfile)this.MemberwiseClone();
        }

        public void Normalize()
        {
            FirstName = FirstName?.Trim();
            LastName = LastName?.Trim();
            Email = Email?.Trim().ToLower();
            StudentNumber = StudentNumber?.Trim().ToUpper();

            // ✅ Required for SharePoint
            Title = $"{StudentNumber} - {FullName}";
        }

        public void Validate()
        {
            if (string.IsNullOrWhiteSpace(StudentNumber))
                throw new Exception("Student Number is required");

            if (string.IsNullOrWhiteSpace(Email))
                throw new Exception("Email is required");

            if (!Email.Contains("@"))
                throw new Exception("Invalid email format");
        }
    }
}