using System;
using System.Collections.Generic;

namespace MITANZ360Edu.Web.Models
{
    public class CourseModel
    {
        // ======================================================
        // SharePoint System Fields
        // ======================================================
        public int Id { get; set; } // SharePoint List Item ID (PRIMARY KEY)
        public DateTime? Created { get; set; }
        public DateTime? Modified { get; set; }
        public string? CreatedBy { get; set; }
        public string? ModifiedBy { get; set; }
        public string? CreatedByEmail { get; set; }
        public string? ModifiedByEmail { get; set; }

        // ======================================================
        // Core Identity
        // ======================================================
        public string CourseCode { get; set; } = string.Empty;
        public string Title { get; set; } = string.Empty;

        // ======================================================
        // Classification
        // ======================================================
        public string CourseCategory { get; set; } = string.Empty;
        public string CourseContentType { get; set; } = string.Empty; // Single-choice (SharePoint Choice)
        public string CourseType { get; set; } = string.Empty;
        public string Level { get; set; } = string.Empty;
        public string Language { get; set; } = string.Empty;

        // ======================================================
        // Status & Lifecycle
        // ======================================================
        public string CourseStatus { get; set; } = string.Empty;
        public string ApprovalStatus { get; set; } = string.Empty;
        public bool Archived { get; set; }
        public string CourseVersion { get; set; } = string.Empty;
        public DateTime? EffectiveFrom { get; set; }
        public decimal? DurationMinutes { get; set; }

        // ======================================================
        // Ownership (READ-ONLY SAFE FOR NOW)
        // ======================================================
        public string? CourseOwnerDisplayName { get; set; }
        public string? CourseOwnerEmail { get; set; }

        // ======================================================
        // Delivery & Enrollment
        // ======================================================
        public string DeliveryMode { get; set; } = string.Empty;
        public bool IsSelfPaced { get; set; }
        public bool EnrollmentOpen { get; set; }
        public string EnrollmentType { get; set; } = string.Empty;
        public decimal? CreditValue { get; set; }

        // ======================================================
        // Learning Content
        // ======================================================
        public string? Description { get; set; }
        public string? LearningOutcomes { get; set; }

        // ======================================================
        // Certification
        // ======================================================
        public bool CertificateIssued { get; set; }

        // ======================================================
        // Image & Ai
        // ======================================================
        public string? ImageUrl { get; set; }
        public string? AiFeed { get; set; }
    }
}