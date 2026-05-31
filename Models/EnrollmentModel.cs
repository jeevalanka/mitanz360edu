using System;

namespace MITANZ360Edu.Web.Models
{
    // ======================================================
    // SHAREPOINT CHOICE → ENUM (LOCKED)
    // ======================================================

    // SharePoint Status choices:
    // Pending, Approved, Active, Rejected, On Hold, Cancelled
    public enum EnrollmentStatus
    {
        Pending,
        Approved,
        Active,
        Rejected,
        OnHold,     // maps to "On Hold"
        Cancelled
    }

    // SharePoint PaymentStatus choices:
    // Pending, Paid, Waived, Rejected, Refunded
    public enum PaymentStatus
    {
        Pending,
        Paid,
        Waived,
        Rejected,
        Refunded
    }

    // ======================================================
    // ENROLLMENT MODEL (FINAL & SAFE)
    // ======================================================
    public class EnrollmentModel
    {
        // SharePoint Item ID
        public string? Id { get; set; }

        // Identity
        public string? EnrollmentCode { get; set; }
        public string? Title { get; set; }

        // Lookups (LookupId values)
        public int CourseId { get; set; }
        public int StudentId { get; set; }

        // Workflow (ENUMS ONLY — NO STRINGS ALLOWED)
        public EnrollmentStatus Status { get; set; }
        public PaymentStatus PaymentStatus { get; set; }

        // Dates
        public DateTime EnrollmentDate { get; set; }
        public DateTime? StartDate { get; set; }
        public DateTime? EndDate { get; set; }

        // Audit / Notes (Append-only in SharePoint)
        public string? Notes { get; set; }

        // Accountability (Person fields – display names)
        public string? ApprovedBy { get; set; }
        public string? PaymentApprovedBy { get; set; }

        // System (read-only, optional usage)
        public DateTime Created { get; set; }
        public DateTime Modified { get; set; }
        public string? CreatedBy { get; set; }
        public string? ModifiedBy { get; set; }
    }
}