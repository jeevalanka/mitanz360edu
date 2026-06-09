using System;

    // ======================================================
    // ENROLLMENT MODEL (FINAL & SAFE)
    // ======================================================
    namespace MITANZ360Edu.Web.Models
    {
        public class EnrollmentModel
        {
            // ======================================================
            // SharePoint Item ID
            // ======================================================
            public string? Id { get; set; }

            // ======================================================
            // Identity
            // ======================================================
            public string? EnrollmentCode { get; set; }
            public string? Title { get; set; }

            // ======================================================
            // Lookups (LookupId values)
            // ======================================================
            public int CourseId { get; set; }
            public int StudentId { get; set; }

            // ======================================================
            // ✅ WORKFLOW (REFERENCE DATA — NO ENUMS)
            // ======================================================

            /// <summary>
            /// Enrollment lifecycle status (from SharePoint ReferenceData)
            /// Example: PENDING, APPROVED, ACTIVE
            /// </summary>
            public string Status { get; set; } = "PENDING";

            /// <summary>
            /// Payment lifecycle status (from SharePoint ReferenceData)
            /// Example: PENDING, PAID, FAILED
            /// </summary>
            public string PaymentStatus { get; set; } = "PENDING";

            // ======================================================
            // Dates
            // ======================================================
            public DateTime EnrollmentDate { get; set; }
            public DateTime? StartDate { get; set; }
            public DateTime? EndDate { get; set; }

            // ======================================================
            // Audit / Notes (Append-only in SharePoint)
            // ======================================================
            public string? Notes { get; set; }

            // ======================================================
            // Accountability (Person fields – display names)
            // ======================================================
            public string? ApprovedBy { get; set; }
            public string? PaymentApprovedBy { get; set; }

        }
    }
