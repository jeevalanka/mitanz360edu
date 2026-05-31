using System.Collections.Generic;

namespace MITANZ360Edu.Web.Models
{

    // =====================================================
    // ENUMS (CANONICAL)
    // =====================================================
    public enum ContentType
    {
        Learning,
        Assessment,
        TutorGuide,
        Moderation,
        Reference
    }
    public enum LearningSubType
    {
        Video,
        PDF,
        SCORM,
        Presentation
    }
    public enum AssessmentType
    {
        Quiz,
        Assignment,
        FinalExam,
        Practical,
        Viva
    }
    public enum AudienceType
    {
        Student,
        Tutor,
        Assessor,
        Admin
    }
    public enum CourseModelType
    {
        Model01,
        Model02,
        Model03,
        Model04,
        Model05,
        Model06,
        Model07,
        Model08,
        Model09,
        Model10,
        Model11,
        Model12,
        Model13,
        Model14,
        Model15,
        Model16,
        Model17,
        Model18,
        Model19,
        Model20,
        Model21,
        Model22,
        Model23,
        Model24,
        Model25
    }
    public enum SourceType
    {
        Internal,
        External,
        Microsoft,
        NZQA,
        AQF,
        AI,
        Student,
        Partner,
        Research,
        Other
    }

    // =====================================================
    // EXT_METADATA BASE (POLYMORPHIC)
    // =====================================================
    public class LmsMetadataBase
    {
        public ContentInfo Content { get; set; } = new();
        public VisibilityInfo Visibility { get; set; } = new();
    }
    public class ContentInfo
    {
        public string Category { get; set; } = string.Empty;
        public string SubType { get; set; } = string.Empty;
    }
    public class VisibilityInfo
    {
        public List<string> AllowedAudiences { get; set; } = new();
        public bool StudentVisible { get; set; }
    }
    // =====================================================
    // LEARNING METADATA
    // =====================================================
    public class LearningMetadata : LmsMetadataBase
    {
        public EngagementInfo Engagement { get; set; } = new();
        public DeliveryInfo Delivery { get; set; } = new();
    }
    public class EngagementInfo
    {
        public int EstimatedMinutes { get; set; }
        public bool EngagementRequired { get; set; }
        public bool CompletionRequired { get; set; }
    }
    public class DeliveryInfo
    {
        public string Mode { get; set; } = string.Empty;
        public string Audience { get; set; } = string.Empty;
    }
    // =====================================================
    // ASSESSMENT METADATA
    // =====================================================
    public class AssessmentMetadata : LmsMetadataBase
    {
        public AssessmentInfo Assessment { get; set; } = new();
    }
    public class AssessmentInfo
    {
        public string AssessmentType { get; set; } = string.Empty;
        public int PassMark { get; set; }
        public int MaxAttempts { get; set; }
        public int WeightPercentage { get; set; }
    }
    // =====================================================
    // RUBRIC METADATA (TUTOR GUIDE)
    // =====================================================
    public class RubricMetadata : LmsMetadataBase
    {
        public RubricInfo Rubric { get; set; } = new();
    }
    public class RubricInfo
    {
        public string RubricId { get; set; } = string.Empty;
        public string LinkedAssessmentContentId { get; set; } = string.Empty;
        public int TotalMarks { get; set; }
        public List<RubricCriterion> Criteria { get; set; } = new();
    }
    public class RubricCriterion
    {
        public string Criterion { get; set; } = string.Empty;
        public int MaxMarks { get; set; }
    }
}