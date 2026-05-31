using System.ComponentModel.DataAnnotations;
using MITANZ360Edu.Web.Components.Pages.ModelContent;

namespace MITANZ360Edu.Web.Models;

public class ModelContentModel
{
    // =====================================================
    // SYSTEM
    // =====================================================

    public int Id
    {
        get;
        set;
    }

    // =====================================================
    // BASIC
    // =====================================================

    [Required]
    [MaxLength(255)]
    public string Title
    {
        get;
        set;
    } = string.Empty;

    [Required]
    [MaxLength(100)]
    public string CourseId
    {
        get;
        set;
    } = string.Empty;

    [MaxLength(100)]
    public string ModuleNo
    {
        get;
        set;
    } = string.Empty;

    // =====================================================
    // CONTENT
    // =====================================================

    public ContentTypeEnum ContentType
    {
        get;
        set;
    } = ContentTypeEnum.Module;

    public string ContentBody
    {
        get;
        set;
    } = string.Empty;

    public string LearningOutcomes
    {
        get;
        set;
    } = string.Empty;

    // =====================================================
    // RESOURCES
    // =====================================================

    public string FileUrl
    {
        get;
        set;
    } = string.Empty;

    public string ResourceUrl
    {
        get;
        set;
    } = string.Empty;

    public string ParentId
    {
        get;
        set;
    } = string.Empty;

    // =====================================================
    // FLAGS
    // =====================================================

    public bool IsPublished
    {
        get;
        set;
    }

    public bool IsMandatory
    {
        get;
        set;
    }

    public bool IsVisibleToStudents
    {
        get;
        set;
    }

    // =====================================================
    // METADATA
    // =====================================================

    public decimal? DurationMinutes
    {
        get;
        set;
    }

    public int? Order
    {
        get;
        set;
    }

    // =====================================================
    // AUDIT
    // =====================================================

    public string CreatedBy
    {
        get;
        set;
    } = string.Empty;

    public DateTime? CreatedDate
    {
        get;
        set;
    }

    public string ModifiedBy
    {
        get;
        set;
    } = string.Empty;

    public DateTime? ModifiedDate
    {
        get;
        set;
    }

    public DateTime? Created
    {
        get;
        set;
    }

    public DateTime? Modified
    {
        get;
        set;
    }

    public string PublishedBy
    {
        get;
        set;
    } = string.Empty;

    public DateTime? PublishedDate
    {
        get;
        set;
    }

    // =====================================================
    // AI
    // =====================================================

    public bool IsAiGenerated
    {
        get;
        set;
    }

    public string AiSummary
    {
        get;
        set;
    } = string.Empty;

    public string AiTags
    {
        get;
        set;
    } = string.Empty;

    public string AiMetadata
    {
        get;
        set;
    } = string.Empty;
    public decimal AiScore { get; set; }

    // =====================================================
    // UI HELPERS
    // =====================================================

    public string PublishStatus =>
        IsPublished
            ? "Published"
            : "Draft";

    public string VisibilityStatus =>
        IsVisibleToStudents
            ? "Visible"
            : "Hidden";

    public string DurationDisplay =>
        DurationMinutes.HasValue
            ? $"{DurationMinutes:N0} Minutes"
            : "N/A";

    public string ContentTypeDisplay =>
        ContentType.ToString();

    // =====================================================
    // HELPERS
    // =====================================================

    public void Normalize()
    {
        Title =
            Title?.Trim()
            ?? string.Empty;

        CourseId =
            CourseId?.Trim()
            ?? string.Empty;

        ModuleNo =
            ModuleNo?.Trim()
            ?? string.Empty;

        ContentBody =
            ContentBody?.Trim()
            ?? string.Empty;

        LearningOutcomes =
            LearningOutcomes?.Trim()
            ?? string.Empty;

        FileUrl =
            FileUrl?.Trim()
            ?? string.Empty;

        ResourceUrl =
            ResourceUrl?.Trim()
            ?? string.Empty;

        ParentId =
            ParentId?.Trim()
            ?? string.Empty;

        AiSummary =
            AiSummary?.Trim()
            ?? string.Empty;

        AiTags =
            AiTags?.Trim()
            ?? string.Empty;

        AiMetadata =
            AiMetadata?.Trim()
            ?? string.Empty;
    }
}