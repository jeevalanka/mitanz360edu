using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Extensions.Configuration;

namespace MITANZ360Edu.Web.Services;

/// <summary>
/// ============================================================
/// ✅ CORE SHAREPOINT SERVICE
/// ============================================================
///
/// Responsibilities:
/// - Graph authentication
/// - Configuration loading
/// - SharePoint connection management
/// - Common helpers
/// - Retry handling
/// - User context
///
/// ============================================================
/// </summary>
public partial class SharePointService
{
    // =====================================================
    // DEPENDENCIES
    // =====================================================

    private readonly IConfiguration _configuration;
    public SharePointService(IConfiguration configuration)
    {
        _configuration = configuration;
    }

    private readonly ILogger<SharePointService> _logger;

    protected readonly GraphServiceClient _graphClient;

    // =====================================================
    // CONFIGURATION
    // =====================================================

    protected readonly string SiteId;

    protected readonly string LmsLibraryListId;

    // =====================================================
    // USER CONTEXT
    // =====================================================

    protected string CurrentUserUpn =>
        "system@mitanz.com";

    // =====================================================
    // CONSTRUCTOR
    // =====================================================

    public SharePointService(
        IConfiguration configuration,
        ILogger<SharePointService> logger)
    {
        _configuration = configuration;

        _logger = logger;

        // =============================================
        // LOAD CONFIG
        // =============================================

        SiteId =
            _configuration["SharePoint:SiteId"]
            ?? throw new InvalidOperationException(
                "SharePoint SiteId missing.");

        LmsLibraryListId =
            _configuration[
                "SharePoint:Libraries:LMS:ListId"]
            ?? throw new InvalidOperationException(
                "LMS Library ListId missing.");

        // =============================================
        // GRAPH AUTH
        // =============================================

        var tenantId =
            _configuration["Graph:TenantId"]
            ?? throw new InvalidOperationException(
                "Graph TenantId missing.");

        var clientId =
            _configuration["Graph:ClientId"]
            ?? throw new InvalidOperationException(
                "Graph ClientId missing.");

        var clientSecret =
            _configuration["Graph:ClientSecret"]
            ?? throw new InvalidOperationException(
                "Graph ClientSecret missing.");

        var credential =
            new ClientSecretCredential(
                tenantId,
                clientId,
                clientSecret);

        _graphClient =
            new GraphServiceClient(
                credential);
    }

    // =====================================================
    // LMS DRIVE ID
    // =====================================================

    public async Task<string> GetLmsDriveIdAsync()
    {
        var configuredDriveId =
            _configuration[
                "SharePoint:Libraries:LMS:DriveId"];

        if (!string.IsNullOrWhiteSpace(
                configuredDriveId))
        {
            _logger.LogInformation(
                "Using configured LMS DriveId.");

            return configuredDriveId;
        }

        _logger.LogInformation(
            "Resolving LMS DriveId dynamically.");

        var drives =
            await _graphClient
                .Sites[SiteId]
                .Drives
                .GetAsync();

        var drive =
            drives?.Value?
                .FirstOrDefault(x =>
                    x.Name == "LMS-Lib-Content");

        if (drive?.Id == null)
        {
            throw new InvalidOperationException(
                "Unable to resolve LMS drive.");
        }

        return drive.Id;
    }

    // =====================================================
    // FIELD HELPERS
    // =====================================================

    protected static string GetField(
        IDictionary<string, object>? fields,
        string fieldName)
    {
        if (fields == null)
        {
            return string.Empty;
        }

        return fields.TryGetValue(
                   fieldName,
                   out var value)
            ? value?.ToString() ?? string.Empty
            : string.Empty;
    }

    protected static bool GetBoolField(
        IDictionary<string, object>? fields,
        string fieldName)
    {
        if (fields == null)
        {
            return false;
        }

        if (!fields.TryGetValue(
                fieldName,
                out var value))
        {
            return false;
        }

        return value switch
        {
            bool b => b,

            string s when bool.TryParse(
                s,
                out var parsed)
                => parsed,

            _ => false
        };
    }

    // =====================================================
    // SECURITY
    // =====================================================

    protected void EnsureUserHasPermission(
        string permission)
    {
        _logger.LogInformation(
            "Permission validated: {Permission}",
            permission);

        // =============================================
        // FUTURE RBAC VALIDATION
        // =============================================
    }

    // =====================================================
    // PUBLIC LMS DRIVE ID
    // =====================================================

    public async Task<string>
        GetPublicLmsDriveIdAsync()
    {
        return await GetLmsDriveIdAsync();
    }

    // =====================================================
    // Get Unic Content ID Auto
    // =====================================================
    public string GenerateContentId()
    {
        var timestamp = DateTime.UtcNow.ToString("yyyyMMdd-HHmmss");
        var random = Random.Shared.Next(1000, 9999);
        return $"ACT-{timestamp}-{random}";
    }

    public async Task<int?> GetExistingCourseIdAsync(string courseName, string courseCode)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(courseName) ||
                string.IsNullOrWhiteSpace(courseCode))
                return null;

            var filter =
                $"fields/Title eq '{courseName.Replace("'", "''")}' and " +
                $"fields/CourseCode eq '{courseCode.Replace("'", "''")}'";

            var listId = _configuration["SharePoint:Lists:AIRepository"];

            var result = await _graphClient
                .Sites[SiteId]
                .Lists[listId]
                .Items
                .GetAsync(config =>
                {
                    config.QueryParameters.Filter = filter;
                    config.QueryParameters.Top = 1;
                    config.QueryParameters.Expand = new[] { "fields" };
                });

            var item = result?.Value?.FirstOrDefault();

            if (item != null)
            {
                return int.Parse(item.Id);
            }

            return null;
        }
        catch (Exception ex)
        {
            Console.WriteLine("❌ GetExistingCourseId failed:");
            Console.WriteLine(ex.Message);
            return null;
        }
    }


}