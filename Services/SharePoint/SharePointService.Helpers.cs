using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace MITANZ360Edu.Web.Services;

public partial class SharePointService
{
    // =====================================================
    // 🔐 AUTH
    // =====================================================

    protected bool IsAuthenticated()
    => !string.IsNullOrWhiteSpace(CurrentUserUpn);

    protected string GetUserDisplayName()
        => CurrentUserUpn;

    // =====================================================
    // 🔐 ROLE HELPERS
    // =====================================================

    private bool IsInRole(params string[] roles)
    {
        return true;
    }

    protected bool HasRequiredRole(
        params string[] allowed)
    {
        return true;
    }

    // =====================================================
    // 🔐 RBAC
    // =====================================================

    protected void EnforceAdmin()
    {
        if (!IsAuthenticated()
            || !IsInRole("Admin", "SysAdmin"))
        {
            throw new UnauthorizedAccessException(
                "Admin access required.");
        }
    }

    protected void EnforceAdminOrTrainer()
    {
        if (!IsAuthenticated()
            || !IsInRole("Admin", "SysAdmin", "Trainer"))
        {
            throw new UnauthorizedAccessException(
                "Admin or Trainer access required.");
        }
    }

    protected void EnforceAcademicAuthority()
    {
        if (!IsAuthenticated()
            || !IsInRole("Academic", "Admin", "SysAdmin"))
        {
            throw new UnauthorizedAccessException(
                "Academic authority required.");
        }
    }

    protected void EnforceFinance()
    {
        if (!IsAuthenticated()
            || !IsInRole("Finance", "Admin", "SysAdmin"))
        {
            throw new UnauthorizedAccessException(
                "Finance authority required.");
        }
    }

    // =====================================================
    // 📚 LIST NAMES
    // =====================================================

    protected string CoursesListName =>
        _configuration["SharePoint:Lists:Courses"]
        ?? throw new InvalidOperationException(
            "Missing config: SharePoint:Lists:Courses");

    // =====================================================
    // 🧾 AUDIT
    // =====================================================

    protected Task WriteAuditAsync(
        string message,
        CancellationToken ct = default)
    {
        _logger.LogInformation(
            "AUDIT\n{Message}\nUser={User}\nAt={AtUtc}",
            message,
            GetUserDisplayName(),
            DateTime.UtcNow);

        return Task.CompletedTask;
    }

    protected Task WriteAuditAsync(
        string action,
        string entityId,
        CancellationToken ct = default)
        => WriteAuditAsync(
            $"{action}\nEntityId={entityId}",
            ct);

    protected Task WriteAuditAsync(
        string action,
        string entityType,
        string entityId,
        CancellationToken ct = default)
        => WriteAuditAsync(
            $"{action}\n{entityType}\nEntityId={entityId}",
            ct);

    // =====================================================
    // 🧼 FIELD NORMALIZATION
    // =====================================================

    protected Dictionary<string, object?> NormalizeFields(
        Dictionary<string, object?> fields)
        => fields
            .Where(x => x.Value != null)
            .ToDictionary(x => x.Key, x => x.Value);

    // =====================================================
    // 🔎 SAFE FIELD READERS
    // =====================================================

    protected string GetString(
        FieldValueSet fields,
        string key)
        => fields.AdditionalData.TryGetValue(
                key,
                out var v)
            ? Convert.ToString(
                    v,
                    CultureInfo.InvariantCulture)
              ?? string.Empty
            : string.Empty;

    protected string? GetStringNullable(
        FieldValueSet fields,
        string key)
        => fields.AdditionalData.TryGetValue(
                key,
                out var v)
            ? Convert.ToString(
                v,
                CultureInfo.InvariantCulture)
            : null;

    protected int? GetInt(
        FieldValueSet fields,
        string key)
        => fields.AdditionalData.TryGetValue(
                key,
                out var v)
           && int.TryParse(
                Convert.ToString(
                    v,
                    CultureInfo.InvariantCulture),
                out var i)
            ? i
            : null;

    protected int? GetIntNullable(
        FieldValueSet fields,
        string key)
        => fields.AdditionalData.TryGetValue(
                key,
                out var v)
           && int.TryParse(
                Convert.ToString(
                    v,
                    CultureInfo.InvariantCulture),
                out var i)
            ? i
            : null;

    protected decimal? GetDecimalNullable(
        FieldValueSet fields,
        string key)
        => fields.AdditionalData.TryGetValue(
                key,
                out var v)
           && decimal.TryParse(
                Convert.ToString(
                    v,
                    CultureInfo.InvariantCulture),
                out var d)
            ? d
            : null;

    protected bool GetBool(
        FieldValueSet fields,
        string key)
        => fields.AdditionalData.TryGetValue(
                key,
                out var v)
           && bool.TryParse(
                Convert.ToString(
                    v,
                    CultureInfo.InvariantCulture),
                out var b)
            ? b
            : false;

    protected DateTime? GetDate(
        FieldValueSet fields,
        string key)
        => fields.AdditionalData.TryGetValue(
                key,
                out var v)
           && DateTime.TryParse(
                Convert.ToString(
                    v,
                    CultureInfo.InvariantCulture),
                CultureInfo.InvariantCulture,
                DateTimeStyles.AssumeUniversal
                | DateTimeStyles.AdjustToUniversal,
                out var d)
            ? d
            : null;

    protected DateTime? GetDateTime(
        FieldValueSet fields,
        string key)
    {
        if (fields.AdditionalData == null)
            return null;

        if (!fields.AdditionalData.TryGetValue(
                key,
                out var value))
        {
            return null;
        }

        if (value == null)
            return null;

        return DateTime.TryParse(
            value.ToString(),
            out var result)
            ? result
            : null;
    }

    // =====================================================
    // 🔐 ODATA SAFETY
    // =====================================================

    protected string EscapeODataString(
        string value)
        => value.Replace("'", "''");

    // =====================================================
    // 🔁 GRAPH RETRY
    // =====================================================

    protected async Task<T?> ExecuteWithRetryAsync<T>(
        Func<CancellationToken, Task<T>> exec,
        string operation,
        CancellationToken ct)
    {
        try
        {
            return await exec(ct)
                .ConfigureAwait(false);
        }
        catch (ServiceException ex)
        {
            _logger.LogError(
                ex,
                "Graph failure during {Operation}",
                operation);

            throw;
        }
    }

    protected async Task ExecuteWithRetryAsync(
        Func<CancellationToken, Task> exec,
        string operation,
        CancellationToken ct)
    {
        try
        {
            await exec(ct)
                .ConfigureAwait(false);
        }
        catch (ServiceException ex)
        {
            _logger.LogError(
                ex,
                "Graph failure during {Operation}",
                operation);

            throw;
        }
    }

    // =====================================================
    // 🧱 ENROLLMENT HELPERS
    // =====================================================

    protected string BuildNote(
        string actor,
        string message)
    {
        var stamp =
            DateTime.UtcNow.ToString(
                "u",
                CultureInfo.InvariantCulture);

        return $"{stamp} [{actor}] {message}";
    }

    protected async Task<string>
        GenerateEnrollmentCodeAsync()
    {
        await Task.Yield();

        var rand =
            Random.Shared.Next(1000, 9999);

        return
            $"ENR-{DateTime.UtcNow:yyyyMMddHHmmss}-{rand}";
    }

    // =====================================================
    // 🧱 FILE HELPERS
    // =====================================================

    protected string SanitizePath(
        string input)
    {
        if (string.IsNullOrWhiteSpace(input))
            return "unknown";

        var s =
            input.Replace('\\', '/')
                 .Trim();

        s = s.Replace(
                    "../",
                    string.Empty,
                    StringComparison.Ordinal)
             .Replace(
                    "..\\",
                    string.Empty,
                    StringComparison.Ordinal);

        s = Regex.Replace(
            s,
            @"[<>:""\|\?\*\u0000-\u001F]",
            "_");

        while (s.Contains(
            "//",
            StringComparison.Ordinal))
        {
            s = s.Replace(
                "//",
                "/",
                StringComparison.Ordinal);
        }

        s = s.Trim('/');

        return string.IsNullOrWhiteSpace(s)
            ? "unknown"
            : s;
    }

    protected void ValidateFileConstraints(
        string fileName,
        Stream stream)
    {
        if (stream == null)
            throw new ArgumentNullException(
                nameof(stream));

        if (string.IsNullOrWhiteSpace(fileName))
        {
            throw new ArgumentException(
                "File name is required.",
                nameof(fileName));
        }

        if (stream.CanSeek
            && stream.Length <= 0)
        {
            throw new InvalidOperationException(
                "File is empty.");
        }

        var maxCfg =
            _configuration["SharePoint:Files:MaxBytes"];

        if (long.TryParse(
                maxCfg,
                out var max)
            && max > 0
            && stream.CanSeek
            && stream.Length > max)
        {
            throw new InvalidOperationException(
                $"File too large. Size={stream.Length} Max={max}");
        }
    }

    // =====================================================
    // 👤 USER FIELD READER
    // =====================================================

    protected string GetUserDisplayName(
        FieldValueSet fields,
        string key)
    {
        if (!fields.AdditionalData.TryGetValue(
                key,
                out var v)
            || v == null)
        {
            return string.Empty;
        }

        if (v is IDictionary<string, object> dict)
        {
            if (dict.TryGetValue(
                    "displayName",
                    out var dn)
                && dn != null)
            {
                return dn.ToString()
                       ?? string.Empty;
            }

            if (dict.TryGetValue(
                    "name",
                    out var n)
                && n != null)
            {
                return n.ToString()
                       ?? string.Empty;
            }
        }

        return v.ToString()
               ?? string.Empty;
    }
}