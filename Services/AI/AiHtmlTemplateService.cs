using System.Text;
using Microsoft.Extensions.Logging;

namespace MITANZ360Edu.Web.Services.AI;

public sealed class AiHtmlTemplateService
{
    private readonly ILogger<AiHtmlTemplateService> _logger;

    public AiHtmlTemplateService(
        ILogger<AiHtmlTemplateService> logger)
    {
        _logger = logger;
    }

    public string GenerateHtml(
        string title,
        object? content)
    {
        try
        {
            var html =
                new StringBuilder();

            html.AppendLine("<!DOCTYPE html>");
            html.AppendLine("<html>");
            html.AppendLine("<head>");
            html.AppendLine("<meta charset='utf-8' />");
            html.AppendLine("<title>MITANZ360Edu AI Report</title>");
            html.AppendLine("<style>");
            html.AppendLine("body { font-family: Arial; margin: 20px; }");
            html.AppendLine("h1 { color: #003366; }");
            html.AppendLine("pre { background: #f5f5f5; padding: 15px; }");
            html.AppendLine("</style>");
            html.AppendLine("</head>");
            html.AppendLine("<body>");

            html.AppendLine($"<h1>{title}</h1>");

            html.AppendLine("<pre>");
            html.AppendLine(content?.ToString());
            html.AppendLine("</pre>");

            html.AppendLine("</body>");
            html.AppendLine("</html>");

            _logger.LogInformation(
                "HTML template generated successfully.");

            return html.ToString();
        }
        catch (Exception ex)
        {
            _logger.LogError(
                ex,
                "HTML generation failed.");

            throw;
        }
    }
}
