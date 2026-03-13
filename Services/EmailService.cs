using System.Globalization;
using System.Text;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Users.Item.SendMail;
using SharepointDailyDigest.Models;

namespace SharepointDailyDigest.Services;

/// <summary>Brand styling for digest emails (Streamflo, Masterflo, Dycor).</summary>
internal record BrandInfo(string DisplayName, string AccentColorHex, string? LogoUrl = null, string? Tagline = null);

public class EmailService : IEmailService
{
    private static readonly Dictionary<string, BrandInfo> Brands = new(StringComparer.OrdinalIgnoreCase)
    {
        ["Streamflo"] = new BrandInfo("Stream-Flo", "#003366", "https://streamflo.com/wp-content/themes/streamflo/images/logo.jpg", null),
        ["Stream-Flo"] = new BrandInfo("Stream-Flo", "#003366", "https://streamflo.com/wp-content/themes/streamflo/images/logo.jpg", null),
        ["Masterflo"] = new BrandInfo("Master Flo", "#0066b3", "https://masterflo.com/wp-content/themes/masterflo/images/logo.png", "A Lifetime of Uptime"),
        ["Master Flo"] = new BrandInfo("Master Flo", "#0066b3", "https://masterflo.com/wp-content/themes/masterflo/images/logo.png", "A Lifetime of Uptime"),
        ["Dycor"] = new BrandInfo("Dycor", "#0d7a7a", "https://dycor.com/wp-content/uploads/2020/01/dycor-logo-dark-220.png", "Data Acquisition, Controls, Innovation and Technology"),
    };
    private GraphServiceClient? _graph;
    private string? _sendFromUserId;
    private readonly object _initLock = new();

    public EmailService()
    {
        // Do not validate here so the worker process can start; validate on first use.
    }

    private void EnsureInitialized()
    {
        if (_graph != null) return;
        lock (_initLock)
        {
            if (_graph != null) return;
            var tenantId = Environment.GetEnvironmentVariable("AZURE_TENANT_ID") ?? "";
            var clientId = Environment.GetEnvironmentVariable("AZURE_CLIENT_ID") ?? "";
            var clientSecret = Environment.GetEnvironmentVariable("AZURE_CLIENT_SECRET") ?? "";
            var missing = new List<string>();
            if (string.IsNullOrEmpty(tenantId)) missing.Add("AZURE_TENANT_ID");
            if (string.IsNullOrEmpty(clientId)) missing.Add("AZURE_CLIENT_ID");
            if (string.IsNullOrEmpty(clientSecret)) missing.Add("AZURE_CLIENT_SECRET");
            if (missing.Count > 0)
                throw new InvalidOperationException("Add these Application settings in the Function App: " + string.Join(", ", missing));
            _sendFromUserId = Environment.GetEnvironmentVariable("SEND_FROM_USER_ID") ?? "";
            if (string.IsNullOrEmpty(_sendFromUserId))
                throw new InvalidOperationException("Add Application setting SEND_FROM_USER_ID (object ID or UPN of the mailbox to send from).");
            var credential = new Azure.Identity.ClientSecretCredential(tenantId, clientId, clientSecret);
            _graph = new GraphServiceClient(credential);
        }
    }

    public async Task SendDigestAsync(string toEmail, string listOrLibraryName, IReadOnlyList<ChangedItem> changes, string? brand = null, string? siteName = null, CancellationToken cancellationToken = default)
    {
        if (changes.Count == 0)
            return;
        EnsureInitialized();

        var hasSiteName = !string.IsNullOrWhiteSpace(siteName);
        var subject = hasSiteName
            ? $"SharePoint digest: {siteName} – {listOrLibraryName} – {changes.Count} new or updated item(s) in the last 24 hours"
            : $"SharePoint digest: {listOrLibraryName} – {changes.Count} new or updated item(s) in the last 24 hours";
        var body = BuildDigestHtml(listOrLibraryName, changes, brand, siteName);

        var requestBody = new SendMailPostRequestBody
        {
            Message = new Message
            {
                Subject = subject,
                Body = new ItemBody
                {
                    ContentType = BodyType.Html,
                    Content = body,
                },
                ToRecipients = new List<Recipient>
                {
                    new()
                    {
                        EmailAddress = new EmailAddress { Address = toEmail },
                    },
                },
            },
            SaveToSentItems = true,
        };

        await _graph!.Users[_sendFromUserId!].SendMail.PostAsync(requestBody, cancellationToken: cancellationToken).ConfigureAwait(false);
    }

    private static string BuildDigestHtml(string listOrLibraryName, IReadOnlyList<ChangedItem> changes, string? brand = null, string? siteName = null)
    {
        var sb = new StringBuilder();
        var brandInfo = !string.IsNullOrWhiteSpace(brand) && Brands.TryGetValue(brand.Trim(), out var b) ? b : null;
        var accent = brandInfo?.AccentColorHex ?? "#2563eb";
        var hasSiteName = !string.IsNullOrWhiteSpace(siteName);

        // Email-safe: inline styles, no external CSS
        sb.Append("<html><body style=\"margin:0; padding:0; font-family: 'Segoe UI', system-ui, -apple-system, sans-serif; font-size: 15px; line-height: 1.5; color: #1e293b; background: #f1f5f9;\">");
        sb.Append("<table role=\"presentation\" width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" style=\"background: #f1f5f9; padding: 24px 16px;\"><tr><td align=\"center\">");
        sb.Append("<table role=\"presentation\" width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" style=\"max-width: 560px; background: #ffffff; border-radius: 12px; box-shadow: 0 4px 6px rgba(0,0,0,0.06); overflow: hidden;\"><tr><td>");

        // Header
        sb.Append("<div style=\"background: ").Append(accent).Append("; color: #ffffff; padding: 24px 28px;\">");
        if (brandInfo != null && !string.IsNullOrEmpty(brandInfo.LogoUrl))
            sb.Append("<img src=\"").Append(System.Net.WebUtility.HtmlEncode(brandInfo.LogoUrl)).Append("\" alt=\"").Append(System.Net.WebUtility.HtmlEncode(brandInfo.DisplayName)).Append("\" style=\"max-height: 40px; max-width: 180px; display: block; margin-bottom: 12px;\" />");
        sb.Append("<div style=\"font-size: 22px; font-weight: 700; letter-spacing: -0.02em;\">").Append(System.Net.WebUtility.HtmlEncode(brandInfo?.DisplayName ?? "SharePoint Digest")).Append("</div>");
        if (brandInfo?.Tagline != null)
            sb.Append("<div style=\"font-size: 13px; opacity: 0.9; margin-top: 4px;\">").Append(System.Net.WebUtility.HtmlEncode(brandInfo.Tagline)).Append("</div>");
        sb.Append("</div>");

        // Body
        sb.Append("<div style=\"padding: 28px;\">");
        if (hasSiteName)
            sb.Append("<p style=\"margin: 0 0 6px; font-size: 13px; color: #64748b;\">").Append(System.Net.WebUtility.HtmlEncode(siteName!)).Append("</p>");
        sb.Append("<p style=\"margin: 0 0 20px; font-size: 15px; color: #475569;\">New or changed in the last 24 hours in <strong style=\"color: #1e293b;\">").Append(System.Net.WebUtility.HtmlEncode(listOrLibraryName)).Append("</strong></p>");

        foreach (var c in changes)
        {
            sb.Append("<table role=\"presentation\" width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" style=\"margin-bottom: 12px; background: #f8fafc; border-radius: 8px; border-left: 4px solid ").Append(accent).Append(";\"><tr><td style=\"padding: 14px 16px;\">");
            sb.Append("<a href=\"").Append(System.Net.WebUtility.HtmlEncode(c.WebUrl)).Append("\" style=\"font-weight: 600; color: ").Append(accent).Append("; text-decoration: none; font-size: 15px;\">").Append(System.Net.WebUtility.HtmlEncode(c.Title)).Append("</a>");
            sb.Append("<div style=\"font-size: 13px; color: #64748b; margin-top: 4px;\">");
            sb.Append(FormatModifiedDate(c.Modified));
            if (!string.IsNullOrEmpty(c.ModifiedBy))
                sb.Append(" · ").Append(System.Net.WebUtility.HtmlEncode(c.ModifiedBy));
            sb.Append("</div>");
            sb.Append("</td></tr></table>");
        }

        sb.Append("</div>");

        // Footer
        sb.Append("<div style=\"padding: 16px 28px; border-top: 1px solid #e2e8f0; font-size: 12px; color: #94a3b8;\">");
        sb.Append(System.Net.WebUtility.HtmlEncode(brandInfo?.DisplayName ?? "SharePoint")).Append(" · Daily Digest");
        sb.Append("</div>");

        sb.Append("</td></tr></table></td></tr></table></body></html>");
        return sb.ToString();
    }

    /// <summary>Format modified date for email; shows "Unknown date" if unparsed (default).</summary>
    private static string FormatModifiedDate(DateTimeOffset modified)
    {
        if (modified == default) return "Unknown date";
        var formatted = modified.ToString("MMM d, yyyy h:mm tt", CultureInfo.InvariantCulture);
        var offset = modified.Offset;
        if (offset == TimeSpan.Zero)
            return formatted + " UTC";
        return formatted + " " + (offset >= TimeSpan.Zero ? "+" : "") + offset.ToString(@"hh\:mm", CultureInfo.InvariantCulture);
    }
}
