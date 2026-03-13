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

    public async Task SendDigestAsync(string toEmail, string listOrLibraryName, IReadOnlyList<ChangedItem> changes, string? brand = null, CancellationToken cancellationToken = default)
    {
        if (changes.Count == 0)
            return;
        EnsureInitialized();

        var subject = $"SharePoint digest: {listOrLibraryName} – {changes.Count} new or updated item(s) in the last 24 hours";
        var body = BuildDigestHtml(listOrLibraryName, changes, brand);

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

    private static string BuildDigestHtml(string listOrLibraryName, IReadOnlyList<ChangedItem> changes, string? brand = null)
    {
        var sb = new StringBuilder();
        var brandInfo = !string.IsNullOrWhiteSpace(brand) && Brands.TryGetValue(brand.Trim(), out var b) ? b : null;

        sb.Append("<html><body style=\"font-family: Segoe UI, Arial, sans-serif; color: #333; max-width: 600px;\">");

        if (brandInfo != null)
        {
            sb.Append("<div style=\"border-bottom: 4px solid ").Append(brandInfo.AccentColorHex).Append("; padding-bottom: 12px; margin-bottom: 20px;\">");
            if (!string.IsNullOrEmpty(brandInfo.LogoUrl))
                sb.Append("<img src=\"").Append(System.Net.WebUtility.HtmlEncode(brandInfo.LogoUrl)).Append("\" alt=\"").Append(System.Net.WebUtility.HtmlEncode(brandInfo.DisplayName)).Append("\" style=\"max-height: 48px; max-width: 220px; display: block; margin-bottom: 8px;\" />");
            sb.Append("<span style=\"font-size: 22px; font-weight: bold; color: ").Append(brandInfo.AccentColorHex).Append(";\">").Append(System.Net.WebUtility.HtmlEncode(brandInfo.DisplayName)).Append("</span>");
            if (!string.IsNullOrEmpty(brandInfo.Tagline))
                sb.Append("<br/><span style=\"font-size: 12px; color: #666;\">").Append(System.Net.WebUtility.HtmlEncode(brandInfo.Tagline)).Append("</span>");
            sb.Append("</div>");
        }

        sb.Append("<p>Summary of new or changed items in the past 24 hours for <strong>").Append(System.Net.WebUtility.HtmlEncode(listOrLibraryName)).Append("</strong>.</p>");
        sb.Append("<ul>");
        foreach (var c in changes)
        {
            sb.Append("<li>");
            sb.Append("<a href=\"").Append(System.Net.WebUtility.HtmlEncode(c.WebUrl)).Append("\" style=\"color: ").Append(brandInfo?.AccentColorHex ?? "#0066cc").Append(";\">").Append(System.Net.WebUtility.HtmlEncode(c.Title)).Append("</a>");
            sb.Append(" – ").Append(FormatModifiedDate(c.Modified));
            if (!string.IsNullOrEmpty(c.ModifiedBy))
                sb.Append(" (by ").Append(System.Net.WebUtility.HtmlEncode(c.ModifiedBy)).Append(")");
            sb.Append("</li>");
        }
        sb.Append("</ul>");

        if (brandInfo != null)
            sb.Append("<p style=\"margin-top: 24px; font-size: 11px; color: #888;\">").Append(System.Net.WebUtility.HtmlEncode(brandInfo.DisplayName)).Append(" · SharePoint Daily Digest</p>");

        sb.Append("</body></html>");
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
