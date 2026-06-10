using System.Globalization;
using System.Text;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Users.Item.SendMail;
using SharepointDailyDigest.Models;

namespace SharepointDailyDigest.Services;

public class EmailService : IEmailService
{
    // Unified Stream-Flo Group identity. The per-row Brand value is intentionally ignored:
    // every digest uses one template that represents all three companies together.
    private const string GroupName = "Stream-Flo Group of Companies";
    private const string GroupAccent = "#003366";   // navy — used for links, button, borders, chips

    // Heavy wordmark/heading font stack (reads as an intentional logo across clients).
    private const string HeadFont = "'Arial Black','Segoe UI',Arial,sans-serif";

    /// <summary>The three group companies (display name, accent color, legal name). Used for the header ribbon.</summary>
    private static readonly (string Name, string Color, string Legal)[] GroupBrands =
    {
        ("Stream-Flo", "#003366", "Stream-Flo USA LLC"),
        ("Master Flo", "#0066b3", "Master Flo Valve USA Inc."),
        ("Dycor",      "#0d7a7a", "Dycor Technologies"),
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

    public async Task SendDigestAsync(string toEmail, string listOrLibraryName, IReadOnlyList<ChangedItem> changes, string? brand = null, string? siteName = null, string? listOrLibraryUrl = null, CancellationToken cancellationToken = default)
    {
        if (changes.Count == 0)
            return;
        EnsureInitialized();

        // brand is intentionally unused: the unified Stream-Flo Group template is always used.
        var hasSiteName = !string.IsNullOrWhiteSpace(siteName);
        var subject = hasSiteName
            ? $"SharePoint digest: {siteName} – {listOrLibraryName} – {changes.Count} new or updated item(s) in the last 24 hours"
            : $"SharePoint digest: {listOrLibraryName} – {changes.Count} new or updated item(s) in the last 24 hours";
        var body = BuildDigestHtml(listOrLibraryName, changes, siteName, listOrLibraryUrl);

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

    private static string BuildDigestHtml(string listOrLibraryName, IReadOnlyList<ChangedItem> changes, string? siteName = null, string? listOrLibraryUrl = null)
    {
        static string Enc(string? s) => System.Net.WebUtility.HtmlEncode(s ?? "");

        var accent = GroupAccent;
        var accentDark = Darken(accent, 0.72);
        var accentSoft = Tint(accent, 0.90);   // very light navy tint for the count chip

        var hasSiteName = !string.IsNullOrWhiteSpace(siteName);
        var hasLibraryUrl = !string.IsNullOrWhiteSpace(listOrLibraryUrl);

        var latestChange = changes
            .OrderByDescending(c => c.Modified == default ? DateTimeOffset.MinValue : c.Modified)
            .FirstOrDefault();
        var latestChangedBy = !string.IsNullOrWhiteSpace(latestChange?.ModifiedBy) ? latestChange!.ModifiedBy : "Unknown user";
        var latestChangedAt = latestChange != null ? FormatModifiedDate(latestChange.Modified) : "Unknown date";
        var itemWord = changes.Count == 1 ? "item" : "items";

        var sb = new StringBuilder(8192);

        sb.Append("<!DOCTYPE html><html lang=\"en\"><head><meta charset=\"utf-8\"/>");
        sb.Append("<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\"/>");
        sb.Append("<meta name=\"color-scheme\" content=\"light only\"/>");
        sb.Append("<title>SharePoint Daily Digest</title></head>");
        sb.Append("<body style=\"margin:0; padding:0; width:100%; background:#eef2f7; font-family:'Segoe UI',Arial,sans-serif; font-size:15px; line-height:1.5; color:#1e293b; -webkit-font-smoothing:antialiased;\">");

        // Hidden preheader (inbox preview text)
        sb.Append("<div style=\"display:none; max-height:0; overflow:hidden; opacity:0; mso-hide:all;\">")
          .Append(changes.Count).Append(' ').Append("new or updated ").Append(itemWord).Append(" in ").Append(Enc(listOrLibraryName))
          .Append(" &mdash; last 24 hours</div>");

        sb.Append("<table role=\"presentation\" width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" style=\"background:#eef2f7; padding:32px 16px;\"><tr><td align=\"center\">");
        sb.Append("<table role=\"presentation\" width=\"600\" cellpadding=\"0\" cellspacing=\"0\" style=\"width:600px; max-width:600px; background:#ffffff; border-radius:14px; border:1px solid #e3e8ef; box-shadow:0 8px 24px rgba(15,23,42,0.08); overflow:hidden;\">");

        // ---- Tri-color ribbon (one segment per company) ----
        sb.Append("<tr><td style=\"padding:0;\"><table role=\"presentation\" width=\"100%\" cellpadding=\"0\" cellspacing=\"0\"><tr>");
        foreach (var g in GroupBrands)
            sb.Append("<td style=\"height:6px; line-height:6px; font-size:0; background:").Append(g.Color).Append(";\">&nbsp;</td>");
        sb.Append("</tr></table></td></tr>");

        // ---- Group header ----
        sb.Append("<tr><td style=\"background-color:").Append(accent)
          .Append("; background-image:linear-gradient(135deg,").Append(accent).Append(" 0%,").Append(accentDark).Append(" 100%); padding:32px 36px 28px;\">");
        sb.Append("<div style=\"font-family:").Append(HeadFont).Append("; font-size:24px; font-weight:800; letter-spacing:0.2px; color:#ffffff; line-height:1.15;\">").Append(Enc(GroupName)).Append("</div>");
        sb.Append("<div style=\"width:46px; height:3px; background:rgba(255,255,255,0.55); border-radius:2px; margin:14px 0 0;\"></div>");
        sb.Append("</td></tr>");

        // ---- Sub-bar ----
        sb.Append("<tr><td style=\"background:").Append(accentDark).Append("; padding:9px 36px;\">");
        sb.Append("<span style=\"font-family:").Append(HeadFont).Append("; color:rgba(255,255,255,0.92); font-size:11px; font-weight:700; letter-spacing:1.4px; text-transform:uppercase;\">SharePoint Daily Digest</span>");
        sb.Append("</td></tr>");

        // ---- Body ----
        sb.Append("<tr><td style=\"padding:30px 36px 8px;\">");

        if (hasSiteName)
            sb.Append("<div style=\"font-family:").Append(HeadFont).Append("; font-size:11px; font-weight:700; letter-spacing:1.3px; text-transform:uppercase; color:#94a3b8; margin:0 0 6px;\">").Append(Enc(siteName)).Append("</div>");

        sb.Append("<div style=\"font-family:").Append(HeadFont).Append("; font-size:21px; font-weight:800; color:#0f172a; margin:0 0 12px; line-height:1.25;\">").Append(Enc(listOrLibraryName)).Append("</div>");

        // Count pill
        sb.Append("<span style=\"display:inline-block; background:").Append(accentSoft).Append("; color:").Append(accent)
          .Append("; font-size:12px; font-weight:700; letter-spacing:0.3px; padding:6px 14px; border-radius:999px;\">")
          .Append(changes.Count).Append(' ').Append(itemWord).Append(" updated &middot; last 24 hours</span>");

        // Summary card
        sb.Append("<table role=\"presentation\" width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" style=\"margin:18px 0 0; background:#f8fafc; border:1px solid #e6ebf2; border-radius:10px;\"><tr><td style=\"padding:16px 18px;\">");
        sb.Append("<table role=\"presentation\" width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" style=\"font-size:13px;\">");
        AppendSummaryRow(sb, "Items changed", changes.Count.ToString(CultureInfo.InvariantCulture), false);
        AppendSummaryRow(sb, "Latest update", Enc(latestChangedAt), false);
        AppendSummaryRow(sb, "Latest changed by", Enc(latestChangedBy), true);
        sb.Append("</table>");
        sb.Append("</td></tr></table>");

        // Open button
        if (hasLibraryUrl)
        {
            sb.Append("<table role=\"presentation\" cellpadding=\"0\" cellspacing=\"0\" style=\"margin:20px 0 4px;\"><tr><td style=\"background:").Append(accent).Append("; border-radius:8px;\">");
            sb.Append("<a href=\"").Append(Enc(listOrLibraryUrl)).Append("\" style=\"display:inline-block; padding:12px 24px; font-size:14px; font-weight:700; color:#ffffff; text-decoration:none; letter-spacing:0.2px;\">Open ").Append(Enc(listOrLibraryName)).Append(" &rarr;</a>");
            sb.Append("</td></tr></table>");
        }

        sb.Append("</td></tr>");

        // ---- Section heading ----
        sb.Append("<tr><td style=\"padding:20px 36px 4px;\">");
        sb.Append("<div style=\"font-family:").Append(HeadFont).Append("; font-size:11px; font-weight:700; letter-spacing:1.3px; text-transform:uppercase; color:#94a3b8; border-top:1px solid #eceff4; padding-top:18px;\">What changed</div>");
        sb.Append("</td></tr>");

        // ---- Item cards ----
        sb.Append("<tr><td style=\"padding:8px 36px 8px;\">");
        var idx = 0;
        foreach (var c in changes)
        {
            idx++;
            sb.Append("<table role=\"presentation\" width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" style=\"margin:0 0 10px; background:#ffffff; border:1px solid #eceff4; border-left:4px solid ").Append(accent).Append("; border-radius:10px;\"><tr><td style=\"padding:14px 16px;\">");
            sb.Append("<div style=\"font-family:").Append(HeadFont).Append("; font-size:11px; font-weight:700; color:").Append(accent).Append("; letter-spacing:0.5px; margin:0 0 4px;\">#").Append(idx).Append("</div>");
            sb.Append("<a href=\"").Append(Enc(c.WebUrl)).Append("\" style=\"font-size:15px; font-weight:600; color:#0f172a; text-decoration:none;\">").Append(Enc(c.Title)).Append("</a>");
            sb.Append("<div style=\"font-size:12.5px; color:#64748b; margin-top:6px;\">");
            sb.Append("<span style=\"color:#475569; font-weight:600;\">Updated</span> ").Append(Enc(FormatModifiedDate(c.Modified)));
            sb.Append(" &nbsp;&middot;&nbsp; <span style=\"color:#475569; font-weight:600;\">By</span> ").Append(Enc(string.IsNullOrWhiteSpace(c.ModifiedBy) ? "Unknown user" : c.ModifiedBy));
            sb.Append("</div>");
            sb.Append("</td></tr></table>");
        }
        sb.Append("</td></tr>");

        // ---- Footer ----
        sb.Append("<tr><td style=\"padding:22px 36px 26px; background:#fafbfc; border-top:1px solid #eceff4;\">");
        sb.Append("<div style=\"font-size:11px; color:#94a3b8; line-height:1.6;\">This is an automated message from the Stream-Flo Group of Companies SharePoint notification system. Please do not reply to this email.</div>");
        sb.Append("</td></tr>");

        sb.Append("</table></td></tr></table></body></html>");
        return sb.ToString();
    }

    private static void AppendSummaryRow(StringBuilder sb, string label, string encodedValue, bool last)
    {
        var pad = last ? "0" : "0 0 8px";
        sb.Append("<tr><td style=\"padding:").Append(pad).Append("; width:140px; color:#64748b; font-weight:600; vertical-align:top;\">").Append(System.Net.WebUtility.HtmlEncode(label)).Append("</td>");
        sb.Append("<td style=\"padding:").Append(pad).Append("; color:#1e293b; vertical-align:top;\">").Append(encodedValue).Append("</td></tr>");
    }

    /// <summary>Darken a hex color (for the gradient base and sub-bar under the group header).</summary>
    private static string Darken(string hex, double factor = 0.7)
    {
        try
        {
            hex = hex.TrimStart('#');
            if (hex.Length != 6) return "#002a57";
            int r = (int)(Convert.ToInt32(hex.Substring(0, 2), 16) * factor);
            int g = (int)(Convert.ToInt32(hex.Substring(2, 2), 16) * factor);
            int bl = (int)(Convert.ToInt32(hex.Substring(4, 2), 16) * factor);
            return $"#{r:X2}{g:X2}{bl:X2}";
        }
        catch { return "#002a57"; }
    }

    /// <summary>Lighten a hex color toward white by <paramref name="amount"/> (0..1) for soft brand-tinted chips/panels.</summary>
    private static string Tint(string hex, double amount = 0.9)
    {
        try
        {
            hex = hex.TrimStart('#');
            if (hex.Length != 6) return "#eef2f7";
            int Mix(int c) => (int)Math.Round(c + (255 - c) * amount);
            int r = Mix(Convert.ToInt32(hex.Substring(0, 2), 16));
            int g = Mix(Convert.ToInt32(hex.Substring(2, 2), 16));
            int bl = Mix(Convert.ToInt32(hex.Substring(4, 2), 16));
            return $"#{r:X2}{g:X2}{bl:X2}";
        }
        catch { return "#eef2f7"; }
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
