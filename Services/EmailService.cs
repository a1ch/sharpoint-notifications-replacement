using System.Text;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Users.Item.SendMail;
using SharepointDailyDigest.Models;

namespace SharepointDailyDigest.Services;

public class EmailService : IEmailService
{
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

    public async Task SendDigestAsync(string toEmail, string listOrLibraryName, IReadOnlyList<ChangedItem> changes, CancellationToken cancellationToken = default)
    {
        if (changes.Count == 0)
            return;
        EnsureInitialized();

        var subject = $"SharePoint digest: {listOrLibraryName} – {changes.Count} new or updated item(s) in the last 24 hours";
        var body = BuildDigestHtml(listOrLibraryName, changes);

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

    private static string BuildDigestHtml(string listOrLibraryName, IReadOnlyList<ChangedItem> changes)
    {
        var sb = new StringBuilder();
        sb.Append("<html><body>");
        sb.Append($"<p>Summary of new or changed items in the past 24 hours for <strong>").Append(System.Net.WebUtility.HtmlEncode(listOrLibraryName)).Append("</strong>.</p>");
        sb.Append("<ul>");
        foreach (var c in changes)
        {
            sb.Append("<li>");
            sb.Append("<a href=\"").Append(System.Net.WebUtility.HtmlEncode(c.WebUrl)).Append("\">").Append(System.Net.WebUtility.HtmlEncode(c.Title)).Append("</a>");
            sb.Append(" – ").Append(c.Modified.ToString("g"));
            if (!string.IsNullOrEmpty(c.ModifiedBy))
                sb.Append(" (by ").Append(System.Net.WebUtility.HtmlEncode(c.ModifiedBy)).Append(")");
            sb.Append("</li>");
        }
        sb.Append("</ul></body></html>");
        return sb.ToString();
    }
}
