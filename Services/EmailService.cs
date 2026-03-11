using System.Text;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using SharepointDailyDigest.Models;

namespace SharepointDailyDigest.Services;

public class EmailService : IEmailService
{
    private readonly GraphServiceClient _graph;
    private readonly string _sendFromUserId;

    public EmailService()
    {
        var tenantId = Environment.GetEnvironmentVariable("AZURE_TENANT_ID") ?? "";
        var clientId = Environment.GetEnvironmentVariable("AZURE_CLIENT_ID") ?? "";
        var clientSecret = Environment.GetEnvironmentVariable("AZURE_CLIENT_SECRET") ?? "";
        var credential = new Azure.Identity.ClientSecretCredential(tenantId, clientId, clientSecret);
        _graph = new GraphServiceClient(credential);

        // User or shared mailbox to send from (object ID or UPN). Requires Mail.Send application permission.
        _sendFromUserId = Environment.GetEnvironmentVariable("SEND_FROM_USER_ID") ?? "";
        if (string.IsNullOrEmpty(_sendFromUserId))
            throw new InvalidOperationException("SEND_FROM_USER_ID must be set (object ID or user principal name of the mailbox to send from).");
    }

    public async Task SendDigestAsync(string toEmail, string listOrLibraryName, IReadOnlyList<ChangedItem> changes, CancellationToken cancellationToken = default)
    {
        if (changes.Count == 0)
            return;

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

        await _graph.Users[_sendFromUserId].SendMail.PostAsync(requestBody, cancellationToken: cancellationToken).ConfigureAwait(false);
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
