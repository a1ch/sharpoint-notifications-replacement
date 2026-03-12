using System.Globalization;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using SharepointDailyDigest.Models;

namespace SharepointDailyDigest.Services;

public class SharePointDigestService : ISharePointDigestService
{
    private readonly GraphServiceClient _graph;
    private readonly string _configSitePath;
    private readonly string _configListName;

    public SharePointDigestService()
    {
        var tenantId = Environment.GetEnvironmentVariable("AZURE_TENANT_ID") ?? "";
        var clientId = Environment.GetEnvironmentVariable("AZURE_CLIENT_ID") ?? "";
        var clientSecret = Environment.GetEnvironmentVariable("AZURE_CLIENT_SECRET") ?? "";
        if (string.IsNullOrEmpty(tenantId) || string.IsNullOrEmpty(clientId) || string.IsNullOrEmpty(clientSecret))
            throw new InvalidOperationException("AZURE_TENANT_ID, AZURE_CLIENT_ID, and AZURE_CLIENT_SECRET must be set.");

        var credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
        _graph = new GraphServiceClient(credential);

        // Config list location: site that contains the list with Title + Email columns
        var configSiteUrl = Environment.GetEnvironmentVariable("CONFIG_SITE_URL") ?? "";
        if (string.IsNullOrEmpty(configSiteUrl))
            throw new InvalidOperationException("CONFIG_SITE_URL must be set (e.g. https://tenant.sharepoint.com/sites/MySite).");
        _configListName = Environment.GetEnvironmentVariable("CONFIG_LIST_NAME") ?? "Digest Subscriptions";
        _configSitePath = GetSitePathFromUrl(configSiteUrl);
    }

    public async Task<IReadOnlyList<ConfigListItem>> GetConfigListItemsAsync(CancellationToken cancellationToken = default)
    {
        var site = await GetSiteByPathAsync(_configSitePath, cancellationToken).ConfigureAwait(false);
        if (site?.Id == null)
            return Array.Empty<ConfigListItem>();

        var list = await GetListByNameAsync(site.Id, _configListName, cancellationToken).ConfigureAwait(false);
        if (list?.Id == null)
            return Array.Empty<ConfigListItem>();

        var items = await _graph.Sites[site.Id].Lists[list.Id].Items
            .GetAsync(r =>
            {
                r.QueryParameters.Expand = new[] { "fields" };
            }, cancellationToken).ConfigureAwait(false);

        var result = new List<ConfigListItem>();
        if (items?.Value == null)
            return result;

        foreach (var item in items.Value)
        {
            var fields = item.Fields?.AdditionalData;
            if (fields == null) continue;
            var title = GetFieldString(fields, "Title");
            var email = GetFieldString(fields, "Email");
            if (string.IsNullOrWhiteSpace(title) || string.IsNullOrWhiteSpace(email))
                continue;
            result.Add(new ConfigListItem(title.Trim(), email.Trim()));
        }
        return result;
    }

    public async Task<IReadOnlyList<ChangedItem>> GetRecentChangesAsync(string listOrLibraryUrl, DateTimeOffset sinceUtc, CancellationToken cancellationToken = default)
    {
        var (sitePath, listName) = ParseListOrLibraryUrl(listOrLibraryUrl);
        if (string.IsNullOrEmpty(sitePath) || string.IsNullOrEmpty(listName))
            return Array.Empty<ChangedItem>();

        var site = await GetSiteByPathAsync(sitePath, cancellationToken).ConfigureAwait(false);
        if (site?.Id == null)
            return Array.Empty<ChangedItem>();

        var list = await GetListByNameAsync(site.Id, listName, cancellationToken).ConfigureAwait(false);
        if (list?.Id == null)
            return Array.Empty<ChangedItem>();

        // Graph filter: use Modified (OData date format)
        var filterDate = sinceUtc.UtcDateTime.ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture);
        var filter = $"fields/Modified ge '{filterDate}'";

        var allItems = new List<ListItem>();
        var page = await _graph.Sites[site.Id].Lists[list.Id].Items.GetAsync(r =>
        {
            r.QueryParameters.Expand = new[] { "fields" };
            r.QueryParameters.Filter = filter;
            r.QueryParameters.Top = 999;
        }, cancellationToken).ConfigureAwait(false);

        if (page?.Value != null)
            allItems.AddRange(page.Value);

        var baseWebUrl = site.WebUrl ?? "";
        var results = new List<ChangedItem>();
        foreach (var item in allItems)
        {
            var fields = item.Fields?.AdditionalData;
            if (fields == null) continue;
            var title = GetFieldString(fields, "Title") ?? GetFieldString(fields, "FileLeafRef") ?? "Item";
            var modifiedObj = fields.TryGetValue("Modified", out var modVal) ? modVal : null;
            DateTimeOffset modified = default;
            if (modifiedObj is DateTimeOffset dto)
                modified = dto;
            else if (modifiedObj is string s && DateTimeOffset.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal, out var parsed))
                modified = parsed;
            var modifiedBy = GetFieldString(fields, "Editor") ?? GetFieldString(fields, "ModifiedBy");
            var webUrl = baseWebUrl;
            if (item.WebUrl != null)
                webUrl = item.WebUrl;
            else if (item.Id != null)
                webUrl = $"{baseWebUrl.TrimEnd('/')}/_layouts/15/listform.aspx?PageType=4&ListId={list.Id}&ID={item.Id}";
            results.Add(new ChangedItem(title, webUrl, modified, modifiedBy));
        }
        return results;
    }

    private static string? GetFieldString(IDictionary<string, object>? fields, string name)
    {
        if (fields == null || !fields.TryGetValue(name, out var o) || o == null)
            return null;
        if (o is string s)
            return s;
        if (o is Microsoft.Graph.Models.Json j && j.AdditionalData?.TryGetValue("displayName", out var dn) == true)
            return dn?.ToString();
        return o.ToString();
    }

    private static string GetSitePathFromUrl(string url)
    {
        try
        {
            var uri = new Uri(url);
            var path = uri.AbsolutePath.TrimEnd('/');
            // /sites/SiteName or /sites/SiteName/...
            var sitesIndex = path.IndexOf("/sites/", StringComparison.OrdinalIgnoreCase);
            if (sitesIndex >= 0)
            {
                var afterSites = path.Substring(sitesIndex);
                var nextSlash = afterSites.IndexOf('/', 7);
                var sitePath = nextSlash > 0 ? afterSites.Substring(0, nextSlash) : afterSites;
                return $"{uri.Host}:{sitePath}";
            }
            return $"{uri.Host}:{path}";
        }
        catch
        {
            return url;
        }
    }

    private static (string sitePath, string listName) ParseListOrLibraryUrl(string url)
    {
        try
        {
            var uri = new Uri(url);
            var path = uri.AbsolutePath.TrimEnd('/');
            // .../sites/SiteName/Lists/ListName/... or .../sites/SiteName/ListName/...
            var sitesIndex = path.IndexOf("/sites/", StringComparison.OrdinalIgnoreCase);
            if (sitesIndex < 0)
                return ("", "");

            var afterSites = path.Substring(sitesIndex);
            var listName = "";
            var listsIndex = afterSites.IndexOf("/Lists/", StringComparison.OrdinalIgnoreCase);
            if (listsIndex >= 0)
            {
                var listPart = afterSites.Substring(listsIndex + 7);
                var end = listPart.IndexOf('/');
                listName = end > 0 ? Uri.UnescapeDataString(listPart.Substring(0, end)) : Uri.UnescapeDataString(listPart);
            }
            else
            {
                var segments = afterSites.Split('/');
                if (segments.Length >= 3)
                    listName = Uri.UnescapeDataString(segments[2]);
            }
            var sitePathEnd = path.IndexOf("/Lists/", StringComparison.OrdinalIgnoreCase);
            if (sitePathEnd < 0)
                sitePathEnd = path.Length;
            var sitePath = $"{uri.Host}:{path.Substring(0, sitePathEnd)}";
            return (sitePath, listName);
        }
        catch
        {
            return ("", "");
        }
    }

    private async Task<Site?> GetSiteByPathAsync(string hostAndPath, CancellationToken cancellationToken)
    {
        try
        {
            return await _graph.Sites[hostAndPath].GetAsync(cancellationToken).ConfigureAwait(false);
        }
        catch
        {
            return null;
        }
    }

    private async Task<Microsoft.Graph.Models.List?> GetListByNameAsync(string siteId, string listName, CancellationToken cancellationToken)
    {
        try
        {
            var lists = await _graph.Sites[siteId].Lists.GetAsync(r => r.QueryParameters.Top = 500, cancellationToken).ConfigureAwait(false);
            var match = lists?.Value?.FirstOrDefault(l =>
                string.Equals(l.DisplayName, listName, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(l.Name, listName, StringComparison.OrdinalIgnoreCase));
            return match;
        }
        catch
        {
            return null;
        }
    }
}
