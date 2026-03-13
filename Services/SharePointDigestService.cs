using System.Globalization;
using System.Text.Json;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Extensions.Logging;
using SharepointDailyDigest.Models;

namespace SharepointDailyDigest.Services;

using GraphList = Microsoft.Graph.Models.List;

public class SharePointDigestService : ISharePointDigestService
{
    private GraphServiceClient? _graph;
    private string? _configSitePath;
    private string? _configSiteUrl;
    private string? _configListName;
    private readonly object _initLock = new();
    private readonly ILogger<SharePointDigestService>? _logger;

    public SharePointDigestService(ILoggerFactory? loggerFactory = null)
    {
        _logger = loggerFactory?.CreateLogger<SharePointDigestService>();
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
            var configSiteUrl = Environment.GetEnvironmentVariable("CONFIG_SITE_URL") ?? "";
            if (string.IsNullOrEmpty(configSiteUrl))
                throw new InvalidOperationException("Add Application setting CONFIG_SITE_URL (e.g. https://tenant.sharepoint.com/sites/MySite).");
            var credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
            _graph = new GraphServiceClient(credential);
            _configListName = Environment.GetEnvironmentVariable("CONFIG_LIST_NAME") ?? "Digest Subscriptions";
            _configSiteUrl = configSiteUrl;
            _configSitePath = GetSitePathFromUrl(configSiteUrl);
        }
    }

    public async Task<IReadOnlyList<ConfigListItem>> GetConfigListItemsAsync(CancellationToken cancellationToken = default)
    {
        EnsureInitialized();
        var site = await GetSiteByPathAsync(_configSitePath!, cancellationToken).ConfigureAwait(false);
        if (site?.Id == null)
        {
            _logger?.LogWarning("Config site not found. CONFIG_SITE_URL={ConfigSiteUrl} (resolved to path: {SitePath}). Check URL and app permission Sites.Read.All.", _configSiteUrl, _configSitePath);
            return Array.Empty<ConfigListItem>();
        }

        var list = await GetListByNameAsync(site.Id, _configListName!, cancellationToken).ConfigureAwait(false);
        if (list?.Id == null)
        {
            _logger?.LogWarning("Config list '{ListName}' not found on site. Check CONFIG_LIST_NAME and that the list exists.", _configListName);
            return Array.Empty<ConfigListItem>();
        }

        var items = await _graph!.Sites[site.Id].Lists[list.Id].Items
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
            var brand = GetFieldString(fields, "Brand");
            var brandTrimmed = string.IsNullOrWhiteSpace(brand) ? null : brand.Trim();
            result.Add(new ConfigListItem(title.Trim(), email.Trim(), brandTrimmed));
        }

        var totalRows = items.Value?.Count ?? 0;
        if (result.Count == 0 && totalRows > 0)
            _logger?.LogInformation("Config list has {Total} row(s) but none with both Title and Email. Ensure columns are named 'Title' and 'Email'.", totalRows);

        return result;
    }

    public async Task<IReadOnlyList<ChangedItem>> GetRecentChangesAsync(string listOrLibraryUrl, DateTimeOffset sinceUtc, CancellationToken cancellationToken = default)
    {
        EnsureInitialized();
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
        var page = await _graph!.Sites[site.Id].Lists[list.Id].Items.GetAsync(r =>
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
            DateTimeOffset modified = default;
            if (TryGetFieldValue(fields, "Modified", out var modifiedObj) && modifiedObj != null)
                modified = ParseModifiedValue(modifiedObj) ?? default;
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

    /// <summary>Parse Modified from Graph (DateTimeOffset, string, JsonElement, Json, or ToString). Uses RoundtripKind for ISO 8601.</summary>
    private static DateTimeOffset? ParseModifiedValue(object modifiedObj)
    {
        if (modifiedObj is DateTimeOffset dto)
            return dto;
        if (modifiedObj is string s && TryParseDate(s, out var fromStr))
            return fromStr;
        if (modifiedObj is JsonElement je)
        {
            if (je.TryGetDateTimeOffset(out var fromJe)) return fromJe;
            var raw = je.GetRawText().Trim('"');
            if (TryParseDate(raw, out var fromRaw)) return fromRaw;
        }
        if (modifiedObj is Microsoft.Graph.Models.Json json && json.AdditionalData != null)
        {
            foreach (var kv in json.AdditionalData)
            {
                var str = kv.Value?.ToString()?.Trim('"');
                if (!string.IsNullOrEmpty(str) && TryParseDate(str, out var p)) return p;
            }
        }
        var fallback = modifiedObj.ToString();
        if (!string.IsNullOrEmpty(fallback) && TryParseDate(fallback.Trim('"'), out var parsed)) return parsed;
        return null;
    }

    private static bool TryParseDate(string? s, out DateTimeOffset result)
    {
        result = default;
        if (string.IsNullOrWhiteSpace(s)) return false;
        return DateTimeOffset.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out result)
            || DateTimeOffset.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal, out result);
    }

    private static string? GetFieldString(IDictionary<string, object>? fields, string name)
    {
        if (fields == null)
            return null;
        if (!TryGetFieldValue(fields, name, out var o) || o == null)
            return null;
        if (o is string s)
            return s;
        if (o is Microsoft.Graph.Models.Json j && j.AdditionalData?.TryGetValue("displayName", out var dn) == true)
            return dn?.ToString();
        return o.ToString();
    }

    private static bool TryGetFieldValue(IDictionary<string, object> fields, string name, out object? value)
    {
        value = null;
        if (fields.TryGetValue(name, out var o))
        {
            value = o;
            return true;
        }
        var key = fields.Keys.FirstOrDefault(k => string.Equals(k, name, StringComparison.OrdinalIgnoreCase));
        if (key != null && fields.TryGetValue(key, out var o2))
        {
            value = o2;
            return true;
        }
        return false;
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
            // Root site: Graph expects host + ":/" + server-relative path, e.g. host:/sites/root
            if (string.IsNullOrEmpty(path) || path == "/")
                return $"{uri.Host}:/sites/root";
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
            var sitesIndex = path.IndexOf("/sites/", StringComparison.OrdinalIgnoreCase);

            // Root site: e.g. /Shared Documents or /Lists/ListName
            if (sitesIndex < 0)
            {
                if (string.IsNullOrEmpty(path) || path == "/")
                    return ("", "");
                var sitePath = $"{uri.Host}:/sites/root";
                var listName = "";
                var listsIndex = path.IndexOf("/Lists/", StringComparison.OrdinalIgnoreCase);
                if (listsIndex >= 0)
                {
                    var listPart = path.Substring(listsIndex + 7);
                    var end = listPart.IndexOf('/');
                    listName = end > 0 ? Uri.UnescapeDataString(listPart.Substring(0, end)) : Uri.UnescapeDataString(listPart);
                }
                else
                {
                    var segments = path.TrimStart('/').Split('/');
                    listName = segments.Length >= 1 ? Uri.UnescapeDataString(segments[0]) : "";
                }
                return (sitePath, listName);
            }

            // .../sites/SiteName/Lists/ListName/... or .../sites/SiteName/ListName/...
            var afterSites = path.Substring(sitesIndex);
            var listName2 = "";
            var listsIndex2 = afterSites.IndexOf("/Lists/", StringComparison.OrdinalIgnoreCase);
            if (listsIndex2 >= 0)
            {
                var listPart = afterSites.Substring(listsIndex2 + 7);
                var end = listPart.IndexOf('/');
                listName2 = end > 0 ? Uri.UnescapeDataString(listPart.Substring(0, end)) : Uri.UnescapeDataString(listPart);
            }
            else
            {
                var segments = afterSites.Split('/');
                if (segments.Length >= 3)
                    listName2 = Uri.UnescapeDataString(segments[2]);
            }
            var sitePathEnd = path.IndexOf("/Lists/", StringComparison.OrdinalIgnoreCase);
            if (sitePathEnd < 0)
                sitePathEnd = path.Length;
            var sitePath2 = $"{uri.Host}:{path.Substring(0, sitePathEnd)}";
            return (sitePath2, listName2);
        }
        catch
        {
            return ("", "");
        }
    }

    private async Task<Site?> GetSiteByPathAsync(string hostAndPath, CancellationToken cancellationToken)
    {
        var isRoot = hostAndPath.EndsWith(":/sites/root", StringComparison.OrdinalIgnoreCase);
        var hostname = isRoot ? hostAndPath.AsSpan(0, hostAndPath.IndexOf(":/", StringComparison.Ordinal)).ToString() : null;

        // Tenant root: try "root" first (many tenants only accept this), then hostname
        var toTry = isRoot ? new[] { "root", hostname! } : new[] { hostAndPath };

        Exception? lastEx = null;
        foreach (var siteId in toTry)
        {
            if (string.IsNullOrEmpty(siteId)) continue;
            try
            {
                return await _graph!.Sites[siteId].GetAsync(requestConfig => { }, cancellationToken).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                lastEx = ex;
            }
        }

        _logger?.LogWarning(lastEx, "GetSiteByPathAsync failed for root site (tried 'root' and hostname). Ensure Microsoft Graph permission Sites.Read.All (Application) is granted with admin consent.");
        return null;
    }

    private async Task<GraphList?> GetListByNameAsync(string siteId, string listName, CancellationToken cancellationToken)
    {
        try
        {
            var lists = await _graph!.Sites[siteId].Lists.GetAsync(r => r.QueryParameters.Top = 500, cancellationToken).ConfigureAwait(false);
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
