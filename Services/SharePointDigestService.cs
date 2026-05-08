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
                throw new InvalidOperationException("Add Application setting CONFIG_SITE_URL (e.g. https://tenant.sharepoint.com/itsp).");
            var credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
            _graph = new GraphServiceClient(credential);
            _configListName = Environment.GetEnvironmentVariable("CONFIG_LIST_NAME") ?? "Digest Subscriptions";
            _configSiteUrl = configSiteUrl;
        }
    }

    public async Task<IReadOnlyList<ConfigListItem>> GetConfigListItemsAsync(CancellationToken cancellationToken = default)
    {
        EnsureInitialized();

        var sitePathsToTry = BuildSitePathCandidates(_configSiteUrl!);
        string? resolvedSiteId = null;
        GraphList? list = null;

        foreach (var sp in sitePathsToTry)
        {
            if (string.IsNullOrEmpty(sp)) continue;
            var candidate = await GetSiteByPathAsync(sp, cancellationToken).ConfigureAwait(false);
            if (candidate?.Id == null) continue;
            var candidateList = await GetListByNameAsync(candidate.Id, _configListName!, cancellationToken).ConfigureAwait(false);
            if (candidateList?.Id == null) continue;
            resolvedSiteId = candidate.Id;
            list = candidateList;
            _logger?.LogInformation("Config list '{ListName}' resolved on site '{SitePath}'", _configListName, sp);
            break;
        }

        if (list?.Id == null || resolvedSiteId == null)
        {
            _logger?.LogWarning("Config list '{ListName}' not found. CONFIG_SITE_URL={Url}. Tried: {Tried}",
                _configListName, _configSiteUrl, string.Join(", ", sitePathsToTry));
            return Array.Empty<ConfigListItem>();
        }

        var items = await _graph!.Sites[resolvedSiteId].Lists[list.Id].Items
            .GetAsync(r => { r.QueryParameters.Expand = new[] { "fields" }; }, cancellationToken)
            .ConfigureAwait(false);

        var result = new List<ConfigListItem>();
        if (items?.Value == null) return result;

        foreach (var item in items.Value)
        {
            var fields = item.Fields?.AdditionalData;
            if (fields == null) continue;
            var title = GetFieldString(fields, "Title");
            var email = GetFieldString(fields, "Email");
            if (string.IsNullOrWhiteSpace(title) || string.IsNullOrWhiteSpace(email)) continue;
            var brand = GetFieldString(fields, "Brand");
            result.Add(new ConfigListItem(title.Trim(), email.Trim(), string.IsNullOrWhiteSpace(brand) ? null : brand.Trim()));
        }

        var totalRows = items.Value?.Count ?? 0;
        if (result.Count == 0 && totalRows > 0)
            _logger?.LogInformation("Config list has {Total} row(s) but none with both Title and Email. Ensure columns are named 'Title' and 'Email'.", totalRows);

        return result;
    }

    public async Task<IReadOnlyList<ChangedItem>> GetRecentChangesAsync(string listOrLibraryUrl, DateTimeOffset sinceUtc, CancellationToken cancellationToken = default)
    {
        EnsureInitialized();
        var (sitePath, subsitePath, listName) = ParseListOrLibraryUrl(listOrLibraryUrl);
        if (string.IsNullOrEmpty(sitePath) || string.IsNullOrEmpty(listName))
        {
            _logger?.LogWarning("Could not parse site path or list name from URL: {Url}", listOrLibraryUrl);
            return Array.Empty<ChangedItem>();
        }

        _logger?.LogInformation("Parsed URL '{Url}' -> site='{SitePath}', subsite='{SubsitePath}', list='{ListName}'",
            listOrLibraryUrl, sitePath, subsitePath ?? "(none)", listName);

        // Reconstruct a full URL from the parsed components so BuildSitePathCandidates
        // can generate correctly formatted Graph site paths (host:/path) for all cases.
        var reconstructedUrl = subsitePath != null
            ? $"https://{sitePath.Replace(":/", "/")}/{subsitePath}"
            : $"https://{sitePath.Replace(":/", "/")}";
        var sitePathsToTry = BuildSitePathCandidates(reconstructedUrl);

        string? resolvedSiteId = null;
        GraphList? list = null;
        Site? resolvedSite = null;

        foreach (var sp in sitePathsToTry)
        {
            var candidate = await GetSiteByPathAsync(sp, cancellationToken).ConfigureAwait(false);
            if (candidate?.Id == null) continue;
            var candidateList = await GetListByNameAsync(candidate.Id, listName, cancellationToken).ConfigureAwait(false);
            if (candidateList?.Id == null) continue;
            resolvedSite = candidate;
            resolvedSiteId = candidate.Id;
            list = candidateList;
            _logger?.LogInformation("Resolved list '{ListName}' on site '{SitePath}' (id={SiteId})", listName, sp, resolvedSiteId);
            break;
        }

        if (list?.Id == null || resolvedSiteId == null)
        {
            _logger?.LogWarning("List '{ListName}' not found on any candidate site (from URL: {Url}). Tried: {Tried}",
                listName, listOrLibraryUrl, string.Join(", ", sitePathsToTry));
            return Array.Empty<ChangedItem>();
        }

        var filterDate = sinceUtc.UtcDateTime.ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture);
        var filter = $"fields/Modified ge '{filterDate}'";

        var allItems = new List<ListItem>();
        var page = await _graph!.Sites[resolvedSiteId].Lists[list.Id].Items.GetAsync(r =>
        {
            r.QueryParameters.Expand = new[] { "fields" };
            r.QueryParameters.Filter = filter;
            r.QueryParameters.Top = 999;
        }, cancellationToken).ConfigureAwait(false);

        if (page?.Value != null) allItems.AddRange(page.Value);

        var baseWebUrl = resolvedSite?.WebUrl ?? "";

        var results = new List<ChangedItem>();
        foreach (var item in allItems)
        {
            var fields = item.Fields?.AdditionalData;
            if (fields == null) continue;
            var title = GetFieldString(fields, "Title") ?? GetFieldString(fields, "FileLeafRef") ?? "Item";
            DateTimeOffset modified = default;
            if (TryGetFieldValue(fields, "Modified", out var modifiedObj) && modifiedObj != null)
                modified = ParseModifiedValue(modifiedObj) ?? default;
            var modifiedBy = GetIdentitySetDisplayName(item.LastModifiedBy)
                ?? GetModifiedByDisplayName(fields)
                ?? GetLastModifiedByFromSerializedListItem(item);
            var webUrl = baseWebUrl;
            if (item.WebUrl != null)
                webUrl = item.WebUrl;
            else if (item.Id != null)
                webUrl = $"{baseWebUrl.TrimEnd('/')}/_layouts/15/listform.aspx?PageType=4&ListId={list.Id}&ID={item.Id}";
            results.Add(new ChangedItem(title, webUrl, modified, modifiedBy));
        }
        return results;
    }

    /// <summary>
    /// Given a site URL, returns ordered Graph site path candidates to try, most specific first.
    /// Root collection subsite: host:/itsp, host:/itsp/itst, host
    /// Named site subsite:      host:/sites/OFForms/SupplierForms, host:/sites/OFForms
    /// </summary>
    private static List<string> BuildSitePathCandidates(string siteUrl)
    {
        try
        {
            var uri = new Uri(siteUrl);
            var path = uri.AbsolutePath.TrimEnd('/');
            var candidates = new List<string>();
            var sitesIndex = path.IndexOf("/sites/", StringComparison.OrdinalIgnoreCase);

            if (sitesIndex < 0)
            {
                // Root collection subsite — Graph expects host:/path format
                var segments = path.TrimStart('/').Split('/', StringSplitOptions.RemoveEmptyEntries);
                for (var i = segments.Length; i >= 1; i--)
                    candidates.Add($"{uri.Host}:/{string.Join("/", segments[..i])}");
                candidates.Add(uri.Host); // root fallback
            }
            else
            {
                // Named site e.g. /sites/OFForms or /sites/OFForms/SupplierForms
                var afterSites = path[sitesIndex..];
                var parts = afterSites.TrimStart('/').Split('/', StringSplitOptions.RemoveEmptyEntries);
                for (var i = parts.Length; i >= 2; i--)
                    candidates.Add($"{uri.Host}:/{string.Join("/", parts[..i])}");
            }

            return candidates;
        }
        catch { return new List<string>(); }
    }

    /// <summary>
    /// Parses a SharePoint list/library URL into (sitePath, subsitePath, listName).
    /// sitePath: host:/sites/SiteName for named sites, or hostname for root collection.
    /// subsitePath: segments between the site and the list/library.
    /// </summary>
    private static (string sitePath, string? subsitePath, string listName) ParseListOrLibraryUrl(string url)
    {
        try
        {
            var uri = new Uri(url);
            var path = uri.AbsolutePath.TrimEnd('/');
            path = StripViewSuffix(path);
            var sitesIndex = path.IndexOf("/sites/", StringComparison.OrdinalIgnoreCase);

            if (sitesIndex < 0)
            {
                if (string.IsNullOrEmpty(path) || path == "/") return ("", null, "");
                var rootSitePath = uri.Host;
                var listsIdx = path.IndexOf("/Lists/", StringComparison.OrdinalIgnoreCase);
                if (listsIdx >= 0)
                {
                    var listPart = path[(listsIdx + 7)..];
                    var end = listPart.IndexOf('/');
                    var listName = end > 0 ? Uri.UnescapeDataString(listPart[..end]) : Uri.UnescapeDataString(listPart);
                    var beforeLists = path[..listsIdx].TrimStart('/');
                    return (rootSitePath, string.IsNullOrEmpty(beforeLists) ? null : beforeLists, listName);
                }
                var segments = path.TrimStart('/').Split('/');
                if (segments.Length == 1) return (rootSitePath, null, Uri.UnescapeDataString(segments[0]));
                return (rootSitePath, string.Join("/", segments[..^1]), Uri.UnescapeDataString(segments[^1]));
            }

            var afterSites = path[sitesIndex..];
            var siteNameEnd = afterSites.IndexOf('/', 7);
            var siteRelativePath = siteNameEnd > 0 ? afterSites[..siteNameEnd] : afterSites;
            var sitePath2 = $"{uri.Host}:{siteRelativePath}";
            if (siteNameEnd < 0) return (sitePath2, null, "");

            var afterSiteName = afterSites[siteNameEnd..];
            var listsIndex2 = afterSiteName.IndexOf("/Lists/", StringComparison.OrdinalIgnoreCase);
            if (listsIndex2 >= 0)
            {
                var listPart = afterSiteName[(listsIndex2 + 7)..];
                var end = listPart.IndexOf('/');
                var listName = end > 0 ? Uri.UnescapeDataString(listPart[..end]) : Uri.UnescapeDataString(listPart);
                var beforeLists = afterSiteName[..listsIndex2].Trim('/');
                return (sitePath2, string.IsNullOrEmpty(beforeLists) ? null : beforeLists, listName);
            }

            var remainingSegments = afterSiteName.Trim('/').Split('/', StringSplitOptions.RemoveEmptyEntries);
            if (remainingSegments.Length == 0) return (sitePath2, null, "");
            if (remainingSegments.Length == 1) return (sitePath2, null, Uri.UnescapeDataString(remainingSegments[0]));

            var libraryName = Uri.UnescapeDataString(remainingSegments[^1]);
            if (libraryName.StartsWith("_", StringComparison.Ordinal) ||
                libraryName.Equals("Forms", StringComparison.OrdinalIgnoreCase))
                return (sitePath2, null, "");

            return (sitePath2, string.Join("/", remainingSegments[..^1]), libraryName);
        }
        catch { return ("", null, ""); }
    }

    private static string StripViewSuffix(string path)
    {
        if (path.EndsWith(".aspx", StringComparison.OrdinalIgnoreCase))
        {
            var lastSlash = path.LastIndexOf('/');
            if (lastSlash > 0) path = path[..lastSlash];
        }
        if (path.EndsWith("/Forms", StringComparison.OrdinalIgnoreCase))
            path = path[..^6];
        return path.TrimEnd('/');
    }

    private static DateTimeOffset? ParseModifiedValue(object modifiedObj)
    {
        if (modifiedObj is DateTimeOffset dto) return dto;
        if (modifiedObj is string s && TryParseDate(s, out var fromStr)) return fromStr;
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
        if (fields == null) return null;
        if (!TryGetFieldValue(fields, name, out var o) || o == null) return null;
        if (o is string s) return s;
        if (o is Microsoft.Graph.Models.Json j && j.AdditionalData?.TryGetValue("displayName", out var dn) == true)
            return dn?.ToString();
        return o.ToString();
    }

    private static string? GetLastModifiedByFromSerializedListItem(ListItem item)
    {
        try
        {
            var json = JsonSerializer.Serialize(item);
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;
            if (root.ValueKind != JsonValueKind.Object) return null;
            foreach (var prop in root.EnumerateObject())
            {
                if (!prop.Name.Equals("lastModifiedBy", StringComparison.OrdinalIgnoreCase)) continue;
                var lmb = prop.Value;
                if (lmb.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined) return null;
                if (lmb.TryGetProperty("user", out var user) && user.ValueKind == JsonValueKind.Object)
                    return ParsePersonFromJsonElement(user);
                return ParsePersonFromJsonElement(lmb);
            }
        }
        catch { }
        return null;
    }

    private static string? GetIdentitySetDisplayName(IdentitySet? identitySet)
    {
        if (identitySet == null) return null;
        foreach (var identity in new[] { identitySet.User, identitySet.Application, identitySet.Device })
        {
            if (identity == null) continue;
            if (!string.IsNullOrWhiteSpace(identity.DisplayName)) return identity.DisplayName.Trim();
            if (identity.AdditionalData != null)
                foreach (var want in new[] { "email", "mail", "userPrincipalName" })
                    foreach (var kv in identity.AdditionalData)
                    {
                        if (!kv.Key.Equals(want, StringComparison.OrdinalIgnoreCase)) continue;
                        var s = GraphValueToTrimmedString(kv.Value);
                        if (!string.IsNullOrWhiteSpace(s)) return s;
                    }
            if (!string.IsNullOrWhiteSpace(identity.Id) && identity.Id.Contains('@', StringComparison.Ordinal))
                return identity.Id.Trim();
        }
        return null;
    }

    private static string? GetModifiedByDisplayName(IDictionary<string, object> fields)
    {
        foreach (var fieldName in new[] { "Editor", "ModifiedBy", "Modified_x0020_By" })
        {
            if (!TryGetFieldValue(fields, fieldName, out var o) || o == null) continue;
            var name = ParsePersonOrClaimsField(o);
            if (!string.IsNullOrWhiteSpace(name)) return name;
        }
        foreach (var kv in fields)
        {
            if (kv.Value == null) continue;
            var k = kv.Key;
            if (k.Contains("LookupId", StringComparison.OrdinalIgnoreCase) || k.Contains("LookupValueId", StringComparison.OrdinalIgnoreCase)) continue;
            if (!k.Contains("ditor", StringComparison.OrdinalIgnoreCase) && !k.Contains("ModifiedBy", StringComparison.OrdinalIgnoreCase)) continue;
            var name = ParsePersonOrClaimsField(kv.Value);
            if (!string.IsNullOrWhiteSpace(name)) return name;
        }
        return null;
    }

    private static string? ParsePersonOrClaimsField(object? o)
    {
        if (o == null) return null;
        if (o is string s) return NormalizeClaimsOrString(s);
        if (o is JsonElement je) return ParsePersonFromJsonElement(je);
        if (o is Microsoft.Graph.Models.Json j && j.AdditionalData != null)
            foreach (var want in new[] { "displayName", "LookupValue", "Email", "Title", "Name", "preferredName", "mail" })
                foreach (var kv in j.AdditionalData)
                {
                    if (!kv.Key.Equals(want, StringComparison.OrdinalIgnoreCase)) continue;
                    var str = GraphValueToTrimmedString(kv.Value);
                    if (!string.IsNullOrWhiteSpace(str)) return str;
                }
        if (o is IDictionary<string, object> nested)
            foreach (var want in new[] { "displayName", "LookupValue", "Email", "Title", "Name", "preferredName", "mail" })
            {
                if (!TryGetFieldValue(nested, want, out var inner) || inner == null) continue;
                var str = ParsePersonOrClaimsField(inner);
                if (!string.IsNullOrWhiteSpace(str)) return str;
            }
        try
        {
            var json = JsonSerializer.Serialize(o);
            if (json.Length > 2 && json[0] == '{')
            {
                using var doc = JsonDocument.Parse(json);
                var fromJson = ParsePersonFromJsonElement(doc.RootElement);
                if (!string.IsNullOrWhiteSpace(fromJson)) return fromJson;
            }
        }
        catch { }
        return null;
    }

    private static string? ParsePersonFromJsonElement(JsonElement je)
    {
        if (je.ValueKind == JsonValueKind.String) return NormalizeClaimsOrString(je.GetString() ?? "");
        if (je.ValueKind != JsonValueKind.Object) return null;
        foreach (var want in new[] { "displayName", "LookupValue", "Email", "Title", "Name", "preferredName", "mail" })
            foreach (var prop in je.EnumerateObject())
            {
                if (!prop.Name.Equals(want, StringComparison.OrdinalIgnoreCase)) continue;
                if (prop.Value.ValueKind == JsonValueKind.String)
                {
                    var v = prop.Value.GetString();
                    if (!string.IsNullOrWhiteSpace(v)) return v.Trim();
                }
            }
        foreach (var prop in je.EnumerateObject())
        {
            if (prop.Value.ValueKind != JsonValueKind.String) continue;
            var v = prop.Value.GetString()?.Trim();
            if (!string.IsNullOrWhiteSpace(v) && v.Contains('@', StringComparison.Ordinal)) return v;
        }
        return null;
    }

    private static string? GraphValueToTrimmedString(object? v)
    {
        if (v == null) return null;
        if (v is string ss) return string.IsNullOrWhiteSpace(ss) ? null : ss.Trim();
        if (v is JsonElement jee)
        {
            if (jee.ValueKind == JsonValueKind.String) { var t = jee.GetString()?.Trim(); return string.IsNullOrWhiteSpace(t) ? null : t; }
            if (jee.ValueKind == JsonValueKind.Object) return ParsePersonFromJsonElement(jee);
        }
        return null;
    }

    private static string? NormalizeClaimsOrString(string s)
    {
        s = s.Trim();
        if (s.Length == 0) return null;
        foreach (var marker in new[] { "|membership|", "|windows|", "|claims|" })
        {
            var idx = s.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
            if (idx < 0) continue;
            var tail = s[(idx + marker.Length)..].Trim();
            if (tail.Length > 0) return tail;
        }
        return s;
    }

    private static bool TryGetFieldValue(IDictionary<string, object> fields, string name, out object? value)
    {
        value = null;
        if (fields.TryGetValue(name, out var o)) { value = o; return true; }
        var key = fields.Keys.FirstOrDefault(k => string.Equals(k, name, StringComparison.OrdinalIgnoreCase));
        if (key != null && fields.TryGetValue(key, out var o2)) { value = o2; return true; }
        return false;
    }

    private async Task<Site?> GetSiteByPathAsync(string hostAndPath, CancellationToken cancellationToken)
    {
        try { return await _graph!.Sites[hostAndPath].GetAsync(r => { }, cancellationToken).ConfigureAwait(false); }
        catch { return null; }
    }

    private async Task<GraphList?> GetListByNameAsync(string siteId, string listName, CancellationToken cancellationToken)
    {
        try
        {
            var page = await _graph!.Sites[siteId].Lists
                .GetAsync(r => r.QueryParameters.Top = 1000, cancellationToken)
                .ConfigureAwait(false);

            while (page?.Value != null)
            {
                var match = page.Value.FirstOrDefault(l =>
                    string.Equals(l.DisplayName, listName, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(l.Name, listName, StringComparison.OrdinalIgnoreCase));
                if (match != null) return match;
                if (page.OdataNextLink == null) break;
                page = await _graph.Sites[siteId].Lists
                    .WithUrl(page.OdataNextLink)
                    .GetAsync(cancellationToken: cancellationToken)
                    .ConfigureAwait(false);
            }
            return null;
        }
        catch { return null; }
    }
}
