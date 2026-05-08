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
        var (sitePath, subsitePath, listName) = ParseListOrLibraryUrl(listOrLibraryUrl);
        if (string.IsNullOrEmpty(sitePath) || string.IsNullOrEmpty(listName))
        {
            _logger?.LogWarning("Could not parse site path or list name from URL: {Url}", listOrLibraryUrl);
            return Array.Empty<ChangedItem>();
        }

        _logger?.LogInformation("Parsed URL '{Url}' -> site='{SitePath}', subsite='{SubsitePath}', list='{ListName}'", listOrLibraryUrl, sitePath, subsitePath ?? "(none)", listName);

        var site = await GetSiteByPathAsync(sitePath, cancellationToken).ConfigureAwait(false);
        if (site?.Id == null)
        {
            _logger?.LogWarning("Site not found for path '{SitePath}' (from URL: {Url})", sitePath, listOrLibraryUrl);
            return Array.Empty<ChangedItem>();
        }

        // Try to find the list — first on the parent site, then on the subsite if one exists
        string? resolvedSiteId = null;
        GraphList? list = null;

        // If a subsite path is present, try the subsite first
        if (!string.IsNullOrEmpty(subsitePath))
        {
            var subsite = await GetSubsiteByPathAsync(site.Id, subsitePath, cancellationToken).ConfigureAwait(false);
            if (subsite?.Id != null)
            {
                list = await GetListByNameAsync(subsite.Id, listName, cancellationToken).ConfigureAwait(false);
                if (list?.Id != null)
                {
                    resolvedSiteId = subsite.Id;
                    _logger?.LogInformation("Found list '{ListName}' on subsite '{SubsiteId}'", listName, subsite.Id);
                }
            }
        }

        // Fall back to parent site
        if (list?.Id == null)
        {
            list = await GetListByNameAsync(site.Id, listName, cancellationToken).ConfigureAwait(false);
            if (list?.Id != null)
                resolvedSiteId = site.Id;
        }

        if (list?.Id == null || resolvedSiteId == null)
        {
            _logger?.LogWarning("List '{ListName}' not found on site or subsite (from URL: {Url})", listName, listOrLibraryUrl);
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

        if (page?.Value != null)
            allItems.AddRange(page.Value);

        var baseSite = await _graph!.Sites[resolvedSiteId].GetAsync(cancellationToken: cancellationToken).ConfigureAwait(false);
        var baseWebUrl = baseSite?.WebUrl ?? site.WebUrl ?? "";

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
    /// Parses a SharePoint URL into (sitePath, subsitePath, listName).
    /// subsitePath is the server-relative path of the subsite within the parent site (e.g. "SupplierForms" or "ed74a958-...").
    /// Handles:
    ///   /sites/SiteName/Lists/ListName/AllItems.aspx
    ///   /sites/SiteName/SubSite/Lists/ListName/AllItems.aspx
    ///   /sites/SiteName/SubSite/LibraryName/Forms/AllItems.aspx
    ///   /sites/SiteName/LibraryName/Forms/AllItems.aspx
    ///   /rootSubsite/childSubsite/LibraryName
    /// </summary>
    private static (string sitePath, string? subsitePath, string listName) ParseListOrLibraryUrl(string url)
    {
        try
        {
            var uri = new Uri(url);
            var path = uri.AbsolutePath.TrimEnd('/');

            // Strip trailing view suffixes: /AllItems.aspx, /Forms/AllItems.aspx
            path = StripViewSuffix(path);

            var sitesIndex = path.IndexOf("/sites/", StringComparison.OrdinalIgnoreCase);

            // Root site collection (no /sites/ segment) e.g. /itsp/itst/LibraryName
            if (sitesIndex < 0)
            {
                if (string.IsNullOrEmpty(path) || path == "/")
                    return ("", null, "");

                var rootSitePath = $"{uri.Host}:/sites/root";
                var segments = path.TrimStart('/').Split('/');

                // Check for /Lists/ pattern first
                var listsIdx = path.IndexOf("/Lists/", StringComparison.OrdinalIgnoreCase);
                if (listsIdx >= 0)
                {
                    var listPart = path[(listsIdx + 7)..];
                    var end = listPart.IndexOf('/');
                    var listName = end > 0 ? Uri.UnescapeDataString(listPart[..end]) : Uri.UnescapeDataString(listPart);
                    // Everything before /Lists/ after the first segment is the subsite path
                    var beforeLists = path[..listsIdx].TrimStart('/');
                    var subsiteSegments = beforeLists.Split('/', StringSplitOptions.RemoveEmptyEntries);
                    var subsite = subsiteSegments.Length > 0 ? string.Join("/", subsiteSegments) : null;
                    return (rootSitePath, subsite, listName);
                }

                // No /Lists/ — last segment is library name, everything before is subsite path
                if (segments.Length == 1)
                    return (rootSitePath, null, Uri.UnescapeDataString(segments[0]));

                var libName = Uri.UnescapeDataString(segments[^1]);
                var subsitePath = string.Join("/", segments[..^1]);
                return (rootSitePath, subsitePath, libName);
            }

            // /sites/SiteName/...
            var afterSites = path[sitesIndex..];
            var siteNameEnd = afterSites.IndexOf('/', 7); // skip "/sites/"
            var siteRelativePath = siteNameEnd > 0 ? afterSites[..siteNameEnd] : afterSites;
            var sitePath2 = $"{uri.Host}:{siteRelativePath}";

            if (siteNameEnd < 0)
                return (sitePath2, null, "");

            var afterSiteName = afterSites[siteNameEnd..]; // starts with /

            // Check for /Lists/ anywhere after the site name
            var listsIndex2 = afterSiteName.IndexOf("/Lists/", StringComparison.OrdinalIgnoreCase);
            if (listsIndex2 >= 0)
            {
                var listPart = afterSiteName[(listsIndex2 + 7)..];
                var end = listPart.IndexOf('/');
                var listName = end > 0 ? Uri.UnescapeDataString(listPart[..end]) : Uri.UnescapeDataString(listPart);

                // Anything between the site name and /Lists/ is a subsite path
                var beforeLists = afterSiteName[..listsIndex2].Trim('/');
                var subsite = string.IsNullOrEmpty(beforeLists) ? null : beforeLists;
                return (sitePath2, subsite, listName);
            }

            // No /Lists/ — walk segments: last is library name, everything before is subsite
            var remainingSegments = afterSiteName.Trim('/').Split('/', StringSplitOptions.RemoveEmptyEntries);

            if (remainingSegments.Length == 0)
                return (sitePath2, null, "");

            if (remainingSegments.Length == 1)
                return (sitePath2, null, Uri.UnescapeDataString(remainingSegments[0]));

            // Last segment is the library, everything before is subsite
            var libraryName = Uri.UnescapeDataString(remainingSegments[^1]);
            var subsitePart = string.Join("/", remainingSegments[..^1]);

            // Skip known non-subsite segments
            if (libraryName.StartsWith("_", StringComparison.Ordinal) ||
                libraryName.Equals("Forms", StringComparison.OrdinalIgnoreCase))
                return (sitePath2, null, "");

            return (sitePath2, subsitePart, libraryName);
        }
        catch
        {
            return ("", null, "");
        }
    }

    /// <summary>
    /// Strips trailing view/form suffixes from a SharePoint path.
    ///   .../VendorCreationFormData/AllItems.aspx  -> .../VendorCreationFormData
    ///   .../Shared Documents/Forms/AllItems.aspx  -> .../Shared Documents
    /// </summary>
    private static string StripViewSuffix(string path)
    {
        // Strip .aspx file
        if (path.EndsWith(".aspx", StringComparison.OrdinalIgnoreCase))
        {
            var lastSlash = path.LastIndexOf('/');
            if (lastSlash > 0)
                path = path[..lastSlash];
        }

        // Strip trailing /Forms segment
        if (path.EndsWith("/Forms", StringComparison.OrdinalIgnoreCase))
            path = path[..^6];

        return path.TrimEnd('/');
    }

    private static string GetSitePathFromUrl(string url)
    {
        try
        {
            var uri = new Uri(url);
            var path = uri.AbsolutePath.TrimEnd('/');
            var sitesIndex = path.IndexOf("/sites/", StringComparison.OrdinalIgnoreCase);
            if (sitesIndex >= 0)
            {
                var afterSites = path[sitesIndex..];
                var nextSlash = afterSites.IndexOf('/', 7);
                var sitePath = nextSlash > 0 ? afterSites[..nextSlash] : afterSites;
                return $"{uri.Host}:{sitePath}";
            }
            if (string.IsNullOrEmpty(path) || path == "/")
                return $"{uri.Host}:/sites/root";
            return $"{uri.Host}:{path}";
        }
        catch
        {
            return url;
        }
    }

    /// <summary>
    /// Looks up a subsite within a parent site by its server-relative subsite path.
    /// e.g. subsitePath = "SupplierForms" or "ed74a958-62b4-45b9-9426-a8ca925037f7"
    /// </summary>
    private async Task<Site?> GetSubsiteByPathAsync(string parentSiteId, string subsitePath, CancellationToken cancellationToken)
    {
        try
        {
            var subsites = await _graph!.Sites[parentSiteId].Sites
                .GetAsync(r => r.QueryParameters.Top = 200, cancellationToken)
                .ConfigureAwait(false);

            if (subsites?.Value == null) return null;

            // Match by the last segment of the subsite's server-relative URL
            var subsiteSegment = subsitePath.Split('/').Last();

            var match = subsites.Value.FirstOrDefault(s =>
            {
                if (s.WebUrl == null) return false;
                var rel = new Uri(s.WebUrl).AbsolutePath.TrimEnd('/');
                var lastSeg = rel.Split('/').Last();
                return string.Equals(lastSeg, subsiteSegment, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(rel.TrimStart('/'), subsitePath, StringComparison.OrdinalIgnoreCase);
            });

            if (match != null)
                _logger?.LogInformation("Resolved subsite '{SubsitePath}' -> '{WebUrl}'", subsitePath, match.WebUrl);
            else
                _logger?.LogWarning("Subsite '{SubsitePath}' not found under site '{ParentSiteId}'", subsitePath, parentSiteId);

            return match;
        }
        catch (Exception ex)
        {
            _logger?.LogWarning(ex, "Failed to enumerate subsites for '{ParentSiteId}'", parentSiteId);
            return null;
        }
    }

    /// <summary>Parse Modified from Graph (DateTimeOffset, string, JsonElement, Json, or ToString).</summary>
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
            {
                foreach (var want in new[] { "email", "mail", "userPrincipalName" })
                    foreach (var kv in identity.AdditionalData)
                    {
                        if (!kv.Key.Equals(want, StringComparison.OrdinalIgnoreCase)) continue;
                        var s = GraphValueToTrimmedString(kv.Value);
                        if (!string.IsNullOrWhiteSpace(s)) return s;
                    }
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
        {
            foreach (var want in new[] { "displayName", "LookupValue", "Email", "Title", "Name", "preferredName", "mail" })
                foreach (var kv in j.AdditionalData)
                {
                    if (!kv.Key.Equals(want, StringComparison.OrdinalIgnoreCase)) continue;
                    var str = GraphValueToTrimmedString(kv.Value);
                    if (!string.IsNullOrWhiteSpace(str)) return str;
                }
        }
        if (o is IDictionary<string, object> nested)
        {
            foreach (var want in new[] { "displayName", "LookupValue", "Email", "Title", "Name", "preferredName", "mail" })
            {
                if (!TryGetFieldValue(nested, want, out var inner) || inner == null) continue;
                var str = ParsePersonOrClaimsField(inner);
                if (!string.IsNullOrWhiteSpace(str)) return str;
            }
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
        var isRoot = hostAndPath.EndsWith(":/sites/root", StringComparison.OrdinalIgnoreCase);
        var hostname = isRoot ? hostAndPath.AsSpan(0, hostAndPath.IndexOf(":/", StringComparison.Ordinal)).ToString() : null;

        var toTry = new List<string>();
        if (isRoot)
        {
            toTry.Add("root");
            if (!string.IsNullOrEmpty(hostname)) toTry.Add(hostname!);
        }
        else
        {
            toTry.Add(hostAndPath);
            var sep = hostAndPath.IndexOf(":/", StringComparison.Ordinal);
            if (sep > 0)
            {
                var host = hostAndPath[..sep];
                var rel = hostAndPath[(sep + 2)..].TrimStart('/');
                if (!string.IsNullOrEmpty(rel) && !rel.StartsWith("sites/", StringComparison.OrdinalIgnoreCase))
                {
                    var firstSeg = rel.Split('/', StringSplitOptions.RemoveEmptyEntries).FirstOrDefault();
                    if (!string.IsNullOrEmpty(firstSeg))
                        toTry.Add($"{host}:/sites/{firstSeg}");
                }
            }
        }

        Exception? lastEx = null;
        foreach (var siteId in toTry.Distinct(StringComparer.OrdinalIgnoreCase))
        {
            if (string.IsNullOrEmpty(siteId)) continue;
            try { return await _graph!.Sites[siteId].GetAsync(requestConfig => { }, cancellationToken).ConfigureAwait(false); }
            catch (Exception ex) { lastEx = ex; }
        }

        _logger?.LogWarning(lastEx,
            "GetSiteByPathAsync failed. Tried: {Ids}. Last error: {Message}",
            string.Join(", ", toTry.Distinct(StringComparer.OrdinalIgnoreCase)),
            lastEx?.Message ?? "(none)");
        return null;
    }

    private async Task<GraphList?> GetListByNameAsync(string siteId, string listName, CancellationToken cancellationToken)
    {
        try
        {
            var lists = await _graph!.Sites[siteId].Lists.GetAsync(r => r.QueryParameters.Top = 500, cancellationToken).ConfigureAwait(false);
            return lists?.Value?.FirstOrDefault(l =>
                string.Equals(l.DisplayName, listName, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(l.Name, listName, StringComparison.OrdinalIgnoreCase));
        }
        catch { return null; }
    }
}
