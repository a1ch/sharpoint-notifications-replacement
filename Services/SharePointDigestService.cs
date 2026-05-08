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

        // Build the full site URL from parsed components and use BuildSitePathCandidates
        // so the correct host:/path format is used for both named and root collection subsites.
        var fullSiteUrl = subsitePath != null
            ? $"https://{sitePath.Replace(":/", "/")}/{subsitePath}"
            : $"https://{sitePath.Replace(":/", "/")}";
        var sitePathsToTry = BuildSitePathCandidates(fullSiteUrl);
        // Always include the bare parent site path as final fallback
        if (!sitePathsToTry.Contains(sitePath, StringComparer.OrdinalIgnoreCase))
            sitePathsToTry.Add(sitePath);

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
