    /// <summary>
    /// Given a site URL (not a list URL), returns ordered Graph site path candidates to try.
    /// e.g. https://streamflogroup.sharepoint.com/itsp ->
    ///   ["streamflogroup.sharepoint.com:/itsp", "streamflogroup.sharepoint.com"]
    /// e.g. https://streamflogroup.sharepoint.com/sites/OFForms/SupplierForms ->
    ///   ["streamflogroup.sharepoint.com:/sites/OFForms/SupplierForms", "streamflogroup.sharepoint.com:/sites/OFForms"]
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
                // Root collection subsite e.g. /itsp or /itsp/itst
                // Graph expects host:/path format for subsites
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
                // parts[0] = "sites", parts[1] = SiteName, parts[2+] = subsite segments
                for (var i = parts.Length; i >= 2; i--)
                    candidates.Add($"{uri.Host}:/{string.Join("/", parts[..i])}");
            }

            return candidates;
        }
        catch { return new List<string>(); }
    }
