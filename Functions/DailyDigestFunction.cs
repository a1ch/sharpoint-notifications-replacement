using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using SharepointDailyDigest.Services;

namespace SharepointDailyDigest.Functions;

public class DailyDigestFunction
{
    private readonly ISharePointDigestService _sharePoint;
    private readonly IEmailService _email;
    private readonly ILogger _logger;

    public DailyDigestFunction(
        ISharePointDigestService sharePoint,
        IEmailService email,
        ILoggerFactory loggerFactory)
    {
        _sharePoint = sharePoint;
        _email = email;
        _logger = loggerFactory.CreateLogger<DailyDigestFunction>();
    }

    /// <summary>
    /// Runs daily at 8:00 AM. Set WEBSITE_TIME_ZONE in Azure (e.g. Mountain Standard Time) for local time.
    /// </summary>
    [Function("DailyDigest")]
    public async Task Run([TimerTrigger("0 0 8 * * *")] TimerInfo timer, CancellationToken cancellationToken)
    {
        if (!IsDigestEnabled())
        {
            _logger.LogInformation(
                "Daily digest skipped: DIGEST_ENABLED is not true. Set DIGEST_ENABLED to true in app settings when you want digests (timer still runs at 8 AM).");
            return;
        }

        _logger.LogInformation("Daily digest started at {Time}", DateTime.UtcNow);

        IReadOnlyList<Models.ConfigListItem> configRows;
        try
        {
            configRows = await _sharePoint.GetConfigListItemsAsync(cancellationToken).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to read config list.");
            throw;
        }

        if (configRows.Count == 0)
        {
            _logger.LogInformation("No config rows found. Exiting.");
            return;
        }

        var since = DateTimeOffset.UtcNow.AddHours(-24);

        foreach (var row in configRows)
        {
            try
            {
                var changes = await _sharePoint.GetRecentChangesAsync(row.ListOrLibraryUrl, since, cancellationToken).ConfigureAwait(false);
                var listName = GetListNameFromUrl(row.ListOrLibraryUrl);
                var siteName = GetSiteNameFromUrl(row.ListOrLibraryUrl);
                if (changes.Count > 0)
                    await _email.SendDigestAsync(row.Email, listName, changes, row.Brand, siteName, row.ListOrLibraryUrl, cancellationToken).ConfigureAwait(false);
                else
                    _logger.LogInformation("No changes for {Url}, skipping email to {Email}", row.ListOrLibraryUrl, row.Email);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error processing {Url} for {Email}", row.ListOrLibraryUrl, row.Email);
                // Continue with next row
            }
        }

        _logger.LogInformation("Daily digest finished.");
    }

    /// <summary>Digest work runs only when DIGEST_ENABLED is true/1/yes. Default off so deploys do not send mail until explicitly enabled.</summary>
    private static bool IsDigestEnabled()
    {
        var v = Environment.GetEnvironmentVariable("DIGEST_ENABLED");
        if (string.IsNullOrWhiteSpace(v)) return false;
        return v.Equals("true", StringComparison.OrdinalIgnoreCase)
            || v == "1"
            || v.Equals("yes", StringComparison.OrdinalIgnoreCase);
    }

    private static string GetListNameFromUrl(string url)
    {
        try
        {
            var path = new Uri(url).AbsolutePath.TrimEnd('/');
            var listsIndex = path.IndexOf("/Lists/", StringComparison.OrdinalIgnoreCase);
            if (listsIndex >= 0)
            {
                var listPart = path.Substring(listsIndex + 7);
                var end = listPart.IndexOf('/');
                return end > 0 ? Uri.UnescapeDataString(listPart.Substring(0, end)) : Uri.UnescapeDataString(listPart);
            }
            var segments = path.TrimStart('/').Split('/');
            var last = segments.Length > 0 ? segments[^1] : null;
            return !string.IsNullOrEmpty(last) ? Uri.UnescapeDataString(last) : path;
        }
        catch
        {
            return url;
        }
    }

    private static string GetSiteNameFromUrl(string url)
    {
        try
        {
            var path = new Uri(url).AbsolutePath.TrimEnd('/');
            var sitesIndex = path.IndexOf("/sites/", StringComparison.OrdinalIgnoreCase);
            if (sitesIndex >= 0)
            {
                var afterSites = path.Substring(sitesIndex + 7);
                var nextSlash = afterSites.IndexOf('/');
                var segment = nextSlash > 0 ? afterSites.Substring(0, nextSlash) : afterSites;
                return !string.IsNullOrEmpty(segment) ? Uri.UnescapeDataString(segment) : "SharePoint";
            }
            return "Root";
        }
        catch
        {
            return "";
        }
    }
}
