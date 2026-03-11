using SharepointDailyDigest.Models;

namespace SharepointDailyDigest.Services;

public interface ISharePointDigestService
{
    /// <summary>
    /// Reads the config list (Title = list/library URL, Email = recipient) and returns rows.
    /// </summary>
    Task<IReadOnlyList<ConfigListItem>> GetConfigListItemsAsync(CancellationToken cancellationToken = default);

    /// <summary>
    /// Gets items from the given list or library that were modified in the last 24 hours.
    /// </summary>
    Task<IReadOnlyList<ChangedItem>> GetRecentChangesAsync(string listOrLibraryUrl, DateTimeOffset sinceUtc, CancellationToken cancellationToken = default);
}
