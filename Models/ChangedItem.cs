namespace SharepointDailyDigest.Models;

/// <summary>
/// A list or library item that was new or changed in the last 24 hours.
/// </summary>
public record ChangedItem(string Title, string WebUrl, DateTimeOffset Modified, string? ModifiedBy);
