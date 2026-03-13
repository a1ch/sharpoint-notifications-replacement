namespace SharepointDailyDigest.Models;

/// <summary>
/// One row from the config SharePoint list: Title = list/library URL, Email = recipient, optional Brand for email styling.
/// </summary>
public record ConfigListItem(string ListOrLibraryUrl, string Email, string? Brand = null);
