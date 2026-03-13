using SharepointDailyDigest.Models;

namespace SharepointDailyDigest.Services;

public interface IEmailService
{
    Task SendDigestAsync(string toEmail, string listOrLibraryName, IReadOnlyList<ChangedItem> changes, string? brand = null, string? siteName = null, CancellationToken cancellationToken = default);
}
