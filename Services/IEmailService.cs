using SharepointDailyDigest.Models;

namespace SharepointDailyDigest.Services;

public interface IEmailService
{
    Task SendDigestAsync(string toEmail, string listOrLibraryName, IReadOnlyList<ChangedItem> changes, CancellationToken cancellationToken = default);
}
