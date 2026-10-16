using RustPlusDesk.Core.Accounts;

namespace RustPlusDesk.Core.Cloud;

/// <summary>Cloud account details safe for presentation; session tokens remain in infrastructure.</summary>
public sealed record CloudIdentity(string Id, string DisplayName, string Email, string SteamId,
    IReadOnlyList<string> Providers, bool HasPassword, bool SyncAccepted);

public sealed record CloudSyncConsent(string UserId, string SteamId, bool Granted);

public sealed record CloudAccount(AccountState State, CloudIdentity? Identity = null, DateTimeOffset? ExpiresAt = null,
    CloudSyncConsent? Consent = null);

public interface ICloudAccountService
{
    Task<CloudAccount> CheckAsync(CancellationToken cancellationToken);
    Task<CloudAccount> SignInAsync(CancellationToken cancellationToken);
    Task SignOutAsync(string expectedUserId, CancellationToken cancellationToken);
    Task<CloudAccount> SetSyncConsentAsync(string expectedUserId, string? steamId, bool granted, CancellationToken cancellationToken);
}

public enum CloudFailure
{
    Rejected, AccessDenied, Unavailable, InvalidConfiguration, ChangedConfiguration, InvalidResponse,
    InvalidCallback, UpgradeRequired, RateLimited, IdentityMismatch, SessionNotSaved, ConsentNotSaved
}

public sealed class CloudAccountException(CloudFailure failure) : Exception(failure.ToString())
{
    public CloudFailure Failure { get; } = failure;
}
