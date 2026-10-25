namespace RustPlusDesk.Core.Accounts;

public enum CloudAccountState { SignedOut, Active, Offline, UpdateRequired }

/// <summary>Cloud identity only; no bearer or companion credentials enter presentation.</summary>
public sealed record CloudAccount(CloudAccountState State, string Id = "", string DisplayName = "", string Email = "", DateTimeOffset? ExpiresAt = null);

public interface ICloudAccountService
{
    Task<CloudAccount> CheckAsync(CancellationToken cancellationToken);
    Task<CloudAccount> SignInAsync(CancellationToken cancellationToken);
    Task SignOutAsync(CancellationToken cancellationToken);
}

public enum CloudAccountFailure { Unavailable, Rejected, UpdateRequired, InvalidStorage, ChangedStorage, InvalidResponse }

public sealed class CloudAccountException(CloudAccountFailure failure) : Exception(failure.ToString())
{
    public CloudAccountFailure Failure { get; } = failure;
}
