using RustPlusDesk.Core.Alerts;
using RustPlusDesk.Core.Cloud;

namespace RustPlusDesk.Infrastructure.Cloud;

/// <summary>Signed in = the stored cloud session has a token that has not expired (reference IsCloudConnected).</summary>
public sealed class CloudSignInState(string sessionPath) : ICloudSignInState
{
    public async Task<bool> IsSignedInAsync(CancellationToken ct)
    {
        try
        {
            var session = await CloudSessionFile.ReadAsync(sessionPath, ct).ConfigureAwait(false);
            return session.Token.Length > 0 && (session.ExpiresAt is not { } expires || expires > DateTimeOffset.UtcNow);
        }
        catch (CloudAccountException) { return false; }
    }
}
