using System.Globalization;
using RustPlusDesk.Core.Connections;
using RustPlusDesk.Core.Tracking;

namespace RustPlusDesk.Infrastructure.Steam;

/// <summary>
/// Reference TrackingService.FetchPlayerNameAsync: the name of a player is read from their public Steam profile.
///
/// The reference asks for the profile XML and takes the name out of it, and falls back to the id it was given when
/// the profile is private, rate-limited or the machine is offline. The port asks for the same XML the avatar comes
/// from and keeps the same fallback.
/// </summary>
public sealed class SteamCommunityProfileNameSource(HttpClient http) : IPlayerNameSource
{
    public async Task<string> ResolveAsync(string id, CancellationToken ct)
    {
        if (!TrackedPlayerRules.IsSteamProfileId(id)) return id;
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(ct);
        timeout.CancelAfter(TimeSpan.FromSeconds(8));
        try
        {
            string xml = await http.GetStringAsync(string.Format(CultureInfo.InvariantCulture, SteamAvatarXml.ProfileUrlFormat, id), timeout.Token).ConfigureAwait(false);
            return SteamAvatarXml.PlayerName(xml) ?? id;
        }
        catch (OperationCanceledException) when (ct.IsCancellationRequested) { throw; }
        catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException or InvalidOperationException) { return id; }
    }
}
