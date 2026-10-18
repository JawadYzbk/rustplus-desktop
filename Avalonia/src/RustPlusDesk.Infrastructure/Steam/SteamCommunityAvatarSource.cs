using System.Globalization;
using RustPlusDesk.Core.Connections;

namespace RustPlusDesk.Infrastructure.Steam;

/// <summary>
/// Reference AvatarLoader.FetchSteamAvatarInternalAsync: the public profile XML names the avatar (full, else medium), whose
/// bytes are downloaded; private, rate-limited or offline profiles give null. Requests use an 8 second budget.
/// </summary>
public sealed class SteamCommunityAvatarSource(HttpClient http) : ISteamAvatarSource
{
    public async Task<byte[]?> LoadAsync(ulong steamId, CancellationToken ct)
    {
        if (steamId == 0) return null;
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(ct);
        timeout.CancelAfter(TimeSpan.FromSeconds(8));
        try
        {
            string xml = await http.GetStringAsync(string.Format(CultureInfo.InvariantCulture, SteamAvatarXml.ProfileUrlFormat, steamId), timeout.Token).ConfigureAwait(false);
            if (SteamAvatarXml.AvatarUrl(xml) is not { } url || !Uri.TryCreate(url, UriKind.Absolute, out var uri) || uri.Scheme != Uri.UriSchemeHttps) return null;
            byte[] bytes = await http.GetByteArrayAsync(uri, timeout.Token).ConfigureAwait(false);
            return bytes.Length == 0 ? null : bytes;
        }
        catch (OperationCanceledException) when (ct.IsCancellationRequested) { throw; }
        catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException or InvalidOperationException) { return null; }
    }
}
