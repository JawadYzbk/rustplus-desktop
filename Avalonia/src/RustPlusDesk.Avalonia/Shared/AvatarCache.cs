using Avalonia.Media.Imaging;
using RustPlusDesk.Core.Connections;

namespace RustPlusDesk.Avalonia.Shared;

/// <summary>
/// Reference AvatarLoader + EnsureAvatarAsync: Steam avatars decoded once and kept in memory, one request per Steam ID at a
/// time, and a failed or missing avatar is tried again after 30 seconds at the earliest.
/// </summary>
public sealed class AvatarCache(ISteamAvatarSource source, Func<DateTime>? now = null)
{
    private readonly Dictionary<ulong, Bitmap> avatars = [];
    private readonly HashSet<ulong> loading = [];
    private readonly Dictionary<ulong, DateTime> nextTry = [];
    private readonly Func<DateTime> now = now ?? (() => DateTime.UtcNow);

    /// <summary>Raised on the caller's context after an avatar arrived.</summary>
    public event Action<ulong>? Loaded;
    /// <summary>Changes whenever a new avatar arrives (drives re-rendering).</summary>
    public int Version { get; private set; }

    public Bitmap? Get(ulong steamId) => avatars.GetValueOrDefault(steamId);

    /// <summary>Starts loading the avatar unless it is cached, loading or waiting for its retry.</summary>
    public Task RequestAsync(ulong steamId)
    {
        if (steamId == 0 || avatars.ContainsKey(steamId) || loading.Contains(steamId)
            || nextTry.TryGetValue(steamId, out var next) && now() < next) return Task.CompletedTask;
        return LoadAsync(steamId);
    }

    private async Task LoadAsync(ulong steamId)
    {
        loading.Add(steamId);
        try
        {
            byte[]? bytes = await source.LoadAsync(steamId, CancellationToken.None);
            Bitmap? bitmap = null;
            if (bytes is not null)
            {
                try { using var stream = new MemoryStream(bytes, writable: false); bitmap = new Bitmap(stream); }
                catch (Exception) { bitmap = null; }
            }
            if (bitmap is null) { nextTry[steamId] = now() + SteamAvatarXml.RetryInterval; return; }
            avatars[steamId] = bitmap;
            nextTry.Remove(steamId);
            Version++;
            Loaded?.Invoke(steamId);
        }
        catch (Exception ex) when (ex is not OperationCanceledException) { nextTry[steamId] = now() + SteamAvatarXml.RetryInterval; }
        finally { loading.Remove(steamId); }
    }
}
