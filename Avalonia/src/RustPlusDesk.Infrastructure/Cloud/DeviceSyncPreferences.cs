using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;
using RustPlusDesk.Core.Cloud;
using RustPlusDesk.Infrastructure.Storage;

namespace RustPlusDesk.Infrastructure.Cloud;

/// <summary>Separate, default-off local opt-in; disabling suppresses this process even if persistence fails.</summary>
public sealed class DeviceSyncPreferences(string path) : IDeviceSyncPreferences
{
    private static readonly SemaphoreSlim Writes = new(1, 1);
    private int suppressed;
    private int stopVersion;
    public static string DefaultPath => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "RustPlusDesk", "cache", "avalonia_device_sync.json");

    public async Task<DeviceSyncPreference> ReadAsync(CancellationToken ct)
    {
        ct.ThrowIfCancellationRequested();
        if (Volatile.Read(ref suppressed) != 0) return new("", "", false);
        try
        {
            byte[]? bytes = await AtomicPrivateFile.ReadAsync(path, ct).ConfigureAwait(false);
            if (bytes is null) return new("", "", false);
            if (bytes.Length > 65536) throw new JsonException();
            var node = JsonNode.Parse(bytes)!.AsObject();
            bool enabled = node["Enabled"]?.GetValue<bool>() ?? false;
            string user = CloudSessionFile.Text(node["UserId"], 128);
            string steam = CloudSessionFile.Text(node["SteamId"], 32);
            if (enabled) Validate(user, steam);
            return new(user, steam, enabled && Volatile.Read(ref suppressed) == 0, CloudSessionFile.Text(node["Revision"], 64));
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException or InvalidOperationException or ArgumentException or FormatException)
        { throw new DeviceSyncPreferencesException(); }
    }

    public async Task SetAsync(string userId, string steamId, bool enabled, CancellationToken ct)
    {
        int version = Volatile.Read(ref stopVersion);
        if (!enabled) { Interlocked.Increment(ref stopVersion); Interlocked.Exchange(ref suppressed, 1); }
        else Validate(userId, steamId);
        await Writes.WaitAsync(ct).ConfigureAwait(false);
        try
        {
            byte[]? before = await AtomicPrivateFile.ReadAsync(path, ct).ConfigureAwait(false);
            JsonObject node;
            try
            {
                if (before?.Length > 65536) throw new JsonException();
                node = before is null ? new() : JsonNode.Parse(before)!.AsObject();
            }
            catch (Exception ex) when (!enabled && ex is JsonException or InvalidOperationException or ArgumentException)
            { node = new(); } // A private backup preserves an unreadable preference when explicitly disabling.
            node["UserId"] = enabled ? userId : "";
            node["SteamId"] = enabled ? steamId : "";
            node["Enabled"] = enabled;
            node["Revision"] = Guid.NewGuid().ToString("N");
            await AtomicPrivateFile.WriteAsync(path, JsonSerializer.SerializeToUtf8Bytes(node), before, true, ct).ConfigureAwait(false);
            if (enabled && version == Volatile.Read(ref stopVersion)) Interlocked.Exchange(ref suppressed, 0);
        }
        catch (OperationCanceledException) { throw; }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException or InvalidOperationException or ArgumentException)
        { throw new DeviceSyncPreferencesException(); }
        finally { Writes.Release(); }
    }

    private static void Validate(string userId, string steamId)
    {
        if (string.IsNullOrWhiteSpace(userId) || userId.Length > 128 || userId.Any(char.IsControl) ||
            steamId.Length != 17 || !ulong.TryParse(steamId, NumberStyles.None, CultureInfo.InvariantCulture, out _))
            throw new DeviceSyncPreferencesException();
    }
}
