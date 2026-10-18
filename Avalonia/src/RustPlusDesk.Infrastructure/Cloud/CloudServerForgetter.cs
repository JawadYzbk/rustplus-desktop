using System.Globalization;
using System.Text.Json;
using RustPlusDesk.Core.Cloud;
using RustPlusDesk.Core.Servers;

namespace RustPlusDesk.Infrastructure.Cloud;

/// <summary>
/// Reference CloudPairingImporter.ForgetServerAsync: the removal goes into the shared pending list
/// (%APPDATA%/RustPlusDesk/cache/cloud_pairings_removed.json, "host:port" in lower case → UTC time) so a cloud restore
/// never brings the server back, then DELETE me/servers/{host-port} removes the pairing while signed in. A removal that
/// could not be sent stays pending for the next sync.
/// </summary>
public sealed class CloudServerForgetter(HttpClient http, string sessionPath, string clientVersion, string removedPath) : ICloudServerForgetter
{
    public static string DefaultRemovedPath => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "RustPlusDesk", "cache", "cloud_pairings_removed.json");
    private static readonly SemaphoreSlim Gate = new(1, 1);

    /// <summary>Reference RemovalKey.</summary>
    public static string RemovalKey(string host, int port) => $"{host.Trim().ToLowerInvariant()}:{port.ToString(CultureInfo.InvariantCulture)}";

    public async Task<bool> ForgetAsync(string host, int port, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(host) || port <= 0) return false;
        string key = RemovalKey(host, port);
        await ChangeAsync(removed => removed[key] = DateTime.UtcNow, ct).ConfigureAwait(false);

        string token;
        try
        {
            var session = await CloudSessionFile.ReadAsync(sessionPath, ct).ConfigureAwait(false);
            token = session.Token;
            if (token.Length == 0 || session.ExpiresAt is { } expires && expires <= DateTimeOffset.UtcNow) return false;
        }
        catch (CloudAccountException) { return false; }

        try
        {
            int separator = key.LastIndexOf(':');
            await new CloudApiClient(http, clientVersion).DeleteAsync("me/servers/" + Uri.EscapeDataString($"{key[..separator]}-{key[(separator + 1)..]}"), token, ct)
                .ConfigureAwait(false);
        }
        catch (OperationCanceledException) when (ct.IsCancellationRequested) { throw; }
        catch (Exception ex) when (ex is HttpRequestException or CloudAccountException or TaskCanceledException) { return false; }
        await ChangeAsync(removed => removed.Remove(key), ct).ConfigureAwait(false);
        return true;
    }

    // Reference StorageService cache file: indented JSON, unreadable content counts as empty.
    private async Task ChangeAsync(Action<Dictionary<string, DateTime>> change, CancellationToken ct)
    {
        await Gate.WaitAsync(ct).ConfigureAwait(false);
        try
        {
            Dictionary<string, DateTime> removed = [];
            try
            {
                if (File.Exists(removedPath)) removed = JsonSerializer.Deserialize<Dictionary<string, DateTime>>(await File.ReadAllTextAsync(removedPath, ct).ConfigureAwait(false)) ?? [];
            }
            catch (JsonException) { removed = []; }
            removed = new(removed, StringComparer.Ordinal);
            change(removed);
            Directory.CreateDirectory(Path.GetDirectoryName(removedPath)!);
            await File.WriteAllTextAsync(removedPath, JsonSerializer.Serialize(removed, new JsonSerializerOptions { WriteIndented = true }), ct).ConfigureAwait(false);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { }
        finally { Gate.Release(); }
    }
}
