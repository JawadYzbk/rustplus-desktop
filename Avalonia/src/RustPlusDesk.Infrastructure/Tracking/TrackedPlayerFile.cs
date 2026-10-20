using System.Text.Json;
using System.Text.Json.Serialization;
using RustPlusDesk.Core.Tracking;
using RustPlusDesk.Infrastructure.Storage;

namespace RustPlusDesk.Infrastructure.Tracking;

/// <summary>
/// Reference TrackingService's tracked players: the bookmarks the user keeps, in the file the reference keeps them
/// in.
///
/// The reference writes the file with File.WriteAllText; this port writes it privately and atomically, so a write
/// that is cut short cannot leave a half-written file behind. The file's shape is the reference's: a list of the
/// players, each with its profile id, its name, the server it was last seen on, its group and the sessions that were
/// collected before.
/// </summary>
public sealed class TrackedPlayerFile(string path) : ITrackedPlayerStore
{
    public static string DefaultPath => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "RustPlusDesk", "tracked_players.json");

    private static readonly JsonSerializerOptions Options = new() { WriteIndented = true };

    public async Task<IReadOnlyList<TrackedPlayer>> LoadAsync(CancellationToken ct)
    {
        byte[]? bytes = await AtomicPrivateFile.ReadAsync(path, ct).ConfigureAwait(false);
        if (bytes is null) return [];
        try
        {
            return JsonSerializer.Deserialize<List<TrackedPlayer>>(bytes) ?? [];
        }
        catch (JsonException)
        {
            // A file that cannot be read is a list that is not there, not a broken app.
            return [];
        }
    }

    public async Task SaveAsync(IReadOnlyList<TrackedPlayer> players, CancellationToken ct)
    {
        // Reference SaveDB: the sessions older than twelve weeks are dropped when the file is written.
        var cutoff = DateTime.UtcNow.AddDays(-TrackedPlayerRules.SessionRetentionDays);
        var kept = players.Select(player => TrackedPlayerRules.Prune(player, cutoff)).ToList();
        byte[]? original = await AtomicPrivateFile.ReadAsync(path, ct).ConfigureAwait(false);
        byte[] data = JsonSerializer.SerializeToUtf8Bytes(kept, Options);
        await AtomicPrivateFile.WriteAsync(path, data, original, backup: false, ct).ConfigureAwait(false);
    }
}
