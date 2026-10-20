using System.Globalization;
using System.Text.Json;
using RustPlusDesk.Core.Tracking;

namespace RustPlusDesk.Infrastructure.Tracking;

/// <summary>
/// Reference PlayerWipeTrackerStore: the observations of the players the user tracks, in the app's data folder.
///
/// The layout is the reference's: a folder per server, a folder per wipe in it, and one file per player in that. The
/// file is the compact record file the reference writes (see <see cref="Core.Tracking.TrackerRecordFile"/>); the
/// history that app wrote as JSON lines is converted when it is read.
///
/// The port writes the files directly under the lock of the file, where the reference queues the writes so that
/// capturing a poll never waits for the disk.
/// </summary>
public sealed class PlayerWipeTrackerStore(string rootDirectory) : IDisposable
{
    private const string RecordExtension = ".pwt";
    private const string LegacyExtension = ".jsonl";

    private static readonly JsonSerializerOptions Json = new() { WriteIndented = true };

    private readonly string root = Path.GetFullPath(rootDirectory);
    private readonly Dictionary<string, SemaphoreSlim> gates = new(StringComparer.OrdinalIgnoreCase);
    private readonly object gateLock = new();

    /// <summary>Reference PlayerWipeTrackerStore: the folder the observations are kept in.</summary>
    public static string DefaultRoot => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "RustPlusDesk", "player-wipes");

    public string Root => root;

    /// <summary>Reference StorageBytes: how much the store holds.</summary>
    public long StorageBytes => Directory.Exists(root)
        ? Directory.EnumerateFiles(root, "*", SearchOption.AllDirectories).Sum(path => new FileInfo(path).Length)
        : 0;

    /// <summary>Reference Append: the observation is written to the player's file.</summary>
    public async Task AppendAsync(string serverKey, string wipeKey, ulong steamId, PlayerObservation observation,
        CancellationToken ct = default)
    {
        string path = FilePath(serverKey, wipeKey, steamId);
        var gate = Gate(path);
        await gate.WaitAsync(ct).ConfigureAwait(false);
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            MigrateLegacy(path, steamId);

            var state = EnsureState(path, steamId);
            if (state is null && File.Exists(path))
            {
                // Not our format, or a header torn by a crash on the very first write. The bytes are kept aside instead
                // of appending records nobody could read back.
                File.Move(path, path + ".unreadable", overwrite: true);
            }

            byte[] record;
            if (state is null)
            {
                state = TrackerRecordFile.CreateNew(observation.TimestampUtc);
                record = state.Encode(observation);
                await WriteAllAsync(path, TrackerRecordFile.Header(state.BaseTicks), record, ct).ConfigureAwait(false);
                state.Appended(record.Length);
                return;
            }

            record = state.Encode(observation);
            await AppendAllAsync(path, record, state.ValidLength, ct).ConfigureAwait(false);
            state.Appended(record.Length);
        }
        finally { gate.Release(); }
    }

    /// <summary>Reference Load: the observations the player has in that wipe.</summary>
    public IReadOnlyList<PlayerObservation> Load(string serverKey, string wipeKey, ulong steamId)
    {
        string path = FilePath(serverKey, wipeKey, steamId);
        var gate = Gate(path);
        gate.Wait();
        try
        {
            var observations = new List<PlayerObservation>();
            if (MigrateLegacy(path, steamId) && !File.Exists(path)) return observations;

            if (File.Exists(path))
            {
                byte[] bytes = File.ReadAllBytes(path);
                var state = TrackerRecordFile.Parse(bytes, steamId, observations);
                if (state is not null) States[path] = state;
            }
            return Distinct(observations);
        }
        finally { gate.Release(); }
    }

    /// <summary>The record file's state per path, so an append does not read the history again.</summary>
    private readonly Dictionary<string, TrackerRecordFile> States = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>Reference PlayerWipeTrackerStore.Distinct: the observations in time order, with exact repeats dropped.</summary>
    private static IReadOnlyList<PlayerObservation> Distinct(List<PlayerObservation> observations) =>
        observations
            .OrderBy(observation => observation.TimestampUtc)
            .Distinct()
            .ToArray();

    /// <summary>Reference SaveWipeMetadata: what the store keeps about a wipe.</summary>
    public void SaveWipeMetadata(string serverKey, string wipeKey, StoredWipeMetadata metadata)
    {
        string directory = WipeDirectory(serverKey, wipeKey);
        Directory.CreateDirectory(directory);
        File.WriteAllText(Path.Combine(directory, "wipe_info.json"), JsonSerializer.Serialize(metadata, Json));
    }

    /// <summary>Reference LoadWipeMetadata: what the store keeps about a wipe, when it is there.</summary>
    public StoredWipeMetadata? LoadWipeMetadata(string serverKey, string wipeKey)
    {
        string path = Path.Combine(WipeDirectory(serverKey, wipeKey), "wipe_info.json");
        if (!File.Exists(path)) return null;
        try { return JsonSerializer.Deserialize<StoredWipeMetadata>(File.ReadAllText(path)); }
        catch (JsonException) { return null; }
    }

    /// <summary>Reference SaveWipeMap: the map of the wipe the tracker draws on.</summary>
    public void SaveWipeMap(string serverKey, string wipeKey, TrackerWipeMap map)
    {
        string directory = WipeDirectory(serverKey, wipeKey);
        Directory.CreateDirectory(directory);
        File.WriteAllBytes(Path.Combine(directory, "map.png"), map.PngBytes);
        File.WriteAllText(Path.Combine(directory, "map.json"),
            JsonSerializer.Serialize(map with { PngBytes = Array.Empty<byte>() }, Json));
    }

    /// <summary>Reference LoadWipeMap: the map of the wipe, when it is there.</summary>
    public TrackerWipeMap? LoadWipeMap(string serverKey, string wipeKey)
    {
        string directory = WipeDirectory(serverKey, wipeKey);
        string imagePath = Path.Combine(directory, "map.png");
        string metadataPath = Path.Combine(directory, "map.json");
        if (!File.Exists(imagePath) || !File.Exists(metadataPath)) return null;
        try
        {
            var map = JsonSerializer.Deserialize<TrackerWipeMap>(File.ReadAllText(metadataPath));
            if (map is null) return null;
            return map with { PngBytes = File.ReadAllBytes(imagePath) };
        }
        catch (JsonException) { return null; }
        catch (IOException) { return null; }
    }

    /// <summary>Reference HasWipeMap: whether the wipe has a map.</summary>
    public bool HasWipeMap(string serverKey, string wipeKey)
    {
        string directory = WipeDirectory(serverKey, wipeKey);
        return File.Exists(Path.Combine(directory, "map.png")) && File.Exists(Path.Combine(directory, "map.json"));
    }

    /// <summary>Reference LoadPlayerIds: the players the wipe holds.</summary>
    public IReadOnlyList<ulong> LoadPlayerIds(string serverKey, string wipeKey)
    {
        string directory = WipeDirectory(serverKey, wipeKey);
        if (!Directory.Exists(directory)) return [];
        var ids = new List<ulong>();
        foreach (string? name in PlayerFiles(directory).Select(Path.GetFileNameWithoutExtension))
            if (ulong.TryParse(name, NumberStyles.None, CultureInfo.InvariantCulture, out var steamId)) ids.Add(steamId);
        return ids.Distinct().ToArray();
    }

    /// <summary>Reference GetStoredWipes: the wipes the store holds.</summary>
    public IReadOnlyList<StoredWipeSummary> GetStoredWipes()
    {
        var wipes = new List<StoredWipeSummary>();
        if (!Directory.Exists(root)) return wipes;
        foreach (var serverDirectory in Directory.EnumerateDirectories(root))
        {
            string serverKey = Path.GetFileName(serverDirectory);
            foreach (var wipeDirectory in Directory.EnumerateDirectories(serverDirectory))
            {
                string wipeKey = Path.GetFileName(wipeDirectory);
                var metadata = LoadWipeMetadata(serverKey, wipeKey);
                var files = PlayerFiles(wipeDirectory).ToArray();
                long bytes = files.Sum(path => new FileInfo(path).Length);
                var observations = LoadPlayerIds(serverKey, wipeKey);
                DateTime? lastObserved = null;
                foreach (ulong steamId in observations)
                {
                    var last = Load(serverKey, wipeKey, steamId).LastOrDefault()?.TimestampUtc;
                    if (last is { } value && (lastObserved is null || value > lastObserved)) lastObserved = value;
                }
                wipes.Add(new StoredWipeSummary(
                    serverKey,
                    metadata?.ServerName ?? serverKey,
                    wipeKey,
                    metadata?.WipeStartedAtUtc,
                    lastObserved,
                    observations.Count,
                    0,
                    bytes,
                    HasWipeMap(serverKey, wipeKey)));
            }
        }
        return wipes;
    }

    /// <summary>Reference DeleteWipe: the wipe is gone.</summary>
    public void DeleteWipe(string serverKey, string wipeKey)
    {
        string directory = WipeDirectory(serverKey, wipeKey);
        if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
    }

    /// <summary>Reference DeleteAll: every wipe is gone.</summary>
    public void DeleteAll()
    {
        if (Directory.Exists(root)) Directory.Delete(root, recursive: true);
        Directory.CreateDirectory(root);
    }

    public void Dispose()
    {
        lock (gateLock) gates.Clear();
    }

    // ---- the files -----------------------------------------------------------------------------

    private string FilePath(string serverKey, string wipeKey, ulong steamId) =>
        Path.Combine(WipeDirectory(serverKey, wipeKey), $"{steamId}{RecordExtension}");

    private string WipeDirectory(string serverKey, string wipeKey) =>
        Path.Combine(root, Safe(serverKey), Safe(wipeKey));

    /// <summary>Reference PlayerWipeTrackerStore.Safe: a folder name that is safe on every platform.</summary>
    private static string Safe(string value)
    {
        var invalid = Path.GetInvalidFileNameChars();
        var safe = new string(value.Select(ch => invalid.Contains(ch) ? '_' : ch).ToArray()).Trim();
        return string.IsNullOrWhiteSpace(safe) ? "unknown" : safe[..Math.Min(safe.Length, 120)];
    }

    private static IEnumerable<string> PlayerFiles(string directory) =>
        Directory.Exists(directory)
            ? Directory.EnumerateFiles(directory, "*" + RecordExtension, SearchOption.TopDirectoryOnly)
                .Concat(Directory.EnumerateFiles(directory, "*" + LegacyExtension, SearchOption.TopDirectoryOnly))
            : [];

    private SemaphoreSlim Gate(string path)
    {
        lock (gateLock)
        {
            if (!gates.TryGetValue(path, out var gate)) { gate = new SemaphoreSlim(1, 1); gates[path] = gate; }
            return gate;
        }
    }

    /// <summary>The record file's state, read from the file the first time it is needed.</summary>
    private TrackerRecordFile? EnsureState(string path, ulong steamId)
    {
        lock (gateLock)
        {
            if (States.TryGetValue(path, out var state)) return state;
        }
        if (!File.Exists(path)) return null;
        byte[] bytes = File.ReadAllBytes(path);
        var parsed = TrackerRecordFile.Parse(bytes, steamId, null);
        if (parsed is null) return null;
        lock (gateLock) States[path] = parsed;
        return parsed;
    }

    /// <summary>
    /// Reference MigrateLegacy: the history the app wrote as JSON lines is converted to the record file. It reports
    /// whether the conversion happened, so a file that could not be converted is read as it is.
    /// </summary>
    private bool MigrateLegacy(string path, ulong steamId)
    {
        string legacy = Path.ChangeExtension(path, LegacyExtension);
        if (!File.Exists(legacy)) return false;
        var observations = ReadLegacy(legacy);
        try
        {
            if (observations.Count > 0)
            {
                var state = TrackerRecordFile.CreateNew(observations[0].TimestampUtc);
                using var buffer = new MemoryStream();
                buffer.Write(TrackerRecordFile.Header(state.BaseTicks));
                foreach (var observation in observations)
                {
                    byte[] record = state.Encode(observation);
                    buffer.Write(record);
                    state.Appended(record.Length);
                }
                Directory.CreateDirectory(Path.GetDirectoryName(path)!);
                File.WriteAllBytes(path, buffer.ToArray());
                lock (gateLock) States[path] = state;
            }
            File.Delete(legacy);
            return true;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            // Locked or full: the old file is read as it is.
            return false;
        }
    }

    /// <summary>Reference ReadLegacy: the JSON-lines history, one malformed line at a time.</summary>
    private static List<PlayerObservation> ReadLegacy(string path)
    {
        var result = new List<PlayerObservation>();
        if (!File.Exists(path)) return result;
        foreach (var line in File.ReadAllLines(path))
        {
            if (string.IsNullOrWhiteSpace(line)) continue;
            try
            {
                var item = JsonSerializer.Deserialize<TrackerPersistedObservation>(line);
                if (item?.Observation is not null && item.Kind == "observation")
                    result.Add(item.Observation with { TimestampUtc = item.Observation.TimestampUtc.ToUniversalTime() });
            }
            catch (JsonException)
            {
                // One malformed JSONL line must not hide the remaining history.
            }
        }
        return result;
    }

    private static async Task WriteAllAsync(string path, byte[] header, byte[] record, CancellationToken ct)
    {
        using var stream = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.Read, 4096, useAsync: true);
        await stream.WriteAsync(header, ct).ConfigureAwait(false);
        await stream.WriteAsync(record, ct).ConfigureAwait(false);
        await stream.FlushAsync(ct).ConfigureAwait(false);
    }

    /// <summary>
    /// Appends to the file. A record that was torn by a crash is cut away first, so the file stays readable.
    /// </summary>
    private static async Task AppendAllAsync(string path, byte[] record, long validLength, CancellationToken ct)
    {
        using var stream = new FileStream(path, FileMode.OpenOrCreate, FileAccess.Write, FileShare.Read, 4096, useAsync: true);
        if (stream.Length > validLength) stream.SetLength(validLength);
        stream.Seek(0, SeekOrigin.End);
        await stream.WriteAsync(record, ct).ConfigureAwait(false);
        await stream.FlushAsync(ct).ConfigureAwait(false);
    }
}
