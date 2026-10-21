using System.Text.Json;
using System.Text.Json.Serialization;
using RustPlusDesk.Core.Tutorials;

namespace RustPlusDesk.Infrastructure.Tutorials;

/// <summary>
/// Reference TutorialProgressStore: <c>%APPDATA%\RustPlusDesk\tutorial-progress.json</c>, the same file the WPF app
/// reads and writes. The file keeps every tutorial's progress under <c>Tutorials</c> and the switches under
/// <c>Preferences</c>.
///
/// Unknown members survive a round-trip, both at the root and inside each tutorial and the preferences, so a newer
/// build can add fields without this one dropping them. A file that cannot be read is copied to a timestamped backup
/// before the first save replaces it, and every write is made through a temporary file so a crash cannot leave a
/// half-written store.
/// </summary>
public sealed class LegacyTutorialProgressStore(string? path = null) : ITutorialProgressStore
{
    private sealed class StoreData
    {
        public Dictionary<string, TutorialProgress> Tutorials { get; set; } = new(StringComparer.Ordinal);
        public TutorialPreferences Preferences { get; set; } = new();

        [JsonExtensionData]
        public Dictionary<string, JsonElement>? AdditionalData { get; set; }
    }

    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };
    private readonly string _path = path ?? DefaultPath;
    private readonly SemaphoreSlim _gate = new(1, 1);
    private bool _unreadable;

    public static string DefaultPath => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "RustPlusDesk", "tutorial-progress.json");

    public async Task<TutorialProgress> GetAsync(TutorialDefinition definition, CancellationToken cancellationToken = default)
    {
        await _gate.WaitAsync(cancellationToken);
        try
        {
            var data = await ReadAsync(cancellationToken);
            if (!data.Tutorials.TryGetValue(definition.Id, out var progress))
                return NewProgress(definition);

            // Reference: a finished tour whose chapter has been rewritten is offered again as "Updated".
            if (progress.TutorialVersion < definition.Version && progress.Status is TutorialStatus.Completed or TutorialStatus.Skipped)
                progress.Status = TutorialStatus.Updated;
            progress.TutorialVersion = definition.Version;
            progress.TutorialId = definition.Id;
            return progress;
        }
        finally { _gate.Release(); }
    }

    public async Task<IReadOnlyDictionary<string, TutorialProgress>> GetAllAsync(CancellationToken cancellationToken = default)
    {
        await _gate.WaitAsync(cancellationToken);
        try { return (await ReadAsync(cancellationToken)).Tutorials; }
        finally { _gate.Release(); }
    }

    public async Task SaveAsync(TutorialProgress progress, CancellationToken cancellationToken = default)
    {
        await _gate.WaitAsync(cancellationToken);
        try
        {
            var data = await ReadAsync(cancellationToken);
            data.Tutorials[progress.TutorialId] = progress;
            await WriteAsync(data, cancellationToken);
        }
        finally { _gate.Release(); }
    }

    public async Task ResetAsync(string tutorialId, CancellationToken cancellationToken = default)
    {
        await _gate.WaitAsync(cancellationToken);
        try
        {
            var data = await ReadAsync(cancellationToken);
            data.Tutorials.Remove(tutorialId);
            await WriteAsync(data, cancellationToken);
        }
        finally { _gate.Release(); }
    }

    public async Task ResetAllAsync(CancellationToken cancellationToken = default)
    {
        await _gate.WaitAsync(cancellationToken);
        try
        {
            var data = await ReadAsync(cancellationToken);
            data.Tutorials.Clear();
            data.Preferences.LastTutorialId = null;
            await WriteAsync(data, cancellationToken);
        }
        finally { _gate.Release(); }
    }

    public async Task<TutorialPreferences> GetPreferencesAsync(CancellationToken cancellationToken = default)
    {
        await _gate.WaitAsync(cancellationToken);
        try { return (await ReadAsync(cancellationToken)).Preferences; }
        finally { _gate.Release(); }
    }

    public async Task SavePreferencesAsync(TutorialPreferences preferences, CancellationToken cancellationToken = default)
    {
        await _gate.WaitAsync(cancellationToken);
        try
        {
            var data = await ReadAsync(cancellationToken);
            data.Preferences = preferences;
            await WriteAsync(data, cancellationToken);
        }
        finally { _gate.Release(); }
    }

    private static TutorialProgress NewProgress(TutorialDefinition definition) => new()
    {
        TutorialId = definition.Id,
        TutorialVersion = definition.Version,
        Status = TutorialStatus.NotStarted
    };

    private async Task<StoreData> ReadAsync(CancellationToken cancellationToken)
    {
        if (!File.Exists(_path)) return new();
        try
        {
            await using var stream = File.OpenRead(_path);
            return await JsonSerializer.DeserializeAsync<StoreData>(stream, JsonOptions, cancellationToken).ConfigureAwait(false) ?? new();
        }
        catch (Exception ex) when (ex is JsonException or IOException or UnauthorizedAccessException)
        {
            _unreadable = true;
            return new();
        }
    }

    private async Task WriteAsync(StoreData data, CancellationToken cancellationToken)
    {
        string full = Path.GetFullPath(_path);
        string? directory = Path.GetDirectoryName(full);
        if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);

        // A file that could not be read is kept beside the store before it is replaced, the way the reference's other
        // stores keep one; the user's older progress is then still recoverable by hand.
        if (_unreadable && File.Exists(_path))
        {
            File.Copy(_path, _path + ".unreadable-" + DateTime.UtcNow.ToString("yyyyMMddHHmmss", System.Globalization.CultureInfo.InvariantCulture) + ".bak", overwrite: true);
            _unreadable = false;
        }

        string temporary = _path + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try
        {
            await using (var stream = File.Create(temporary))
                await JsonSerializer.SerializeAsync(stream, data, JsonOptions, cancellationToken).ConfigureAwait(false);
            File.Move(temporary, _path, overwrite: true);
        }
        finally
        {
            if (File.Exists(temporary)) File.Delete(temporary);
        }
    }
}
