using System.Text.Json;
using System.Text.Json.Nodes;
using RustPlusDesk.Core.Settings;

namespace RustPlusDesk.Infrastructure.Settings;

/// <summary>
/// Reads and writes the reference tracking_settings.json. Only keys changed by this app are written;
/// every save re-reads the file and applies just those keys, so settings the WPF app changed meanwhile
/// and keys this app does not know survive. An unreadable file is backed up before it is replaced.
/// </summary>
public sealed class TrackingSettingsStore : IAppSettingsStore, IReloadableSettingsStore
{
    private static readonly JsonSerializerOptions Indented = new() { WriteIndented = true };
    private readonly string path;
    private readonly TimeSpan debounce;
    private readonly object gate = new();
    private readonly SemaphoreSlim writes = new(1, 1);
    private readonly JsonObject document;
    private readonly Dictionary<string, JsonNode?> pending = new(StringComparer.Ordinal);
    private CancellationTokenSource? scheduled;
    private bool unreadable;

    public static string DefaultPath => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "RustPlusDesk", "tracking_settings.json");

    public TrackingSettingsStore(string path, TimeSpan? debounce = null)
    {
        this.path = path;
        this.debounce = debounce ?? TimeSpan.FromMilliseconds(300);
        document = Read(out unreadable) ?? new JsonObject();
    }

    private JsonObject? Read(out bool failed)
    {
        failed = false;
        try
        {
            if (!File.Exists(path)) return null;
            return JsonNode.Parse(File.ReadAllBytes(path)) as JsonObject ?? throw new JsonException();
        }
        catch (Exception ex) when (ex is JsonException or IOException or UnauthorizedAccessException or InvalidOperationException)
        {
            failed = true;
            return null;
        }
    }

    public T Get<T>(string name, T fallback)
    {
        lock (gate)
        {
            try { return document[name] is { } node ? node.Deserialize<T>() ?? fallback : fallback; }
            catch (Exception ex) when (ex is JsonException or InvalidOperationException or FormatException) { return fallback; }
        }
    }

    public void Set<T>(string name, T value)
    {
        lock (gate)
        {
            var node = JsonSerializer.SerializeToNode(value);
            if (JsonNode.DeepEquals(document[name], node)) return;
            document[name] = node;
            pending[name] = node?.DeepClone();
            scheduled?.Cancel();
            scheduled = new CancellationTokenSource();
            _ = SaveLaterAsync(scheduled.Token);
        }
    }

    private async Task SaveLaterAsync(CancellationToken token)
    {
        try { await Task.Delay(debounce, token).ConfigureAwait(false); }
        catch (OperationCanceledException) { return; }
        try { await FlushAsync(CancellationToken.None).ConfigureAwait(false); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { /* Retried on the next change or flush. */ }
    }

    /// <summary>
    /// Reference ReloadApplicationData's re-read of the settings: the file a restored backup replaced is the one the
    /// app reads from here on, while a change this app made and has not written yet is kept.
    /// </summary>
    public void Reload()
    {
        lock (gate)
        {
            var latest = Read(out bool failed);
            if (failed || latest is null) return;
            foreach (var pair in latest)
                if (!pending.ContainsKey(pair.Key)) document[pair.Key] = pair.Value?.DeepClone();
        }
    }

    public async Task FlushAsync(CancellationToken ct = default)
    {
        await writes.WaitAsync(ct).ConfigureAwait(false);
        try
        {
            Dictionary<string, JsonNode?> changes;
            lock (gate)
            {
                if (pending.Count == 0) return;
                changes = pending.ToDictionary(pair => pair.Key, pair => pair.Value?.DeepClone(), StringComparer.Ordinal);
            }
            Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path))!);
            var latest = Read(out bool failed);
            if ((failed || unreadable) && File.Exists(path))
                File.Copy(path, path + "." + Guid.NewGuid().ToString("N") + ".bak");
            unreadable = false;
            var merged = latest ?? new JsonObject();
            foreach (var (name, value) in changes) merged[name] = value?.DeepClone();
            string temporary = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
            try
            {
                await File.WriteAllTextAsync(temporary, merged.ToJsonString(Indented), ct).ConfigureAwait(false);
                File.Move(temporary, path, overwrite: true);
            }
            finally { if (File.Exists(temporary)) File.Delete(temporary); }
            lock (gate)
            {
                foreach (var (name, value) in changes)
                    if (pending.TryGetValue(name, out var current) && JsonNode.DeepEquals(current, value)) pending.Remove(name);
                // Adopt keys the other app changed, without losing local edits made during the write.
                foreach (var (name, value) in merged)
                    if (!pending.ContainsKey(name)) document[name] = value?.DeepClone();
            }
        }
        finally { writes.Release(); }
    }
}
