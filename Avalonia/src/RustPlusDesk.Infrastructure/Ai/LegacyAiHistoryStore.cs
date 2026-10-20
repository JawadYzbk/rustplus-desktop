using System.Text.Json;
using RustPlusDesk.Core.Ai;

namespace RustPlusDesk.Infrastructure.Ai;

/// <summary>
/// Reference AiHistory: the last few exchanges, kept in the app's data folder.
///
/// The file is the reference's own, so the WPF app and this one read each other's history. A file that is not there
/// and a file that cannot be read both give an empty history: a companion that has said nothing yet is not a broken
/// app.
/// </summary>
public sealed class LegacyAiHistoryStore(string path) : IAiHistoryStore
{
    /// <summary>The reference's file: %APPDATA%/RustPlusDesk/ai-history.json, a plain indented list.</summary>
    public static string DefaultPath => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "RustPlusDesk", "ai-history.json");

    private static readonly JsonSerializerOptions Options = new() { WriteIndented = true };

    private readonly List<AiExchange> entries = [];
    private bool loaded;

    /// <summary>The entries, newest first.</summary>
    public IReadOnlyList<AiExchange> Entries
    {
        get
        {
            Load();
            return entries;
        }
    }

    /// <summary>Adds one exchange at the front and keeps the newest twenty.</summary>
    public void Add(AiExchange exchange)
    {
        Load();
        entries.Insert(0, exchange);
        AiHistory.KeepNewest(entries);
        Save();
    }

    /// <summary>Empties the history.</summary>
    public void Clear()
    {
        Load();
        entries.Clear();
        Save();
    }

    /// <summary>Reads the file once; a file that is not there, or cannot be read, is an empty history.</summary>
    private void Load()
    {
        if (loaded) return;
        loaded = true;
        if (!File.Exists(path)) return;
        try
        {
            var read = JsonSerializer.Deserialize<List<AiExchange>>(File.ReadAllText(path)) ?? [];
            entries.Clear();
            entries.AddRange(read);
        }
        catch (Exception ex) when (ex is JsonException or IOException or UnauthorizedAccessException)
        {
            // A file that cannot be read is an empty history, not a broken app.
            entries.Clear();
        }
    }

    /// <summary>Writes the history back; a file that will not budge is not worth an exception.</summary>
    private void Save()
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path))!);
            File.WriteAllText(path, JsonSerializer.Serialize(entries, Options));
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException)
        {
            // A history that does not survive the session is a smaller problem than an exception thrown out of the
            // middle of an answer.
        }
    }
}
