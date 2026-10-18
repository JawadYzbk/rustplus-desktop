using System.Globalization;
using System.Text;
using System.Text.Json;
using RustPlusDesk.Core.Deaths;

namespace RustPlusDesk.Infrastructure.Deaths;

/// <summary>
/// Reference DeathReporter local log: one JSON line per death in %APPDATA%/RustPlusDesk/deaths/{server}.jsonl, shared
/// with the WPF app (same folder, file names and fields). The pre-rename RustPlusDesktop/deaths folder is merged in
/// once, as the reference does.
/// </summary>
public sealed class DeathLogFile(string folder, string? legacyFolder = null) : IDeathLog
{
    public static string DefaultFolder => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "RustPlusDesk", "deaths");
    public static string DefaultLegacyFolder => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "RustPlusDesktop", "deaths");

    private readonly Lock gate = new();
    private bool migrated;

    public string Folder => folder;

    /// <summary>Reference LogPathFor: invalid file-name characters become "_", an empty key becomes "server".</summary>
    public string PathFor(string serverKey)
    {
        EnsureMigrated();
        return Path.Combine(folder, SafeFileName(serverKey) + ".jsonl");
    }

    public void Append(string serverKey, DeathRecord death)
    {
        try
        {
            string path = PathFor(serverKey);
            Directory.CreateDirectory(folder);
            // Same anonymous shape and default serializer options as the reference, so both apps write identical lines.
            string line = JsonSerializer.Serialize(new
            {
                steam_id = death.SteamId.ToString(CultureInfo.InvariantCulture),
                name = death.Name,
                died_at = death.DeathTime,
                spawn_at = death.SpawnTime,
                x = death.X,
                y = death.Y,
                grid = death.Grid,
                location_type = death.LocationType,
                location_name = death.LocationName,
            });
            lock (gate) File.AppendAllText(path, line + Environment.NewLine, Encoding.UTF8);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { }
    }

    public IReadOnlyList<string> ReadLines(string serverKey) => Read(PathFor(serverKey));

    public void Clear(string serverKey) => Delete(PathFor(serverKey));

    /// <summary>Reference KillerLogStore path: the death log's name with ".killers.jsonl".</summary>
    public string KillerPathFor(string serverKey) => PathFor(serverKey).Replace(".jsonl", ".killers.jsonl", StringComparison.Ordinal);

    public IReadOnlyList<string> ReadKillerLines(string serverKey) => Read(KillerPathFor(serverKey));

    public void WriteKillerLines(string serverKey, IEnumerable<string> lines)
    {
        try
        {
            string path = KillerPathFor(serverKey);
            if (File.Exists(path)) lock (gate) File.WriteAllLines(path, lines);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { }
    }

    public void ClearKillers(string serverKey) => Delete(KillerPathFor(serverKey));

    private IReadOnlyList<string> Read(string path)
    {
        try
        {
            lock (gate) return File.Exists(path) ? File.ReadAllLines(path, Encoding.UTF8) : [];
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { return []; }
    }

    private void Delete(string path)
    {
        try { lock (gate) File.Delete(path); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { }
    }

    /// <summary>Reference MigrateLegacyDeaths: move each legacy file, or merge new lines into an existing one.</summary>
    private void EnsureMigrated()
    {
        lock (gate)
        {
            if (migrated) return;
            migrated = true;
            if (legacyFolder is null || !Directory.Exists(legacyFolder)) return;
            try
            {
                Directory.CreateDirectory(folder);
                foreach (string file in Directory.GetFiles(legacyFolder, "*.jsonl"))
                {
                    string target = Path.Combine(folder, Path.GetFileName(file));
                    if (!File.Exists(target)) { File.Move(file, target); continue; }
                    var existing = new HashSet<string>(File.ReadAllLines(target, Encoding.UTF8));
                    using (var writer = File.AppendText(target))
                        foreach (string line in File.ReadAllLines(file, Encoding.UTF8))
                            if (!string.IsNullOrWhiteSpace(line) && existing.Add(line)) writer.WriteLine(line);
                    File.Delete(file);
                }
                if (Directory.GetFileSystemEntries(legacyFolder).Length == 0)
                {
                    Directory.Delete(legacyFolder, false);
                    if (Path.GetDirectoryName(legacyFolder) is { Length: > 0 } parent && Directory.Exists(parent) && Directory.GetFileSystemEntries(parent).Length == 0)
                        Directory.Delete(parent, false);
                }
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { }
        }
    }

    private static string SafeFileName(string value)
    {
        char[] invalid = Path.GetInvalidFileNameChars();
        string cleaned = new(value.Select(c => invalid.Contains(c) ? '_' : c).ToArray());
        return string.IsNullOrWhiteSpace(cleaned) ? "server" : cleaned;
    }
}
