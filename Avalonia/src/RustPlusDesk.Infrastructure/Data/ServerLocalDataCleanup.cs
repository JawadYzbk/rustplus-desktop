using System.Text.Json;

namespace RustPlusDesk.Infrastructure.Data;

/// <summary>
/// Reference Services/Data/ServerLocalDataCleanup: removes what the app stored on disk for a deleted server profile — its
/// map caches and 3D map data, chat history, deaths, wipe tracking, drawings and API caches. Most of it is keyed by host
/// and port, so it is only removed when no remaining profile points at the same server. Offline maps (port 0 or the
/// "offline" token) only lose their own 3D map folder. Every removal is best effort.
/// </summary>
public static class ServerLocalDataCleanup
{
    /// <summary>%APPDATA%/RustPlusDesk (reference DataManager.AppDir).</summary>
    public static string DefaultAppDir => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "RustPlusDesk");
    /// <summary>LocalAppData/RustPlusDesk.</summary>
    public static string DefaultLocalDir => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "RustPlusDesk");

    // Windows' Path.GetInvalidFileNameChars, so keys match the reference on every platform.
    private static readonly char[] Invalid = [.. Enumerable.Range(0, 32).Select(code => (char)code), '"', '<', '>', '|', ':', '*', '?', '\\', '/'];

    /// <summary>
    /// Deletes <paramref name="ownMapFolder"/> and, unless <paramref name="keepServerData"/>, everything stored under the
    /// server's host and port. Returns how many files and folders were removed.
    /// </summary>
    public static int DeleteFor(string host, int port, bool keepServerData, string? ownMapFolder, string appDir, string localDir)
    {
        int removed = 0;
        if (!string.IsNullOrWhiteSpace(ownMapFolder)) removed += DeleteDirectory(ownMapFolder);
        if (keepServerData || string.IsNullOrWhiteSpace(host)) return removed;

        string dashKey = Safe($"{host}-{port}");
        string underscoreKey = Safe($"{host}_{port}");

        removed += DeleteFile(Path.Combine(appDir, "deaths", dashKey + ".jsonl"));
        removed += DeleteFile(Path.Combine(appDir, "deaths", dashKey + ".killers.jsonl"));
        removed += DeleteDirectory(Path.Combine(appDir, "player-wipes", dashKey));
        removed += DeleteDirectory(Path.Combine(appDir, "Overlays", dashKey));
        removed += DeleteMatching(Path.Combine(appDir, "cache"), underscoreKey + "_*.json");

        removed += DeleteFile(Path.Combine(localDir, "chat", underscoreKey + ".json"));
        removed += DeleteFile(Path.Combine(localDir, "chat", underscoreKey + "_clan.json"));
        removed += DeleteFile(Path.Combine(localDir, "map_cache", underscoreKey + ".json"));
        removed += DeleteFile(Path.Combine(localDir, "map_cache", underscoreKey + ".png"));
        removed += DeleteFile(Path.Combine(localDir, "map_cache", underscoreKey + "_custom.png"));

        // 3D map folders are named after the profile and wipe, so earlier wipes are found by their manifest's server.
        string mapsRoot = Path.Combine(appDir, "3DMaps");
        if (Directory.Exists(mapsRoot))
            foreach (string folder in Directory.EnumerateDirectories(mapsRoot))
                if (ManifestIsForServer(Path.Combine(folder, "map3d_manifest.json"), host, port)) removed += DeleteDirectory(folder);
        return removed;
    }

    /// <summary>Reference Map3DLocalBuildService.GetPreparedFolderPath: 3DMaps/{safe name}_{sha256(host:port:mapId:wipe)[..12]}.</summary>
    public static string PreparedMapFolder(string appDir, string? name, string host, int port, string? rustMapsMapId, DateTime? wipeTime)
    {
        string wipeKey = wipeTime?.ToUniversalTime().ToString("O", System.Globalization.CultureInfo.InvariantCulture) ?? "unknown";
        string raw = $"{host}:{port}:{rustMapsMapId ?? "unknown"}:{wipeKey}".ToLowerInvariant();
        string hash = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(System.Text.Encoding.UTF8.GetBytes(raw))).ToLowerInvariant()[..12];
        string display = string.IsNullOrWhiteSpace(name) ? host : name;
        string safeName = new string(display.Select(ch => char.IsLetterOrDigit(ch) ? ch : '_').ToArray()).Trim('_');
        if (string.IsNullOrWhiteSpace(safeName)) safeName = "server";
        return Path.Combine(appDir, "3DMaps", $"{safeName}_{hash}");
    }

    private static bool ManifestIsForServer(string manifestPath, string host, int port)
    {
        if (!File.Exists(manifestPath)) return false;
        try
        {
            using var doc = JsonDocument.Parse(File.ReadAllText(manifestPath));
            var root = doc.RootElement;
            string? manifestHost = root.TryGetProperty("Host", out var hostEl) && hostEl.ValueKind == JsonValueKind.String ? hostEl.GetString() : null;
            int manifestPort = root.TryGetProperty("Port", out var portEl) && portEl.ValueKind == JsonValueKind.Number ? portEl.GetInt32() : -1;
            return manifestPort == port && string.Equals(manifestHost, host, StringComparison.OrdinalIgnoreCase);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException or FormatException or InvalidOperationException) { return false; }
    }

    private static int DeleteMatching(string directory, string pattern)
    {
        if (!Directory.Exists(directory)) return 0;
        int removed = 0;
        foreach (string file in Directory.EnumerateFiles(directory, pattern)) removed += DeleteFile(file);
        return removed;
    }

    private static int DeleteFile(string path)
    {
        try
        {
            if (!File.Exists(path)) return 0;
            File.Delete(path);
            return 1;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { return 0; }
    }

    private static int DeleteDirectory(string path)
    {
        try
        {
            if (!Directory.Exists(path)) return 0;
            Directory.Delete(path, recursive: true);
            return 1;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { return 0; }
    }

    private static string Safe(string value) => new(value.Select(ch => Invalid.Contains(ch) ? '_' : ch).ToArray());
}
