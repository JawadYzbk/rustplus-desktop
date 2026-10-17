using System.Text.Json;
using RustPlusDesk.Core.Raid;

namespace RustPlusDesk.Infrastructure.Raid;

/// <summary>
/// Shares the reference app's raid-plan.json (same path and shape). An unreadable file loads as an empty
/// plan like the reference, but is copied to a backup before the first save would replace it.
/// </summary>
public sealed class RaidPlanStore(string path) : IRaidPlanStore
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };
    private readonly SemaphoreSlim writes = new(1, 1);
    private bool unreadable;

    public static string DefaultPath => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "RustPlusDesk", "raid-plan.json");

    public async Task<IReadOnlyList<RaidPlanEntry>> LoadAsync(CancellationToken cancellationToken)
    {
        if (!File.Exists(path)) return [];
        try
        {
            await using var stream = File.OpenRead(path);
            var entries = await JsonSerializer.DeserializeAsync<List<RaidPlanEntry>>(stream, JsonOptions, cancellationToken).ConfigureAwait(false) ?? [];
            return entries.Where(entry => entry is not null).ToList();
        }
        catch (Exception ex) when (ex is JsonException or IOException or UnauthorizedAccessException)
        {
            unreadable = true;
            return [];
        }
    }

    public async Task SaveAsync(IReadOnlyList<RaidPlanEntry> entries, CancellationToken cancellationToken)
    {
        await writes.WaitAsync(cancellationToken).ConfigureAwait(false);
        string temporary = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path))!);
            if (unreadable && File.Exists(path))
            {
                File.Copy(path, path + "." + Guid.NewGuid().ToString("N") + ".bak");
                unreadable = false;
            }
            await using (var stream = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None, 4096, FileOptions.Asynchronous))
                await JsonSerializer.SerializeAsync(stream, entries, JsonOptions, cancellationToken).ConfigureAwait(false);
            File.Move(temporary, path, overwrite: true);
        }
        finally
        {
            if (File.Exists(temporary)) File.Delete(temporary);
            writes.Release();
        }
    }
}
