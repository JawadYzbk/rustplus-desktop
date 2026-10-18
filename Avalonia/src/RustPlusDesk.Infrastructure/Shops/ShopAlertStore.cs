using System.Text.Json;
using RustPlusDesk.Core.Map;

namespace RustPlusDesk.Infrastructure.Shops;

/// <summary>
/// The reference shop_alerts.json beside the executable: a JSON array of saved alerts (QueryText, MatchSellSide,
/// MatchBuySide, NotifyChat, NotifySound). An unreadable file loads as no alerts, like the reference, but is
/// copied to a backup before the first save replaces it.
/// </summary>
public sealed class ShopAlertStore(string path) : IShopAlertStore
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };
    private readonly SemaphoreSlim writes = new(1, 1);
    private bool unreadable;

    public static string DefaultPath => Path.Combine(AppContext.BaseDirectory, "shop_alerts.json");

    public async Task<IReadOnlyList<SavedShopAlert>> LoadAsync(CancellationToken ct)
    {
        if (!File.Exists(path)) return [];
        try
        {
            await using var stream = File.OpenRead(path);
            var alerts = await JsonSerializer.DeserializeAsync<List<SavedShopAlert?>>(stream, JsonOptions, ct).ConfigureAwait(false) ?? [];
            return alerts.Where(alert => alert is { QueryText.Length: > 0 }).Select(alert => alert!).ToList();
        }
        catch (Exception ex) when (ex is JsonException or IOException or UnauthorizedAccessException)
        {
            unreadable = true;
            return [];
        }
    }

    public async Task SaveAsync(IReadOnlyList<SavedShopAlert> alerts, CancellationToken ct)
    {
        await writes.WaitAsync(ct).ConfigureAwait(false);
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
                await JsonSerializer.SerializeAsync(stream, alerts, JsonOptions, ct).ConfigureAwait(false);
            File.Move(temporary, path, overwrite: true);
        }
        finally
        {
            if (File.Exists(temporary)) File.Delete(temporary);
            writes.Release();
        }
    }
}
