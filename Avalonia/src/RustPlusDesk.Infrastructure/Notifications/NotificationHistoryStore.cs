using System.Text.Json;
using RustPlusDesk.Core.Notifications;

namespace RustPlusDesk.Infrastructure.Notifications;

/// <summary>
/// The reference cache/notifications_history.json (DataManager cache "notifications_history": a JSON list with the
/// RustPlusNotification property names). An unreadable file loads as empty, like the reference, but is copied to a
/// backup before the first save replaces it.
/// </summary>
public sealed class NotificationHistoryStore(string path) : INotificationHistoryStore
{
    private static readonly JsonSerializerOptions Options = new() { WriteIndented = true };
    private readonly SemaphoreSlim writes = new(1, 1);
    private bool unreadable;

    public static string DefaultPath => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "RustPlusDesk", "cache", "notifications_history.json");

    public async Task<IReadOnlyList<InboxNotification>> LoadAsync(CancellationToken ct)
    {
        if (!File.Exists(path)) return [];
        try
        {
            await using var stream = File.OpenRead(path);
            var items = await JsonSerializer.DeserializeAsync<List<InboxNotification?>>(stream, Options, ct).ConfigureAwait(false) ?? [];
            return items.OfType<InboxNotification>().ToList();
        }
        catch (Exception ex) when (ex is JsonException or IOException or UnauthorizedAccessException)
        {
            unreadable = true;
            return [];
        }
    }

    public async Task SaveAsync(IReadOnlyList<InboxNotification> items, CancellationToken ct)
    {
        await writes.WaitAsync(ct).ConfigureAwait(false);
        string temporary = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path))!);
            if (unreadable && File.Exists(path))
            {
                File.Copy(path, path + ".unreadable-" + DateTime.UtcNow.ToString("yyyyMMddHHmmss", System.Globalization.CultureInfo.InvariantCulture) + ".bak", overwrite: true);
                unreadable = false;
            }
            await using (var stream = File.Create(temporary))
                await JsonSerializer.SerializeAsync(stream, items, Options, ct).ConfigureAwait(false);
            File.Move(temporary, path, overwrite: true);
        }
        finally
        {
            if (File.Exists(temporary)) File.Delete(temporary);
            writes.Release();
        }
    }
}
