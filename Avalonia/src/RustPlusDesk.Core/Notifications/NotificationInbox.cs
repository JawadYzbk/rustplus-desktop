using RustPlusDesk.Core.Accounts;

namespace RustPlusDesk.Core.Notifications;

/// <summary>
/// One notification centre row, shaped like the reference RustPlusNotification so the shared
/// cache/notifications_history.json keeps working in both apps (same property names).
/// </summary>
public sealed class InboxNotification
{
    public string Id { get; set; } = Guid.NewGuid().ToString();
    public DateTime Timestamp { get; set; } = DateTime.Now;
    /// <summary>Alarm, Death, Pairing, OilRigAlarm (reference types).</summary>
    public string Type { get; set; } = "";
    public string Title { get; set; } = "";
    public string Message { get; set; } = "";
    public string? ServerIp { get; set; }
    public int? ServerPort { get; set; }
    public string? ServerName { get; set; }
    public bool IsRead { get; set; }
    public uint? EntityId { get; set; }
    public string? EntityName { get; set; }
    public string? AttackerName { get; set; }
    public string? ChatAuthor { get; set; }
    public string? FcmNotificationId { get; set; }
    public string? ImageUrl { get; set; }

    public string? ServerKey => !string.IsNullOrEmpty(ServerIp) && ServerPort.HasValue ? $"{ServerIp}:{ServerPort.Value}" : null;
}

public enum NotificationFilter { All, Alarms, Deaths }

/// <summary>Persists the reference notification history.</summary>
public interface INotificationHistoryStore
{
    Task<IReadOnlyList<InboxNotification>> LoadAsync(CancellationToken ct);
    Task SaveAsync(IReadOnlyList<InboxNotification> items, CancellationToken ct);
}

/// <summary>
/// The reference NotificationCenterService rules: newest first, retention in days (default 30), at most 500 rows,
/// muted servers dropped, de-duplication by push ID and by same type/message/server within four seconds, rows without
/// a picture borrow the server's last one, and old fixed titles reworded on load.
/// </summary>
public sealed class NotificationInbox
{
    public const int MaxItems = 500;
    public const int DefaultRetentionDays = 30;
    private readonly List<InboxNotification> items = [];

    public IReadOnlyList<InboxNotification> Items => items;
    public int UnreadCount => items.Count(item => !item.IsRead);

    /// <summary>Reference NotificationsRetentionDays: zero or less means the 30-day default.</summary>
    public static int Retention(int days) => days <= 0 ? DefaultRetentionDays : days;

    public void Load(IEnumerable<InboxNotification> saved, DateTime now, int retentionDays)
    {
        var cutoff = now.AddDays(-Retention(retentionDays));
        items.Clear();
        foreach (var item in saved.Where(item => item is not null && item.Timestamp >= cutoff).OrderByDescending(item => item.Timestamp))
        {
            Reword(item);
            items.Add(item);
        }
    }

    /// <summary>Reference AddNotification; false when the row was muted or a duplicate.</summary>
    public bool Add(InboxNotification notification, DateTime now, int retentionDays, ICollection<string> mutedServers)
    {
        if (notification.ServerKey is { } key && mutedServers.Contains(key)) return false;
        if (string.IsNullOrEmpty(notification.ImageUrl) && !string.IsNullOrEmpty(notification.ServerIp))
            notification.ImageUrl = items.FirstOrDefault(item => item.ServerIp == notification.ServerIp && item.ServerPort == notification.ServerPort
                && !string.IsNullOrEmpty(item.ImageUrl))?.ImageUrl;
        if (!string.IsNullOrEmpty(notification.FcmNotificationId) && items.Any(item => item.FcmNotificationId == notification.FcmNotificationId)) return false;
        if (items.Any(item => item.Type == notification.Type && item.Message == notification.Message && item.ServerIp == notification.ServerIp
                && item.ServerPort == notification.ServerPort && Math.Abs((item.Timestamp - notification.Timestamp).TotalSeconds) < 4)) return false;
        items.Insert(0, notification);
        var cutoff = now.AddDays(-Retention(retentionDays));
        items.RemoveAll(item => item.Timestamp < cutoff);
        if (items.Count > MaxItems) items.RemoveRange(MaxItems, items.Count - MaxItems);
        return true;
    }

    public bool MarkRead(string id)
    {
        var item = items.FirstOrDefault(entry => entry.Id == id);
        if (item is null || item.IsRead) return false;
        item.IsRead = true;
        return true;
    }

    public bool MarkAllRead()
    {
        bool changed = false;
        foreach (var item in items.Where(item => !item.IsRead)) { item.IsRead = true; changed = true; }
        return changed;
    }

    public bool Delete(string id) => items.RemoveAll(item => item.Id == id) > 0;
    public void Clear() => items.Clear();

    /// <summary>Reference view: text in title, message or server name, then the Alarms / Deaths category, newest first by time.</summary>
    /// <param name="matchServer">False in streamer mode: the reference search then ignores server names.</param>
    public IEnumerable<InboxNotification> Filter(string? search, NotificationFilter filter, bool matchServer = true)
    {
        string text = (search ?? "").Trim();
        return items.Where(item => text.Length == 0
                || item.Title.Contains(text, StringComparison.OrdinalIgnoreCase)
                || item.Message.Contains(text, StringComparison.OrdinalIgnoreCase)
                || (matchServer && item.ServerName?.Contains(text, StringComparison.OrdinalIgnoreCase) == true))
            .Where(item => filter switch
            {
                NotificationFilter.Alarms => item.Type.Equals("Alarm", StringComparison.OrdinalIgnoreCase),
                NotificationFilter.Deaths => item.Type.Equals("Death", StringComparison.OrdinalIgnoreCase),
                _ => true
            })
            .OrderByDescending(item => item.Timestamp);
    }

    /// <summary>Reference Reword: rows saved with the two old fixed titles get the live-map wording.</summary>
    public static void Reword(InboxNotification n)
    {
        if (n.Type == "Death" && n.Title == "Offline Death" && !string.IsNullOrWhiteSpace(n.AttackerName))
        {
            n.Title = $"You were killed by {n.AttackerName}";
            if (!string.IsNullOrWhiteSpace(n.ServerName)) n.Message = n.ServerName;
        }
        else if (n.Type == "Pairing" && n.Title == "Pairing Successful")
        {
            if (n.EntityId.HasValue)
            {
                n.Title = n.EntityName ?? "Smart Device";
                n.Message = $"Device paired (ID {n.EntityId.Value})";
            }
            else if (!string.IsNullOrWhiteSpace(n.ServerName))
            {
                n.Title = n.ServerName;
                n.Message = "Server paired.";
            }
        }
    }

    /// <summary>
    /// The reference rows for Rust+ pushes: alarms keep the push title (else "Alarm"), deaths read "You were killed by
    /// …" over the server, pairings show the paired thing over what happened. Other pushes are not inbox rows.
    /// </summary>
    public static InboxNotification? FromPush(CompanionUpdate update)
    {
        string? host = string.IsNullOrWhiteSpace(update.Host) ? null : update.Host;
        string? server = string.IsNullOrWhiteSpace(update.ServerName) ? null : update.ServerName;
        var row = new InboxNotification
        {
            Timestamp = update.OccurredAt.LocalDateTime, ServerIp = host, ServerPort = update.Port, ServerName = server, EntityId = update.EntityId
        };
        switch (update.Kind)
        {
            case CompanionUpdateKind.Alarm:
                row.Type = "Alarm";
                row.Title = string.IsNullOrWhiteSpace(update.Title) ? "Alarm" : update.Title;
                row.Message = update.Message;
                return row;
            case CompanionUpdateKind.Death:
                row.Type = "Death";
                row.AttackerName = update.Message;
                row.Title = $"You were killed by {update.Message}";
                row.Message = server ?? "";
                return row;
            case CompanionUpdateKind.Pairing:
                row.Type = "Pairing";
                if (update.EntityId is { } entity)
                {
                    row.EntityName = string.IsNullOrWhiteSpace(update.Title) ? "Smart Device" : update.Title;
                    row.Title = row.EntityName;
                    row.Message = $"Device paired (ID {entity})";
                }
                else
                {
                    row.Title = server ?? $"{update.Host}:{update.Port}";
                    row.Message = "Server paired.";
                }
                return row;
            default:
                return null;
        }
    }
}
