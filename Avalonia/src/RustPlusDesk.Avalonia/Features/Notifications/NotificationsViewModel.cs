using System.Globalization;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Avalonia.Shared.Icons;
using RustPlusDesk.Core.Accounts;
using RustPlusDesk.Core.Notifications;
using RustPlusDesk.Core.Settings;

namespace RustPlusDesk.Avalonia.Features.Notifications;

/// <summary>One inbox row as shown: type icon, unread dot, title, message, server tag and time.</summary>
public sealed partial class NotificationRow : ObservableObject
{
    public NotificationRow(InboxNotification item, bool hideServer = false)
    {
        Item = item;
        HideServer = hideServer;
    }

    /// <summary>The notification this row shows.</summary>
    public InboxNotification Item { get; }

    /// <summary>Whether the server behind the notification is hidden (streamer mode).</summary>
    public bool HideServer { get; }

    public string Title => Item.Title;
    public string Message => Item.Message;
    public bool IsUnread => !Item.IsRead;
    public bool HasServer => Item.ServerKey is not null;
    /// <summary>Server name, else "host:port"; streamer mode shows the reference "SERVER HIDDEN" instead.</summary>
    public string ServerTag => HideServer ? StreamerMode.HiddenServer : Item.ServerName is { Length: > 0 } name ? name : Item.ServerKey ?? "";
    public string Time => Item.Timestamp.ToString("g", CultureInfo.CurrentCulture);
    /// <summary>Symbol per notification type.</summary>
    public FluentSymbol Symbol => Item.Type switch
    {
        "Alarm" => FluentSymbol.AlertUrgent24,
        "Death" => FluentSymbol.HeartBroken24,
        "Pairing" => FluentSymbol.PlugConnected24,
        "OilRigAlarm" => FluentSymbol.Fire24,
        _ => FluentSymbol.Info24
    };
    public bool IsAlarm => Item.Type is "Alarm" or "OilRigAlarm";
    public bool IsDeath => Item.Type == "Death";

    /// <summary>
    /// The server's own picture, when the push carried one.
    ///
    /// The reference's notification centre shows the picture the push sent with the server's logo, and falls back to the
    /// type's glyph when there is none.
    /// </summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasPicture))]
    private global::Avalonia.Media.Imaging.Bitmap? picture;

    /// <summary>Whether the row has a picture to show rather than the glyph.</summary>
    public bool HasPicture => Picture is not null;
}
/// <summary>
/// The reference Notifications tab (NotificationsTabContent + NotificationCenterService): durable history in
/// cache/notifications_history.json, unread count, search, All/Alarms/Deaths filter, click to mark read, mark all read,
/// clear (confirmed), delete, mute a server (confirmed; MutedNotificationServers in tracking_settings.json) and the
/// server tag that selects and connects to that paired server.
/// </summary>
public sealed partial class NotificationsViewModel : ObservableObject
{
    private readonly NotificationInbox inbox = new();
    private readonly INotificationHistoryStore? store;
    private readonly IAppSettingsStore? settings;
    private readonly Func<DateTime> now;
    private Task pendingSave = Task.CompletedTask;

    public NotificationsViewModel(INotificationHistoryStore? store, IAppSettingsStore? settings, Func<DateTime>? now = null)
    {
        this.store = store;
        this.settings = settings;
        this.now = now ?? (() => DateTime.Now);
    }

    [ObservableProperty] private IReadOnlyList<NotificationRow> rows = [];
    [ObservableProperty] private string searchText = "";
    [ObservableProperty] private NotificationFilter filter = NotificationFilter.All;
    [ObservableProperty] private int unreadCount;
    [ObservableProperty] private bool isConfirmingClear;
    [ObservableProperty] private NotificationRow? muteCandidate;
    [ObservableProperty] private string notice = "";
    /// <summary>Reference IsStreamerMode: server tags read "SERVER HIDDEN" and search ignores server names.</summary>
    [ObservableProperty] private bool isStreamerMode;
    partial void OnIsStreamerModeChanged(bool value) => Refresh();

    public bool HasRows => Rows.Count > 0;
    public bool HasUnread => UnreadCount > 0;
    public string UnreadText => UnreadCount > 99 ? "99+" : UnreadCount.ToString(CultureInfo.InvariantCulture);
    public bool IsMuteConfirmOpen => MuteCandidate is not null;
    public string MuteQuestion => MuteCandidate is { } row ? $"Mute all future Notification Center alerts from {(row.HideServer ? "this server" : row.ServerTag)}? You can unmute it in Settings > Alerts." : "";
    public bool ShowAll => Filter == NotificationFilter.All;
    public bool ShowAlarms => Filter == NotificationFilter.Alarms;
    public bool ShowDeaths => Filter == NotificationFilter.Deaths;

    [RelayCommand] private void SetFilter(NotificationFilter value) => Filter = value;

    /// <summary>Selects and connects to a paired server by host and port; false when none matches.</summary>
    public Func<string, int, bool>? ConnectToServer { get; set; }

    partial void OnRowsChanged(IReadOnlyList<NotificationRow> value) => OnPropertyChanged(nameof(HasRows));
    partial void OnSearchTextChanged(string value) => Refresh();
    partial void OnFilterChanged(NotificationFilter value)
    {
        OnPropertyChanged(nameof(ShowAll));
        OnPropertyChanged(nameof(ShowAlarms));
        OnPropertyChanged(nameof(ShowDeaths));
        Refresh();
    }
    partial void OnUnreadCountChanged(int value)
    {
        OnPropertyChanged(nameof(HasUnread));
        OnPropertyChanged(nameof(UnreadText));
    }
    partial void OnMuteCandidateChanged(NotificationRow? value)
    {
        OnPropertyChanged(nameof(IsMuteConfirmOpen));
        OnPropertyChanged(nameof(MuteQuestion));
    }

    private int RetentionDays => settings?.Get(SettingKeys.NotificationsRetentionDays, NotificationInbox.DefaultRetentionDays) ?? NotificationInbox.DefaultRetentionDays;

    public async Task LoadAsync(CancellationToken ct)
    {
        if (store is null) return;
        inbox.Load(await store.LoadAsync(ct), now(), RetentionDays);
        Refresh();
    }

    /// <summary>Adds a Rust+ push (alarm, death, pairing) like the reference handlers.</summary>
    public bool Handle(CompanionUpdate update) => NotificationInbox.FromPush(update) is { } row && Add(row);

    public bool Add(InboxNotification notification)
    {
        var muted = settings?.Get<List<string>>(SettingKeys.MutedNotificationServers, []) ?? [];
        if (!inbox.Add(notification, now(), RetentionDays, muted)) return false;
        Changed();
        Added?.Invoke(notification);
        return true;
    }

    /// <summary>Reference NotificationAdded: raised for every row that made it into the inbox.</summary>
    public event Action<InboxNotification>? Added;

    /// <summary>Reference row click: marks it read.</summary>
    [RelayCommand]
    private void Open(NotificationRow? row)
    {
        if (row is not null && inbox.MarkRead(row.Item.Id)) Changed();
    }

    [RelayCommand]
    private void MarkAllRead()
    {
        if (inbox.MarkAllRead()) Changed();
    }

    [RelayCommand] private void RequestClear() => IsConfirmingClear = inbox.Items.Count > 0;
    [RelayCommand] private void CancelClear() => IsConfirmingClear = false;

    [RelayCommand]
    private void ConfirmClear()
    {
        IsConfirmingClear = false;
        inbox.Clear();
        Changed();
    }

    [RelayCommand]
    private void Delete(NotificationRow? row)
    {
        if (row is not null && inbox.Delete(row.Item.Id)) Changed();
    }

    [RelayCommand] private void RequestMute(NotificationRow? row) => MuteCandidate = row?.HasServer == true ? row : null;
    [RelayCommand] private void CancelMute() => MuteCandidate = null;

    /// <summary>Reference TrackingService.MuteServer: adds "host:port" once and remembers the server's name.</summary>
    [RelayCommand]
    private void ConfirmMute()
    {
        if (MuteCandidate is not { Item.ServerKey: { } key } row) return;
        MuteCandidate = null;
        if (settings is not null)
        {
            var muted = settings.Get<List<string>>(SettingKeys.MutedNotificationServers, []) ?? [];
            var names = settings.Get<Dictionary<string, string>>(SettingKeys.MutedNotificationServerNames, []) ?? [];
            if (!string.IsNullOrWhiteSpace(row.Item.ServerName))
            {
                names[key] = row.Item.ServerName;
                settings.Set(SettingKeys.MutedNotificationServerNames, names);
            }
            if (!muted.Contains(key))
            {
                muted.Add(key);
                settings.Set(SettingKeys.MutedNotificationServers, muted);
            }
        }
        Notice = $"Notifications from {(row.HideServer ? "this server" : row.ServerTag)} are now muted.";
    }

    /// <summary>Reference server tag: select that paired server and connect.</summary>
    [RelayCommand]
    private void ConnectServer(NotificationRow? row)
    {
        if (row?.Item is not { ServerIp: { } host, ServerPort: { } port }) return;
        Notice = ConnectToServer?.Invoke(host, port) == true ? "" : $"Could not find a paired server profile for {host}:{port}. Please pair this server first.";
    }

    private void Changed()
    {
        Refresh();
        QueueSave();
    }

    private void Refresh()
    {
        Rows = inbox.Filter(SearchText, Filter, !IsStreamerMode).Select(item => new NotificationRow(item, IsStreamerMode)).ToList();
        LoadPictures();
        UnreadCount = inbox.UnreadCount;
    }

    /// <summary>
    /// The server's own picture, for the rows that have one.
    ///
    /// The pictures are fetched in the background and put on the rows as they arrive, so the list is never held up by a
    /// picture that is slow to come.
    /// </summary>
    private void LoadPictures()
    {
        foreach (var row in Rows)
        {
            if (row.Item.ImageUrl is not { Length: > 0 } url) continue;

            _ = LoadPictureAsync(row, url);
        }
    }

    private static async Task LoadPictureAsync(NotificationRow row, string url)
    {
        var picture = await NotificationPictures.LoadAsync(url).ConfigureAwait(true);
        if (picture is null) return;

        row.Picture = picture;
    }

    private void QueueSave()
    {
        if (store is null) return;
        var snapshot = inbox.Items.ToList();
        pendingSave = SaveAfterAsync(pendingSave, snapshot);
    }

    private async Task SaveAfterAsync(Task previous, IReadOnlyList<InboxNotification> snapshot)
    {
        await previous;
        try { await store!.SaveAsync(snapshot, CancellationToken.None); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { Notice = "Notification history could not be saved. Check file permissions and disk space."; }
    }

    public Task FlushAsync() => pendingSave;
}
