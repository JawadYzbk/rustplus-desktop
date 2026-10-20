using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Core.Localization;
using RustPlusDesk.Core.Settings;
using RustPlusDesk.Core.Tracking;

namespace RustPlusDesk.Avalonia.Features.Players;

/// <summary>
/// One tracked player as the list shows them (reference TrackedPlayer). The name and the group are edited in place,
/// so the row keeps what the user is typing while the list is rebuilt.
/// </summary>
public sealed partial class TrackedPlayerRow : ObservableObject
{
    public TrackedPlayerRow(TrackedPlayer player)
    {
        Player = player;
        name = player.Name;
        groupName = TrackedPlayerRules.GroupName(player.GroupName);
    }

    public TrackedPlayer Player { get; private set; }

    /// <summary>Reference TrackedPlayer.BMId: the profile id the player is known by.</summary>
    public string BMId => Player.BMId;

    /// <summary>Reference TrackedPlayer.Name: the name the user gave the player.</summary>
    [ObservableProperty] private string name;

    /// <summary>Reference TrackedPlayer.GroupName: the group the player is in.</summary>
    [ObservableProperty] private string groupName;

    /// <summary>Reference BtnRenameTracked_Click: the name is edited in place.</summary>
    [ObservableProperty] private bool isRenaming;

    /// <summary>Reference BtnGroupTracked_Click: the group is chosen in place.</summary>
    [ObservableProperty] private bool isGrouping;

    /// <summary>Reference ShowBulkGroupEditorDialog: the player is checked in the group manager.</summary>
    [ObservableProperty] private bool isSelected;

    /// <summary>Reference TrackedPlayer.IsOnline: the last session has not ended yet.</summary>
    public bool IsOnline => Player.IsOnline;

    /// <summary>Reference TrackedPlayer.PlayTimeStr: how long the player has been on.</summary>
    public string PlayTime => Player.PlayTimeStr(DateTime.UtcNow);

    /// <summary>Reference TrackedPlayer.GroupColor: the colour of the group the player is in.</summary>
    public string GroupColor => string.IsNullOrEmpty(Player.GroupColor) ? "None" : Player.GroupColor;

    /// <summary>The player's initial, for the row's avatar.</summary>
    public string Initial => Name.Length > 0 ? Name[..1].ToUpperInvariant() : "?";

    /// <summary>Takes the player as it is now; the edited name and group are kept while the row is being edited.</summary>
    public void Update(TrackedPlayer player)
    {
        Player = player;
        if (!IsRenaming) Name = player.Name;
        if (!IsGrouping) GroupName = TrackedPlayerRules.GroupName(player.GroupName);
    }
}

/// <summary>One group of tracked players, as the list shows it (reference the group headers).</summary>
public sealed record TrackedPlayerGroupRow(string ServerName, string GroupName, string GroupColor, bool IsExpanded, IReadOnlyList<TrackedPlayerRow> Players)
{
    /// <summary>Reference TrackingService.GetGroupState: the state is kept per server and group.</summary>
    public string Key => $"{ServerName}|{GroupName}";

    /// <summary>
    /// Reference RefreshTrackedPlayersList: the group header carries the colour of the group. The reference stores the
    /// colour as the text the user picked, so it is read the same way here.
    /// </summary>
    public global::Avalonia.Media.IBrush? Colour => ParseColour(GroupColor);

    private static global::Avalonia.Media.IBrush? ParseColour(string colour)
    {
        if (string.IsNullOrEmpty(colour) || colour == "None") return null;
        try { return new global::Avalonia.Media.SolidColorBrush(global::Avalonia.Media.Color.Parse(colour)); }
        catch (FormatException) { return null; }
    }
}

/// <summary>One server's groups, as the list shows them.</summary>
public sealed record TrackedPlayerServerRow(string Name, IReadOnlyList<TrackedPlayerGroupRow> Groups);

/// <summary>
/// Reference MainWindow.Players: the players the user keeps an eye on.
///
/// The reference collected how long a player stays on a server; that half was removed, so what is left is the
/// bookmark: the players the user added, the groups they are in and the order the groups are in. The file the
/// reference keeps them in is read and written as it is, so both apps see the same bookmarks.
/// </summary>
public sealed partial class PlayersViewModel : ObservableObject
{
    private readonly ITrackedPlayerStore store;
    private readonly IAppSettingsStore? settings;
    private readonly IPlayerNameSource? names;
    private readonly Dictionary<string, TrackedPlayerRow> rows = new(StringComparer.Ordinal);
    private IReadOnlyList<TrackedPlayer> players = [];
    private string serverName = "";
    private bool isConnected;

    public PlayersViewModel(ITrackedPlayerStore store, IAppSettingsStore? settings = null, IPlayerNameSource? names = null)
    {
        this.store = store;
        this.settings = settings;
        this.names = names;
    }

    /// <summary>
    /// Reference TrackingService.StatusMessage: the reference says the player tracking was removed and points at the
    /// wipe tracker, which is where the team history is.
    /// </summary>
    public static string TrackingRemovedMessage => Loc.Text("PlayerActivityTrackingRemoved",
        "Player activity tracking has been removed. Team history is available in Wipe Tracker.");

    /// <summary>Reference TxtTrackedFilter_TextChanged: the filter the list is shown with.</summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(Servers))]
    [NotifyPropertyChangedFor(nameof(HasNoMatches))]
    private string filter = "";

    /// <summary>Reference TxtManualBMId: the name or the profile id the user types in.</summary>
    [ObservableProperty] private string manualId = "";

    /// <summary>
    /// The reference's Players tab has an Online half and a Tracked half; this is which one is shown.
    /// </summary>
    [ObservableProperty] private bool isTrackedTab;

    /// <summary>Reference TxtOnlinePlayersStatus: what the online half says.</summary>
    [ObservableProperty] private string status = TrackingRemovedMessage;

    /// <summary>The tracked players, grouped the way the reference shows them.</summary>
    public IReadOnlyList<TrackedPlayerServerRow> Servers => TrackedPlayerRules.Group(Filtered, GroupOrder)
        .Select(server => new TrackedPlayerServerRow(server.Name, server.Groups
            .Select(group => new TrackedPlayerGroupRow(server.Name, group.GroupName, group.GroupColor,
                IsExpanded(server.Name, group.GroupName),
                group.Players.Select(Row).ToList()))
            .ToList()))
        .ToList();

    /// <summary>Reference RefreshTrackedPlayersList: no results for the filter.</summary>
    public bool HasNoMatches => Filtered.Count == 0 && !string.IsNullOrWhiteSpace(Filter);

    /// <summary>Reference BtnManageGroups_Click: the groups the user has.</summary>
    public IReadOnlyList<string> Groups => players
        .Select(player => TrackedPlayerRules.GroupName(player.GroupName))
        .Distinct(StringComparer.Ordinal)
        .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
        .ToArray();

    /// <summary>The players the filter keeps.</summary>
    private IReadOnlyList<TrackedPlayer> Filtered => TrackedPlayerRules.Filter(players, Filter);

    /// <summary>Reference OpenPlayersWorkspace: the workspace opens on the server that is selected.</summary>
    public void Open(string? selectedServerName, bool connected)
    {
        serverName = selectedServerName ?? "";
        isConnected = connected;
        Status = connected
            ? TrackingRemovedMessage
            : Loc.Text("ConnectToLoadPlayers", "Connect to a server to load players list");
    }

    /// <summary>
    /// Reference BtnShowOnline_Click: the online half is read again. The reference asks Steam for the players of the
    /// server and then says that the tracking was removed; the port says the same thing.
    /// </summary>
    [RelayCommand]
    private void RefreshOnline()
    {
        Status = isConnected
            ? TrackingRemovedMessage
            : Loc.Text("ConnectToLoadPlayers", "Connect to a server to load players list");
    }

    /// <summary>Reference TrackingService.GetTrackedPlayers: the bookmarks are read from the file.</summary>
    public async Task LoadAsync(CancellationToken ct = default)
    {
        players = await store.LoadAsync(ct).ConfigureAwait(false);
        Refresh();
    }

    /// <summary>Reference BtnAddManual_Click: the player the user typed in is added.</summary>
    [RelayCommand]
    private async Task AddManualAsync()
    {
        string id = ManualId.Trim();
        if (id.Length == 0) return;
        string name = names is null ? id : await names.ResolveAsync(id, CancellationToken.None).ConfigureAwait(false);
        var player = new TrackedPlayer(id, name, serverName, "", "", [], IsBMOnly: true);
        players = players.Where(existing => existing.BMId != id).Append(player).ToList();
        await SaveAsync().ConfigureAwait(false);
        ManualId = "";
        Refresh();
    }

    /// <summary>Reference BtnRemoveTracked_Click: the player is no longer tracked.</summary>
    [RelayCommand]
    private async Task RemoveAsync(TrackedPlayerRow row)
    {
        players = players.Where(existing => existing.BMId != row.BMId).ToList();
        rows.Remove(row.BMId);
        await SaveAsync().ConfigureAwait(false);
        Refresh();
    }

    /// <summary>Reference BtnRenameTracked_Click: the name is edited in place.</summary>
    [RelayCommand]
    private void EditName(TrackedPlayerRow row) => row.IsRenaming = true;

    /// <summary>Reference BtnGroupTracked_Click: the group is chosen in place.</summary>
    [RelayCommand]
    private void EditGroup(TrackedPlayerRow row) => row.IsGrouping = true;

    /// <summary>Reference BtnRenameTracked_Click: the player is given another name.</summary>
    [RelayCommand]
    private async Task RenameAsync(TrackedPlayerRow row)
    {
        string name = row.Name.Trim();
        if (name.Length == 0) return;
        await ChangeAsync(row, player => player with { Name = name }).ConfigureAwait(false);
        row.IsRenaming = false;
    }

    /// <summary>Reference BtnGroupTracked_Click: the player is put in a group, which may be a new one.</summary>
    [RelayCommand]
    private async Task SetGroupAsync(TrackedPlayerRow row)
    {
        string group = row.GroupName.Trim();
        await ChangeAsync(row, player => player with { GroupName = group == TrackedPlayerRules.UngroupedName ? "" : group }).ConfigureAwait(false);
        row.IsGrouping = false;
    }

    /// <summary>Keeps what the user edited, whether it was the name or the group.</summary>
    [RelayCommand]
    private async Task SaveAsync(TrackedPlayerRow row)
    {
        if (row.IsRenaming) await RenameAsync(row).ConfigureAwait(false);
        else if (row.IsGrouping) await SetGroupAsync(row).ConfigureAwait(false);
    }

    /// <summary>Leaves the row as it was.</summary>
    [RelayCommand]
    private void Cancel(TrackedPlayerRow row)
    {
        row.IsRenaming = false;
        row.IsGrouping = false;
        row.Update(row.Player);
    }

    /// <summary>Reference SetGroupState: the group is folded away or opened, and the choice is kept.</summary>
    [RelayCommand]
    private void ToggleGroup(TrackedPlayerGroupRow group)
    {
        if (settings is null) return;
        var states = settings.Get(SettingKeys.GroupStates, new Dictionary<string, bool>());
        states[group.Key] = !group.IsExpanded;
        settings.Set(SettingKeys.GroupStates, states);
        Refresh();
    }

    /// <summary>Reference ShowBulkGroupEditorDialog: the group manager is open.</summary>
    [RelayCommand]
    private void OpenGroupManager()
    {
        GroupName = "";
        GroupColour = "None";
        GroupError = "";
        foreach (var row in GroupPlayers) row.IsSelected = false;
        IsManagingGroups = true;
    }

    /// <summary>
    /// Reference ShowBulkGroupEditorDialog: an existing group is picked, and its name, its colour and the players in it
    /// are filled in.
    /// </summary>
    [RelayCommand]
    private void ChooseGroup(string name)
    {
        GroupName = name;
        var sample = players.FirstOrDefault(player => player.GroupName == name);
        GroupColour = string.IsNullOrEmpty(sample?.GroupColor) ? "None" : sample!.GroupColor;
        foreach (var row in GroupPlayers) row.IsSelected = row.Player.GroupName == name;
    }

    /// <summary>
    /// Reference ShowBulkGroupEditorDialog: the group is saved - the players that are checked are in it, and the players
    /// that are not and were in it are out of it.
    /// </summary>
    [RelayCommand]
    private async Task SaveGroupAsync()
    {
        string name = GroupName.Trim();
        if (name.Length == 0)
        {
            GroupError = Loc.Text("EnterGroupName", "Please enter a group name.");
            return;
        }
        GroupError = "";
        players = players.Select(player =>
        {
            bool selected = rows.TryGetValue(player.BMId, out var row) && row.IsSelected;
            if (selected) return player with { GroupName = name, GroupColor = GroupColour };
            if (player.GroupName == name) return player with { GroupName = "", GroupColor = "" };
            return player;
        }).ToList();
        await SaveAsync().ConfigureAwait(false);
        IsManagingGroups = false;
        Refresh();
    }

    /// <summary>Reference ShowBulkGroupEditorDialog: the group manager is closed without saving.</summary>
    [RelayCommand]
    private void CancelGroupManager()
    {
        IsManagingGroups = false;
        GroupError = "";
        foreach (var row in GroupPlayers) row.IsSelected = false;
    }

    /// <summary>Reference ShowBulkGroupEditorDialog: the group manager is open.</summary>
    [ObservableProperty] private bool isManagingGroups;

    /// <summary>Reference ShowBulkGroupEditorDialog: the name of the group being edited.</summary>
    [ObservableProperty] private string groupName = "";

    /// <summary>Reference ShowBulkGroupEditorDialog: the colour of the group being edited.</summary>
    [ObservableProperty] private string groupColour = "None";

    /// <summary>Reference ShowBulkGroupEditorDialog: what is wrong with the group that is being edited.</summary>
    [ObservableProperty] private string groupError = "";

    /// <summary>Reference ShowBulkGroupEditorDialog: the players the manager lists, by their name.</summary>
    public IReadOnlyList<TrackedPlayerRow> GroupPlayers =>
        players.OrderBy(player => player.Name, StringComparer.OrdinalIgnoreCase).Select(Row).ToList();

    /// <summary>Reference CreateColorSelector: the colours the reference offers.</summary>
    public static IReadOnlyList<string> GroupColours { get; } =
        ["None", "Red", "Green", "Blue", "Yellow", "Purple", "Cyan", "Orange", "Pink", "White", "Gray"];

    /// <summary>Reference SetGroupOrder: the groups of a server are put in the order the user wants them in.</summary>
    [RelayCommand]
    private void MoveGroupUp(TrackedPlayerGroupRow group) => MoveGroup(group, up: true);

    /// <summary>Reference SetGroupOrder: the groups of a server are put in the order the user wants them in.</summary>
    [RelayCommand]
    private void MoveGroupDown(TrackedPlayerGroupRow group) => MoveGroup(group, up: false);

    private void MoveGroup(TrackedPlayerGroupRow group, bool up)
    {
        if (settings is null) return;
        var order = GroupOrder(group.ServerName).ToList();
        if (order.Count == 0)
        {
            // Reference TrackedGroup_Drop: the order starts with the groups the server has.
            order = Servers.FirstOrDefault(server => server.Name == group.ServerName)?.Groups
                .Select(server => server.GroupName).ToList() ?? [];
        }
        int index = order.IndexOf(group.GroupName);
        if (index < 0) order.Add(group.GroupName);
        else
        {
            int target = up ? index - 1 : index + 1;
            if (target < 0 || target >= order.Count) return;
            order.RemoveAt(index);
            order.Insert(target, group.GroupName);
        }
        var orders = settings.Get(SettingKeys.GroupOrder, new Dictionary<string, List<string>>());
        orders[group.ServerName] = order;
        settings.Set(SettingKeys.GroupOrder, orders);
        Refresh();
    }

    private async Task ChangeAsync(TrackedPlayerRow row, Func<TrackedPlayer, TrackedPlayer> change)
    {
        players = players.Select(player => player.BMId == row.BMId ? change(player) : player).ToList();
        await SaveAsync().ConfigureAwait(false);
        Refresh();
    }

    private async Task SaveAsync() => await store.SaveAsync(players, CancellationToken.None).ConfigureAwait(false);

    /// <summary>Reference RefreshTrackedPlayersList: the list is built again.</summary>
    private void Refresh()
    {
        OnPropertyChanged(nameof(Servers));
        OnPropertyChanged(nameof(HasNoMatches));
        OnPropertyChanged(nameof(Groups));
    }

    /// <summary>Reference RefreshTrackedPlayersList: the row of a player, kept while the list is rebuilt.</summary>
    private TrackedPlayerRow Row(TrackedPlayer player)
    {
        if (!rows.TryGetValue(player.BMId, out var row))
        {
            row = new TrackedPlayerRow(player);
            rows[player.BMId] = row;
        }
        row.Update(player);
        return row;
    }

    /// <summary>Reference TrackingService.GetGroupState: a group is open unless the user folded it away.</summary>
    private bool IsExpanded(string server, string group) =>
        settings?.Get(SettingKeys.GroupStates, new Dictionary<string, bool>()).GetValueOrDefault($"{server}|{group}", true) ?? true;

    /// <summary>Reference TrackingService.GetGroupOrder: the order the user put the groups of a server in.</summary>
    private IReadOnlyList<string> GroupOrder(string server) =>
        settings?.Get(SettingKeys.GroupOrder, new Dictionary<string, List<string>>()).GetValueOrDefault(server) ?? [];
}
