using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Core.Alerts;
using RustPlusDesk.Core.Connections;
using RustPlusDesk.Core.Settings;

namespace RustPlusDesk.Avalonia.Features.Teams;

public sealed record TeamMemberRow(ServerTeamMember Member, bool Abbreviate = false, string? AfkText = null)
{
    /// <summary>Reference TeamMemberVM.IsAfk: shown after 5 idle minutes with the "AFK: m:ss" text.</summary>
    public bool IsAfk => AfkText is not null;
    /// <summary>The member's Steam avatar when loaded (reference TeamMemberVM.Avatar).</summary>
    public global::Avalonia.Media.Imaging.Bitmap? Avatar { get; init; }
    public bool HasAvatar => Avatar is not null;
    /// <summary>First letter shown while there is no avatar.</summary>
    public string Initial => Name.Length > 0 ? Name[..1].ToUpperInvariant() : "?";
    /// <summary>The member's name (streamer mode: first character + "..."), else the abbreviated Steam ID.</summary>
    public string Name => string.IsNullOrWhiteSpace(Member.Name) ? SteamId
        : StreamerMode.Name(new string(Member.Name.Where(c => !char.IsControl(c)).Take(256).ToArray()), Abbreviate);
    public string SteamId => StreamerMode.SteamId(Member.SteamId, Abbreviate);
    public bool IsOnline => Member.IsOnline == true;
    public bool IsDead => Member.IsAlive == false;
    public string State => (Member.IsOnline switch { true => "Online", false => "Offline", _ => "Connection unknown" }) +
        (Member.IsAlive switch { true => " · Alive", false => " · Dead", _ => "" });
}

/// <summary>Displays only the current owned server's roster; never publishes membership or master preferences.</summary>
public sealed partial class TeamViewModel : ObservableObject
{
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(Caption))]
    private IReadOnlyList<TeamMemberRow> members = [];
    [ObservableProperty] private string message = "Connect to a server to see its team.";
    [ObservableProperty] private bool isExpanded;
    /// <summary>Reference TeamMemberVM.Abbreviate (streamer mode).</summary>
    [ObservableProperty] private bool abbreviate;
    partial void OnAbbreviateChanged(bool value) => Members = Members.Select(row => row with { Abbreviate = value }).ToArray();
    public string Caption => Members.Count == 0 ? "Team" : $"Team ({Members.Count})";
    [RelayCommand] private void Toggle() => IsExpanded = !IsExpanded;
    /// <summary>Raised for every successful team read (reference LoadTeamAsync feeds the death tracker).</summary>
    public event Action<ServerTeamSnapshot>? SnapshotRead;
    /// <summary>Raised when the roster is reset (connection teardown); per-connection trackers start over.</summary>
    public event Action? Cleared;
    public void Reset() { Members = []; Message = "Connect to a server to see its team."; afk.Reset(); Cleared?.Invoke(); }

    private readonly AfkTracker afk = new();
    /// <summary>Clock for AFK tracking (UTC); replaced by tests.</summary>
    public Func<DateTime> Now { get; set; } = () => DateTime.UtcNow;
    /// <summary>Reference AfkAlertMinutes: the idle minutes after which an AFK alert is raised.</summary>
    public Func<int> AfkThresholdMinutes { get; set; } = () => AfkTracker.DefaultThresholdMinutes;
    /// <summary>Raised by the one-second AFK tick for every member that went AFK or came back (reference AfkTimer_Tick).</summary>
    public event Action<AfkAlert>? AfkAlertRaised;
    /// <summary>Steam avatars by Steam ID (reference AvatarLoader cache).</summary>
    public Func<string, global::Avalonia.Media.Imaging.Bitmap?>? Avatars { get; set; }

    /// <summary>Reference TeamMembers.Where(t =&gt; t.IsAfk) for the chat AFK command: the AFK rows with their idle clock.</summary>
    public IReadOnlyList<RustPlusDesk.Core.Chat.ChatAfkMember> AfkMembers()
    {
        DateTime now = Now();
        var result = new List<RustPlusDesk.Core.Chat.ChatAfkMember>();
        foreach (var row in Members)
        {
            if (!row.IsAfk || afk.Idle(row.Member.SteamId, now) is not { } idle) continue;
            result.Add(new RustPlusDesk.Core.Chat.ChatAfkMember(row.Name, now - idle));
        }
        return result;
    }

    /// <summary>Reference EnsureAvatarAsync completion: rows pick up avatars that arrived.</summary>
    public void RefreshAvatars()
    {
        if (Avatars is not { } avatars || Members.All(row => ReferenceEquals(row.Avatar, avatars(row.Member.SteamId)))) return;
        Members = Members.Select(row => row with { Avatar = avatars(row.Member.SteamId) }).ToArray();
    }

    /// <summary>Reference AfkTimer_Tick: ages the idle clocks, refreshes the AFK texts and raises the alerts.</summary>
    public void TickAfk()
    {
        int threshold = AfkThresholdMinutes();
        var alerts = afk.Tick(Now(), threshold);
        ApplyAfkTexts();
        foreach (var alert in alerts) AfkAlertRaised?.Invoke(alert);
    }

    private void ApplyAfkTexts()
    {
        if (Members.All(row => row.AfkText == afk.AfkText(row.Member.SteamId))) return;
        Members = Members.Select(row => row with { AfkText = afk.AfkText(row.Member.SteamId) }).ToArray();
    }

    private async Task RunAfkAsync(CancellationToken ct)
    {
        using var timer = new PeriodicTimer(TimeSpan.FromSeconds(1));
        try
        {
            while (await timer.WaitForNextTickAsync(ct)) TickAfk();
        }
        catch (OperationCanceledException) when (ct.IsCancellationRequested) { }
    }

    public async Task RunAsync(IServerTeamSession session, CancellationToken ct)
    {
        Message = "Loading team…";
        var afkTicks = RunAfkAsync(ct);
        using var timer = new PeriodicTimer(TimeSpan.FromSeconds(10));
        try
        {
            do
            {
                try
                {
                    var snapshot = await session.ReadTeamSnapshotAsync(ct).WaitAsync(ct);
                    ct.ThrowIfCancellationRequested();
                    var current = snapshot.Members;
                    SnapshotRead?.Invoke(snapshot);
                    afk.Observe(current, Now());
                    if (!Members.Select(row => row.Member).SequenceEqual(current))
                        Members = current.Select(member => new TeamMemberRow(member, Abbreviate, afk.AfkText(member.SteamId)) { Avatar = Avatars?.Invoke(member.SteamId) }).ToArray();
                    Message = current.Count == 0 ? "No team members reported by this server." : "Current server team";
                }
                catch (OperationCanceledException) when (ct.IsCancellationRequested) { throw; }
                catch (Exception) { Message = "Team update unavailable. Displayed members may be out of date."; }
            } while (await timer.WaitForNextTickAsync(ct));
        }
        catch (OperationCanceledException) when (ct.IsCancellationRequested) { }
        await afkTicks;
    }
}
