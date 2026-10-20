using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Core.Clans;

namespace RustPlusDesk.Avalonia.Features.Clans;

/// <summary>One clan member as the tab shows them (reference ClanMemberVM).</summary>
public sealed record ClanMemberRow(ClanMember Member, string RoleName)
{
    public ulong SteamId => Member.SteamId;
    public string Name => string.IsNullOrWhiteSpace(Member.Name) ? Member.SteamId.ToString() : Member.Name;
    public bool IsOnline => Member.IsOnline;
    public string State => IsOnline ? "Online" : "Offline";
    public string Initial => Name.Length > 0 ? Name[..1].ToUpperInvariant() : "?";
}

/// <summary>
/// Reference MainWindow.Clan.Core: the clan of the connected player with its roles, members and message of the day.
/// The list is polled every fifteen seconds (reference TeamTimer_Tick) and the role permissions are reconciled
/// against it, so a role the clan gained arrives switched off rather than inheriting permission nobody gave.
/// </summary>
public sealed partial class ClanViewModel : ObservableObject
{
    /// <summary>Reference TeamTimer_Tick: the clan is pulled every fifteen seconds.</summary>
    public static readonly TimeSpan PollInterval = TimeSpan.FromSeconds(15);

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasClan))]
    [NotifyPropertyChangedFor(nameof(Members))]
    [NotifyPropertyChangedFor(nameof(Roles))]
    [NotifyPropertyChangedFor(nameof(Name))]
    [NotifyPropertyChangedFor(nameof(Motd))]
    [NotifyPropertyChangedFor(nameof(MemberCountText))]
    private ClanSnapshot? clan;
    [ObservableProperty] private string message = "Connect to a server to see its clan.";
    [ObservableProperty] private bool isLoading;
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasClan))]
    private bool isFailed;

    public bool HasClan => Clan is not null;
    public string Name => Clan?.Name ?? "";
    public string Motd => Clan?.Motd ?? "";
    public string MemberCountText => Clan is { } clan ? $"{clan.OnlineCount} online · {clan.MemberCount} / {clan.MaxMemberCount ?? 100}" : "";

    /// <summary>The members in the order the server reports them (reference ClanMembers).</summary>
    public IReadOnlyList<ClanMemberRow> Members => (Clan?.Members ?? [])
        .Select(member => new ClanMemberRow(member, member.RoleName))
        .ToArray();

    public IReadOnlyList<ClanRole> Roles => Clan?.Roles ?? [];

    /// <summary>Reference ApplyClanEmptyState: the empty and the failed state are told apart.</summary>
    public void Reset()
    {
        Clan = null;
        IsFailed = false;
        IsLoading = false;
        Message = "Connect to a server to see its clan.";
    }

    /// <summary>One pull: the clan snapshot replaces the previous one, or the empty/failed state is shown.</summary>
    public void Apply(ClanSnapshot? snapshot, bool failed)
    {
        IsLoading = false;
        if (snapshot is null)
        {
            IsFailed = failed;
            if (failed) Message = "The clan could not be read. The list may be out of date.";
            else { Clan = null; Message = "You are not in a clan."; }
            return;
        }

        IsFailed = false;
        Clan = snapshot;
        Message = snapshot.Members.Count == 0 ? "This clan has no members." : $"{snapshot.MemberCount} member(s).";
    }

    /// <summary>Reads the clan; set by the connection once a session is live.</summary>
    public Func<CancellationToken, Task<ClanSnapshot?>>? Reader { get; set; }

    [RelayCommand]
    private async Task RefreshAsync()
    {
        if (Reader is not { } read) return;
        IsLoading = true;
        try
        {
            Apply(await read(CancellationToken.None).ConfigureAwait(false), false);
        }
        catch (Exception) { Apply(null, true); }
    }

    /// <summary>Polls the clan while the connection lives (reference TeamTimer_Tick).</summary>
    public async Task RunAsync(IClanSession session, CancellationToken ct)
    {
        Reader = token => session.ReadClanAsync(token);
        await RefreshAsync();
        using var timer = new PeriodicTimer(PollInterval);
        try
        {
            while (await timer.WaitForNextTickAsync(ct))
            {
                ct.ThrowIfCancellationRequested();
                await RefreshAsync();
            }
        }
        catch (OperationCanceledException) when (ct.IsCancellationRequested) { }
    }
}
