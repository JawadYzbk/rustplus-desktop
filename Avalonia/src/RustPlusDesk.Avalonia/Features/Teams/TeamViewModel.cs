using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Core.Connections;

namespace RustPlusDesk.Avalonia.Features.Teams;

public sealed record TeamMemberRow(ServerTeamMember Member)
{
    public string Name => string.IsNullOrWhiteSpace(Member.Name) ? Member.SteamId
        : new string(Member.Name.Where(c => !char.IsControl(c)).Take(256).ToArray());
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
    public string Caption => Members.Count == 0 ? "Team" : $"Team ({Members.Count})";
    [RelayCommand] private void Toggle() => IsExpanded = !IsExpanded;
    public void Reset() { Members = []; Message = "Connect to a server to see its team."; }

    public async Task RunAsync(IServerTeamSession session, CancellationToken ct)
    {
        Message = "Loading team…";
        using var timer = new PeriodicTimer(TimeSpan.FromSeconds(10));
        try
        {
            do
            {
                try
                {
                    var current = await session.ReadTeamAsync(ct).WaitAsync(ct);
                    ct.ThrowIfCancellationRequested();
                    if (!Members.Select(row => row.Member).SequenceEqual(current)) Members = current.Select(member => new TeamMemberRow(member)).ToArray();
                    Message = current.Count == 0 ? "No team members reported by this server." : "Current server team";
                }
                catch (OperationCanceledException) when (ct.IsCancellationRequested) { throw; }
                catch (Exception) { Message = "Team update unavailable. Displayed members may be out of date."; }
            } while (await timer.WaitForNextTickAsync(ct));
        }
        catch (OperationCanceledException) when (ct.IsCancellationRequested) { }
    }
}
