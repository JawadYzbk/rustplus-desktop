using CommunityToolkit.Mvvm.ComponentModel;
using RustPlusDesk.Core.Chat;

namespace RustPlusDesk.Avalonia.Features.Chat;

public sealed partial class ServerChatViewModel : ObservableObject
{
    public ChatLaneViewModel Team { get; } = new(ChatChannel.Team) { IsSelected = true };
    public ChatLaneViewModel Clan { get; } = new(ChatChannel.Clan);
    private object? active;
    [ObservableProperty] private int selectedIndex;
    [ObservableProperty] private bool isVisible = true;
    partial void OnSelectedIndexChanged(int value) => UpdateReadState();
    partial void OnIsVisibleChanged(bool value) => UpdateReadState();
    private void UpdateReadState()
    {
        Team.IsSelected = IsVisible && SelectedIndex == 0;
        Clan.IsSelected = IsVisible && SelectedIndex == 1;
    }

    public void Clear()
    {
        active = null;
        Team.Clear();
        Clan.Clear();
    }

    public async Task RunAsync(IServerChat client, CancellationToken cancellationToken)
    {
        var context = new object();
        active = context;
        var team = Team.RunAsync(client, cancellationToken);
        var clan = Clan.RunAsync(client, cancellationToken);
        try
        {
            await foreach (var message in client.WatchAsync(cancellationToken))
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (active != context) return;
                if (message.Channel == ChatChannel.Team) Team.Receive(message);
                else if (message.Channel == ChatChannel.Clan) Clan.Receive(message);
            }
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { }
        catch (Exception) { /* Periodic history reads continue if broadcast delivery fails. */ }
        finally { await Task.WhenAll(team, clan); }
    }
}
