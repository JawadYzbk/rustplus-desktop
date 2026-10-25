using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Core.Map;

namespace RustPlusDesk.Avalonia.Features.Map;

public sealed partial class MapViewModel : ObservableObject
{
    private sealed record ActiveMap(IServerMap Client, CancellationToken Token);
    private ActiveMap? active;
    private readonly SemaphoreSlim liveReads = new(1, 1);
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasMap))]
    private ServerMapData? data;
    public bool HasMap => Data is not null;
    [ObservableProperty] private string status = "Connect to load the server map.";
    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(RefreshCommand))]
    [NotifyCanExecuteChangedFor(nameof(RefreshLiveCommand))]
    private bool isConnected;
    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(RefreshCommand))]
    private bool isLoading;
    [ObservableProperty] private bool showMonuments = true;
    [ObservableProperty] private bool showGrid = true;
    [ObservableProperty] private bool showTeam = true;
    [ObservableProperty] private bool showEvents = true;
    [ObservableProperty] private bool showShops = true;
    [ObservableProperty] private MapLiveData? liveData;
    [ObservableProperty] private string liveStatus = "Connect to load live markers.";
    [ObservableProperty] private IReadOnlyList<MapTeamMember> teamMembers = [];
    [ObservableProperty] private MapTeamMember? selectedMember;
    [ObservableProperty] private bool followSelectedMember;
    [ObservableProperty] private MapPoint? focusPosition;
    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(RefreshLiveCommand))]
    private bool isRefreshingLive;
    private bool CanRefreshLive => IsConnected && !IsRefreshingLive;
    partial void OnSelectedMemberChanged(MapTeamMember? value)
    {
        if (value is not null) FollowSelectedMember = true;
        UpdateFocus();
    }
    partial void OnFollowSelectedMemberChanged(bool value) => UpdateFocus();
    private void UpdateFocus() => FocusPosition = FollowSelectedMember && SelectedMember is { HasPosition: true } member
        ? new MapPoint(member.X!.Value, member.Y!.Value) : null;
    private bool CanRefresh => IsConnected && !IsLoading;

    public void Clear()
    {
        active = null;
        Data = null;
        ResetLive();
        IsConnected = IsLoading = false;
        Status = "Connect to load the server map.";
    }

    public async Task RunAsync(IServerMap client, CancellationToken cancellationToken)
    {
        var context = new ActiveMap(client, cancellationToken);
        active = context;
        IsConnected = true;
        var liveUpdates = UpdateLiveAsync(context, cancellationToken);
        try
        {
            await LoadAsync(context, cancellationToken);
            await Task.Delay(Timeout.Infinite, cancellationToken);
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { }
        finally
        {
            await liveUpdates;
            await Task.WhenAll(new[] { RefreshCommand.ExecutionTask, RefreshLiveCommand.ExecutionTask }.OfType<Task>());
            if (active == context)
            {
                active = null;
                ResetLive();
                IsConnected = IsLoading = false;
                Status = HasMap ? "Disconnected — last loaded terrain." : "Connect to load the server map.";
            }
        }
    }

    [RelayCommand(CanExecute = nameof(CanRefresh))]
    private async Task RefreshAsync(CancellationToken cancellationToken)
    {
        if (active is not { } context) return;
        using var linked = CancellationTokenSource.CreateLinkedTokenSource(context.Token, cancellationToken);
        try { await LoadAsync(context, linked.Token); }
        catch (OperationCanceledException) when (linked.IsCancellationRequested) { }
    }

    private async Task LoadAsync(ActiveMap context, CancellationToken cancellationToken)
    {
        IsLoading = true;
        Status = "Loading server terrain…";
        try
        {
            var loaded = await context.Client.ReadAsync(cancellationToken).WaitAsync(cancellationToken);
            cancellationToken.ThrowIfCancellationRequested();
            if (active != context) return;
            Data = loaded;
            Status = loaded.WorldSize is { } size ? $"World size {size:N0} m · {loaded.Monuments.Count} monuments"
                : "Terrain loaded. World size unavailable; monument positions cannot be calibrated.";
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { throw; }
        catch (Exception)
        {
            if (active == context) Status = HasMap ? "Refresh failed — showing the last loaded terrain."
                : "Map unavailable. Check server support and refresh to retry.";
        }
        finally { if (active == context) IsLoading = false; }
    }

    private void ResetLive()
    {
        LiveData = null;
        TeamMembers = [];
        SelectedMember = null;
        FollowSelectedMember = false;
        FocusPosition = null;
        IsRefreshingLive = false;
        LiveStatus = "Live markers disconnected.";
    }

    private async Task UpdateLiveAsync(ActiveMap context, CancellationToken cancellationToken)
    {
        try
        {
            await LoadLiveAsync(context, cancellationToken);
            using var timer = new PeriodicTimer(TimeSpan.FromSeconds(3));
            while (await timer.WaitForNextTickAsync(cancellationToken)) await LoadLiveAsync(context, cancellationToken);
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { }
    }

    [RelayCommand(CanExecute = nameof(CanRefreshLive))]
    private async Task RefreshLiveAsync(CancellationToken cancellationToken)
    {
        if (active is not { } context) return;
        using var linked = CancellationTokenSource.CreateLinkedTokenSource(context.Token, cancellationToken);
        try { await LoadLiveAsync(context, linked.Token); }
        catch (OperationCanceledException) when (linked.IsCancellationRequested) { }
    }

    private async Task LoadLiveAsync(ActiveMap context, CancellationToken cancellationToken)
    {
        if (!await liveReads.WaitAsync(0, cancellationToken)) return;
        try
        {
            if (active != context) return;
            IsRefreshingLive = true;
            var snapshot = await context.Client.ReadLiveAsync(cancellationToken).WaitAsync(cancellationToken);
            cancellationToken.ThrowIfCancellationRequested();
            if (active != context) return;
            var selectedId = SelectedMember?.SteamId;
            bool following = FollowSelectedMember;
            LiveData = snapshot;
            TeamMembers = (snapshot.Team ?? []).OrderByDescending(m => m.IsLeader).ThenByDescending(m => m.IsOnline).ThenBy(m => m.Name).ToArray();
            SelectedMember = TeamMembers.FirstOrDefault(m => m.SteamId == selectedId);
            FollowSelectedMember = following;
            UpdateFocus();
            LiveStatus = $"{(snapshot.Team is null ? "Team unavailable" : $"{TeamMembers.Count} team members")} · "
                + (snapshot.Markers is null ? "Map events unavailable" : $"{snapshot.Markers.Count} live markers");
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { throw; }
        catch (Exception)
        {
            if (active == context)
            {
                ResetLive();
                LiveStatus = "Live markers unavailable. Retry or wait for automatic refresh.";
            }
        }
        finally
        {
            if (active == context) IsRefreshingLive = false;
            liveReads.Release();
        }
    }
}
