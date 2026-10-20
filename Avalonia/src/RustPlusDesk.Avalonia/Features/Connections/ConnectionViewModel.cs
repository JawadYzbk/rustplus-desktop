using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Core.Connections;
using RustPlusDesk.Core.Servers;
using RustPlusDesk.Core.Time;
using RustPlusDesk.Avalonia.Features.Devices;
using RustPlusDesk.Core.Cloud;
using RustPlusDesk.Core.Devices;
using RustPlusDesk.Avalonia.Features.Teams;

namespace RustPlusDesk.Avalonia.Features.Connections;

/// <summary>Owns one selected server's connection, recovery loop, and status presentation.</summary>
public sealed partial class ConnectionViewModel(IServerSessionFactory factory) : ObservableObject
{
    public DevicesViewModel Devices { get; } = new();
    public TeamViewModel Team { get; } = new();
    public Clans.ClanViewModel Clan { get; } = new();
    public Map.MapViewModel Map { get; } = new();
    private Chat.ChatViewModel? chat;
    public Chat.ChatViewModel Chat => chat ??= new() { CenterOnMap = Map.CenterOnPlayer };
    private string profilesPath = "";
    private ServerSummary? selectedServer;
    private ServerClock clock = new();
    private DateTime? sampleTime;
    private double? sampleHours;
    private (IServerSession Session, CancellationTokenSource Lifetime, string Id, DateTimeOffset? Wipe)? activeSession;
    private Task<DeviceSynchronizationResult>? synchronization;
    private Task? deviceImport;
    public bool CanImportDevices => IsConnected && activeSession?.Session is IServerTeamSession;

    public async Task<T> RunDeviceImportAsync<T>(string path, ServerSummary server, string? expectedConnection,
        Func<IServerTeamSession, string, CancellationToken, Task<T>> operation, CancellationToken ct)
    {
        if (!CanImportDevices || activeSession is not { } connected || connected.Session is not IServerTeamSession team ||
            profilesPath != path || selectedServer != server || deviceImport is not null ||
            expectedConnection is not null && connected.Id != expectedConnection)
            throw new DeviceImportException(DeviceImportFailure.Disconnected);
        using var lifetime = CancellationTokenSource.CreateLinkedTokenSource(ct, connected.Lifetime.Token);
        var task = operation(team, connected.Id, lifetime.Token);
        deviceImport = task;
        try { return await task; }
        finally { if (deviceImport == task) deviceImport = null; }
    }

    public async Task<DeviceSynchronizationResult?> SynchronizeDevicesAsync(ICloudDeviceSynchronizer service, string userId, string steamId, CancellationToken ct)
    {
        if (!IsConnected || activeSession is not { } connected || selectedServer is not { } server || synchronization is not null) return null;
        using var lifetime = CancellationTokenSource.CreateLinkedTokenSource(ct, connected.Lifetime.Token);
        var task = service.SynchronizeAsync(profilesPath, server, userId, steamId, connected.Id, connected.Wipe, connected.Session.Devices, lifetime.Token);
        synchronization = task;
        try { return await task; }
        finally { if (synchronization == task) synchronization = null; }
    }

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsConnected))]
    [NotifyPropertyChangedFor(nameof(CanImportDevices))]
    [NotifyPropertyChangedFor(nameof(IsActive))]
    [NotifyPropertyChangedFor(nameof(IsPending))]
    [NotifyPropertyChangedFor(nameof(RequiresRePair))]
    [NotifyCanExecuteChangedFor(nameof(ConnectCommand))]
    [NotifyCanExecuteChangedFor(nameof(DisconnectCommand))]
    private ConnectionPhase phase;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(DisplayServerName))]
    private string serverName = "Select a server";
    /// <summary>Reference streamer mode: the selected server's name is shown as "SERVER HIDDEN".</summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(DisplayServerName))]
    private bool hideServerName;
    public string DisplayServerName => HideServerName && selectedServer is not null ? RustPlusDesk.Core.Settings.StreamerMode.HiddenServer : ServerName;
    [ObservableProperty] private string statusMessage = "Disconnected";
    [ObservableProperty] private string population = "—";
    [ObservableProperty] private string queue = "—";
    /// <summary>Reference _vm.ServerPlayers: the raw count the chat "pop" command answers with.</summary>
    public uint? PlayerCount { get; private set; }
    /// <summary>Reference _vm.ServerQueue: the raw queue length the chat "pop" command reports.</summary>
    public uint? QueueCount { get; private set; }
    [ObservableProperty] private string serverTime = "—";
    [ObservableProperty] private string nextPhase = "";

    public bool IsConnected => Phase == ConnectionPhase.Connected;
    public bool IsActive => Phase is ConnectionPhase.Connecting or ConnectionPhase.Connected or ConnectionPhase.Reconnecting;
    public bool IsPending => Phase is ConnectionPhase.Connecting or ConnectionPhase.Reconnecting;
    public bool RequiresRePair => Phase == ConnectionPhase.RePairRequired;
    private bool CanConnect => selectedServer is not null && !IsActive;
    private bool CanDisconnect => IsActive;

    public void Select(string path, ServerSummary? server)
    {
        if (profilesPath == path && selectedServer == server) return;
        ConnectCommand.Cancel();
        profilesPath = path;
        selectedServer = server;
        Devices.Select(server?.Devices ?? []);
        ServerName = server?.Name ?? "Select a server";
        ResetStatus();
        Phase = ConnectionPhase.Disconnected;
        StatusMessage = "Disconnected";
        ConnectCommand.NotifyCanExecuteChanged();
    }

    [RelayCommand(CanExecute = nameof(CanConnect))]
    private async Task ConnectAsync(CancellationToken cancellationToken)
    {
        if (selectedServer is null) return;
        var target = selectedServer;
        string source = profilesPath;
        bool hasConnected = false;
        int retryDelaySeconds = 2;
        Phase = ConnectionPhase.Connecting;
        StatusMessage = "Connecting…";
        ResetStatus();
        try
        {
            while (true)
            {
                cancellationToken.ThrowIfCancellationRequested();
                try
                {
                    await using var session = await factory.ConnectAsync(source, target.ProfileKey, cancellationToken);
                    cancellationToken.ThrowIfCancellationRequested();
                    hasConnected = true;
                    retryDelaySeconds = 2;
                    clock = new ServerClock();
                    ApplyStatus(session.InitialStatus);
                    using var connectedCancellation = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
                    activeSession = (session, connectedCancellation, Guid.NewGuid().ToString("N"), session.InitialStatus.WipeTime);
                    Phase = ConnectionPhase.Connected;
                    StatusMessage = "Connected";
                    var disconnectWatch = CancelOnDisconnectAsync(session.Disconnected, connectedCancellation);
                    var clockUpdates = UpdateClockAsync(connectedCancellation.Token);
                    var teamUpdates = session is IServerTeamSession team ? Team.RunAsync(team, connectedCancellation.Token) : Task.CompletedTask;
                    var clanUpdates = session is RustPlusDesk.Core.Clans.IClanSession clanSession ? Clan.RunAsync(clanSession, connectedCancellation.Token) : Task.CompletedTask;
                    var mapUpdates = session is RustPlusDesk.Core.Map.IServerMapSession mapSession ? Map.RunAsync(mapSession, connectedCancellation.Token) : Task.CompletedTask;
                    var chatUpdates = session is RustPlusDesk.Core.Chat.IServerChatSession chatSession ? Chat.RunAsync(chatSession, connectedCancellation.Token) : Task.CompletedTask;
                    var deviceUpdates = target.Devices.Count > 0
                        ? Devices.RunAsync(session.Devices, connectedCancellation.Token) : Task.CompletedTask;
                    try
                    {
                        while (true)
                        {
                            await Task.Delay(TimeSpan.FromSeconds(3), connectedCancellation.Token);
                            var status = await session.ReadStatusAsync(connectedCancellation.Token).WaitAsync(connectedCancellation.Token);
                            connectedCancellation.Token.ThrowIfCancellationRequested();
                            ApplyStatus(status);
                            if (status.WipeTime != session.InitialStatus.WipeTime)
                                throw new ServerConnectionException(ConnectionFailure.Unreachable);
                        }
                    }
                    catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested && session.Disconnected.IsCompletedSuccessfully)
                    { throw new ServerConnectionException(ConnectionFailure.Unreachable); }
                    finally
                    {
                        connectedCancellation.Cancel();
                        activeSession = null;
                        if (synchronization is { } pending)
                        {
                            try { await pending; }
                            catch (Exception) { /* The sync owner reports its failure; disposal still awaits its completion. */ }
                        }
                        if (deviceImport is { } importing)
                        {
                            try { await importing; }
                            catch (Exception) { /* The import owner reports its failure after connected work has stopped. */ }
                        }
                        await disconnectWatch;
                        try { await clockUpdates; }
                        catch (OperationCanceledException) when (connectedCancellation.IsCancellationRequested) { }
                        await deviceUpdates;
                        await teamUpdates;
                        await clanUpdates;
                        await mapUpdates;
                        await chatUpdates;
                    }
                }
                catch (ServerConnectionException ex) when (ex.Failure == ConnectionFailure.Unreachable && hasConnected)
                {
                    ResetStatus();
                    Phase = ConnectionPhase.Reconnecting;
                    StatusMessage = $"Connection lost. Retrying in {retryDelaySeconds}s…";
                    await Task.Delay(TimeSpan.FromSeconds(retryDelaySeconds), cancellationToken);
                    retryDelaySeconds = Math.Min(retryDelaySeconds * 2, 60);
                }
            }
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { }
        catch (ServerConnectionException ex)
        {
            ResetStatus();
            Phase = ex.Failure switch
            {
                ConnectionFailure.RePairRequired => ConnectionPhase.RePairRequired,
                ConnectionFailure.CompanionLoginRequired => ConnectionPhase.SignInRequired,
                _ => ConnectionPhase.Failed
            };
            StatusMessage = ex.Failure switch
            {
                ConnectionFailure.RePairRequired => "Re-pair needed. Pair this server again from Rust.",
                ConnectionFailure.CompanionLoginRequired => "Companion sign-in needed. Sign in again in the existing app, then retry.",
                ConnectionFailure.InvalidCredentials => "The saved pairing credentials are invalid. Re-pair this server.",
                ConnectionFailure.ProfileUnavailable => "The saved profile changed or is unavailable. Refresh the server list.",
                _ => "Server unreachable. Check the connection and try again."
            };
        }
        catch (Exception)
        {
            ResetStatus();
            Phase = ConnectionPhase.Failed;
            StatusMessage = "Unable to connect. Refresh the server list and try again.";
        }
        finally
        {
            if (cancellationToken.IsCancellationRequested || IsActive)
            {
                ResetStatus();
                Phase = ConnectionPhase.Disconnected;
                StatusMessage = "Disconnected";
            }
        }
    }

    private static async Task CancelOnDisconnectAsync(Task disconnected, CancellationTokenSource lifetime)
    {
        try
        {
            await disconnected.WaitAsync(lifetime.Token).ConfigureAwait(false);
            lifetime.Cancel();
        }
        catch (OperationCanceledException) when (lifetime.IsCancellationRequested) { }
    }

    [RelayCommand(CanExecute = nameof(CanDisconnect))]
    private Task DisconnectAsync() => StopAsync();

    public async Task StopAsync()
    {
        ConnectCommand.Cancel();
        if (ConnectCommand.ExecutionTask is { } task) await task;
    }

    /// <summary>Whether the live session can hand the team leadership to a teammate (the chat "promote" command).</summary>
    public bool CanPromote => activeSession is { Session: IServerLeaderSession };

    /// <summary>Reference real.PromoteToLeaderAsync through the live session; false when it cannot or the server refused.</summary>
    public Task<bool> PromoteToLeaderAsync(ulong steamId, CancellationToken ct) =>
        activeSession is { Session: IServerLeaderSession leader }
            ? leader.PromoteToLeaderAsync(steamId, ct)
            : Task.FromResult(false);

    private void ApplyStatus(ServerStatus status)
    {
        if (!string.IsNullOrWhiteSpace(status.Name)) ServerName = status.Name;
        PlayerCount = status.Players;
        QueueCount = status.Queue;
        Population = $"{status.Players?.ToString() ?? "—"} / {status.MaxPlayers?.ToString() ?? "—"}";
        Queue = status.Queue?.ToString() ?? "—";
        if (status.GameHours is { } hours && double.IsFinite(hours))
        {
            sampleHours = ((hours % 24) + 24) % 24;
            sampleTime = DateTime.UtcNow;
            clock.Observe(sampleHours.Value, sampleTime.Value, status.Sunrise, status.Sunset, status.DayLengthMinutes, status.TimeScale);
        }
        UpdateClock();
    }

    private void UpdateClock()
    {
        if (sampleTime is null || sampleHours is null || DateTime.UtcNow - sampleTime > TimeSpan.FromMinutes(2))
        {
            ServerTime = "—";
            NextPhase = "";
            return;
        }
        double hours = clock.Advance(sampleHours.Value, (DateTime.UtcNow - sampleTime.Value).TotalMinutes);
        int minutes = (int)Math.Round(hours * 60, MidpointRounding.AwayFromZero) % 1440;
        ServerTime = $"{minutes / 60:00}:{minutes % 60:00}";
        if (clock.IsPaused) { NextPhase = "Time paused"; return; }
        bool day = clock.IsDay(hours);
        double remainingHours = day ? clock.Sunset - hours : hours >= clock.Sunset ? 24 - hours + clock.Sunrise : clock.Sunrise - hours;
        int remainingMinutes = (int)Math.Ceiling(remainingHours / (day ? clock.DaySpeed : clock.NightSpeed));
        NextPhase = $"{remainingMinutes}m until {(day ? "night" : "day")}";
    }

    private async Task UpdateClockAsync(CancellationToken cancellationToken)
    {
        using var timer = new PeriodicTimer(TimeSpan.FromSeconds(1));
        while (await timer.WaitForNextTickAsync(cancellationToken))
        {
            UpdateClock();
            Devices.UpdateTime();
        }
    }

    private void ResetStatus()
    {
        Team.Reset();
        Clan.Reset();
        Map.Reset();
        Chat.Reset();
        sampleTime = null;
        sampleHours = null;
        PlayerCount = QueueCount = null;
        Population = Queue = ServerTime = "—";
        NextPhase = "";
    }
}
