using RustPlusApi;
using RustPlusDesk.Core.Connections;
using RustPlusDesk.Core.Devices;
using RustPlusDesk.Core.Map;
using RustPlusDesk.Core.Chat;
using System.Threading.Channels;

namespace RustPlusDesk.Infrastructure.Connections;

internal sealed class RustPlusTransport : IRustTransport, IServerTeamSession, IServerMarkerSession, IServerChatSession
{
    private readonly RustPlus api;
    private readonly bool companionSocket;
    private bool disposed;
    private readonly RustRequestGate requests = new();
    private readonly RustPlusDeviceClient devices;
    private readonly TaskCompletionSource disconnected = new(TaskCreationOptions.RunContinuationsAsynchronously);
    public Task Disconnected => disconnected.Task;
    public IDeviceClient Devices => devices;

    public RustPlusTransport(RustPlusConnection connection)
    {
        api = new RustPlus(connection, new RustPlusSocketOptions { RequestTimeout = TimeSpan.FromSeconds(6) });
        devices = new RustPlusDeviceClient(api, requests);
        api.Disconnected += OnDisconnected;
        api.OnTeamChatReceived += OnTeamChat;
        api.OnClanChatReceived += OnClanChat;
        companionSocket = connection.ServerId is not null;
    }

    public ServerStatus InitialStatus { get; private set; } = null!;
    private void OnDisconnected(object? sender, EventArgs e) => disconnected.TrySetResult();

    private readonly Channel<ChatMessage> chat = Channel.CreateBounded<ChatMessage>(new BoundedChannelOptions(500) { FullMode = BoundedChannelFullMode.DropOldest });
    private void OnTeamChat(object? sender, RustPlusApi.Data.Events.TeamMessageEventArg message)
        => chat.Writer.TryWrite(new ChatMessage(ChatChannel.Team, message.SteamId, message.Name ?? "", message.Message ?? "", Time(message.Time), Hex(message.Color)));
    private void OnClanChat(object? sender, RustPlusApi.Data.Events.ClanMessageEventArg message)
        => chat.Writer.TryWrite(new ChatMessage(ChatChannel.Clan, message.SteamId, message.Name ?? "", message.Message ?? "", Time(message.Time)));
    private static string Hex(System.Drawing.Color color) => color.IsEmpty ? "" : $"#{color.R:X2}{color.G:X2}{color.B:X2}";
    private static DateTimeOffset Time(DateTime? utc) => utc is { } value ? new DateTimeOffset(DateTime.SpecifyKind(value, DateTimeKind.Utc)) : DateTimeOffset.UtcNow;

    public async IAsyncEnumerable<ChatMessage> WatchChatAsync([System.Runtime.CompilerServices.EnumeratorCancellation] CancellationToken ct)
    {
        await foreach (var message in chat.Reader.ReadAllAsync(ct).ConfigureAwait(false)) yield return message;
    }

    public async Task<IReadOnlyList<ChatMessage>> ReadChatAsync(ChatChannel channel, CancellationToken ct)
    {
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(ct);
        timeout.CancelAfter(TimeSpan.FromSeconds(8));
        try
        {
            await requests.WaitAsync(timeout.Token).ConfigureAwait(false);
            if (channel == ChatChannel.Clan)
            {
                var clan = await api.GetClanChatAsync(timeout.Token).WaitAsync(timeout.Token).ConfigureAwait(false);
                if (!clan.IsSuccess || clan.Data is null) throw new ServerConnectionException(ConnectionFailureClassifier.Response(clan.Error?.Message));
                return (clan.Data.Messages ?? []).Select(message => new ChatMessage(ChatChannel.Clan, message.SteamId, message.Name ?? "", message.Message ?? "", Time(message.Time))).ToArray();
            }
            var team = await api.GetTeamChatAsync(timeout.Token).WaitAsync(timeout.Token).ConfigureAwait(false);
            if (!team.IsSuccess || team.Data is null) throw new ServerConnectionException(ConnectionFailureClassifier.Response(team.Error?.Message));
            return (team.Data.Messages ?? []).Select(message => new ChatMessage(ChatChannel.Team, message.SteamId, message.Name ?? "", message.Message ?? "", Time(message.Time), Hex(message.Color))).ToArray();
        }
        catch (OperationCanceledException) when (ct.IsCancellationRequested) { throw; }
        catch (ServerConnectionException) { throw; }
        catch (Exception ex) { throw new ServerConnectionException(ConnectionFailureClassifier.Exception(ex, companionSocket)); }
    }

    public async Task SendChatAsync(ChatChannel channel, string text, CancellationToken ct)
    {
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(ct);
        timeout.CancelAfter(TimeSpan.FromSeconds(8));
        try
        {
            await requests.WaitAsync(timeout.Token).ConfigureAwait(false);
            if (channel == ChatChannel.Clan)
            {
                var clan = await api.SendClanMessageAsync(text, timeout.Token).WaitAsync(timeout.Token).ConfigureAwait(false);
                if (!clan.IsSuccess) throw new ServerConnectionException(ConnectionFailureClassifier.Response(clan.Error?.Message));
            }
            else
            {
                var team = await api.SendTeamMessageAsync(text, timeout.Token).WaitAsync(timeout.Token).ConfigureAwait(false);
                if (!team.IsSuccess) throw new ServerConnectionException(ConnectionFailureClassifier.Response(team.Error?.Message));
            }
        }
        catch (OperationCanceledException) when (ct.IsCancellationRequested) { throw; }
        catch (ServerConnectionException) { throw; }
        catch (Exception ex) { throw new ServerConnectionException(ConnectionFailureClassifier.Exception(ex, companionSocket)); }
    }

    public async Task InitializeAsync(CancellationToken cancellationToken)
    {
        await api.ConnectAsync(cancellationToken).ConfigureAwait(false);
        InitialStatus = await ReadStatusAsync(cancellationToken).ConfigureAwait(false);
    }

    public async Task<ServerStatus> ReadStatusAsync(CancellationToken cancellationToken)
    {
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        timeout.CancelAfter(TimeSpan.FromSeconds(6));
        try
        {
            await requests.WaitAsync(timeout.Token).ConfigureAwait(false);
            var info = await api.GetInfoAsync(timeout.Token).WaitAsync(timeout.Token).ConfigureAwait(false);
            if (!info.IsSuccess || info.Data is null)
                throw new ServerConnectionException(ConnectionFailureClassifier.Response(info.Error?.Message));

            var data = info.Data;
            var status = new ServerStatus(data.Name ?? "", data.PlayerCount, data.MaxPlayerCount, data.QueuedPlayerCount,
                WipeTime: data.WipeTime is { } wipe ? new DateTimeOffset(DateTime.SpecifyKind(wipe, DateTimeKind.Utc)) : null);
            // Time is optional on some modded servers; authenticated server info still proves connectivity.
            try
            {
                await requests.WaitAsync(timeout.Token).ConfigureAwait(false);
                var time = await api.GetTimeAsync(timeout.Token).WaitAsync(timeout.Token).ConfigureAwait(false);
                if (time.IsSuccess && time.Data is not null)
                    status = status with
                    {
                        GameHours = time.Data.Time,
                        Sunrise = time.Data.Sunrise,
                        Sunset = time.Data.Sunset,
                        DayLengthMinutes = time.Data.DayLengthMinutes,
                        TimeScale = time.Data.TimeScale
                    };
            }
            catch (Exception) when (!cancellationToken.IsCancellationRequested) { }
            cancellationToken.ThrowIfCancellationRequested();
            return status;
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { throw; }
        catch (Exception ex)
        {
            throw new ServerConnectionException(ConnectionFailureClassifier.Exception(ex, companionSocket));
        }
    }

    public async ValueTask DisposeAsync()
    {
        if (disposed) return;
        disposed = true;
        api.Disconnected -= OnDisconnected;
        api.OnTeamChatReceived -= OnTeamChat;
        api.OnClanChatReceived -= OnClanChat;
        chat.Writer.TryComplete();
        disconnected.TrySetResult();
        devices.Dispose();
        try { await api.DisconnectAsync(true).WaitAsync(TimeSpan.FromSeconds(2)).ConfigureAwait(false); }
        catch (Exception) { /* Teardown must still release the socket after a failed disconnect. */ }
        finally { api.Dispose(); }
    }

    public async Task<IReadOnlyList<ServerTeamMember>> ReadTeamAsync(CancellationToken ct) => (await ReadTeamSnapshotAsync(ct).ConfigureAwait(false)).Members;

    /// <summary>Reference GetTeamInfoAsync: members with position and spawn/death times, plus member and leader map notes.</summary>
    public async Task<ServerTeamSnapshot> ReadTeamSnapshotAsync(CancellationToken ct)
    {
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(ct);
        timeout.CancelAfter(TimeSpan.FromSeconds(6));
        try
        {
            await requests.WaitAsync(timeout.Token).ConfigureAwait(false);
            var response = await api.GetTeamInfoAsync(timeout.Token).WaitAsync(timeout.Token).ConfigureAwait(false);
            if (!response.IsSuccess || response.Data is null)
                throw new ServerConnectionException(ConnectionFailureClassifier.Response(response.Error?.Message));
            var members = (response.Data.Members ?? []).Select(member => new ServerTeamMember(
                member.SteamId.ToString(System.Globalization.CultureInfo.InvariantCulture), member.Name ?? "", member.IsOnline, member.IsAlive)
            {
                X = member.X,
                Y = member.Y,
                SpawnTime = UnixSeconds(member.LastSpawnTime),
                DeathTime = UnixSeconds(member.LastDeathTime)
            }).ToArray();
            return new ServerTeamSnapshot(members, Notes(response.Data.Notes), Notes(response.Data.LeaderNotes));
        }
        catch (OperationCanceledException) when (ct.IsCancellationRequested) { throw; }
        catch (Exception ex) { throw new ServerConnectionException(ConnectionFailureClassifier.Exception(ex, companionSocket)); }
    }

    private static TeamMapNote[] Notes(IEnumerable<RustPlusApi.Data.Notes.PlayerNote>? notes) =>
        (notes ?? []).Select(note => new TeamMapNote((int)note.Icon, note.X, note.Y, note.Text ?? "")).ToArray();

    /// <summary>Rust+ sends unix seconds where 0 means "not set"; the SDK turns them into DateTime values.</summary>
    internal static long? UnixSeconds(DateTime value)
    {
        var utc = value.Kind == DateTimeKind.Local ? value.ToUniversalTime() : DateTime.SpecifyKind(value, DateTimeKind.Utc);
        long seconds = new DateTimeOffset(utc).ToUnixTimeSeconds();
        return seconds > 0 ? seconds : null;
    }

    /// <summary>Map image and monuments, with the world size from server info (the reference's GetMapWithMonumentsAsync).</summary>
    public async Task<ServerMap> ReadMapAsync(CancellationToken ct)
    {
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(ct);
        // Map payloads are large JPEGs; the reference allowed them a longer budget than status calls.
        timeout.CancelAfter(TimeSpan.FromSeconds(30));
        try
        {
            await requests.WaitAsync(timeout.Token).ConfigureAwait(false);
            var info = await api.GetInfoAsync(timeout.Token).WaitAsync(timeout.Token).ConfigureAwait(false);
            if (!info.IsSuccess || info.Data is null)
                throw new ServerConnectionException(ConnectionFailureClassifier.Response(info.Error?.Message));
            await requests.WaitAsync(timeout.Token).ConfigureAwait(false);
            var map = await api.GetMapAsync(timeout.Token).WaitAsync(timeout.Token).ConfigureAwait(false);
            if (!map.IsSuccess || map.Data?.JpgImage is not { Length: > 0 } image)
                throw new ServerConnectionException(ConnectionFailureClassifier.Response(map.Error?.Message));
            var data = map.Data;
            var monuments = (data.Monuments ?? []).Where(monument => !string.IsNullOrWhiteSpace(monument.Name) && monument.X is not null && monument.Y is not null)
                .Select(monument => new MapMonument(monument.Name!, monument.X!.Value, monument.Y!.Value, monument.IsCustomName == true)).ToArray();
            if (info.Data.MapSize is not { } worldSize || worldSize == 0 || data.Width is not { } width || data.Height is not { } height)
                throw new ServerConnectionException(ConnectionFailure.Unreachable);
            string background = $"#{data.Background.R:X2}{data.Background.G:X2}{data.Background.B:X2}";
            return new ServerMap(image, (int)width, (int)height, (int)(data.OceanMargin ?? 0), worldSize, monuments, background);
        }
        catch (OperationCanceledException) when (ct.IsCancellationRequested) { throw; }
        catch (ServerConnectionException) { throw; }
        catch (Exception ex) { throw new ServerConnectionException(ConnectionFailureClassifier.Exception(ex, companionSocket)); }
    }

    public async Task<IReadOnlyList<MapMarker>> ReadMarkersAsync(CancellationToken ct)
    {
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(ct);
        timeout.CancelAfter(TimeSpan.FromSeconds(8)); // reference marker poll budget
        try
        {
            await requests.WaitAsync(timeout.Token).ConfigureAwait(false);
            var response = await api.GetMapMarkersAsync(timeout.Token).WaitAsync(timeout.Token).ConfigureAwait(false);
            if (!response.IsSuccess || response.Data is null)
                throw new ServerConnectionException(ConnectionFailureClassifier.Response(response.Error?.Message));
            return RustPlusMarkerMapper.Map(response.Data);
        }
        catch (OperationCanceledException) when (ct.IsCancellationRequested) { throw; }
        catch (ServerConnectionException) { throw; }
        catch (Exception ex) { throw new ServerConnectionException(ConnectionFailureClassifier.Exception(ex, companionSocket)); }
    }
}
