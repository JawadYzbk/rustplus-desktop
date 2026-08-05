using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace RustPlusDesk.Services.PlayerWipeTracker;

/// <summary>Coordinates the existing team snapshots, pure engines, and local JSONL storage.</summary>
public sealed class PlayerWipeTrackerService : IAsyncDisposable
{
    private readonly PlayerWipeTrackerStore _store;
    private readonly PlayerWipeTrackerCapabilityService _capabilities;
    private readonly LaravelPlayerWipeTrackerClient _cloudClient = new();
    private readonly PlayerWipeTrackerCloudSyncQueue _cloudQueue;
    private readonly Dictionary<ulong, PlayerWipeTrackerEngine> _engines = new();
    private string? _serverKey;
    private string? _wipeKey;
    private string? _sessionId;
    private ulong _ownSteamId;
    private DateTime? _wipeStartedAtUtc;

    public PlayerWipeTrackerService(PlayerWipeTrackerStore store, PlayerWipeTrackerCapabilityService capabilities)
    {
        _store = store ?? throw new ArgumentNullException(nameof(store));
        _capabilities = capabilities ?? throw new ArgumentNullException(nameof(capabilities));
        _cloudQueue = new PlayerWipeTrackerCloudSyncQueue((request, cancellationToken) => _cloudClient.PutDayAsync(request, cancellationToken));
    }

    public bool Enabled { get; set; }
    public bool CloudBackupEnabled { get; set; }
    public string? CurrentWipeKey => _wipeKey;
    public string? CurrentSessionId => _sessionId;
    public IReadOnlyCollection<ulong> TrackedPlayers => _engines.Keys.ToArray();
    public PlayerWipeTrackerCapabilities Capabilities => _capabilities.Current;

    public void UpdateCapabilities(JsonElement bootstrap) => _capabilities.Update(bootstrap);

    public void StartConnection(string serverKey, DateTime? wipeTimeUtc, string? mapIdentity, ulong ownSteamId, string? sessionId = null)
    {
        _serverKey = serverKey;
        _wipeKey = BuildWipeKey(serverKey, wipeTimeUtc, mapIdentity);
        _wipeStartedAtUtc = wipeTimeUtc?.ToUniversalTime();
        _sessionId = string.IsNullOrWhiteSpace(sessionId) ? Guid.NewGuid().ToString("N") : sessionId;
        _ownSteamId = ownSteamId;
        _engines.Clear();
        foreach (var steamId in _store.LoadPlayerIds(serverKey, _wipeKey))
        {
            if (_capabilities.Current.CanTrackPlayer(steamId, ownSteamId))
                _engines[steamId] = LoadEngine(steamId);
        }
    }

    public void Observe(PlayerObservation observation)
    {
        if (!Enabled || _serverKey is null || _wipeKey is null ||
            !_capabilities.Current.IsTrackerAvailable ||
            !_capabilities.Current.CanTrackPlayer(observation.SteamId, _ownSteamId))
            return;

        if (_sessionId is null)
            _sessionId = observation.SessionId;

        if (!_engines.TryGetValue(observation.SteamId, out var engine))
        {
            engine = LoadEngine(observation.SteamId);
            _engines[observation.SteamId] = engine;
        }

        if (!engine.Observe(observation with { SessionId = _sessionId }))
            return;

        _store.Append(_serverKey, _wipeKey, observation.SteamId,
            new TrackerPersistedObservation(1, "observation", observation with { SessionId = _sessionId }));

        if (CloudBackupEnabled && _capabilities.Current.CanUseCloudSync)
            _ = QueueCloudDayAsync(observation.SteamId, observation.TimestampUtc, observation.Name);
    }

    public void Disconnect(DateTime? timestampUtc = null)
    {
        var timestamp = (timestampUtc ?? DateTime.UtcNow).ToUniversalTime();
        foreach (var (steamId, engine) in _engines)
        {
            var last = engine.LastObservation;
            engine.EndSession(timestamp);
            if (_serverKey is not null && _wipeKey is not null && last is not null && timestamp > last.TimestampUtc)
            {
                _store.Append(_serverKey, _wipeKey, steamId,
                    new TrackerPersistedObservation(1, "observation", last with
                    {
                        TimestampUtc = timestamp,
                        IsConnected = false,
                        SnapshotValid = false,
                    }));
            }
        }
        _sessionId = null;
    }

    public TrackerSummary GetSummary(ulong steamId)
        => _engines.TryGetValue(steamId, out var engine) ? engine.Summarize() : new TrackerSummary(TimeSpan.Zero, TimeSpan.Zero, TimeSpan.Zero, TimeSpan.Zero, TimeSpan.Zero, TimeSpan.Zero, TimeSpan.Zero, 0, 0, Array.Empty<MonumentVisit>());

    public IReadOnlyList<TrackerSegment> GetSegments(ulong steamId)
        => _engines.TryGetValue(steamId, out var engine) ? engine.Segments : Array.Empty<TrackerSegment>();

    public long StorageBytes => _store.StorageBytes;

    public CloudDayUploadRequest? BuildCloudDay(ulong steamId, DateOnly day, string? playerName)
    {
        if (_serverKey is null || _wipeKey is null)
            return null;

        var observations = _store.Load(_serverKey, _wipeKey, steamId)
            .Where(item => item.Kind == "observation" && DateOnly.FromDateTime(item.Observation.TimestampUtc.ToUniversalTime()) == day)
            .Select(item => item.Observation)
            .OrderBy(item => item.TimestampUtc)
            .ToArray();
        if (observations.Length == 0)
            return null;

        var cloud = new List<CloudTrackerObservation>(observations.Length);
        PlayerObservation? previous = null;
        foreach (var observation in observations)
        {
            var displacement = previous is null || previous.X is null || previous.Y is null || observation.X is null || observation.Y is null
                ? 0
                : Math.Sqrt(Math.Pow(previous.X.Value - observation.X.Value, 2) + Math.Pow(previous.Y.Value - observation.Y.Value, 2));
            var continuity = previous is not null && previous.SessionId == observation.SessionId &&
                observation.TimestampUtc > previous.TimestampUtc &&
                (observation.TimestampUtc - previous.TimestampUtc).TotalSeconds <= PlayerWipeTrackerEngine.MaxContinuityGapSeconds &&
                previous.IsConnected && previous.SnapshotValid && observation.IsConnected && observation.SnapshotValid;
            var state = !continuity && previous is not null ? PlayerActivityState.Unknown :
                PlayerWipeTrackerEngine.Classify(observation, displacement);
            var eventName = previous is not null && !previous.Dead && observation.Dead ? "death" :
                previous is not null && previous.Dead && !observation.Dead ? "respawn" : null;
            cloud.Add(new CloudTrackerObservation
            {
                Timestamp = observation.TimestampUtc.ToUniversalTime().ToString("O"),
                X = observation.X,
                Y = observation.Y,
                State = state.ToString().ToLowerInvariant(),
                LocationType = observation.LocationType.ToString().ToLowerInvariant(),
                LocationName = observation.LocationName,
                Grid = observation.Grid,
                Event = eventName,
            });
            previous = observation;
        }

        var payload = new CloudTrackerDayPayload
        {
            GeneratedAt = DateTime.UtcNow.ToString("O"),
            ObservationSessions = observations.Select(item => item.SessionId).Distinct(StringComparer.Ordinal).ToArray(),
            Observations = cloud,
        };
        var json = System.Text.Json.JsonSerializer.Serialize(payload, new JsonSerializerOptions(JsonSerializerDefaults.Web));
        return new CloudDayUploadRequest
        {
            ServerKey = _serverKey,
            WipeKey = _wipeKey,
            WipeStartedAt = _wipeStartedAtUtc?.ToString("O"),
            PlayerSteamId = steamId.ToString(),
            PlayerName = playerName,
            Day = day.ToString("yyyy-MM-dd"),
            Payload = payload,
            Checksum = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(json))).ToLowerInvariant(),
        };
    }

    public void DeleteWipe(string serverKey, string wipeKey) => _store.DeleteWipe(serverKey, wipeKey);
    public void DeleteAll() => _store.DeleteAll();

    public static string BuildWipeKey(string serverKey, DateTime? wipeTimeUtc, string? mapIdentity)
    {
        var normalized = wipeTimeUtc.HasValue
            ? wipeTimeUtc.Value.ToUniversalTime().ToString("O")
            : "unknown";
        var source = $"{serverKey.Trim()}|{normalized}|{mapIdentity?.Trim() ?? "unknown"}";
        return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(source))).ToLowerInvariant();
    }

    private PlayerWipeTrackerEngine LoadEngine(ulong steamId)
    {
        var engine = new PlayerWipeTrackerEngine();
        if (_serverKey is null || _wipeKey is null)
            return engine;
        foreach (var item in _store.Load(_serverKey, _wipeKey, steamId).Where(x => x.Kind == "observation"))
            engine.Observe(item.Observation);
        return engine;
    }

    private async Task QueueCloudDayAsync(ulong steamId, DateTime timestampUtc, string playerName)
    {
        try
        {
            await _store.FlushAsync().ConfigureAwait(false);
            var request = BuildCloudDay(steamId, DateOnly.FromDateTime(timestampUtc.ToUniversalTime()), playerName);
            if (request is not null)
                _cloudQueue.Enqueue(request);
        }
        catch { }
    }

    public async ValueTask DisposeAsync()
    {
        try { await _store.FlushAsync().ConfigureAwait(false); } catch { }
        await _cloudQueue.DisposeAsync().ConfigureAwait(false);
        await _store.DisposeAsync().ConfigureAwait(false);
    }
}
