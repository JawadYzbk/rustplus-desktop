using System;
using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace RustPlusDesk.Services.PlayerWipeTracker;

public enum PlayerActivityState
{
    Moving,
    Stationary,
    Afk,
    Dead,
    Offline,
    Unknown,
}

public enum TrackerLocationType
{
    Monument,
    Base,
    Open,
    Unknown,
}

public sealed record PlayerObservation(
    ulong SteamId,
    string Name,
    DateTime TimestampUtc,
    string SessionId,
    bool IsConnected,
    bool SnapshotValid,
    bool Online,
    bool Dead,
    bool Afk,
    double? X,
    double? Y,
    TrackerLocationType LocationType,
    string? LocationName,
    string? Grid,
    long? SpawnTime,
    long? DeathTime);

public sealed record TrackerSegment(
    DateTime StartUtc,
    DateTime EndUtc,
    PlayerActivityState State,
    TrackerLocationType LocationType,
    string? LocationName,
    string? Grid,
    double? StartX,
    double? StartY,
    double? EndX,
    double? EndY,
    string SessionId);

public sealed record MonumentVisit(
    string Name,
    DateTime StartUtc,
    DateTime EndUtc,
    double? EntryX,
    double? EntryY,
    double? ExitX,
    double? ExitY)
{
    public TimeSpan EstimatedDuration => EndUtc - StartUtc;
}

public sealed record TrackerSummary(
    TimeSpan Coverage,
    TimeSpan Unknown,
    TimeSpan Moving,
    TimeSpan Stationary,
    TimeSpan Afk,
    TimeSpan Dead,
    TimeSpan Offline,
    double EstimatedDistance,
    int Deaths,
    IReadOnlyList<MonumentVisit> MonumentVisits);

public sealed record TrackerPersistedObservation(
    int SchemaVersion,
    string Kind,
    PlayerObservation Observation);

public sealed class CloudTrackerDayPayload
{
    [JsonPropertyName("schema_version")] public int SchemaVersion { get; init; } = 1;
    [JsonPropertyName("generated_at")] public string GeneratedAt { get; init; } = string.Empty;
    [JsonPropertyName("observation_sessions")] public IReadOnlyList<string> ObservationSessions { get; init; } = Array.Empty<string>();
    [JsonPropertyName("observations")] public IReadOnlyList<CloudTrackerObservation> Observations { get; init; } = Array.Empty<CloudTrackerObservation>();
}

public sealed class CloudTrackerObservation
{
    [JsonPropertyName("timestamp")] public string Timestamp { get; init; } = string.Empty;
    [JsonPropertyName("x")] public double? X { get; init; }
    [JsonPropertyName("y")] public double? Y { get; init; }
    [JsonPropertyName("state")] public string State { get; init; } = "unknown";
    [JsonPropertyName("location_type")] public string LocationType { get; init; } = "unknown";
    [JsonPropertyName("location_name")] public string? LocationName { get; init; }
    [JsonPropertyName("grid")] public string? Grid { get; init; }
    [JsonPropertyName("event")] public string? Event { get; init; }
}

public sealed class CloudDayUploadRequest
{
    [JsonPropertyName("server_key")] public string ServerKey { get; init; } = string.Empty;
    [JsonPropertyName("wipe_key")] public string WipeKey { get; init; } = string.Empty;
    [JsonPropertyName("wipe_started_at")] public string? WipeStartedAt { get; init; }
    [JsonPropertyName("player_steam_id")] public string PlayerSteamId { get; init; } = string.Empty;
    [JsonPropertyName("player_name")] public string? PlayerName { get; init; }
    [JsonPropertyName("day")] public string Day { get; init; } = string.Empty;
    [JsonPropertyName("schema_version")] public int SchemaVersion { get; init; } = 1;
    [JsonPropertyName("payload")] public CloudTrackerDayPayload Payload { get; init; } = new();
    [JsonPropertyName("checksum")] public string Checksum { get; init; } = string.Empty;
}
