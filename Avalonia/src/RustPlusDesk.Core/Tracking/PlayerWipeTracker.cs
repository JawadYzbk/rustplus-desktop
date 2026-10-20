namespace RustPlusDesk.Core.Tracking;

/// <summary>Reference PlayerActivityState: what the tracked player was doing.</summary>
public enum PlayerActivityState
{
    Moving,
    Stationary,
    Afk,
    Dead,
    Offline,
    Unknown,
}

/// <summary>
/// Reference TrackerLocationType: where the tracked player was. The order matters: the record file stores these as
/// numbers, and the ones that were added later are only ever appended.
/// </summary>
public enum TrackerLocationType
{
    Monument,
    Base,
    Open,
    Unknown,
    Water,
    DeepSea,
}

/// <summary>Reference TrackerLocationTypeText: the wording for people and for the wire.</summary>
public static class TrackerLocationTypeText
{
    /// <summary>For people: "Deep Sea" rather than the identifier.</summary>
    public static string Label(this TrackerLocationType type) =>
        type == TrackerLocationType.DeepSea ? "Deep Sea" : type.ToString();

    /// <summary>The backend's and the export's spelling; ToString() would give "deepsea".</summary>
    public static string Wire(this TrackerLocationType type) =>
        type == TrackerLocationType.DeepSea ? "deep_sea" : type.ToString().ToLowerInvariant();
}

/// <summary>Reference PlayerObservation: one poll of a tracked player.</summary>
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

/// <summary>Reference TrackerPersistedObservation: an observation as it is kept on disk.</summary>
public sealed record TrackerPersistedObservation(int SchemaVersion, string Kind, PlayerObservation Observation);

/// <summary>Reference TrackerSegment: a stretch of time the player spent in one state at one place.</summary>
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

/// <summary>Reference MonumentVisit: one stay at a monument.</summary>
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

/// <summary>Reference TrackerSummary: how the player's wipe was spent.</summary>
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

/// <summary>Reference TrackerInsights: the read-at-a-glance intel for one player over the current wipe.</summary>
public sealed record TrackerInsights(
    DateTime? FirstSeenUtc,
    DateTime? LastSeenUtc,
    int SessionCount,
    string? TopMonument,
    TimeSpan TopMonumentDuration,
    int TopMonumentVisits,
    TimeSpan LongestBlindGap,
    DateTime? LongestBlindGapStartUtc,
    int? PeakHourLocal,
    TimeSpan PeakHourActive,
    PlayerActivityState CurrentState,
    TrackerLocationType CurrentLocationType,
    string? CurrentLocationName,
    string? CurrentGrid,
    DateTime? CurrentAsOfUtc,
    bool IsLikelyOnline)
{
    public static TrackerInsights Empty { get; } = new(
        null, null, 0, null, TimeSpan.Zero, 0, TimeSpan.Zero, null, null,
        TimeSpan.Zero, PlayerActivityState.Unknown, TrackerLocationType.Unknown, null, null, null, false);
}

/// <summary>Reference TrackerPoint: one point of the player's route, as the map draws it.</summary>
public sealed record TrackerPoint(
    DateTime TimestampUtc,
    double X,
    double Y,
    PlayerActivityState State,
    TrackerLocationType LocationType,
    string? LocationName,
    string? Grid,
    string? Event,
    string SessionId,
    /// <summary>
    /// True when the player did not travel here from the previous point (death, respawn, Deep Sea, a gap); routes
    /// break here and the step is not counted as distance.
    /// </summary>
    bool BreakBefore = false);

/// <summary>Reference TrackerMonument: a single monument on the wipe map.</summary>
public sealed record TrackerMonument(string Name, double X, double Y, string? Size = null);

/// <summary>Reference TrackerWipeMap: the map of the wipe the tracker draws on.</summary>
public sealed record TrackerWipeMap(
    byte[] PngBytes,
    int WorldSize,
    double WorldRectX,
    double WorldRectY,
    double WorldRectWidth,
    double WorldRectHeight,
    double OceanMargin = 0,
    IReadOnlyList<TrackerMonument>? Monuments = null);

/// <summary>Reference StoredWipeMetadata: what the store keeps about a wipe.</summary>
public sealed record StoredWipeMetadata(
    string ServerKey,
    string ServerName,
    string WipeKey,
    DateTime? WipeStartedAtUtc,
    DateTime CreatedAtUtc);

/// <summary>Reference StoredWipeSummary: what the wipe holds, for the list of wipes.</summary>
public sealed record StoredWipeSummary(
    string ServerKey,
    string ServerName,
    string WipeKey,
    DateTime? WipeStartedAtUtc,
    DateTime? LastObservedAtUtc,
    int PlayerCount,
    long TotalObservations,
    long StoredBytes,
    bool HasMap);
