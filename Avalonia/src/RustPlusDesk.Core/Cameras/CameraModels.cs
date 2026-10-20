namespace RustPlusDesk.Core.Cameras;

/// <summary>
/// Reference CameraSession: the kind of camera a subscription is. The reference reads it from the camera controller
/// (a drone, an auto-turret, a PTZ camera, a static camera) and falls back to "Camera" for anything else.
/// </summary>
public enum CameraKind
{
    Drone,
    AutoTurret,
    Ptz,
    Static,
    Camera,
}

/// <summary>Reference CameraEntityType: what the camera sees.</summary>
public enum CameraEntityKind
{
    Tree,
    Player,
}

/// <summary>
/// Reference CameraButtons: the keys the camera takes. The numbers are the protocol's own, so a button means the same
/// thing to the server as it does here.
/// </summary>
[Flags]
public enum CameraButtons
{
    None = 0,
    Forward = 2,
    Backward = 4,
    Left = 8,
    Right = 16,
    Jump = 32,
    Duck = 64,
    Sprint = 128,
    Use = 256,
    FirePrimary = 1024,
    FireSecondary = 2048,
    Reload = 8192,
    FireThird = 134217728,
}

/// <summary>Reference CameraControlFlags: what the camera lets the user do.</summary>
[Flags]
public enum CameraControlFlags
{
    None = 0,
    Movement = 1,
    Mouse = 2,
    SprintAndDuck = 4,
    Fire = 8,
    Reload = 16,
    Crosshair = 32,
}

/// <summary>Reference CameraInfo: what a camera is and what it can do.</summary>
public sealed record CameraInfo(
    string Id,
    string Name,
    CameraKind Kind,
    int Width,
    int Height,
    CameraControlFlags ControlFlags)
{
    /// <summary>Whether the camera can be moved with the movement pad.</summary>
    public bool CanMove => ControlFlags.HasFlag(CameraControlFlags.Movement);

    /// <summary>Whether the camera can be looked around with the mouse.</summary>
    public bool CanLook => ControlFlags.HasFlag(CameraControlFlags.Mouse);

    /// <summary>Whether the camera can sprint and duck.</summary>
    public bool CanSprintAndDuck => ControlFlags.HasFlag(CameraControlFlags.SprintAndDuck);

    /// <summary>Whether the camera can fire.</summary>
    public bool CanFire => ControlFlags.HasFlag(CameraControlFlags.Fire);

    /// <summary>Whether the camera can reload.</summary>
    public bool CanReload => ControlFlags.HasFlag(CameraControlFlags.Reload);

    /// <summary>Whether the camera draws a crosshair.</summary>
    public bool CanCrosshair => ControlFlags.HasFlag(CameraControlFlags.Crosshair);
}

/// <summary>Reference CameraEntity: something the camera sees.</summary>
public sealed record CameraEntity(CameraEntityKind Kind, string Name, double X, double Y, double Z);

/// <summary>
/// Reference CameraSession: one live camera.
///
/// The server tracks one camera subscription per connection, so a session owns a connection of its own; the frames
/// are rendered by the protocol's own renderer, so what the port shows is what the server sends.
/// </summary>
public interface ICameraSession : IAsyncDisposable
{
    /// <summary>The camera this session is subscribed to.</summary>
    string CameraId { get; }

    /// <summary>What the camera is and what it can do.</summary>
    CameraInfo Info { get; }

    /// <summary>Reference CameraSession.TargetFps: how many frames a second are rendered.</summary>
    int TargetFps { get; set; }

    /// <summary>Reference CameraSession.FrameRendered: a rendered frame, as PNG bytes.</summary>
    event Action<byte[]>? FrameRendered;

    /// <summary>Reference CameraSession.EntitiesUpdated: what the camera sees and the vertical field of view.</summary>
    event Action<IReadOnlyList<CameraEntity>, double>? EntitiesUpdated;

    /// <summary>Reference CameraSession.KeepAliveFailed: the camera is gone or the connection dropped.</summary>
    event Action<string>? KeepAliveFailed;

    /// <summary>Reference CameraSession.SendInputAsync: the keys and the mouse movement.</summary>
    Task SendInputAsync(CameraButtons buttons, float mouseDeltaX, float mouseDeltaY, CancellationToken ct);

    /// <summary>Reference CameraSession.LookAsync: the mouse-look pan and tilt.</summary>
    Task LookAsync(float deltaX, float deltaY, CancellationToken ct);

    /// <summary>Reference CameraSession.ShootAsync: fire (an auto-turret).</summary>
    Task ShootAsync(CancellationToken ct);

    /// <summary>Reference CameraSession.ReloadAsync: reload (an auto-turret).</summary>
    Task ReloadAsync(CancellationToken ct);

    /// <summary>Reference CameraSession.ZoomAsync: zoom (a PTZ camera).</summary>
    Task ZoomAsync(CancellationToken ct);
}

/// <summary>
/// Reference CameraSession.StartAsync: a camera is opened on a connection of its own. The server keeps one camera
/// subscription per connection, so several cameras at once need several connections.
/// </summary>
public interface ICameraSessionFactory
{
    /// <summary>Opens the camera, or throws when the connection or the subscription fails.</summary>
    Task<ICameraSession> OpenAsync(string profilesPath, string profileKey, string cameraId, CancellationToken ct);
}

/// <summary>One camera the user added: its identifier and the name it is shown with.</summary>
public sealed record CameraEntry(string Id, string Name)
{
    /// <summary>Reference CamDisplayName: the name the user gave it, or the identifier.</summary>
    public string DisplayName => string.IsNullOrWhiteSpace(Name) ? Id : Name;
}

/// <summary>
/// Reference ServerProfile.CameraIds and CameraNames: the cameras the user added for a server. They live in the
/// profile, so the WPF app and this one show the same cameras.
/// </summary>
public interface ICameraStore
{
    /// <summary>The cameras the user added for the server.</summary>
    Task<IReadOnlyList<CameraEntry>> LoadAsync(string profilesPath, string profileKey, CancellationToken ct);

    /// <summary>The cameras are written back to the profile.</summary>
    Task SaveAsync(string profilesPath, string profileKey, IReadOnlyList<CameraEntry> cameras, CancellationToken ct);
}

/// <summary>
/// Reference GrabCameraSnapshotAsync: a single picture of a camera, for the tile.
///
/// The reference opens the camera for a moment, lets a few frames accumulate - the picture sharpens over the frames -
/// and takes the last one.
/// </summary>
public interface ICameraSnapshotSource
{
    /// <summary>The camera's picture, or null when the camera could not be reached.</summary>
    Task<byte[]?> GrabAsync(string profilesPath, string profileKey, string cameraId, CancellationToken ct);
}
