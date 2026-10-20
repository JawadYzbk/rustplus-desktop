using RustPlusApi;
using RustPlusApi.Camera;
using RustPlusApi.Data.Events;
using RustPlusDesk.Core.Cameras;
using RustPlusDesk.Core.Connections;
using RustPlusDesk.Infrastructure.Connections;
using SdkCameraButtons = RustPlusApi.Data.Cameras.CameraButtons;
using SdkCameraControlFlags = RustPlusApi.Data.Cameras.CameraControlFlags;
using SdkCameraEntity = RustPlusApi.Data.Cameras.CameraEntity;
using SdkCameraEntityType = RustPlusApi.Data.Cameras.CameraEntityType;

namespace RustPlusDesk.Infrastructure.Cameras;

/// <summary>
/// Reference CameraSession: one live camera.
///
/// The server tracks one camera subscription per connection, so a session owns a connection of its own and drives
/// the protocol's camera controller for the subscribe, the keep-alive and the input. The frames are rendered by the
/// protocol's own renderer, so what the app shows is what the server sends.
/// </summary>
internal sealed class CameraSession : ICameraSession
{
    private readonly RustPlus api;
    private readonly CameraController controller;
    private readonly CameraRenderer renderer;
    private readonly object renderLock = new();
    private DateTime lastRenderUtc = DateTime.MinValue;
    private volatile bool disposed;

    public CameraSession(RustPlus api, CameraController controller, string name)
    {
        this.api = api;
        this.controller = controller;
        renderer = new CameraRenderer(controller.Info?.Width > 0 ? controller.Info.Width : 1,
            controller.Info?.Height > 0 ? controller.Info.Height : 1);
        Info = new CameraInfo(controller.CameraId, name, Kind(controller), controller.Info?.Width ?? 0,
            controller.Info?.Height ?? 0, (CameraControlFlags)(int)(controller.Info?.ControlFlags ?? SdkCameraControlFlags.None));
        controller.OnFrameReceived += OnFrameReceived;
        controller.OnKeepAliveFailed += OnKeepAliveFailed;
    }

    public string CameraId => controller.CameraId;

    public CameraInfo Info { get; }

    public event Action<byte[]>? FrameRendered;

    public event Action<IReadOnlyList<CameraEntity>, double>? EntitiesUpdated;

    public event Action<string>? KeepAliveFailed;

    /// <summary>Reference CameraSession: the kind of camera, read from the controller.</summary>
    private static CameraKind Kind(CameraController controller) =>
        controller.IsDrone ? CameraKind.Drone
        : controller.IsAutoTurret ? CameraKind.AutoTurret
        : controller.IsPtzCamera ? CameraKind.Ptz
        : controller.IsStaticCamera ? CameraKind.Static
        : CameraKind.Camera;

    private void OnFrameReceived(object? sender, CameraRaysEventArg frame)
    {
        if (disposed) return;

        IReadOnlyList<CameraEntity> entities = frame.Entities is null
            ? []
            : frame.Entities.Select(entity => Entities(entity)).ToArray();
        double verticalFov = frame.VerticalFov;

        byte[]? png = null;
        lock (renderLock)
        {
            if (disposed) return;
            // The rays always accumulate, so the image keeps sharpening over the frames.
            renderer.AddRays(frame);

            var minimumInterval = TimeSpan.FromMilliseconds(1000.0 / Math.Max(1, TargetFps));
            var now = DateTime.UtcNow;
            if (now - lastRenderUtc < minimumInterval) return;
            lastRenderUtc = now;

            try { png = renderer.Render(); } catch { png = null; }
        }

        if (png is not null) FrameRendered?.Invoke(png);
        EntitiesUpdated?.Invoke(entities, verticalFov);
    }

    /// <summary>Reference CameraEntity: what the camera sees.</summary>
    private static Core.Cameras.CameraEntity Entities(SdkCameraEntity entity) =>
        new(entity.Type == SdkCameraEntityType.Player ? CameraEntityKind.Player : CameraEntityKind.Tree,
            entity.Name ?? string.Empty, entity.Position.X, entity.Position.Y, entity.Position.Z);

    private void OnKeepAliveFailed(object? sender, RustPlusApi.Data.ErrorMessage error) =>
        KeepAliveFailed?.Invoke(error?.Message ?? "keep-alive failed");

    /// <summary>Reference CameraSession.TargetFps: how many frames a second are rendered.</summary>
    public int TargetFps { get; set; } = 4;

    public Task SendInputAsync(CameraButtons buttons, float mouseDeltaX, float mouseDeltaY, CancellationToken ct) =>
        disposed ? Task.CompletedTask
            : controller.SendInputAsync((SdkCameraButtons)(int)buttons, mouseDeltaX, mouseDeltaY, ct);

    public Task LookAsync(float deltaX, float deltaY, CancellationToken ct) =>
        disposed ? Task.CompletedTask : controller.LookAsync(deltaX, deltaY, ct);

    public Task ShootAsync(CancellationToken ct) =>
        disposed ? Task.CompletedTask : controller.ShootAsync(ct);

    public Task ReloadAsync(CancellationToken ct) =>
        disposed ? Task.CompletedTask : controller.ReloadAsync(ct);

    public Task ZoomAsync(CancellationToken ct) =>
        disposed ? Task.CompletedTask : controller.ZoomAsync(ct);

    public async ValueTask DisposeAsync()
    {
        if (disposed) return;
        disposed = true;

        controller.OnFrameReceived -= OnFrameReceived;
        controller.OnKeepAliveFailed -= OnKeepAliveFailed;

        try { await controller.DisposeAsync().ConfigureAwait(false); } catch { /* best effort */ }
        try { await api.DisposeAsync().ConfigureAwait(false); } catch { /* best effort */ }
    }
}

/// <summary>
/// Reference CameraSession.StartAsync: the camera is opened on a connection of its own.
///
/// The camera subscription goes over the game server's own connection - the companion socket carries the app's traffic,
/// not the game's - so the direct route is the one that is used.
/// </summary>
public sealed class CameraSessionFactory : ICameraSessionFactory
{
    public async Task<ICameraSession> OpenAsync(string profilesPath, string profileKey, string cameraId, CancellationToken ct)
    {
        var routes = await LegacyConnectionRoutes.ReadAsync(profilesPath, profileKey, ct).ConfigureAwait(false);
        var route = routes.Routes.LastOrDefault()
            ?? throw new ServerConnectionException(ConnectionFailure.ProfileUnavailable);
        var api = new RustPlus(route, new RustPlusSocketOptions { RequestTimeout = TimeSpan.FromSeconds(6) });
        try
        {
            await api.ConnectAsync(ct).ConfigureAwait(false);
            var response = await CameraController.SubscribeAsync(api, cameraId, null, ct).ConfigureAwait(false);
            if (!response.IsSuccess || response.Data is null)
                throw new InvalidOperationException(response.Error?.Message ?? $"Failed to subscribe to camera '{cameraId}'.");
            return new CameraSession(api, response.Data, cameraId);
        }
        catch
        {
            try { await api.DisposeAsync().ConfigureAwait(false); } catch { /* best effort */ }
            throw;
        }
    }
}
