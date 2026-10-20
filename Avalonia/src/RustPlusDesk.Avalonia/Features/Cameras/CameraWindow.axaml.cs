using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using RustPlusDesk.Core.Cameras;

namespace RustPlusDesk.Avalonia.Features.Cameras;

/// <summary>
/// Reference CameraWindow: one camera, live.
///
/// The camera is opened on a connection of its own, the frames are rendered by the protocol's renderer and the
/// picture is shown as it comes. The controls under the picture are the ones the camera supports: the stick moves a
/// drone or looks around, the drone goes up and down, and the auto-turret fires and reloads.
/// </summary>
public sealed partial class CameraWindow : Window
{
    private const double StickRadius = 59;
    private const double StickCentre = 59;

    private ICameraSessionFactory? sessions;
    private string profilesPath = "";
    private string profileKey = "";
    private string cameraId = "";
    private Func<IReadOnlySet<string>>? teamNames;
    private ICameraSession? session;
    private CancellationTokenSource? continuous;
    private bool stickActive;
    private double stickX;
    private double stickY;
    private int frameWidth;
    private int frameHeight;
    private double verticalFov = 65;
    private bool crosshairSupported;
    private IReadOnlyList<CameraEntity> entities = [];

    /// <summary>Reference RefreshTeamAsync: the names of the team players, whose names are drawn green.</summary>
    private IReadOnlySet<string> TeamNames
    {
        get
        {
            try { return teamNames?.Invoke() ?? new HashSet<string>(StringComparer.OrdinalIgnoreCase); }
            catch { return new HashSet<string>(StringComparer.OrdinalIgnoreCase); }
        }
    }

    /// <summary>The window the XAML loader builds; it shows no camera until one is opened in it.</summary>
    public CameraWindow() => InitializeComponent();

    public CameraWindow(ICameraSessionFactory sessions, string profilesPath, string profileKey, string cameraId, string name,
        Func<IReadOnlySet<string>>? teamNames = null)
    {
        InitializeComponent();
        this.sessions = sessions;
        this.profilesPath = profilesPath;
        this.profileKey = profileKey;
        this.cameraId = cameraId;
        this.teamNames = teamNames;
        TitleText.Text = name;
        Opened += OnOpened;
        Closed += OnClosed;
        FpsBox.SelectionChanged += (_, _) => ApplyFps();
    }

    private async void OnOpened(object? sender, EventArgs e)
    {
        if (sessions is null) return;
        try
        {
            session = await sessions.OpenAsync(profilesPath, profileKey, cameraId, CancellationToken.None);
            session.FrameRendered += OnFrame;
            session.EntitiesUpdated += OnEntities;
            session.KeepAliveFailed += OnKeepAliveFailed;
            Configure(session.Info);
            ApplyFps();
            StatusText.Text = "Live";
            Banner.IsVisible = false;
        }
        catch (Exception ex)
        {
            BannerText.Text = ex.Message;
            Banner.IsVisible = true;
            StatusText.Text = "Not connected";
        }
    }

    /// <summary>Reference ConfigureHud: the controls the camera supports.</summary>
    private void Configure(CameraInfo info)
    {
        frameWidth = info.Width;
        frameHeight = info.Height;
        bool movement = info.Kind == CameraKind.Drone || info.CanMove;
        bool look = info.CanLook;
        bool vertical = info.ControlFlags.HasFlag(CameraControlFlags.SprintAndDuck);
        bool reload = info.CanReload;

        Crosshair.IsVisible = info.CanCrosshair;
        crosshairSupported = info.CanCrosshair;
        MovementPad.IsVisible = movement || look;
        VerticalPad.IsVisible = vertical;
        ZoomButton.IsVisible = info.Kind == CameraKind.Ptz;
        FireButton.IsVisible = info.Kind == CameraKind.AutoTurret;
        ReloadButton.IsVisible = reload;
        Controls.IsVisible = movement || look || vertical || reload || info.Kind is CameraKind.Ptz or CameraKind.AutoTurret;

        KindText.Text = info.Kind switch
        {
            CameraKind.Drone => "Drone",
            CameraKind.AutoTurret => "Auto-turret",
            CameraKind.Ptz => "PTZ camera",
            CameraKind.Static => "Static camera",
            _ => "Camera",
        };
        KindBadge.IsVisible = true;

        if (info.Kind == CameraKind.Static)
        {
            BannerText.Text = "Static camera - no controls";
            Banner.IsVisible = true;
        }
    }

    /// <summary>Reference ApplyFps: how many frames a second are rendered.</summary>
    private void ApplyFps()
    {
        if (session is null) return;
        int fps = FpsBox.SelectedIndex switch { 0 => 2, 1 => 5, 2 => 10, _ => 15 };
        session.TargetFps = fps;
    }

    private void OnFrame(byte[] png)
    {
        Dispatcher.UIThread.Post(() =>
        {
            using var stream = new System.IO.MemoryStream(png);
            Bitmap? next;
            try { next = new Bitmap(stream); }
            catch { return; }
            var previous = Picture.Source as IDisposable;
            Picture.Source = next;
            previous?.Dispose();
        });
    }

    private void OnEntities(IReadOnlyList<CameraEntity> found, double fov)
    {
        entities = found;
        if (fov > 0) verticalFov = fov;
        Dispatcher.UIThread.Post(DrawOverlay);
    }

    private void OnKeepAliveFailed(string message)
    {
        Dispatcher.UIThread.Post(() =>
        {
            BannerText.Text = message;
            Banner.IsVisible = true;
            StatusText.Text = "No signal";
        });
    }

    /// <summary>Reference DrawOverlay: the name of every player the camera sees, drawn where the player is.</summary>
    private void DrawOverlay()
    {
        Overlay.Children.Clear();
        if (entities.Count == 0 || Picture.Source is null) return;
        double viewWidth = Picture.Bounds.Width;
        double viewHeight = Picture.Bounds.Height;
        if (viewWidth <= 1 || viewHeight <= 1 || frameWidth <= 0 || frameHeight <= 0) return;

        double scale = Math.Min(viewWidth / frameWidth, viewHeight / frameHeight);
        double offsetX = (Overlay.Bounds.Width - frameWidth * scale) / 2.0;
        double offsetY = (Overlay.Bounds.Height - frameHeight * scale) / 2.0;

        foreach (var entity in entities)
        {
            if (entity.Kind != CameraEntityKind.Player && string.IsNullOrWhiteSpace(entity.Name)) continue;
            if (CameraOverlay.Place(entity, verticalFov, frameWidth, frameHeight) is not { } place) continue;

            // Reference DrawOverlay: the team is green, everybody else is red.
            bool isTeam = !string.IsNullOrWhiteSpace(entity.Name) && TeamNames.Contains(entity.Name);
            var name = new TextBlock
            {
                Text = string.IsNullOrWhiteSpace(entity.Name) ? "player" : entity.Name,
                FontSize = 14,
                Foreground = new global::Avalonia.Media.SolidColorBrush(
                    isTeam ? global::Avalonia.Media.Color.FromRgb(50, 205, 50) : global::Avalonia.Media.Color.FromRgb(255, 69, 0)),
                Background = new global::Avalonia.Media.SolidColorBrush(global::Avalonia.Media.Color.FromArgb(120, 0, 0, 0)),
                Padding = new global::Avalonia.Thickness(4, 1, 4, 1),
            };
            name.Measure(new global::Avalonia.Size(double.PositiveInfinity, double.PositiveInfinity));
            var size = name.DesiredSize;
            Canvas.SetLeft(name, offsetX + place.X * scale - size.Width / 2.0);
            Canvas.SetTop(name, offsetY + (place.Y - 12) * scale - size.Height / 2.0);
            Overlay.Children.Add(name);
        }
    }

    // ---- the stick -----------------------------------------------------------------------------

    private void OnStickPressed(object? sender, PointerPressedEventArgs e)
    {
        if (session is null) return;
        stickActive = true;
        e.Pointer.Capture(MovementPad);
        UpdateStick(e.GetPosition(MovementPad));
        StartContinuous(StickTick);
    }

    private void OnStickMoved(object? sender, PointerEventArgs e)
    {
        if (!stickActive) return;
        UpdateStick(e.GetPosition(MovementPad));
    }

    private void OnStickReleased(object? sender, PointerReleasedEventArgs e)
    {
        if (!stickActive) return;
        stickActive = false;
        e.Pointer.Capture(null);
        stickX = 0;
        stickY = 0;
        Canvas.SetLeft(StickThumb, StickCentre - 24);
        Canvas.SetTop(StickThumb, StickCentre - 24);
        StopContinuous();
    }

    /// <summary>Reference JoyUpdate: the stick's place, clamped to the circle.</summary>
    private void UpdateStick(global::Avalonia.Point point)
    {
        double dx = point.X - StickCentre;
        double dy = point.Y - StickCentre;
        var (x, y) = CameraInput.Stick(dx, dy, StickRadius);
        Canvas.SetLeft(StickThumb, StickCentre - 24 + x * StickRadius);
        Canvas.SetTop(StickThumb, StickCentre - 24 - y * StickRadius);
        stickX = x;
        stickY = y;
    }

    /// <summary>Reference JoystickTick: the camera is told to move, or to look, every tick.</summary>
    private Task StickTick(ICameraSession camera)
    {
        if (camera.Info.CanMove)
        {
            var buttons = CameraInput.MovementButtons(stickX, stickY);
            return camera.SendInputAsync(buttons, 0f, 0f, CancellationToken.None);
        }
        if (camera.Info.CanLook)
            return camera.LookAsync(CameraInput.Response(stickX, 16.0), CameraInput.Response(stickY, 16.0), CancellationToken.None);
        return Task.CompletedTask;
    }

    // ---- the keys ------------------------------------------------------------------------------

    private void OnJump(object? sender, PointerPressedEventArgs e) => Hold(camera => camera.SendInputAsync(CameraButtons.Jump, 0f, 0f, CancellationToken.None));

    private void OnDuck(object? sender, PointerPressedEventArgs e) => Hold(camera => camera.SendInputAsync(CameraButtons.Duck, 0f, 0f, CancellationToken.None));

    private void OnSprint(object? sender, PointerPressedEventArgs e) => Hold(camera => camera.SendInputAsync(CameraButtons.Sprint, 0f, 0f, CancellationToken.None));

    private void OnFire(object? sender, PointerPressedEventArgs e)
    {
        // Reference HookHold: the auto-turret's crosshair is shown while the button is held.
        Crosshair.IsVisible = true;
        Hold(camera => camera.SendInputAsync(CameraButtons.FirePrimary, 0f, 0f, CancellationToken.None));
    }

    /// <summary>Reference HookHold: the button is let go, so the key is let go too.</summary>
    private void OnHoldReleased(object? sender, PointerReleasedEventArgs e) => ReleaseHold();

    /// <summary>Reference HookHold: the pointer leaves the button, so the key is let go too.</summary>
    private void OnHoldExited(object? sender, PointerEventArgs e) => ReleaseHold();

    private void ReleaseHold()
    {
        StopContinuous();
        if (!crosshairSupported) Crosshair.IsVisible = false;
    }

    private async void OnReload(object? sender, PointerPressedEventArgs e)
    {
        if (session is null) return;
        await session.ReloadAsync(CancellationToken.None);
    }

    private async void OnZoom(object? sender, PointerPressedEventArgs e)
    {
        if (session is null) return;
        await session.ZoomAsync(CancellationToken.None);
    }

    /// <summary>Reference StartContinuousButtons: the key is held down, so it is sent again and again.</summary>
    private void Hold(Func<ICameraSession, Task> action) => StartContinuous(action);

    /// <summary>Reference StartContinuous: the action runs every hundred milliseconds until it is stopped.</summary>
    private void StartContinuous(Func<ICameraSession, Task> action)
    {
        StopContinuous();
        var camera = session;
        if (camera is null) return;

        var cts = new CancellationTokenSource();
        continuous = cts;
        var token = cts.Token;
        _ = Task.Run(async () =>
        {
            try
            {
                while (!token.IsCancellationRequested)
                {
                    await action(camera);
                    await Task.Delay(100, token);
                }
            }
            catch (OperationCanceledException) { }
            finally
            {
                try { await camera.SendInputAsync(CameraButtons.None, 0f, 0f, CancellationToken.None); } catch { }
            }
        }, token);
    }

    /// <summary>Reference StopContinuous: the key is let go.</summary>
    private void StopContinuous()
    {
        var cts = continuous;
        continuous = null;
        if (cts is null) return;
        try { cts.Cancel(); } catch { }
        try { cts.Dispose(); } catch { }
    }

    private async void OnClosed(object? sender, EventArgs e)
    {
        StopContinuous();
        if (session is null) return;
        session.FrameRendered -= OnFrame;
        session.EntitiesUpdated -= OnEntities;
        session.KeepAliveFailed -= OnKeepAliveFailed;
        try { await session.DisposeAsync(); } catch { }
        session = null;
    }
}
