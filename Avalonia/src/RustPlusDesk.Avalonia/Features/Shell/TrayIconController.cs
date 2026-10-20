using System.Globalization;
using Avalonia.Platform;
using RustPlusDesk.Core.Localization;
using RustPlusDesk.Core.Platform;

namespace RustPlusDesk.Avalonia.Features.Shell;

/// <summary>
/// The app's side of the system tray: it owns the tray icon, fills its menu with what the app is doing now, and
/// tells the shell what the user chose.
///
/// Reference SetupTrayIcon in App.xaml.cs: the menu holds the tracking status and the time of the last pull, both
/// greyed out because they are there to be read, and below them "Open Rust+ Desk" and "Exit".
/// </summary>
public sealed class TrayIconController : IDisposable
{
    /// <summary>The id of the line that opens the app.</summary>
    public const string OpenId = "open";

    /// <summary>The id of the line that leaves the app.</summary>
    public const string ExitId = "exit";

    /// <summary>
    /// How long the second click may take to follow the first. The reference's tray icon opens the app on a double
    /// click and does nothing on a single one, so the port waits for the second click as well.
    /// </summary>
    public static readonly TimeSpan DoubleClickInterval = TimeSpan.FromMilliseconds(400);

    private readonly ITrayIcon trayIcon;
    private readonly Func<bool> isTracking;
    private DateTime? lastUpdate;
    private DateTime? lastClick;
    private bool disposed;

    /// <summary>Builds the controller over the tray icon the app shows.</summary>
    /// <param name="trayIcon">The icon in the notification area.</param>
    /// <param name="isTracking">Whether the app is following a server right now.</param>
    /// <param name="icon">The picture of the icon; the app's own, so the tray shows what the window shows.</param>
    public TrayIconController(ITrayIcon trayIcon, Func<bool> isTracking, byte[]? icon = null)
    {
        this.trayIcon = trayIcon;
        this.isTracking = isTracking;
        trayIcon.Clicked += OnClicked;
        trayIcon.MenuItemInvoked += OnMenuItemInvoked;
        trayIcon.ContextMenuRequested += OnContextMenuRequested;
        Loc.LanguageChanged += OnLanguageChanged;
        if (icon is not null) trayIcon.Icon = icon;
        Refresh();
        trayIcon.IsVisible = true;
    }

    /// <summary>The menu the reference's tray shows, in the reference's order.</summary>
    public IReadOnlyList<TrayMenuItem> MenuItems { get; private set; } = [];

    /// <summary>The text the reference shows while the pointer rests on the icon.</summary>
    public string Tooltip { get; private set; } = "";

    /// <summary>When the app last heard from the server; the reference's last pull.</summary>
    public DateTime? LastUpdate => lastUpdate;

    /// <summary>The user wants the app in front of them.</summary>
    public event EventHandler? OpenRequested;

    /// <summary>The user wants the app to close.</summary>
    public event EventHandler? ExitRequested;

    /// <summary>
    /// The app's own icon, so the tray shows what the window shows. The reference takes the icon of the running
    /// program; this is the same picture, read from the app's own assets.
    /// </summary>
    public static byte[]? AppIcon()
    {
        try
        {
            var uri = new Uri($"avares://{typeof(TrayIconController).Assembly.GetName().Name}/Assets/app.ico");
            if (!AssetLoader.Exists(uri)) return null;
            using var stream = AssetLoader.Open(uri);
            if (stream is null) return null;
            using var buffer = new MemoryStream();
            stream.CopyTo(buffer);
            return buffer.ToArray();
        }
        catch (Exception ex) when (ex is ArgumentException or InvalidOperationException or IOException or NotSupportedException)
        {
            // An app without its icon still has a tray icon, only without a picture.
            return null;
        }
    }

    /// <summary>
    /// The app heard from the server: the reference remembers the time of its last pull and shows it in the menu.
    /// </summary>
    public void NoteServerUpdate(DateTime? when = null)
    {
        lastUpdate = when ?? DateTime.Now;
        Refresh();
    }

    /// <summary>Fills the menu and the tooltip with what the app is doing now.</summary>
    public void Refresh()
    {
        var items = BuildMenu();
        bool changed = !SameMenu(items, MenuItems);
        MenuItems = items;
        Tooltip = BuildTooltip();
        trayIcon.Tooltip = Tooltip;
        // The menu is only handed over when it says something new, so a menu that is open keeps its own lines.
        if (changed) trayIcon.ShowContextMenu(items);
    }

    /// <summary>Let the tray icon go and stop listening to it.</summary>
    public void Dispose()
    {
        if (disposed) return;
        disposed = true;
        trayIcon.Clicked -= OnClicked;
        trayIcon.MenuItemInvoked -= OnMenuItemInvoked;
        trayIcon.ContextMenuRequested -= OnContextMenuRequested;
        Loc.LanguageChanged -= OnLanguageChanged;
        trayIcon.Dispose();
    }

    /// <summary>The reference's menu: two lines to read, then the two actions, with a separator between each.</summary>
    private IReadOnlyList<TrayMenuItem> BuildMenu() =>
    [
        new TrayMenuItem { Header = Format("TrayTrackingStatus", "Tracking: {0}", TrackingText), IsEnabled = false },
        new TrayMenuItem { Header = Format("TrayLastUpdate", "Last update: {0}", LastUpdateText), IsEnabled = false },
        TrayMenuItem.Separator,
        new TrayMenuItem { Header = Loc.Text("OpenRustPlusDesk", "Open Rust+ Desk"), Id = OpenId },
        TrayMenuItem.Separator,
        new TrayMenuItem { Header = Loc.Text("Exit", "Exit"), Id = ExitId },
    ];

    /// <summary>The reference's tooltip: its own name while the app is idle, and the last pull while it is tracking.</summary>
    private string BuildTooltip() => isTracking()
        ? Format("TrayIconTracking", "Rust+ Desk (Tracking {0})", LastUpdateText)
        : Loc.Text("TrayIconDefault", "Rust+ Desk Tracker");

    /// <summary>The reference says Active while it tracks and Idle while it does not.</summary>
    private string TrackingText => isTracking() ? "Active" : "Idle";

    /// <summary>The time of the last pull, as the reference shows it; "--:--:--" when there has been none.</summary>
    private string LastUpdateText => lastUpdate?.ToString("HH:mm:ss", CultureInfo.CurrentCulture) ?? "--:--:--";

    /// <summary>Whether two menus say the same thing, so the platform is not handed the same menu over and over.</summary>
    private static bool SameMenu(IReadOnlyList<TrayMenuItem> left, IReadOnlyList<TrayMenuItem> right)
    {
        if (left.Count != right.Count) return false;
        for (int index = 0; index < left.Count; index++)
        {
            var one = left[index];
            var other = right[index];
            if (one.Header != other.Header || one.Id != other.Id ||
                one.IsSeparator != other.IsSeparator || one.IsEnabled != other.IsEnabled) return false;
        }
        return true;
    }

    /// <summary>
    /// The reference's wording, read from the reference's own catalogue. A translation without the placeholder is
    /// shown as it stands rather than as an error.
    /// </summary>
    private static string Format(string key, string fallback, object argument)
    {
        string template = Loc.Text(key, fallback);
        try { return string.Format(CultureInfo.CurrentCulture, template, argument); }
        catch (FormatException) { return template; }
    }

    /// <summary>
    /// Reference: the reference's tray icon opens the app on a double click and does nothing on a single one, so the
    /// port waits for the second click as well.
    /// </summary>
    private void OnClicked(object? sender, EventArgs e)
    {
        DateTime now = DateTime.UtcNow;
        if (lastClick is { } first && now - first <= DoubleClickInterval)
        {
            lastClick = null;
            OpenRequested?.Invoke(this, EventArgs.Empty);
            return;
        }
        lastClick = now;
    }

    /// <summary>The user chose a line of the menu; the line's id says what it means.</summary>
    private void OnMenuItemInvoked(object? sender, TrayMenuItem item)
    {
        switch (item.Id)
        {
            case OpenId:
                OpenRequested?.Invoke(this, EventArgs.Empty);
                break;
            case ExitId:
                ExitRequested?.Invoke(this, EventArgs.Empty);
                break;
        }
    }

    /// <summary>The platform is about to show the menu, so it is filled with what is true now.</summary>
    private void OnContextMenuRequested(object? sender, EventArgs e) => Refresh();

    /// <summary>The reference reads its tray tooltip again when the language changes.</summary>
    private void OnLanguageChanged(object? sender, EventArgs e) => Refresh();
}
