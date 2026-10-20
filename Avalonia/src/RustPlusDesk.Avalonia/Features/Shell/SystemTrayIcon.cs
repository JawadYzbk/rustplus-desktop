using RustPlusDesk.Core.Platform;
using AvaloniaApplication = global::Avalonia.Application;
using AvaloniaTrayIcon = global::Avalonia.Controls.TrayIcon;
using AvaloniaTrayIcons = global::Avalonia.Controls.TrayIcons;
using NativeMenu = global::Avalonia.Controls.NativeMenu;
using NativeMenuItem = global::Avalonia.Controls.NativeMenuItem;
using NativeMenuItemBase = global::Avalonia.Controls.NativeMenuItemBase;
using NativeMenuItemSeparator = global::Avalonia.Controls.NativeMenuItemSeparator;
using WindowIcon = global::Avalonia.Controls.WindowIcon;

namespace RustPlusDesk.Avalonia.Features.Shell;

/// <summary>
/// The app's icon in the operating system's notification area, drawn by Avalonia's own cross-platform tray icon.
///
/// The reference builds the same thing with Windows Forms, which is Windows only; Avalonia's tray icon talks to the
/// notification area of Windows, Linux and macOS in the same way, so the port needs no P/Invoke of its own.
///
/// The icon itself is made when the app first asks for it, so nothing here needs a desktop to be built or tested.
/// </summary>
public sealed class SystemTrayIcon : ITrayIcon
{
    private readonly AvaloniaTrayIcons icons;
    private readonly Func<AvaloniaTrayIcon> factory;
    private AvaloniaTrayIcon? icon;
    private NativeMenu? menu;
    private WindowIcon? picture;
    private MemoryStream? pictureStream;
    private byte[]? pictureBytes;
    private bool disposed;

    /// <summary>Attaches a tray icon to the running app; the platform's own icon is made when the app first uses it.</summary>
    public SystemTrayIcon(AvaloniaApplication application) : this(application, () => new AvaloniaTrayIcon()) { }

    /// <summary>
    /// The same, with the platform's tray icon made by the given factory. A caller that has no desktop - a test, say -
    /// can leave the factory alone and never draw anything.
    /// </summary>
    public SystemTrayIcon(AvaloniaApplication application, Func<AvaloniaTrayIcon> factory)
    {
        this.factory = factory;
        icons = AvaloniaTrayIcon.GetIcons(application) ?? new AvaloniaTrayIcons();
        if (AvaloniaTrayIcon.GetIcons(application) is null) AvaloniaTrayIcon.SetIcons(application, icons);
    }

    /// <summary>Whether the icon is in the notification area at all.</summary>
    public bool IsVisible
    {
        get => icon?.IsVisible ?? false;
        set { if (Ensure() is { } tray) tray.IsVisible = value; }
    }

    /// <summary>The text shown while the pointer rests on the icon.</summary>
    public string? Tooltip
    {
        get => icon?.ToolTipText;
        set { if (Ensure() is { } tray) tray.ToolTipText = value; }
    }

    /// <summary>The picture of the icon, so the tray shows what the window shows.</summary>
    public byte[]? Icon
    {
        get => pictureBytes;
        set
        {
            pictureBytes = value;
            pictureStream?.Dispose();
            pictureStream = null;
            picture = null;
            if (value is not null)
            {
                // The platform reads the picture while the icon is set, so the stream is kept alive with it.
                var stream = new MemoryStream(value, writable: false);
                try
                {
                    picture = new WindowIcon(stream);
                    pictureStream = stream;
                }
                catch (Exception ex) when (ex is IOException or ArgumentException or NotSupportedException)
                {
                    stream.Dispose();
                }
            }
            if (icon is { } tray) tray.Icon = picture;
        }
    }

    /// <summary>The last message the app asked to be shown in a balloon.</summary>
    public (string Title, string Message)? Balloon { get; private set; }

    /// <summary>
    /// Tells the user something without opening the app. Avalonia's cross-platform tray icon draws no balloon of its
    /// own, so the message is kept here for the app to show in a window of its own.
    /// </summary>
    public void ShowBalloon(string title, string message) => Balloon = (title, message);

    /// <summary>Puts the menu in the tray icon, replacing the one that was there.</summary>
    public void ShowContextMenu(IReadOnlyList<TrayMenuItem> items)
    {
        if (Ensure() is not { } tray) return;
        // A new menu is made every time, so a menu that is open keeps the lines it was opened with.
        var built = new NativeMenu();
        foreach (var item in items) built.Add(Build(item));
        // The platform says when it is about to show the menu; that is the moment the app can fill it again.
        built.Opening += (_, _) => ContextMenuRequested?.Invoke(this, EventArgs.Empty);
        menu = built;
        tray.Menu = built;
    }

    /// <summary>The user clicked the icon.</summary>
    public event EventHandler? Clicked;

    /// <summary>The platform is about to show the menu, so the app can fill it with what is true now.</summary>
    public event EventHandler? ContextMenuRequested;

    /// <summary>The user chose a line of the menu.</summary>
    public event EventHandler<TrayMenuItem>? MenuItemInvoked;

    /// <summary>Takes the icon out of the notification area.</summary>
    public void Dispose()
    {
        if (disposed) return;
        disposed = true;
        if (icon is not null)
        {
            icons.Remove(icon);
            icon.Dispose();
            icon = null;
        }
        pictureStream?.Dispose();
        pictureStream = null;
        menu = null;
    }

    /// <summary>One line of the menu, as the platform draws it.</summary>
    private NativeMenuItemBase Build(TrayMenuItem item)
    {
        if (item.IsSeparator) return new NativeMenuItemSeparator();
        var entry = new NativeMenuItem(item.Header) { IsEnabled = item.IsEnabled };
        if (item.Items.Count > 0)
        {
            var submenu = new NativeMenu();
            foreach (var child in item.Items) submenu.Add(Build(child));
            entry.Menu = submenu;
            return entry;
        }
        // The platform only offers a line to the user when it has a handler, so the click is what carries the choice.
        entry.Click += (_, _) => MenuItemInvoked?.Invoke(this, item);
        return entry;
    }

    /// <summary>The platform's tray icon, made the first time the app asks for one.</summary>
    private AvaloniaTrayIcon? Ensure()
    {
        if (disposed) return null;
        if (icon is not null) return icon;
        var tray = factory();
        tray.Clicked += (_, _) => Clicked?.Invoke(this, EventArgs.Empty);
        tray.Icon = picture;
        if (menu is not null) tray.Menu = menu;
        icon = tray;
        try { icons.Add(tray); }
        catch (Exception ex) when (ex is InvalidOperationException or PlatformNotSupportedException) { }
        return tray;
    }
}
