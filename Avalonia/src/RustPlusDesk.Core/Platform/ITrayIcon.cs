namespace RustPlusDesk.Core.Platform;

/// <summary>
/// One line of the tray menu: a heading with its own entries, or a separator.
///
/// The reference's tray menu holds two read-only lines that say what the app is doing and the two actions below them;
/// this is the same shape, so the menu the port shows is the menu the reference shows.
/// </summary>
public sealed record TrayMenuItem
{
    /// <summary>A separator line: the reference puts one between the status lines and the actions.</summary>
    public static TrayMenuItem Separator { get; } = new() { IsSeparator = true };

    /// <summary>The text of the line, e.g. "Open Rust+ Desk".</summary>
    public string Header { get; init; } = "";

    /// <summary>The entries of a submenu; empty for an ordinary line.</summary>
    public IReadOnlyList<TrayMenuItem> Items { get; init; } = [];

    /// <summary>True for a separator, which carries no text of its own.</summary>
    public bool IsSeparator { get; init; }

    /// <summary>False for the reference's status lines, which are there to be read and not chosen.</summary>
    public bool IsEnabled { get; init; } = true;

    /// <summary>What the line means, e.g. "open"; the shell turns it into a command of its own.</summary>
    public string Id { get; init; } = "";
}

/// <summary>
/// The app's icon in the operating system's notification area, without the app having to know the windowing toolkit.
///
/// The reference builds its tray icon with Windows Forms, which is Windows only; this is the same thing for every
/// platform, so the port can show a tray icon on Linux and macOS as well.
/// </summary>
public interface ITrayIcon : IDisposable
{
    /// <summary>Whether the icon is in the notification area at all.</summary>
    bool IsVisible { get; set; }

    /// <summary>The text shown while the pointer rests on the icon.</summary>
    string? Tooltip { get; set; }

    /// <summary>The picture of the icon, so the tray shows what the window shows.</summary>
    byte[]? Icon { get; set; }

    /// <summary>Tells the user something without opening the app, where the platform can show it.</summary>
    void ShowBalloon(string title, string message);

    /// <summary>Puts the menu in the tray icon, replacing the one that was there.</summary>
    void ShowContextMenu(IReadOnlyList<TrayMenuItem> items);

    /// <summary>The user clicked the icon.</summary>
    event EventHandler? Clicked;

    /// <summary>The platform is about to show the menu, so the app can fill it with what is true now.</summary>
    event EventHandler? ContextMenuRequested;

    /// <summary>The user chose a line of the menu.</summary>
    event EventHandler<TrayMenuItem>? MenuItemInvoked;
}
