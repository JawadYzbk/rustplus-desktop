using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Core.Dock;

namespace RustPlusDesk.Avalonia.Features.Dock;

/// <summary>One of the colours a tile's text can be drawn in.</summary>
/// <param name="Key">The colour's key, as it is written to the arrangement.</param>
/// <param name="Name">The name it is shown under.</param>
public sealed record DockTextColorChoice(string Key, string Name);

/// <summary>
/// Reference MiniMapWindow.CommandDock.TileSettings: one tile's own settings.
///
/// The tile keeps its own appearance, and the dock's defaults are behind it: a tile that has never been given one of its
/// own follows the dock, and one that has follows itself. The settings are the tile's own fields, so they are saved
/// with the arrangement and are read back with it.
/// </summary>
public sealed partial class DockTileSettingsViewModel : ObservableObject
{
    private readonly CommandDockTile tile;
    private readonly DockViewModel dock;

    public DockTileSettingsViewModel(DockViewModel dock, CommandDockTile tile)
    {
        this.dock = dock;
        this.tile = tile;

        opacity = tile.Opacity ?? dock.Layout.DefaultOpacity;
        fontScale = tile.FontScale ?? dock.Layout.DefaultFontScale;
        iconScale = tile.IconScale ?? 1.0;
        textColorKey = tile.TextColorKey ?? CommandDockTextColors.Auto;
        colSpan = tile.ColSpan;
        rowSpan = tile.RowSpan;

        clockStyle = tile.ClockStyle;
        clockShowDayNight = tile.ClockShowDayNight;
        clock12Hour = tile.Clock12Hour;
        serverInfoMinimal = tile.ServerInfoMinimal;
        showGraph = tile.ShowGraph;
        showDeviceIcon = tile.ShowDeviceIcon;
        eventHideLabel = tile.EventHideLabel;
        chatAbbreviateNames = tile.ChatAbbreviateNames;
        deathWipeAlwaysVisible = tile.DeathWipeAlwaysVisible;
        deathWipeKeepLatest = tile.DeathWipeKeepLatest;
        collapseIncludesMap = tile.CollapseIncludesMap;
        deathTrackAlwaysVisible = tile.DeathTrackAlwaysVisible;
    }

    /// <summary>The tile these settings belong to.</summary>
    public CommandDockTile Tile => tile;

    /// <summary>The tile's name, so the window says whose settings it is showing.</summary>
    public string Title => tile.Kind switch
    {
        CommandDockTileKinds.Clock => "Clock",
        CommandDockTileKinds.ServerInfo => "Server information",
        CommandDockTileKinds.TeamChat => "Team chat",
        CommandDockTileKinds.ClanChat => "Clan chat",
        CommandDockTileKinds.Collapse => "Collapse",
        CommandDockTileKinds.Map => "Map",
        CommandDockTileKinds.Device => "Device",
        CommandDockTileKinds.Event => "Event",
        CommandDockTileKinds.Rule => "Rule",
        CommandDockTileKinds.Session => "Session",
        CommandDockTileKinds.DeathWipe => "Death markers",
        _ => tile.Kind,
    };

    /// <summary>How see-through the tile is, as a percentage.</summary>
    [ObservableProperty] private double opacity;

    /// <summary>How big the tile's text is, as a multiplier.</summary>
    [ObservableProperty] private double fontScale;

    /// <summary>How big the tile's icon is, as a multiplier.</summary>
    [ObservableProperty] private double iconScale;

    /// <summary>The colour the tile's text is drawn in.</summary>
    [ObservableProperty] private string textColorKey = CommandDockTextColors.Auto;

    /// <summary>How many cells wide the tile is.</summary>
    [ObservableProperty] private int colSpan = 1;

    /// <summary>How many cells high the tile is.</summary>
    [ObservableProperty] private int rowSpan = 1;

    /// <summary>Which clock the tile shows.</summary>
    [ObservableProperty] private int clockStyle;

    /// <summary>Whether the clock shows the day and night phase.</summary>
    [ObservableProperty] private bool clockShowDayNight = true;

    /// <summary>Whether the clock shows the twelve-hour time.</summary>
    [ObservableProperty] private bool clock12Hour;

    /// <summary>Whether the server information is the small one.</summary>
    [ObservableProperty] private bool serverInfoMinimal;

    /// <summary>Whether the server information draws its population line.</summary>
    [ObservableProperty] private bool showGraph = true;

    /// <summary>Whether the device tile shows the device's picture rather than its name.</summary>
    [ObservableProperty] private bool showDeviceIcon = true;

    /// <summary>Whether the event tile hides its label.</summary>
    [ObservableProperty] private bool eventHideLabel;

    /// <summary>Whether the chat tile shortens the names.</summary>
    [ObservableProperty] private bool chatAbbreviateNames = true;

    /// <summary>Whether the death-wipe tile stays on the dock when there is nothing to clear.</summary>
    [ObservableProperty] private bool deathWipeAlwaysVisible;

    /// <summary>Whether the death-wipe tile keeps the newest marker of each player.</summary>
    [ObservableProperty] private bool deathWipeKeepLatest;

    /// <summary>Whether the collapse button takes the map with it.</summary>
    [ObservableProperty] private bool collapseIncludesMap;

    /// <summary>Whether the death-tracking tile stays on the dock while the player is alive.</summary>
    [ObservableProperty] private bool deathTrackAlwaysVisible;

    /// <summary>The colours the tile's text can be drawn in, with the name each one is shown under.</summary>
    public IReadOnlyList<DockTextColorChoice> TextColors { get; } =
        CommandDockTextColors.All.Select(key => new DockTextColorChoice(key, ColorName(key))).ToArray();

    /// <summary>The name of a colour, for the picker.</summary>
    public static string ColorName(string key) => key switch
    {
        CommandDockTextColors.White => "White",
        CommandDockTextColors.Black => "Black",
        CommandDockTextColors.Cyan => "Cyan",
        CommandDockTextColors.Amber => "Amber",
        CommandDockTextColors.Red => "Red",
        CommandDockTextColors.Green => "Green",
        _ => "Theme colour",
    };

    /// <summary>The clock styles the tile can show.</summary>
    public IReadOnlyList<string> ClockStyles { get; } = ["Digital", "Analogue", "Rust", "Server time, small"];

    /// <summary>Whether this tile is a clock.</summary>
    public bool IsClock => tile.Kind == CommandDockTileKinds.Clock;

    /// <summary>Whether this tile is the server information.</summary>
    public bool IsServerInfo => tile.Kind == CommandDockTileKinds.ServerInfo;

    /// <summary>Whether this tile is a device.</summary>
    public bool IsDevice => tile.Kind == CommandDockTileKinds.Device;

    /// <summary>Whether this tile is an event.</summary>
    public bool IsEvent => tile.Kind == CommandDockTileKinds.Event;

    /// <summary>Whether this tile is a chat.</summary>
    public bool IsChat => tile.Kind is CommandDockTileKinds.TeamChat or CommandDockTileKinds.ClanChat;

    /// <summary>Whether this tile clears the death markers.</summary>
    public bool IsDeathWipe => tile.Kind == CommandDockTileKinds.DeathWipe;

    /// <summary>Whether this tile is the collapse button.</summary>
    public bool IsCollapse => tile.Kind == CommandDockTileKinds.Collapse;

    /// <summary>Whether this tile follows the player.</summary>
    public bool IsSession => tile.Kind == CommandDockTileKinds.Session;

    /// <summary>
    /// Reference the tile settings' save: the settings are the tile's own, and the arrangement is written with them.
    /// </summary>
    [RelayCommand]
    public void Save()
    {
        tile.Opacity = Opacity;
        tile.FontScale = FontScale;
        tile.IconScale = IconScale;
        tile.TextColorKey = TextColorKey;
        tile.ColSpan = Math.Max(1, ColSpan);
        tile.RowSpan = Math.Max(1, RowSpan);

        tile.ClockStyle = ClockStyle;
        tile.ClockShowDayNight = ClockShowDayNight;
        tile.Clock12Hour = Clock12Hour;
        tile.ServerInfoMinimal = ServerInfoMinimal;
        tile.ShowGraph = ShowGraph;
        tile.ShowDeviceIcon = ShowDeviceIcon;
        tile.EventHideLabel = EventHideLabel;
        tile.ChatAbbreviateNames = ChatAbbreviateNames;
        tile.DeathWipeAlwaysVisible = DeathWipeAlwaysVisible;
        tile.DeathWipeKeepLatest = DeathWipeKeepLatest;
        tile.CollapseIncludesMap = CollapseIncludesMap;
        tile.DeathTrackAlwaysVisible = DeathTrackAlwaysVisible;

        dock.Save();
        dock.RebuildTiles();
    }

    /// <summary>
    /// Reference the tile settings' reset: the tile goes back to the dock's own appearance.
    ///
    /// Only the appearance is reset: what the tile shows is the tile's own business.
    /// </summary>
    [RelayCommand]
    public void ResetAppearance()
    {
        tile.Opacity = null;
        tile.FontScale = null;
        tile.IconScale = null;
        tile.TextColorKey = null;

        Opacity = dock.Layout.DefaultOpacity;
        FontScale = dock.Layout.DefaultFontScale;
        IconScale = 1.0;
        TextColorKey = CommandDockTextColors.Auto;

        dock.Save();
        dock.RebuildTiles();
    }
}
