namespace RustPlusDesk.Core.Dock;

/// <summary>
/// Reference CommandDockTile: one tile on the mini-map's command dock.
///
/// <see cref="Kind"/> is a string rather than an enum on purpose: a layout saved by a newer build can be read by an
/// older one, which drops the tiles it cannot render instead of failing to parse the file and wiping the whole dock.
/// </summary>
public sealed class CommandDockTile
{
    /// <summary>Reference CommandDockTile.Id: the tile's own identifier, a fresh one unless the file says otherwise.</summary>
    public string Id { get; set; } = Guid.NewGuid().ToString("N");

    /// <summary>Reference CommandDockTile.Kind: Clock | Device | TeamChat | ClanChat | Event | Rule, and the rest.</summary>
    public string Kind { get; set; } = CommandDockTileKinds.Clock;

    /// <summary>Reference CommandDockTile.Col: the column the tile sits in, counted from the map's corner.</summary>
    public int Col { get; set; }

    /// <summary>Reference CommandDockTile.Row: the row the tile sits in.</summary>
    public int Row { get; set; }

    /// <summary>Reference CommandDockTile.ColSpan: how many columns wide the tile is.</summary>
    public int ColSpan { get; set; } = 1;

    /// <summary>Reference CommandDockTile.RowSpan: how many rows high the tile is.</summary>
    public int RowSpan { get; set; } = 1;

    /// <summary>Reference CommandDockTile.EntityId: device tiles, the paired entity this tile controls or reports on.</summary>
    public uint EntityId { get; set; }

    /// <summary>
    /// Reference CommandDockTile.ServerKey: the server this tile belongs to, as "{host}-{port}", or null for one that
    /// belongs everywhere. Only device tiles get one: an entity id means nothing on another server.
    /// </summary>
    public string? ServerKey { get; set; }

    /// <summary>Reference CommandDockTile.RuleId: rule tiles, the Logic Engine rule this tile launches.</summary>
    public string? RuleId { get; set; }

    /// <summary>Reference CommandDockTile.EventKey: event tiles, cargo | deepsea | oilrig | heli | chinook | vendor.</summary>
    public string? EventKey { get; set; }

    /// <summary>
    /// Reference CommandDockTile.EventHideLabel: event tiles, leave off the name under the icon and show only the
    /// countdown. At one cell the name does not fit however the text is scaled, and the icon already says which event it is.
    /// </summary>
    public bool EventHideLabel { get; set; }

    /// <summary>Reference CommandDockTile.ClockStyle: clock tiles, 0 digital, 1 analog, 2 Rust style, 3 minimal.</summary>
    public int ClockStyle { get; set; }

    /// <summary>Reference CommandDockTile.ClockShowDayNight: clock tiles, append the time left until sunrise or sunset.</summary>
    public bool ClockShowDayNight { get; set; } = true;

    /// <summary>Reference CommandDockTile.Clock12Hour: clock tiles, 3:20 PM rather than 15:20.</summary>
    public bool Clock12Hour { get; set; }

    /// <summary>
    /// Reference CommandDockTile.ShowDeviceIcon: device tiles, show the icon chosen in the device list instead of the
    /// device's name. At one cell there is room for the icon and a state word, but not for a name as well.
    /// </summary>
    public bool ShowDeviceIcon { get; set; } = true;

    /// <summary>Reference CommandDockTile.ChatAbbreviateNames: chat tiles, shorten long names so the message still fits.</summary>
    public bool ChatAbbreviateNames { get; set; } = true;

    /// <summary>
    /// Reference CommandDockTile.DiscordChannel: discord tiles, which configured notification channel to post into -
    /// "chat", "events" or "shop". Null follows whichever is configured first.
    /// </summary>
    public string? DiscordChannel { get; set; }

    /// <summary>Reference CommandDockTile.DiscordMention: discord tiles, 0 none, 1 @here, 2 @everyone.</summary>
    public int DiscordMention { get; set; }

    /// <summary>Reference CommandDockTile.ShowGraph: server info tiles, draw the population line, not just the figure.</summary>
    public bool ShowGraph { get; set; } = true;

    /// <summary>
    /// Reference CommandDockTile.ServerInfoMinimal: server info tiles, a glyph and the figure in one cell. Not a smaller
    /// version of the full tile but a different one - no caption, no graph, no room for either.
    /// </summary>
    public bool ServerInfoMinimal { get; set; }

    /// <summary>
    /// Reference CommandDockTile.TranslateTarget: translate tiles, the language to translate into, as a culture name such
    /// as "ru-RU". Null means the language the app is running in.
    /// </summary>
    public string? TranslateTarget { get; set; }

    /// <summary>
    /// Reference CommandDockTile.TranslateUseAi: translate tiles, write down speech with the configured AI model rather
    /// than with Windows. Off by default, because it spends the user's own API credit.
    /// </summary>
    public bool TranslateUseAi { get; set; }

    /// <summary>
    /// Reference CommandDockTile.TranslateTextWithAi: translate tiles, translate with the configured AI model rather than
    /// with Google. Off by default, again because it spends the user's own credit.
    /// </summary>
    public bool TranslateTextWithAi { get; set; }

    /// <summary>
    /// Reference CommandDockTile.CollapseIncludesMap: collapse tiles, hide the map along with everything else. Off by
    /// default, because the usual reason to clear the dock is to see the game with the map still on it.
    /// </summary>
    public bool CollapseIncludesMap { get; set; }

    // ── Death tracking ──────────────────────────────────────────────────
    //
    // Where on the screen the killer's name is, as fractions of it rather than pixels. The band sits in the same place
    // whatever the resolution; a saved rectangle in pixels would be wrong the first time somebody changed it.

    /// <summary>Reference CommandDockTile.DeathRegionLeft: the left edge of the death band, as a fraction of the screen.</summary>
    public double DeathRegionLeft { get; set; } = 0.180;

    /// <summary>Reference CommandDockTile.DeathRegionTop: the top edge of the death band, as a fraction of the screen.</summary>
    public double DeathRegionTop { get; set; } = 0.020;

    /// <summary>Reference CommandDockTile.DeathRegionWidth: the width of the death band, as a fraction of the screen.</summary>
    public double DeathRegionWidth { get; set; } = 0.640;

    /// <summary>Reference CommandDockTile.DeathRegionHeight: the height of the death band, as a fraction of the screen.</summary>
    public double DeathRegionHeight { get; set; } = 0.034;

    /// <summary>Reference CommandDockTile.DeathTrackAlwaysVisible: death tracking tiles, stay on the dock while alive.</summary>
    public bool DeathTrackAlwaysVisible { get; set; }

    /// <summary>
    /// Reference CommandDockTile.DeathWipeKeepLatest: death wipe tiles, keep the newest marker per player instead of
    /// clearing every one. Off by default, so one press does the obvious thing.
    /// </summary>
    public bool DeathWipeKeepLatest { get; set; }

    /// <summary>
    /// Reference CommandDockTile.DeathWipeAlwaysVisible: death wipe tiles, stay on the dock even when the map has no
    /// markers to clear. A button with nothing to do is a button in the way.
    /// </summary>
    public bool DeathWipeAlwaysVisible { get; set; }

    /// <summary>
    /// Reference CommandDockTile.DeathReadWithAi: death tracking tiles, read the death screen with the AI model rather
    /// than with Windows. Off by default, because it spends the user's own API credit.
    /// </summary>
    public bool DeathReadWithAi { get; set; }

    // ── Appearance ──────────────────────────────────────────────────────
    //
    // All three are null until the tile's own settings are touched, and fall back to the dock's defaults. Two places that
    // set the same thing drift apart - one says 70% while the tile shows 40% and nobody can tell which wins - so the
    // global values are defaults rather than a second switch, and "back to global" is putting null back.

    /// <summary>Reference CommandDockTile.Opacity: background and border opacity, 0 to 1. Text is never faded.</summary>
    public double? Opacity { get; set; }

    /// <summary>Reference CommandDockTile.FontScale: multiplier on every font size in the tile.</summary>
    public double? FontScale { get; set; }

    /// <summary>
    /// Reference CommandDockTile.IconScale: device, event and rule tiles, how much bigger the icon is drawn than the
    /// text scale alone would make it. Null follows the text.
    /// </summary>
    public double? IconScale { get; set; }

    /// <summary>Reference CommandDockTile.TextColorKey: a key from <see cref="CommandDockTextColors"/>, or null for the theme's own.</summary>
    public string? TextColorKey { get; set; }
}

/// <summary>
/// Reference CommandDockTextColors: the text colours a tile can be set to.
///
/// Deliberately few, and each one picked to hold up on both a bright desert and a night-time screen - a tile turned fully
/// transparent sits directly on the game, and the theme's grey stops being readable the moment its background goes.
/// </summary>
public static class CommandDockTextColors
{
    /// <summary>Reference CommandDockTextColors.Auto: the theme's own colour.</summary>
    public const string Auto = "auto";

    /// <summary>Reference CommandDockTextColors.White: white text.</summary>
    public const string White = "white";

    /// <summary>Reference CommandDockTextColors.Black: black text.</summary>
    public const string Black = "black";

    /// <summary>Reference CommandDockTextColors.Cyan: cyan text.</summary>
    public const string Cyan = "cyan";

    /// <summary>Reference CommandDockTextColors.Amber: amber text.</summary>
    public const string Amber = "amber";

    /// <summary>Reference CommandDockTextColors.Red: red text.</summary>
    public const string Red = "red";

    /// <summary>Reference CommandDockTextColors.Green: green text.</summary>
    public const string Green = "green";

    /// <summary>Reference CommandDockTextColors.All: every colour a tile can be set to.</summary>
    public static readonly string[] All = { Auto, White, Black, Cyan, Amber, Red, Green };
}

/// <summary>
/// Reference CommandDockTileKinds: the kinds of tile the dock can hold.
///
/// A string rather than an enum, so a layout written by a newer build still loads in an older one: the tile it cannot
/// render is dropped, and the rest of the dock is left alone.
/// </summary>
public static class CommandDockTileKinds
{
    /// <summary>
    /// Reference CommandDockTileKinds.Map: the mini-map. One per dock, and the only tile whose size is free rather than
    /// derived from its cell span - but otherwise an ordinary occupant of the grid.
    /// </summary>
    public const string Map = "Map";

    /// <summary>
    /// Reference CommandDockTileKinds.MapTileId: the map tile's id is fixed rather than a fresh Guid. Its element is the
    /// one declared in XAML and survives every rebuild, so its mouse handlers are attached once and capture this id.
    /// </summary>
    public const string MapTileId = "map";

    /// <summary>Reference CommandDockTileKinds.Clock: a clock tile.</summary>
    public const string Clock = "Clock";

    /// <summary>Reference CommandDockTileKinds.Device: a paired device tile.</summary>
    public const string Device = "Device";

    /// <summary>Reference CommandDockTileKinds.TeamChat: the team chat tile.</summary>
    public const string TeamChat = "TeamChat";

    /// <summary>Reference CommandDockTileKinds.ClanChat: the clan chat tile.</summary>
    public const string ClanChat = "ClanChat";

    /// <summary>Reference CommandDockTileKinds.Event: a game event tile.</summary>
    public const string Event = "Event";

    /// <summary>Reference CommandDockTileKinds.Rule: a Logic Engine rule tile.</summary>
    public const string Rule = "Rule";

    /// <summary>Reference CommandDockTileKinds.Session: the session tile.</summary>
    public const string Session = "Session";

    /// <summary>Reference CommandDockTileKinds.Discord: the Discord notification tile.</summary>
    public const string Discord = "Discord";

    /// <summary>Reference CommandDockTileKinds.ServerInfo: the server information tile.</summary>
    public const string ServerInfo = "ServerInfo";

    /// <summary>
    /// Reference CommandDockTileKinds.AiCompanion: the AI companion. One per dock in practice - a second one would
    /// share the same recorder and the same key, so the two would only fight over one microphone.
    /// </summary>
    public const string AiCompanion = "AiCompanion";

    /// <summary>
    /// Reference CommandDockTileKinds.Translate: type or dictate something and read it back in another language. Shares
    /// the recorder with the AI companion, and for the same reason: there is one microphone.
    /// </summary>
    public const string Translate = "Translate";

    /// <summary>
    /// Reference CommandDockTileKinds.Collapse: hides every other tile, and shows them again. One button, one cell. The
    /// tiles are hidden rather than moved, so expanding puts the arrangement back exactly as it was.
    /// </summary>
    public const string Collapse = "Collapse";

    /// <summary>
    /// Reference CommandDockTileKinds.DeathWipe: one cell that clears the map's death markers. This one is about the
    /// ones that already did, not about the death that just happened.
    /// </summary>
    public const string DeathWipe = "DeathWipe";

    /// <summary>
    /// Reference CommandDockTileKinds.DeathTrack: after a death, one press to read the killer's name off the death
    /// screen. Only on the dock while the player is dead, because that is the only time the screen it reads is on screen.
    /// </summary>
    public const string DeathTrack = "DeathTrack";
}
