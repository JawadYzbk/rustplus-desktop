using RustPlusDesk.Avalonia.Shared.Icons;
using RustPlusDesk.Core.Dock;

namespace RustPlusDesk.Avalonia.Features.Dock;

/// <summary>
/// One thing the picker can put on the dock.
///
/// The kind is part of the entry rather than buried in the factory, because the factory is what decides whether the
/// entry can be offered at all: a kind this build cannot draw is not offered.
/// </summary>
/// <param name="Section">The heading the entry sits under.</param>
/// <param name="Title">The name shown in the picker.</param>
/// <param name="Hint">The line under the name that says what it does.</param>
/// <param name="Icon">The glyph beside the name.</param>
/// <param name="Kind">The tile kind the entry adds, and the gate on whether it is offered.</param>
/// <param name="Make">Builds the tile, with the settings the entry stands for.</param>
public sealed record DockTileOption(
    string Section,
    string Title,
    string Hint,
    FluentSymbol Icon,
    string Kind,
    Func<CommandDockTile> Make);

/// <summary>A heading in the picker and the entries under it.</summary>
public sealed record DockTileSection(string Title, FluentSymbol Icon, IReadOnlyList<DockTileOption> Options);

/// <summary>
/// The devices a picker can offer, when the connection has any.
/// </summary>
public interface IDockHostDevices
{
    /// <summary>The devices that are paired on this server, in the order the picker shows them.</summary>
    IReadOnlyList<DockHostDevice> DockDevices { get; }
}

/// <summary>One device, as the picker sees it: what it is called and which entity it is.</summary>
/// <param name="EntityId">The entity the tile controls.</param>
/// <param name="DisplayName">The name the user gave it.</param>
/// <param name="Kind">What the device is, for the icon and the line under the name.</param>
public sealed record DockHostDevice(uint EntityId, string DisplayName, string Kind);

/// <summary>
/// Reference CommandDockTilePicker: the catalogue of everything that can be added to the dock.
///
/// The whole catalogue is written down here, including the kinds this build cannot draw yet, and each entry is offered
/// only when <see cref="DockTiles.IsSupported"/> says this build draws it. That is the whole arrangement: the kinds that
/// are not drawn are skipped rather than shown as empty boxes, and a kind that gains a builder appears in the picker on
/// its own, with no second list to keep in step.
/// </summary>
public static class DockTileCatalogue
{
    /// <summary>
    /// The catalogue, as the picker shows it.
    /// </summary>
    /// <param name="mapPresent">Whether the dock already carries the map, which is the one tile that cannot be added twice.</param>
    /// <param name="host">The connection, for the kinds that are built from what the server actually has.</param>
    public static IReadOnlyList<DockTileSection> Sections(bool mapPresent, IDockHost? host = null)
    {
        var entries = new List<DockTileOption>
        {
            // ---- General: the widgets that are about the dock itself ----
            new("General", Loc("CommandDockCollapseTitle", "Collapse"),
                Loc("CommandDockCollapsePickerHint", "One cell that hides every other tile, and brings them back"),
                FluentSymbol.ArrowCollapseAll24, CommandDockTileKinds.Collapse,
                () => new CommandDockTile { Kind = CommandDockTileKinds.Collapse }),

            new("General", Loc("CommandDockDeathWipeTitle", "Wipe death markers"),
                Loc("CommandDockDeathWipePickerHint", "One cell that clears the map's death markers, and counts them"),
                FluentSymbol.Eraser24, CommandDockTileKinds.DeathWipe,
                () => new CommandDockTile { Kind = CommandDockTileKinds.DeathWipe }),

            // ---- Map ----
            // Offered only while no map is on the dock, which is the same condition the reference uses.
            new("Map", Loc("MiniMap", "Mini-map"),
                Loc("CommandDockAddMapHint", "Brings the map back with all layers on"),
                FluentSymbol.Map24, CommandDockTileKinds.Map,
                () => new CommandDockTile { Id = CommandDockTileKinds.MapTileId, Kind = CommandDockTileKinds.Map }),

            // ---- Clock ----
            new("Clock", Loc("CommandDockClockDigital", "Digital clock"),
                Loc("CommandDockClockDigitalHint", "Server time with the day and night phase"),
                FluentSymbol.Clock24, CommandDockTileKinds.Clock,
                () => new CommandDockTile { Kind = CommandDockTileKinds.Clock, ClockStyle = 0 }),

            new("Clock", Loc("CommandDockClockAnalog", "Analogue clock"),
                Loc("CommandDockClockAnalogHint", "A dial that follows the server clock"),
                FluentSymbol.Timer24, CommandDockTileKinds.Clock,
                () => new CommandDockTile { Kind = CommandDockTileKinds.Clock, ClockStyle = DockTiles.ClockStyleAnalogue }),

            new("Clock", Loc("CommandDockClockRust", "Rust clock"),
                Loc("CommandDockClockRustHint", "Server time in the game's own lettering"),
                FluentSymbol.Clock24, CommandDockTileKinds.Clock,
                () => new CommandDockTile { Kind = CommandDockTileKinds.Clock, ClockStyle = DockTiles.ClockStyleRust }),

            new("Clock", Loc("CommandDockClockMinimal", "Server time, small"),
                Loc("CommandDockClockMinimalHint", "One cell: the sun or the moon, and the time - as on the mini-map"),
                FluentSymbol.WeatherSunny24, CommandDockTileKinds.Clock,
                () => new CommandDockTile { Kind = CommandDockTileKinds.Clock, ClockStyle = DockTiles.ClockStyleMinimal }),

            // ---- Chat ----
            new("Chat", Loc("TeamChat", "Team chat"),
                Loc("CommandDockChatHint", "Two cells wide, resizable in edit mode"),
                FluentSymbol.PeopleCommunity24, CommandDockTileKinds.TeamChat,
                () => new CommandDockTile { Kind = CommandDockTileKinds.TeamChat, ColSpan = 2, RowSpan = 2 }),

            new("Chat", Loc("ClanChat", "Clan chat"),
                Loc("CommandDockChatHint", "Two cells wide, resizable in edit mode"),
                FluentSymbol.Shield24, CommandDockTileKinds.ClanChat,
                () => new CommandDockTile { Kind = CommandDockTileKinds.ClanChat, ColSpan = 2, RowSpan = 2 }),

            // ---- Server ----
            new("Server", Loc("CommandDockServerInfoTitle", "Server info"),
                Loc("CommandDockServerInfoHint", "Players on now, and how that has moved this session"),
                FluentSymbol.Server24, CommandDockTileKinds.ServerInfo,
                () => new CommandDockTile { Kind = CommandDockTileKinds.ServerInfo, ColSpan = 2, RowSpan = 1 }),

            new("Server", Loc("CommandDockServerMinimal", "Players, small"),
                Loc("CommandDockServerMinimalHint", "One cell: how many are on, and how many fit - as on the mini-map"),
                FluentSymbol.People24, CommandDockTileKinds.ServerInfo,
                () => new CommandDockTile
                {
                    Kind = CommandDockTileKinds.ServerInfo,
                    ServerInfoMinimal = true,
                    ShowGraph = false,
                }),

            // ---- Session ----
            new("Session", Loc("CommandDockSessionTitle", "Session stats"),
                Loc("CommandDockSessionHint", "Online time, distance, idle time and deaths"),
                FluentSymbol.DataUsage24, CommandDockTileKinds.Session,
                () => new CommandDockTile { Kind = CommandDockTileKinds.Session, ColSpan = 3, RowSpan = 1 }),

            // ---- AI companion ----
            new("AI companion", Loc("CommandDockAiTitle", "Ask AI"),
                Loc("CommandDockAiPickerHint", "Record a question, attach the screen, get an answer"),
                FluentSymbol.Sparkle24, CommandDockTileKinds.AiCompanion,
                () => new CommandDockTile { Kind = CommandDockTileKinds.AiCompanion, ColSpan = 2, RowSpan = 1 }),

            new("AI companion", Loc("CommandDockTranslateTitle", "Translate"),
                Loc("CommandDockTranslatePickerHint", "Paste a line, read it back in your language"),
                FluentSymbol.Translate24, CommandDockTileKinds.Translate,
                () => new CommandDockTile { Kind = CommandDockTileKinds.Translate, ColSpan = 3, RowSpan = 1 }),

            new("AI companion", Loc("CommandDockDeathTrackTitle", "Who killed you?"),
                Loc("CommandDockDeathTrackPickerHint", "Appears when you die: one press reads the name off the death screen"),
                FluentSymbol.HeartBroken24, CommandDockTileKinds.DeathTrack,
                () => new CommandDockTile { Kind = CommandDockTileKinds.DeathTrack, ColSpan = 2, RowSpan = 2 }),

            // ---- Events ----
            new("Events", "Oil Rig Smart Alarm", "",
                FluentSymbol.Box24, CommandDockTileKinds.Event,
                () => new CommandDockTile { Kind = CommandDockTileKinds.Event, EventKey = "oilrig" }),
        };

        // The devices of the server that is connected, and only those: a device that was never paired cannot be
        // put on the dock, and an entity id means nothing on another server.
        if (host is IDockHostDevices { DockDevices.Count: > 0 } devices)
        {
            foreach (var device in devices.DockDevices.OrderBy(device => device.DisplayName, StringComparer.CurrentCulture))
            {
                uint entityId = device.EntityId;
                var icon = device.Kind switch
                {
                    "SmartSwitch" or "Smart Switch" => FluentSymbol.Power24,
                    "StorageMonitor" => FluentSymbol.Database24,
                    "SmartAlarm" => FluentSymbol.Alert24,
                    _ => FluentSymbol.PlugConnected24,
                };

                entries.Add(new DockTileOption("Devices", device.DisplayName, device.Kind, icon, CommandDockTileKinds.Device,
                    () => new CommandDockTile { Kind = CommandDockTileKinds.Device, EntityId = entityId }));
            }
        }

        // A kind this build cannot draw is skipped rather than shown as an empty box, and the map is only offered
        // while the dock does not already carry it.
        var offered = entries
            .Where(entry => DockTiles.IsSupported(entry.Kind))
            .Where(entry => entry.Kind != CommandDockTileKinds.Map || !mapPresent)
            .ToList();

        // The sections keep the order the reference shows them in, and an empty one is left out entirely.
        var sections = new List<DockTileSection>();
        foreach (var group in offered.GroupBy(entry => entry.Section))
        {
            var section = new DockTileSection(group.Key, SectionIcon(group.Key), group.ToList());
            if (section.Options.Count > 0) sections.Add(section);
        }

        return sections;
    }

    /// <summary>The glyph beside a section's heading.</summary>
    private static FluentSymbol SectionIcon(string section) => section switch
    {
        "General" => FluentSymbol.Grid24,
        "Map" => FluentSymbol.Map24,
        "Clock" => FluentSymbol.Clock24,
        "Chat" => FluentSymbol.Chat24,
        "Server" => FluentSymbol.Server24,
        "Session" => FluentSymbol.DataUsage24,
        "AI companion" => FluentSymbol.Sparkle24,
        "Events" => FluentSymbol.AlertUrgent24,
        "Devices" => FluentSymbol.PlugConnected24,
        "Rules" => FluentSymbol.BrainCircuit24,
        _ => FluentSymbol.Grid24,
    };

    /// <summary>The text of a picker entry, in the app's language when it is translated.</summary>
    private static string Loc(string key, string fallback) => RustPlusDesk.Core.Localization.Loc.Text(key, fallback);
}
