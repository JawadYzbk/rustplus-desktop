using System.Globalization;

namespace RustPlusDesk.Core.Chat;

/// <summary>
/// Reference ServerProfile chat-command settings: the enable switch, the prefix, the response delay and the
/// per-server command names (Cmd*). The names are validated the way the reference property setters validate
/// them, so what the settings card stores is what the engine matches.
/// </summary>
public sealed record ChatCommandSettings(
    bool Enabled = true,
    string Prefix = "!",
    double ResponseDelaySeconds = 0.5,
    string Pop = "pop",
    string List = "commands",
    string Time = "time",
    string Promote = "promote",
    string DeepSea = "deepsea",
    string Cargo = "cargo",
    string Afk = "afk",
    string OilRig = "oilrig",
    string Heli = "heli",
    string Vendor = "vendor",
    string UpkeepDetail = "upkeepdetail",
    string BaseCodes = "code",
    string Ai = "ai",
    string CustomTimer = "timer")
{
    public static ChatCommandSettings Defaults { get; } = new();

    /// <summary>Reference ChatCommandPrefix getter: an empty prefix reads as "!".</summary>
    public string EffectivePrefix => string.IsNullOrEmpty(Prefix) ? "!" : Prefix;

    /// <summary>Reference ChatCommandPrefix setter: only "!", ".", "," and "\" are accepted.</summary>
    public static bool IsValidPrefix(string? value) => value is "!" or "." or "," or "\\";

    /// <summary>
    /// Reference ServerProfile.ValidateCommand: trimmed, a leading "!" removed, and the fallback when the
    /// result is empty or starts with a digit (a command name may not begin with a number).
    /// </summary>
    public static string ValidateCommand(string? value, string fallback)
    {
        if (string.IsNullOrWhiteSpace(value)) return fallback;
        string trimmed = value.Trim().TrimStart('!');
        return trimmed.Length > 0 && char.IsDigit(trimmed[0]) ? fallback : trimmed;
    }
}

/// <summary>The reference Resources.resx chat-command strings, verbatim.</summary>
public static class ChatCommandTexts
{
    public const string AgoHoursMinutes = "{0}h {1}m";
    public const string AgoMinutes = "{0}m";
    public const string PopResponse = "{0} players currently online.{1}";
    public const string PopQueue = " Queue: {0} players.";
    public const string TimeResponse = "Current in-game time: {0}.";
    public const string PromoteResponse = "{0} was promoted to leader.";
    public const string WhoResponse = "Controlled by Rust+ Desktop (PC App Active)";
    public const string DeepSeaActive = "Deep Sea is active! Running for {0}.";
    public const string DeepSeaActiveMidEvent = "Deep Sea is currently active (connected mid-event – spawn time unknown).";
    public const string DeepSeaEndedMinutesAgo = "Deep Sea ended {0} minutes ago this session.";
    public const string DeepSeaStatusUnknown = "Deep Sea event status unknown (not seen this session).";
    public const string CargoNotActive = "Cargo Ship not active.";
    public const string CargoDockedDeparts = "Cargo Ship docked at {0}. Departs in approx. {1} minutes.";
    public const string CargoDockedPreparingDepart = "Cargo Ship docked at {0} and preparing to depart.";
    public const string CargoDockedUnknown = "Cargo Ship docked at {0} (departure time unknown).";
    public const string CargoActiveLeaves = "Cargo Ship active. Leaves in approx. {0} minutes.";
    public const string CargoActivePreparingLeave = "Cargo Ship active and preparing to leave soon.";
    public const string CargoActiveDurationNotLearned = "Cargo Ship active (route duration not yet learned for this server).";
    public const string CargoActiveMidRoute = "Cargo Ship active (connected after spawn – remaining time unknown).";
    public const string CargoDespawnedMinutesAgo = "Cargo Ship not active. Despawned {0} minutes ago this session.";
    public const string HeliActive = "Patrol Heli is active! Running for {0}.";
    public const string HeliActiveMidEvent = "Patrol Heli is currently active (connected mid-event – spawn time unknown).";
    public const string HeliNotActiveAgo = "Patrol Heli not active. It was {0} {1} ago this session.";
    public const string HeliReasonShotDown = "shot down";
    public const string HeliReasonLeftMap = "left the map";
    public const string HeliStatusUnknown = "Patrol Heli event status unknown (not seen this session).";
    public const string VendorActive = "Travelling Vendor is active! Running for {0}.";
    public const string VendorActiveMidEvent = "Travelling Vendor is currently active (connected mid-event – spawn time unknown).";
    public const string VendorDespawnedAgo = "Travelling Vendor not active. Despawned {0} ago this session.";
    public const string VendorStatusUnknown = "Travelling Vendor event status unknown (not seen this session).";
    public const string ListHeader = "COMMANDS: {0}";
    public const string NoOneAfk = "No one is AFK.";
    public const string AfkTeamScope = "Team";
    public const string HarborFallback = "harbor";
}

/// <summary>One teammate the AFK answer reports: the reference reads TeamMembers.Where(t =&gt; t.IsAfk).</summary>
public sealed record ChatAfkMember(string Name, DateTime LastMoveUtc);

/// <summary>
/// The raw event state the reference heli/vendor/deep-sea answers read (_heliSpawnTime, _dynStates,
/// _vendorDespawnTime, _deepSeaSpawnTime, ...). <see cref="HeliGrid"/> is the grid label of the heli's last
/// real position, which the reference derives from the marker.
/// </summary>
public sealed record ChatEventCommandState(
    bool DeepSeaActive = false,
    DateTime? DeepSeaSpawnTime = null,
    DateTime? DeepSeaDespawnTime = null,
    bool HeliActive = false,
    DateTime? HeliSpawnTime = null,
    string? HeliGrid = null,
    DateTime? HeliLastEventUtc = null,
    bool HeliLastEventWasCrash = false,
    bool VendorActive = false,
    DateTime? VendorSpawnTime = null,
    DateTime? VendorDespawnTime = null);

/// <summary>
/// The cargo state the reference answer reads (_cargoDockStates entry, _cargoLastDespawnUtc and the two
/// learned durations for the server).
/// </summary>
public sealed record ChatCargoCommandState(
    bool Active = false,
    bool Docked = false,
    DateTime? DockTime = null,
    bool WasAlreadyDocked = false,
    string? HarborName = null,
    bool SeenAtEdge = false,
    DateTime? FirstSeen = null,
    int LearnedDockingMinutes = 0,
    int LearnedFullLifeMinutes = 0,
    DateTime? LastDespawnUtc = null);

/// <summary>
/// Everything a command answers from. The reference reads these off MainWindow fields; the port passes the same
/// values in, so the answers do not depend on where the port keeps them.
///
/// <see cref="Connected"/> mirrors the reference's "if (_rust is not RustPlusClientReal real) return;": without a
/// live socket every game-backed command is ignored, and the AI command (a later slice) is the only one answered
/// before that check.
/// <see cref="CanTrackDeepSea"/> / <see cref="CanTrackCargo"/> mirror EventCapabilities.IsTrackable: on a server
/// whose event source cannot see the event, the command is ignored silently rather than answering "unknown".
/// <see cref="CanPromote"/> is false when the live session cannot hand out leadership: the reference answers the
/// author that they were promoted and fires the request at the game, so a port without the action must stay silent
/// rather than say something it did not do.
/// </summary>
public sealed record ChatCommandState(
    bool Connected = false,
    bool CanTrackDeepSea = true,
    bool CanTrackCargo = true,
    uint? Players = null,
    uint? Queue = null,
    string? ServerTime = null,
    string? TimeUntilNextPhase = null,
    IReadOnlyList<ChatAfkMember>? Afk = null,
    ChatEventCommandState? Events = null,
    ChatCargoCommandState? Cargo = null,
    bool CanPromote = false,
    bool HasBaseCodes = false,
    bool CanListAi = false,
    string? FirstCustomTimerCommand = null);

/// <summary>An incoming chat message as the command engine sees it (reference TeamChatMessage + channel).</summary>
public sealed record ChatCommandRequest(ChatChannel Channel, string Author, ulong SteamId, string Text);

/// <summary>
/// The reference ServerProfile chat-command fields in profiles.json, one profile at a time. Like the other legacy
/// stores, a write goes through the serialized, backed-up profile writer and adopts the latest document.
/// </summary>
public interface IChatCommandSettingsStore
{
    Task<ChatCommandSettings> GetAsync(string profilesPath, string profileKey, CancellationToken ct);
    /// <summary>Throws <c>ProfileChangedException</c> when the profile no longer exists.</summary>
    Task SaveAsync(string profilesPath, string profileKey, ChatCommandSettings settings, CancellationToken ct);
}

/// <summary>What a command produced: the reply to send (if any), whether the author asked to be promoted, and a log line.</summary>
public sealed record ChatCommandReply(string? Message = null, bool PromoteToLeader = false, string? Log = null)
{
    /// <summary>Nothing to answer: not a command, disabled, cooling down, or an event that cannot be tracked.</summary>
    public static readonly ChatCommandReply Ignored = new();
}

/// <summary>Reference _lastChatCommandTime / ChatCommandCooldownSeconds (2s, for system stability).</summary>
public sealed class ChatCommandCooldown(double seconds = ChatCommands.CooldownSeconds)
{
    private DateTime last = DateTime.MinValue;

    /// <summary>Consumes the window when it is open; false while a command is still cooling down.</summary>
    public bool TryEnter(DateTime now)
    {
        if ((now - last).TotalSeconds < seconds) return false;
        last = now;
        return true;
    }

    public void Reset() => last = DateTime.MinValue;
}

/// <summary>
/// Reference ProcessChatCommands: matches the message against the profile's command names and builds the reply.
/// The parts that need features this port has not ported yet (chat master election, clan role gating, switches,
/// upkeep, base codes, custom timers, logic rules, AI) stay with their slice; everything here is the ported rule
/// verbatim, including the exact answer texts.
/// </summary>
public static class ChatCommands
{
    public const int CooldownSeconds = 2;

    /// <summary>Reference FormatAgo: "1h 5m" above an hour, "45m" below.</summary>
    public static string Ago(TimeSpan span) => span.TotalHours >= 1
        ? string.Format(CultureInfo.InvariantCulture, ChatCommandTexts.AgoHoursMinutes, (int)span.TotalHours, span.Minutes)
        : string.Format(CultureInfo.InvariantCulture, ChatCommandTexts.AgoMinutes, (int)span.TotalMinutes);

    /// <summary>Reference 128-character cut: keep 125 and mark the cut with an ellipsis.</summary>
    public static string Truncate(string text) => text.Length > ChatLimits.MaxLength
        ? text[..(ChatLimits.MaxLength - 3)] + "..."
        : text;

    /// <summary>
    /// Reference ProcessChatCommands. The cooldown is consumed at the point the reference consumes it: after the
    /// prefix match and the chat-master check, before the socket check, so an ignored command on a dead socket
    /// still blocks the next one for two seconds exactly as the reference does.
    /// </summary>
    public static ChatCommandReply Process(ChatCommandRequest request, ChatCommandState state,
        ChatCommandSettings settings, ChatCommandCooldown cooldown, DateTime now)
    {
        if (!settings.Enabled) return ChatCommandReply.Ignored;

        string prefix = settings.EffectivePrefix;
        string cmd = request.Text.Trim().ToLowerInvariant();
        if (string.IsNullOrEmpty(cmd) || !cmd.StartsWith(prefix, StringComparison.Ordinal)) return ChatCommandReply.Ignored;

        if (!cooldown.TryEnter(now))
            return new ChatCommandReply(null, Log: $"[ChatCommand] Ignoring '{cmd}' from {request.Author} (Cooldown active)");

        cmd = cmd[prefix.Length..];

        if (!state.Connected) return ChatCommandReply.Ignored;

        string Log(string what) => $"[ChatCommand] {what} executed by {request.Author}";

        if (cmd == settings.List.ToLowerInvariant())
            return new ChatCommandReply(ListAnswer(settings, state), Log: Log("List"));

        if (cmd == settings.Pop.ToLowerInvariant())
        {
            string queue = state.Queue is > 0
                ? string.Format(CultureInfo.InvariantCulture, ChatCommandTexts.PopQueue, state.Queue.Value)
                : "";
            return new ChatCommandReply(
                string.Format(CultureInfo.InvariantCulture, ChatCommandTexts.PopResponse,
                    state.Players?.ToString(CultureInfo.InvariantCulture) ?? "-", queue), Log: Log("Pop"));
        }

        if (cmd == settings.Time.ToLowerInvariant())
        {
            string message = string.Format(CultureInfo.InvariantCulture, ChatCommandTexts.TimeResponse, state.ServerTime ?? "-");
            if (!string.IsNullOrWhiteSpace(state.TimeUntilNextPhase)) message += $" ({state.TimeUntilNextPhase})";
            return new ChatCommandReply(message.Trim(), Log: Log("Time"));
        }

        if (!string.IsNullOrWhiteSpace(settings.Afk) && cmd == settings.Afk.ToLowerInvariant())
            return new ChatCommandReply(AfkAnswer(state.Afk ?? [], request.Channel, now), Log: Log("AFK"));

        if (cmd == settings.Promote.ToLowerInvariant())
        {
            if (!state.CanPromote) return ChatCommandReply.Ignored;
            return new ChatCommandReply(
                string.Format(CultureInfo.InvariantCulture, ChatCommandTexts.PromoteResponse, request.Author),
                PromoteToLeader: true, Log: Log("Promote"));
        }

        if (cmd is "who" or "controller" or "bot")
            return new ChatCommandReply(ChatCommandTexts.WhoResponse, Log: Log("Who"));

        if (cmd == settings.DeepSea.ToLowerInvariant())
        {
            if (!state.CanTrackDeepSea) return ChatCommandReply.Ignored;
            return new ChatCommandReply(DeepSeaAnswer(state.Events ?? new ChatEventCommandState(), now), Log: Log("DeepSea"));
        }

        if (cmd == settings.Cargo.ToLowerInvariant())
        {
            if (!state.CanTrackCargo) return ChatCommandReply.Ignored;
            return new ChatCommandReply(CargoAnswer(state.Cargo ?? new ChatCargoCommandState(), now), Log: Log("Cargo"));
        }

        if (cmd == settings.Heli.ToLowerInvariant())
            return new ChatCommandReply(HeliAnswer(state.Events ?? new ChatEventCommandState(), now), Log: Log("Heli"));

        if (cmd == settings.Vendor.ToLowerInvariant())
            return new ChatCommandReply(VendorAnswer(state.Events ?? new ChatEventCommandState(), now), Log: Log("Vendor"));

        return ChatCommandReply.Ignored;
    }

    /// <summary>
    /// Reference list command. Advertises a command only when it can do something: the timer entry only once a
    /// timer exists, base codes only to a channel allowed to see them and only once a code is set, and the AI
    /// command only to an author allowed to use it.
    /// </summary>
    public static string ListAnswer(ChatCommandSettings settings, ChatCommandState state)
    {
        string prefix = settings.EffectivePrefix;
        var standard = new List<string>();
        if (!string.IsNullOrWhiteSpace(state.FirstCustomTimerCommand)) standard.Add(prefix + state.FirstCustomTimerCommand);
        Add(standard, prefix, settings.Pop);
        Add(standard, prefix, settings.Time);
        Add(standard, prefix, settings.Promote);
        Add(standard, prefix, settings.DeepSea);
        Add(standard, prefix, settings.Cargo);
        Add(standard, prefix, settings.OilRig);
        Add(standard, prefix, settings.Heli);
        Add(standard, prefix, settings.Vendor);
        Add(standard, prefix, settings.UpkeepDetail);
        Add(standard, prefix, settings.Afk);
        standard.Add(prefix + "who");
        if (!string.IsNullOrWhiteSpace(settings.BaseCodes) && state.HasBaseCodes) Add(standard, prefix, settings.BaseCodes);
        if (!string.IsNullOrWhiteSpace(settings.Ai) && state.CanListAi) Add(standard, prefix, settings.Ai);
        return Truncate(string.Format(CultureInfo.InvariantCulture, ChatCommandTexts.ListHeader, string.Join(", ", standard)));

        static void Add(List<string> commands, string prefix, string name)
        {
            if (!string.IsNullOrWhiteSpace(name)) commands.Add(prefix + name);
        }
    }

    /// <summary>Reference AFK answer: "Name - m:ss | ...", scoped when the question came from clan chat.</summary>
    public static string AfkAnswer(IReadOnlyList<ChatAfkMember> members, ChatChannel channel, DateTime now)
    {
        string scope = channel == ChatChannel.Clan ? ChatCommandTexts.AfkTeamScope : "";
        string prefix = string.IsNullOrEmpty(scope) ? "AFK: " : $"AFK ({scope}): ";
        if (members.Count == 0)
        {
            string none = ChatCommandTexts.NoOneAfk;
            return string.IsNullOrEmpty(scope) ? none : $"({scope}) {none}";
        }

        var parts = members.Select(member =>
        {
            int totalSeconds = (int)(now - member.LastMoveUtc).TotalSeconds;
            return $"{member.Name} - {totalSeconds / 60}:{totalSeconds % 60:D2}";
        });
        return prefix + string.Join(" | ", parts);
    }

    public static string DeepSeaAnswer(ChatEventCommandState events, DateTime now)
    {
        if (events.DeepSeaActive)
        {
            return events.DeepSeaSpawnTime.HasValue
                ? string.Format(CultureInfo.InvariantCulture, ChatCommandTexts.DeepSeaActive, Ago(now - events.DeepSeaSpawnTime.Value))
                : ChatCommandTexts.DeepSeaActiveMidEvent;
        }
        if (events.DeepSeaDespawnTime.HasValue)
            return string.Format(CultureInfo.InvariantCulture, ChatCommandTexts.DeepSeaEndedMinutesAgo,
                (int)(now - events.DeepSeaDespawnTime.Value).TotalMinutes);
        return ChatCommandTexts.DeepSeaStatusUnknown;
    }

    public static string CargoAnswer(ChatCargoCommandState cargo, DateTime now)
    {
        if (!cargo.Active)
        {
            return cargo.LastDespawnUtc.HasValue
                ? string.Format(CultureInfo.InvariantCulture, ChatCommandTexts.CargoDespawnedMinutesAgo,
                    (int)(now - cargo.LastDespawnUtc.Value).TotalMinutes)
                : ChatCommandTexts.CargoNotActive;
        }

        string harbor = cargo.HarborName is { Length: > 0 } ? cargo.HarborName : ChatCommandTexts.HarborFallback;
        if (cargo.Docked && cargo.DockTime.HasValue)
        {
            if (cargo.LearnedDockingMinutes > 0 && !cargo.WasAlreadyDocked)
            {
                var remain = TimeSpan.FromMinutes(cargo.LearnedDockingMinutes) - (now - cargo.DockTime.Value);
                return remain.TotalMinutes > 0
                    ? string.Format(CultureInfo.InvariantCulture, ChatCommandTexts.CargoDockedDeparts, harbor, (int)remain.TotalMinutes)
                    : string.Format(CultureInfo.InvariantCulture, ChatCommandTexts.CargoDockedPreparingDepart, harbor);
            }
            return string.Format(CultureInfo.InvariantCulture, ChatCommandTexts.CargoDockedUnknown, harbor);
        }

        if (cargo.SeenAtEdge)
        {
            if (cargo.LearnedFullLifeMinutes > 0 && cargo.FirstSeen.HasValue)
            {
                var remain = TimeSpan.FromMinutes(cargo.LearnedFullLifeMinutes) - (now - cargo.FirstSeen.Value);
                return remain.TotalMinutes > 0
                    ? string.Format(CultureInfo.InvariantCulture, ChatCommandTexts.CargoActiveLeaves, (int)remain.TotalMinutes)
                    : ChatCommandTexts.CargoActivePreparingLeave;
            }
            return ChatCommandTexts.CargoActiveDurationNotLearned;
        }

        return ChatCommandTexts.CargoActiveMidRoute;
    }

    public static string HeliAnswer(ChatEventCommandState events, DateTime now)
    {
        if (events.HeliActive)
        {
            string suffix = events.HeliGrid is { Length: > 0 } grid ? $" [{grid}]" : "";
            return (events.HeliSpawnTime.HasValue
                ? string.Format(CultureInfo.InvariantCulture, ChatCommandTexts.HeliActive, Ago(now - events.HeliSpawnTime.Value))
                : ChatCommandTexts.HeliActiveMidEvent) + suffix;
        }
        if (events.HeliLastEventUtc.HasValue)
        {
            string reason = events.HeliLastEventWasCrash ? ChatCommandTexts.HeliReasonShotDown : ChatCommandTexts.HeliReasonLeftMap;
            return string.Format(CultureInfo.InvariantCulture, ChatCommandTexts.HeliNotActiveAgo, reason,
                Ago(now - events.HeliLastEventUtc.Value));
        }
        return ChatCommandTexts.HeliStatusUnknown;
    }

    public static string VendorAnswer(ChatEventCommandState events, DateTime now)
    {
        if (events.VendorActive)
        {
            return events.VendorSpawnTime.HasValue
                ? string.Format(CultureInfo.InvariantCulture, ChatCommandTexts.VendorActive, Ago(now - events.VendorSpawnTime.Value))
                : ChatCommandTexts.VendorActiveMidEvent;
        }
        if (events.VendorDespawnTime.HasValue)
            return string.Format(CultureInfo.InvariantCulture, ChatCommandTexts.VendorDespawnedAgo,
                Ago(now - events.VendorDespawnTime.Value));
        return ChatCommandTexts.VendorStatusUnknown;
    }
}
