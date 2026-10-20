using System.Globalization;
using RustPlusDesk.Core.Devices;
using RustPlusDesk.Core.Localization;
using RustPlusDesk.Core.Map;

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
    /// <summary>Reference ServerProfile.MaxBaseCodes: a profile holds at most five door codes.</summary>
    public const int MaxBaseCodes = 5;

    /// <summary>Shared empty mapping list, so two settings records with no mappings compare equal.</summary>
    public static IReadOnlyList<ChatCommandMapping> NoMappings { get; } = [];

    /// <summary>Shared empty base-code list, so two settings records with no codes compare equal.</summary>
    public static IReadOnlyList<ChatBaseCode> NoBaseCodes { get; } = [];

    /// <summary>Shared empty role-permission list, so two settings records with no roles compare equal.</summary>
    public static IReadOnlyList<ChatRolePermission> NoRolePermissions { get; } = [];

    public static ChatCommandSettings Defaults { get; } = new();

    /// <summary>Reference ServerProfile.SwitchCommandMappings: the switch names the chat can toggle.</summary>
    public IReadOnlyList<ChatCommandMapping> SwitchMappings { get; init; } = NoMappings;

    /// <summary>Reference ServerProfile.UpkeepCommandMappings: the storage monitors the upkeep commands report.</summary>
    public IReadOnlyList<ChatCommandMapping> UpkeepMappings { get; init; } = NoMappings;

    /// <summary>Reference ServerProfile.BaseCodes: the labelled door codes the base-code command answers with.</summary>
    public IReadOnlyList<ChatBaseCode> BaseCodeRows { get; init; } = NoBaseCodes;

    /// <summary>Reference ServerProfile.ClanCommandsAllowBaseCodes: clan chat may read the base codes.</summary>
    public bool ClanCommandsAllowBaseCodes { get; init; }

    /// <summary>Reference ServerProfile.ClanCommandsAllowSwitches: clan chat may toggle mapped switches.</summary>
    public bool ClanCommandsAllowSwitches { get; init; }

    /// <summary>
    /// Reference ServerProfile.ClanChatCommandsEnabled: commands are also answered when they arrive in clan chat.
    /// Team chat is always served once commands are on; the clan is an addition to it, never a replacement.
    /// </summary>
    public bool ClanCommandsEnabled { get; init; }

    /// <summary>Reference ServerProfile.ClanCommandsAllowAi: clan chat may use the AI command.</summary>
    public bool ClanCommandsAllowAi { get; init; }

    /// <summary>
    /// Reference ServerProfile.ClanRolePermissions: one row per clan role, carrying whether that role may use
    /// commands. Persisted so the choice survives a restart, and reconciled against the live clan on every pull.
    /// </summary>
    public IReadOnlyList<ChatRolePermission> ClanRolePermissions { get; init; } = NoRolePermissions;

    /// <summary>Whether any role at all may use commands (reference ServerProfile.HasAllowedClanRole).</summary>
    public bool HasAllowedClanRole => ClanRolePermissions.Any(permission => permission.Allowed);

    /// <summary>
    /// Reference ServerProfile.IsClanRoleAllowed: an unknown role is refused rather than allowed, so a role the clan
    /// gained since the last pull waits for the owner to grant it instead of inheriting permission nobody gave.
    /// </summary>
    public bool IsClanRoleAllowed(int roleId) => ClanRolePermissions.Any(permission => permission.RoleId == roleId && permission.Allowed);

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
/// <summary>
/// The reference Resources.resx strings the chat answers use, and the English wording they fall back to.
///
/// The answers are read from the app's resource catalogue in the language the user chose, exactly as the
/// reference does; the fallback is the wording the reference has, so a missing translation still answers in
/// English and the drift check still has something to compare against.
/// </summary>
public static class ChatCommandTexts
{
    /// <summary>The resource key for each answer, when the reference has one.</summary>
    public static IReadOnlyDictionary<string, string> Keys { get; } = new Dictionary<string, string>
    {
        ["AgoHoursMinutes"] = "AgoHoursMinutes",
        ["AgoMinutes"] = "AgoMinutes",
        ["PopResponse"] = "ChatCmdPopResponse",
        ["PopQueue"] = "ChatCmdPopQueue",
        ["TimeResponse"] = "ChatCmdTimeResponse",
        ["PromoteResponse"] = "ChatCmdPromoteResponse",
        ["DeepSeaActive"] = "ChatCmdDeepSeaActive",
        ["DeepSeaActiveMidEvent"] = "ChatCmdDeepSeaActiveMidEvent",
        ["DeepSeaEndedMinutesAgo"] = "ChatCmdDeepSeaEndedMinutesAgo",
        ["DeepSeaStatusUnknown"] = "ChatCmdDeepSeaStatusUnknown",
        ["CargoNotActive"] = "ChatCmdCargoNotActive",
        ["CargoDockedDeparts"] = "ChatCmdCargoDockedDeparts",
        ["CargoDockedPreparingDepart"] = "ChatCmdCargoDockedPreparingDepart",
        ["CargoDockedUnknown"] = "ChatCmdCargoDockedUnknown",
        ["CargoActiveLeaves"] = "ChatCmdCargoActiveLeaves",
        ["CargoActivePreparingLeave"] = "ChatCmdCargoActivePreparingLeave",
        ["CargoActiveDurationNotLearned"] = "ChatCmdCargoActiveDurationNotLearned",
        ["CargoActiveMidRoute"] = "ChatCmdCargoActiveMidRoute",
        ["CargoDespawnedMinutesAgo"] = "ChatCmdCargoDespawnedMinutesAgo",
        ["HeliActive"] = "ChatCmdHeliActive",
        ["HeliActiveMidEvent"] = "ChatCmdHeliActiveMidEvent",
        ["HeliNotActiveAgo"] = "ChatCmdHeliNotActiveAgo",
        ["HeliReasonShotDown"] = "ChatCmdHeliReasonShotDown",
        ["HeliReasonLeftMap"] = "ChatCmdHeliReasonLeftMap",
        ["HeliStatusUnknown"] = "ChatCmdHeliStatusUnknown",
        ["VendorActive"] = "ChatCmdVendorActive",
        ["VendorActiveMidEvent"] = "ChatCmdVendorActiveMidEvent",
        ["VendorDespawnedAgo"] = "ChatCmdVendorDespawnedAgo",
        ["VendorStatusUnknown"] = "ChatCmdVendorStatusUnknown",
        ["ListHeader"] = "ChatCmdListHeader",
        ["NoOneAfk"] = "ChatCmdNoOneAfk",
        ["AfkTeamScope"] = "ChatCmdAfkTeamScope",
        ["HarborFallback"] = "HarborFallback",
        ["SwitchToggled"] = "ChatCmdSwitchToggled",
        ["SwitchStateOn"] = "ChatCmdSwitchStateOn",
        ["SwitchStateOff"] = "ChatCmdSwitchStateOff",
        ["SwitchNotPaired"] = "ChatCmdSwitchNotPaired",
        ["UpkeepNoTcMapped"] = "ChatCmdUpkeepNoTcMapped",
        ["UpkeepTcEmptyExpired"] = "ChatCmdUpkeepTcEmptyExpired",
        ["UpkeepNeed24h"] = "ChatCmdUpkeepNeed24h",
        ["UpkeepTcTime"] = "ChatCmdUpkeepTcTime",
        ["UpkeepDays"] = "ChatCmdUpkeepDays",
        ["UpkeepHours"] = "ChatCmdUpkeepHours",
        ["UpkeepMinutes"] = "ChatCmdUpkeepMinutes",
        ["UpkeepEmptyExpiredShort"] = "ChatCmdUpkeepEmptyExpiredShort",
        ["UpkeepTimeShort"] = "ChatCmdUpkeepTimeShort",
        ["UpkeepHeader"] = "ChatCmdUpkeepHeader",
        ["UpkeepNotPaired"] = "ChatCmdUpkeepNotPaired",
        ["UpkeepNotPairedSingle"] = "ChatCmdUpkeepNotPairedSingle",
        ["MaterialWood"] = "MaterialWood",
        ["MaterialStone"] = "MaterialStone",
        ["MaterialMetal"] = "MaterialMetal",
        ["MaterialHQM"] = "MaterialHQM",
        ["OilRigHackRunning"] = "ChatCmdOilRigHackRunning",
        ["OilRigSmallLabel"] = "UiBadgeSmallOil",
        ["OilRigLargeLabel"] = "UiBadgeLargeOil",
        ["BaseCodeNameDefault"] = "BaseCodesNameDefault",
        ["NoBaseCodes"] = "ChatCmdNoBaseCodes",
        ["TimerMaxReached"] = "ChatCmdTimerMaxReached",
        ["TimerNameMustStartWithLetter"] = "TimerNameMustStartWithLetter",
        ["TimerCreated"] = "TimerCreated",
        ["TimerDurationRequired"] = "TimerDurationRequired",
        ["CountdownOneMinute"] = "CountdownOneMinute",
        ["TimerExpiredTitle"] = "TimerExpiredTitle",
    };

    /// <summary>The English wording of each answer, as the reference has it.</summary>
    public static IReadOnlyDictionary<string, string> Fallbacks { get; } = new Dictionary<string, string>
    {
        ["AgoHoursMinutes"] = "{0}h {1}m",
        ["AgoMinutes"] = "{0}m",
        ["PopResponse"] = "{0} players currently online.{1}",
        ["PopQueue"] = " Queue: {0} players.",
        ["TimeResponse"] = "Current in-game time: {0}.",
        ["PromoteResponse"] = "{0} was promoted to leader.",
        ["WhoResponse"] = "Controlled by Rust+ Desktop (PC App Active)",
        ["DeepSeaActive"] = "Deep Sea is active! Running for {0}.",
        ["DeepSeaActiveMidEvent"] = "Deep Sea is currently active (connected mid-event – spawn time unknown).",
        ["DeepSeaEndedMinutesAgo"] = "Deep Sea ended {0} minutes ago this session.",
        ["DeepSeaStatusUnknown"] = "Deep Sea event status unknown (not seen this session).",
        ["CargoNotActive"] = "Cargo Ship not active.",
        ["CargoDockedDeparts"] = "Cargo Ship docked at {0}. Departs in approx. {1} minutes.",
        ["CargoDockedPreparingDepart"] = "Cargo Ship docked at {0} and preparing to depart.",
        ["CargoDockedUnknown"] = "Cargo Ship docked at {0} (departure time unknown).",
        ["CargoActiveLeaves"] = "Cargo Ship active. Leaves in approx. {0} minutes.",
        ["CargoActivePreparingLeave"] = "Cargo Ship active and preparing to leave soon.",
        ["CargoActiveDurationNotLearned"] = "Cargo Ship active (route duration not yet learned for this server).",
        ["CargoActiveMidRoute"] = "Cargo Ship active (connected after spawn – remaining time unknown).",
        ["CargoDespawnedMinutesAgo"] = "Cargo Ship not active. Despawned {0} minutes ago this session.",
        ["HeliActive"] = "Patrol Heli is active! Running for {0}.",
        ["HeliActiveMidEvent"] = "Patrol Heli is currently active (connected mid-event – spawn time unknown).",
        ["HeliNotActiveAgo"] = "Patrol Heli not active. It was {0} {1} ago this session.",
        ["HeliReasonShotDown"] = "shot down",
        ["HeliReasonLeftMap"] = "left the map",
        ["HeliStatusUnknown"] = "Patrol Heli event status unknown (not seen this session).",
        ["VendorActive"] = "Travelling Vendor is active! Running for {0}.",
        ["VendorActiveMidEvent"] = "Travelling Vendor is currently active (connected mid-event – spawn time unknown).",
        ["VendorDespawnedAgo"] = "Travelling Vendor not active. Despawned {0} ago this session.",
        ["VendorStatusUnknown"] = "Travelling Vendor event status unknown (not seen this session).",
        ["ListHeader"] = "COMMANDS: {0}",
        ["NoOneAfk"] = "No one is AFK.",
        ["AfkTeamScope"] = "Team",
        ["HarborFallback"] = "harbor",
        ["SwitchToggled"] = "{0} turned {1}.",
        ["SwitchStateOn"] = "ON",
        ["SwitchStateOff"] = "OFF",
        ["SwitchNotPaired"] = "Bound Smart Switch not found or not paired.",
        ["UpkeepNoTcMapped"] = "No Tool Cupboards are mapped for upkeep commands.",
        ["UpkeepTcEmptyExpired"] = "Upkeep in {0} TC: Empty or expired.",
        ["UpkeepNeed24h"] = " Need/24h: {0}.",
        ["UpkeepTcTime"] = "Upkeep in {0} TC: {1}.",
        ["UpkeepDays"] = "{0} days",
        ["UpkeepHours"] = "{0} hours",
        ["UpkeepMinutes"] = "{0} minutes",
        ["UpkeepEmptyExpiredShort"] = "{0}: Empty/expired",
        ["UpkeepTimeShort"] = "{0}: {1}d, {2}h",
        ["UpkeepHeader"] = "Upkeep: {0}",
        ["UpkeepNotPaired"] = "Bound Tool Cupboard monitors not found or not paired.",
        ["UpkeepNotPairedSingle"] = "Bound Tool Cupboard monitor not found or not paired.",
        ["MaterialWood"] = "Wood",
        ["MaterialStone"] = "Stone",
        ["MaterialMetal"] = "Metal",
        ["MaterialHQM"] = "HQM",
        ["OilRigHackRunning"] = "{0}: crate unlocks in {1}m {2}s",
        ["OilRigNone"] = "{0}: None",
        ["OilRigNoAlarm"] = "No Oil Rig Smart Alarm notification received.",
        ["OilRigSmallLabel"] = "Small Oil",
        ["OilRigLargeLabel"] = "Large Oil",
        ["OilRigGenericLabel"] = "Oil Rig",
        ["BaseCodeNameDefault"] = "Doorcode",
        ["NoBaseCodes"] = "No base codes are set.",
        ["NoActiveTimers"] = "No active timers.",
        ["TimerMaxReached"] = "Maximum of 5 timers allowed.",
        ["TimerNameMustStartWithLetter"] = "Timer name must start with a letter.",
        ["TimerCreated"] = "Timer {0} created for {1}h {2}m {3}s",
        ["TimerDurationRequired"] = "Please enter a timer duration (hours, minutes, or seconds).",
        ["CountdownOneMinute"] = "1-minute countdown...",
        ["TimerExpiredTitle"] = "Timer Expired",
    };

    public static string AgoHoursMinutes => Keys.TryGetValue(nameof(AgoHoursMinutes), out string? key) ? Loc.Text(key, Fallbacks[nameof(AgoHoursMinutes)]) : Fallbacks[nameof(AgoHoursMinutes)];
    public static string AgoMinutes => Keys.TryGetValue(nameof(AgoMinutes), out string? key) ? Loc.Text(key, Fallbacks[nameof(AgoMinutes)]) : Fallbacks[nameof(AgoMinutes)];
    public static string PopResponse => Keys.TryGetValue(nameof(PopResponse), out string? key) ? Loc.Text(key, Fallbacks[nameof(PopResponse)]) : Fallbacks[nameof(PopResponse)];
    public static string PopQueue => Keys.TryGetValue(nameof(PopQueue), out string? key) ? Loc.Text(key, Fallbacks[nameof(PopQueue)]) : Fallbacks[nameof(PopQueue)];
    public static string TimeResponse => Keys.TryGetValue(nameof(TimeResponse), out string? key) ? Loc.Text(key, Fallbacks[nameof(TimeResponse)]) : Fallbacks[nameof(TimeResponse)];
    public static string PromoteResponse => Keys.TryGetValue(nameof(PromoteResponse), out string? key) ? Loc.Text(key, Fallbacks[nameof(PromoteResponse)]) : Fallbacks[nameof(PromoteResponse)];
    public static string WhoResponse => Keys.TryGetValue(nameof(WhoResponse), out string? key) ? Loc.Text(key, Fallbacks[nameof(WhoResponse)]) : Fallbacks[nameof(WhoResponse)];
    public static string DeepSeaActive => Keys.TryGetValue(nameof(DeepSeaActive), out string? key) ? Loc.Text(key, Fallbacks[nameof(DeepSeaActive)]) : Fallbacks[nameof(DeepSeaActive)];
    public static string DeepSeaActiveMidEvent => Keys.TryGetValue(nameof(DeepSeaActiveMidEvent), out string? key) ? Loc.Text(key, Fallbacks[nameof(DeepSeaActiveMidEvent)]) : Fallbacks[nameof(DeepSeaActiveMidEvent)];
    public static string DeepSeaEndedMinutesAgo => Keys.TryGetValue(nameof(DeepSeaEndedMinutesAgo), out string? key) ? Loc.Text(key, Fallbacks[nameof(DeepSeaEndedMinutesAgo)]) : Fallbacks[nameof(DeepSeaEndedMinutesAgo)];
    public static string DeepSeaStatusUnknown => Keys.TryGetValue(nameof(DeepSeaStatusUnknown), out string? key) ? Loc.Text(key, Fallbacks[nameof(DeepSeaStatusUnknown)]) : Fallbacks[nameof(DeepSeaStatusUnknown)];
    public static string CargoNotActive => Keys.TryGetValue(nameof(CargoNotActive), out string? key) ? Loc.Text(key, Fallbacks[nameof(CargoNotActive)]) : Fallbacks[nameof(CargoNotActive)];
    public static string CargoDockedDeparts => Keys.TryGetValue(nameof(CargoDockedDeparts), out string? key) ? Loc.Text(key, Fallbacks[nameof(CargoDockedDeparts)]) : Fallbacks[nameof(CargoDockedDeparts)];
    public static string CargoDockedPreparingDepart => Keys.TryGetValue(nameof(CargoDockedPreparingDepart), out string? key) ? Loc.Text(key, Fallbacks[nameof(CargoDockedPreparingDepart)]) : Fallbacks[nameof(CargoDockedPreparingDepart)];
    public static string CargoDockedUnknown => Keys.TryGetValue(nameof(CargoDockedUnknown), out string? key) ? Loc.Text(key, Fallbacks[nameof(CargoDockedUnknown)]) : Fallbacks[nameof(CargoDockedUnknown)];
    public static string CargoActiveLeaves => Keys.TryGetValue(nameof(CargoActiveLeaves), out string? key) ? Loc.Text(key, Fallbacks[nameof(CargoActiveLeaves)]) : Fallbacks[nameof(CargoActiveLeaves)];
    public static string CargoActivePreparingLeave => Keys.TryGetValue(nameof(CargoActivePreparingLeave), out string? key) ? Loc.Text(key, Fallbacks[nameof(CargoActivePreparingLeave)]) : Fallbacks[nameof(CargoActivePreparingLeave)];
    public static string CargoActiveDurationNotLearned => Keys.TryGetValue(nameof(CargoActiveDurationNotLearned), out string? key) ? Loc.Text(key, Fallbacks[nameof(CargoActiveDurationNotLearned)]) : Fallbacks[nameof(CargoActiveDurationNotLearned)];
    public static string CargoActiveMidRoute => Keys.TryGetValue(nameof(CargoActiveMidRoute), out string? key) ? Loc.Text(key, Fallbacks[nameof(CargoActiveMidRoute)]) : Fallbacks[nameof(CargoActiveMidRoute)];
    public static string CargoDespawnedMinutesAgo => Keys.TryGetValue(nameof(CargoDespawnedMinutesAgo), out string? key) ? Loc.Text(key, Fallbacks[nameof(CargoDespawnedMinutesAgo)]) : Fallbacks[nameof(CargoDespawnedMinutesAgo)];
    public static string HeliActive => Keys.TryGetValue(nameof(HeliActive), out string? key) ? Loc.Text(key, Fallbacks[nameof(HeliActive)]) : Fallbacks[nameof(HeliActive)];
    public static string HeliActiveMidEvent => Keys.TryGetValue(nameof(HeliActiveMidEvent), out string? key) ? Loc.Text(key, Fallbacks[nameof(HeliActiveMidEvent)]) : Fallbacks[nameof(HeliActiveMidEvent)];
    public static string HeliNotActiveAgo => Keys.TryGetValue(nameof(HeliNotActiveAgo), out string? key) ? Loc.Text(key, Fallbacks[nameof(HeliNotActiveAgo)]) : Fallbacks[nameof(HeliNotActiveAgo)];
    public static string HeliReasonShotDown => Keys.TryGetValue(nameof(HeliReasonShotDown), out string? key) ? Loc.Text(key, Fallbacks[nameof(HeliReasonShotDown)]) : Fallbacks[nameof(HeliReasonShotDown)];
    public static string HeliReasonLeftMap => Keys.TryGetValue(nameof(HeliReasonLeftMap), out string? key) ? Loc.Text(key, Fallbacks[nameof(HeliReasonLeftMap)]) : Fallbacks[nameof(HeliReasonLeftMap)];
    public static string HeliStatusUnknown => Keys.TryGetValue(nameof(HeliStatusUnknown), out string? key) ? Loc.Text(key, Fallbacks[nameof(HeliStatusUnknown)]) : Fallbacks[nameof(HeliStatusUnknown)];
    public static string VendorActive => Keys.TryGetValue(nameof(VendorActive), out string? key) ? Loc.Text(key, Fallbacks[nameof(VendorActive)]) : Fallbacks[nameof(VendorActive)];
    public static string VendorActiveMidEvent => Keys.TryGetValue(nameof(VendorActiveMidEvent), out string? key) ? Loc.Text(key, Fallbacks[nameof(VendorActiveMidEvent)]) : Fallbacks[nameof(VendorActiveMidEvent)];
    public static string VendorDespawnedAgo => Keys.TryGetValue(nameof(VendorDespawnedAgo), out string? key) ? Loc.Text(key, Fallbacks[nameof(VendorDespawnedAgo)]) : Fallbacks[nameof(VendorDespawnedAgo)];
    public static string VendorStatusUnknown => Keys.TryGetValue(nameof(VendorStatusUnknown), out string? key) ? Loc.Text(key, Fallbacks[nameof(VendorStatusUnknown)]) : Fallbacks[nameof(VendorStatusUnknown)];
    public static string ListHeader => Keys.TryGetValue(nameof(ListHeader), out string? key) ? Loc.Text(key, Fallbacks[nameof(ListHeader)]) : Fallbacks[nameof(ListHeader)];
    public static string NoOneAfk => Keys.TryGetValue(nameof(NoOneAfk), out string? key) ? Loc.Text(key, Fallbacks[nameof(NoOneAfk)]) : Fallbacks[nameof(NoOneAfk)];
    public static string AfkTeamScope => Keys.TryGetValue(nameof(AfkTeamScope), out string? key) ? Loc.Text(key, Fallbacks[nameof(AfkTeamScope)]) : Fallbacks[nameof(AfkTeamScope)];
    public static string HarborFallback => Keys.TryGetValue(nameof(HarborFallback), out string? key) ? Loc.Text(key, Fallbacks[nameof(HarborFallback)]) : Fallbacks[nameof(HarborFallback)];
    public static string SwitchToggled => Keys.TryGetValue(nameof(SwitchToggled), out string? key) ? Loc.Text(key, Fallbacks[nameof(SwitchToggled)]) : Fallbacks[nameof(SwitchToggled)];
    public static string SwitchStateOn => Keys.TryGetValue(nameof(SwitchStateOn), out string? key) ? Loc.Text(key, Fallbacks[nameof(SwitchStateOn)]) : Fallbacks[nameof(SwitchStateOn)];
    public static string SwitchStateOff => Keys.TryGetValue(nameof(SwitchStateOff), out string? key) ? Loc.Text(key, Fallbacks[nameof(SwitchStateOff)]) : Fallbacks[nameof(SwitchStateOff)];
    public static string SwitchNotPaired => Keys.TryGetValue(nameof(SwitchNotPaired), out string? key) ? Loc.Text(key, Fallbacks[nameof(SwitchNotPaired)]) : Fallbacks[nameof(SwitchNotPaired)];
    public static string UpkeepNoTcMapped => Keys.TryGetValue(nameof(UpkeepNoTcMapped), out string? key) ? Loc.Text(key, Fallbacks[nameof(UpkeepNoTcMapped)]) : Fallbacks[nameof(UpkeepNoTcMapped)];
    public static string UpkeepTcEmptyExpired => Keys.TryGetValue(nameof(UpkeepTcEmptyExpired), out string? key) ? Loc.Text(key, Fallbacks[nameof(UpkeepTcEmptyExpired)]) : Fallbacks[nameof(UpkeepTcEmptyExpired)];
    public static string UpkeepNeed24h => Keys.TryGetValue(nameof(UpkeepNeed24h), out string? key) ? Loc.Text(key, Fallbacks[nameof(UpkeepNeed24h)]) : Fallbacks[nameof(UpkeepNeed24h)];
    public static string UpkeepTcTime => Keys.TryGetValue(nameof(UpkeepTcTime), out string? key) ? Loc.Text(key, Fallbacks[nameof(UpkeepTcTime)]) : Fallbacks[nameof(UpkeepTcTime)];
    public static string UpkeepDays => Keys.TryGetValue(nameof(UpkeepDays), out string? key) ? Loc.Text(key, Fallbacks[nameof(UpkeepDays)]) : Fallbacks[nameof(UpkeepDays)];
    public static string UpkeepHours => Keys.TryGetValue(nameof(UpkeepHours), out string? key) ? Loc.Text(key, Fallbacks[nameof(UpkeepHours)]) : Fallbacks[nameof(UpkeepHours)];
    public static string UpkeepMinutes => Keys.TryGetValue(nameof(UpkeepMinutes), out string? key) ? Loc.Text(key, Fallbacks[nameof(UpkeepMinutes)]) : Fallbacks[nameof(UpkeepMinutes)];
    public static string UpkeepEmptyExpiredShort => Keys.TryGetValue(nameof(UpkeepEmptyExpiredShort), out string? key) ? Loc.Text(key, Fallbacks[nameof(UpkeepEmptyExpiredShort)]) : Fallbacks[nameof(UpkeepEmptyExpiredShort)];
    public static string UpkeepTimeShort => Keys.TryGetValue(nameof(UpkeepTimeShort), out string? key) ? Loc.Text(key, Fallbacks[nameof(UpkeepTimeShort)]) : Fallbacks[nameof(UpkeepTimeShort)];
    public static string UpkeepHeader => Keys.TryGetValue(nameof(UpkeepHeader), out string? key) ? Loc.Text(key, Fallbacks[nameof(UpkeepHeader)]) : Fallbacks[nameof(UpkeepHeader)];
    public static string UpkeepNotPaired => Keys.TryGetValue(nameof(UpkeepNotPaired), out string? key) ? Loc.Text(key, Fallbacks[nameof(UpkeepNotPaired)]) : Fallbacks[nameof(UpkeepNotPaired)];
    public static string UpkeepNotPairedSingle => Keys.TryGetValue(nameof(UpkeepNotPairedSingle), out string? key) ? Loc.Text(key, Fallbacks[nameof(UpkeepNotPairedSingle)]) : Fallbacks[nameof(UpkeepNotPairedSingle)];
    public static string MaterialWood => Keys.TryGetValue(nameof(MaterialWood), out string? key) ? Loc.Text(key, Fallbacks[nameof(MaterialWood)]) : Fallbacks[nameof(MaterialWood)];
    public static string MaterialStone => Keys.TryGetValue(nameof(MaterialStone), out string? key) ? Loc.Text(key, Fallbacks[nameof(MaterialStone)]) : Fallbacks[nameof(MaterialStone)];
    public static string MaterialMetal => Keys.TryGetValue(nameof(MaterialMetal), out string? key) ? Loc.Text(key, Fallbacks[nameof(MaterialMetal)]) : Fallbacks[nameof(MaterialMetal)];
    public static string MaterialHQM => Keys.TryGetValue(nameof(MaterialHQM), out string? key) ? Loc.Text(key, Fallbacks[nameof(MaterialHQM)]) : Fallbacks[nameof(MaterialHQM)];
    public static string OilRigHackRunning => Keys.TryGetValue(nameof(OilRigHackRunning), out string? key) ? Loc.Text(key, Fallbacks[nameof(OilRigHackRunning)]) : Fallbacks[nameof(OilRigHackRunning)];
    public static string OilRigNone => Keys.TryGetValue(nameof(OilRigNone), out string? key) ? Loc.Text(key, Fallbacks[nameof(OilRigNone)]) : Fallbacks[nameof(OilRigNone)];
    public static string OilRigNoAlarm => Keys.TryGetValue(nameof(OilRigNoAlarm), out string? key) ? Loc.Text(key, Fallbacks[nameof(OilRigNoAlarm)]) : Fallbacks[nameof(OilRigNoAlarm)];
    public static string OilRigSmallLabel => Keys.TryGetValue(nameof(OilRigSmallLabel), out string? key) ? Loc.Text(key, Fallbacks[nameof(OilRigSmallLabel)]) : Fallbacks[nameof(OilRigSmallLabel)];
    public static string OilRigLargeLabel => Keys.TryGetValue(nameof(OilRigLargeLabel), out string? key) ? Loc.Text(key, Fallbacks[nameof(OilRigLargeLabel)]) : Fallbacks[nameof(OilRigLargeLabel)];
    public static string OilRigGenericLabel => Keys.TryGetValue(nameof(OilRigGenericLabel), out string? key) ? Loc.Text(key, Fallbacks[nameof(OilRigGenericLabel)]) : Fallbacks[nameof(OilRigGenericLabel)];
    public static string BaseCodeNameDefault => Keys.TryGetValue(nameof(BaseCodeNameDefault), out string? key) ? Loc.Text(key, Fallbacks[nameof(BaseCodeNameDefault)]) : Fallbacks[nameof(BaseCodeNameDefault)];
    public static string NoBaseCodes => Keys.TryGetValue(nameof(NoBaseCodes), out string? key) ? Loc.Text(key, Fallbacks[nameof(NoBaseCodes)]) : Fallbacks[nameof(NoBaseCodes)];
    public static string NoActiveTimers => Keys.TryGetValue(nameof(NoActiveTimers), out string? key) ? Loc.Text(key, Fallbacks[nameof(NoActiveTimers)]) : Fallbacks[nameof(NoActiveTimers)];
    public static string TimerMaxReached => Keys.TryGetValue(nameof(TimerMaxReached), out string? key) ? Loc.Text(key, Fallbacks[nameof(TimerMaxReached)]) : Fallbacks[nameof(TimerMaxReached)];
    public static string TimerNameMustStartWithLetter => Keys.TryGetValue(nameof(TimerNameMustStartWithLetter), out string? key) ? Loc.Text(key, Fallbacks[nameof(TimerNameMustStartWithLetter)]) : Fallbacks[nameof(TimerNameMustStartWithLetter)];
    public static string TimerCreated => Keys.TryGetValue(nameof(TimerCreated), out string? key) ? Loc.Text(key, Fallbacks[nameof(TimerCreated)]) : Fallbacks[nameof(TimerCreated)];
    public static string TimerDurationRequired => Keys.TryGetValue(nameof(TimerDurationRequired), out string? key) ? Loc.Text(key, Fallbacks[nameof(TimerDurationRequired)]) : Fallbacks[nameof(TimerDurationRequired)];
    public static string CountdownOneMinute => Keys.TryGetValue(nameof(CountdownOneMinute), out string? key) ? Loc.Text(key, Fallbacks[nameof(CountdownOneMinute)]) : Fallbacks[nameof(CountdownOneMinute)];
    public static string TimerExpiredTitle => Keys.TryGetValue(nameof(TimerExpiredTitle), out string? key) ? Loc.Text(key, Fallbacks[nameof(TimerExpiredTitle)]) : Fallbacks[nameof(TimerExpiredTitle)];
}

/// <summary>Reference ChatCommandMapping: one chat command name bound to one paired entity.</summary>
public sealed record ChatCommandMapping(uint EntityId, string Command);

/// <summary>
/// Reference ServerProfile.BaseCode: one labelled door code. Only complete codes are answered, because a
/// half-typed code in team chat would be tried on a door.
/// </summary>
public sealed record ChatBaseCode(string Label, string Code)
{
    public bool IsComplete => Code.Length == 4;

    /// <summary>Reference answer label: the trimmed custom name, else "Doorcode".</summary>
    public string DisplayLabel => string.IsNullOrWhiteSpace(Label) ? ChatCommandTexts.BaseCodeNameDefault : Label.Trim();
}

/// <summary>
/// Reference ServerProfile.ClanRolePermission: one clan role and whether its members may use chat commands.
/// </summary>
public sealed record ChatRolePermission(int RoleId, string Name, int Rank, bool Allowed);

/// <summary>
/// One clan member as the command gate sees them: the author of a clan message and the role they hold
/// (reference ClanMemberVM.SteamId / RoleId / RoleName).
/// </summary>
public sealed record ChatClanMember(ulong SteamId, int RoleId, string RoleName);

/// <summary>
/// One paired device as the switch and upkeep commands see it: the saved definition plus the live snapshot the device
/// panel last read. <see cref="UpkeepSeconds"/> is the reference ProtectSeconds (time left on the tool cupboard).
/// </summary>
public sealed record ChatDeviceState(PairedDevice Device, bool? IsOn = null, DeviceStorage? Storage = null,
    int? UpkeepSeconds = null)
{
    /// <summary>Reference SmartDevice.PureName: the alias, else the name, else the kind.</summary>
    public string PureName => string.IsNullOrWhiteSpace(Device.Alias)
        ? string.IsNullOrWhiteSpace(Device.Name) ? Device.Kind.ToString() : Device.Name
        : Device.Alias;

    /// <summary>Reference storage lookup test: unknown, protected (tool cupboard) or empty storage.</summary>
    public bool LooksLikeToolCupboard => Storage is null || Storage.HasProtection == true || Storage.Items.Count == 0;
}

/// <summary>One answer in a sequence, with the delay that precedes it (reference inter-TC delay).</summary>
public sealed record ChatMessageStep(string Text, double DelaySeconds = 0);

/// <summary>
/// One oil rig as the reference BuildOilRigAlarmAnswer reads it: the running crate timer (when the rig was called)
/// and the latest Smart Alarm for that rig. The port has no logic rules yet, so only alarms make a rig show up.
/// </summary>
public sealed record ChatOilRigState(string CanonicalRig, int? SecondsLeft = null, DateTime? LastAlarmUtc = null);

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
    bool CanListAi = false,
    IReadOnlyList<ChatCustomTimer>? CustomTimers = null,
    bool AlertCustomTimer = true,
    IReadOnlyList<ChatDeviceState>? Devices = null,
    Func<uint, bool, CancellationToken, Task<bool>>? SetSwitch = null,
    IReadOnlyList<ChatOilRigState>? OilRigs = null,
    IReadOnlyList<ChatClanMember>? ClanMembers = null);

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
public sealed record ChatCommandReply(string? Message = null, bool PromoteToLeader = false, string? Log = null,
    Func<CancellationToken, Task<ChatCommandReply?>>? Deferred = null, IReadOnlyList<ChatMessageStep>? FollowUp = null,
    ChatCustomTimer? NewCustomTimer = null)
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

        // Reference IsClanCommandAuthorAllowed, checked before the prefix and the cooldown: a clan message from an
        // author whose role was never granted is ignored outright, and it does not spend the cooldown either.
        if (request.Channel == ChatChannel.Clan)
        {
            if (!settings.ClanCommandsEnabled) return ChatCommandReply.Ignored;
            var member = (state.ClanMembers ?? []).FirstOrDefault(clan => clan.SteamId == request.SteamId);
            if (member is null)
                return new ChatCommandReply(null,
                    Log: $"[ChatCommand] Clan command from {request.Author} ignored: author not in the last clan pull.");
            if (!settings.IsClanRoleAllowed(member.RoleId))
                return new ChatCommandReply(null,
                    Log: $"[ChatCommand] Clan command from {request.Author} ignored: role '{member.RoleName}' is not permitted.");
        }

        string prefix = settings.EffectivePrefix;
        string cmd = request.Text.Trim().ToLowerInvariant();
        if (string.IsNullOrEmpty(cmd) || !cmd.StartsWith(prefix, StringComparison.Ordinal)) return ChatCommandReply.Ignored;

        if (!cooldown.TryEnter(now))
            return new ChatCommandReply(null, Log: $"[ChatCommand] Ignoring '{cmd}' from {request.Author} (Cooldown active)");

        cmd = cmd[prefix.Length..];

        if (!state.Connected) return ChatCommandReply.Ignored;

        string Log(string what) => $"[ChatCommand] {what} executed by {request.Author}";

        // Reference: clan chat sees switches and base codes only when the profile allows it. Team chat always does.
        bool allowSwitches = request.Channel != ChatChannel.Clan || settings.ClanCommandsAllowSwitches;
        bool allowBaseCodes = request.Channel != ChatChannel.Clan || settings.ClanCommandsAllowBaseCodes;

        if (cmd == settings.List.ToLowerInvariant())
        {
            string? deviceList = DeviceListAnswer(settings, state, allowSwitches);
            return new ChatCommandReply(ListAnswer(settings, state, allowBaseCodes), Log: Log("List"),
                FollowUp: deviceList is null ? null : [new ChatMessageStep(deviceList, 3)]);
        }

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

        if (cmd == settings.OilRig.ToLowerInvariant())
            return new ChatCommandReply(OilRigAnswer(state));

        if (cmd == settings.Heli.ToLowerInvariant())
            return new ChatCommandReply(HeliAnswer(state.Events ?? new ChatEventCommandState(), now), Log: Log("Heli"));

        if (cmd == settings.Vendor.ToLowerInvariant())
            return new ChatCommandReply(VendorAnswer(state.Events ?? new ChatEventCommandState(), now), Log: Log("Vendor"));

        // Custom timers answer to their own command names before the create/list command, as the reference does.
        foreach (var timer in state.CustomTimers ?? [])
        {
            if (!string.Equals(cmd, timer.Command.ToLowerInvariant(), StringComparison.Ordinal)) continue;
            TimeSpan remaining = timer.EndTimeUtc - now;
            return remaining.TotalSeconds <= 0
                ? ChatCommandReply.Ignored
                : new ChatCommandReply($"{timer.Name}: {timer.RemainingText(now)}",
                    Log: $"[ChatCommand] Timer '{timer.Name}' checked by {request.Author}");
        }

        string timerCommand = settings.CustomTimer.ToLowerInvariant();
        if (cmd == timerCommand)
        {
            var timers = state.CustomTimers ?? [];
            string message = timers.Count == 0
                ? ChatCommandTexts.NoActiveTimers
                : string.Join(" | ", timers.Select(timer => $"{prefix}{timer.Command} : {timer.RemainingText(now)}"));
            return new ChatCommandReply(message);
        }
        if (cmd.StartsWith(timerCommand + " ", StringComparison.Ordinal))
        {
            if ((state.CustomTimers?.Count ?? 0) >= ChatCustomTimer.MaxTimers)
                return new ChatCommandReply(ChatCommandTexts.TimerMaxReached);

            string[] args = cmd[(timerCommand.Length + 1)..].Split(',');
            if (args.Length != 2) return ChatCommandReply.Ignored;
            string timerName = args[0].Trim();
            string timePart = args[1].Trim();

            int hours = 0, minutes = 0, seconds = 0;
            if (timePart.Contains(':'))
            {
                string[] parts = timePart.Split(':');
                if (parts.Length == 3)
                {
                    _ = int.TryParse(parts[0], out hours);
                    _ = int.TryParse(parts[1], out minutes);
                    _ = int.TryParse(parts[2], out seconds);
                }
                else if (parts.Length == 2)
                {
                    _ = int.TryParse(parts[0], out minutes);
                    _ = int.TryParse(parts[1], out seconds);
                }
            }
            else
            {
                _ = int.TryParse(timePart, out minutes);
            }

            int totalSeconds = hours * 3600 + minutes * 60 + seconds;
            if (totalSeconds <= 0) return ChatCommandReply.Ignored;
            if (!ChatCustomTimer.IsValidName(timerName))
                return new ChatCommandReply(ChatCommandTexts.TimerNameMustStartWithLetter);

            string command = ChatCustomTimer.CommandFor(timerName);
            double totalMinutes = totalSeconds / 60.0;
            var timer = new ChatCustomTimer(Guid.NewGuid().ToString(), timerName, command,
                now.AddSeconds(totalSeconds), request.Channel, CreatedNotified: false,
                Notified60: totalMinutes <= 60, Notified30: totalMinutes <= 30,
                Notified10: totalMinutes <= 10, Notified3: totalMinutes <= 3);
            string? created = state.AlertCustomTimer
                ? string.Format(CultureInfo.InvariantCulture, ChatCommandTexts.TimerCreated, prefix + command,
                    hours, minutes, seconds)
                : null;
            return new ChatCommandReply(created,
                Log: $"[ChatCommand] Timer created by {request.Author}: {timerName} for {hours}h {minutes}m {seconds}s",
                NewCustomTimer: timer);
        }

        var switchMappings = allowSwitches ? Mappings(settings.SwitchMappings, cmd) : [];
        if (switchMappings.Count > 0)
        {
            var seeds = switchMappings
                .Select(mapping => FindDevice(state, mapping.EntityId, DeviceKind.Switch, allowGroup: true))
                .Where(device => device is not null).Select(device => device!).ToList();
            // Reference: a mapping whose device is gone stays silent, while a group without switches says "not paired".
            if (seeds.Count == 0) return ChatCommandReply.Ignored;

            var leaves = seeds.SelectMany(seed => SwitchLeaves(state, seed.Device))
                .GroupBy(leaf => leaf.Device.EntityId).Select(group => group.First()).ToList();
            if (leaves.Count == 0) return new ChatCommandReply(ChatCommandTexts.SwitchNotPaired);

            bool targetOn = !(leaves[0].IsOn ?? false);
            if (state.SetSwitch is not { } setSwitch) return ChatCommandReply.Ignored;
            return new ChatCommandReply(Deferred: ct => ToggleAsync(setSwitch, leaves, targetOn, request.Author, ct));
        }

        if (cmd == settings.UpkeepDetail.ToLowerInvariant())
        {
            var cupboards = (settings.UpkeepMappings ?? []).Where(mapping => mapping.EntityId != 0).ToList();
            if (cupboards.Count == 0)
                return new ChatCommandReply(ChatCommandTexts.UpkeepNoTcMapped, Log: Log("UpkeepDetail"));

            var steps = new List<ChatMessageStep>();
            foreach (var mapping in cupboards)
            {
                var device = FindDevice(state, mapping.EntityId, DeviceKind.StorageMonitor);
                if (device is null || !device.LooksLikeToolCupboard) continue;
                int seconds = device.UpkeepSeconds ?? 0;
                string text = seconds <= 0
                    ? string.Format(CultureInfo.InvariantCulture, ChatCommandTexts.UpkeepTcEmptyExpired, device.PureName)
                    : string.Format(CultureInfo.InvariantCulture, ChatCommandTexts.UpkeepTcTime, device.PureName,
                        UpkeepTime(seconds)) + UpkeepMaterials(device, seconds);
                steps.Add(new ChatMessageStep(text,
                    steps.Count == 0 ? 0 : Math.Max(2.0, settings.ResponseDelaySeconds)));
            }
            return new ChatCommandReply(Log: Log("UpkeepDetail"), FollowUp: steps);
        }

        if (allowBaseCodes && !string.IsNullOrWhiteSpace(settings.BaseCodes) && cmd == settings.BaseCodes.ToLowerInvariant())
        {
            var parts = (settings.BaseCodeRows ?? [])
                .Where(code => code.IsComplete)
                .Select(code => $"{code.DisplayLabel}: {code.Code}")
                .ToList();
            if (parts.Count == 0) return new ChatCommandReply(ChatCommandTexts.NoBaseCodes, Log: Log("BaseCodes"));

            string message = string.Join(" | ", parts);
            // Team chat truncates at 128 characters: drop whole entries rather than cut a code in half.
            while (message.Length > ChatLimits.MaxLength && parts.Count > 1)
            {
                parts.RemoveAt(parts.Count - 1);
                message = string.Join(" | ", parts) + " ...";
            }
            return new ChatCommandReply(message, Log: Log("BaseCodes"));
        }

        var upkeepMappings = Mappings(settings.UpkeepMappings, cmd);
        if (upkeepMappings.Count == 1)
        {
            var device = FindDevice(state, upkeepMappings[0].EntityId, DeviceKind.StorageMonitor);
            string? upkeepLog = device is not null && device.LooksLikeToolCupboard
                ? $"[ChatCommand] Upkeep for {device.Device.Name} executed by {request.Author}"
                : null;
            return new ChatCommandReply(UpkeepSingleAnswer(device), Log: upkeepLog);
        }
        if (upkeepMappings.Count > 1)
        {
            var parts = new List<string>();
            foreach (var mapping in upkeepMappings)
            {
                var device = FindDevice(state, mapping.EntityId, DeviceKind.StorageMonitor);
                if (device is null || !device.LooksLikeToolCupboard) continue;
                int seconds = device.UpkeepSeconds ?? 0;
                parts.Add(seconds <= 0
                    ? string.Format(CultureInfo.InvariantCulture, ChatCommandTexts.UpkeepEmptyExpiredShort, device.PureName)
                    : string.Format(CultureInfo.InvariantCulture, ChatCommandTexts.UpkeepTimeShort, device.PureName,
                        seconds / 86400, seconds % 86400 / 3600));
            }
            return new ChatCommandReply(parts.Count > 0
                ? string.Format(CultureInfo.InvariantCulture, ChatCommandTexts.UpkeepHeader, string.Join(" | ", parts))
                : ChatCommandTexts.UpkeepNotPaired,
                Log: $"[ChatCommand] Multi-Upkeep for cmd={cmd} executed by {request.Author}");
        }

        return ChatCommandReply.Ignored;
    }

    private static List<ChatCommandMapping> Mappings(IReadOnlyList<ChatCommandMapping>? mappings, string cmd) =>
        (mappings ?? []).Where(mapping => mapping.EntityId != 0 &&
            cmd == (mapping.Command ?? "").ToLowerInvariant()).ToList();

    /// <summary>
    /// Reference BuildOilRigAlarmAnswer (without the logic-rule half, which is not ported yet): small and large
    /// answers when both are known, otherwise the one that is, then a generic alarm, then the first alarm seen.
    /// </summary>
    private static string OilRigAnswer(ChatCommandState state)
    {
        var rigs = state.OilRigs ?? [];
        var small = rigs.FirstOrDefault(rig => rig.CanonicalRig == OilRigTimer.SmallRig);
        var large = rigs.FirstOrDefault(rig => rig.CanonicalRig == OilRigTimer.LargeRig);
        var generic = rigs.FirstOrDefault(rig => rig.CanonicalRig == OilRigTimer.GenericRig);

        if (small is not null && large is not null)
            return string.Join(" | ", SingleRig(OilRigTimer.SmallRig, ChatCommandTexts.OilRigSmallLabel, small),
                SingleRig(OilRigTimer.LargeRig, ChatCommandTexts.OilRigLargeLabel, large));
        if (small is not null) return SingleRig(OilRigTimer.SmallRig, ChatCommandTexts.OilRigSmallLabel, small);
        if (large is not null) return SingleRig(OilRigTimer.LargeRig, ChatCommandTexts.OilRigLargeLabel, large);
        if (generic is not null) return SingleRig(OilRigTimer.GenericRig, ChatCommandTexts.OilRigGenericLabel, generic);
        if (rigs.Count > 0)
        {
            var first = rigs[0];
            string label = first.CanonicalRig switch
            {
                OilRigTimer.SmallRig => ChatCommandTexts.OilRigSmallLabel,
                OilRigTimer.LargeRig => ChatCommandTexts.OilRigLargeLabel,
                _ => ChatCommandTexts.OilRigGenericLabel
            };
            return SingleRig(first.CanonicalRig, label, first);
        }
        return ChatCommandTexts.OilRigNoAlarm;
    }

    /// <summary>Reference BuildSingleRigAnswer: running crate timer, else the last alarm time, else "None".</summary>
    private static string SingleRig(string canonicalRig, string label, ChatOilRigState rig)
    {
        if (rig.SecondsLeft is { } seconds && seconds > 0)
            return string.Format(CultureInfo.InvariantCulture, ChatCommandTexts.OilRigHackRunning, label,
                seconds / 60, seconds % 60);
        if (rig.LastAlarmUtc is { } alarm)
            return $"{label}: {alarm.ToLocalTime():g}";
        return string.Format(CultureInfo.InvariantCulture, ChatCommandTexts.OilRigNone, label);
    }

    private static ChatDeviceState? FindDevice(ChatCommandState state, uint entityId, DeviceKind kind, bool allowGroup = false) =>
        state.Devices?.FirstOrDefault(device => device.Device.EntityId == entityId &&
            (device.Device.Kind == kind || allowGroup && device.Device.Kind == DeviceKind.Group));

    /// <summary>Reference AddSwitches: recurse through groups and collect the SmartSwitch leaves.</summary>
    private static IEnumerable<ChatDeviceState> SwitchLeaves(ChatCommandState state, PairedDevice device)
    {
        if (device.Kind == DeviceKind.Group)
        {
            foreach (var child in device.Children)
                foreach (var leaf in SwitchLeaves(state, child)) yield return leaf;
        }
        else if (device.Kind == DeviceKind.Switch && state.Devices is { } devices)
        {
            var live = devices.FirstOrDefault(item => item.Device.EntityId == device.EntityId);
            if (live is not null) yield return live;
        }
    }

    /// <summary>Reference switch branch: one target state from the first leaf, 800 ms between toggles.</summary>
    private static async Task<ChatCommandReply?> ToggleAsync(Func<uint, bool, CancellationToken, Task<bool>> setSwitch,
        IReadOnlyList<ChatDeviceState> leaves, bool targetOn, string author, CancellationToken ct)
    {
        var names = new List<string>();
        foreach (var leaf in leaves)
        {
            if (leaf.IsOn == targetOn) continue;
            try
            {
                if (await setSwitch(leaf.Device.EntityId, targetOn, ct).ConfigureAwait(false))
                    names.Add(leaf.PureName);
            }
            catch (OperationCanceledException) when (ct.IsCancellationRequested) { return null; }
            catch (Exception) { }
            try { await Task.Delay(800, ct).ConfigureAwait(false); }
            catch (OperationCanceledException) { break; }
        }
        if (names.Count == 0) return null;

        string stateName = targetOn ? ChatCommandTexts.SwitchStateOn : ChatCommandTexts.SwitchStateOff;
        string text = names.Count == 1 ? names[0] : string.Join(", ", names);
        if (names.Count > 1 && text.Length > 80) text = text[..77] + "...";
        return new ChatCommandReply(
            string.Format(CultureInfo.InvariantCulture, ChatCommandTexts.SwitchToggled, text, stateName),
            Log: $"[ChatCommand] Toggled {names.Count} switches to {targetOn} by {author}");
    }

    private static string UpkeepSingleAnswer(ChatDeviceState? device)
    {
        if (device is null || !device.LooksLikeToolCupboard) return ChatCommandTexts.UpkeepNotPairedSingle;
        int seconds = device.UpkeepSeconds ?? 0;
        return seconds <= 0
            ? string.Format(CultureInfo.InvariantCulture, ChatCommandTexts.UpkeepTcEmptyExpired, device.PureName)
            : string.Format(CultureInfo.InvariantCulture, ChatCommandTexts.UpkeepTcTime, device.PureName,
                UpkeepTime(seconds));
    }

    private static string UpkeepTime(int seconds)
    {
        int days = seconds / 86400;
        int remainder = seconds % 86400;
        int hours = remainder / 3600;
        int minutes = remainder % 3600 / 60;
        var parts = new List<string>();
        if (days > 0) parts.Add(string.Format(CultureInfo.InvariantCulture, ChatCommandTexts.UpkeepDays, days));
        if (hours > 0 || days > 0) parts.Add(string.Format(CultureInfo.InvariantCulture, ChatCommandTexts.UpkeepHours, hours));
        parts.Add(string.Format(CultureInfo.InvariantCulture, ChatCommandTexts.UpkeepMinutes, minutes));
        return string.Join(", ", parts);
    }

    private static string UpkeepMaterials(ChatDeviceState device, int seconds)
    {
        if (seconds <= 0 || device.Storage?.Items is not { Count: > 0 } items) return "";
        string materials = string.Join(", ", items
            .Select(item => (Item: item, Sort: UpkeepMaterialSort(item.ItemId)))
            .Where(entry => entry.Sort > 0)
            .GroupBy(entry => entry.Sort)
            .Select(group => new
            {
                Sort = group.Key,
                Name = UpkeepMaterialName(group.Key),
                Amount = (int)Math.Ceiling(group.Sum(entry => Math.Max(0, entry.Item.Quantity ?? 0)) * 86400.0 / seconds)
            })
            .Where(material => material.Amount > 0)
            .OrderBy(material => material.Sort)
            .Select(material => $"{material.Name} {material.Amount.ToString("N0", CultureInfo.CurrentCulture).Replace(",", "")}"));
        return string.IsNullOrEmpty(materials)
            ? ""
            : string.Format(CultureInfo.InvariantCulture, ChatCommandTexts.UpkeepNeed24h, materials);
    }

    /// <summary>
    /// Reference hardcoded upkeep item IDs: wood, stone, metal fragments, high quality metal. The reference also
    /// matches short names; the port's snapshot only carries the ID, which is the reliable half of that test.
    /// </summary>
    private static int UpkeepMaterialSort(int itemId) => itemId switch
    {
        -151838493 => 10,
        -2099697608 => 20,
        69511070 => 30,
        317398316 => 40,
        _ => 0
    };

    private static string UpkeepMaterialName(int sort) => sort switch
    {
        10 => ChatCommandTexts.MaterialWood,
        20 => ChatCommandTexts.MaterialStone,
        30 => ChatCommandTexts.MaterialMetal,
        40 => ChatCommandTexts.MaterialHQM,
        _ => ""
    };

    /// <summary>Reference list: each switch/TC mapping whose device still exists, sent three seconds later.</summary>
    public static string? DeviceListAnswer(ChatCommandSettings settings, ChatCommandState state, bool allowSwitches = true)
    {
        if (state.Devices is null) return null;
        string prefix = settings.EffectivePrefix;
        var parts = new List<string>();
        if (allowSwitches)
        {
            foreach (var mapping in settings.SwitchMappings ?? [])
            {
                if (string.IsNullOrWhiteSpace(mapping.Command) || mapping.EntityId == 0) continue;
                var device = FindDevice(state, mapping.EntityId, DeviceKind.Switch);
                if (device is not null) parts.Add($"[{device.PureName}]: {prefix}{mapping.Command}");
            }
        }
        foreach (var mapping in settings.UpkeepMappings ?? [])
        {
            if (string.IsNullOrWhiteSpace(mapping.Command) || mapping.EntityId == 0) continue;
            var device = FindDevice(state, mapping.EntityId, DeviceKind.StorageMonitor);
            if (device is not null) parts.Add($"[{device.PureName}]: {prefix}{mapping.Command}");
        }
        return parts.Count == 0 ? null : Truncate(string.Join(" | ", parts));
    }

    /// <summary>
    /// Reference list command. Advertises a command only when it can do something: the timer entry only once a
    /// timer exists, base codes only to a channel allowed to see them and only once a code is set, and the AI
    /// command only to an author allowed to use it.
    /// </summary>
    public static string ListAnswer(ChatCommandSettings settings, ChatCommandState state, bool allowBaseCodes = true)
    {
        string prefix = settings.EffectivePrefix;
        var standard = new List<string>();
        if (!string.IsNullOrWhiteSpace(state.CustomTimers?.FirstOrDefault()?.Command))
            standard.Add(prefix + state.CustomTimers![0].Command);
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
        if (!string.IsNullOrWhiteSpace(settings.BaseCodes) && (settings.BaseCodeRows ?? []).Any(code => code.IsComplete) && allowBaseCodes)
            Add(standard, prefix, settings.BaseCodes);
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