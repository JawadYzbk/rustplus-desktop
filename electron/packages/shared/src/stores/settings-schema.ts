/**
 * tracking_settings.json contract — full key catalog + defaults transcribed from the legacy app
 * (audit DATA_STORES §2; defaults source TrackingService.cs:61-194). Zero-key drift tolerated:
 * the parity test pins documented defaults, and the store rejects unknown keys so renames fail loudly.
 */
import { z } from "zod";

/** NaN cannot exist in JSON; the legacy app (System.Text.Json named-float literals) persists "NaN". */
const windowCoord = z.union([z.number(), z.literal("NaN")]);

export const trackingSettingsSchema = z
  .object({
    // --- Connection/session ---
    LastHost: z.string(),
    LastPort: z.number().int(),
    LastServerName: z.string(),
    LastBMId: z.string().nullable(),
    LastSelectedServerKey: z.string(),
    AutoConnectEnabled: z.boolean(),
    AutoStartEnabled: z.boolean(),
    HideConsole: z.boolean(),
    BackgroundTrackingEnabled: z.boolean(),
    SteamId64: z.string(),
    FcmIssuedAt: z.string().nullable(),
    FcmExpiresAt: z.string().nullable(),
    LastSeenVersion: z.string(),
    SuppressVersion8Notice: z.boolean(),

    // --- Window/UI ---
    SidebarWidth: z.number(),
    SidebarPinned: z.boolean(),
    WindowWidth: z.number(),
    WindowHeight: z.number(),
    WindowLeft: windowCoord,
    WindowTop: windowCoord,
    WindowMaximized: z.boolean(),
    CloseToTrayEnabled: z.boolean(),
    StartMinimizedEnabled: z.boolean(),
    SelectedLanguage: z.string(),

    // --- Map display ---
    MapShowSteamMarkers: z.boolean(),
    MapShowPlayerArrows: z.boolean(),
    MapShowDeathTags: z.boolean(),
    MapShowDeathHeatmap: z.boolean(),
    MaxSelfDeathMarkers: z.number().int(),
    MaxTeamDeathMarkers: z.number().int(),
    MapAbbreviateNames: z.boolean(),
    MapPlayerIconScale: z.number(),
    MapUseMonumentText: z.boolean(),
    MapMonumentDisplayMode: z.number().int(),
    MapMonumentScale: z.number(),
    MapMonumentOpacity: z.number(),
    MapGridOpacity: z.number(),
    HiddenExtraMonumentTypes: z.array(z.unknown()),

    // --- Events/announcements ---
    AnnounceCargo: z.boolean(),
    AnnounceHeli: z.boolean(),
    AnnounceChinook: z.boolean(),
    AnnounceVendor: z.boolean(),
    AnnounceOilRig: z.boolean(),
    AnnounceDeepSea: z.boolean(),
    AnnounceCargoDocking: z.boolean(),
    AnnounceCargoEgress: z.boolean(),
    AnnounceCargoArrival: z.boolean(),
    ListenForServerEvents: z.boolean(),
    TrustOwnDetections: z.boolean(),
    AnnounceSpawnsMaster: z.boolean(),
    ChatMasterOfferSoundEnabled: z.boolean(),
    // Learned dictionaries — value shapes are legacy-runtime internals; kept opaque to stay lossless.
    LearnedDockingDurations: z.record(z.unknown()),
    LearnedCargoFullLifeMinutes: z.record(z.unknown()),
    LearnedCargoTravelMinutes: z.record(z.unknown()),
    ServerHarbors: z.record(z.unknown()),
    ServerCargoTriggers: z.record(z.unknown()),

    // --- Players/deaths ---
    AnnouncePlayerOnline: z.boolean(),
    AnnouncePlayerOffline: z.boolean(),
    AnnouncePlayerAfk: z.boolean(),
    AnnouncePlayerAfkReturn: z.boolean(),
    AfkAlertMinutes: z.number().int(),
    AnnouncePlayerDeathSelf: z.boolean(),
    AnnouncePlayerDeathTeam: z.boolean(),
    AnnouncePlayerRespawnSelf: z.boolean(),
    AnnouncePlayerRespawnTeam: z.boolean(),
    OfflineDeathAlertsEnabled: z.boolean(),
    OfflineDeathSoundPath: z.string(),
    OfflineDeathSoundLoopEnabled: z.boolean(),
    OfflineDeathDiscordEnabled: z.boolean(),
    OfflineDeathHistory: z.array(z.unknown()),

    // --- Shops ---
    AutoLoadShops: z.boolean(),
    AnnounceNewShops: z.boolean(),
    AnnounceSuspiciousShops: z.boolean(),
    AnnounceTradeAlerts: z.boolean(),

    // --- Alarms/devices ---
    AnnounceSmartAlerts: z.boolean(),
    GenericAlarmPopupEnabled: z.boolean(),
    GenericAlarmOverlayEnabled: z.boolean(),
    GenericAlarmAudioEnabled: z.boolean(),
    GenericAlarmAudioFilePath: z.string(),
    GroupStates: z.record(z.unknown()),
    GroupOrder: z.record(z.unknown()),
    HotkeyTriggerChatAlertsEnabled: z.boolean(),
    /** keyed "host:port|entityId" → bool */
    HotkeyTriggerChatAlertEnabled: z.record(z.boolean()),
    LearnedQueryPorts: z.record(z.unknown()),

    // --- Integrations (token-bearing URLs; encrypted at rest by the store layer) ---
    DiscordWebhookUrl: z.string(),
    /** "None" | "@everyone" | "@here" (legacy free-form string; validated loosely) */
    DiscordWebhookMention: z.string(),
    SmartHomeWebhookUrl: z.string(),
    TelegramCallWebhookUrl: z.string(),
    TelegramCallUser: z.string(),
    TelegramCallMsg: z.string(),
    TelegramCallLang: z.string(),
    TelegramCallIncTitle: z.boolean(),
    IncMsg: z.boolean(),
    IncType: z.boolean(),

    // --- Cloud consent/freemium ---
    TranslationConsentGiven: z.boolean(),
    UploadConsentGiven: z.boolean(),
    CloudSyncEnabled: z.boolean(),
    PlayerWipeTrackerEnabled: z.boolean(),
    PlayerWipeTrackerCloudBackupEnabled: z.boolean(),

    // --- Notification center ---
    NotificationsToastEnabled: z.boolean(),
    NotificationsSoundsEnabled: z.boolean(),
    NotificationsRetentionDays: z.number().int(),
    MutedNotificationServers: z.array(z.unknown()),
    MutedNotificationServerNames: z.record(z.unknown()),
    ServerFollowingSteamId: z.record(z.unknown()),

    // --- Misc ---
    SaveAlertSelection: z.boolean(),
    AnnounceTracking: z.boolean(),
    LastCrosshairStyle: z.string(),
    LastCustomCrosshairId: z.string(),
  })
  .strict();

export type TrackingSettings = z.infer<typeof trackingSettingsSchema>;

/** Exact legacy defaults (audit DATA_STORES §2). Order mirrors TrackingService.cs:61-194 groupings. */
export const TRACKING_SETTINGS_DEFAULTS: TrackingSettings = {
  LastHost: "",
  LastPort: 0,
  LastServerName: "",
  LastBMId: null,
  LastSelectedServerKey: "",
  AutoConnectEnabled: false,
  AutoStartEnabled: false,
  HideConsole: false,
  BackgroundTrackingEnabled: true,
  SteamId64: "",
  FcmIssuedAt: null,
  FcmExpiresAt: null,
  LastSeenVersion: "",
  SuppressVersion8Notice: false,

  SidebarWidth: 420,
  SidebarPinned: true,
  WindowWidth: 1280,
  WindowHeight: 720,
  WindowLeft: "NaN",
  WindowTop: "NaN",
  WindowMaximized: false,
  CloseToTrayEnabled: false,
  StartMinimizedEnabled: false,
  SelectedLanguage: "",

  MapShowSteamMarkers: true,
  MapShowPlayerArrows: true,
  MapShowDeathTags: false,
  MapShowDeathHeatmap: false,
  MaxSelfDeathMarkers: 3,
  MaxTeamDeathMarkers: 3,
  MapAbbreviateNames: false,
  MapPlayerIconScale: 1.0,
  MapUseMonumentText: false,
  MapMonumentDisplayMode: 0,
  MapMonumentScale: 1.0,
  MapMonumentOpacity: 1.0,
  MapGridOpacity: 0.7,
  HiddenExtraMonumentTypes: [],

  AnnounceCargo: false,
  AnnounceHeli: false,
  AnnounceChinook: false,
  AnnounceVendor: false,
  AnnounceOilRig: false,
  AnnounceDeepSea: false,
  AnnounceCargoDocking: false,
  AnnounceCargoEgress: false,
  AnnounceCargoArrival: false,
  ListenForServerEvents: true,
  TrustOwnDetections: true,
  AnnounceSpawnsMaster: false,
  ChatMasterOfferSoundEnabled: true,
  LearnedDockingDurations: {},
  LearnedCargoFullLifeMinutes: {},
  LearnedCargoTravelMinutes: {},
  ServerHarbors: {},
  ServerCargoTriggers: {},

  AnnouncePlayerOnline: false,
  AnnouncePlayerOffline: false,
  AnnouncePlayerAfk: false,
  AnnouncePlayerAfkReturn: false,
  AfkAlertMinutes: 5,
  AnnouncePlayerDeathSelf: false,
  AnnouncePlayerDeathTeam: false,
  AnnouncePlayerRespawnSelf: false,
  AnnouncePlayerRespawnTeam: false,
  OfflineDeathAlertsEnabled: true,
  OfflineDeathSoundPath: "",
  OfflineDeathSoundLoopEnabled: false,
  OfflineDeathDiscordEnabled: false,
  OfflineDeathHistory: [],

  AutoLoadShops: true,
  AnnounceNewShops: false,
  AnnounceSuspiciousShops: false,
  AnnounceTradeAlerts: false,

  AnnounceSmartAlerts: false,
  GenericAlarmPopupEnabled: false,
  GenericAlarmOverlayEnabled: true,
  GenericAlarmAudioEnabled: true,
  GenericAlarmAudioFilePath: "",
  GroupStates: {},
  GroupOrder: {},
  HotkeyTriggerChatAlertsEnabled: true,
  HotkeyTriggerChatAlertEnabled: {},
  LearnedQueryPorts: {},

  DiscordWebhookUrl: "",
  DiscordWebhookMention: "None",
  SmartHomeWebhookUrl: "",
  TelegramCallWebhookUrl: "",
  TelegramCallUser: "",
  TelegramCallMsg: "Alarm ausgeloest!",
  TelegramCallLang: "de-DE-Standard-A",
  TelegramCallIncTitle: true,
  IncMsg: true,
  IncType: false,

  TranslationConsentGiven: false,
  UploadConsentGiven: false,
  CloudSyncEnabled: false,
  PlayerWipeTrackerEnabled: false,
  PlayerWipeTrackerCloudBackupEnabled: false,

  NotificationsToastEnabled: true,
  NotificationsSoundsEnabled: true,
  NotificationsRetentionDays: 30,
  MutedNotificationServers: [],
  MutedNotificationServerNames: {},
  ServerFollowingSteamId: {},

  SaveAlertSelection: true,
  AnnounceTracking: false,
  LastCrosshairStyle: "GreenDot",
  LastCustomCrosshairId: "",
};
