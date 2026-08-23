/**
 * Contracts for the remaining legacy JSON stores (audit DATA_STORES §1), transcribed from source:
 *  - hotkeys.json          MainWindow.xaml.cs:7817/7931   Dict<serverKey, Dict<gesture, List<long>>>
 *  - hotkey_options.json   MainWindow.xaml.cs:7825        { ParallelMode=false, ToggleDelayMs=150 }
 *  - custom_alerts.json    AlertTemplateService.cs:18     Dict<culture, Dict<key, template>>
 *  - tracked_players.json  TrackingService.cs:13/203      List<TrackedPlayer>
 *  - tutorial-progress.json TutorialProgressStore.cs:24   { Tutorials{}, Preferences }
 *
 * Field names stay PascalCase (System.Text.Json default) — byte-compatible with files the C# app wrote.
 * TutorialStatus is a numeric enum there (no JsonStringEnumConverter): NotStarted=0…Updated=4.
 */
import { z } from "zod";

// --- Device hotkeys ----------------------------------------------------------

/** serverKey ("host:port|steamId") → gesture ("Ctrl+F1", …) → entityIds */
export const deviceHotkeysSchema = z.record(z.record(z.array(z.number())));
export type DeviceHotkeys = z.infer<typeof deviceHotkeysSchema>;

export const HOTKEY_OPTIONS_DEFAULTS = {
  ParallelMode: false,
  ToggleDelayMs: 150,
};
export const hotkeyOptionsSchema = z
  .object({
    ParallelMode: z.boolean(),
    ToggleDelayMs: z.number().int(),
  })
  .strict();
export type HotkeyOptions = z.infer<typeof hotkeyOptionsSchema>;

// --- Alert template overrides -------------------------------------------------

/** culture → template key → override text (top-level keys matched case-insensitively by the legacy app) */
export const customAlertsSchema = z.record(z.record(z.string()));
export type CustomAlerts = z.infer<typeof customAlertsSchema>;

export function findAlertOverride(alerts: CustomAlerts, culture: string, key: string): string | undefined {
  const target = Object.keys(alerts).find((k) => k.toLowerCase() === culture.toLowerCase());
  return target ? alerts[target]?.[key] : undefined;
}

// --- Tracked players -----------------------------------------------------------

export const playerSessionSchema = z.object({
  Name: z.string(),
  BMId: z.string(),
  SessionStartTimeUtc: z.string(),
  /** Legacy System.TimeSpan — serialized in "c" format ("1.02:03:04"). Kept opaque; math lands with stage 4/5. */
  Duration: z.string(),
  IsTracked: z.boolean(),
});
export type PlayerSessionModel = z.infer<typeof playerSessionSchema>;

export const trackedPlayerSchema = z.object({
  BMId: z.string(),
  Name: z.string(),
  LastServerName: z.string(),
  GroupName: z.string(),
  GroupColor: z.string(),
  Sessions: z.array(playerSessionSchema),
  IsBMOnly: z.boolean().optional(),
});
export type TrackedPlayerModel = z.infer<typeof trackedPlayerSchema>;

export const trackedPlayersFileSchema = z.object({
  schemaVersion: z.number().int(),
  players: z.array(trackedPlayerSchema.passthrough()),
});
export type TrackedPlayersFile = z.infer<typeof trackedPlayersFileSchema>;

// --- Tutorial progress -----------------------------------------------------------

export const TUTORIAL_STATUS = {
  NotStarted: 0,
  InProgress: 1,
  Completed: 2,
  Skipped: 3,
  Updated: 4,
} as const;
export type TutorialStatusValue = (typeof TUTORIAL_STATUS)[keyof typeof TUTORIAL_STATUS];

const tutorialStatusSchema = z.union([
  z.literal(0),
  z.literal(1),
  z.literal(2),
  z.literal(3),
  z.literal(4),
]);

const tutorialProgressSchema = z.object({
  TutorialId: z.string(),
  TutorialVersion: z.number().int(),
  Status: tutorialStatusSchema,
  LastCompletedStepId: z.string().nullable().optional(),
  CompletedStepIds: z.array(z.string()),
  StartedAtUtc: z.string().nullable().optional(),
  CompletedAtUtc: z.string().nullable().optional(),
  SkippedAtUtc: z.string().nullable().optional(),
});
export type TutorialProgressModel = z.infer<typeof tutorialProgressSchema>;
export { tutorialProgressSchema };

export const tutorialPreferencesSchema = z.object({
  FirstRunPromptDismissed: z.boolean(),
  AutoStartBasicTutorial: z.boolean(),
  AutoStartNewFeatureTutorials: z.boolean(),
  OfferedTutorialIds: z.array(z.string()),
  LastTutorialId: z.string().nullable().optional(),
});
export type TutorialPreferences = z.infer<typeof tutorialPreferencesSchema>;

export const TUTORIAL_PREFERENCES_DEFAULTS: TutorialPreferences = {
  FirstRunPromptDismissed: false,
  AutoStartBasicTutorial: true,
  AutoStartNewFeatureTutorials: true,
  OfferedTutorialIds: [],
};

export const tutorialProgressFileSchema = z.object({
  schemaVersion: z.number().int(),
  Tutorials: z.record(tutorialProgressSchema),
  Preferences: tutorialPreferencesSchema,
});
export type TutorialProgressFile = z.infer<typeof tutorialProgressFileSchema>;
