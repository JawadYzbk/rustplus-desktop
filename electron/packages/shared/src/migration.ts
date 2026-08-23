/**
 * Legacy-data migration (M3) contract — importing C#-era stores from %APPDATA%\RustPlusDesk et al.
 * into the new RustPlusDesk-Electron root. Audit DATA_STORES §1/§4/§6 is the source of truth for
 * what migrates, what copies verbatim, and what defers to later stages.
 */
import { z } from "zod";

export const LEGACY_SOURCE_IDS = [
  "tracking_settings",
  "profiles",
  "hotkeys",
  "hotkey_options",
  "custom_alerts",
  "tracked_players",
  "tutorial_progress",
  "fcm_config",
  "cache_files",
] as const;
export type LegacySourceId = (typeof LEGACY_SOURCE_IDS)[number];

export const migrationStatusSchema = z.enum(["migrated", "copied", "deferred", "warning", "failed", "missing"]);
export type MigrationStatus = z.infer<typeof migrationStatusSchema>;

export interface LegacyRootInfo {
  kind: "appData" | "localData" | "deathsDir";
  path: string;
  exists: boolean;
}

export interface LegacySourceInfo {
  id: LegacySourceId;
  label: string;
  /** Relative file/dir under the legacy root this source maps to ("deferred" rows name their stage). */
  location: string;
  exists: boolean;
  bytes: number | null;
}

export const migrationRowSchema = z.object({
  source: z.string(),
  status: migrationStatusSchema,
  detail: z.string().nullable().optional(),
  warnings: z.array(z.string()).optional(),
});
export type MigrationRow = z.infer<typeof migrationRowSchema>;
