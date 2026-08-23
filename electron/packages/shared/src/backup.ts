/**
 * Backup/restore (M-format v2) + granular reset contract.
 *
 * Crypto upgrade vs legacy (audit DATA_STORES §4, decision logged): AES-256-GCM (authenticated) with
 * PBKDF2-SHA256 @ 210k iterations replaces AES-256-CBC/no-MAC @ 10k. Legacy .zip backups are NOT read
 * by this format; reading them is pending owner answer to audit open question §8.3.
 *
 * File layout (encrypted): "RPDENC2" | u8 version(=2) | 16B salt | 12B iv | 16B gcm tag | zip bytes.
 * Unencrypted: a plain zip (starts with PK).
 */
import { z } from "zod";
import { defineChannel } from "./ipc/framework.js";

export const BACKUP_MAGIC = "RPDENC2";
export const PBKDF2_ITERATIONS = 210_000;

export const backupCreate = defineChannel(
  "backup/create",
  z.object({
    /** Null/undefined → plain unencrypted zip. */
    password: z.string().min(1).nullable().optional(),
  }),
  z.object({
    path: z.string(),
    bytes: z.number().int(),
    encrypted: z.boolean(),
  }),
  "Create a backup zip of all user data stores in the new root.",
);

export const backupRestore = defineChannel(
  "backup/restore",
  z.object({
    path: z.string(),
    password: z.string().nullable().optional(),
  }),
  z.object({
    restored: z.array(z.string()),
    skipped: z.array(z.string()),
  }),
  "Restore a backup zip (v2 format, optionally encrypted) into the new root.",
);

export const resetTargetsSchema = z.array(z.enum(["profiles", "steam", "pairing", "crosshairs", "cache", "three_d"])).min(1);
export type ResetTarget = z.infer<typeof resetTargetsSchema>[number];

export const resetPerform = defineChannel(
  "reset/perform",
  z.object({ targets: resetTargetsSchema }),
  z.object({ performed: z.array(z.string()) }),
  "Granular data reset (legacy Connection.Reset semantics; Connection itself is runtime state, stage 4).",
);
