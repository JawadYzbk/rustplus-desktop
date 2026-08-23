/**
 * Server profiles store — profiles.json parity, lossless-first.
 *
 * The legacy file is a bare JSON array of ServerProfile objects (plaintext PlayerToken). Full typed
 * modeling of every nested structure (SmartDevice tree, LogicRule, DeviceAutomationRule, CustomTimer,
 * DeathMarkerData…) lands with the devices/automation stages. Until then this store is DELIBERATELY
 * schema-loose per profile: it preserves every unknown field byte-faithfully on rewrite (zero data loss),
 * while providing typed accessors for the subset already needed and sealing PlayerToken at rest.
 *
 * New format: { "schemaVersion": 1, "profiles": [ …profile objects… ] } — the migrator converts legacy
 * arrays; reading a legacy array directly is supported for tests/migrator tooling.
 */
import { join } from "node:path";
import { z } from "zod";
import { JsonStore } from "./json-store.js";
import type { SecretCodec } from "./secret-codec.js";

const SCHEMA_VERSION = 1;

/** Typed subset — grows as stages port the nested models. Everything else passes through untouched. */
export const profileCoreSchema = z.object({
  Name: z.string(),
  Description: z.string().optional(),
  Host: z.string(),
  Port: z.number().int(),
  SteamId64: z.string().optional().default(""),
  BattleMetricsId: z.string().nullable().optional(),
  UseFacepunchProxy: z.boolean().optional(),
});

export type ProfileCore = z.infer<typeof profileCoreSchema>;

/** Loose profile record: core fields validated, everything else opaque but preserved. */
const looseProfileSchema = profileCoreSchema.passthrough();
type LooseProfile = z.infer<typeof looseProfileSchema>;

const profilesFileSchema = z.object({
  schemaVersion: z.literal(SCHEMA_VERSION),
  profiles: z.array(z.record(z.unknown())),
});

interface StoredDoc {
  schemaVersion: typeof SCHEMA_VERSION;
  profiles: Record<string, unknown>[];
}

export class ProfilesStore {
  private readonly json: JsonStore<StoredDoc>;
  private cache: LooseProfile[] | null = null;

  constructor(
    userDataDir: string,
    private readonly secrets: SecretCodec,
    log?: (level: "warn" | "error", message: string) => void,
  ) {
    this.json = new JsonStore<StoredDoc>({
      file: join(userDataDir, "profiles.json"),
      schemaVersion: SCHEMA_VERSION,
      validate: (doc): doc is StoredDoc =>
        profilesFileSchema.safeParse(doc).success &&
        (doc as StoredDoc).profiles.every((p) => looseProfileSchema.safeParse(p).success),
      log,
    });
  }

  /** All profiles with PlayerToken unsealed in memory only. */
  list(): ProfileCore[] {
    return this.load().map((p) => coreOf(p));
  }

  /** Stable identity parity with C# MatchKey ("host:port|steamId"); token deliberately excluded. */
  matchKey(p: ProfileCore): string {
    return `${p.Host}:${p.Port}|${p.SteamId64}`;
  }

  upsert(profile: ProfileCore & Record<string, unknown>): void {
    const profiles = this.load();
    const key = this.matchKey(profile);
    const idx = profiles.findIndex((p) => this.matchKey(coreOf(p)) === key);

    // Cache holds PLAINTEXT tokens; sealing happens uniformly in persist(). Never pre-seal here
    // or a re-upsert would double-wrap the blob.
    if (idx >= 0) {
      const merged: Record<string, unknown> = { ...profiles[idx], ...profile };
      const parsed = looseProfileSchema.safeParse(merged);
      if (!parsed.success) throw new Error(`profile update breached contract: ${parsed.error.issues[0]?.message}`);
      profiles[idx] = parsed.data as LooseProfile;
    } else {
      const parsed = looseProfileSchema.safeParse(profile);
      if (!parsed.success) throw new Error(`profile insert breached contract: ${parsed.error.issues[0]?.message}`);
      profiles.unshift(parsed.data as LooseProfile); // newest first, matching legacy list behavior
    }
    this.persist(profiles);
  }

  removeByMatchKey(key: string): boolean {
    const profiles = this.load();
    const next = profiles.filter((p) => this.matchKey(coreOf(p)) !== key);
    if (next.length === profiles.length) return false;
    this.persist(next);
    return true;
  }

  /** Unseals PlayerToken for connection use; returns "" when unset. */
  tokenFor(key: string): string {
    const p = this.load().find((c) => this.matchKey(coreOf(c)) === key);
    if (!p) return "";
    const raw = p["PlayerToken"];
    if (typeof raw !== "string" || raw === "") return "";
    return this.secrets.open(raw);
  }

  /** Raw Devices array of one profile (opaque records — callers parse with server-profile.ts). */
  devicesFor(key: string): Record<string, unknown>[] | null {
    const p = this.load().find((c) => this.matchKey(coreOf(c)) === key);
    if (!p) return null;
    const raw = p["Devices"];
    return Array.isArray(raw) ? (raw as Record<string, unknown>[]) : [];
  }

  /** Whole-tree write for one profile (legacy Save() parity); false when the profile is gone. */
  saveDevices(key: string, devices: Record<string, unknown>[]): boolean {
    const profiles = this.load();
    const p = profiles.find((c) => this.matchKey(coreOf(c)) === key);
    if (!p) return false;
    p["Devices"] = devices;
    this.persist(profiles);
    return true;
  }

  /** Opaque field read (LogicRules, CustomTimers, flags…). Undefined when profile/field absent. */
  field(key: string, name: string): unknown {
    return this.load().find((c) => this.matchKey(coreOf(c)) === key)?.[name];
  }

  /** Opaque field write with whole-document persist; false when the profile is gone. */
  setField(key: string, name: string, value: unknown): boolean {
    const profiles = this.load();
    const p = profiles.find((c) => this.matchKey(coreOf(c)) === key);
    if (!p) return false;
    p[name] = value;
    this.persist(profiles);
    return true;
  }

  private load(): LooseProfile[] {
    if (this.cache) return this.cache;
    const outcome = this.json.load();
    let rawProfiles: Record<string, unknown>[];
    if (outcome.status === "loaded") {
      rawProfiles = outcome.doc.profiles;
    } else if (outcome.status === "missing") {
      rawProfiles = [];
    } else {
      throw new Error(
        outcome.status === "quarantined"
          ? `profiles.json quarantined: ${outcome.reason} (see ${outcome.quarantinePath})`
          : `profiles.json migration failed: ${outcome.reason}`,
      );
    }
    this.cache = rawProfiles.map((p) => unsealProfile(p, this.secrets));
    return this.cache;
  }

  private persist(profiles: LooseProfile[]): void {
    const sealedProfiles = profiles.map((p) => sealProfile(p, this.secrets));
    this.json.save({ schemaVersion: SCHEMA_VERSION, profiles: sealedProfiles });
    // Keep cache consistent with what's on disk (unsealed view).
    this.cache = profiles;
  }
}

function coreOf(p: LooseProfile): ProfileCore {
  return {
    Name: p.Name,
    Description: p.Description,
    Host: p.Host,
    Port: p.Port,
    SteamId64: p.SteamId64,
    BattleMetricsId: p.BattleMetricsId,
    UseFacepunchProxy: p.UseFacepunchProxy,
  };
}

function sealProfile(p: LooseProfile, codec: SecretCodec): Record<string, unknown> {
  const out: Record<string, unknown> = { ...p };
  const token = out["PlayerToken"];
  out["PlayerToken"] = codec.seal(typeof token === "string" ? token : "");
  return out;
}

function unsealProfile(p: Record<string, unknown>, codec: SecretCodec): LooseProfile {
  const out: Record<string, unknown> = { ...p };
  const token = out["PlayerToken"];
  if (typeof token === "string" && token !== "") {
    out["PlayerToken"] = codec.open(token);
  }
  const parsed = looseProfileSchema.safeParse(out);
  if (!parsed.success) {
    throw new Error(`stored profile breached contract: ${parsed.error.issues[0]?.message}`);
  }
  return parsed.data as LooseProfile;
}
