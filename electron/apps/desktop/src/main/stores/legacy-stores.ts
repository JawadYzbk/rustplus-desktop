/**
 * Small legacy stores over JsonStore (hotkeys, hotkey options, alert template overrides, tracked players).
 * Each mirrors the exact file shape the C# app wrote (PascalCase, same nesting); only the wrapper adds
 * schemaVersion + atomic writes. Consumers wire these in their feature stages; they exist and are tested now.
 */
import { join } from "node:path";
import { z } from "zod";
import {
  HOTKEY_OPTIONS_DEFAULTS,
  customAlertsSchema,
  deviceHotkeysSchema,
  hotkeyOptionsSchema,
  trackedPlayersFileSchema,
  type CustomAlerts,
  type DeviceHotkeys,
  type HotkeyOptions,
  type TrackedPlayerModel,
} from "@rpd/shared";
import { JsonStore } from "./json-store.js";

const V = 1;
type Log = ((level: "warn" | "error", message: string) => void) | undefined;

export class DeviceHotkeysStore {
  private readonly json: JsonStore<{ schemaVersion: number } & DeviceHotkeys>;
  constructor(userDataDir: string, log?: Log) {
    this.json = new JsonStore({
      file: join(userDataDir, "hotkeys.json"),
      schemaVersion: V,
      validate: (d): d is { schemaVersion: number } & DeviceHotkeys =>
        isObj(d) && d["schemaVersion"] === V && deviceHotkeysSchema.safeParse(withoutVersion(d)).success,
      log,
    });
  }

  all(): DeviceHotkeys {
    const o = this.json.load();
    return o.status === "loaded" ? withoutVersion(o.doc) : {};
  }

  setForServer(serverKey: string, gestures: Record<string, number[]>): DeviceHotkeys {
    const next = { ...this.all(), [serverKey]: gestures };
    this.json.save(Object.assign({ schemaVersion: V }, next));
    return next;
  }

  removeServer(serverKey: string): DeviceHotkeys {
    const { [serverKey]: _removed, ...rest } = this.all();
    void _removed;
    this.json.save(Object.assign({ schemaVersion: V }, rest));
    return rest;
  }
}

export class HotkeyOptionsStore {
  private readonly json: JsonStore<{ schemaVersion: number } & HotkeyOptions>;
  constructor(userDataDir: string, log?: Log) {
    this.json = new JsonStore({
      file: join(userDataDir, "hotkey_options.json"),
      schemaVersion: V,
      validate: (d): d is { schemaVersion: number } & HotkeyOptions =>
        isObj(d) && d["schemaVersion"] === V && hotkeyOptionsSchema.safeParse(withoutVersion(d)).success,
      log,
    });
  }

  get(): HotkeyOptions {
    const o = this.json.load();
    return o.status === "loaded" ? withoutVersion(o.doc) : { ...HOTKEY_OPTIONS_DEFAULTS };
  }

  set(patch: Partial<HotkeyOptions>): HotkeyOptions {
    const next = { ...this.get(), ...patch };
    const parsed = hotkeyOptionsSchema.safeParse(next);
    if (!parsed.success) throw new Error(`hotkey options breach contract: ${parsed.error.issues[0]?.message}`);
    this.json.save({ schemaVersion: V, ...parsed.data });
    return parsed.data;
  }
}

export class AlertTemplateStore {
  private readonly json: JsonStore<{ schemaVersion: number } & CustomAlerts>;
  constructor(userDataDir: string, log?: Log) {
    this.json = new JsonStore({
      file: join(userDataDir, "custom_alerts.json"),
      schemaVersion: V,
      validate: (d): d is { schemaVersion: number } & CustomAlerts =>
        isObj(d) && d["schemaVersion"] === V && customAlertsSchema.safeParse(withoutVersion(d)).success,
      log,
    });
  }

  /** Culture keys are preserved verbatim on disk; lookups are case-insensitive (legacy OrdinalIgnoreCase). */
  all(): CustomAlerts {
    const o = this.json.load();
    return o.status === "loaded" ? withoutVersion(o.doc) : {};
  }

  override(culture: string, key: string, template: string): CustomAlerts {
    const current = this.all();
    const existingCulture =
      Object.keys(current).find((k) => k.toLowerCase() === culture.toLowerCase()) ?? culture;
    const next: CustomAlerts = { ...current, [existingCulture]: { ...current[existingCulture], [key]: template } };
    this.json.save(Object.assign({ schemaVersion: V }, next));
    return next;
  }
}

export class TrackedPlayersStore {
  private readonly json: JsonStore<z.infer<typeof trackedPlayersFileSchema>>;
  constructor(userDataDir: string, log?: Log) {
    this.json = new JsonStore({
      file: join(userDataDir, "tracked_players.json"),
      schemaVersion: V,
      validate: (d): d is z.infer<typeof trackedPlayersFileSchema> => trackedPlayersFileSchema.safeParse(d).success,
      log,
    });
  }

  list(): TrackedPlayerModel[] {
    const o = this.json.load();
    return o.status === "loaded" ? o.doc.players : [];
  }

  replace(players: TrackedPlayerModel[]): TrackedPlayerModel[] {
    const parsed = trackedPlayersFileSchema.safeParse({ schemaVersion: V, players });
    if (!parsed.success) throw new Error(`tracked players breach contract: ${parsed.error.issues[0]?.message}`);
    this.json.save(parsed.data);
    return players;
  }
}

function withoutVersion<T extends Record<string, unknown>>(doc: T): Omit<T, "schemaVersion"> {
  const { schemaVersion: _v, ...rest } = doc;
  void _v;
  return rest;
}

function isObj(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}
