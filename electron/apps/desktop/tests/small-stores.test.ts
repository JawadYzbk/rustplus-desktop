/**
 * Small legacy store tests — real fs, PascalCase byte-compatibility with C#-written files.
 */
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { mkdtempSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import {
  HOTKEY_OPTIONS_DEFAULTS,
  TUTORIAL_PREFERENCES_DEFAULTS,
  TUTORIAL_STATUS,
  findAlertOverride,
} from "@rpd/shared";
import {
  AlertTemplateStore,
  DeviceHotkeysStore,
  HotkeyOptionsStore,
  TrackedPlayersStore,
} from "../src/main/stores/legacy-stores.js";
import { TutorialProgressStore } from "../src/main/stores/tutorial-progress-store.js";

let dir = "";

beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), "rpd-small-stores-"));
});

afterEach(() => {
  rmSync(dir, { recursive: true, force: true });
});

describe("DeviceHotkeysStore", () => {
  it("round-trips serverKey → gesture → entityIds and removes servers", () => {
    const store = new DeviceHotkeysStore(dir);
    store.setForServer("1.2.3.4:28082|42", { "Ctrl+F1": [12345, 67890], Toggle: [111] });
    store.setForServer("5.6.7.8:1|43", { "Alt+X": [9] });

    const reloaded = new DeviceHotkeysStore(dir).all();
    expect(reloaded["1.2.3.4:28082|42"]).toEqual({ "Ctrl+F1": [12345, 67890], Toggle: [111] });

    const afterRemove = new DeviceHotkeysStore(dir).removeServer("1.2.3.4:28082|42");
    expect(Object.keys(afterRemove)).toEqual(["5.6.7.8:1|43"]);
  });
});

describe("HotkeyOptionsStore", () => {
  it("defaults to ParallelMode=false / ToggleDelayMs=150 (audit §1)", () => {
    expect(new HotkeyOptionsStore(dir).get()).toEqual(HOTKEY_OPTIONS_DEFAULTS);
  });

  it("persists patches and rejects unknown keys", () => {
    const store = new HotkeyOptionsStore(dir);
    store.set({ ParallelMode: true, ToggleDelayMs: 250 });
    expect(new HotkeyOptionsStore(dir).get()).toEqual({ ParallelMode: true, ToggleDelayMs: 250 });
    expect(() => store.set({ Nope: 1 } as never)).toThrowError(/breach contract/);
  });
});

describe("AlertTemplateStore", () => {
  it("stores overrides per culture; lookups are case-insensitive while disk keys stay verbatim", () => {
    const store = new AlertTemplateStore(dir);
    store.override("de-DE", "CargoSpawned", "Cargo ist da!");
    store.override("EN-US", "HeliSpawned", "Heli up");

    // Disk preserves exactly what the caller wrote (legacy Dictionary behavior).
    const raw = JSON.parse(readFileSync(join(dir, "custom_alerts.json"), "utf8")) as Record<string, unknown>;
    expect(Object.keys(raw).filter((k) => k !== "schemaVersion").sort()).toEqual(["EN-US", "de-DE"]);

    const reloaded = new AlertTemplateStore(dir).all();
    expect(findAlertOverride(reloaded, "DE-de", "CargoSpawned")).toBe("Cargo ist da!");
    expect(findAlertOverride(reloaded, "en-us", "HeliSpawned")).toBe("Heli up");
    expect(findAlertOverride(reloaded, "fr-FR", "CargoSpawned")).toBeUndefined();
  });
});

describe("TrackedPlayersStore", () => {
  it("round-trips players incl. TimeSpan-string sessions", () => {
    const store = new TrackedPlayersStore(dir);
    const players = [
      {
        BMId: "bm-1",
        Name: "Neelo",
        LastServerName: "US Main",
        GroupName: "Team",
        GroupColor: "#60CDFF",
        Sessions: [
          {
            Name: "Neelo",
            BMId: "bm-1",
            SessionStartTimeUtc: "2026-08-23T10:00:00Z",
            Duration: "01:30:00",
            IsTracked: true,
          },
        ],
        IsBMOnly: false,
      },
    ];
    store.replace(players);
    const listed = new TrackedPlayersStore(dir).list();
    expect(listed).toHaveLength(1);
    expect(listed[0]!.Sessions[0]!.Duration).toBe("01:30:00");
  });
});

describe("TutorialProgressStore", () => {
  it("fresh read returns NotStarted at the definition's version", () => {
    const p = new TutorialProgressStore(dir).get({ id: "basics", version: 3 });
    expect(p.Status).toBe(TUTORIAL_STATUS.NotStarted);
    expect(p.TutorialVersion).toBe(3);
  });

  it("Completed tutorial on older definition version flips to Updated (legacy GetAsync parity)", () => {
    const store = new TutorialProgressStore(dir);
    store.save({
      TutorialId: "basics",
      TutorialVersion: 2,
      Status: TUTORIAL_STATUS.Completed,
      CompletedStepIds: ["s1"],
    });

    const bumped = new TutorialProgressStore(dir).get({ id: "basics", version: 3 });
    expect(bumped.Status).toBe(TUTORIAL_STATUS.Updated);
    expect(bumped.TutorialVersion).toBe(3);

    // Legacy parity: GetAsync's bump is in-memory only — disk still holds Completed/v2, so reading
    // again WITHOUT saving yields Updated once more (the re-offer keeps firing until acted on).
    expect(new TutorialProgressStore(dir).get({ id: "basics", version: 3 }).Status).toBe(TUTORIAL_STATUS.Updated);

    // Acting on it (user completes again at v3 → save) persists; afterwards the read stays Completed.
    new TutorialProgressStore(dir).save({
      TutorialId: "basics",
      TutorialVersion: 3,
      Status: TUTORIAL_STATUS.Completed,
      CompletedStepIds: ["s1"],
    });
    const settled = new TutorialProgressStore(dir).get({ id: "basics", version: 3 });
    expect(settled.Status).toBe(TUTORIAL_STATUS.Completed);
    expect(settled.TutorialVersion).toBe(3);
  });

  it("Skipped also flips to Updated on version bump", () => {
    const store = new TutorialProgressStore(dir);
    store.save({ TutorialId: "maps", TutorialVersion: 1, Status: TUTORIAL_STATUS.Skipped, CompletedStepIds: [] });
    expect(new TutorialProgressStore(dir).get({ id: "maps", version: 2 }).Status).toBe(TUTORIAL_STATUS.Updated);
  });

  it("reset / resetAll / preferences behave like the C# store", () => {
    const store = new TutorialProgressStore(dir);
    store.save({ TutorialId: "a", TutorialVersion: 1, Status: TUTORIAL_STATUS.InProgress, CompletedStepIds: [] });
    store.save({ TutorialId: "b", TutorialVersion: 1, Status: TUTORIAL_STATUS.InProgress, CompletedStepIds: [] });
    store.reset("a");
    expect(Object.keys(store.all())).toEqual(["b"]);
    store.resetAll();
    expect(store.all()).toEqual({});

    store.savePreferences({ ...TUTORIAL_PREFERENCES_DEFAULTS, FirstRunPromptDismissed: true, LastTutorialId: "a" });
    const prefs = new TutorialProgressStore(dir).preferences();
    expect(prefs.FirstRunPromptDismissed).toBe(true);
    expect(prefs.AutoStartBasicTutorial).toBe(true); // default preserved
    expect(prefs.LastTutorialId).toBe("a");
  });
});
