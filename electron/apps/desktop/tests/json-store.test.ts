/**
 * JsonStore behavior tests — real filesystem via mkdtemp (no mocks): atomicity and rename semantics
 * are exactly what we're testing, so they must run against actual NTFS behavior.
 */
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { mkdtempSync, existsSync, writeFileSync, readFileSync, readdirSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { JsonStore } from "../src/main/stores/json-store.js";

interface PrefsDoc {
  schemaVersion: number;
  volume: number;
  theme: string;
}

const CURRENT = 2;

function validate(doc: unknown): doc is PrefsDoc {
  if (typeof doc !== "object" || doc === null) return false;
  const d = doc as Record<string, unknown>;
  return (
    d["schemaVersion"] === CURRENT &&
    typeof d["volume"] === "number" &&
    typeof d["theme"] === "string"
  );
}

function makeStore(migrate?: (raw: Record<string, unknown>, from: number) => PrefsDoc | null): JsonStore<PrefsDoc> {
  return new JsonStore<PrefsDoc>({ file: filePath, schemaVersion: CURRENT, validate, migrate });
}

let dir = "";
let filePath = "";

beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), "rpd-store-"));
  filePath = join(dir, "prefs.json");
});

afterEach(() => {
  rmSync(dir, { recursive: true, force: true });
});

describe("JsonStore", () => {
  it("round-trips save→load preserving document content", () => {
    const store = makeStore();
    const doc: PrefsDoc = { schemaVersion: CURRENT, volume: 0.8, theme: "dark" };
    store.save(doc);
    expect(store.load()).toEqual({ status: "loaded", doc });
    // Atomicity: no temp leftovers after save.
    expect(existsSync(`${filePath}.tmp`)).toBe(false);
  });

  it("reports missing for a nonexistent file", () => {
    expect(makeStore().load()).toEqual({ status: "missing" });
  });

  it("quarantines invalid JSON and keeps the bytes on disk", () => {
    writeFileSync(filePath, "{ this is not json", "utf8");
    const result = makeStore().load();
    expect(result.status).toBe("quarantined");
    if (result.status === "quarantined") {
      expect(result.quarantinePath).toMatch(/corrupt/);
      expect(existsSync(result.quarantinePath)).toBe(true);
      expect(readFileSync(result.quarantinePath, "utf8")).toContain("not json");
    }
    expect(existsSync(filePath)).toBe(false); // moved, not deleted
  });

  it("quarantines a current-version document failing validation", () => {
    writeFileSync(filePath, JSON.stringify({ schemaVersion: CURRENT, volume: "loud", theme: 3 }), "utf8");
    expect(makeStore().load().status).toBe("quarantined");
  });

  it("quarantines future-version documents (downgrade protection)", () => {
    writeFileSync(filePath, JSON.stringify({ schemaVersion: 99, volume: 1, theme: "x" }), "utf8");
    const result = makeStore().load();
    expect(result.status).toBe("quarantined");
    if (result.status === "quarantined") expect(result.reason).toContain("future");
  });

  it("migrates an older version through the migrate hook", () => {
    const store = makeStore((raw) => ({
      schemaVersion: CURRENT,
      volume: typeof raw["volume"] === "number" ? raw["volume"] : 1,
      theme: typeof raw["theme"] === "string" ? raw["theme"] : "dark",
    }));
    writeFileSync(filePath, JSON.stringify({ schemaVersion: 1, volume: 0.5 }), "utf8");
    expect(store.load()).toEqual({
      status: "loaded",
      doc: { schemaVersion: CURRENT, volume: 0.5, theme: "dark" },
    });
  });

  it("treats an unversioned legacy file as version 0 for migration", () => {
    const seen: number[] = [];
    const store = makeStore((_raw, from) => {
      seen.push(from);
      return { schemaVersion: CURRENT, volume: 1, theme: "dark" };
    });
    writeFileSync(filePath, JSON.stringify({ volume: 1, theme: "dark" }), "utf8"); // no schemaVersion
    const result = store.load();
    expect(seen).toEqual([0]);
    expect(result.status).toBe("loaded");
  });

  it("keeps the original file untouched when migration fails (non-destructive)", () => {
    const original = JSON.stringify({ schemaVersion: 1, volume: 0.25, note: "keep me" });
    writeFileSync(filePath, original, "utf8");
    const result = makeStore(() => null).load();
    expect(result.status).toBe("migrate-failed");
    expect(existsSync(filePath)).toBe(true);
    expect(readFileSync(filePath, "utf8")).toBe(original);
    expect(existsSync(join(dir, "corrupt"))).toBe(false);
  });

  it("overwrites cleanly on repeated saves without temp accumulation", () => {
    const store = makeStore();
    for (let i = 0; i < 5; i++) {
      store.save({ schemaVersion: CURRENT, volume: i / 10, theme: "dark" });
    }
    const result = store.load();
    expect(result.status).toBe("loaded");
    if (result.status === "loaded") expect(result.doc.volume).toBe(0.4);
    expect(readdirSync(dir).filter((f) => f.endsWith(".tmp"))).toHaveLength(0);
  });
});
