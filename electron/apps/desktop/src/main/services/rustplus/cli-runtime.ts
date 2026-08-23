/**
 * CLI runtime resolution — port of RuntimeHelper.cs (node.exe discovery, rustplus-cli entry
 * candidates) and PairingListenerRealProcess.FindChromiumBrowser (browser preference order for the
 * fcm-register login window). All filesystem/registry access is injected so the resolution logic
 * itself is golden-testable.
 */
import * as fs from "node:fs";
import * as path from "node:path";

type Exists = (p: string) => boolean;
const fileExists: Exists = (p) => {
  try {
    return fs.existsSync(p) && fs.statSync(p).isFile();
  } catch {
    return false;
  }
};

/**
 * FindBundledNode parity: probe each base dir for runtime/node-win-x64/node.exe, then
 * node-win-x64/node.exe, then a flat node.exe. `devWalkUp` mirrors the debug deep-search that
 * walks up at most 5 levels from the first base.
 */
export function findBundledNode(
  bases: string[],
  opts: { exists?: Exists; devWalkUp?: boolean } = {},
): string | null {
  const exists = opts.exists ?? fileExists;

  // Deduplicate case-insensitively (Windows), preserving order (HashSet+ordinal-ignore-case parity).
  const seen = new Set<string>();
  const unique: string[] = [];
  for (const b of bases) {
    if (!b) continue;
    const key = process.platform === "win32" ? b.toLowerCase() : b;
    if (!seen.has(key)) {
      seen.add(key);
      unique.push(b);
    }
  }

  for (const base of unique) {
    for (const rel of ["runtime/node-win-x64/node.exe", "node-win-x64/node.exe", "node.exe"]) {
      const p = path.join(base, ...rel.split("/"));
      if (exists(p)) return p;
    }
  }

  if (opts.devWalkUp !== false && unique.length > 0) {
    let cur = unique[0]!;
    for (let i = 0; i < 5; i++) {
      const pDev = path.join(cur, "runtime", "node-win-x64", "node.exe");
      if (exists(pDev)) return path.resolve(pDev);
      const next = path.dirname(cur);
      if (!next || next === cur) break;
      cur = next;
    }
  }
  return null;
}

/** ResolveCliEntry parity — first existing candidate wins; working dir is the root. */
export const CLI_ENTRY_CANDIDATES = [
  "cli.js",
  "rustplus.js",
  "index.js",
  path.join("node_modules", "@liamcottle", "rustplus.js", "cli", "index.js"),
] as const;

export function resolveCliEntry(root: string, exists: Exists = fileExists): string | null {
  for (const rel of CLI_ENTRY_CANDIDATES) {
    const p = path.join(root, rel);
    if (exists(p)) return p;
  }
  return null;
}

export interface BrowserCandidate {
  exe: string;
  name: string;
  relatives: string[];
}

/** Order is preference, not availability — legacy comment preserved verbatim in spirit. */
export const BROWSER_CANDIDATES: BrowserCandidate[] = [
  { exe: "chrome.exe", name: "Google Chrome", relatives: ["Google\\Chrome\\Application\\chrome.exe"] },
  { exe: "msedge.exe", name: "Microsoft Edge", relatives: ["Microsoft\\Edge\\Application\\msedge.exe"] },
  { exe: "brave.exe", name: "Brave", relatives: ["BraveSoftware\\Brave-Browser\\Application\\brave.exe"] },
  { exe: "vivaldi.exe", name: "Vivaldi", relatives: ["Vivaldi\\Application\\vivaldi.exe"] },
  { exe: "opera.exe", name: "Opera", relatives: ["Opera\\opera.exe"] },
  { exe: "chrome.exe", name: "Chromium", relatives: ["Chromium\\Application\\chrome.exe"] },
];

export interface FoundBrowser {
  path: string;
  name: string;
}

/**
 * FindChromiumBrowser parity: App Paths registry first (where Windows installers register
 * executables), then ProgramFiles/ProgramFiles(x86)/LocalAppData relative probes.
 */
export function findChromiumBrowser(
  roots: string[],
  opts: {
    exists?: Exists;
    registryLookup?: (exeName: string) => string | null;
    onlyThese?: string[];
  } = {},
): FoundBrowser | null {
  const exists = opts.exists ?? fileExists;
  const registryLookup = opts.registryLookup ?? (() => null);
  const only = opts.onlyThese ?? [];

  for (const cand of BROWSER_CANDIDATES) {
    if (only.length > 0 && !only.some((o) => o.toLowerCase() === cand.exe.toLowerCase())) continue;

    const fromRegistry = registryLookup(cand.exe);
    if (fromRegistry) return { path: fromRegistry, name: cand.name };

    for (const root of roots) {
      if (!root) continue;
      for (const relative of cand.relatives) {
        const full = path.join(root, ...relative.split("\\"));
        if (exists(full)) return { path: full, name: cand.name };
      }
    }
  }
  return null;
}

/** PUPPETEER_EXECUTABLE_PATH / CHROME_PATH env pair for one fcm-register attempt. */
export interface RegisterBrowserEnv {
  label: string;
  env: Record<string, string>;
}

export function browserRegisterEnvs(browser: FoundBrowser): RegisterBrowserEnv {
  return {
    label: browser.name,
    env: { PUPPETEER_EXECUTABLE_PATH: browser.path, CHROME_PATH: browser.path },
  };
}

/**
 * LookUpAppPath parity: reads HKCU/HKLM "App Paths", where Windows installers register
 * executables. Synchronous (pairing start is already an explicit user action); returns null on any
 * failure — filesystem probes remain as fallback.
 */
export function windowsRegistryAppPath(exeName: string): string | null {
  if (process.platform !== "win32") return null;
  const { execFileSync } = require("node:child_process") as typeof import("node:child_process");
  for (const hive of ["HKCU", "HKLM"]) {
    try {
      const out = execFileSync(
        "reg",
        ["query", `${hive}\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\${exeName}`, "/ve"],
        { encoding: "utf8", timeout: 2000, stdio: ["ignore", "pipe", "ignore"] },
      );
      const match = /REG_SZ\s+(.+)/.exec(out);
      if (!match) continue;
      const p = match[1]!.trim().replace(/^"|"$/g, "");
      if (p && fileExists(p)) return p;
    } catch {
      /* key missing or reg unavailable — try next hive */
    }
  }
  return null;
}
