import { existsSync, readFileSync } from "node:fs";
import { join } from "node:path";

export interface RecyclerYield {
  shortName: string;
  quantity: number;
  probability: number;
}

export interface RecyclerNode {
  yields: RecyclerYield[];
}

export interface RecyclerItem {
  id: string;
  shortName: string;
  displayName: string;
  category: string;
  stackSize: number;
  wild: RecyclerNode | null;
  safe: RecyclerNode | null;
  recycleInfo: RecycleInfo[];
}

interface RecycleOutput {
  shortName: string;
  amount: number;
}

interface RecycleInfo {
  recyclerId: string;
  guaranteed: RecycleOutput[];
  percentage: RecycleOutput[];
}

export interface RecyclerMetric {
  expected: number;
  guaranteed: number;
  chance: number;
  chancePercent: number;
  min: number;
  max: number;
}

export interface RecyclerOutput {
  shortName: string;
  displayName: string;
  wild: RecyclerMetric;
  safe: RecyclerMetric;
}

export interface RecyclerCalculation {
  outputs: RecyclerOutput[];
  wildSeconds: number;
  safeSeconds: number;
}

interface RawRecyclerItem {
  id?: unknown;
  shortName?: unknown;
  displayName?: unknown;
  category?: unknown;
  stackSize?: unknown;
  canBeRecycled?: unknown;
  recycleInfo?: unknown;
}

const knownOutputNames: Record<string, string> = {
  scrap: "Scrap",
  "metal.refined": "High Quality Metal",
  "metal.fragments": "Metal Fragments",
  cloth: "Cloth",
  rope: "Rope",
  techparts: "Tech Parts",
  wood: "Wood",
  stones: "Stones",
  sulfur: "Sulfur",
  gunpowder: "Gunpowder",
  leather: "Leather",
  "fat.animal": "Animal Fat",
  lowgradefuel: "Low Grade Fuel",
  "bone.fragments": "Bone Fragments",
  charcoal: "Charcoal",
  "crude.oil": "Crude Oil",
  riflebody: "Rifle Body",
  semibody: "Semi Body",
  smgbody: "SMG Body",
  metalpipe: "Metal Pipe",
  metalspring: "Metal Spring",
  gears: "Gears",
  metalblade: "Metal Blade",
  roadsigns: "Road Signs",
  sheetmetal: "Sheet Metal",
  sewingkit: "Sewing Kit",
  tarp: "Tarp",
  propanetank: "Propane Tank",
  "cctv.camera": "CCTV Camera",
  "targeting.computer": "Targeting Computer",
  fuse: "Fuse",
};

const asString = (value: unknown): string => (typeof value === "string" ? value : "");
const asNumber = (value: unknown, fallback = 0): number => (typeof value === "number" && Number.isFinite(value) ? value : fallback);
const asObject = (value: unknown): Record<string, unknown> | null =>
  value && typeof value === "object" && !Array.isArray(value) ? (value as Record<string, unknown>) : null;

function findDataPath(fileName: string): string {
  const candidates = [
    join(process.cwd(), "assets", "data", fileName),
    join(__dirname, "..", "..", "assets", "data", fileName),
    join(__dirname, "..", "..", "..", "..", "assets", "data", fileName),
  ];
  const path = candidates.find((candidate) => existsSync(candidate));
  if (!path) throw new Error(`The packaged ${fileName} asset is missing.`);
  return path;
}

function parseNode(value: unknown): RecyclerNode | null {
  const node = asObject(value);
  if (!node || !Array.isArray(node["yield"])) return null;
  return {
    yields: node["yield"].flatMap((entry): RecyclerYield[] => {
      const item = asObject(entry);
      if (!item) return [];
      const shortName = asString(item["shortname"]);
      const quantity = asNumber(item["quantity"]);
      const probability = asNumber(item["probability"]);
      return shortName && quantity >= 0 && probability >= 0 ? [{ shortName, quantity, probability }] : [];
    }),
  };
}

function parseRecycleInfo(value: unknown): RecycleInfo[] {
  if (!Array.isArray(value)) return [];
  return value.flatMap((entry): RecycleInfo[] => {
    const info = asObject(entry);
    if (!info) return [];
    const parseOutputs = (raw: unknown): RecycleOutput[] =>
      Array.isArray(raw)
        ? raw.flatMap((output): RecycleOutput[] => {
            const item = asObject(output);
            const shortName = asString(item?.["itemId"]);
            const amount = asNumber(item?.["amount"]);
            return shortName && amount >= 0 ? [{ shortName, amount }] : [];
          })
        : [];
    return [{
      recyclerId: asString(info["recyclerId"]),
      guaranteed: parseOutputs(info["guaranteedOutput"]),
      percentage: parseOutputs(info["percentageBasedOutput"]),
    }];
  });
}

export function loadRecyclerItems(): RecyclerItem[] {
  const rawItems = JSON.parse(readFileSync(findDataPath("recycler-items.json"), "utf8")) as unknown;
  const rawRecycling = JSON.parse(readFileSync(findDataPath("Recycling-Data.json"), "utf8")) as unknown;
  const recycling = asObject(rawRecycling) ?? {};
  if (!Array.isArray(rawItems)) throw new Error("recycler-items.json must contain an array.");

  return rawItems.flatMap((raw): RecyclerItem[] => {
    const item = asObject(raw) as RawRecyclerItem | null;
    if (!item || item["canBeRecycled"] !== true) return [];
    const shortName = asString(item["shortName"]);
    if (!shortName) return [];
    const override = asObject(recycling[shortName]);
    return [{
      id: asString(item["id"]) || shortName,
      shortName,
      displayName: asString(item["displayName"]) || humanize(shortName),
      category: asString(item["category"]) || "Other",
      stackSize: Math.max(1, Math.floor(asNumber(override?.["stackSize"], asNumber(item["stackSize"], 1)))),
      wild: parseNode(override?.["recycler"]),
      safe: parseNode(override?.["safe-zone-recycler"]),
      recycleInfo: parseRecycleInfo(item["recycleInfo"]),
    }];
  }).sort((a, b) => a.displayName.localeCompare(b.displayName));
}

export function calculateRecycler(
  items: readonly RecyclerItem[],
  quantities: Readonly<Record<string, number>>,
): RecyclerCalculation {
  const wild = new Map<string, { expected: number; min: number; max: number }>();
  const safe = new Map<string, { expected: number; min: number; max: number }>();
  let wildSeconds = 0;
  let safeSeconds = 0;

  const add = (map: Map<string, { expected: number; min: number; max: number }>, shortName: string, expected: number, min: number, max: number): void => {
    const current = map.get(shortName) ?? { expected: 0, min: 0, max: 0 };
    map.set(shortName, { expected: current.expected + expected, min: current.min + min, max: current.max + max });
  };

  const processNode = (node: RecyclerNode | null, map: Map<string, { expected: number; min: number; max: number }>, quantity: number): void => {
    for (const output of node?.yields ?? []) {
      add(map, output.shortName, quantity * output.quantity * output.probability, output.probability >= 1 ? quantity * output.quantity : 0, quantity * output.quantity);
    }
  };

  for (const item of items) {
    const quantity = Math.max(0, Math.floor(quantities[item.shortName] ?? 0));
    if (quantity === 0) continue;
    const unitsPerTick = Math.max(1, Math.ceil(item.stackSize * 0.1));
    const ticks = Math.ceil(quantity / unitsPerTick);
    wildSeconds += ticks * 5;
    safeSeconds += ticks * 8;
    if (item.wild || item.safe) {
      processNode(item.wild, wild, quantity);
      processNode(item.safe, safe, quantity);
      continue;
    }
    for (const info of item.recycleInfo) {
      const map = info.recyclerId === "recycler-radtown" ? wild : info.recyclerId === "recycler-safezone" ? safe : null;
      if (!map) continue;
      for (const output of info.guaranteed) add(map, output.shortName, quantity * output.amount, quantity * output.amount, quantity * output.amount);
      for (const output of info.percentage) add(map, output.shortName, quantity * output.amount / 100, 0, quantity);
    }
  }

  const names = new Set([...wild.keys(), ...safe.keys()]);
  const metric = (value: { expected: number; min: number; max: number } | undefined): RecyclerMetric => {
    const current = value ?? { expected: 0, min: 0, max: 0 };
    const chance = Math.max(0, current.max - current.min);
    return {
      ...current,
      guaranteed: current.min,
      chance,
      chancePercent: chance > 0 ? ((current.expected - current.min) / chance) * 100 : 0,
    };
  };
  const outputs = [...names]
    .map((shortName) => ({ shortName, displayName: knownOutputNames[shortName] ?? humanize(shortName), wild: metric(wild.get(shortName)), safe: metric(safe.get(shortName)) }))
    .filter((output) => output.wild.expected > 0 || output.safe.expected > 0)
    .sort((a, b) => a.displayName.localeCompare(b.displayName));
  return { outputs, wildSeconds, safeSeconds };
}

function humanize(shortName: string): string {
  return shortName
    .replace(/[._]/g, " ")
    .replace(/\b\w/g, (letter) => letter.toUpperCase());
}
