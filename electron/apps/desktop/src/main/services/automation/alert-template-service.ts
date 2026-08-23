/**
 * AlertTemplateService — port of Services/AlertTemplateService.cs: per-culture custom alert
 * template overrides persisted as { culture: { key: template } }, read BEFORE resource fallback,
 * with a malformed-placeholder fallback chain. The file format is preserved byte-for-byte so
 * existing custom_alerts.json overrides keep working.
 */
import { existsSync, mkdirSync, readFileSync, writeFileSync, renameSync } from "node:fs";
import { dirname } from "node:path";

export interface AlertTemplateDeps {
  filePath: string;
  culture?: string;
  /** Resource-manager equivalent: default translations per key. */
  defaults: (key: string) => string | null;
  log?: (level: "warn" | "error", message: string) => void;
}

type Overrides = Record<string, Record<string, string>>;

/** .NET string.Format subset: literal {{ / }} escapes and {n} argument slots. */
export function dotNetFormat(template: string, args: readonly unknown[]): string {
  return template.replace(/\{\{|\}\}|\{(\d+)\}/g, (match, slot?: string) => {
    if (match === "{{") return "{";
    if (match === "}}") return "}";
    const idx = Number(slot);
    if (idx >= args.length) throw new FormatError(`missing argument ${idx}`);
    const v = args[idx];
    if (v === null || v === undefined) throw new FormatError(`null argument ${idx}`);
    return String(v);
  });
}
class FormatError extends Error {}

export class AlertTemplateService {
  private overrides: Overrides = {};

  constructor(private readonly deps: AlertTemplateDeps) {
    this.load();
  }

  get culture(): string {
    return this.deps.culture ?? "en";
  }

  load(): void {
    try {
      if (!existsSync(this.deps.filePath)) return;
      const json = readFileSync(this.deps.filePath, "utf8");
      const data = JSON.parse(json) as Overrides;
      this.overrides = data ?? {};
    } catch (err) {
      // Fallback to empty on any corruption (parity), but surface it in the log.
      this.overrides = {};
      this.deps.log?.("warn", `custom alerts unreadable — starting empty: ${String(err)}`);
    }
  }

  save(): void {
    try {
      const dir = dirname(this.deps.filePath);
      if (!existsSync(dir)) mkdirSync(dir, { recursive: true });
      // Atomic write (tmp + rename) — same durability contract as the other stores.
      const tmp = `${this.deps.filePath}.tmp`;
      writeFileSync(tmp, JSON.stringify(this.overrides, null, 2), "utf8");
      renameSync(tmp, this.deps.filePath);
    } catch (err) {
      // Ignore write errors to prevent crashes (parity comment), but log them.
      this.deps.log?.("warn", `custom alerts save failed: ${String(err)}`);
    }
  }

  getAlertTemplate(key: string): string {
    const cultureOverrides = this.overrides[this.culture];
    const custom = cultureOverrides?.[key];
    if (custom !== undefined) return custom;
    return this.deps.defaults(key) ?? "";
  }

  getFormattedAlert(key: string, ...args: readonly unknown[]): string {
    const template = this.getAlertTemplate(key);
    if (template.length === 0) return "";
    try {
      return dotNetFormat(template, args);
    } catch {
      // Malformed custom template → retry with the default translation; if that also fails,
      // return the raw template (parity chain).
      const fallback = this.deps.defaults(key) ?? "";
      try {
        return dotNetFormat(fallback, args);
      } catch {
        return template;
      }
    }
  }

  setOverride(key: string, template: string): void {
    const cultureOverrides = (this.overrides[this.culture] ??= {});
    cultureOverrides[key] = template;
    this.save();
  }

  removeOverride(key: string): void {
    const cultureOverrides = this.overrides[this.culture];
    if (!cultureOverrides) return;
    delete cultureOverrides[key];
    if (Object.keys(cultureOverrides).length === 0) delete this.overrides[this.culture];
    this.save();
  }

  hasOverride(key: string): boolean {
    return this.overrides[this.culture]?.[key] !== undefined;
  }
}
