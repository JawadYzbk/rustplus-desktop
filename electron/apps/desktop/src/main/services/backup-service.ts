/**
 * Backup/restore service — v2 format with upgraded crypto (decision log, stage 3).
 *
 * Create: collect user stores from the new root → zip (fflate) → manifest with per-file sha256 →
 *   optional AES-256-GCM encryption (PBKDF2-SHA256, 210k iters).
 * Restore: magic-detect encryption → decrypt+verify GCM tag → unzip → verify manifest hashes →
 *   copy into the new root. Never deletes unrelated files; JSON files are parse-checked before copy.
 */
import { createHash, createCipheriv, createDecipheriv, pbkdf2Sync, randomBytes } from "node:crypto";
import { existsSync, mkdirSync, readdirSync, readFileSync, writeFileSync } from "node:fs";
import { join, relative } from "node:path";
import { zipSync, unzipSync, strFromU8, strToU8 } from "fflate";
import { APP_VERSION, BACKUP_MAGIC, PBKDF2_ITERATIONS } from "@rpd/shared";

const BACKUP_FORMAT = 2;

/** Store files backed up (new-root-relative). Directories are walked (Overlays\ lands here at stage 6). */
const BACKUP_FILES = [
  "profiles.json",
  "tracking_settings.json",
  "hotkeys.json",
  "hotkey_options.json",
  "custom_alerts.json",
  "tracked_players.json",
  "tutorial-progress.json",
  "ui-prefs.json",
  "rustplusjs-config.json",
];

const BACKUP_DIRS = ["cache", "Overlays"];

interface Manifest {
  format: number;
  appVersion: string;
  createdAt: string;
  files: Array<{ path: string; sha256: string }>;
}

export interface BackupResult {
  path: string;
  bytes: number;
  encrypted: boolean;
}

export class BackupService {
  constructor(
    private readonly userDataDir: string,
    private readonly backupsDir: string,
    private readonly log?: (level: "info" | "warn" | "error", message: string) => void,
  ) {}

  create(password?: string | null): BackupResult {
    const files = this.collectFiles();

    const manifest: Manifest = {
      format: BACKUP_FORMAT,
      appVersion: APP_VERSION,
      createdAt: new Date().toISOString(),
      files: files.map((f) => ({ path: f.rel, sha256: sha256(f.data) })),
    };

    const zip = zipSync({
      ...Object.fromEntries(files.map((f) => [f.rel, f.data])),
      "manifest.json": strToU8(JSON.stringify(manifest, null, 2)),
    });

    const encrypted = Boolean(password);
    const payload = encrypted ? encrypt(zip, password as string) : zip;

    const stamp = new Date().toISOString().replace(/[:.]/g, "-");
    const name = `RustPlusDesk-Backup-${stamp}${encrypted ? ".enc" : ""}.zip`;
    const outPath = join(this.backupsDir, name);
    mkdirSync(this.backupsDir, { recursive: true });
    writeFileSync(outPath, payload);
    this.log?.("info", `backup created: ${name} (${payload.length} B, ${files.length} files, enc=${encrypted})`);
    return { path: outPath, bytes: payload.length, encrypted };
  }

  restore(backupPath: string, password?: string | null): { restored: string[]; skipped: string[] } {
    if (!existsSync(backupPath)) throw new Error(`backup file not found: ${backupPath}`);
    const raw = new Uint8Array(readFileSync(backupPath));

    let zipBytes: Uint8Array;
    if (isEncrypted(raw)) {
      if (!password) throw new Error("backup is encrypted but no password was provided");
      zipBytes = decrypt(raw, password);
    } else {
      if (password) throw new Error("password provided but backup is not encrypted");
      zipBytes = raw;
    }

    let entries: Record<string, Uint8Array>;
    try {
      entries = unzipSync(zipBytes);
    } catch (err) {
      throw new Error(`not a valid backup zip: ${err instanceof Error ? err.message : String(err)}`);
    }

    // Manifest first; tamper check via recorded hashes.
    const manifestEntry = entries["manifest.json"];
    if (!manifestEntry) throw new Error("manifest.json missing — not a v2 backup");
    const manifest = JSON.parse(strFromU8(manifestEntry)) as Manifest;
    if (manifest.format !== BACKUP_FORMAT) {
      throw new Error(`unsupported backup format version ${String(manifest.format)} (expected ${BACKUP_FORMAT})`);
    }
    for (const f of manifest.files) {
      const data = entries[f.path];
      if (!data) throw new Error(`backup is incomplete: ${f.path} missing`);
      if (sha256(data) !== f.sha256) throw new Error(`integrity check failed for ${f.path}`);
    }

    const restored: string[] = [];
    const skipped: string[] = [];
    for (const [rel, data] of Object.entries(entries)) {
      if (rel === "manifest.json") continue;
      if (rel.endsWith("/")) continue;
      if (rel.endsWith(".json")) {
        try {
          JSON.parse(strFromU8(data));
        } catch {
          skipped.push(`${rel} (invalid JSON)`);
          continue;
        }
      }
      const dest = join(this.userDataDir, rel);
      mkdirSync(join(dest, ".."), { recursive: true });
      writeFileSync(dest, data);
      restored.push(rel);
    }
    this.log?.("info", `backup restored from ${relative(this.userDataDir, backupPath) || backupPath}: ${restored.length} file(s), ${skipped.length} skipped`);
    return { restored, skipped };
  }

  private collectFiles(): Array<{ rel: string; data: Uint8Array }> {
    const out: Array<{ rel: string; data: Uint8Array }> = [];
    for (const rel of BACKUP_FILES) {
      const p = join(this.userDataDir, rel);
      if (existsSync(p)) out.push({ rel, data: new Uint8Array(readFileSync(p)) });
    }
    for (const dir of BACKUP_DIRS) {
      const base = join(this.userDataDir, dir);
      if (!existsSync(base)) continue;
      walk(base, (p) => {
        // Zip entry names are POSIX-separated regardless of platform.
        const rel = join(dir, relative(base, p)).split("\\").join("/");
        out.push({ rel, data: new Uint8Array(readFileSync(p)) });
      });
    }
    return out;
  }
}

function walk(dir: string, visit: (file: string) => void): void {
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    const p = join(dir, entry.name);
    if (entry.isDirectory()) walk(p, visit);
    else if (entry.isFile()) visit(p);
  }
}

function isEncrypted(bytes: Uint8Array): boolean {
  if (bytes.length < BACKUP_MAGIC.length + 1 + 16 + 12 + 16) return false;
  const head = Buffer.from(bytes.slice(0, BACKUP_MAGIC.length)).toString("latin1");
  return head === BACKUP_MAGIC;
}

/** "RPDENC2" | u8=2 | salt(16) | iv(12) | tag(16) | ciphertext */
export function encrypt(plaintext: Uint8Array, password: string): Buffer {
  const salt = randomBytes(16);
  const iv = randomBytes(12);
  const key = pbkdf2Sync(password, salt, PBKDF2_ITERATIONS, 32, "sha256");
  const cipher = createCipheriv("aes-256-gcm", key, iv);
  const ciphertext = Buffer.concat([cipher.update(Buffer.from(plaintext)), cipher.final()]);
  const tag = cipher.getAuthTag();
  return Buffer.concat([
    Buffer.from(BACKUP_MAGIC, "latin1"),
    Buffer.from([BACKUP_FORMAT]),
    salt,
    iv,
    tag,
    ciphertext,
  ]);
}

export function decrypt(blob: Uint8Array, password: string): Uint8Array {
  const buf = Buffer.from(blob);
  let off = BACKUP_MAGIC.length;
  const version = buf[off]!;
  if (version !== BACKUP_FORMAT) throw new Error(`unsupported encryption version ${version}`);
  off += 1;
  const salt = buf.subarray(off, off + 16);
  off += 16;
  const iv = buf.subarray(off, off + 12);
  off += 12;
  const tag = buf.subarray(off, off + 16);
  off += 16;
  const key = pbkdf2Sync(password, salt, PBKDF2_ITERATIONS, 32, "sha256");
  const decipher = createDecipheriv("aes-256-gcm", key, iv);
  decipher.setAuthTag(tag); // throws on wrong password/tamper
  return new Uint8Array(Buffer.concat([decipher.update(buf.subarray(off)), decipher.final()]));
}

function sha256(data: Uint8Array): string {
  return createHash("sha256").update(Buffer.from(data)).digest("hex");
}
