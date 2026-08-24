import { join } from "node:path";
import { unlinkSync } from "node:fs";
import { JsonStore } from "./json-store.js";
import type { SecretCodec } from "./secret-codec.js";

export interface CloudSessionUser {
  id: string;
  steamId: string | null;
  name: string | null;
  displayName: string | null;
  email: string | null;
  providers: string[];
  hasPassword: boolean;
}

export interface CloudSession {
  token: string;
  user: CloudSessionUser;
  expiresAt: string | null;
}

type StoredSession = { schemaVersion: number; sealed: string };

export class CloudSessionStore {
  private readonly json: JsonStore<StoredSession>;
  private readonly file: string;

  constructor(
    userDataDir: string,
    private readonly codec: SecretCodec,
    private readonly log?: (level: "warn" | "error", message: string) => void,
  ) {
    this.file = join(userDataDir, "cloud-session.bin");
    this.json = new JsonStore<StoredSession>({
      file: this.file,
      schemaVersion: 1,
      validate: (doc): doc is StoredSession =>
        typeof doc === "object" && doc !== null && (doc as StoredSession).schemaVersion === 1 && typeof (doc as StoredSession).sealed === "string",
      log,
    });
  }

  load(): CloudSession | null {
    const outcome = this.json.load();
    if (outcome.status !== "loaded") return null;
    try {
      const value: unknown = JSON.parse(this.codec.open(outcome.doc.sealed));
      if (!isSession(value)) throw new Error("invalid session payload");
      return value;
    } catch (error) {
      this.log?.("error", `cloud session unreadable: ${error instanceof Error ? error.message : String(error)}`);
      return null;
    }
  }

  save(session: CloudSession): void {
    this.json.save({ schemaVersion: 1, sealed: this.codec.seal(JSON.stringify(session)) });
  }

  clear(): void {
    try { unlinkSync(this.file); } catch (error) {
      if ((error as NodeJS.ErrnoException).code !== "ENOENT") throw error;
    }
  }
}

function isSession(value: unknown): value is CloudSession {
  if (typeof value !== "object" || value === null) return false;
  const session = value as Partial<CloudSession>;
  return typeof session.token === "string" && session.token.length > 0 && isUser(session.user) &&
    (session.expiresAt === null || typeof session.expiresAt === "string");
}

function isUser(value: unknown): value is CloudSessionUser {
  if (typeof value !== "object" || value === null) return false;
  const user = value as Partial<CloudSessionUser>;
  return typeof user.id === "string" && (user.steamId === null || typeof user.steamId === "string") &&
    (user.name === null || typeof user.name === "string") && (user.displayName === null || typeof user.displayName === "string") &&
    (user.email === null || typeof user.email === "string") && Array.isArray(user.providers) &&
    user.providers.every((item) => typeof item === "string") && typeof user.hasPassword === "boolean";
}
