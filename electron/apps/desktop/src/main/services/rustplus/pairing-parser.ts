/**
 * Pairing-listener stdout line parser — faithful port of HandleListenOutput +
 * ProcessBodyJson/PublishServerDescription/BufferAlarm/FlushBufferedAlarm/TryFlushChat/
 * TryFlushOfflineDeath (PairingListenerRealProcess.cs:99-298, 428-903).
 *
 * The legacy app consumes rustplus-cli's stdout text; this state machine turns those lines into
 * typed events. All quirks preserved deliberately:
 *  - alarm bodies buffer until the FCM persistentId arrives (dedup survives restarts);
 *  - "expirtyDate" typo key parity;
 *  - raid alarms without explicit type ("Raid Alarm" name fallback);
 *  - listener-side dedup: identical host:port|steam|token|entity within 20 s ignored;
 *  - chat bundles open/close on channelId == "chat" and reset last ip/port.
 */
export type ListenerEvent =
  | { kind: "listening" }
  | { kind: "failed"; line: string }
  | { kind: "paired"; payload: PairingPayload }
  | { kind: "serverInfo"; payload: ServerInfoPayload }
  | { kind: "alarm"; alarm: AlarmNotification }
  | { kind: "chat"; message: TeamChatMessage }
  | { kind: "offlineDeath"; death: OfflineDeathNotification };

export interface PairingPayload {
  Host: string;
  Port: number;
  ServerName?: string | null;
  ServerDescription?: string | null;
  SteamId64?: string;
  PlayerToken?: string;
  EntityId?: number | null;
  EntityName?: string | null;
  EntityType?: string | null;
  IssueDate?: string | null;
  ExpiryDate?: string | null;
}

export interface ServerInfoPayload {
  Host: string;
  Port: number;
  ServerName?: string | null;
  ServerDescription: string;
}

export interface AlarmNotification {
  ts: number;
  server: string;
  deviceName: string;
  entityId?: number | null;
  message: string;
  ip?: string | null;
  port?: number | null;
  title?: string | null;
  fcmId?: string | null;
}

export interface TeamChatMessage {
  ts: number;
  author: string;
  text: string;
  ip?: string | null;
  port?: number | null;
}

export interface OfflineDeathNotification {
  ts: number;
  server: string;
  attacker: string;
  ip?: string | null;
  port?: number | null;
}

const ANSI = /\x1B\[[0-9;]*[A-Za-z]/g;
const RUST_URL = /rustplus:\/\/[^\s'">]+/i;
// eslint-disable-next-line no-control-regex
const KV_LINE = /\{\s*key:\s*'([^']+)'\s*,\s*value:\s*(?:'|\"|`)([\s\S]*?)(?:'|\"|`)\s*\}/;
const DEATH_TITLE = /^(?:You were killed by|Du wurdest getötet von)\s+(?<attacker>.+)/i;
const TOP_LEVEL_ID = /^\s*(?<type>id|persistentId):\s*["'](?<id>[^"']+)["']/;
const BODY_JSON = /value:\s*(?:'|`)(?<json>\{[\s\S]*?\})(?:'|`)/;
const MSG_LINE = /\{\s*key:\s*'(?:message|gcm\.notification\.body)'\s*,\s*value:\s*'([^']+)'\s*\}/;

const PAIR_DEDUP_WINDOW_MS = 20_000;

interface PendingAlarmCtx {
  server: string | null;
  entityName: string | null;
  entityId: number | null;
  host: string | null;
  port: number | null;
}

function jget(root: Record<string, unknown>, ...names: string[]): string | null {
  for (const n of names) {
    const v = root[n];
    if (v !== undefined && v !== null && typeof v !== "object") return String(v);
  }
  const lower = new Map(Object.keys(root).map((k) => [k.toLowerCase(), k]));
  for (const n of names) {
    const actual = lower.get(n.toLowerCase());
    if (actual !== undefined) {
      const v = root[actual];
      if (v !== undefined && v !== null && typeof v !== "object") return String(v);
    }
  }
  return null;
}

function parseUint(s: string | null): number | null {
  if (!s) return null;
  const u = Number(s);
  return Number.isFinite(u) && u >= 0 ? Math.floor(u) : null;
}

function parseIntStrict(s: string | null): number | null {
  if (!s || !/^-?\d+$/.test(s.trim())) return null;
  const n = Number(s.trim());
  return Number.isSafeInteger(n) ? n : null;
}

/** Extracts "serverDescription"/"server_description"/"description", descending into server/serverInfo. */
export function getServerDescription(root: Record<string, unknown>): string | null {
  const direct = jget(root, "serverDescription", "server_description", "description");
  if (direct && direct.trim()) return direct;
  for (const container of ["server", "serverInfo"]) {
    const nested = root[container];
    if (nested && typeof nested === "object") {
      const d = jget(nested as Record<string, unknown>, "description");
      if (d && d.trim()) return d;
    }
  }
  return null;
}

function tryParseRustPlusUrl(url: string): PairingPayload | null {
  try {
    const qIndex = url.indexOf("?");
    if (qIndex < 0) return null;
    const query = url.slice(qIndex + 1);

    let ip: string | undefined;
    let portStr: string | undefined;
    let name: string | undefined;
    let playerId: string | undefined;
    let playerToken: string | undefined;

    for (const part of query.split("&")) {
      if (!part) continue;
      const eq = part.indexOf("=");
      const k = decodeURIComponent(eq < 0 ? part : part.slice(0, eq)).toLowerCase();
      const v = eq < 0 ? "" : decodeURIComponent(part.slice(eq + 1));
      switch (k) {
        case "ip":
        case "host":
          ip = v;
          break;
        case "port":
          portStr = v;
          break;
        case "name":
          name = v;
          break;
        case "playerid":
          playerId = v;
          break;
        case "playertoken":
          playerToken = v;
          break;
      }
    }

    if (!ip?.trim() || !playerId?.trim() || !playerToken?.trim()) return null;
    const port = parseIntStrict(portStr ?? null) ?? 28082;
    return {
      Host: ip,
      Port: port,
      ServerName: name?.trim() ? name : null,
      SteamId64: playerId,
      PlayerToken: playerToken,
    };
  } catch {
    return null;
  }
}

export class PairingLineParser {
  // Per-notification context (reset on "Notification Received").
  private pendingAlarmTitle: string | null = null;
  private pendingFcmId: string | null = null;
  private bufferedAlarm: AlarmNotification | null = null;

  // Key/value streaming state.
  private lastIp: string | null = null;
  private lastPort: number | null = null;

  // Chat bundle.
  private chatBundleOpen = false;
  private pendingChatTitle: string | null = null;
  private pendingChatMsg: string | null = null;
  private pendingChatTs: number | null = null;
  private pendingChatIp: string | null = null;
  private pendingChatPort: number | null = null;

  // Offline death.
  private pendingDeathAttacker: string | null = null;
  private pendingDeathServer: string | null = null;
  private pendingDeathTs: number | null = null;
  private pendingDeathIp: string | null = null;
  private pendingDeathPort: number | null = null;

  // Alarm assembly.
  private pendingAlarm: PendingAlarmCtx | null = null;
  private pendingAlarmMsg: string | null = null;
  private pendingAlarmMsgTs: number | null = null;

  // Multiline JSON collection.
  private collectingJson = false;
  private jsonBuffer: string[] = [];

  // Listener-side pairing dedup.
  private lastPairKey: string | null = null;
  private lastPairAt = -Infinity;

  constructor(private readonly now: () => number = Date.now) {}

  feed(line: string): ListenerEvent[] {
    if (!line || !line.trim()) return [];
    const s = line.replace(ANSI, "").trim();
    const out: ListenerEvent[] = [];

    // New FCM notification starts – flush any buffered alarm, reset per-message context.
    if (s.toLowerCase().includes("notification received")) {
      this.flushBufferedAlarm(out);
      this.pendingAlarmTitle = null;
      this.pendingFcmId = null;
    }

    // Top-level FCM id (preferred persistentId).
    const topId = TOP_LEVEL_ID.exec(s);
    if (topId?.groups) {
      this.pendingFcmId = topId.groups["id"]!;
      if (topId.groups["type"]!.toLowerCase() === "persistentid") this.flushBufferedAlarm(out);
    }

    // End of FCM notification object.
    if (s === "}") this.flushBufferedAlarm(out);

    // CLI status markers.
    if (s.toLowerCase().includes("listening for fcm notifications")) out.push({ kind: "listening" });
    if (s.toLowerCase().includes("error") || s.includes("ERR!")) out.push({ kind: "failed", line: s });

    // 0) rustplus:// deep link.
    const lm = RUST_URL.exec(s);
    if (lm) {
      const payload = tryParseRustPlusUrl(lm[0]);
      if (payload) {
        out.push({ kind: "paired", payload });
        return out;
      }
    }

    // 0.1) Single-line JSON check bypasses multiline collection.
    const mSingle = BODY_JSON.exec(s);
    if (mSingle?.groups) {
      this.collectingJson = false;
      this.processBodyJson(mSingle.groups["json"]!, out);
      this.rest(s, out);
      return out;
    }

    // A) raw key/value lines (channelId/title/body/ip/port)
    const kv = KV_LINE.exec(s);
    if (kv) {
      const k = kv[1]!;
      const v = kv[2]!;

      if (k.toLowerCase() === "ip" || k.toLowerCase() === "gcm.notification.ip") {
        this.lastIp = v;
        if (this.chatBundleOpen) this.pendingChatIp = v;
        else if (this.pendingDeathAttacker) this.pendingDeathIp = v;
        else if (this.pendingAlarm) this.pendingAlarm.host = v;
      } else if (k.toLowerCase() === "port" || k.toLowerCase() === "gcm.notification.port") {
        const portVal = parseIntStrict(v);
        if (portVal !== null) {
          this.lastPort = portVal;
          if (this.chatBundleOpen) this.pendingChatPort = portVal;
          else if (this.pendingDeathAttacker) this.pendingDeathPort = portVal;
          else if (this.pendingAlarm) this.pendingAlarm.port = portVal;
        }
      }

      if (k.toLowerCase() === "gcm.notification.android_channel_id" || k.toLowerCase() === "channelid") {
        this.chatBundleOpen = v.toLowerCase() === "chat";
        if (!this.chatBundleOpen) {
          this.pendingChatMsg = null;
          this.pendingChatTitle = null;
          this.pendingChatTs = null;
        }
        this.lastIp = null;
        this.lastPort = null;
      }

      // Offline-death title capture runs BEFORE generic title handling and returns early on match.
      if (k.toLowerCase() === "title" || k.toLowerCase() === "gcm.notification.title") {
        const mDeath = DEATH_TITLE.exec(v);
        if (mDeath?.groups) {
          this.pendingDeathAttacker = mDeath.groups["attacker"]!.replace(/^['"]|['"]$/g, "");
          this.pendingDeathTs = this.now();
          return out;
        }
        this.pendingAlarmTitle = v;
        if (this.chatBundleOpen) {
          this.pendingChatTitle = v;
          this.tryFlushChat(out);
          return out;
        }
      }

      if (k.toLowerCase() === "body" || k.toLowerCase() === "gcm.notification.body") {
        if (this.pendingDeathAttacker) {
          this.pendingDeathServer = v;
          this.tryFlushOfflineDeath(out);
          return out;
        }
      }
    } else if (
      s.includes("key: 'body'") ||
      s.includes("key: 'gcm.notification.body'") ||
      (s.includes(`value: '{"`) && !mSingle)
    ) {
      // Start of a multiline value block.
      this.collectingJson = true;
      this.jsonBuffer = [s];
      return out;
    }

    // B) message/body lines
    const mm = MSG_LINE.exec(s);
    if (mm && this.chatBundleOpen) {
      this.pendingChatMsg = mm[1]!;
      this.pendingChatTs = this.now();
      this.tryFlushChat(out);
      return out;
    }

    // C) JSON "value: '...'"
    const m = BODY_JSON.exec(s);
    if (m?.groups) this.processBodyJson(m.groups["json"]!, out);

    this.rest(s, out);
    return out;
  }

  /** Second pass mirroring HandleListenOutputRest: alarm message flush + body JSON pairing/alarm. */
  private rest(s: string, out: ListenerEvent[]): void {
    const mm = MSG_LINE.exec(s);
    const m = BODY_JSON.exec(s);

    // 1) ALARM via message lines (may come before or after the body)
    if (mm) {
      this.pendingAlarmMsg = mm[1]!;
      this.pendingAlarmMsgTs = this.now();
      const ctx = this.pendingAlarm;
      if (ctx) {
        // C# parity: BufferAlarm only — flush happens on Notification Received / persistentId / "}".
        this.bufferedAlarm = {
          ts: this.pendingAlarmMsgTs,
          server: ctx.server ?? "-",
          deviceName: `${ctx.entityName ?? "Alarm"}${ctx.entityId != null ? `#${ctx.entityId}` : ""}`,
          entityId: ctx.entityId,
          message: this.pendingAlarmMsg,
          ip: ctx.host,
          port: ctx.port,
          title: this.pendingAlarmTitle,
          fcmId: null,
        };
        this.pendingAlarm = null;
        this.pendingAlarmMsg = null;
        this.pendingAlarmMsgTs = null;
      }
      return;
    }

    // 2) appData body JSON
    if (m?.groups) {
      let root: Record<string, unknown>;
      try {
        root = JSON.parse(m.groups["json"]!) as Record<string, unknown>;
      } catch {
        return; // legacy logs and falls through without acting
      }

      const type = jget(root, "type");
      const host = jget(root, "ip");
      const portRaw = parseIntStrict(jget(root, "port"));
      const port = portRaw ?? 28082;
      const name = jget(root, "name");
      const description = getServerDescription(root);
      const playerId = jget(root, "playerId");
      const playerToken = jget(root, "playerToken");
      const entityId = parseUint(jget(root, "entityId", "entityID"));
      const entityName = jget(root, "entityName");
      const entityType = jget(root, "entityType"); // "1" Switch / "2" Alarm

      // Kind inference: entityType code first, then entity-name heuristics.
      let kind: string | null = null;
      if (entityType) {
        if (entityType === "1") kind = "SmartSwitch";
        else if (entityType === "2") kind = "SmartAlarm";
      }
      if (kind === null && entityName) {
        if (/switch/i.test(entityName)) kind = "SmartSwitch";
        else if (/alarm/i.test(kind ?? "") || /alarm/i.test(entityName)) kind = "SmartAlarm";
      }

      // SERVER / ENTITY → Paired
      if (
        host?.trim() &&
        playerId?.trim() &&
        playerToken?.trim() &&
        (type?.toLowerCase() === "server" || type?.toLowerCase() === "entity")
      ) {
        const payload: PairingPayload = {
          Host: host,
          Port: port,
          ServerName: name?.trim() ? name : null,
          ServerDescription: description?.trim() ? description : null,
          SteamId64: playerId,
          PlayerToken: playerToken,
          EntityId: entityId,
          EntityName: entityName?.trim() ? entityName : null,
          EntityType: kind ?? type,
          IssueDate: jget(root, "issueDate"),
          ExpiryDate: jget(root, "expiryDate", "expirtyDate"), // typo key parity
        };

        const key = `${payload.Host}:${payload.Port}|${payload.SteamId64}|${payload.PlayerToken}|${payload.EntityId}`;
        if (this.lastPairKey === key && this.now() - this.lastPairAt < PAIR_DEDUP_WINDOW_MS) {
          return; // same-pairing bounce inside 20 s ignored
        }
        this.lastPairKey = key;
        this.lastPairAt = this.now();
        out.push({ kind: "paired", payload });
        return;
      }

      // ALARM with explicit type
      if (type?.toLowerCase() === "alarm") {
        this.pendingAlarm = { server: name, entityName, entityId, host, port };
        if (this.pendingAlarmMsg != null) {
          this.bufferedAlarm = {
            ts: this.pendingAlarmMsgTs ?? this.now(),
            server: name ?? "-",
            deviceName: `${entityName ?? "Alarm"}${entityId != null ? `#${entityId}` : ""}`,
            entityId,
            message: this.pendingAlarmMsg,
            ip: host,
            port,
            title: this.pendingAlarmTitle,
            fcmId: null,
          };
          this.pendingAlarm = null;
          this.pendingAlarmMsg = null;
          this.pendingAlarmMsgTs = null;
        }
        return;
      }

      // Raid Alarm without explicit type
      if (!type && host) {
        this.pendingAlarm = { server: name, entityName: "Raid Alarm", entityId: null, host, port };
        this.lastIp = host;
        this.lastPort = port;
        if (this.pendingAlarmMsg != null) {
          this.bufferedAlarm = {
            ts: this.pendingAlarmMsgTs ?? this.now(),
            server: name ?? "-",
            deviceName: "Raid Alarm",
            entityId: null,
            message: this.pendingAlarmMsg,
            ip: host,
            port,
            title: this.pendingAlarmTitle,
            fcmId: null,
          };
          this.pendingAlarm = null;
          this.pendingAlarmMsg = null;
          this.pendingAlarmMsgTs = null;
        }
        return;
      }
    }
  }

  private processBodyJson(json: string, out: ListenerEvent[]): void {
    let root: Record<string, unknown>;
    try {
      root = JSON.parse(json) as Record<string, unknown>;
    } catch {
      return;
    }

    const type = jget(root, "type");
    const ip = jget(root, "ip");
    const port = parseIntStrict(jget(root, "port"));
    this.publishServerDescription(root, ip, port, out);

    if (type?.toLowerCase() === "chat") {
      const author = jget(root, "name") ?? jget(root, "username") ?? "Team";
      const text = jget(root, "message") ?? this.pendingChatMsg ?? "";
      out.push({
        kind: "chat",
        message: {
          ts: this.now(),
          author,
          text,
          ip: ip ?? this.pendingChatIp ?? this.lastIp,
          port: port ?? this.pendingChatPort ?? this.lastPort,
        },
      });
      this.pendingChatMsg = null;
      this.pendingChatTitle = null;
      this.pendingChatTs = null;
      this.pendingChatIp = null;
      this.pendingChatPort = null;
    } else if (type?.toLowerCase() === "death") {
      const serverName = jget(root, "name");
      if (ip) this.pendingDeathIp = ip;
      if (port !== null) this.pendingDeathPort = port;
      if (serverName) this.pendingDeathServer = serverName;
      this.tryFlushOfflineDeath(out);
    }
  }

  private publishServerDescription(
    root: Record<string, unknown>,
    ip: string | null,
    port: number | null,
    out: ListenerEvent[],
  ): void {
    let nestedHost: string | null = null;
    let nestedPort: number | null = null;
    let nestedName: string | null = null;
    for (const container of ["server", "serverInfo"]) {
      const nested = root[container];
      if (nested && typeof nested === "object") {
        const o = nested as Record<string, unknown>;
        nestedHost ??= jget(o, "ip");
        nestedPort ??= parseIntStrict(jget(o, "port"));
        nestedName ??= jget(o, "name");
      }
    }

    const description = getServerDescription(root);
    if (!description || !description.trim()) return;

    out.push({
      kind: "serverInfo",
      payload: {
        Host: ip ?? nestedHost ?? this.lastIp ?? "",
        Port: port ?? nestedPort ?? this.lastPort ?? 0,
        ServerName: jget(root, "serverName") ?? jget(root, "name") ?? nestedName,
        ServerDescription: description.trim(),
      },
    });
  }

  private bufferedForFlush(): ListenerEvent | null {
    if (!this.bufferedAlarm) return null;
    const alarm = this.bufferedAlarm;
    this.bufferedAlarm = null;
    this.pendingAlarmTitle = null;
    const fcmId = this.pendingFcmId;
    this.pendingFcmId = null;
    return { kind: "alarm", alarm: { ...alarm, fcmId } };
  }

  private flushBufferedAlarm(out: ListenerEvent[]): void {
    const evt = this.bufferedForFlush();
    if (evt) out.push(evt);
  }

  private tryFlushChat(out: ListenerEvent[]): void {
    if (!this.chatBundleOpen || !this.pendingChatMsg) return;
    const author = this.pendingChatTitle?.trim() ? this.pendingChatTitle : "Team";
    out.push({
      kind: "chat",
      message: {
        ts: this.pendingChatTs ?? this.now(),
        author,
        text: this.pendingChatMsg,
        ip: this.pendingChatIp ?? this.lastIp,
        port: this.pendingChatPort ?? this.lastPort,
      },
    });
    this.pendingChatMsg = null;
    this.pendingChatTitle = null;
    this.pendingChatTs = null;
    this.pendingChatIp = null;
    this.pendingChatPort = null;
  }

  private tryFlushOfflineDeath(out: ListenerEvent[]): void {
    if (!this.pendingDeathAttacker || !this.pendingDeathServer) return;
    out.push({
      kind: "offlineDeath",
      death: {
        ts: this.pendingDeathTs ?? this.now(),
        server: this.pendingDeathServer,
        attacker: this.pendingDeathAttacker,
        ip: this.pendingDeathIp ?? this.lastIp,
        port: this.pendingDeathPort ?? this.lastPort,
      },
    });
    this.pendingDeathAttacker = null;
    this.pendingDeathServer = null;
    this.pendingDeathTs = null;
    this.pendingDeathIp = null;
    this.pendingDeathPort = null;
  }
}
