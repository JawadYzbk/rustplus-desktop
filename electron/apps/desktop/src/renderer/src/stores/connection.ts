import { create } from "zustand";
import {
  connectProfile as connectProfileIpc,
  disconnect as disconnectIpc,
  getConnectionStatus,
  type ConnSnapshot,
} from "../lib/ipc.js";

export interface TeamMember {
  steamId: string;
  name: string;
  online: boolean;
  dead: boolean;
  x: number | null;
  y: number | null;
}

export interface TeamSnapshot {
  leaderSteamId: string | null;
  members: TeamMember[];
  receivedAt: number;
}

export interface ServerStatus {
  players: number;
  maxPlayers: number;
  queuedPlayers: number;
  timeString: string | null;
}

export type ConnectionPhase = "disconnected" | "connecting" | "connected" | "reconnecting";

const EMPTY_SNAPSHOT: ConnSnapshot = {
  connected: false,
  activeProxy: null,
  host: null,
  port: null,
  consecutiveTimeouts: 0,
  teamChatPrimed: false,
  clanChatPrimed: false,
};

function record(value: unknown): Record<string, unknown> {
  return value !== null && typeof value === "object" ? value as Record<string, unknown> : {};
}

function text(value: unknown, fallback: string): string {
  return typeof value === "string" && value.trim() ? value.trim() : fallback;
}

function coordinate(value: unknown): number | null {
  return typeof value === "number" && Number.isFinite(value) ? value : null;
}

/** Normalize the raw Rust+ teamInfo payload at the renderer boundary. */
export function normalizeTeamSnapshot(value: unknown, receivedAt = Date.now()): TeamSnapshot {
  const raw = record(value);
  const leader = raw.leaderSteamId ?? raw.leaderSteamId64 ?? raw.leaderId;
  const members = Array.isArray(raw.members) ? raw.members.flatMap((value): TeamMember[] => {
    const member = record(value);
    const rawId = member.steamId ?? member.steamId64 ?? member.userId ?? member.playerId;
    if (rawId === undefined || String(rawId) === "0") return [];
    const position = record(member.position ?? member.pos);
    const dead = member.dead ?? member.isDead ?? (member.alive === undefined ? false : !Boolean(member.alive));
    return [{
      steamId: String(rawId),
      name: text(member.name ?? member.displayName, "(player)"),
      online: Boolean(member.online ?? member.isOnline),
      dead: Boolean(dead),
      x: coordinate(member.x ?? position.x),
      y: coordinate(member.y ?? position.y),
    }];
  }) : [];
  return {
    leaderSteamId: leader === undefined || leader === null ? null : String(leader),
    members,
    receivedAt,
  };
}

interface ConnectionState {
  snapshot: ConnSnapshot;
  phase: ConnectionPhase;
  team: TeamSnapshot | null;
  status: ServerStatus | null;
  error: string | null;
  hydrate: () => Promise<void>;
  connectProfile: (matchKey: string, useProxy?: boolean) => Promise<ConnSnapshot>;
  disconnect: () => Promise<void>;
  applyPush: (stream: string, event: unknown) => void;
}

export const useConnectionStore = create<ConnectionState>((set) => ({
  snapshot: EMPTY_SNAPSHOT,
  phase: "disconnected",
  team: null,
  status: null,
  error: null,

  hydrate: async () => {
    try {
      const snapshot = await getConnectionStatus();
      set({ snapshot, phase: snapshot.connected ? "connected" : "disconnected", error: null });
    } catch (reason: unknown) {
      set({ error: reason instanceof Error ? reason.message : String(reason) });
    }
  },

  connectProfile: async (matchKey, useProxy) => {
    set({ phase: "connecting", error: null, team: null, status: null });
    try {
      const snapshot = await connectProfileIpc(matchKey, useProxy);
      set({ snapshot, phase: "connected", error: null });
      return snapshot;
    } catch (reason: unknown) {
      const error = reason instanceof Error ? reason.message : String(reason);
      set({ phase: "disconnected", error });
      throw reason;
    }
  },

  disconnect: async () => {
    try {
      const snapshot = await disconnectIpc();
      set({ snapshot, phase: "disconnected", team: null, status: null, error: null });
    } catch (reason: unknown) {
      set({ error: reason instanceof Error ? reason.message : String(reason) });
      throw reason;
    }
  },

  applyPush: (stream, event) => {
    const raw = record(event);
    if (stream === "conn") {
      const kind = raw.kind;
      if (kind === "connecting") set({ phase: "connecting", error: null });
      else if (kind === "connected") set({ phase: "connected", error: null });
      else if (kind === "reconnectingIn") set({ phase: "reconnecting" });
      else if (kind === "lost") set({ phase: "reconnecting", snapshot: { ...EMPTY_SNAPSHOT, host: null, port: null } });
      else if (kind === "disconnected") set({ phase: "disconnected", snapshot: EMPTY_SNAPSHOT, team: null, status: null });
      return;
    }
    if (stream === "poll" && raw.kind === "status") {
      const value = record(raw.status);
      const numberOr = (candidate: unknown, fallback: number): number => typeof candidate === "number" && Number.isFinite(candidate) ? candidate : fallback;
      set({ status: {
        players: numberOr(value.players, 0),
        maxPlayers: numberOr(value.maxPlayers, 0),
        queuedPlayers: numberOr(value.queuedPlayers ?? value.queue, 0),
        timeString: typeof value.timeString === "string" ? value.timeString : null,
      } });
      return;
    }
    if (stream === "poll" && raw.kind === "team") {
      set({ team: normalizeTeamSnapshot(raw.team) });
    }
  },
}));
