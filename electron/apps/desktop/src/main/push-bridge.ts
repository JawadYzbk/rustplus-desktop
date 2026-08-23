/**
 * Push bridge — forwards main-process runtime events to renderer windows over the one-way
 * `conn/push` channel. Destroyed windows are skipped; missing main window simply means no-op.
 */
import { pushChannel } from "@rpd/shared";

export interface PushTarget {
  send(channel: string, payload: unknown): void;
  isDestroyed(): boolean;
}

export type PushEmitter = (stream: "conn" | "poll" | "device", event: unknown) => void;

export function createPushBridge(getTargets: () => PushTarget[]): PushEmitter {
  return (stream, event) => {
    const payload = { stream, event };
    for (const target of getTargets()) {
      if (target.isDestroyed()) continue;
      try {
        target.send(pushChannel, payload);
      } catch {
        /* window died between check and send — next tick filters it out */
      }
    }
  };
}
