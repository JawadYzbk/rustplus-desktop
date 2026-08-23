/**
 * Query-port discovery — port of TrackingService.AutoDiscoverQueryPortAsync + PollSingleServerAsync
 * port-candidate logic (TrackingService.cs:1512-1659):
 *  - Steam Web API GetServersAtAddress filtered to appid 252490; single hit trusted, multiple hits
 *    pick the port closest to the companion port;
 *  - learned ports are persisted and skip discovery entirely;
 *  - otherwise probe candidates: [queryPort, appPort-67, 28015, appPort-1, appPort] (dedup'd, positive).
 *
 * Network + persistence are injected for testability.
 */
import { queryPlayers } from "./a2s.js";

export interface SteamApiFetcher {
  /** Returns parsed response.servers entries or null on any failure (5 s budget in legacy). */
  fetchServersAtAddress(host: string): Promise<Array<{ appid?: number; addr?: string }> | null>;
}

export interface LearnedPortStore {
  get(key: string): number | undefined;
  set(key: string, port: number): void;
}

export interface QueryPortDeps {
  steam: SteamApiFetcher;
  learned: LearnedPortStore;
  /** Probe a candidate query port; resolves players on success. Defaults to real A2S with 8 s budget. */
  probe?(host: string, port: number): Promise<unknown>;
}

const RUST_APPID = 252490;

export function buildPortCandidates(appPort: number, discovered: number | null): number[] {
  const candidates = discovered !== null ? [discovered] : [appPort];
  if (discovered === null) {
    // Legacy fallback order preserved exactly.
    for (const fb of [appPort - 67, 28015, appPort - 1, appPort]) {
      if (fb > 0 && !candidates.includes(fb)) candidates.push(fb);
    }
  }
  return candidates.filter((p) => p > 0);
}

export class QueryPortResolver {
  private readonly probe: (host: string, port: number) => Promise<unknown>;

  constructor(private readonly deps: QueryPortDeps) {
    this.probe = deps.probe ?? ((host: string, port: number) => queryPlayers(host, port, { timeoutMs: 8_000 }));
  }

  async discover(host: string, appPort: number): Promise<number | null> {
    try {
      const servers = await this.deps.steam.fetchServersAtAddress(host);
      if (!servers) return null;
      const rustPorts: number[] = [];
      for (const s of servers) {
        if (s.appid === RUST_APPID && typeof s.addr === "string") {
          const parts = s.addr.split(":");
          if (parts.length === 2) {
            const sp = Number(parts[1]);
            if (Number.isInteger(sp)) rustPorts.push(sp);
          }
        }
      }
      if (rustPorts.length === 1) return rustPorts[0]!;
      if (rustPorts.length > 1) {
        // Closest to the companion port.
        return rustPorts.reduce((best, p) =>
          Math.abs(p - appPort) < Math.abs(best - appPort) ? p : best,
        );
      }
      return null;
    } catch {
      return null;
    }
  }

  /**
   * Resolve the working query port — legacy order:
   * learned cache → Steam API discovery (persisted on hit, used DIRECTLY — the caller's main
   * query is its own validation) → candidate-chain probing only when discovery gave nothing.
   * Returns the port, or null when nothing answered.
   */
  async resolveQueryPort(host: string, appPort: number): Promise<{ port: number; learned: boolean } | null> {
    const key = `${host}:${appPort}`;
    const known = this.deps.learned.get(key);
    if (known !== undefined) return { port: known, learned: true };

    const discovered = await this.discover(host, appPort);
    if (discovered !== null) {
      this.deps.learned.set(key, discovered); // save immediately so later calls skip discovery
      return { port: discovered, learned: false };
    }

    // API gave nothing — try the most common Rust query ports (legacy fallback chain).
    for (const candidate of buildPortCandidates(appPort, null)) {
      try {
        await this.probe(host, candidate);
        return { port: candidate, learned: false };
      } catch {
        continue;
      }
    }
    return null;
  }
}
