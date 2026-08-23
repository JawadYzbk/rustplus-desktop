/**
 * A2S query, query-port discovery, chat priming, watchdog — golden tests with fake UDP/network seams.
 */
import { describe, expect, it } from "vitest";
import { queryPlayers, type UdpSocket } from "../src/main/services/rustplus/a2s.js";
import {
  QueryPortResolver,
  buildPortCandidates,
} from "../src/main/services/rustplus/query-port.js";
import { ChatPrimer, StatusWatchdog } from "../src/main/services/rustplus/watchdog.js";

/** Scripted fake UDP socket: queue of datagrams to emit per recv(), recording sends. */
function fakeUdp(script: Array<Uint8Array | Error>, sent: Uint8Array[]): UdpSocket {
  const queue = [...script];
  return {
    async send(data) {
      sent.push(data);
    },
    recv() {
      const next = queue.shift();
      if (!next) return new Promise<Uint8Array>(() => undefined); // hang → timeout path
      if (next instanceof Error) return Promise.reject(next);
      return Promise.resolve(next);
    },
    close: () => undefined,
  };
}

const u8 = (...bytes: number[]): Uint8Array => Uint8Array.from(bytes);
const le32 = (v: number): number[] => [v & 0xff, (v >>> 8) & 0xff, (v >>> 16) & 0xff, (v >>> 24) & 0xff];

/** Builds a 0x44 player-list payload: count, then (index, name\0, score i32, duration f32). */
function playersPayload(entries: Array<[string, number]>): Uint8Array {
  const parts: number[] = [entries.length];
  for (const [name, duration] of entries) {
    parts.push(1); // index
    for (const b of Buffer.from(name, "utf8")) parts.push(b);
    parts.push(0); // null terminator
    parts.push(...le32(1337)); // score
    const dv = new DataView(new ArrayBuffer(4));
    dv.setFloat32(0, duration, true);
    for (let i = 0; i < 4; i++) parts.push(dv.getUint8(i));
  }
  return u8(...le32(0xffffffff), 0x44, ...parts);
}

describe("A2S queryPlayers", () => {
  it("direct 0x44 without challenge", async () => {
    const sent: Uint8Array[] = [];
    const udp = fakeUdp([playersPayload([["Alice", 12.5], ["Bob", 0]])], sent);
    const res = await queryPlayers("1.2.3.4", 28015, { udp: () => udp, timeoutMs: 1000 });
    expect(res).toEqual([
      { name: "Alice", duration: 12.5 },
      { name: "Bob", duration: 0 },
    ]);
    // Initial dummy-challenge request shape.
    expect(Array.from(sent[0]!)).toEqual([0xff, 0xff, 0xff, 0xff, 0x55, 0xff, 0xff, 0xff, 0xff]);
  });

  it("challenge resend: first 0x41 echoes token, second answer is the list", async () => {
    const sent: Uint8Array[] = [];
    const udp = fakeUdp(
      [u8(...le32(0xffffffff), 0x41, 0xde, 0xad, 0xbe, 0xef), playersPayload([["Cara", 3]])],
      sent,
    );
    const res = await queryPlayers("1.2.3.4", 28015, { udp: () => udp, timeoutMs: 1000 });
    expect(res).toEqual([{ name: "Cara", duration: 3 }]);
    expect(Array.from(sent[1]!.slice(0, 5))).toEqual([0xff, 0xff, 0xff, 0xff, 0x55]);
    expect(Array.from(sent[1]!.slice(5))).toEqual([0xde, 0xad, 0xbe, 0xef]);
  });

  it("split packets reassemble in order and parse the combined 0x44 payload", async () => {
    // Split payloads concatenate into a COMPLETE inner packet (own FF..FF header + cmd).
    const inner = playersPayload([
      ["PlayerOne", 60],
      ["PlayerTwo", 120],
    ]);
    const half = Math.ceil(inner.length / 2);
    const splitHeader = (id: number, total: number, num: number): Uint8Array =>
      u8(...le32(0xfffffffe), ...le32(id), total, num, 0x00, 0x10);

    const p1 = u8(...splitHeader(0x11223344, 2, 0), ...inner.subarray(0, half));
    const p2 = u8(...splitHeader(0x11223344, 2, 1), ...inner.subarray(half));

    // Out-of-order delivery must still reassemble by index.
    const udp = fakeUdp([u8(...le32(0xffffffff), 0x41, 0, 0, 0, 1), p2, p1], []);
    const res = await queryPlayers("1.2.3.4", 28015, { udp: () => udp, timeoutMs: 1000 });
    expect(res.map((p) => p.name)).toEqual(["PlayerOne", "PlayerTwo"]);
  });

  it("BZIP2-compressed splits are rejected outright", async () => {
    const id = 0x80000001;
    const inner = playersPayload([["X", 1]]);
    const udp = fakeUdp(
      [
        u8(...le32(0xffffffff), 0x41, 0, 0, 0, 1),
        u8(...le32(0xfffffffe), ...le32(id), 1, 0, 0x00, 0x04, ...inner),
      ],
      [],
    );
    await expect(queryPlayers("1.2.3.4", 28015, { udp: () => udp, timeoutMs: 1000 })).rejects.toThrowError(
      /BZIP2 Compression not supported/,
    );
  });

  it("timeout carries the legacy received/total split message", async () => {
    // Two split packets announced, only one ever delivered.
    const inner = playersPayload([["X", 1]]);
    const p1 = u8(...le32(0xfffffffe), ...le32(7), 2, 0, 0x00, 0x10, ...inner);
    const udp = fakeUdp([u8(...le32(0xffffffff), 0x41, 0, 0, 0, 1), p1], []);
    await expect(queryPlayers("1.2.3.4", 28015, { udp: () => udp, timeoutMs: 30 })).rejects.toThrowError(
      /Timeout waiting for player packets on port 28015\. Received 1\/2 split packets\./,
    );
  });

  it("reassembled garbage header throws the legacy invalid-header error", async () => {
    const junk = u8(0x11, 0x22, 0x33, 0x44);
    const udp = fakeUdp(
      [u8(...le32(0xffffffff), 0x41, 0, 0, 0, 1), u8(...le32(0xfffffffe), ...le32(9), 1, 0, 0x00, 0x04, ...junk)],
      [],
    );
    await expect(queryPlayers("1.2.3.4", 28015, { udp: () => udp, timeoutMs: 1000 })).rejects.toThrowError(
      /Invalid header in reassembled packet/,
    );
  });
});

describe("QueryPortResolver", () => {
  const okProbe = (host: string, port: number): Promise<unknown> =>
    port === 28016 ? Promise.resolve([]) : Promise.reject(new Error("timeout"));

  it("Steam API single Rust port trusted immediately and persisted as learned", async () => {
    const learned = new Map<string, number>();
    const r = new QueryPortResolver({
      steam: { fetchServersAtAddress: async () => [{ appid: 252490, addr: "1.2.3.4:28016" }] },
      learned: {
        get: (k) => learned.get(k),
        set: (k, v) => void learned.set(k, v),
      },
    });
    const res = await r.resolveQueryPort("1.2.3.4", 28082);
    expect(res).toEqual({ port: 28016, learned: false });
    expect(learned.get("1.2.3.4:28082")).toBe(28016);

    // Next call short-circuits from the learned cache.
    const again = await r.resolveQueryPort("1.2.3.4", 28082);
    expect(again).toEqual({ port: 28016, learned: true });
  });

  it("multiple ports pick the one closest to the companion port", async () => {
    const r = new QueryPortResolver({
      steam: {
        fetchServersAtAddress: async () => [
          { appid: 252490, addr: "1.2.3.4:27015" },
          { appid: 252490, addr: "1.2.3.4:28100" }, // closest to 28082
          { appid: 999999, addr: "1.2.3.4:1234" }, // filtered out
        ],
      },
      learned: { get: () => undefined, set: () => undefined },
      probe: async (host, port) => {
        if (host === "1.2.3.4" && port === 28100) return [];
        throw new Error("nope");
      },
    });
    expect((await r.resolveQueryPort("1.2.3.4", 28082))?.port).toBe(28100);
  });

  it("API failure falls back to the legacy candidate chain in exact order", () => {
    // appPort=28082 → [28082, 28015(=28082-67), 28015 dup skipped, 28081, 28082]
    expect(buildPortCandidates(28082, null)).toEqual([28082, 28015, 28081]);
    expect(buildPortCandidates(30000, null)).toEqual([30000, 29933, 28015, 29999]);
    expect(buildPortCandidates(28082, 28100)).toEqual([28100]); // discovered → no fallback chain
  });

  it("candidates probed in order until one answers", async () => {
    const probedPorts: number[] = [];
    const r = new QueryPortResolver({
      steam: { fetchServersAtAddress: async () => null },
      learned: { get: () => undefined, set: () => undefined },
      probe: async (_h, port) => {
        probedPorts.push(port);
        if (port === 28015) return [];
        throw new Error("timeout");
      },
    });
    const res = await r.resolveQueryPort("1.2.3.4", 28082);
    expect(probedPorts).toEqual([28082, 28015]);
    expect(res?.port).toBe(28015);
  });

  it("discovered ports are used directly without an extra probe (legacy parity)", async () => {
    let probes = 0;
    const r = new QueryPortResolver({
      steam: { fetchServersAtAddress: async () => [{ appid: 252490, addr: "1.2.3.4:28016" }] },
      learned: { get: () => undefined, set: () => undefined },
      probe: async () => {
        probes++;
        return [];
      },
    });
    const res = await r.resolveQueryPort("1.2.3.4", 28082);
    expect(res).toEqual({ port: 28016, learned: false });
    expect(probes).toBe(0); // the caller's main query validates it
  });
});

describe("ChatPrimer + StatusWatchdog", () => {
  it("chat primes exactly once per connection flag reset", async () => {
    const state = { teamChatPrimed: false, clanChatPrimed: false };
    let sends = 0;
    const primer = new ChatPrimer({
      raw: { sendRequestAsync: async () => ({ response: {} }) },
      send: async () => {
        sends++;
        return {};
      },
      toggleSwitch: async () => ({}),
      switchState: async () => null,
    });
    expect(await primer.primeTeamChat(state)).toBe(true);
    expect(await primer.primeTeamChat(state)).toBe(false);
    expect(sends).toBe(1);
    state.teamChatPrimed = false; // new connection
    expect(await primer.primeTeamChat(state)).toBe(true);
    expect(sends).toBe(2);
  });

  it("watchdog fires silent refresh after 5 consecutive failures; success resets", () => {
    let refreshes = 0;
    const w = new StatusWatchdog(5, () => refreshes++);
    for (let i = 0; i < 4; i++) w.recordFailure();
    expect(refreshes).toBe(0);
    w.recordSuccess(); // resets
    expect(w.current).toBe(0);
    for (let i = 0; i < 5; i++) w.recordFailure();
    expect(refreshes).toBe(1);
    expect(w.current).toBe(0); // counter restarted after firing
  });
});
