/**
 * A2S_PLAYER UDP query — faithful port of A2SClient.QueryPlayersAsync (A2SClient.cs):
 *  - initial A2S_PLAYER request with dummy challenge FFFFFFFF;
 *  - first response may be 0x41 challenge (re-send with real token) or direct 0x44 player list;
 *  - multi-packet reassembly on header 0xFFFFFFFE (id/total/num/size);
 *  - BZIP2-compressed splits explicitly rejected (parity);
 *  - missing split packets are skipped during concat exactly like the C# (may produce an
 *    "Invalid header" error rather than hanging — preserved).
 *
 * The UDP socket is injected so golden tests run without a network.
 */
import { createSocket, type RemoteInfo } from "node:dgram";
import { realClock, type Clock } from "./timing.js";

export interface A2SPlayer {
  name: string;
  duration: number;
}

export interface UdpSocket {
  send(data: Uint8Array): Promise<void>;
  /** Resolves with the next datagram; rejects on close/error. */
  recv(): Promise<Uint8Array>;
  close(): void;
}

export type UdpFactory = (host: string, port: number) => UdpSocket;

export const realUdpFactory: UdpFactory = (host, port) => {
  const sock = createSocket("udp4");
  sock.bind({ address: "0.0.0.0" }, () => sock.connect(port, host));
  return {
    send: (data) =>
      new Promise((resolve, reject) =>
        sock.send(Buffer.from(data), (err?: Error | null) => (err ? reject(err) : resolve())),
      ),
    recv: () =>
      new Promise<Uint8Array>((resolve, reject) => {
        const onMsg = (msg: Buffer, _rinfo: RemoteInfo): void => {
          sock.off("message", onMsg);
          sock.off("error", onErr);
          resolve(new Uint8Array(msg));
        };
        const onErr = (err: Error): void => {
          sock.off("message", onMsg);
          sock.off("error", onErr);
          reject(err);
        };
        sock.on("message", onMsg);
        sock.on("error", onErr);
      }),
    close: () => {
      try {
        sock.close();
      } catch {
        /* already closed */
      }
    },
  };
};

async function resolveHost(host: string): Promise<string> {
  // IPAddress.TryParse parity: plain IPv4 passes through, otherwise DNS.
  if (/^(\d{1,3}\.){3}\d{1,3}$/.test(host)) return host;
  const { lookup } = await import("node:dns/promises");
  const addrs = await lookup(host, { family: 4 });
  return addrs.address;
}

function u32le(b: Uint8Array, off: number): number {
  return (b[off]! | (b[off + 1]! << 8) | (b[off + 2]! << 16) | (b[off + 3]! << 24)) >>> 0;
}

class ByteReader {
  private off = 0;
  constructor(private readonly buf: Uint8Array) {}
  get remaining(): number {
    return this.buf.length - this.off;
  }
  u8(): number {
    return this.buf[this.off++]!;
  }
  i32(): number {
    const v = u32le(this.buf, this.off);
    this.off += 4;
    return v | 0;
  }
  f32(): number {
    const view = new DataView(this.buf.buffer, this.buf.byteOffset + this.off, 4);
    this.off += 4;
    return view.getFloat32(0, true);
  }
  nullString(): string {
    const bytes: number[] = [];
    while (this.remaining > 0) {
      const b = this.u8();
      if (b === 0) break;
      bytes.push(b);
    }
    return new TextDecoder().decode(Uint8Array.from(bytes));
  }
}

function parsePlayers(br: ByteReader): A2SPlayer[] {
  const list: A2SPlayer[] = [];
  if (br.remaining <= 0) return list;
  const count = br.u8();
  for (let i = 0; i < count; i++) {
    if (br.remaining <= 0) break;
    br.u8(); // index
    const name = br.nullString();
    if (br.remaining < 8) break; // needs score(4) + duration(4)
    br.i32(); // score — tracked players match by name in the caller (legacy parity)
    list.push({ name, duration: br.f32() });
  }
  return list;
}

const sleep = (ms: number): Promise<void> => new Promise((r) => setTimeout(r, ms));

/** Query A2S_PLAYERS. Throws with legacy-shaped messages on timeout/compression/bad headers. */
export async function queryPlayers(
  hostInput: string,
  port: number,
  opts: { timeoutMs?: number; udp?: UdpFactory; clock?: Clock } = {},
): Promise<A2SPlayer[]> {
  const timeoutMs = opts.timeoutMs ?? 3_000;
  const clock = opts.clock ?? realClock;
  const makeSocket = opts.udp ?? realUdpFactory;

  const host = await resolveHost(hostInput);
  const udp = makeSocket(host, port);
  const start = clock.now();

  try {
    // 1. Initial A2S_PLAYER with dummy challenge.
    await udp.send(Uint8Array.of(0xff, 0xff, 0xff, 0xff, 0x55, 0xff, 0xff, 0xff, 0xff));

    // 2. First response: 0x41 challenge → re-send; 0x44 → direct list.
    let needLoop = true;
    {
      const res = await recvOrTimeout(udp, deadlineLeft(start, timeoutMs, clock));
      if (res.length >= 5 && u32le(res, 0) === 0xffffffff) {
        const cmd = res[4]!;
        if (cmd === 0x44) return parsePlayers(new ByteReader(res.subarray(5)));
        if (cmd === 0x41) {
          const challenge = res.subarray(5, 9);
          const req2 = Uint8Array.of(0xff, 0xff, 0xff, 0xff, 0x55, ...challenge);
          await udp.send(req2);
          needLoop = true;
        } else {
          needLoop = true; // unexpected first packet — keep listening (legacy Debug.WriteLine path)
        }
      }
    }

    // 3. Player list — may arrive as single packet or splits.
    if (!needLoop) throw new Error("no usable response");
    const packets = new Map<number, Uint8Array>();
    let totalPackets = -1;
    let packetId = 0;

    for (;;) {
      const left = deadlineLeft(start, timeoutMs, clock);
      if (left <= 0) throwTimeout(packets.size, totalPackets, port);
      let buf: Uint8Array;
      try {
        buf = await recvOrTimeout(udp, left);
      } catch (err) {
        if (err instanceof RecvTimeoutError) {
          // Legacy message shape preserved.
          throw new Error(
            totalPackets >= 0
              ? `Timeout waiting for player packets on port ${port}. Received ${packets.size}/${totalPackets} split packets.`
              : "Timeout waiting for full player list.",
          );
        }
        throw err;
      }

      if (u32le(buf, 0) === 0xffffffff) {
        const cmd = buf[4]!;
        if (cmd === 0x44) return parsePlayers(new ByteReader(buf.subarray(5)));
        continue; // unexpected single-packet cmd — ignore (legacy parity)
      }
      if (u32le(buf, 0) === 0xfffffffe) {
        const id = u32le(buf, 4);
        const total = buf[8]!;
        const num = buf[9]!;
        packets.set(num, buf.subarray(12)); // size field not needed — datagram is self-delimiting

        if (totalPackets === -1) {
          totalPackets = total;
          packetId = id;
        }
        if (id !== packetId) continue;

        if (packets.size === totalPackets) {
          // BZIP2 compression flag (high bit of id) — rejected, never decompressed (legacy parity).
          if ((packetId & 0x80000000) !== 0) {
            throw new Error("BZIP2 Compression not supported.");
          }
          const parts: Uint8Array[] = [];
          let len = 0;
          for (let i = 0; i < totalPackets; i++) {
            const p = packets.get(i);
            if (p) {
              parts.push(p);
              len += p.length;
            }
          }
          const combined = new Uint8Array(len);
          let o = 0;
          for (const p of parts) {
            combined.set(p, o);
            o += p.length;
          }
          const cHeader = u32le(combined, 0);
          if (cHeader !== 0xffffffff) {
            throw new Error(`Invalid header in reassembled packet: ${cHeader}`);
          }
          const cmd = combined[4]!;
          if (cmd === 0x44) return parsePlayers(new ByteReader(combined.subarray(5)));
        }
      }
    }
  } finally {
    udp.close();
  }
}

function deadlineLeft(start: number, timeoutMs: number, clock: Clock): number {
  return timeoutMs - (clock.now() - start);
}

async function recvOrTimeout(udp: UdpSocket, leftMs: number): Promise<Uint8Array> {
  let timer: ReturnType<typeof setTimeout> | undefined;
  try {
    return await Promise.race([
      udp.recv(),
      new Promise<never>((_, reject) => {
        timer = setTimeout(() => reject(new RecvTimeoutError()), Math.max(1, leftMs));
      }),
    ]);
  } finally {
    clearTimeout(timer);
  }
}

class RecvTimeoutError extends Error {
  constructor() {
    super("recv timeout");
    this.name = "RecvTimeoutError";
  }
}

function throwTimeout(received: number, total: number, port: number): never {
  // Legacy message shape: "Timeout waiting for player packets on port X. Received n/m split packets."
  throw new Error(
    total >= 0
      ? `Timeout waiting for player packets on port ${port}. Received ${received}/${total} split packets.`
      : "Timeout waiting for full player list.",
  );
}

export { sleep };
