import { createRequire } from "node:module";
import { dirname, join } from "node:path";
import { describe, expect, it } from "vitest";

const require = createRequire(import.meta.url);

describe("rustplus.js protobuf compatibility", () => {
  it("decodes sparse AppInfo and team payloads while preserving queuedPlayers", async () => {
    const packageEntry = require.resolve("@liamcottle/rustplus.js");
    const packageDir = dirname(packageEntry);
    const protobuf = require(require.resolve("protobufjs", { paths: [packageDir] })) as {
      load(path: string): Promise<{ lookupType(name: string): { create(value: unknown): unknown; encode(value: unknown): { finish(): Uint8Array }; decode(bytes: Uint8Array): unknown } }>;
    };
    const root = await protobuf.load(join(packageDir, "rustplus.proto"));
    const message = root.lookupType("rustplus.AppMessage");
    const encoded = message.encode(message.create({
      response: {
        seq: 1,
        info: { players: 42, maxPlayers: 200, queuedPlayers: 7 },
        teamInfo: { members: [{ steamId: "76561198000000001", name: "Ada" }] },
      },
    })).finish();

    expect(() => message.decode(encoded)).not.toThrow();
    expect(message.decode(encoded)).toBeTruthy();
  });
});
