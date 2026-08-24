import { describe, expect, it, vi } from "vitest";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { CloudApiClient, CloudApiError } from "../src/main/services/cloud/cloud-api-client.js";
import { CloudService } from "../src/main/services/cloud/cloud-service.js";
import { CloudSessionStore } from "../src/main/stores/cloud-session-store.js";
import { PassthroughSecretCodec } from "../src/main/stores/secret-codec.js";

describe("CloudApiClient", () => {
  it("adds Laravel headers, repeated query values, and unwraps data", async () => {
    const fetchImpl = vi.fn<typeof fetch>().mockResolvedValue(new Response(JSON.stringify({ data: { ok: true } }), { status: 200 }));
    const client = new CloudApiClient({ baseUrl: "https://cloud.test/api/", clientVersion: "8.1.0", token: () => "secret", fetchImpl });
    await expect(client.request("GET", "/client/bootstrap", { query: { "steam_ids[]": ["1", "2"] } })).resolves.toEqual({ ok: true });
    const [url, init] = fetchImpl.mock.calls[0]!;
    expect(String(url)).toBe("https://cloud.test/api/v1/client/bootstrap?steam_ids%5B%5D=1&steam_ids%5B%5D=2");
    expect((init?.headers as Headers).get("Authorization")).toBe("Bearer secret");
    expect((init?.headers as Headers).get("X-Client-Version")).toBe("8.1.0");
  });

  it("preserves safe Laravel validation errors and unauthorized callbacks", async () => {
    const onUnauthorized = vi.fn();
    const fetchImpl = vi.fn<typeof fetch>().mockResolvedValue(new Response(JSON.stringify({ message: "nope", errors: { email: ["invalid email"] } }), { status: 422 }));
    const client = new CloudApiClient({ baseUrl: "https://cloud.test/api", clientVersion: "8.1.0", fetchImpl, onUnauthorized });
    await expect(client.request("POST", "auth/token", { json: {} })).rejects.toMatchObject({ status: 422, reason: "invalid email" });
    expect(onUnauthorized).not.toHaveBeenCalled();

    fetchImpl.mockResolvedValueOnce(new Response(JSON.stringify({ message: "expired" }), { status: 401 }));
    await expect(client.request("GET", "me")).rejects.toBeInstanceOf(CloudApiError);
    expect(onUnauthorized).toHaveBeenCalledOnce();
  });
});

describe("CloudService", () => {
  it("logs in, persists the bearer, and maps Player Wipe Tracker limits", async () => {
    const fetchImpl = vi.fn<typeof fetch>()
      .mockResolvedValueOnce(new Response(JSON.stringify({ data: {
        token: "token-1",
        user: { id: "u1", email: "a@example.com", providers: ["email"], has_password: true },
        expires_at: null,
      } }), { status: 200 }))
      .mockResolvedValueOnce(new Response(JSON.stringify({ data: {
        plan_code: "supporter",
        limits: { player_wipe_tracker: {
          access: { enabled: true }, team_tracking: { enabled: true }, cloud_sync: { enabled: true },
          advanced_views: { enabled: false }, route_replay: { enabled: false }, export: { enabled: true },
          max_tracked_players: { value: 8 }, retained_wipes: { value: 4 }, cloud_retention_days: { value: 30 },
        } },
      } }), { status: 200 }));
    const directory = mkdtempSync(join(tmpdir(), "rpd-cloud-"));
    const store = new CloudSessionStore(directory, new PassthroughSecretCodec());
    const service = new CloudService(store, "https://cloud.test/api", fetchImpl);

    expect((await service.bootstrap()).signedIn).toBe(false);
    await expect(service.login("a@example.com", "password")).resolves.toMatchObject({ signedIn: true });
    const result = await service.bootstrap();
    expect(result.capabilities).toMatchObject({ planCode: "supporter", canTrackTeam: true, maxTrackedPlayers: 8, retainedWipes: 4 });
    expect((fetchImpl.mock.calls[1]?.[1]?.headers as Headers).get("Authorization")).toBe("Bearer token-1");
    service.logout();
    expect(store.load()).toBeNull();
    rmSync(directory, { recursive: true, force: true });
  });
});
