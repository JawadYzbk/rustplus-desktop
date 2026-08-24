import { APP_VERSION } from "@rpd/shared";
import { z } from "zod";
import { CloudApiClient, CloudApiError } from "./cloud-api-client.js";
import { CloudSessionStore, type CloudSession, type CloudSessionUser } from "../../stores/cloud-session-store.js";

const apiUserSchema = z.object({
  id: z.union([z.string(), z.number()]),
  steam_id: z.string().nullable().optional(),
  name: z.string().nullable().optional(),
  display_name: z.string().nullable().optional(),
  email: z.string().nullable().optional(),
  providers: z.array(z.string()).optional(),
  has_password: z.boolean().optional(),
});

export interface PlayerWipeCapabilities {
  planCode: string;
  isTrackerAvailable: boolean;
  canTrackTeam: boolean;
  canUseCloudSync: boolean;
  canUseAdvancedViews: boolean;
  canUseRouteReplay: boolean;
  canExport: boolean;
  maxTrackedPlayers: number;
  retainedWipes: number;
  cloudRetentionDays: number;
  fetchedAt: string;
}

export interface CloudBootstrapResult {
  signedIn: boolean;
  user: CloudSessionUser | null;
  capabilities: PlayerWipeCapabilities | null;
  error: string | null;
}

export class CloudService {
  private session: CloudSession | null;
  private lastCapabilities: PlayerWipeCapabilities | null = null;
  private readonly client: CloudApiClient;

  constructor(
    private readonly sessions: CloudSessionStore,
    baseUrl: string,
    fetchImpl?: typeof fetch,
  ) {
    this.session = sessions.load();
    this.client = new CloudApiClient({
      baseUrl,
      clientVersion: APP_VERSION,
      token: () => this.session?.token ?? null,
      fetchImpl,
      onUnauthorized: () => this.logout(),
    });
  }

  get trackerCapabilities(): PlayerWipeCapabilities | null { return this.lastCapabilities; }
  get api(): CloudApiClient { return this.client; }

  async login(email: string, password: string): Promise<{ signedIn: true; user: CloudSessionUser }> {
    const payload = await this.client.request<unknown>("POST", "auth/token", {
      json: { email, password, device_name: "RustPlusDesk Electron" },
    });
    const parsed = z.object({
      token: z.string().min(1),
      expires_at: z.string().nullable().optional(),
      user: apiUserSchema,
    }).parse(payload);
    const user = mapUser(parsed.user);
    this.session = { token: parsed.token, user, expiresAt: parsed.expires_at ?? null };
    this.sessions.save(this.session);
    return { signedIn: true, user };
  }

  async bootstrap(): Promise<CloudBootstrapResult> {
    if (!this.session) return { signedIn: false, user: null, capabilities: null, error: null };
    try {
      const payload = await this.client.request<unknown>("GET", "client/bootstrap");
      this.lastCapabilities = parseCapabilities(payload);
      return {
        signedIn: true,
        user: this.session.user,
        capabilities: this.lastCapabilities,
        error: null,
      };
    } catch (error) {
      if (error instanceof CloudApiError && (error.status === 401 || error.status === 403)) {
        return { signedIn: false, user: null, capabilities: null, error: error.reason };
      }
      return {
        signedIn: true,
        user: this.session.user,
        capabilities: null,
        error: error instanceof Error ? error.message : String(error),
      };
    }
  }

  logout(): void {
    this.session = null;
    this.lastCapabilities = null;
    this.sessions.clear();
  }
}

function mapUser(user: z.infer<typeof apiUserSchema>): CloudSessionUser {
  return {
    id: String(user.id),
    steamId: user.steam_id ?? null,
    name: user.name ?? null,
    displayName: user.display_name ?? null,
    email: user.email ?? null,
    providers: user.providers ?? [],
    hasPassword: user.has_password ?? false,
  };
}

function parseCapabilities(payload: unknown): PlayerWipeCapabilities {
  const data = record(payload);
  const limits = record(data["limits"]);
  const tracker = record(limits["player_wipe_tracker"]);
  const enabled = (key: string): boolean => record(tracker[key])["enabled"] === true;
  const value = (key: string, fallback: number): number => {
    const candidate = record(tracker[key])["value"];
    return typeof candidate === "number" && Number.isFinite(candidate) ? Math.max(0, Math.trunc(candidate)) : fallback;
  };
  return {
    planCode: typeof data["plan_code"] === "string" && data["plan_code"] ? data["plan_code"] : "free",
    isTrackerAvailable: enabled("access"),
    canTrackTeam: enabled("team_tracking"),
    canUseCloudSync: enabled("cloud_sync"),
    canUseAdvancedViews: enabled("advanced_views"),
    canUseRouteReplay: enabled("route_replay"),
    canExport: enabled("export"),
    maxTrackedPlayers: Math.max(1, value("max_tracked_players", 1)),
    retainedWipes: Math.max(1, value("retained_wipes", 1)),
    cloudRetentionDays: value("cloud_retention_days", 0),
    fetchedAt: new Date().toISOString(),
  };
}

function record(value: unknown): Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value) ? value as Record<string, unknown> : {};
}
