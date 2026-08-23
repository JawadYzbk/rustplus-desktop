/**
 * `profile/*` channel handlers — bridges the lossless ProfilesStore to typed renderer DTOs.
 * Device conversion goes through server-profile.ts's dual-casing parser and PascalCase
 * serializer so files on disk keep their legacy shape while the bridge speaks camelCase.
 */
import type { DeviceNodeDto } from "@rpd/shared";
import { parseDevices, serializeDevices } from "./services/devices/server-profile.js";
import { countActualDevicesTree, findDeviceById, type SmartDeviceNode } from "./services/devices/device-data.js";
import type { ProfilesStore } from "./stores/profiles-store.js";

export interface ProfileHandlersDeps {
  profiles: Pick<ProfilesStore, "list" | "matchKey" | "devicesFor" | "saveDevices">;
}

/** Counts actual (non-group) devices recursively — legacy CountActualDevices parity. */
function deviceCount(devices: SmartDeviceNode[]): number {
  return countActualDevicesTree(devices);
}

export function buildProfileHandlers(deps: ProfileHandlersDeps): {
  "profile/list": () => { profiles: Array<{ matchKey: string; name: string; host: string; port: number; steamId64: string; deviceCount: number }> };
  "profile/getDevices": (req: { matchKey: string }) => { devices: DeviceNodeDto[]; found: boolean };
  "profile/saveDevices": (req: { matchKey: string; devices: SmartDeviceNode[] }) => { saved: boolean };
} {
  return {
    "profile/list": () => ({
      profiles: deps.profiles.list().map((p) => {
        const key = deps.profiles.matchKey(p);
        const raw = deps.profiles.devicesFor(key) ?? [];
        return {
          matchKey: key,
          name: p.Name,
          host: p.Host,
          port: p.Port,
          steamId64: p.SteamId64,
          deviceCount: deviceCount(parseDevices(raw)),
        };
      }),
    }),

    "profile/getDevices": (req) => {
      const raw = deps.profiles.devicesFor(req.matchKey);
      if (raw === null) return { devices: [], found: false };
      return { devices: parseDevices(raw) as unknown as DeviceNodeDto[], found: true };
    },

    "profile/saveDevices": (req) => ({
      saved: deps.profiles.saveDevices(req.matchKey, serializeDevices(req.devices)),
    }),
  };
}

// findDeviceById is re-exported for renderer-adjacent helpers/tests that resolve nodes after IPC round-trips.
export { findDeviceById };
