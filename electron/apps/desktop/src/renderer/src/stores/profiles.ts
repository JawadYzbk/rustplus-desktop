/**
 * Profiles store (renderer) — mirrors the C# _vm.Selected semantics: one active profile,
 * its device tree cached per match key, edits mutate the tree in place and persist
 * whole-tree (legacy Save() parity).
 */
import { create } from "zustand";
import { getDevices, listProfiles, saveDevices, type DeviceNode, type ProfileSummary } from "../lib/ipc.js";

/** Recursive find over the tree (main-side findDeviceById mirror for optimistic UI). */
export function findNode(devices: readonly DeviceNode[], entityId: number): DeviceNode | null {
  for (const d of devices) {
    if (d.entityId === entityId) return d;
    const hit = findNode(d.children, entityId);
    if (hit) return hit;
  }
  return null;
}

/** Alias > Name > #ID display parity (DeviceImportItem.DisplayName / legacy tree labels). */
export function deviceLabel(d: Pick<DeviceNode, "alias" | "name" | "entityId">): string {
  const label = d.alias && d.alias.trim().length > 0 ? d.alias : d.name && d.name.trim().length > 0 ? d.name : null;
  return label ? label : `#${d.entityId}`;
}

interface ProfilesState {
  profiles: ProfileSummary[];
  activeKey: string | null;
  devices: Record<string, DeviceNode[]>;
  loading: boolean;
  error: string | null;
  loadProfiles: () => Promise<void>;
  selectProfile: (matchKey: string | null) => void;
  ensureDevices: (matchKey: string) => Promise<void>;
  /** In-place device mutation + whole-tree persist; returns false when the node is gone. */
  updateDevice: (matchKey: string, entityId: number, patch: Partial<DeviceNode>) => Promise<boolean>;
}

export const useProfilesStore = create<ProfilesState>((set, get) => ({
  profiles: [],
  activeKey: null,
  devices: {},
  loading: false,
  error: null,

  loadProfiles: async () => {
    set({ loading: true, error: null });
    try {
      const profiles = await listProfiles();
      const prevActive = get().activeKey;
      const activeKey =
        prevActive && profiles.some((p) => p.matchKey === prevActive)
          ? prevActive
          : (profiles[0]?.matchKey ?? null);
      set({ profiles, activeKey, loading: false });
      if (activeKey) await get().ensureDevices(activeKey);
    } catch (err) {
      set({ loading: false, error: err instanceof Error ? err.message : String(err) });
    }
  },

  selectProfile: (matchKey) => {
    set({ activeKey: matchKey });
    if (matchKey) void get().ensureDevices(matchKey);
  },

  ensureDevices: async (matchKey) => {
    if (get().devices[matchKey]) return;
    const devices = await getDevices(matchKey);
    if (devices !== null) set((s) => ({ devices: { ...s.devices, [matchKey]: devices } }));
  },

  updateDevice: async (matchKey, entityId, patch) => {
    const tree = get().devices[matchKey];
    if (!tree) return false;
    const node = findNode(tree, entityId);
    if (!node) return false;
    Object.assign(node, patch);
    set((s) => ({ devices: { ...s.devices, [matchKey]: [...tree] } })); // new top-level array → re-render
    return saveDevices(matchKey, tree);
  },
}));
