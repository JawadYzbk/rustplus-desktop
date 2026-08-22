/**
 * UI shell state (parity anchors: audit UI_SHELL §1–2).
 *
 * - Sidebar: 64px icon rail collapsed ↔ expanded panel; hover-expand after 200 ms with 180 ms width
 *   animation and a pin toggle. Width persistence lands with the settings store in stage 3.
 * - Workspace tabs (GeneticsLab, PlayerWipeTracker, DeathStats, RaidCalculator, Recycler) take over the full
 *   content area and return to `_lastWorkspaceTabIndex` on close — replicated via `lastContentRoute`.
 */
import { create } from "zustand";

export type RailTab =
  | "devices"
  | "team"
  | "clan"
  | "cameras"
  | "players"
  | "notifications"
  | "genetics-lab"
  | "wipe-tracker"
  | "death-stats"
  | "raid-calculator"
  | "recycler-calculator";

/** Tabs whose real views are full-window takeover panels in the legacy app. */
export const WORKSPACE_TABS: ReadonlySet<RailTab> = new Set<RailTab>([
  "genetics-lab",
  "wipe-tracker",
  "death-stats",
  "raid-calculator",
  "recycler-calculator",
]);

interface UiState {
  sidebarExpanded: boolean;
  sidebarPinned: boolean;
  /** Expanded width in px; legacy range 360–480, default 420 (SidebarWidth=420). */
  sidebarWidth: number;
  activeTab: RailTab;
  /** Route to restore when a workspace takeover closes. */
  lastContentRoute: RailTab;
  setSidebarExpanded: (expanded: boolean) => void;
  togglePinned: () => void;
  setSidebarWidth: (width: number) => void;
  setActiveTab: (tab: RailTab) => void;
  closeWorkspace: () => void;
}

export const useUiStore = create<UiState>((set, get) => ({
  sidebarExpanded: false,
  sidebarPinned: true,
  sidebarWidth: 420,
  activeTab: "devices",
  lastContentRoute: "devices",
  setSidebarExpanded: (expanded) => set({ sidebarExpanded: expanded }),
  togglePinned: () => set({ sidebarPinned: !get().sidebarPinned }),
  setSidebarWidth: (width) => set({ sidebarWidth: Math.min(480, Math.max(360, width)) }),
  setActiveTab: (tab) =>
    set((state) => ({
      activeTab: tab,
      // Non-workspace navigation always records the fallback route for workspace close semantics.
      lastContentRoute: WORKSPACE_TABS.has(tab)
        ? WORKSPACE_TABS.has(state.activeTab)
          ? state.lastContentRoute
          : state.activeTab
        : tab,
    })),
  closeWorkspace: () => set({ activeTab: get().lastContentRoute }),
}));
