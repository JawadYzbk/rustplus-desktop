import { useEffect } from "react";
import type * as React from "react";
import { Navigate, Route, Routes, useParams, useLocation } from "react-router-dom";
import { TitleBar } from "./components/shell/TitleBar.js";
import { SidebarRail } from "./components/shell/SidebarRail.js";
import { FeaturePending } from "./components/shell/FeaturePending.js";
import { useUiStore, WORKSPACE_TABS, type RailTab } from "./stores/ui.js";
import { useProfilesStore } from "./stores/profiles.js";
import { getUiPrefs, setUiPrefsDebounced } from "./lib/ipc.js";
import { MigratePage } from "./pages/MigratePage.js";
import { DevicesPage } from "./pages/DevicesPage.js";
import { RaidCalculatorPage } from "./pages/RaidCalculatorPage.js";
import { RecyclerCalculatorPage } from "./pages/RecyclerCalculatorPage.js";
import { PlayerWipeTrackerPage } from "./pages/PlayerWipeTrackerPage.js";
import { DeathStatsPage } from "./pages/DeathStatsPage.js";
import { PairingPage } from "./pages/PairingPage.js";

/**
 * Shell layout (audit UI_SHELL §2): titlebar row; below it the icon rail + content column where the right
 * side is reserved for the permanent live-map pane (map lands in stage 6 — the split is scaffolded now).
 * Workspace tabs take over the full content area and restore `lastContentRoute` on close.
 */
export function App(): React.JSX.Element {
  const activeTab = useUiStore((s) => s.activeTab);
  const setActiveTab = useUiStore((s) => s.setActiveTab);
  const location = useLocation();

  // URL → store sync so deep links / restores land on the right tab.
  useEffect(() => {
    const match = /^\/tab\/([a-z-]+)$/.exec(location.pathname);
    if (match?.[1]) {
      const tab = match[1] as RailTab;
      if (tab !== activeTab) setActiveTab(tab);
    }
  }, [location.pathname]); // eslint-disable-line react-hooks/exhaustive-deps

  useEffect(() => {
    void window.rpd
      .getInfo()
      .then((r) => {
        if (r.ok) void window.rpd.log("info", "shell", `boot ok (v${String((r.data as { version: string }).version)})`);
        else void window.rpd.log("error", "shell", `getInfo failed: ${r.error.message}`);
      })
      .catch(() => undefined);

    // Hydrate persisted shell prefs, then persist subsequent changes (debounced).
    let lastPersisted = "";
    void getUiPrefs().then((prefs) => {
      if (prefs) {
        useUiStore.setState({ sidebarPinned: prefs.sidebarPinned, sidebarWidth: prefs.sidebarWidth });
        lastPersisted = JSON.stringify(prefs);
      }
    });
    const unsub = useUiStore.subscribe((state) => {
      const snapshot = JSON.stringify({ sidebarPinned: state.sidebarPinned, sidebarWidth: state.sidebarWidth });
      if (snapshot !== lastPersisted) {
        lastPersisted = snapshot;
        const parsed = JSON.parse(snapshot) as { sidebarPinned?: boolean; sidebarWidth?: number };
        setUiPrefsDebounced(parsed);
      }
    });
    const unsubPush = window.rpd.onPush(({ stream, event }) => {
      const e = (event ?? {}) as { kind?: string; entityId?: unknown; on?: unknown };
      if (stream === "device" && e.kind === "deviceState" && typeof e.entityId === "number" && typeof e.on === "boolean") {
        useProfilesStore.getState().setDeviceState(e.entityId, e.on);
      }
    });
    return () => {
      unsub();
      unsubPush();
    };
  }, []);

  return (
    <div className="flex h-full flex-col">
      <TitleBar />
      <div className="flex min-h-0 flex-1">
        <SidebarRail />
        <main className="min-w-0 flex-1 bg-background">
          <Routes>
            <Route path="/" element={<Navigate to="/tab/devices" replace />} />
            <Route path="/tab/:tabId" element={<TabOutlet />} />
            {/* Migration UX route (stage 3) — reachable via deep link /migrate; not on the rail. */}
            <Route path="/migrate" element={<MigratePage />} />
            <Route path="/pair" element={<PairingPage />} />
            <Route path="*" element={<Navigate to="/tab/devices" replace />} />
          </Routes>
        </main>
      </div>
    </div>
  );
}

function TabOutlet(): React.JSX.Element {
  const { tabId } = useParams<{ tabId: string }>();
  switch (tabId) {
    case "devices":
      return <DevicesPage />;
    case "team":
      return <FeaturePending title="Team" stage="4+" matrix="3.10, 9.6" />;
    case "clan":
      return <FeaturePending title="Clan" stage="4+" matrix="3.10" />;
    case "cameras":
      return <FeaturePending title="Cameras" stage="4/12" matrix="3.13" />;
    case "players":
      return <FeaturePending title="Players" stage="4/12" matrix="3.12, 5.5" />;
    case "notifications":
      return <FeaturePending title="Notifications" stage="12" matrix="10.9" />;
    case "genetics-lab":
      return <FeaturePending title="GeneticsLab" stage="8" matrix="7.1–7.5" />;
    case "wipe-tracker":
      return <PlayerWipeTrackerPage />;
    case "death-stats":
      return <DeathStatsPage />;
    case "raid-calculator":
      return <RaidCalculatorPage />;
    case "recycler-calculator":
      return <RecyclerCalculatorPage />;
    default:
      return <Navigate to="/tab/devices" replace />;
  }
}

// Referenced by the shell badge; keeps the workspace-set import meaningful until takeover routes land.
export const isWorkspaceTab = (tab: RailTab): boolean => WORKSPACE_TABS.has(tab);
