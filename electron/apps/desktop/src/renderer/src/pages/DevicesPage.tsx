/**
 * Devices page (stage 5) — device tree over the active server profile.
 * Parity anchors: MainWindow.Devices.cs tree listing — groups render as expandable rows,
 * leaves show an online/offline/missing state pill, label is Alias > Name (#id).
 */
import { useEffect, useMemo, useState } from "react";
import type * as React from "react";
import { useProfilesStore } from "../stores/profiles.js";
import { activateProfile, type DeviceNode } from "../lib/ipc.js";
import { RulesPanel } from "./RulesPanel.js";
import { TimersPanel } from "./TimersPanel.js";

function StatusPill({ node }: { node: DeviceNode }): React.JSX.Element {
  if (node.isGroup) return <span className="text-xs text-muted-foreground">{node.children.length} dev</span>;
  if (node.isMissing) {
    return <span className="rounded-full bg-destructive/15 px-2 py-0.5 text-[10px] font-medium text-destructive">missing</span>;
  }
  const on = false; // live on/off arrives via conn/push device events (stage 5 wiring next)
  void on;
  return <span className="rounded-full bg-muted px-2 py-0.5 text-[10px] font-medium text-muted-foreground">idle</span>;
}

function DeviceRow({
  node,
  depth,
}: {
  node: DeviceNode;
  depth: number;
}): React.JSX.Element {
  const [expanded, setExpanded] = useState(depth === 0);
  const hasChildren = node.children.length > 0;
  return (
    <div>
      <div
        className="flex items-center gap-2 rounded-md px-2 py-1.5 hover:bg-muted/60"
        style={{ paddingLeft: `${8 + depth * 16}px` }}
        onClick={() => {
          if (hasChildren) setExpanded((e) => !e);
        }}
      >
        {hasChildren ? (
          <span className="w-3 text-[10px] text-muted-foreground" aria-hidden>
            {expanded ? "▾" : "▸"}
          </span>
        ) : (
          <span className="w-3" aria-hidden />
        )}
        <span className="truncate text-sm">
          {node.alias?.trim() || node.name?.trim() || `#${node.entityId}`}
          {!node.isGroup && <span className="ml-1 text-[11px] text-muted-foreground">#{node.entityId}</span>}
        </span>
        <span className="ml-auto flex items-center gap-2">
          {node.kind && <span className="text-[11px] text-muted-foreground">{node.kind}</span>}
          <StatusPill node={node} />
        </span>
      </div>
      {hasChildren && expanded && (
        <div>
          {node.children.map((c) => (
            <DeviceRow key={c.entityId} node={c} depth={depth + 1} />
          ))}
        </div>
      )}
    </div>
  );
}

export function DevicesPage(): React.JSX.Element {
  const profiles = useProfilesStore((s) => s.profiles);
  const activeKey = useProfilesStore((s) => s.activeKey);
  const devicesBy = useProfilesStore((s) => s.devices);
  const loading = useProfilesStore((s) => s.loading);
  const error = useProfilesStore((s) => s.error);
  const loadProfiles = useProfilesStore((s) => s.loadProfiles);
  const selectProfile = useProfilesStore((s) => s.selectProfile);
  // Devices ↔ Rules ↔ Timers sub-view within the tab (legacy hosts all in the devices window).
  const [view, setView] = useState<"devices" | "rules" | "timers">("devices");

  useEffect(() => {
    void loadProfiles();
  }, [loadProfiles]);

  const active = useMemo(() => profiles.find((p) => p.matchKey === activeKey) ?? null, [profiles, activeKey]);
  const devices = activeKey ? (devicesBy[activeKey] ?? []) : [];

  if (profiles.length === 0 && !loading && !error) {
    return (
      <div className="flex h-full items-center justify-center p-6">
        <div className="max-w-sm text-center">
          <h2 className="text-lg font-semibold">No server profiles</h2>
          <p className="mt-1 text-sm text-muted-foreground">
            Pair a Rust+ server to create one. Legacy profiles import via the migration flow.
          </p>
        </div>
      </div>
    );
  }

  return (
    <div className="flex h-full min-h-0 flex-col">
      <header className="flex items-center gap-3 border-b px-4 py-2.5">
        <h1 className="text-sm font-semibold">Devices</h1>
        {profiles.length > 0 && (
          <select
            aria-label="Active profile"
            value={activeKey ?? ""}
            onChange={(e) => {
              selectProfile(e.target.value);
              void activateProfile(e.target.value);
            }}
            className="rounded-md border bg-transparent px-2 py-1 text-xs"
          >
            {profiles.map((p) => (
              <option key={p.matchKey} value={p.matchKey}>
                {p.name} ({p.host}:{p.port})
              </option>
            ))}
          </select>
        )}
        <div className="ml-auto flex items-center gap-2">
          {activeKey && (
            <div className="flex rounded-md border p-0.5 text-xs">
              {(["devices", "rules", "timers"] as const).map((v) => (
                <button
                  key={v}
                  type="button"
                  onClick={() => setView(v)}
                  className={
                    view === v
                      ? "rounded bg-primary/10 px-2 py-1 font-medium capitalize text-primary"
                      : "rounded px-2 py-1 capitalize text-muted-foreground hover:bg-muted"
                  }
                >
                  {v}
                </button>
              ))}
            </div>
          )}
          {active && view === "devices" && (
            <span className="text-xs text-muted-foreground">
              {active.deviceCount} device{active.deviceCount === 1 ? "" : "s"}
            </span>
          )}
        </div>
      </header>

      {error && <div className="px-4 py-2 text-xs text-destructive">{error}</div>}

      {view === "rules" && activeKey ? (
        <div className="min-h-0 flex-1 overflow-y-auto">
          <RulesPanel matchKey={activeKey} />
        </div>
      ) : view === "timers" && activeKey ? (
        <TimersPanel matchKey={activeKey} />
      ) : (
        <div className="min-h-0 flex-1 overflow-y-auto p-2">
          {loading && profiles.length === 0 ? (
            <div className="p-4 text-sm text-muted-foreground">Loading profiles…</div>
          ) : (
            devices.map((d) => <DeviceRow key={d.entityId} node={d} depth={0} />)
          )}
        </div>
      )}
    </div>
  );
}
