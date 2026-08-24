/**
 * Devices page (stage 5) — device tree over the active server profile.
 * Parity anchors: MainWindow.Devices.cs tree listing — groups render as expandable rows,
 * leaves show an online/offline/missing state pill, label is Alias > Name (#id).
 */
import { useEffect, useMemo, useState } from "react";
import type * as React from "react";
import { useNavigate } from "react-router-dom";
import { useProfilesStore } from "../stores/profiles.js";
import { findNode } from "../stores/profiles.js";
import {
  activateProfile,
  applyImportedDevices,
  deleteDevice,
  exportDevices,
  importDevicesPreview,
  type DeviceImportCandidate,
  type DeviceNode,
} from "../lib/ipc.js";
import { RulesPanel } from "./RulesPanel.js";
import { TimersPanel } from "./TimersPanel.js";
import { DeviceAutomationPanel } from "./DeviceAutomationPanel.js";
import { Badge } from "../components/ui/badge.js";
import { Button } from "../components/ui/button.js";
import { Card } from "../components/ui/card.js";
import { Checkbox } from "../components/ui/checkbox.js";
import { Dialog, DialogContent, DialogDescription, DialogFooter, DialogHeader, DialogTitle } from "../components/ui/dialog.js";
import { Input } from "../components/ui/input.js";
import { Label } from "../components/ui/label.js";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "../components/ui/select.js";
import { useConnectionStore } from "../stores/connection.js";

function StatusPill({ node }: { node: DeviceNode }): React.JSX.Element {
  if (node.isGroup) return <span className="text-xs text-muted-foreground">{node.children.length} dev</span>;
  if (node.isMissing) {
    return <Badge variant="destructive" className="rounded-full px-2 py-0.5 text-[10px]">missing</Badge>;
  }
  return <Badge variant="secondary" className="rounded-full px-2 py-0.5 text-[10px] text-muted-foreground">idle</Badge>;
}

function DeviceRow({
  node,
  depth,
  liveState,
  deviceStates,
  selectedId,
  onSelect,
}: {
  node: DeviceNode;
  depth: number;
  liveState: boolean | undefined;
  deviceStates: Readonly<Record<number, boolean>>;
  selectedId: number | null;
  onSelect: (entityId: number) => void;
}): React.JSX.Element {
  const [expanded, setExpanded] = useState(depth === 0);
  const hasChildren = node.children.length > 0;
  return (
    <div>
      <Button
        type="button"
        variant="ghost"
        className={`h-auto w-full justify-start whitespace-normal rounded-md px-2 py-1.5 ${selectedId === node.entityId ? "bg-primary/10 ring-1 ring-primary/40" : "hover:bg-muted/60"}`}
        style={{ paddingLeft: `${8 + depth * 16}px` }}
        onClick={() => {
          onSelect(node.entityId);
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
          {node.isGroup ? (
            <StatusPill node={node} />
          ) : node.isMissing ? (
            <StatusPill node={node} />
          ) : liveState === undefined ? (
            <span className="rounded-full bg-muted px-2 py-0.5 text-[10px] font-medium text-muted-foreground">unknown</span>
          ) : (
            <span className={liveState ? "rounded-full bg-emerald-500/15 px-2 py-0.5 text-[10px] font-medium text-emerald-400" : "rounded-full bg-slate-500/15 px-2 py-0.5 text-[10px] font-medium text-slate-300"}>
              {liveState ? "ON" : "OFF"}
            </span>
          )}
        </span>
      </Button>
      {hasChildren && expanded && (
        <div>
          {node.children.map((c) => (
            <DeviceRow
              key={c.entityId}
              node={c}
              depth={depth + 1}
              liveState={deviceStates[c.entityId]}
              deviceStates={deviceStates}
              selectedId={selectedId}
              onSelect={onSelect}
            />
          ))}
        </div>
      )}
    </div>
  );
}

function filterTree(nodes: readonly DeviceNode[], query: string, kind: string): DeviceNode[] {
  if (query === "" && kind === "all") return [...nodes];
  const q = query.toLowerCase();
  return nodes.flatMap((node) => {
    const children = node.isGroup ? filterTree(node.children, query, kind) : [];
    const label = `${node.alias ?? ""} ${node.name ?? ""} ${node.entityId}`.toLowerCase();
    const matches = label.includes(q) && (kind === "all" || node.kind === kind);
    if (!matches && children.length === 0) return [];
    return [{ ...node, children: node.isGroup && matches ? node.children : children }];
  });
}

function allKinds(nodes: readonly DeviceNode[]): string[] {
  return nodes.flatMap((node) => [
    ...(node.kind ? [node.kind] : []),
    ...allKinds(node.children),
  ]);
}

type ImportCandidateState = DeviceImportCandidate & { selected: boolean };

export function DevicesPage(): React.JSX.Element {
  const navigate = useNavigate();
  const profiles = useProfilesStore((s) => s.profiles);
  const activeKey = useProfilesStore((s) => s.activeKey);
  const devicesBy = useProfilesStore((s) => s.devices);
  const deviceStates = useProfilesStore((s) => s.deviceStates);
  const loading = useProfilesStore((s) => s.loading);
  const error = useProfilesStore((s) => s.error);
  const loadProfiles = useProfilesStore((s) => s.loadProfiles);
  const selectProfile = useProfilesStore((s) => s.selectProfile);
  const reloadDevices = useProfilesStore((s) => s.reloadDevices);
  const connection = useConnectionStore((s) => s.snapshot);
  const connectionPhase = useConnectionStore((s) => s.phase);
  const connectionError = useConnectionStore((s) => s.error);
  const connectServer = useConnectionStore((s) => s.connectProfile);
  const disconnectServer = useConnectionStore((s) => s.disconnect);
  // Devices ↔ Rules ↔ Timers sub-view within the tab (legacy hosts all in the devices window).
  const [view, setView] = useState<"devices" | "rules" | "timers" | "automation">("devices");
  const [query, setQuery] = useState("");
  const [kind, setKind] = useState("all");
  const [selectedId, setSelectedId] = useState<number | null>(null);
  const [importCandidates, setImportCandidates] = useState<ImportCandidateState[] | null>(null);
  const [actionMessage, setActionMessage] = useState<string | null>(null);

  useEffect(() => {
    void loadProfiles();
  }, [loadProfiles]);

  const active = useMemo(() => profiles.find((p) => p.matchKey === activeKey) ?? null, [profiles, activeKey]);
  const devices = activeKey ? (devicesBy[activeKey] ?? []) : [];
  const kinds = useMemo(
    () => ["all", ...new Set(allKinds(devices))],
    [devices],
  );
  const filteredDevices = useMemo(() => filterTree(devices, query.trim().toLowerCase(), kind), [devices, query, kind]);
  const selectedNode = selectedId === null ? null : findNode(devices, selectedId);

  const handleProfileChange = async (matchKey: string): Promise<void> => {
    selectProfile(matchKey);
    if (await activateProfile(matchKey)) await connectServer(matchKey);
  };

  const handleExport = async (): Promise<void> => {
    if (!activeKey) return;
    try {
      const result = await exportDevices(activeKey);
      setActionMessage(result.saved ? `Exported devices to ${result.path ?? "JSON"}.` : null);
    } catch (err: unknown) {
      setActionMessage(err instanceof Error ? err.message : String(err));
    }
  };

  const handleImportPreview = async (): Promise<void> => {
    if (!activeKey) return;
    try {
      const result = await importDevicesPreview(activeKey);
      if (!result.canceled) {
        setImportCandidates(result.candidates.map((candidate) => ({
          ...candidate,
          selected: !candidate.alreadyPresent && !candidate.fromPreviousWipe,
        })));
      }
    } catch (err: unknown) {
      setActionMessage(err instanceof Error ? err.message : String(err));
    }
  };

  const handleImportApply = async (): Promise<void> => {
    if (!activeKey || !importCandidates) return;
    try {
      const result = await applyImportedDevices(activeKey, importCandidates.filter((candidate) => candidate.selected).map((candidate) => candidate.originalDto));
      if (result.saved) await reloadDevices(activeKey);
      setActionMessage(`Imported ${result.imported} device${result.imported === 1 ? "" : "s"}.`);
      setImportCandidates(null);
    } catch (err: unknown) {
      setActionMessage(err instanceof Error ? err.message : String(err));
    }
  };

  const handleDelete = async (): Promise<void> => {
    if (!activeKey || !selectedNode || selectedNode.isGroup || !selectedNode.isMissing) return;
    if (!window.confirm(`Delete missing device #${selectedNode.entityId}?`)) return;
    try {
      const result = await deleteDevice(activeKey, selectedNode.entityId);
      if (result.removed) {
        await reloadDevices(activeKey);
        setSelectedId(null);
        setActionMessage("Missing device removed.");
      } else setActionMessage("Only missing leaf devices can be removed.");
    } catch (err: unknown) {
      setActionMessage(err instanceof Error ? err.message : String(err));
    }
  };

  if (profiles.length === 0 && !loading && !error) {
    return (
      <div className="flex h-full items-center justify-center p-6">
        <Card className="max-w-sm p-6 text-center">
          <h2 className="text-lg font-semibold">No server profiles</h2>
          <p className="mt-1 text-sm text-muted-foreground">
            Pair a Rust+ server or import your existing RustPlusDesk data to get started.
          </p>
          <div className="mt-4 flex justify-center gap-2">
            <Button type="button" size="sm" onClick={() => navigate("/pair")}>Pair Rust+</Button>
            <Button type="button" size="sm" variant="outline" onClick={() => navigate("/migrate")}>Import old data</Button>
          </div>
        </Card>
      </div>
    );
  }

  return (
    <div className="flex h-full min-h-0 flex-col">
      <header className="flex items-center gap-3 border-b px-4 py-2.5">
        <h1 className="text-sm font-semibold">Devices</h1>
        {profiles.length > 0 && (
          <Select value={activeKey ?? ""} onValueChange={(value) => void handleProfileChange(value)}>
            <SelectTrigger aria-label="Active profile" className="w-56 text-xs"><SelectValue placeholder="Select server" /></SelectTrigger>
            <SelectContent>{profiles.map((p) => <SelectItem key={p.matchKey} value={p.matchKey}>{p.name} ({p.host}:{p.port})</SelectItem>)}</SelectContent>
          </Select>
        )}
        <div className="ml-auto flex items-center gap-2">
          {active && <>
            <Badge variant={connection.connected ? "default" : "secondary"} className="gap-1.5">
              <span className={`h-1.5 w-1.5 rounded-full ${connection.connected ? "bg-emerald-300" : "bg-muted-foreground"}`} />
              {connectionPhase === "reconnecting" ? "Reconnecting" : connection.connected ? "Connected" : "Offline"}
            </Badge>
            {connection.connected ? (
              <Button type="button" variant="outline" size="sm" onClick={() => void disconnectServer()}>Disconnect</Button>
            ) : (
              <Button type="button" size="sm" disabled={connectionPhase === "connecting"} onClick={() => void handleProfileChange(active.matchKey)}>
                {connectionPhase === "connecting" ? "Connecting…" : "Connect"}
              </Button>
            )}
          </>}
          {activeKey && (
            <div className="flex rounded-md border p-0.5 text-xs">
              {(["devices", "rules", "timers", "automation"] as const).map((v) => (
                <Button
                  key={v}
                  type="button"
                  variant={view === v ? "secondary" : "ghost"}
                  size="sm"
                  onClick={() => setView(v)}
                  className={`capitalize ${view === v ? "font-medium text-primary" : "text-muted-foreground"}`}
                >
                  {v}
                </Button>
              ))}
            </div>
          )}
          {active && view === "devices" && (
            <>
              <Input
                value={query}
                onChange={(e) => setQuery(e.target.value)}
                placeholder="Search devices"
                aria-label="Search devices"
                className="h-8 w-36 text-xs"
              />
              <Select value={kind} onValueChange={setKind}>
                <SelectTrigger aria-label="Filter device type" className="h-8 max-w-32 text-xs"><SelectValue placeholder="All types" /></SelectTrigger>
                <SelectContent>{kinds.map((value) => <SelectItem key={value} value={value}>{value === "all" ? "All types" : value}</SelectItem>)}</SelectContent>
              </Select>
              <span className="text-xs text-muted-foreground">
                {active.deviceCount} device{active.deviceCount === 1 ? "" : "s"}
              </span>
              <Button type="button" variant="outline" size="sm" onClick={() => void handleExport()}>Export</Button>
              <Button type="button" variant="outline" size="sm" onClick={() => void handleImportPreview()}>Import</Button>
              <Button type="button" variant="destructive" size="sm" disabled={!selectedNode || selectedNode.isGroup || !selectedNode.isMissing} onClick={() => void handleDelete()}>Delete missing</Button>
            </>
          )}
        </div>
      </header>

      {(error || connectionError) && <div className="px-4 py-2 text-xs text-destructive">{error ?? connectionError}</div>}
      {actionMessage && <div className="border-b px-4 py-2 text-xs text-muted-foreground">{actionMessage}</div>}

      <Dialog open={Boolean(importCandidates)} onOpenChange={(open) => { if (!open) setImportCandidates(null); }}>
          <DialogContent className="flex max-h-[80vh] max-w-2xl flex-col">
            <DialogHeader>
              <DialogTitle>Import devices</DialogTitle>
              <DialogDescription>Select the devices to add. Existing devices stay protected.</DialogDescription>
            </DialogHeader>
            {importCandidates && <>
            <div className="my-3 flex gap-2 text-xs">
              <Button type="button" variant="ghost" size="sm" className="h-auto p-0 text-primary" onClick={() => setImportCandidates((current) => current?.map((candidate) => ({ ...candidate, selected: !candidate.alreadyPresent })) ?? null)}>Select available</Button>
              <Button type="button" variant="ghost" size="sm" className="h-auto p-0 text-muted-foreground" onClick={() => setImportCandidates((current) => current?.map((candidate) => ({ ...candidate, selected: false })) ?? null)}>Select none</Button>
              <span className="ml-auto text-muted-foreground">{importCandidates.filter((candidate) => candidate.selected).length} selected / {importCandidates.length}</span>
            </div>
            <div className="min-h-0 flex-1 overflow-y-auto rounded-md border">
              {importCandidates.map((candidate) => (
                <Label key={candidate.id} className="flex items-center gap-3 border-b px-3 py-2 last:border-b-0 hover:bg-muted/40">
                  <Checkbox checked={candidate.selected} disabled={candidate.alreadyPresent} onCheckedChange={() => setImportCandidates((current) => current?.map((item) => item.id === candidate.id ? { ...item, selected: !item.selected } : item) ?? null)} />
                  <span className="min-w-0 flex-1">
                    <span className="block truncate text-sm">{candidate.alias?.trim() || candidate.name?.trim() || `#${candidate.entityId}`} <span className="text-[11px] text-muted-foreground">#{candidate.entityId}</span></span>
                    <span className="block text-[11px] text-muted-foreground">{candidate.ownerName} · {candidate.kind ?? "Unknown"}</span>
                  </span>
                  {candidate.alreadyPresent && <Badge variant="secondary" className="text-[10px]">already present</Badge>}
                  {candidate.fromPreviousWipe && <Badge variant="outline" className="text-[10px] text-warning">previous wipe</Badge>}
                </Label>
              ))}
              {importCandidates.length === 0 && <p className="p-6 text-center text-sm text-muted-foreground">No device candidates found.</p>}
            </div>
            </>}
            <DialogFooter>
              <Button type="button" variant="outline" size="sm" onClick={() => setImportCandidates(null)}>Cancel</Button>
              <Button type="button" size="sm" onClick={() => void handleImportApply()} disabled={!importCandidates?.some((candidate) => candidate.selected)}>Import selected</Button>
            </DialogFooter>
          </DialogContent>
      </Dialog>

      {view === "rules" && activeKey ? (
        <div className="min-h-0 flex-1 overflow-y-auto">
          <RulesPanel matchKey={activeKey} />
        </div>
      ) : view === "timers" && activeKey ? (
        <TimersPanel matchKey={activeKey} />
      ) : view === "automation" && activeKey ? (
        <DeviceAutomationPanel matchKey={activeKey} />
      ) : (
        <div className="min-h-0 flex-1 overflow-y-auto p-2">
          {loading && profiles.length === 0 ? (
            <div className="p-4 text-sm text-muted-foreground">Loading profiles…</div>
          ) : (
            filteredDevices.map((d) => (
              <DeviceRow
                key={d.entityId}
                node={d}
                depth={0}
                liveState={deviceStates[d.entityId]}
                deviceStates={deviceStates}
                selectedId={selectedId}
                onSelect={setSelectedId}
              />
            ))
          )}
        </div>
      )}
    </div>
  );
}
