import { useEffect, useState } from "react";
import type * as React from "react";
import { cloudBootstrap, cloudLogin, cloudLogout, getWipeSettings, setWipeSettings, type CloudArchive, type CloudBootstrap, type WipeMap, type WipePlayer, type WipeSettings, type WipeStatus, wipeDeleteAllCloud, wipeDeleteCloudArchive, wipeGetCloudArchives, wipeGetMap, wipeGetPlayer, wipeGetStatus, wipeRestoreCloudArchive } from "../lib/ipc.js";
import { WipeReplayPanel } from "./WipeReplayPanel.js";
import { Badge } from "../components/ui/badge.js";
import { Button } from "../components/ui/button.js";
import { Card } from "../components/ui/card.js";
import { Checkbox } from "../components/ui/checkbox.js";
import { Input } from "../components/ui/input.js";
import { Label } from "../components/ui/label.js";

export function PlayerWipeTrackerPage(): React.JSX.Element {
  const [state, setState] = useState<CloudBootstrap | null>(null);
  const [local, setLocal] = useState<WipeStatus | null>(null);
  const [map, setMap] = useState<WipeMap | null>(null);
  const [selectedSteamId, setSelectedSteamId] = useState<string | null>(null);
  const [detail, setDetail] = useState<WipePlayer | null>(null);
  const [archives, setArchives] = useState<CloudArchive[]>([]);
  const [settings, setSettings] = useState<WipeSettings | null>(null);
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const refresh = async (): Promise<void> => {
    setError(null);
    try {
      const [cloud, tracker, wipeSettings, currentMap] = await Promise.all([cloudBootstrap(), wipeGetStatus(), getWipeSettings(), wipeGetMap()]);
      setState(cloud);
      setLocal(tracker);
      setMap(currentMap);
      setSelectedSteamId((current) => current ?? tracker.players[0]?.steamId ?? null);
      setSettings(wipeSettings);
      setArchives(cloud.signedIn && cloud.capabilities?.canUseCloudSync ? await wipeGetCloudArchives() : []);
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    }
  };

  useEffect(() => {
    void refresh();
  }, []);

  useEffect(() => {
    if (!selectedSteamId) {
      setDetail(null);
      return;
    }
    let cancelled = false;
    void wipeGetPlayer(selectedSteamId).then((player) => { if (!cancelled) setDetail(player); }).catch((err: unknown) => { if (!cancelled) setError(err instanceof Error ? err.message : String(err)); });
    return () => { cancelled = true; };
  }, [selectedSteamId, local?.sessionId, local?.players.length]);

  const signIn = async (event: React.FormEvent): Promise<void> => {
    event.preventDefault();
    setBusy(true);
    setError(null);
    try {
      await cloudLogin(email, password);
      setPassword("");
      await refresh();
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setBusy(false);
    }
  };

  const signOut = async (): Promise<void> => {
    setBusy(true);
    try {
      await cloudLogout();
      setState({ signedIn: false, user: null, capabilities: null, error: null });
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setBusy(false);
    }
  };

  const restoreArchive = async (archiveId: string): Promise<void> => {
    if (!window.confirm("Restore this cloud archive into local wipe history? Existing observations are deduplicated.")) return;
    setBusy(true);
    try {
      const result = await wipeRestoreCloudArchive(archiveId);
      await refresh();
      setError(`Restored ${result.observations} observations from ${result.days} day(s).`);
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setBusy(false);
    }
  };

  const deleteArchive = async (archiveId: string): Promise<void> => {
    if (!window.confirm("Delete this cloud archive permanently?")) return;
    setBusy(true);
    try {
      if (await wipeDeleteCloudArchive(archiveId)) setArchives((current) => current.filter((archive) => archive.id !== archiveId));
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setBusy(false);
    }
  };

  const deleteAllArchives = async (): Promise<void> => {
    if (!window.confirm("Delete every Player Wipe Tracker cloud archive permanently?")) return;
    setBusy(true);
    try {
      await wipeDeleteAllCloud();
      setArchives([]);
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setBusy(false);
    }
  };

  if (!state) return <div className="p-6 text-sm text-muted-foreground">Checking Laravel cloud account…</div>;

  return (
    <div className="h-full overflow-y-auto p-6">
      <div className="mx-auto max-w-3xl">
        <header className="mb-6 flex flex-wrap items-start justify-between gap-3">
          <div>
            <h1 className="text-lg font-semibold">Player Wipe Tracker</h1>
            <p className="mt-1 text-sm text-muted-foreground">Cloud entitlements and tracker workspace</p>
          </div>
          {state.signedIn && <Button type="button" variant="outline" size="sm" onClick={() => void signOut()} disabled={busy}>Sign out</Button>}
        </header>

        {error && <p role="alert" className="mb-4 rounded-md border border-destructive/40 bg-destructive/10 px-3 py-2 text-sm text-destructive">{error}</p>}
        {state.error && <p role="status" className="mb-4 rounded-md border border-warning/40 bg-warning/10 px-3 py-2 text-sm text-warning">Cloud refresh failed: {state.error}</p>}
        {settings && <Card className="mb-4 p-4"><div className="flex flex-wrap items-center justify-between gap-3"><div><h2 className="text-sm font-medium">Local tracking</h2><p className="mt-1 text-xs text-muted-foreground">Record team observations into the local wipe history.</p></div><Label className="flex items-center gap-2 text-sm"><Checkbox checked={settings.enabled} onCheckedChange={(checked) => void setWipeSettings({ enabled: checked === true }).then(setSettings).catch((err: unknown) => setError(err instanceof Error ? err.message : String(err)))} /> Enabled</Label></div><Label className="mt-3 flex items-center gap-2 text-xs text-muted-foreground"><Checkbox checked={settings.cloudBackupEnabled} disabled={!settings.enabled} onCheckedChange={(checked) => void setWipeSettings({ cloudBackupEnabled: checked === true }).then(setSettings).catch((err: unknown) => setError(err instanceof Error ? err.message : String(err)))} /> Back up daily payloads to Laravel cloud</Label></Card>}

        {!state.signedIn ? (
          <Card className="max-w-sm p-5"><form onSubmit={(event) => void signIn(event)}>
            <h2 className="font-medium">Sign in to unlock cloud features</h2>
            <p className="mt-1 text-xs text-muted-foreground">Your Laravel session is encrypted in the Electron main process.</p>
            <div className="mt-4 space-y-2"><Label htmlFor="cloud-email">Email</Label><Input id="cloud-email" required type="email" value={email} onChange={(event) => setEmail(event.target.value)} /></div>
            <div className="mt-3 space-y-2"><Label htmlFor="cloud-password">Password</Label><Input id="cloud-password" required type="password" value={password} onChange={(event) => setPassword(event.target.value)} /></div>
            <Button type="submit" disabled={busy} className="mt-4 w-full">{busy ? "Signing in…" : "Sign in"}</Button>
          </form></Card>
        ) : state.capabilities ? (
          <div className="space-y-4">
            <Card className="p-5">
              <p className="text-sm text-muted-foreground">Signed in as <span className="font-medium text-foreground">{state.user?.displayName ?? state.user?.email ?? "cloud user"}</span></p>
              <div className="mt-4 grid gap-3 sm:grid-cols-3">
                <Metric label="Plan" value={state.capabilities.planCode} />
                <Metric label="Tracked players" value={String(state.capabilities.maxTrackedPlayers)} />
                <Metric label="Retained wipes" value={String(state.capabilities.retainedWipes)} />
              </div>
            </Card>
            <div className="grid gap-2 sm:grid-cols-2">
              <Capability label="Tracker access" enabled={state.capabilities.isTrackerAvailable} />
              <Capability label="Team tracking" enabled={state.capabilities.canTrackTeam} />
              <Capability label="Cloud sync" enabled={state.capabilities.canUseCloudSync} />
              <Capability label="Advanced views" enabled={state.capabilities.canUseAdvancedViews} />
              <Capability label="Route replay" enabled={state.capabilities.canUseRouteReplay} />
              <Capability label="Export" enabled={state.capabilities.canExport} />
            </div>
            <Card className="p-5">
              <div className="flex items-center justify-between gap-3">
                <div><h2 className="font-medium">Current wipe</h2><p className="mt-1 text-xs text-muted-foreground">{local?.serverKey ?? "No active Rust+ session"}</p></div>
                <span className="text-xs text-muted-foreground">{local?.players.length ?? 0} tracked</span>
              </div>
              {local?.players.length ? <div className="mt-4 grid gap-2 md:grid-cols-2">{local.players.map((player) => <Button type="button" variant="ghost" key={player.steamId} onClick={() => setSelectedSteamId(player.steamId)} className={`h-auto justify-start rounded-md border px-3 py-2 text-left ${selectedSteamId === player.steamId ? "border-primary/60 bg-primary/10" : ""}`}><div className="w-full"><div className="flex items-center justify-between gap-2"><span className="truncate text-sm font-medium">{player.name}</span><Badge variant="outline" className={player.insights.isLikelyOnline ? "text-xs text-emerald-500" : "text-xs text-muted-foreground"}>{player.insights.currentState}{player.insights.isLikelyOnline ? " · online" : ""}</Badge></div><p className="mt-1 text-[11px] text-muted-foreground">{player.summary.deaths} deaths · {Math.round(player.summary.estimatedDistance).toLocaleString()}m · {player.observationCount} observations</p></div></Button>)}</div> : <p className="mt-4 text-sm text-muted-foreground">No local observations yet. Start a Rust+ team session with the tracker enabled to begin recording.</p>}
            </Card>
            {detail && <WipeReplayPanel player={detail} map={map} />}
            <Card className="p-5">
              <div className="flex flex-wrap items-center justify-between gap-3"><div><h2 className="font-medium">Cloud archives</h2><p className="mt-1 text-xs text-muted-foreground">Laravel backups available for restore.</p></div>{state.capabilities.canUseCloudSync && archives.length > 0 && <Button type="button" variant="ghost" size="sm" disabled={busy} onClick={() => void deleteAllArchives()} className="h-auto p-0 text-xs text-destructive">Delete all</Button>}</div>
              {!state.capabilities.canUseCloudSync ? <p className="mt-4 text-sm text-muted-foreground">Cloud backup is locked for this plan.</p> : archives.length === 0 ? <p className="mt-4 text-sm text-muted-foreground">No cloud archives found.</p> : <div className="mt-4 space-y-2">{archives.map((archive) => <div key={archive.id} className="flex flex-wrap items-center gap-3 rounded-md border px-3 py-2"><div className="min-w-0 flex-1"><p className="truncate text-sm font-medium">{archive.serverName}</p><p className="text-[11px] text-muted-foreground">{archive.wipeStartedAtUtc ? formatDate(archive.wipeStartedAtUtc) : "Unknown wipe"} · {archive.players.length} player(s) · {archive.storedBytes == null ? "size unknown" : formatBytes(archive.storedBytes)}</p></div><Button type="button" variant="outline" size="sm" disabled={busy} onClick={() => void restoreArchive(archive.id)}>Restore</Button><Button type="button" variant="destructive" size="sm" disabled={busy} onClick={() => void deleteArchive(archive.id)}>Delete</Button></div>)}</div>}
            </Card>
          </div>
        ) : (
          <p className="rounded-lg border border-dashed p-5 text-sm text-muted-foreground">Sign in to load Player Wipe Tracker entitlements.</p>
        )}
      </div>
    </div>
  );
}

function Metric({ label, value }: { label: string; value: string }): React.JSX.Element {
  return <Card className="p-3"><p className="text-[11px] text-muted-foreground">{label}</p><p className="mt-1 text-sm font-semibold">{value}</p></Card>;
}

function Capability({ label, enabled }: { label: string; enabled: boolean }): React.JSX.Element {
  return <Card className="flex items-center justify-between rounded-md px-3 py-2 text-sm"><span>{label}</span><Badge variant="outline" className={enabled ? "text-emerald-500" : "text-muted-foreground"}>{enabled ? "Enabled" : "Locked"}</Badge></Card>;
}

function formatDate(value: string): string { const date = new Date(value); return Number.isNaN(date.getTime()) ? value : date.toLocaleString(); }
function formatBytes(value: number): string { return value < 1024 ? `${value} B` : `${(value / 1024).toFixed(1)} KB`; }
