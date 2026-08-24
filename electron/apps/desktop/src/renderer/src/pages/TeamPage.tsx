import { useEffect, useMemo, useState } from "react";
import type * as React from "react";
import { useNavigate } from "react-router-dom";
import { Badge } from "../components/ui/badge.js";
import { Button } from "../components/ui/button.js";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "../components/ui/card.js";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "../components/ui/select.js";
import { activateProfile } from "../lib/ipc.js";
import { useConnectionStore, type TeamMember } from "../stores/connection.js";
import { useProfilesStore } from "../stores/profiles.js";

function memberState(member: TeamMember): { label: string; className: string } {
  if (member.dead) return { label: "Downed", className: "text-destructive" };
  if (member.online) return { label: "Online", className: "text-emerald-400" };
  return { label: "Offline", className: "text-muted-foreground" };
}

function position(member: TeamMember): string {
  return member.x === null || member.y === null ? "Position unavailable" : `${member.x.toFixed(0)} / ${member.y.toFixed(0)}`;
}

export function TeamPage(): React.JSX.Element {
  const navigate = useNavigate();
  const profiles = useProfilesStore((state) => state.profiles);
  const activeKey = useProfilesStore((state) => state.activeKey);
  const loadingProfiles = useProfilesStore((state) => state.loading);
  const loadProfiles = useProfilesStore((state) => state.loadProfiles);
  const selectProfile = useProfilesStore((state) => state.selectProfile);
  const snapshot = useConnectionStore((state) => state.snapshot);
  const phase = useConnectionStore((state) => state.phase);
  const team = useConnectionStore((state) => state.team);
  const serverStatus = useConnectionStore((state) => state.status);
  const connectionError = useConnectionStore((state) => state.error);
  const hydrateConnection = useConnectionStore((state) => state.hydrate);
  const connect = useConnectionStore((state) => state.connectProfile);
  const disconnect = useConnectionStore((state) => state.disconnect);
  const [busy, setBusy] = useState(false);
  const [pageError, setPageError] = useState<string | null>(null);

  useEffect(() => {
    void loadProfiles();
    void hydrateConnection();
  }, [loadProfiles, hydrateConnection]);

  const active = useMemo(() => profiles.find((profile) => profile.matchKey === activeKey) ?? null, [profiles, activeKey]);
  const onlineCount = team?.members.filter((member) => member.online && !member.dead).length ?? 0;
  const deadCount = team?.members.filter((member) => member.dead).length ?? 0;

  const connectActive = async (matchKey: string): Promise<void> => {
    setBusy(true);
    setPageError(null);
    try {
      if (!await activateProfile(matchKey)) throw new Error("Unable to activate server profile");
      selectProfile(matchKey);
      await connect(matchKey);
    } catch (reason: unknown) {
      setPageError(reason instanceof Error ? reason.message : String(reason));
    } finally {
      setBusy(false);
    }
  };

  if (profiles.length === 0 && !loadingProfiles) {
    return (
      <div className="flex h-full items-center justify-center p-6">
        <Card className="max-w-sm p-6 text-center">
          <CardTitle>Team is waiting for a server</CardTitle>
          <CardDescription className="mt-2">Pair Rust+ or import your previous RustPlusDesk data before loading team members.</CardDescription>
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
      <header className="flex flex-wrap items-center gap-3 border-b px-4 py-2.5">
        <div>
          <h1 className="text-sm font-semibold">Team</h1>
          <p className="text-[11px] text-muted-foreground">Live Rust+ team roster</p>
        </div>
        <Select value={activeKey ?? ""} onValueChange={(value) => void connectActive(value)}>
          <SelectTrigger aria-label="Active profile" className="w-56 text-xs"><SelectValue placeholder="Select server" /></SelectTrigger>
          <SelectContent>{profiles.map((profile) => <SelectItem key={profile.matchKey} value={profile.matchKey}>{profile.name} ({profile.host}:{profile.port})</SelectItem>)}</SelectContent>
        </Select>
        <div className="ml-auto flex items-center gap-2">
          <Badge variant={snapshot.connected ? "default" : "secondary"} className="gap-1.5">
            <span className={`h-1.5 w-1.5 rounded-full ${snapshot.connected ? "bg-emerald-300" : "bg-muted-foreground"}`} />
            {phase === "reconnecting" ? "Reconnecting" : snapshot.connected ? "Connected" : "Disconnected"}
          </Badge>
          {snapshot.connected ? (
            <Button type="button" variant="outline" size="sm" disabled={busy} onClick={() => void disconnect()}>Disconnect</Button>
          ) : (
            <Button type="button" size="sm" disabled={!activeKey || busy} onClick={() => activeKey && void connectActive(activeKey)}>{busy ? "Connecting…" : "Connect"}</Button>
          )}
        </div>
      </header>

      {(pageError || connectionError) && <div className="border-b px-4 py-2 text-xs text-destructive">{pageError ?? connectionError}</div>}

      <div className="min-h-0 flex-1 overflow-y-auto p-4">
        <div className="mb-4 grid gap-3 sm:grid-cols-4">
          <Card><CardHeader className="pb-2"><CardDescription>Roster</CardDescription><CardTitle className="text-2xl">{team?.members.length ?? "—"}</CardTitle></CardHeader></Card>
          <Card><CardHeader className="pb-2"><CardDescription>Online and alive</CardDescription><CardTitle className="text-2xl text-emerald-400">{team ? onlineCount : "—"}</CardTitle></CardHeader></Card>
          <Card><CardHeader className="pb-2"><CardDescription>Downed</CardDescription><CardTitle className="text-2xl text-destructive">{team ? deadCount : "—"}</CardTitle></CardHeader></Card>
          <Card><CardHeader className="pb-2"><CardDescription>Queued players</CardDescription><CardTitle className="text-2xl text-primary">{serverStatus?.queuedPlayers ?? "—"}</CardTitle><CardDescription>{serverStatus ? `${serverStatus.players}/${serverStatus.maxPlayers} online` : "Waiting for server status"}</CardDescription></CardHeader></Card>
        </div>

        {!snapshot.connected ? (
          <Card className="border-dashed">
            <CardHeader><CardTitle>Connect to load the team</CardTitle><CardDescription>Select a stored profile and connect. Team data is read from Rust+ and never fabricated locally.</CardDescription></CardHeader>
          </Card>
        ) : !team ? (
          <Card className="border-dashed">
            <CardHeader><CardTitle>Waiting for team data</CardTitle><CardDescription>The connection is live; the first Rust+ team snapshot will appear here.</CardDescription></CardHeader>
          </Card>
        ) : team.members.length === 0 ? (
          <Card className="border-dashed">
            <CardHeader><CardTitle>No team members returned</CardTitle><CardDescription>Rust+ returned an empty team roster for this server.</CardDescription></CardHeader>
          </Card>
        ) : (
          <div className="grid gap-3 sm:grid-cols-2 xl:grid-cols-3">
            {team.members.map((member) => {
              const state = memberState(member);
              const isLeader = member.steamId === team.leaderSteamId;
              return (
                <Card key={member.steamId} className="border-border/80 bg-card/70">
                  <CardContent className="flex items-start gap-3 p-4">
                    <div className={`mt-1 h-2.5 w-2.5 shrink-0 rounded-full ${member.dead ? "bg-destructive" : member.online ? "bg-emerald-400" : "bg-muted-foreground/50"}`} aria-hidden />
                    <div className="min-w-0 flex-1">
                      <div className="flex items-center gap-2">
                        <p className="truncate text-sm font-medium">{member.name}</p>
                        {isLeader && <Badge variant="outline" className="text-[10px]">Leader</Badge>}
                      </div>
                      <p className={`mt-1 text-xs ${state.className}`}>{state.label} · {position(member)}</p>
                      <p className="mt-1 truncate text-[11px] text-muted-foreground">{member.steamId}</p>
                    </div>
                  </CardContent>
                </Card>
              );
            })}
          </div>
        )}
      </div>
    </div>
  );
}
