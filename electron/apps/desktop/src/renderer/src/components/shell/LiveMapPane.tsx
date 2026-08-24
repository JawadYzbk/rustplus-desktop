import { Crosshair, Radio } from "lucide-react";
import type * as React from "react";
import { Badge } from "../ui/badge.js";
import { cn } from "../../lib/cn.js";
import { projectWorldPoint, type MapPoint } from "../../lib/map-projection.js";
import { useConnectionStore, type LiveMapMarker, type MapSnapshot, type TeamMember } from "../../stores/connection.js";

const markerTone: Record<string, string> = {
  Player: "bg-primary ring-primary/40",
  VendingMachine: "bg-amber-300 ring-amber-300/30",
  CargoShip: "bg-orange-300 ring-orange-300/30",
  CH47: "bg-orange-300 ring-orange-300/30",
  PatrolHelicopter: "bg-red-400 ring-red-400/30",
  Crate: "bg-yellow-200 ring-yellow-200/30",
  Explosion: "bg-red-300 ring-red-300/30",
  GenericRadius: "bg-violet-300 ring-violet-300/30",
};


function pointForMember(map: MapSnapshot, member: TeamMember): MapPoint | null {
  return member.x === null || member.y === null ? null : projectWorldPoint(map, member.x, member.y);
}

function pointForMarker(map: MapSnapshot, marker: LiveMapMarker): MapPoint | null {
  return projectWorldPoint(map, marker.x, marker.y);
}

function MarkerDot({ point, label, tone, className }: { point: MapPoint; label: string; tone: string; className?: string }): React.JSX.Element {
  return (
    <span
      className={cn("absolute z-10 h-2.5 w-2.5 -translate-x-1/2 -translate-y-1/2 rounded-full ring-4", tone, className)}
      style={{ left: `${point.left}%`, top: `${point.top}%` }}
      title={label}
      aria-label={label}
    />
  );
}

function mapImageSource(map: MapSnapshot): string | null {
  if (!map.imageBase64) return null;
  return map.imageBase64.startsWith("data:") ? map.imageBase64 : `data:image/jpeg;base64,${map.imageBase64}`;
}

export function LiveMapPane(): React.JSX.Element {
  const connected = useConnectionStore((state) => state.snapshot.connected);
  const phase = useConnectionStore((state) => state.phase);
  const map = useConnectionStore((state) => state.map);
  const markers = useConnectionStore((state) => state.markers);
  const team = useConnectionStore((state) => state.team);
  const imageSource = map ? mapImageSource(map) : null;

  const teamPoints = map && team ? team.members.flatMap((member) => {
    const point = pointForMember(map, member);
    return point ? [{ member, point }] : [];
  }) : [];
  const markerPoints = map ? markers.flatMap((marker) => {
    const point = pointForMarker(map, marker);
    return point ? [{ marker, point }] : [];
  }) : [];
  const monumentPoints = map ? map.monuments.flatMap((monument) => {
    const point = projectWorldPoint(map, monument.x, monument.y);
    return point ? [{ monument, point }] : [];
  }) : [];

  return (
    <aside aria-label="Live Rust+ map" className="flex min-h-0 w-[min(38vw,520px)] min-w-[360px] flex-col border-l bg-card/35">
      <header className="flex items-center gap-3 border-b px-4 py-3">
        <div className="flex h-8 w-8 items-center justify-center rounded-md border border-primary/30 bg-primary/10 text-primary">
          <Crosshair className="h-4 w-4" aria-hidden />
        </div>
        <div className="min-w-0">
          <h2 className="text-sm font-semibold tracking-wide">Live map</h2>
          <p className="text-[10px] uppercase tracking-[0.18em] text-muted-foreground">Rust+ field telemetry</p>
        </div>
        <Badge variant={connected ? "default" : "secondary"} className="ml-auto gap-1.5 text-[10px]">
          <Radio className="h-3 w-3" aria-hidden />
          {phase === "reconnecting" ? "Reconnecting" : connected ? "Live" : "Offline"}
        </Badge>
      </header>

      <div className="min-h-0 flex-1 p-3">
        <div className="relative h-full min-h-[260px] overflow-hidden rounded-lg border border-primary/20 bg-background shadow-[inset_0_0_40px_hsl(var(--primary)/0.05)]">
          {!connected ? (
            <div className="flex h-full items-center justify-center p-6 text-center">
              <div>
                <Crosshair className="mx-auto mb-3 h-8 w-8 text-muted-foreground/40" aria-hidden />
                <p className="text-sm font-medium">Map offline</p>
                <p className="mt-1 text-xs text-muted-foreground">Connect to a Rust+ server to load its map and live positions.</p>
              </div>
            </div>
          ) : !map ? (
            <div className="flex h-full items-center justify-center p-6 text-center">
              <div>
                <Radio className="mx-auto mb-3 h-8 w-8 animate-pulse text-primary/70" aria-hidden />
                <p className="text-sm font-medium">Requesting map snapshot</p>
                <p className="mt-1 text-xs text-muted-foreground">The connection is live; waiting for Rust+ map data.</p>
              </div>
            </div>
          ) : !imageSource ? (
            <div className="flex h-full items-center justify-center p-6 text-center">
              <div>
                <p className="text-sm font-medium">Map payload received</p>
                <p className="mt-1 text-xs text-muted-foreground">Rust+ returned coordinates but no map image.</p>
              </div>
            </div>
          ) : (
            <figure className="relative h-full w-full overflow-hidden">
              <img src={imageSource} alt="Rust+ server map" className="absolute inset-0 h-full w-full object-fill" />
              <div className="map-grid absolute inset-0" aria-hidden />

              {monumentPoints.map(({ monument, point }, index) => (
                <span
                  key={`${monument.token ?? "monument"}-${index}`}
                  className="absolute z-[1] h-1 w-1 -translate-x-1/2 -translate-y-1/2 rounded-full bg-foreground/50"
                  style={{ left: `${point.left}%`, top: `${point.top}%` }}
                  title={monument.token ?? "Monument"}
                  aria-label={monument.token ?? "Monument"}
                />
              ))}

              {markerPoints.map(({ marker, point }) => (
                <MarkerDot
                  key={`marker-${marker.id}`}
                  point={point}
                  tone={markerTone[marker.type] ?? "bg-foreground ring-foreground/30"}
                  label={marker.name ?? marker.type}
                  className={marker.type === "Player" ? "h-3 w-3" : undefined}
                />
              ))}

              {teamPoints.map(({ member, point }) => (
                <MarkerDot
                  key={`team-${member.steamId}`}
                  point={point}
                  tone={member.dead ? "bg-destructive ring-destructive/35" : member.online ? "bg-emerald-300 ring-emerald-300/35" : "bg-muted-foreground ring-muted-foreground/30"}
                  label={`${member.name} · ${member.dead ? "Downed" : member.online ? "Online" : "Offline"}`}
                  className="z-20 h-3.5 w-3.5 border-2 border-background"
                />
              ))}

              <div className="pointer-events-none absolute left-3 top-3 rounded-md border border-primary/30 bg-background/80 px-2.5 py-2 font-mono text-[10px] shadow-lg backdrop-blur-sm">
                <div className="flex items-center gap-2 text-primary"><Crosshair className="h-3 w-3" aria-hidden />LIVE GRID</div>
                <div className="mt-1 text-muted-foreground">{teamPoints.length} team · {markerPoints.length} signals</div>
              </div>
              <div className="pointer-events-none absolute bottom-3 right-3 rounded border border-border/70 bg-background/75 px-2 py-1 font-mono text-[10px] text-muted-foreground backdrop-blur-sm">
                {map.worldSize > 0 ? `SCALE ${map.worldSize}u` : "SCALE CALIBRATING"}
              </div>
            </figure>
          )}
        </div>
      </div>

      <footer className="flex items-center justify-between border-t px-4 py-2 text-[10px] text-muted-foreground">
        <span>Base map: connection snapshot</span>
        <span>Markers: 2s poll</span>
      </footer>
    </aside>
  );
}
