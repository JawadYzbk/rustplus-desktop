import { useMemo } from "react";
import type * as React from "react";
import type { WipeMap, WipePlayer } from "../lib/ipc.js";
import { Card } from "../components/ui/card.js";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "../components/ui/table.js";

const VIEW_WIDTH = 1000;
const VIEW_HEIGHT = 560;
const STATE_COLORS: Record<WipePlayer["observations"][number]["state"], string> = {
  moving: "#38bdf8",
  stationary: "#60a5fa",
  afk: "#a78bfa",
  dead: "#f87171",
  offline: "#94a3b8",
  unknown: "#64748b",
};

export function WipeReplayPanel({ player, map }: { player: WipePlayer; map: WipeMap | null }): React.JSX.Element {
  const points = useMemo(() => player.observations.filter((point) => point.x !== null && point.y !== null), [player.observations]);
  const project = useMemo(() => makeProjector(points, map), [map, points]);
  const polyline = points.map((point) => {
    const mapped = project(point.x!, point.y!);
    return `${mapped.x},${mapped.y}`;
  }).join(" ");
  const totalSeconds = Math.max(1, player.segments.reduce((total, segment) => total + Math.max(0, (Date.parse(segment.endUtc) - Date.parse(segment.startUtc)) / 1000), 0));

  return (
    <Card className="p-5">
      <div className="flex flex-wrap items-start justify-between gap-3">
        <div><h2 className="font-medium">Replay · {player.name}</h2><p className="mt-1 text-xs text-muted-foreground">{player.observationCount} observations · {player.insights.sessionCount} session(s) · {player.insights.currentLocationName ?? player.insights.currentGrid ?? "location unknown"}</p></div>
        <div className="text-right text-xs text-muted-foreground"><p>{formatDate(player.insights.firstSeenUtc)} → {formatDate(player.insights.lastSeenUtc)}</p><p className="mt-1">{formatDuration(totalSeconds)} tracked timeline</p></div>
      </div>

      <div className="mt-4 overflow-hidden rounded-md border bg-slate-950/70">
        <svg viewBox={`0 0 ${VIEW_WIDTH} ${VIEW_HEIGHT}`} className="block h-auto min-h-64 w-full" role="img" aria-label={`Player route replay for ${player.name}`}>
          {map ? <image href={`data:image/png;base64,${map.pngBase64}`} x={0} y={0} width={VIEW_WIDTH} height={VIEW_HEIGHT} preserveAspectRatio="xMidYMid meet" opacity={0.62} /> : <GridBackground />}
          {polyline && <polyline points={polyline} fill="none" stroke="#e2e8f0" strokeOpacity={0.55} strokeWidth={2} />}
          {points.map((point, index) => {
            const mapped = project(point.x!, point.y!);
            return <circle key={`${point.timestampUtc}-${index}`} cx={mapped.x} cy={mapped.y} r={point.event ? 6 : 3.5} fill={STATE_COLORS[point.state]} stroke={point.event ? "#fff" : "none"} strokeWidth={point.event ? 2 : 0} />;
          })}
          {!points.length && <text x={VIEW_WIDTH / 2} y={VIEW_HEIGHT / 2} textAnchor="middle" fill="#94a3b8" fontSize="18">No coordinate points recorded</text>}
        </svg>
      </div>

      <div className="mt-3 flex flex-wrap gap-x-3 gap-y-1 text-[11px] text-muted-foreground">{Object.entries(STATE_COLORS).map(([state, color]) => <span key={state} className="inline-flex items-center gap-1"><span className="h-2 w-2 rounded-full" style={{ backgroundColor: color }} />{state}</span>)}</div>

      <div className="mt-4"><p className="mb-2 text-xs font-medium">Activity timeline</p><div className="flex h-7 overflow-hidden rounded-md bg-muted">{player.segments.map((segment, index) => { const width = Math.max(0.5, (Math.max(0, Date.parse(segment.endUtc) - Date.parse(segment.startUtc)) / 1000) / totalSeconds * 100); return <div key={`${segment.startUtc}-${index}`} title={`${segment.state} · ${formatDuration((Date.parse(segment.endUtc) - Date.parse(segment.startUtc)) / 1000)}`} style={{ width: `${width}%`, backgroundColor: STATE_COLORS[segment.state] }} />; })}</div><div className="mt-2 flex justify-between text-[11px] text-muted-foreground"><span>{formatDate(player.insights.firstSeenUtc)}</span><span>{formatDate(player.insights.lastSeenUtc)}</span></div></div>

      <Table className="min-w-[560px] text-left text-xs"><TableHeader><TableRow><TableHead>Time</TableHead><TableHead>State</TableHead><TableHead>Location</TableHead><TableHead>Event</TableHead></TableRow></TableHeader><TableBody>{player.observations.slice(-12).reverse().map((point, index) => <TableRow key={`${point.timestampUtc}-${index}`}><TableCell className="whitespace-nowrap">{formatDate(point.timestampUtc)}</TableCell><TableCell><span className="inline-flex items-center gap-1"><span className="h-2 w-2 rounded-full" style={{ backgroundColor: STATE_COLORS[point.state] }} />{point.state}</span></TableCell><TableCell>{point.locationName ?? point.grid ?? "—"}</TableCell><TableCell>{point.event ?? "—"}</TableCell></TableRow>)}</TableBody></Table>
    </Card>
  );
}

function GridBackground(): React.JSX.Element {
  return <><rect width={VIEW_WIDTH} height={VIEW_HEIGHT} fill="#0f172a" />{Array.from({ length: 11 }, (_, index) => <line key={`v-${index}`} x1={index * 100} x2={index * 100} y1={0} y2={VIEW_HEIGHT} stroke="#334155" strokeOpacity={0.4} />)}{Array.from({ length: 6 }, (_, index) => <line key={`h-${index}`} x1={0} x2={VIEW_WIDTH} y1={index * 100} y2={index * 100} stroke="#334155" strokeOpacity={0.4} />)}</>;
}

function makeProjector(points: Array<{ x: number | null; y: number | null }>, map: WipeMap | null): (x: number, y: number) => { x: number; y: number } {
  if (map) {
    const scale = Math.min(VIEW_WIDTH / map.imageWidth, VIEW_HEIGHT / map.imageHeight);
    const imageLeft = (VIEW_WIDTH - map.imageWidth * scale) / 2;
    const imageTop = (VIEW_HEIGHT - map.imageHeight * scale) / 2;
    return (worldX, worldY) => ({ x: imageLeft + (map.worldRectX + worldX / map.worldSize * map.worldRectWidth) * scale, y: imageTop + (map.worldRectY + (1 - worldY / map.worldSize) * map.worldRectHeight) * scale });
  }
  const usable = points.filter((point): point is { x: number; y: number } => point.x !== null && point.y !== null);
  const minX = Math.min(...usable.map((point) => point.x), 0);
  const maxX = Math.max(...usable.map((point) => point.x), 1);
  const minY = Math.min(...usable.map((point) => point.y), 0);
  const maxY = Math.max(...usable.map((point) => point.y), 1);
  const spanX = Math.max(1, maxX - minX);
  const spanY = Math.max(1, maxY - minY);
  return (x, y) => ({ x: 36 + (x - minX) / spanX * (VIEW_WIDTH - 72), y: VIEW_HEIGHT - 36 - (y - minY) / spanY * (VIEW_HEIGHT - 72) });
}

function formatDate(value: string | null): string { if (!value) return "—"; const date = new Date(value); return Number.isNaN(date.getTime()) ? value : date.toLocaleString(); }
function formatDuration(seconds: number): string { const safe = Math.max(0, Math.round(seconds)); const hours = Math.floor(safe / 3600); const minutes = Math.floor((safe % 3600) / 60); const remainder = safe % 60; return hours > 0 ? `${hours}h ${minutes}m` : minutes > 0 ? `${minutes}m ${remainder}s` : `${remainder}s`; }
