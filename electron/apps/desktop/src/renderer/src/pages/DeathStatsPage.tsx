import { useEffect, useState } from "react";
import type * as React from "react";
import { clearDeathLog, getDeathStats, type DeathStats } from "../lib/ipc.js";
import { Button } from "../components/ui/button.js";
import { Card } from "../components/ui/card.js";
import { Input } from "../components/ui/input.js";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "../components/ui/select.js";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "../components/ui/table.js";

const empty: DeathStats = {
  total: 0, victims: 0, avgSurvival: "—", longestSurvival: "—", peakHour: "—", deadliestPlace: "—", deadliestGrid: "—",
  byArea: [], byVictim: [], byLocation: [], recent: [], deathsPerDay: [],
};

export function DeathStatsPage(): React.JSX.Element {
  const [summary, setSummary] = useState<DeathStats>(empty);
  const [players, setPlayers] = useState<string[]>([]);
  const [search, setSearch] = useState("");
  const [player, setPlayer] = useState("all");
  const [type, setType] = useState<"all" | "monument" | "base" | "open">("all");
  const [range, setRange] = useState<"all" | "24h" | "7d">("all");
  const [error, setError] = useState<string | null>(null);

  const refresh = (): void => {
    void getDeathStats({ search, player, type, range })
      .then((result) => { setSummary(result.summary); setPlayers(result.players); setError(null); })
      .catch((reason: unknown) => setError(reason instanceof Error ? reason.message : String(reason)));
  };

  useEffect(() => {
    const timer = window.setTimeout(refresh, 120);
    return () => window.clearTimeout(timer);
  }, [search, player, type, range]);

  const clear = (): void => {
    if (!window.confirm("Clear the local death log for this server? This cannot be undone.")) return;
    void clearDeathLog().then(refresh).catch((reason: unknown) => setError(reason instanceof Error ? reason.message : String(reason)));
  };

  const maxDay = Math.max(1, ...summary.deathsPerDay.map((day) => day.count));
  return (
    <div className="flex h-full min-h-0 flex-col">
      <header className="flex flex-wrap items-center gap-3 border-b px-4 py-2.5">
        <div>
          <h1 className="text-sm font-semibold">Death Stats</h1>
          <p className="text-[11px] text-muted-foreground">Where and when your team dies, from the local death log.</p>
        </div>
        <div className="ml-auto flex gap-2 text-xs">
          <Button type="button" variant="outline" size="sm" onClick={refresh}>Refresh</Button>
          <Button type="button" variant="destructive" size="sm" onClick={clear}>Clear log</Button>
        </div>
      </header>
      {error && <div className="border-b px-4 py-2 text-xs text-destructive">{error}</div>}
      <main className="min-h-0 flex-1 overflow-y-auto p-3">
        <Card className="mb-3 grid gap-2 p-3 md:grid-cols-[minmax(0,1fr)_160px_130px_120px]">
          <Input value={search} onChange={(event) => setSearch(event.target.value)} placeholder="Search victim, location, or grid" aria-label="Search deaths" className="text-xs" />
          <Select value={player} onValueChange={setPlayer}>
            <SelectTrigger aria-label="Death player filter" className="text-xs"><SelectValue placeholder="All players" /></SelectTrigger>
            <SelectContent><SelectItem value="all">All players</SelectItem>{players.map((value) => <SelectItem key={value} value={value}>{value}</SelectItem>)}</SelectContent>
          </Select>
          <Select value={type} onValueChange={(value) => setType(value as typeof type)}>
            <SelectTrigger aria-label="Death location filter" className="text-xs"><SelectValue placeholder="All places" /></SelectTrigger>
            <SelectContent><SelectItem value="all">All places</SelectItem><SelectItem value="monument">Monument</SelectItem><SelectItem value="base">Base</SelectItem><SelectItem value="open">Open</SelectItem></SelectContent>
          </Select>
          <Select value={range} onValueChange={(value) => setRange(value as typeof range)}>
            <SelectTrigger aria-label="Death time filter" className="text-xs"><SelectValue placeholder="All time" /></SelectTrigger>
            <SelectContent><SelectItem value="all">All time</SelectItem><SelectItem value="24h">Last 24h</SelectItem><SelectItem value="7d">Last 7d</SelectItem></SelectContent>
          </Select>
        </Card>

        <p className="mb-3 text-sm font-semibold">{summary.total ? `${summary.total} death(s) across ${summary.victims} player(s).` : "No deaths match the current filters."}</p>
        <div className="mb-3 grid grid-cols-2 gap-2 md:grid-cols-3 xl:grid-cols-6">
          <Metric title="Total deaths" value={String(summary.total)} />
          <Metric title="Avg survival" value={summary.avgSurvival} />
          <Metric title="Longest survival" value={summary.longestSurvival} />
          <Metric title="Deadliest place" value={summary.deadliestPlace} />
          <Metric title="Deadliest grid" value={summary.deadliestGrid} />
          <Metric title="Peak hour" value={summary.peakHour} />
        </div>

        <div className="grid gap-3 xl:grid-cols-2">
          <Card className="p-3">
            <h2 className="mb-2 text-xs font-medium">Where you die</h2>
            <div className="space-y-1.5">{summary.byArea.map((area) => <div key={area.type} className="flex items-center gap-2 text-xs"><span className="w-20">{area.name}</span><div className="h-2 flex-1 rounded bg-muted"><div className="h-2 rounded bg-primary" style={{ width: `${Math.max(area.percent, area.deaths ? 3 : 0)}%` }} /></div><span className="w-20 text-right text-muted-foreground">{area.deaths} · {area.percent}%</span></div>)}</div>
          </Card>
          <Card className="p-3">
            <h2 className="mb-2 text-xs font-medium">Deaths per day (last 14 days)</h2>
            <div className="flex h-14 items-end gap-1">{summary.deathsPerDay.map((day) => <div key={day.day} title={`${day.day}: ${day.count}`} className="flex-1 rounded-sm bg-primary/80" style={{ height: `${day.count ? Math.max(8, day.count / maxDay * 100) : 4}%`, opacity: day.count ? 1 : 0.3 }} />)}</div>
          </Card>
          <DataTable title="By location" headers={["Location", "Deaths"]} rows={summary.byLocation.map((item) => [item.location, String(item.deaths)])} />
          <DataTable title="By player" headers={["Victim", "Deaths", "Avg survival"]} rows={summary.byVictim.map((item) => [item.victim, String(item.deaths), item.avgSurvival])} />
          <Card className="p-3 xl:col-span-2">
            <h2 className="mb-2 text-xs font-medium">Recent</h2>
            <Table className="text-left text-xs">
              <TableHeader><TableRow><TableHead>Victim</TableHead><TableHead>Type</TableHead><TableHead>Location</TableHead><TableHead>Grid</TableHead><TableHead>Died</TableHead></TableRow></TableHeader>
              <TableBody>{summary.recent.map((item, index) => <TableRow key={`${item.died}-${index}`}><TableCell>{item.victim}</TableCell><TableCell>{item.type}</TableCell><TableCell>{item.location}</TableCell><TableCell>{item.grid}</TableCell><TableCell>{item.died}</TableCell></TableRow>)}</TableBody>
            </Table>
            {!summary.recent.length && <p className="py-5 text-center text-xs text-muted-foreground">No deaths recorded.</p>}
          </Card>
        </div>
      </main>
    </div>
  );
}

function Metric({ title, value }: { title: string; value: string }): React.JSX.Element {
  return <Card className="px-2 py-2"><p className="text-[10px] font-semibold uppercase text-muted-foreground">{title}</p><p className="truncate text-sm font-semibold" title={value}>{value}</p></Card>;
}

function DataTable({ title, headers, rows }: { title: string; headers: string[]; rows: string[][] }): React.JSX.Element {
  return <Card className="p-3"><h2 className="mb-2 text-xs font-medium">{title}</h2><Table className="text-left text-xs"><TableHeader><TableRow>{headers.map((header) => <TableHead key={header}>{header}</TableHead>)}</TableRow></TableHeader><TableBody>{rows.map((row, index) => <TableRow key={`${row[0]}-${index}`}>{row.map((cell, cellIndex) => <TableCell key={`${cell}-${cellIndex}`}>{cell}</TableCell>)}</TableRow>)}</TableBody></Table>{!rows.length && <p className="py-5 text-center text-xs text-muted-foreground">No data.</p>}</Card>;
}
