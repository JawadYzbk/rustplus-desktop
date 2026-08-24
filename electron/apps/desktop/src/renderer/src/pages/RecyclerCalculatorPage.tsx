import { useEffect, useMemo, useState } from "react";
import type * as React from "react";
import {
  calculateRecycler,
  getRecyclerData,
  type RecyclerCalculation,
  type RecyclerItemDto,
  type RecyclerMetricDto,
  type RecyclerOutputDto,
} from "../lib/ipc.js";
import { Button } from "../components/ui/button.js";
import { Card } from "../components/ui/card.js";
import { Input } from "../components/ui/input.js";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "../components/ui/select.js";

const PAGE_SIZE = 30;
const amount = (value: number): string => Math.round(value).toLocaleString();

function Metric({ value, title }: { value: RecyclerMetricDto; title: string }): React.JSX.Element {
  const chance = value.chance > 0 ? `${amount(value.chance)} (${Math.round(value.chancePercent)}%)` : "0";
  return (
    <div title={value.min < value.max ? `Range: ${amount(value.min)} - ${amount(value.max)}` : undefined} className="rounded-md border bg-background/40 px-2 py-1 text-center">
      <p className="text-[10px] font-semibold uppercase text-muted-foreground">{title}</p>
      <p className="text-xs font-semibold">{amount(value.expected)}</p>
      <p className="text-[10px] text-muted-foreground">G {amount(value.guaranteed)} · C {chance}</p>
    </div>
  );
}

function OutputRow({ output }: { output: RecyclerOutputDto }): React.JSX.Element {
  const active = output.wild.expected > 0 || output.safe.expected > 0;
  return (
    <div className={`grid grid-cols-[minmax(0,1fr)_110px_110px] items-center gap-2 rounded-md px-2 py-1.5 ${active ? "bg-muted/40" : "opacity-30"}`}>
      <span className="truncate text-xs" title={output.shortName}>{output.displayName}</span>
      <Metric value={output.wild} title="Wild" />
      <Metric value={output.safe} title="Outpost" />
    </div>
  );
}

function ItemTile({ item, quantity, onChange }: { item: RecyclerItemDto; quantity: number; onChange: (quantity: number) => void }): React.JSX.Element {
  const change = (event: React.MouseEvent, delta: number): void => {
    event.preventDefault();
    onChange(Math.max(0, quantity + delta * (event.shiftKey ? 10 : 1)));
  };
  return (
    <div
      className="group rounded-md border bg-card p-2 hover:border-primary/60"
      title={`${item.displayName} · stack ${item.stackSize}`}
      onWheel={(event) => onChange(Math.max(0, quantity + (event.deltaY < 0 ? 1 : -1) * (event.shiftKey ? 10 : 1)))}
      onContextMenu={(event) => change(event, -1)}
    >
      <div className="flex h-10 items-center justify-center rounded bg-muted text-lg text-muted-foreground">{item.displayName.slice(0, 1)}</div>
      <p className="mt-1 truncate text-center text-[10px]" title={item.displayName}>{item.displayName}</p>
      <div className="mt-1 flex items-center gap-1">
        <Button type="button" variant="outline" size="icon" aria-label={`Decrease ${item.displayName}`} onClick={(event) => change(event, -1)} className="h-6 w-6 text-xs">−</Button>
        <Input
          aria-label={`${item.displayName} quantity`}
          type="number"
          min={0}
          value={quantity}
          onChange={(event) => onChange(Math.max(0, Math.floor(Number(event.target.value) || 0)))}
          className="min-w-0 flex-1 px-1 py-1 text-center text-[11px]"
        />
        <Button type="button" variant="outline" size="icon" aria-label={`Increase ${item.displayName}`} onClick={(event) => change(event, 1)} className="h-6 w-6 text-xs">+</Button>
      </div>
    </div>
  );
}

export function RecyclerCalculatorPage(): React.JSX.Element {
  const [items, setItems] = useState<RecyclerItemDto[]>([]);
  const [quantities, setQuantities] = useState<Record<string, number>>({});
  const [query, setQuery] = useState("");
  const [category, setCategory] = useState("all");
  const [visibleCount, setVisibleCount] = useState(PAGE_SIZE);
  const [calculation, setCalculation] = useState<RecyclerCalculation | null>(null);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    void getRecyclerData()
      .then(setItems)
      .catch((err: unknown) => setError(err instanceof Error ? err.message : String(err)));
  }, []);

  const categories = useMemo(() => ["all", ...new Set(items.map((item) => item.category).filter(Boolean).sort())], [items]);
  const filteredItems = useMemo(() => {
    const q = query.trim().toLowerCase();
    return items.filter((item) => (category === "all" || item.category === category) && `${item.displayName} ${item.shortName}`.toLowerCase().includes(q));
  }, [category, items, query]);
  const visibleItems = filteredItems.slice(0, visibleCount);
  const selectedQuantityCount = useMemo(() => Object.values(quantities).filter((value) => value > 0).length, [quantities]);

  useEffect(() => {
    setVisibleCount(PAGE_SIZE);
  }, [category, query]);

  useEffect(() => {
    if (items.length === 0) return;
    let cancelled = false;
    const entries = Object.entries(quantities)
      .filter(([, quantity]) => quantity > 0)
      .map(([shortName, quantity]) => ({ shortName, quantity }));
    void calculateRecycler(entries)
      .then((result) => {
        if (!cancelled) {
          setCalculation(result);
          setError(null);
        }
      })
      .catch((err: unknown) => {
        if (!cancelled) setError(err instanceof Error ? err.message : String(err));
      });
    return () => {
      cancelled = true;
    };
  }, [items.length, quantities]);

  const setQuantity = (shortName: string, quantity: number): void => {
    setQuantities((current) => ({ ...current, [shortName]: Math.max(0, Math.floor(quantity)) }));
  };

  const fillStacks = (): void => {
    setQuantities((current) => ({
      ...current,
      ...Object.fromEntries(visibleItems.map((item) => [item.shortName, item.stackSize])),
    }));
  };

  return (
    <div className="flex h-full min-h-0 flex-col">
      <header className="flex flex-wrap items-center gap-3 border-b px-4 py-2.5">
        <div>
          <h1 className="text-sm font-semibold">Recycler Calculator</h1>
          <p className="text-[11px] text-muted-foreground">Wild recycler and safe-zone yield planning</p>
        </div>
        <div className="ml-auto flex items-center gap-2 text-xs text-muted-foreground">
          {selectedQuantityCount} selected
          <Button type="button" variant="outline" size="sm" onClick={() => setQuantities({})}>Clear all</Button>
          <Button type="button" variant="outline" size="sm" onClick={fillStacks}>Fill stacks</Button>
        </div>
      </header>
      {error && <div className="border-b px-4 py-2 text-xs text-destructive">{error}</div>}

      <div className="flex min-h-0 flex-1 flex-col gap-3 overflow-y-auto p-3">
        <Card className="p-3">
          <div className="mb-2 flex flex-wrap gap-2">
            <Input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="Search recyclable items" aria-label="Search recyclable items" className="min-w-48 flex-1 text-xs" />
            <Select value={category} onValueChange={setCategory}>
              <SelectTrigger aria-label="Recycler category" className="text-xs"><SelectValue placeholder="All categories" /></SelectTrigger>
              <SelectContent>{categories.map((value) => <SelectItem key={value} value={value}>{value === "all" ? "All categories" : value}</SelectItem>)}</SelectContent>
            </Select>
          </div>
          <div className="grid grid-cols-2 gap-2 sm:grid-cols-4 md:grid-cols-5 lg:grid-cols-6 xl:grid-cols-8">
            {visibleItems.map((item) => <ItemTile key={item.shortName} item={item} quantity={quantities[item.shortName] ?? 0} onChange={(value) => setQuantity(item.shortName, value)} />)}
          </div>
          {visibleCount < filteredItems.length && <Button type="button" variant="outline" onClick={() => setVisibleCount((count) => count + PAGE_SIZE)} className="mt-3 w-full border-dashed text-xs text-muted-foreground">Load more ({filteredItems.length - visibleCount} remaining)</Button>}
          {items.length === 0 && <p className="p-6 text-center text-sm text-muted-foreground">Loading recycler data…</p>}
        </Card>

        <Card className="min-h-0 p-3">
          <div className="mb-2 flex flex-wrap items-center justify-between gap-2">
            <div>
              <p className="text-xs font-medium">Yield board</p>
              <p className="text-[11px] text-muted-foreground">G = guaranteed · C = chance</p>
            </div>
            <div className="flex gap-2 text-xs">
              <span className="rounded-md border px-2 py-1 text-emerald-400">Wild {formatTime(calculation?.wildSeconds ?? 0)}</span>
              <span className="rounded-md border px-2 py-1 text-orange-300">Outpost {formatTime(calculation?.safeSeconds ?? 0)}</span>
            </div>
          </div>
          <div className="grid grid-cols-[minmax(0,1fr)_110px_110px] gap-2 px-2 pb-1 text-[10px] font-semibold uppercase text-muted-foreground">
            <span>Output</span><span className="text-center">Wild</span><span className="text-center">Outpost</span>
          </div>
          <div className="space-y-0.5">
            {(calculation?.outputs ?? []).map((output) => <OutputRow key={output.shortName} output={output} />)}
            {calculation && calculation.outputs.length === 0 && <p className="p-5 text-center text-sm text-muted-foreground">Set quantities above to calculate yields.</p>}
          </div>
        </Card>
      </div>
    </div>
  );
}

function formatTime(seconds: number): string {
  const total = Math.max(0, Math.floor(seconds));
  const hours = Math.floor(total / 3600).toString().padStart(2, "0");
  const minutes = Math.floor((total % 3600) / 60).toString().padStart(2, "0");
  const remainder = (total % 60).toString().padStart(2, "0");
  return `${hours}:${minutes}:${remainder}`;
}
