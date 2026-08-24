import { useEffect, useMemo, useState } from "react";
import type * as React from "react";
import {
  calculateRaid,
  getRaidData,
  type RaidCalculation,
  type RaidDataPage,
  type RaidMethodDto,
  type RaidTargetDto,
} from "../lib/ipc.js";
import { Button } from "../components/ui/button.js";
import { Card } from "../components/ui/card.js";
import { Checkbox } from "../components/ui/checkbox.js";
import { Input } from "../components/ui/input.js";
import { Label } from "../components/ui/label.js";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "../components/ui/select.js";

type Mode = "LowestSulfur" | "LowestTotalResources" | "FewestRaidItems" | "Custom";

const modeLabels: Record<Mode, string> = {
  LowestSulfur: "Lowest sulfur",
  LowestTotalResources: "Lowest total resources",
  FewestRaidItems: "Fewest raid items",
  Custom: "Custom",
};

const formatAmount = (amount: number): string => Math.round(amount).toLocaleString();

function MethodCard({ method, highlighted }: { method: RaidMethodDto; highlighted?: boolean }): React.JSX.Element {
  const sulfur = method.resources.find((resource) => resource.shortname.toLowerCase() === "sulfur");
  return (
    <Card className={`rounded-lg p-3 ${highlighted ? "border-primary/70 bg-primary/5" : ""}`}>
      <div className="flex items-start gap-3">
        <div className="min-w-0 flex-1">
          <div className="flex items-center gap-2">
            <p className="truncate text-sm font-medium">{method.source.displayName}</p>
            {highlighted && <span className="rounded-full bg-primary/15 px-2 py-0.5 text-[10px] text-primary">recommended</span>}
          </div>
          <p className="mt-1 text-xs text-muted-foreground">
            {formatAmount(method.damagePerItem)} damage/item · {formatAmount(method.totalDamage)} total damage
          </p>
        </div>
        <span className="text-right text-sm font-semibold">{formatAmount(method.requiredItems)}×</span>
      </div>
      <div className="mt-2 flex flex-wrap gap-x-3 gap-y-1 text-[11px] text-muted-foreground">
        {sulfur && <span>{formatAmount(sulfur.amount)} sulfur</span>}
        {method.overkill > 0 && <span>{formatAmount(method.overkill)} overkill</span>}
        {!method.hasCraftCost && <span>no craft cost</span>}
      </div>
    </Card>
  );
}

function TargetButton({ target, selected, onClick }: { target: RaidTargetDto; selected: boolean; onClick: () => void }): React.JSX.Element {
  return (
    <Button
      type="button"
      variant="ghost"
      aria-pressed={selected}
      onClick={onClick}
      className={`h-auto w-full justify-start px-3 py-2 text-left ${selected ? "bg-primary/10 text-primary" : ""}`}
    >
      <span className="block truncate text-sm">{target.displayName}</span>
      <span className="mt-0.5 block text-[11px] text-muted-foreground">{target.category} · {formatAmount(target.startHealth)} HP</span>
    </Button>
  );
}

export function RaidCalculatorPage(): React.JSX.Element {
  const [data, setData] = useState<RaidDataPage | null>(null);
  const [targetId, setTargetId] = useState<number | null>(null);
  const [sourceIds, setSourceIds] = useState<number[]>([]);
  const [targetQuantity, setTargetQuantity] = useState(1);
  const [mode, setMode] = useState<Mode>("LowestSulfur");
  const [query, setQuery] = useState("");
  const [result, setResult] = useState<RaidCalculation | null>(null);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    void getRaidData()
      .then((loaded) => {
        setData(loaded);
        const first = loaded.targets.find((target) => target.displayName === "Armored Door") ?? loaded.targets[0];
        setTargetId(first?.targetId ?? null);
        setSourceIds(loaded.sources.map((source) => source.sourceId));
      })
      .catch((err: unknown) => setError(err instanceof Error ? err.message : String(err)));
  }, []);

  const target = data?.targets.find((item) => item.targetId === targetId) ?? null;
  const visibleTargets = useMemo(() => {
    const q = query.trim().toLowerCase();
    return (data?.targets ?? []).filter((item) => `${item.displayName} ${item.category} ${item.buildingTier ?? ""}`.toLowerCase().includes(q));
  }, [data, query]);

  useEffect(() => {
    if (!targetId || !data) return;
    let cancelled = false;
    void calculateRaid({ targetId, targetQuantity, sourceIds, mode })
      .then((calculated) => {
        if (!cancelled) {
          setResult(calculated);
          setError(null);
        }
      })
      .catch((err: unknown) => {
        if (!cancelled) setError(err instanceof Error ? err.message : String(err));
      });
    return () => {
      cancelled = true;
    };
  }, [data, mode, sourceIds, targetId, targetQuantity]);

  const toggleSource = (sourceId: number): void => {
    setSourceIds((current) => current.includes(sourceId) ? current.filter((id) => id !== sourceId) : [...current, sourceId]);
  };

  return (
    <div className="flex h-full min-h-0 flex-col">
      <header className="flex flex-wrap items-center gap-3 border-b px-4 py-2.5">
        <div>
          <h1 className="text-sm font-semibold">Raid Calculator</h1>
          <p className="text-[11px] text-muted-foreground">Dataset-backed damage and resource planning</p>
        </div>
        <div className="ml-auto flex items-center gap-2">
          <Label className="flex items-center gap-1 text-xs text-muted-foreground">
            Targets
            <Input
              type="number"
              min={1}
              max={100_000}
              value={targetQuantity}
              onChange={(event) => setTargetQuantity(Math.max(1, Number(event.target.value) || 1))}
              className="w-16 text-xs text-foreground"
            />
          </Label>
          <Select value={mode} onValueChange={(value) => setMode(value as Mode)}>
            <SelectTrigger className="w-52 text-xs"><SelectValue placeholder={modeLabels[mode]} /></SelectTrigger>
            <SelectContent>{Object.entries(modeLabels).map(([value, label]) => <SelectItem key={value} value={value}>{label}</SelectItem>)}</SelectContent>
          </Select>
        </div>
      </header>

      {error && <div className="border-b px-4 py-2 text-xs text-destructive">{error}</div>}

      <div className="flex min-h-0 flex-1">
        <aside className="flex w-64 min-w-0 flex-col border-r p-3">
          <Input
            value={query}
            onChange={(event) => setQuery(event.target.value)}
            placeholder="Search targets"
            aria-label="Search raid targets"
            className="mb-2 text-xs"
          />
          <div className="min-h-0 flex-1 space-y-0.5 overflow-y-auto">
            {visibleTargets.map((item) => <TargetButton key={item.targetId} target={item} selected={item.targetId === targetId} onClick={() => setTargetId(item.targetId)} />)}
          </div>
        </aside>

        <section className="min-w-0 flex-1 overflow-y-auto p-4">
          {!target ? (
            <p className="text-sm text-muted-foreground">Loading raid targets…</p>
          ) : (
            <>
              <div className="mb-4 flex flex-wrap items-end justify-between gap-3">
                <div>
                  <p className="text-lg font-semibold">{target.displayName}</p>
                  <p className="text-xs text-muted-foreground">{target.category} · {formatAmount(target.startHealth * targetQuantity)} total HP</p>
                </div>
                <div className="text-right text-xs text-muted-foreground">{result?.methods.length ?? 0} available methods</div>
              </div>

              <Card className="mb-4 p-3">
                <div className="mb-2 flex items-center justify-between">
                  <p className="text-xs font-medium">Raid items</p>
                  <div className="flex gap-2 text-[11px]">
                    <Button type="button" variant="ghost" size="sm" className="h-auto p-0 text-primary" onClick={() => setSourceIds(data?.sources.map((source) => source.sourceId) ?? [])}>all</Button>
                    <Button type="button" variant="ghost" size="sm" className="h-auto p-0 text-muted-foreground" onClick={() => setSourceIds([])}>none</Button>
                  </div>
                </div>
                <div className="flex flex-wrap gap-2">
                  {(data?.sources ?? []).map((source) => (
                    <Label key={source.sourceId} className="flex cursor-pointer items-center gap-1.5 rounded-md border px-2 py-1 text-xs hover:bg-muted">
                      <Checkbox checked={sourceIds.includes(source.sourceId)} onCheckedChange={() => toggleSource(source.sourceId)} />
                      {source.displayName}
                    </Label>
                  ))}
                </div>
              </Card>

              {result?.combination.length ? (
                <div className="mb-4 rounded-lg border border-primary/40 bg-primary/5 p-3">
                  <p className="mb-2 text-xs font-medium text-primary">Best combination · {modeLabels[mode]}</p>
                  <div className="space-y-2">{result.combination.map((method) => <MethodCard key={method.source.sourceId} method={method} highlighted />)}</div>
                  {result.resources.length > 0 && <p className="mt-3 text-xs text-muted-foreground">Resources: {result.resources.map((resource) => `${formatAmount(resource.amount)} ${resource.displayName}`).join(" · ")}</p>}
                </div>
              ) : null}

              <div>
                <p className="mb-2 text-xs font-medium">Individual methods</p>
                {result?.methods.length ? (
                  <div className="grid gap-2 md:grid-cols-2">{result.methods.map((method) => <MethodCard key={method.source.sourceId} method={method} highlighted={method.source.sourceId === result.recommended?.source.sourceId} />)}</div>
                ) : <p className="rounded-lg border border-dashed p-4 text-sm text-muted-foreground">Select at least one raid item.</p>}
              </div>
            </>
          )}
        </section>
      </div>
    </div>
  );
}
