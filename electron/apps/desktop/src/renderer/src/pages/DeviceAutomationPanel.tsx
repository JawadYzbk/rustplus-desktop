/** Device Automation editor — profile rules backed by deviceAutomation/* IPC. */
import { useCallback, useEffect, useMemo, useState } from "react";
import type * as React from "react";
import {
  getDeviceAutomationRules,
  saveDeviceAutomationRules,
  type DeviceAutomationRuleDto,
  type DeviceNode,
} from "../lib/ipc.js";
import { deviceLabel, useProfilesStore } from "../stores/profiles.js";
import { Badge } from "../components/ui/badge.js";
import { Button } from "../components/ui/button.js";
import { Card } from "../components/ui/card.js";
import { Checkbox } from "../components/ui/checkbox.js";
import { Input } from "../components/ui/input.js";
import { Label } from "../components/ui/label.js";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "../components/ui/select.js";

const MATCH_MODES: DeviceAutomationRuleDto["playerMatchMode"][] = [
  "AnyOnline",
  "AllOnline",
  "Specific",
  "AnyOffline",
  "AllOffline",
  "SpecificOffline",
];

function newRule(index: number): DeviceAutomationRuleDto {
  return {
    id: crypto.randomUUID(),
    name: `Automation ${index}`,
    isEnabled: false,
    isExpanded: true,
    conditionType: "PlayerProximity",
    playerMatchMode: "AnyOnline",
    specificPlayerSteamId: "",
    locationEntityId: 0,
    distanceMeters: 250,
    startTime: "20:00",
    endTime: "08:00",
    targetEntityId: 0,
    matchedState: false,
    unmatchedState: true,
  };
}

function leaves(nodes: readonly DeviceNode[]): DeviceNode[] {
  return nodes.flatMap((node) => (node.isGroup ? leaves(node.children) : [node]));
}

function stateLabel(value: boolean): string {
  return value ? "ON" : "OFF";
}

const inputClass = "text-xs";

export function DeviceAutomationPanel({ matchKey }: { matchKey: string }): React.JSX.Element {
  const devices = useProfilesStore((s) => s.devices[matchKey] ?? []);
  const [active, setActive] = useState(false);
  const [rules, setRules] = useState<DeviceAutomationRuleDto[]>([]);
  const [dirty, setDirty] = useState(false);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const allDevices = useMemo(() => leaves(devices), [devices]);
  const switches = useMemo(
    () => allDevices.filter((d) => d.kind?.replace(/ /g, "").toLowerCase() === "smartswitch"),
    [allDevices],
  );

  const refresh = useCallback(async () => {
    setError(null);
    try {
      const data = await getDeviceAutomationRules(matchKey);
      setActive(data.isActive);
      setRules(data.rules);
      setDirty(false);
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    }
  }, [matchKey]);

  useEffect(() => {
    void refresh();
  }, [refresh]);

  const persist = async (nextActive = active, nextRules = rules): Promise<void> => {
    setSaving(true);
    setError(null);
    try {
      if (!(await saveDeviceAutomationRules(matchKey, nextActive, nextRules))) {
        throw new Error("save rejected");
      }
      setDirty(false);
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setSaving(false);
    }
  };

  const patchRule = (index: number, patch: Partial<DeviceAutomationRuleDto>): void => {
    setRules((current) => current.map((rule, i) => (i === index ? { ...rule, ...patch } : rule)));
    setDirty(true);
  };

  const addRule = (): void => {
    const next = [...rules, newRule(rules.length + 1)];
    setRules(next);
    void persist(active, next);
  };

  const removeRule = (index: number): void => {
    const next = rules.filter((_, i) => i !== index);
    setRules(next);
    void persist(active, next);
  };

  return (
    <div className="flex min-h-0 flex-1 flex-col bg-background">
      <header className="flex items-center gap-3 border-b border-emerald-500/20 px-5 py-3">
        <div>
          <p className="text-[10px] font-semibold uppercase tracking-[0.2em] text-emerald-400">Automation</p>
          <h2 className="text-base font-semibold">Device rules</h2>
        </div>
        <span className="hidden text-xs text-muted-foreground sm:block">
          Switches react to player proximity or in-game time.
        </span>
        <div className="ml-auto flex items-center gap-2">
          {dirty && <span className="text-[11px] text-amber-400">Unsaved</span>}
          <Button
            type="button"
            variant="outline"
            size="sm"
            disabled={!dirty || saving}
            onClick={() => void persist()}
            className="border-primary/60 bg-primary/10 text-primary"
          >
            {saving ? "Saving…" : "Save changes"}
          </Button>
          <Button
            type="button"
            size="sm"
            onClick={addRule}
            className="bg-emerald-500 text-slate-950 hover:bg-emerald-400"
          >
            + Add rule
          </Button>
        </div>
      </header>

      <div className="flex items-center gap-3 border-b px-5 py-3">
        <Label className="flex items-center gap-2 text-xs font-medium">
          <Checkbox
            checked={active}
            onCheckedChange={(checked) => {
              const next = checked === true;
              setActive(next);
              void persist(next, rules);
            }}
          />
          Run device automation
        </Label>
        <span className="text-[11px] text-muted-foreground">
          {rules.filter((rule) => rule.isEnabled).length}/{rules.length} rules enabled
        </span>
      </div>

      {error && <p className="border-b border-destructive/20 px-5 py-2 text-xs text-destructive">{error}</p>}

      <div className="min-h-0 flex-1 overflow-y-auto p-4 sm:p-5">
        {rules.length === 0 ? (
          <div className="rounded-lg border border-dashed border-emerald-500/30 bg-emerald-500/[0.04] px-5 py-8 text-center">
            <p className="text-sm font-medium">No device rules yet</p>
            <p className="mt-1 text-xs text-muted-foreground">Add one to control a Smart Switch automatically.</p>
          </div>
        ) : (
          <div className="mx-auto max-w-3xl space-y-3">
            {rules.map((rule, index) => (
              <RuleCard
                key={rule.id}
                rule={rule}
                index={index}
                devices={allDevices}
                switches={switches}
                onChange={patchRule}
                onDelete={removeRule}
              />
            ))}
          </div>
        )}
      </div>
    </div>
  );
}

function RuleCard({
  rule,
  index,
  devices,
  switches,
  onChange,
  onDelete,
}: {
  rule: DeviceAutomationRuleDto;
  index: number;
  devices: DeviceNode[];
  switches: DeviceNode[];
  onChange: (index: number, patch: Partial<DeviceAutomationRuleDto>) => void;
  onDelete: (index: number) => void;
}): React.JSX.Element {
  const set = (patch: Partial<DeviceAutomationRuleDto>): void => onChange(index, patch);
  const specific = rule.playerMatchMode.startsWith("Specific");
  return (
    <Card className="overflow-hidden rounded-lg border-white/[0.08] shadow-[0_12px_32px_rgba(0,0,0,0.14)]">
      <div className="flex items-center gap-3 px-4 py-3">
        <Checkbox
          checked={rule.isEnabled}
          onCheckedChange={(checked) => set({ isEnabled: checked === true })}
          aria-label={`Enable ${rule.name}`}
        />
        <Input
          value={rule.name}
          onChange={(event) => set({ name: event.target.value || "Automation" })}
          className="h-8 min-w-0 flex-1 border-transparent bg-transparent px-0 text-sm font-semibold shadow-none focus-visible:border-emerald-400"
          aria-label="Rule name"
        />
        <Badge variant="outline" className={rule.isEnabled ? "text-[10px] uppercase tracking-wider text-emerald-400" : "text-[10px] uppercase tracking-wider text-muted-foreground"}>
          {rule.isEnabled ? "live" : "off"}
        </Badge>
        <Button
          type="button"
          variant="ghost"
          size="icon"
          onClick={() => set({ isExpanded: !rule.isExpanded })}
          className="h-8 w-8 text-muted-foreground"
          aria-label={rule.isExpanded ? "Collapse rule" : "Expand rule"}
        >
          {rule.isExpanded ? "⌃" : "⌄"}
        </Button>
        <Button
          type="button"
          variant="ghost"
          size="icon"
          onClick={() => onDelete(index)}
          className="h-8 w-8 text-muted-foreground hover:bg-destructive/15 hover:text-destructive"
          aria-label={`Delete ${rule.name}`}
        >
          ×
        </Button>
      </div>

      {rule.isExpanded && (
        <div className="border-t border-white/[0.06] px-4 pb-4 pt-3">
          <div className="mb-3 grid gap-3 sm:grid-cols-2">
            <Label className="grid gap-1 text-[11px] uppercase tracking-wide text-emerald-400">
              When
              <Select value={rule.conditionType} onValueChange={(value) => set({ conditionType: value as DeviceAutomationRuleDto["conditionType"] })}>
                <SelectTrigger className={inputClass}><SelectValue /></SelectTrigger>
                <SelectContent><SelectItem value="PlayerProximity">Player proximity</SelectItem><SelectItem value="GameTime">In-game time</SelectItem></SelectContent>
              </Select>
            </Label>
            {rule.conditionType === "PlayerProximity" ? (
              <Label className="grid gap-1 text-[11px] uppercase tracking-wide text-muted-foreground">
                Anchor device
                <Select value={String(rule.locationEntityId)} onValueChange={(value) => set({ locationEntityId: Number(value) })}>
                  <SelectTrigger className={inputClass}><SelectValue placeholder="Choose a device…" /></SelectTrigger>
                  <SelectContent><SelectItem value="0">Choose a device…</SelectItem>{devices.map((device) => <SelectItem key={device.entityId} value={String(device.entityId)}>{deviceLabel(device)}</SelectItem>)}</SelectContent>
                </Select>
              </Label>
            ) : (
              <div className="grid grid-cols-2 gap-2">
                <Label className="grid gap-1 text-[11px] uppercase tracking-wide text-muted-foreground">
                  From
                  <Input type="time" value={rule.startTime} onChange={(event) => set({ startTime: event.target.value })} className={inputClass} />
                </Label>
                <Label className="grid gap-1 text-[11px] uppercase tracking-wide text-muted-foreground">
                  Until
                  <Input type="time" value={rule.endTime} onChange={(event) => set({ endTime: event.target.value })} className={inputClass} />
                </Label>
              </div>
            )}
          </div>

          {rule.conditionType === "PlayerProximity" && (
            <div className="mb-3 grid gap-3 sm:grid-cols-[1fr_110px]">
              <Label className="grid gap-1 text-[11px] uppercase tracking-wide text-muted-foreground">
                Player condition
                <Select value={rule.playerMatchMode} onValueChange={(value) => set({ playerMatchMode: value as DeviceAutomationRuleDto["playerMatchMode"] })}>
                  <SelectTrigger className={inputClass}><SelectValue /></SelectTrigger>
                  <SelectContent>{MATCH_MODES.map((mode) => <SelectItem key={mode} value={mode}>{mode.replace(/([a-z])([A-Z])/g, "$1 $2")}</SelectItem>)}</SelectContent>
                </Select>
              </Label>
              <Label className="grid gap-1 text-[11px] uppercase tracking-wide text-muted-foreground">
                Radius (m)
                <Input type="number" min={1} value={rule.distanceMeters} onChange={(event) => set({ distanceMeters: Math.max(1, Number(event.target.value) || 1) })} className={inputClass} />
              </Label>
              {specific && (
                <Label className="grid gap-1 text-[11px] uppercase tracking-wide text-muted-foreground sm:col-span-2">
                  Steam ID
                  <Input value={rule.specificPlayerSteamId} onChange={(event) => set({ specificPlayerSteamId: event.target.value })} placeholder="7656119…" className={inputClass} inputMode="numeric" />
                </Label>
              )}
            </div>
          )}

          <div className="grid gap-3 sm:grid-cols-[1fr_120px_120px]">
            <Label className="grid gap-1 text-[11px] uppercase tracking-wide text-emerald-400">
              Then switch
              <Select value={String(rule.targetEntityId)} onValueChange={(value) => set({ targetEntityId: Number(value) })}>
                <SelectTrigger className={inputClass}><SelectValue placeholder="Choose a Smart Switch…" /></SelectTrigger>
                <SelectContent><SelectItem value="0">Choose a Smart Switch…</SelectItem>{switches.map((device) => <SelectItem key={device.entityId} value={String(device.entityId)}>{deviceLabel(device)}</SelectItem>)}</SelectContent>
              </Select>
            </Label>
            <StateSelect label="Matched" value={rule.matchedState} onChange={(value) => set({ matchedState: value })} />
            <StateSelect label="Otherwise" value={rule.unmatchedState} onChange={(value) => set({ unmatchedState: value })} />
          </div>
        </div>
      )}
    </Card>
  );
}

function StateSelect({ label, value, onChange }: { label: string; value: boolean; onChange: (value: boolean) => void }): React.JSX.Element {
  return (
    <Label className="grid gap-1 text-[11px] uppercase tracking-wide text-muted-foreground">
      {label}
      <Select value={String(value)} onValueChange={(next) => onChange(next === "true")}>
        <SelectTrigger className={inputClass}><SelectValue /></SelectTrigger>
        <SelectContent><SelectItem value="true">{stateLabel(true)}</SelectItem><SelectItem value="false">{stateLabel(false)}</SelectItem></SelectContent>
      </Select>
    </Label>
  );
}
