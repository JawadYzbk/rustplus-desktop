/**
 * Logic Engine panel (stage 5) — rule list + header editing over logic/* IPC, with the
 * full step editor mounted when a rule is expanded (logic/getRule|saveRule).
 */
import { useCallback, useEffect, useState } from "react";
import type * as React from "react";
import {
  getLogicStatus,
  getRules,
  runRule,
  saveRules,
  stopLogic,
  type LogicStatus,
  type RuleHeaderInput,
} from "../lib/ipc.js";
import { StepEditor } from "./StepEditor.js";
import { Badge } from "../components/ui/badge.js";
import { Button } from "../components/ui/button.js";
import { Checkbox } from "../components/ui/checkbox.js";
import { Input } from "../components/ui/input.js";
import { Label } from "../components/ui/label.js";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "../components/ui/select.js";
import { Accordion, AccordionContent, AccordionItem, AccordionTrigger } from "../components/ui/accordion.js";

const TRIGGERS = ["SmartAlarm", "SmartSwitch", "ChatCommand", "RuleTriggered", "RuleCompleted"] as const;

function newHeader(n: number): RuleHeaderInput {
  return {
    id: crypto.randomUUID(),
    name: `New Rule ${n}`,
    isEnabled: false,
    isLoopEnabled: false,
    loopCount: 1,
    triggerType: "SmartAlarm",
    triggerEntityId: 0,
    triggerCommand: "rulecommand",
    triggerRuleId: "",
    triggerState: true,
    conditionOperator: "NONE",
    conditionDeviceEntityId: 0,
    conditionDeviceState: true,
  };
}

export function RulesPanel({ matchKey }: { matchKey: string }): React.JSX.Element {
  const [engineActive, setEngineActive] = useState(false);
  const [rules, setRules] = useState<RuleHeaderInput[]>([]);
  const [status, setStatus] = useState<LogicStatus | null>(null);
  const [dirty, setDirty] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [saving, setSaving] = useState(false);
  const [openRuleId, setOpenRuleId] = useState<string | null>(null);

  const refresh = useCallback(async () => {
    try {
      const data = await getRules(matchKey);
      setEngineActive(data.isEngineActive);
      setRules(data.rules.map(({ stepCount: _drop, ...header }) => header));
      setStatus(await getLogicStatus());
      setDirty(false);
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    }
  }, [matchKey]);

  useEffect(() => {
    void refresh();
  }, [refresh]);

  // Light status polling keeps the running indicator honest without push plumbing.
  useEffect(() => {
    const t = setInterval(() => {
      void getLogicStatus().then(setStatus).catch(() => undefined);
    }, 2000);
    return () => clearInterval(t);
  }, []);

  const persist = async (nextActive: boolean, nextRules: RuleHeaderInput[]): Promise<void> => {
    setSaving(true);
    try {
      const ok = await saveRules(matchKey, nextActive, nextRules);
      if (!ok) throw new Error("save rejected");
      await refresh();
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setSaving(false);
    }
  };

  if (error) return <div className="p-3 text-xs text-destructive">{error}</div>;

  return (
    <div className="flex min-h-0 flex-col">
      <div className="flex items-center gap-3 border-b px-4 py-2">
        <Label className="flex items-center gap-2 text-xs font-medium">
          <Checkbox
            checked={engineActive}
            onCheckedChange={(checked) => {
              setEngineActive(checked === true);
              void persist(checked === true, rules);
            }}
          />
          Logic Engine active
        </Label>
        <Button
          type="button"
          variant="outline"
          size="sm"
          onClick={() => {
            setRules((r) => [...r, newHeader(r.length + 1)]);
            setDirty(true);
          }}
        >
          + Rule
        </Button>
        <span className="ml-auto flex items-center gap-2 text-xs text-muted-foreground">
          {status?.isRunning === true && (
            <>
              <Badge variant="outline" className="rounded-full bg-emerald-500/15 px-2 py-0.5 text-[10px] text-emerald-600">
                {status.currentRuleName} · step {status.currentStepNumber} ({status.currentStepType})
              </Badge>
              <Button type="button" variant="outline" size="sm" onClick={() => void stopLogic()}>
                Stop
              </Button>
            </>
          )}
          {dirty && (
            <Button
              type="button"
              variant="outline"
              size="sm"
              disabled={saving}
              className="rounded-md border border-primary bg-primary/10 px-2 py-0.5 text-primary hover:bg-primary/20 disabled:opacity-50"
              onClick={() => void persist(engineActive, rules)}
            >
              {saving ? "Saving…" : "Save changes"}
            </Button>
          )}
        </span>
      </div>

      <div className="min-h-0 flex-1 overflow-y-auto p-2">
        {rules.length === 0 && (
          <p className="p-4 text-xs text-muted-foreground">No rules yet — create one to automate switches and alarms.</p>
        )}
        <Accordion type="single" collapsible value={openRuleId ?? (dirty ? rules[0]?.id : undefined)} onValueChange={(value) => setOpenRuleId(value || null)}>
        {rules.map((rule, i) => (
          <AccordionItem key={rule.id} value={rule.id} className="mb-1 rounded-md border px-2">
            <div className="flex items-center gap-2">
              <Checkbox
                checked={rule.isEnabled}
                onCheckedChange={(checked) => {
                  setRules((rs) => rs.map((r, j) => (j === i ? { ...r, isEnabled: checked === true } : r)));
                  setDirty(true);
                }}
              />
              <AccordionTrigger className="flex-1 py-2 text-sm hover:no-underline">
                <span className={rule.isEnabled ? "" : "text-muted-foreground"}>{rule.name}</span>
                <span className="ml-auto mr-2 text-[11px] text-muted-foreground">
                  {rule.triggerType}
                  {rule.triggerType === "ChatCommand" ? ` · !${rule.triggerCommand}` : ""}
                  {` · ${rule.loopCount}×`}
                </span>
              </AccordionTrigger>
              <Button type="button" variant="outline" size="sm" onClick={() => void runRule(rule.id)}>Run</Button>
              <Button type="button" variant="ghost" size="icon" className="h-8 w-8 text-destructive" onClick={() => { setRules((rs) => rs.filter((_, j) => j !== i)); setDirty(true); }}>✕</Button>
            </div>
            <AccordionContent>
            <div className="grid grid-cols-2 gap-2 px-6 py-2 text-xs">
              <Label className="col-span-2 flex items-center gap-2">
                Name
                <Input
                  value={rule.name}
                  onChange={(e) => {
                    setRules((rs) => rs.map((r, j) => (j === i ? { ...r, name: e.target.value } : r)));
                    setDirty(true);
                  }}
                  className="flex-1"
                />
              </Label>
              <Label className="flex items-center gap-2">
                Trigger
                <Select value={rule.triggerType} onValueChange={(value) => {
                    setRules((rs) =>
                      rs.map((r, j) => (j === i ? { ...r, triggerType: value as typeof r.triggerType } : r)),
                    );
                    setDirty(true);
                  }}>
                  <SelectTrigger className="flex-1"><SelectValue /></SelectTrigger>
                  <SelectContent>{TRIGGERS.map((t) => <SelectItem key={t} value={t}>{t}</SelectItem>)}</SelectContent>
                </Select>
              </Label>
              <Label className="flex items-center gap-2">
                Loop count
                <Input
                  type="number"
                  min={0}
                  value={rule.loopCount}
                  onChange={(e) => {
                    setRules((rs) =>
                      rs.map((r, j) => (j === i ? { ...r, loopCount: Math.max(0, Number(e.target.value) || 0) } : r)),
                    );
                    setDirty(true);
                  }}
                  className="w-20"
                />
              </Label>
              {(rule.triggerType === "SmartAlarm" || rule.triggerType === "SmartSwitch") && (
                <Label className="flex items-center gap-2">
                  Trigger entity #
                  <Input
                    type="number"
                    min={0}
                    value={rule.triggerEntityId}
                    onChange={(e) => {
                      setRules((rs) =>
                        rs.map((r, j) =>
                          j === i ? { ...r, triggerEntityId: Math.max(0, Number(e.target.value) || 0) } : r,
                        ),
                      );
                      setDirty(true);
                    }}
                    className="w-24"
                  />
                </Label>
              )}
              {rule.triggerType === "ChatCommand" && (
                <Label className="flex items-center gap-2">
                  Command
                  <Input
                    value={rule.triggerCommand}
                    onChange={(e) => {
                      setRules((rs) => rs.map((r, j) => (j === i ? { ...r, triggerCommand: e.target.value } : r)));
                      setDirty(true);
                    }}
                    className="flex-1"
                  />
                </Label>
              )}
            </div>
            {openRuleId === rule.id && (
              <StepEditor matchKey={matchKey} ruleId={rule.id} onSaved={() => void refresh()} />
            )}
            </AccordionContent>
          </AccordionItem>
        ))}
        </Accordion>
      </div>
    </div>
  );
}
