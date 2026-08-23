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
        <label className="flex items-center gap-2 text-xs font-medium">
          <input
            type="checkbox"
            checked={engineActive}
            onChange={(e) => {
              setEngineActive(e.target.checked);
              void persist(e.target.checked, rules);
            }}
          />
          Logic Engine active
        </label>
        <button
          type="button"
          className="rounded-md border px-2 py-1 text-xs hover:bg-muted"
          onClick={() => {
            setRules((r) => [...r, newHeader(r.length + 1)]);
            setDirty(true);
          }}
        >
          + Rule
        </button>
        <span className="ml-auto flex items-center gap-2 text-xs text-muted-foreground">
          {status?.isRunning === true && (
            <>
              <span className="rounded-full bg-emerald-500/15 px-2 py-0.5 text-[10px] text-emerald-600">
                {status.currentRuleName} · step {status.currentStepNumber} ({status.currentStepType})
              </span>
              <button type="button" className="rounded border px-2 py-0.5 hover:bg-muted" onClick={() => void stopLogic()}>
                Stop
              </button>
            </>
          )}
          {dirty && (
            <button
              type="button"
              disabled={saving}
              className="rounded-md border border-primary bg-primary/10 px-2 py-0.5 text-primary hover:bg-primary/20 disabled:opacity-50"
              onClick={() => void persist(engineActive, rules)}
            >
              {saving ? "Saving…" : "Save changes"}
            </button>
          )}
        </span>
      </div>

      <div className="min-h-0 flex-1 overflow-y-auto p-2">
        {rules.length === 0 && (
          <p className="p-4 text-xs text-muted-foreground">No rules yet — create one to automate switches and alarms.</p>
        )}
        {rules.map((rule, i) => (
          <details
            key={rule.id}
            className="mb-1 rounded-md border px-2 py-1.5"
            open={openRuleId === rule.id || (i === 0 && dirty)}
            onToggle={(e) => setOpenRuleId(e.currentTarget.open ? rule.id : null)}
          >
            <summary className="flex cursor-pointer items-center gap-2 text-sm">
              <input
                type="checkbox"
                checked={rule.isEnabled}
                onChange={(e) => {
                  setRules((rs) => rs.map((r, j) => (j === i ? { ...r, isEnabled: e.target.checked } : r)));
                  setDirty(true);
                }}
              />
              <span className={rule.isEnabled ? "" : "text-muted-foreground"}>{rule.name}</span>
              <span className="ml-auto text-[11px] text-muted-foreground">
                {rule.triggerType}
                {rule.triggerType === "ChatCommand" ? ` · !${rule.triggerCommand}` : ""}
                {` · ${rule.loopCount}×`}
                <button
                  type="button"
                  className="ml-2 rounded border px-1.5 py-0.5 hover:bg-muted"
                  onClick={(e) => {
                    e.preventDefault();
                    void runRule(rule.id);
                  }}
                >
                  Run
                </button>
                <button
                  type="button"
                  className="ml-1 rounded border px-1.5 py-0.5 text-destructive hover:bg-destructive/10"
                  onClick={(e) => {
                    e.preventDefault();
                    setRules((rs) => rs.filter((_, j) => j !== i));
                    setDirty(true);
                  }}
                >
                  ✕
                </button>
              </span>
            </summary>
            <div className="grid grid-cols-2 gap-2 px-6 py-2 text-xs">
              <label className="col-span-2 flex items-center gap-2">
                Name
                <input
                  value={rule.name}
                  onChange={(e) => {
                    setRules((rs) => rs.map((r, j) => (j === i ? { ...r, name: e.target.value } : r)));
                    setDirty(true);
                  }}
                  className="flex-1 rounded border bg-transparent px-2 py-1"
                />
              </label>
              <label className="flex items-center gap-2">
                Trigger
                <select
                  value={rule.triggerType}
                  onChange={(e) => {
                    setRules((rs) =>
                      rs.map((r, j) => (j === i ? { ...r, triggerType: e.target.value as typeof r.triggerType } : r)),
                    );
                    setDirty(true);
                  }}
                  className="flex-1 rounded border bg-transparent px-2 py-1"
                >
                  {TRIGGERS.map((t) => (
                    <option key={t}>{t}</option>
                  ))}
                </select>
              </label>
              <label className="flex items-center gap-2">
                Loop count
                <input
                  type="number"
                  min={0}
                  value={rule.loopCount}
                  onChange={(e) => {
                    setRules((rs) =>
                      rs.map((r, j) => (j === i ? { ...r, loopCount: Math.max(0, Number(e.target.value) || 0) } : r)),
                    );
                    setDirty(true);
                  }}
                  className="w-20 rounded border bg-transparent px-2 py-1"
                />
              </label>
              {(rule.triggerType === "SmartAlarm" || rule.triggerType === "SmartSwitch") && (
                <label className="flex items-center gap-2">
                  Trigger entity #
                  <input
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
                    className="w-24 rounded border bg-transparent px-2 py-1"
                  />
                </label>
              )}
              {rule.triggerType === "ChatCommand" && (
                <label className="flex items-center gap-2">
                  Command
                  <input
                    value={rule.triggerCommand}
                    onChange={(e) => {
                      setRules((rs) => rs.map((r, j) => (j === i ? { ...r, triggerCommand: e.target.value } : r)));
                      setDirty(true);
                    }}
                    className="flex-1 rounded border bg-transparent px-2 py-1"
                  />
                </label>
              )}
            </div>
            {openRuleId === rule.id && (
              <StepEditor matchKey={matchKey} ruleId={rule.id} onSaved={() => void refresh()} />
            )}
          </details>
        ))}
      </div>
    </div>
  );
}
