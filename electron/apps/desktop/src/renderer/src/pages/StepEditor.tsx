/**
 * Rule step editor (stage 5) — reads/writes one full rule over logic/getRule|saveRule.
 * Per-step fields follow Models/LogicRule.cs: Wait seconds, Toggle target/group +
 * inherit/ON/OFF state, StartTimer (Custom minutes/name or rig targets + crate marker),
 * CheckAvailability gate with operator + CSV ids and inline conditional steps.
 * parseLogicRule on the main side re-applies C# defaults and clamps on save.
 */
import { useCallback, useEffect, useState } from "react";
import type * as React from "react";
import { getRule, saveRule, type FullRuleDto, type StepDto } from "../lib/ipc.js";
import { Button } from "../components/ui/button.js";
import { Checkbox } from "../components/ui/checkbox.js";
import { Input } from "../components/ui/input.js";
import { Label } from "../components/ui/label.js";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "../components/ui/select.js";

const STEP_TYPES = ["Wait", "Toggle", "CheckAvailability", "StartTimer"] as const;
const OPERATORS = ["IS_OFFLINE", "IS_ONLINE", "ALL_OFFLINE", "ANY_OFFLINE", "ALL_ONLINE", "ANY_ONLINE"] as const;
const TIMER_TARGETS = ["Custom", "SmallOilRig", "LargeOilRig"] as const;

function newStep(type: (typeof STEP_TYPES)[number]): StepDto {
  const base: StepDto = { stepType: type };
  if (type === "Wait") base.waitSeconds = 10;
  if (type === "Toggle") {
    base.targetEntityId = 0;
    base.targetGroupName = "";
    base.toggleState = null;
  }
  if (type === "StartTimer") {
    base.timerTarget = "Custom";
    base.timerMinutes = 15;
    base.timerName = "";
    base.showCrateOnMap = true;
  }
  if (type === "CheckAvailability") {
    base.conditionOperator = "ALL_OFFLINE";
    base.conditionDeviceIdsCsv = "";
    base.conditionalSteps = [];
  }
  return base;
}

const inputCls = "text-xs";

export function StepEditor({
  matchKey,
  ruleId,
  onSaved,
}: {
  matchKey: string;
  ruleId: string;
  onSaved?: () => void;
}): React.JSX.Element {
  const [rule, setRule] = useState<FullRuleDto | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [saving, setSaving] = useState(false);

  useEffect(() => {
    let cancelled = false;
    void getRule(matchKey, ruleId)
      .then((r) => {
        if (!cancelled) setRule(r);
      })
      .catch((err: unknown) => {
        if (!cancelled) setError(err instanceof Error ? err.message : String(err));
      });
    return () => {
      cancelled = true;
    };
  }, [matchKey, ruleId]);

  const persist = useCallback(async () => {
    if (!rule) return;
    setSaving(true);
    try {
      const ok = await saveRule(matchKey, rule);
      if (!ok) throw new Error("save rejected");
      onSaved?.();
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setSaving(false);
    }
  }, [matchKey, rule, onSaved]);

  if (error !== null && rule === null) return <p className="px-6 py-1 text-xs text-destructive">{error}</p>;
  if (!rule) return <p className="px-6 py-1 text-xs text-muted-foreground">Loading steps…</p>;

  const setSteps = (steps: StepDto[]): void => setRule({ ...rule, steps });

  return (
    <div className="border-t px-6 py-2">
      <div className="mb-1 flex items-center gap-2">
        <span className="text-[11px] font-semibold uppercase tracking-wide text-muted-foreground">Steps</span>
        <Select value="" onValueChange={(value) => setSteps([...rule.steps, newStep(value as (typeof STEP_TYPES)[number])])}>
          <SelectTrigger className="h-7 w-32 text-xs"><SelectValue placeholder="+ Add step…" /></SelectTrigger>
          <SelectContent>{STEP_TYPES.map((t) => <SelectItem key={t} value={t}>{t}</SelectItem>)}</SelectContent>
        </Select>
        <Button
          type="button"
          variant="outline"
          size="sm"
          disabled={saving}
          onClick={() => void persist()}
          className="ml-auto border-primary bg-primary/10 text-primary hover:bg-primary/20"
        >
          {saving ? "Saving…" : "Save steps"}
        </Button>
      </div>

      {rule.steps.length === 0 && (
        <p className="py-1 text-xs text-muted-foreground">No steps yet — this rule does nothing when it fires.</p>
      )}

      <ol className="space-y-1">
        {rule.steps.map((step, i) => (
          <li key={i} className="flex flex-wrap items-center gap-2 rounded border bg-muted/30 px-2 py-1 text-xs">
            <span className="w-4 text-right text-[10px] text-muted-foreground">{i + 1}.</span>
            <Select value={step.stepType} onValueChange={(value) => {
                const next = [...rule.steps];
                next[i] = { ...newStep(value as StepDto["stepType"]) };
                setSteps(next);
              }}>
              <SelectTrigger className={inputCls}><SelectValue /></SelectTrigger>
              <SelectContent>{STEP_TYPES.map((t) => <SelectItem key={t} value={t}>{t}</SelectItem>)}</SelectContent>
            </Select>

            {step.stepType === "Wait" && (
              <Label className="flex items-center gap-1">
                seconds
                <Input
                  type="number"
                  min={0}
                  value={step.waitSeconds ?? 0}
                  onChange={(e) =>
                    setSteps(
                      rule.steps.map((s, j) =>
                        j === i ? { ...s, waitSeconds: Math.max(0, Number(e.target.value) || 0) } : s,
                      ),
                    )
                  }
                  className={`w-16 ${inputCls}`}
                />
              </Label>
            )}

            {step.stepType === "Toggle" && (
              <>
                <Label className="flex items-center gap-1">
                  entity #
                  <Input
                    type="number"
                    min={0}
                    value={step.targetEntityId ?? 0}
                    onChange={(e) =>
                      setSteps(
                        rule.steps.map((s, j) =>
                          j === i ? { ...s, targetEntityId: Math.max(0, Number(e.target.value) || 0) } : s,
                        ),
                      )
                    }
                    className={`w-20 ${inputCls}`}
                  />
                </Label>
                <Label className="flex items-center gap-1">
                  or group
                  <Input
                    value={step.targetGroupName ?? ""}
                    placeholder="(none)"
                    onChange={(e) =>
                      setSteps(rule.steps.map((s, j) => (j === i ? { ...s, targetGroupName: e.target.value } : s)))
                    }
                    className={`w-28 ${inputCls}`}
                  />
                </Label>
                <Label className="flex items-center gap-1">
                  state
                  <Select value={step.toggleState === null || step.toggleState === undefined ? "invert" : String(step.toggleState)} onValueChange={(value) =>
                      setSteps(
                        rule.steps.map((s, j) => ({
                          ...(j === i ? s : s),
                          toggleState:
                            j !== i ? s.toggleState : value === "invert" ? null : value === "true",
                        })),
                      )
                    }>
                    <SelectTrigger className={inputCls}><SelectValue placeholder="invert" /></SelectTrigger>
                    <SelectContent><SelectItem value="invert">invert</SelectItem><SelectItem value="true">ON</SelectItem><SelectItem value="false">OFF</SelectItem></SelectContent>
                  </Select>
                </Label>
              </>
            )}

            {step.stepType === "StartTimer" && (
              <>
                <Label className="flex items-center gap-1">
                  target
                  <Select value={step.timerTarget ?? "Custom"} onValueChange={(value) =>
                      setSteps(
                        rule.steps.map((s, j) =>
                          j === i ? { ...s, timerTarget: value as StepDto["timerTarget"] } : s,
                        ),
                      )
                    }>
                    <SelectTrigger className={inputCls}><SelectValue /></SelectTrigger>
                    <SelectContent>{TIMER_TARGETS.map((t) => <SelectItem key={t} value={t}>{t}</SelectItem>)}</SelectContent>
                  </Select>
                </Label>
                {(step.timerTarget ?? "Custom") === "Custom" ? (
                  <>
                    <Label className="flex items-center gap-1">
                      minutes
                      <Input
                        type="number"
                        min={1}
                        value={step.timerMinutes ?? 15}
                        onChange={(e) =>
                          setSteps(
                            rule.steps.map((s, j) =>
                              j === i ? { ...s, timerMinutes: Math.max(1, Number(e.target.value) || 1) } : s,
                            ),
                          )
                        }
                        className={`w-16 ${inputCls}`}
                      />
                    </Label>
                    <Input
                      value={step.timerName ?? ""}
                      placeholder="name"
                      onChange={(e) =>
                        setSteps(rule.steps.map((s, j) => (j === i ? { ...s, timerName: e.target.value } : s)))
                      }
                      className={`w-24 ${inputCls}`}
                    />
                  </>
                ) : (
                  <Label className="flex items-center gap-1">
                    <Checkbox
                      checked={step.showCrateOnMap ?? true}
                      onCheckedChange={(checked) =>
                        setSteps(
                          rule.steps.map((s, j) => (j === i ? { ...s, showCrateOnMap: checked === true } : s)),
                        )
                      }
                    />
                    crate marker
                  </Label>
                )}
              </>
            )}

            {step.stepType === "CheckAvailability" && (
              <>
                <Label className="flex items-center gap-1">
                  op
                  <Select value={step.conditionOperator ?? "ALL_OFFLINE"} onValueChange={(value) =>
                      setSteps(
                        rule.steps.map((s, j) =>
                          j === i ? { ...s, conditionOperator: value as StepDto["conditionOperator"] } : s,
                        ),
                      )
                    }>
                    <SelectTrigger className={inputCls}><SelectValue /></SelectTrigger>
                    <SelectContent>{OPERATORS.map((o) => <SelectItem key={o} value={o}>{o}</SelectItem>)}</SelectContent>
                  </Select>
                </Label>
                <Input
                  value={step.conditionDeviceIdsCsv ?? ""}
                  placeholder="entity ids, comma-separated"
                  onChange={(e) =>
                    setSteps(
                      rule.steps.map((s, j) => (j === i ? { ...s, conditionDeviceIdsCsv: e.target.value } : s)),
                    )
                  }
                  className={`w-44 ${inputCls}`}
                />
              </>
            )}

            <span className="ml-auto flex gap-1">
              <Button
                type="button"
                variant="outline"
                size="icon"
                disabled={i === 0}
                onClick={() => {
                  const next = [...rule.steps];
                  [next[i - 1], next[i]] = [next[i]!, next[i - 1]!];
                  setSteps(next);
                }}
                className="h-7 w-7 disabled:opacity-30"
              >
                ↑
              </Button>
              <Button
                type="button"
                variant="outline"
                size="icon"
                disabled={i === rule.steps.length - 1}
                onClick={() => {
                  const next = [...rule.steps];
                  [next[i], next[i + 1]] = [next[i + 1]!, next[i]!];
                  setSteps(next);
                }}
                className="h-7 w-7 disabled:opacity-30"
              >
                ↓
              </Button>
              <Button
                type="button"
                variant="destructive"
                size="icon"
                onClick={() => setSteps(rule.steps.filter((_, j) => j !== i))}
                className="h-7 w-7"
              >
                ✕
              </Button>
            </span>

            {/* Conditional steps run inline when the CheckAvailability gate passes. */}
            {step.stepType === "CheckAvailability" && (
              <ConditionalStepsList
                steps={step.conditionalSteps ?? []}
                onChange={(conditionalSteps) =>
                  setSteps(rule.steps.map((s, j) => (j === i ? { ...s, conditionalSteps } : s)))
                }
              />
            )}
          </li>
        ))}
      </ol>

      {error !== null && <p className="mt-1 text-xs text-destructive">{error}</p>}
    </div>
  );
}

/** Nested conditional-step list for a CheckAvailability gate (same field rules). */
function ConditionalStepsList({
  steps,
  onChange,
}: {
  steps: StepDto[];
  onChange: (steps: StepDto[]) => void;
}): React.JSX.Element {
  return (
    <div className="w-full pl-6 pt-1">
      <div className="flex items-center gap-2">
        <span className="text-[10px] font-semibold uppercase tracking-wide text-muted-foreground">
          If available · run
        </span>
        <Select value="" onValueChange={(value) => onChange([...steps, newStep(value as StepDto["stepType"])])}>
          <SelectTrigger className="h-7 w-24 text-xs"><SelectValue placeholder="+ Add…" /></SelectTrigger>
          <SelectContent>{STEP_TYPES.filter((t) => t !== "CheckAvailability").map((t) => <SelectItem key={t} value={t}>{t}</SelectItem>)}</SelectContent>
        </Select>
      </div>
      <ul className="mt-1 space-y-1">
        {steps.map((step, i) => (
          <li key={i} className="flex flex-wrap items-center gap-2 rounded border bg-background px-2 py-1">
            <Select value={step.stepType} onValueChange={(value) => onChange(steps.map((s, j) => (j === i ? { ...newStep(value as StepDto["stepType"]) } : s)))}>
              <SelectTrigger className={`${inputCls} text-xs`}><SelectValue /></SelectTrigger>
              <SelectContent>{STEP_TYPES.filter((t) => t !== "CheckAvailability").map((t) => <SelectItem key={t} value={t}>{t}</SelectItem>)}</SelectContent>
            </Select>
            {step.stepType === "Wait" && (
              <Label className="flex items-center gap-1 text-xs">
                seconds
                <Input
                  type="number"
                  min={0}
                  value={step.waitSeconds ?? 0}
                  onChange={(e) =>
                    onChange(
                      steps.map((s, j) =>
                        j === i ? { ...s, waitSeconds: Math.max(0, Number(e.target.value) || 0) } : s,
                      ),
                    )
                  }
                  className={`w-14 ${inputCls}`}
                />
              </Label>
            )}
            {step.stepType === "Toggle" && (
              <>
                <Label className="flex items-center gap-1 text-xs">
                  entity #
                  <Input
                    type="number"
                    min={0}
                    value={step.targetEntityId ?? 0}
                    onChange={(e) =>
                      onChange(
                        steps.map((s, j) =>
                          j === i ? { ...s, targetEntityId: Math.max(0, Number(e.target.value) || 0) } : s,
                        ),
                      )
                    }
                  className={`w-16 ${inputCls}`}
                  />
                </Label>
                <Select value={step.toggleState === null || step.toggleState === undefined ? "invert" : String(step.toggleState)} onValueChange={(value) =>
                    onChange(
                      steps.map((s, j) => (j === i ? { ...s, toggleState: value === "invert" ? null : value === "true" } : s)),
                    )
                  }>
                  <SelectTrigger className={`${inputCls} text-xs`}><SelectValue /></SelectTrigger>
                  <SelectContent><SelectItem value="invert">invert</SelectItem><SelectItem value="true">ON</SelectItem><SelectItem value="false">OFF</SelectItem></SelectContent>
                </Select>
              </>
            )}
            {step.stepType === "StartTimer" && (
              <Label className="flex items-center gap-1 text-xs">
                minutes
                <Input
                  type="number"
                  min={1}
                  value={step.timerMinutes ?? 15}
                  onChange={(e) =>
                    onChange(
                      steps.map((s, j) =>
                        j === i ? { ...s, timerMinutes: Math.max(1, Number(e.target.value) || 1) } : s,
                      ),
                    )
                  }
                  className={`w-14 ${inputCls}`}
                />
              </Label>
            )}
            <Button
              type="button"
              variant="destructive"
              size="icon"
              onClick={() => onChange(steps.filter((_, j) => j !== i))}
              className="ml-auto h-7 w-7"
            >
              ✕
            </Button>
          </li>
        ))}
      </ul>
    </div>
  );
}
