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

const inputCls = "rounded border bg-transparent px-2 py-1";

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
        <select
          value=""
          onChange={(e) => {
            const t = e.target.value as (typeof STEP_TYPES)[number];
            if (e.target.value !== "") setSteps([...rule.steps, newStep(t)]);
          }}
          className="rounded border bg-transparent px-1 py-0.5 text-xs"
        >
          <option value="">+ Add step…</option>
          {STEP_TYPES.map((t) => (
            <option key={t} value={t}>
              {t}
            </option>
          ))}
        </select>
        <button
          type="button"
          disabled={saving}
          onClick={() => void persist()}
          className="ml-auto rounded-md border border-primary bg-primary/10 px-2 py-0.5 text-primary hover:bg-primary/20 disabled:opacity-50"
        >
          {saving ? "Saving…" : "Save steps"}
        </button>
      </div>

      {rule.steps.length === 0 && (
        <p className="py-1 text-xs text-muted-foreground">No steps yet — this rule does nothing when it fires.</p>
      )}

      <ol className="space-y-1">
        {rule.steps.map((step, i) => (
          <li key={i} className="flex flex-wrap items-center gap-2 rounded border bg-muted/30 px-2 py-1 text-xs">
            <span className="w-4 text-right text-[10px] text-muted-foreground">{i + 1}.</span>
            <select
              value={step.stepType}
              onChange={(e) => {
                const next = [...rule.steps];
                next[i] = { ...newStep(e.target.value as StepDto["stepType"]) };
                setSteps(next);
              }}
              className={inputCls}
            >
              {STEP_TYPES.map((t) => (
                <option key={t}>{t}</option>
              ))}
            </select>

            {step.stepType === "Wait" && (
              <label className="flex items-center gap-1">
                seconds
                <input
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
              </label>
            )}

            {step.stepType === "Toggle" && (
              <>
                <label className="flex items-center gap-1">
                  entity #
                  <input
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
                </label>
                <label className="flex items-center gap-1">
                  or group
                  <input
                    value={step.targetGroupName ?? ""}
                    placeholder="(none)"
                    onChange={(e) =>
                      setSteps(rule.steps.map((s, j) => (j === i ? { ...s, targetGroupName: e.target.value } : s)))
                    }
                    className={`w-28 ${inputCls}`}
                  />
                </label>
                <label className="flex items-center gap-1">
                  state
                  <select
                    value={step.toggleState === null || step.toggleState === undefined ? "" : String(step.toggleState)}
                    onChange={(e) =>
                      setSteps(
                        rule.steps.map((s, j) => ({
                          ...(j === i ? s : s),
                          toggleState:
                            j !== i ? s.toggleState : e.target.value === "" ? null : e.target.value === "true",
                        })),
                      )
                    }
                    className={inputCls}
                  >
                    <option value="">invert</option>
                    <option value="true">ON</option>
                    <option value="false">OFF</option>
                  </select>
                </label>
              </>
            )}

            {step.stepType === "StartTimer" && (
              <>
                <label className="flex items-center gap-1">
                  target
                  <select
                    value={step.timerTarget ?? "Custom"}
                    onChange={(e) =>
                      setSteps(
                        rule.steps.map((s, j) =>
                          j === i ? { ...s, timerTarget: e.target.value as StepDto["timerTarget"] } : s,
                        ),
                      )
                    }
                    className={inputCls}
                  >
                    {TIMER_TARGETS.map((t) => (
                      <option key={t}>{t}</option>
                    ))}
                  </select>
                </label>
                {(step.timerTarget ?? "Custom") === "Custom" ? (
                  <>
                    <label className="flex items-center gap-1">
                      minutes
                      <input
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
                    </label>
                    <input
                      value={step.timerName ?? ""}
                      placeholder="name"
                      onChange={(e) =>
                        setSteps(rule.steps.map((s, j) => (j === i ? { ...s, timerName: e.target.value } : s)))
                      }
                      className={`w-24 ${inputCls}`}
                    />
                  </>
                ) : (
                  <label className="flex items-center gap-1">
                    <input
                      type="checkbox"
                      checked={step.showCrateOnMap ?? true}
                      onChange={(e) =>
                        setSteps(
                          rule.steps.map((s, j) => (j === i ? { ...s, showCrateOnMap: e.target.checked } : s)),
                        )
                      }
                    />
                    crate marker
                  </label>
                )}
              </>
            )}

            {step.stepType === "CheckAvailability" && (
              <>
                <label className="flex items-center gap-1">
                  op
                  <select
                    value={step.conditionOperator ?? "ALL_OFFLINE"}
                    onChange={(e) =>
                      setSteps(
                        rule.steps.map((s, j) =>
                          j === i ? { ...s, conditionOperator: e.target.value as StepDto["conditionOperator"] } : s,
                        ),
                      )
                    }
                    className={inputCls}
                  >
                    {OPERATORS.map((o) => (
                      <option key={o}>{o}</option>
                    ))}
                  </select>
                </label>
                <input
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
              <button
                type="button"
                disabled={i === 0}
                onClick={() => {
                  const next = [...rule.steps];
                  [next[i - 1], next[i]] = [next[i]!, next[i - 1]!];
                  setSteps(next);
                }}
                className="rounded border px-1 hover:bg-muted disabled:opacity-30"
              >
                ↑
              </button>
              <button
                type="button"
                disabled={i === rule.steps.length - 1}
                onClick={() => {
                  const next = [...rule.steps];
                  [next[i], next[i + 1]] = [next[i + 1]!, next[i]!];
                  setSteps(next);
                }}
                className="rounded border px-1 hover:bg-muted disabled:opacity-30"
              >
                ↓
              </button>
              <button
                type="button"
                onClick={() => setSteps(rule.steps.filter((_, j) => j !== i))}
                className="rounded border px-1 text-destructive hover:bg-destructive/10"
              >
                ✕
              </button>
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
        <select
          value=""
          onChange={(e) => {
            if (e.target.value !== "") onChange([...steps, newStep(e.target.value as StepDto["stepType"])]);
          }}
          className="rounded border bg-transparent px-1 py-0.5 text-xs"
        >
          <option value="">+ Add…</option>
          {STEP_TYPES.filter((t) => t !== "CheckAvailability").map((t) => (
            <option key={t} value={t}>
              {t}
            </option>
          ))}
        </select>
      </div>
      <ul className="mt-1 space-y-1">
        {steps.map((step, i) => (
          <li key={i} className="flex flex-wrap items-center gap-2 rounded border bg-background px-2 py-1">
            <select
              value={step.stepType}
              onChange={(e) => onChange(steps.map((s, j) => (j === i ? { ...newStep(e.target.value as StepDto["stepType"]) } : s)))}
              className={`${inputCls} text-xs`}
            >
              {STEP_TYPES.filter((t) => t !== "CheckAvailability").map((t) => (
                <option key={t}>{t}</option>
              ))}
            </select>
            {step.stepType === "Wait" && (
              <label className="flex items-center gap-1 text-xs">
                seconds
                <input
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
              </label>
            )}
            {step.stepType === "Toggle" && (
              <>
                <label className="flex items-center gap-1 text-xs">
                  entity #
                  <input
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
                </label>
                <select
                  value={step.toggleState === null || step.toggleState === undefined ? "" : String(step.toggleState)}
                  onChange={(e) =>
                    onChange(
                      steps.map((s, j) => (j === i ? { ...s, toggleState: e.target.value === "" ? null : e.target.value === "true" } : s)),
                    )
                  }
                  className={`${inputCls} text-xs`}
                >
                  <option value="">invert</option>
                  <option value="true">ON</option>
                  <option value="false">OFF</option>
                </select>
              </>
            )}
            {step.stepType === "StartTimer" && (
              <label className="flex items-center gap-1 text-xs">
                minutes
                <input
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
              </label>
            )}
            <button
              type="button"
              onClick={() => onChange(steps.filter((_, j) => j !== i))}
              className="ml-auto rounded border px-1 text-destructive hover:bg-destructive/10"
            >
              ✕
            </button>
          </li>
        ))}
      </ul>
    </div>
  );
}
