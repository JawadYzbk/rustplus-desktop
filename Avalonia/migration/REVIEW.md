# Account, pairing, startup, device organization and recycler review — 2026-10-16

Development preview verified; the full migration remains unfinished. Overall self-evaluation: **3.8/5**.

| Axis | Score | Evidence and next improvement |
| --- | --- | --- |
| Accuracy | 4 | 358 cases pass in the worktree and source-only isolation. Debug/Release and three self-contained publishes succeed. Sync/import, native protocol projection, live team lifetime/stale data, minimum-size 24-member docking and direct nested icon selection/save are covered. Live Steam/FCM/cloud, concurrent writers and native Linux/macOS still need acceptance. |
| Completeness | 3 | Pairing, local organization/icons/direct picker access, recycler, responsive panels, live roster, cloud account/plan, own recovery/automatic sync and guarded import are implemented. Full map/team/master/chat publication and most shipped features remain pending. 22 images, double-click shortcuts and localization remain required. These slices do not establish full parity. |
| Clarity | 4 | Core exposes safe contracts; credentials and mutations remain in Infrastructure. HTTP/private writers and recovery are reused. Cloud permission is explicitly separate from default-off local automatic synchronization. |
| Actionability | 4 | README includes build/run and account workflows; tests use disposable fixtures. Live validation requires the user's Steam consent and a paired Rust server. |
| Conciseness | 4 | Account and profile writes share one private atomic writer; the existing transport is reused for direct re-pair probes. Reconnect, persistence and UI coordination need separate tests and cannot safely be reduced to a single event handler. |

Highest-impact improvements: validate real sign-in/pairing/sign-out; finish the documented shipped-feature
backlog; run native Linux/macOS runtime and appearance acceptance. The completeness gap requires further
implementation, not a change to the release label. No production-ready claim is justified.

Self-check: the user can test the new preview workflows and inspect the explicit backlog, but would not
reasonably consider the entire migration complete. Deliver this slice for review while continuing required ports.
