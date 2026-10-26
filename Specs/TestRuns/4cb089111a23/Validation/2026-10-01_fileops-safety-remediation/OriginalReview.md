# I26 File Operations completion review

Reviewed 2026-10-01, against RedSalamander `e860022a0` on `codex/fileops-ui-i26`.

**Verdict: do not accept “finished and trustworthy” for critical file operations.** The qualification packet is real, but it does not cover important decision-lifetime and result-truth failures. Three P1 findings need correction before I would approve this implementation. The new path-detail UI also introduces a focus-restoration defect. Passing thousands of tests does not resolve these concrete counterexamples.

This review separates new I26 behavior from defects that were already present in nearby code. Most engine defects below predate the implementation baseline `1425d56799bbc5fce158e92be93da92ef9982292`; that matters for attribution, but they still invalidate broad safety assurances about the finished flow. The historic Done document should remain an evidence record. Corrections need an indexed follow-up and qualification against the actual fixed implementation.

## Findings

### 1. P1 — A decision for prompt A can be applied to prompt B

The UI action identity is only kind, task ID, and action. [PopupHitTest](Z:/src/RedSalamander/RedSalamander/FolderWindow.FileOperations.Popup.h:28) and the conflict presentation snapshot do not carry the engine prompt generation. [Primary submission](Z:/src/RedSalamander/RedSalamander/FolderWindow.FileOperations.Popup.cpp:11618) rereads the current prompt, verifies that the raw action is currently offered, unlocks, and calls a submitter without an expected generation. [Overflow submission](Z:/src/RedSalamander/RedSalamander/FolderWindow.FileOperations.Popup.cpp:11573) and [Escape](Z:/src/RedSalamander/RedSalamander/FolderWindow.FileOperations.Popup.cpp:6548) have the same weakness. The async More callback explicitly retains an old action list and task ID, then invokes the unfenced submission path [later](Z:/src/RedSalamander/RedSalamander/FolderWindow.FileOperations.Popup.cpp:11739).

Concrete trigger: accept Replace for one conflicting file; the worker reaches the next conflict for the same task before the UI processes a repeated activation or redraws. The old Replace hit/callback is accepted for the next file if Replace is also offered there. The current destination guard may correctly identify file B, but it cannot prove that the user consented to replacing B. This is a consent defect, not an identity-validation substitute.

The engine [SubmitConflictDecision](Z:/src/RedSalamander/RedSalamander/FolderWindow.FileOperations.State.cpp:10732) also accepts whatever prompt is current. Hosted-control matching and focus restoration use the same generation-free hit identity [here](Z:/src/RedSalamander/RedSalamander/FolderWindow.FileOperations.Popup.cpp:7085). In particular, safe focus for successive Recycle escalation prompts is tracked by task ID rather than prompt generation [here](Z:/src/RedSalamander/RedSalamander/FolderWindow.FileOperations.Popup.cpp:7186).

The contract explicitly requires stale actions to fail closed: [UI_FileOperationsPopup.md](Z:/src/RedSalamander/Specs/UI/UI_FileOperationsPopup.md:1226). Existing “stale painted button” coverage tests inspection/navigation [only](Z:/src/RedSalamander/RedSalamander/SelfTest/Commands/Commands.SelfTest.FileOps.cpp:14359). Current-prompt Enter/Escape tests do not replay a destructive action from A against B.

**Fix:** carry `{taskId, promptGeneration, action, placement}` through snapshots, hits, hosted controls, menus, Escape, Apply-to-all, and focus bindings. Validate the expected generation and publication policy atomically with submission under the arbiter lock. Reject stale input without resolving the current prompt. Reset safe initial focus for each new prompt generation.

**Required tests:** publish A; retain each action route; advance to B on the same task with the same offered action; invoke A's input; assert B is still unresolved and its destination bytes are unchanged. Cover primary, overflow, Escape, UIA, held/repeated activation, Apply-to-all, and consecutive Recycle escalation prompts.

**Evidence:** [compiled characterization results](C:/RedSalamander.Perf/evidence/i26-code-review-20261001-e860022a0/conflict-probe-results.txt). The extracted production submit/wait methods accept an action retained from A on B. This is a source-extracted engine reproduction; it is not a physical-input application reproduction. The production UI routes above establish the missing correlation.

**Attribution:** the generation gap predates I26; the revised UI and completion map do not repair or test it.

### 2. P1 — Duplicate submission can wake and cancel the next prompt

[SubmitConflictDecision](Z:/src/RedSalamander/RedSalamander/FolderWindow.FileOperations.State.cpp:10748) overwrites any pending decision and signals the shared manual-reset event after releasing the mutex. [WaitForConflictDecision](Z:/src/RedSalamander/RedSalamander/FolderWindow.FileOperations.State.cpp:5296) trusts a signaled event, reads `decisionAction.value_or(Cancel)`, then clears the prompt. [ClearConflictPrompt](Z:/src/RedSalamander/RedSalamander/FolderWindow.FileOperations.State.cpp:4895) resets state and the event. There is no one-shot acceptance state or generation attached to the signal.

A valid schedule, even with UI submissions serialized:

1. UI submit A1 finishes; the event remains signaled.
2. UI submit A2 starts while A is still active, replaces the action, releases the mutex, and is preempted before `SetEvent`.
3. The worker consumes A1's signal with A2's stored action, closes A and resets the event.
4. The worker encounters the next conflict and publishes B, resetting the event and clearing the pending choice.
5. UI submit A2 resumes its old `SetEvent` tail.
6. B's waiter wakes with no B decision and returns Cancel.

The same code also permits a second submission to replace the first pending choice before the worker consumes it. Moving just one event call under a lock is insufficient unless decision consumption and prompt transition are made coherent too.

**Fix:** make prompt acceptance a generation-checked, one-shot state transition. Tie wakeups to accepted state and consume/close that state atomically. A signaled event with no accepted decision must never resolve a prompt as if Cancel had been chosen.

**Required tests:** deterministic gates after storing the second action and before its event signal; consume A, publish B, release the old signal, and verify B still waits. Test duplicate identical/different actions, Apply-to-all, cancellation, task teardown, and concurrent worker ownership. No timing sleeps should be needed.

**Evidence:** [exact-body probe](C:/RedSalamander.Perf/evidence/i26-code-review-20261001-e860022a0/conflict-probe.cpp), [provenance](C:/RedSalamander.Perf/evidence/i26-code-review-20261001-e860022a0/conflict-probe-provenance.json), and [results](C:/RedSalamander.Perf/evidence/i26-code-review-20261001-e860022a0/conflict-probe-results.txt). A second agent adversarially checked the extraction and schedule against production and confirmed it. The reproduced cross-prompt effect is an unsolicited Cancel, not proof of a destructive action from this event race.

**Attribution:** pre-existing engine race, still present in the qualified candidate.

### 3. P1 — Missing cancellation receipts become false “untouched source” truth

[FinalizeTypedItemResults](Z:/src/RedSalamander/RedSalamander/FolderWindow.FileOperations.State.cpp:11060) reports `NotAttempted / Retained / Canceled` when a receipt is absent and cancellation is present, even if the builder has reached `MutationPossible`. [FileSystemItemCompleted](Z:/src/RedSalamander/RedSalamander/FolderWindow.FileOperations.State.cpp:7535) deliberately clears an earlier receipt on a later null completion, so this is a reachable reducer state, not just malformed data.

The ABI [permits null mutation receipts](Z:/src/RedSalamander/Common/PlugInterfaces/FileSystem.h:136) when a provider cannot expose truth and requires destructive recovery to fail closed. Cancellation proves a terminal cause; it does not prove that storage was untouched. The [file-operation contract](Z:/src/RedSalamander/Specs/FileSystem/FileSystem_FileOperations.md:1615) also explicitly preserves unknown publication/source truth rather than inferring it from cancellation.

The real Recycle command admits bulk execution [here](Z:/src/RedSalamander/RedSalamander/FolderWindow.FileOperations.cpp:4082), which calls [DeleteItems](Z:/src/RedSalamander/RedSalamander/FolderWindow.FileOperations.State.cpp:21014) and then the generic finalizer. A compliant provider can have changed storage before returning cancellation without a receipt. The current result then incorrectly says the source was retained. That truth feeds presentation and completion consumers.

**Limit of the finding:** the shipped Local provider passes receipt pointers in its terminal Delete callbacks; I did not demonstrate this null-receipt trigger in its current implementation. The confirmed host defect concerns other providers allowed by the ABI, or a mutation-possible item for which no terminal truth arrives. Do not describe it as an observed Local-provider data-loss incident.

**Fix:** distinguish proven not-started items from mutation-possible items. Preserve Canceled as the terminal cause where required, but use Unknown publication/source axes after possible destructive mutation unless a valid receipt proves otherwise. Preparing items and proved no-commit receipts must retain their existing safe classification.

**Required tests:** bulk Delete and Move providers that mutate, then cancel with a null receipt; cancellation without a terminal item callback; known no-commit, known commit, explicit unknown, malformed and missing receipts; preceding child/attempt receipts overwritten by null; mixed multi-item outcomes. Assert dense item results, aggregate state, refresh/removal decisions, and recovery eligibility. Existing null-receipt tests [exercise success](Z:/src/RedSalamander/RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.SelfTest.Phases05_06.cpp:1671), not this canceled fallback.

**Evidence:** [extracted reducer results](C:/RedSalamander.Perf/evidence/i26-code-review-20261001-e860022a0/cancellation-reducer-probe-results.txt), [source and fixtures](C:/RedSalamander.Perf/evidence/i26-code-review-20261001-e860022a0/cancellation-reducer-probe.cpp). Both mutation-possible Delete and Move falsely report Retained; the Preparing control case reports Retained correctly. This probe exercises the exact reducer branch, not a full provider integration.

**Attribution:** pre-existing reducer defect; I26 coverage and assurance gap.

### 4. P2 — New path-detail rebuilds can reclaim native focus

The new path detail uses a read-only `TextField`. [SyncPathDetail](Z:/src/RedSalamander/RedSalamander/FolderWindow.FileOperations.Popup.cpp:6539) focuses it with `SetFocusControl(field, false)` whenever the field is created. [Hosted-tree rebuild](Z:/src/RedSalamander/RedSalamander/FolderWindow.FileOperations.Popup.cpp:6747) destroys and recreates that field while retaining the detail ID; [restoration](Z:/src/RedSalamander/RedSalamander/FolderWindow.FileOperations.Popup.cpp:7259) makes another false-flag focus call.

At the actual pinned DxUi revision `19ca44d95a0cd493fc7fb1a57f4d4ad227bdefc1`, the flag suppresses native focus transfer for non-text controls. Text fields still enter [ActivateTextInput](C:/RsI26Adopt/.build/dependencies/DxUi/source/19ca44d95a0cd493fc7fb1a57f4d4ad227bdefc1/src/Controls/DxUi.WindowHost.cpp:1659), which calls [SetFocus](C:/RsI26Adopt/.build/dependencies/DxUi/source/19ca44d95a0cd493fc7fb1a57f4d4ad227bdefc1/src/Controls/DxUi.NativeTextInput.cpp:591). Read-only fields support and export text-input state too. DxUi is behaving according to its documented non-text scope; the consumer incorrectly treats this flag as a general no-focus operation.

Trigger: open path detail, focus the main window or another application, then let another task/descriptor change rebuild hosted controls while the detail stays open. Focus loss deactivates text input but retains logical focus; it does not close detail. Recreation calls native text-input activation again without an outside-focus guard. The popup can reclaim native focus during a background render.

**Fix:** capture native-focus ownership before rebuilding and restore it only when the popup actually owned it. Background reconstruction must preserve logical detail state without activating text input. If a shared logical-focus API is added, qualify it in DxUi and adopt the pin explicitly; consumer changes must not silently assume a wider contract.

**Required test:** detail open with text/selection retained; focus outside; change another task to force a tree rebuild; pump presentation; assert native focus and foreground remain outside. Repeat with text, Copy, and Close logically selected, resize/theme/DPI changes, and detail closure. Current [path tests](Z:/src/RedSalamander/RedSalamander/SelfTest/Commands/Commands.SelfTest.FileOps.cpp:9399) check unchanged foreground while the popup already owns focus.

**Evidence:** [compiled pinned-method probe](C:/RedSalamander.Perf/evidence/i26-code-review-20261001-e860022a0/focus-probe-results.txt), [provenance](C:/RedSalamander.Perf/evidence/i26-code-review-20261001-e860022a0/focus-probe-provenance.json). The actual pinned methods call intercepted Win32 SetFocus for a read-only field with `false`, including the same-control restoration path. No real window was created or activated by this probe. Consumer reachability was independently verified by source review.

**Attribution:** introduced by the new consumer path-detail/restoration code.

### 5. P2 — Hidden Issues panes keep rebuilding and sorting history

[OnCreate](Z:/src/RedSalamander/RedSalamander/FolderWindow.FileOperations.IssuesPane.cpp:1572) starts a 750 ms timer. Close hides and retains the window; [OnShowWindow](Z:/src/RedSalamander/RedSalamander/FolderWindow.FileOperations.IssuesPane.cpp:1621) saves state when hidden but does not stop the timer. [OnTimer](Z:/src/RedSalamander/RedSalamander/FolderWindow.FileOperations.IssuesPane.cpp:1599) calls RefreshRows unconditionally. [RefreshRows](Z:/src/RedSalamander/RedSalamander/FolderWindow.FileOperations.IssuesPane.cpp:1007) rebuilds and sorts before checking equality. [BuildRows](Z:/src/RedSalamander/RedSalamander/FolderWindow.FileOperations.IssuesPane.cpp:646) copies live/completed diagnostics, builds maps and strings, correlates observations, and projects every retained row.

Leaving Issues closed therefore continues allocations and O(n log n) sorting every 750 ms. History is bounded; this is recurring unnecessary hidden work, not an unbounded memory-leak claim. I did not measure its wall-clock cost in this review.

The new hidden integration interval is [350 ms](Z:/src/RedSalamander/RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.SelfTest.IntegratedPresentationPerf.cpp:281), shorter than the timer period, and its assertions cover [render/cache counters](Z:/src/RedSalamander/RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.SelfTest.IntegratedPresentationPerf.cpp:567), not projection work or wakeups. A timer tick could still occur depending on its phase, but the test neither guarantees one nor detects the row work.

**Fix:** stop the timer while hidden/minimized and refresh once on show. Use a cheap model revision to avoid reprojecting unchanged history during visible polling. A notification-based projection can remove polling altogether.

**Required tests:** hide/minimize for more than two timer periods; assert projection count, allocations, and UI-thread wakeups remain idle; append diagnostics while hidden; show and verify fresh rows and selection. Test large retained history and repeated open/close.

**Attribution:** the timer policy predates I26; the expanded projection increases the work and the new qualification misses it.

### 6. P2 — Reused alerts have no presentation identity for UIA commands

[AlertOverlayUiaProvider::Invoke](Z:/src/RedSalamander/RedSalamander/Ui/AlertOverlayWindow.cpp:522) resolves the current visible window from a saved HWND and kind/index. It has no alert-presentation generation. It posts [raw commands](Z:/src/RedSalamander/RedSalamander/Ui/AlertOverlayWindow.cpp:559), and the [WndProc](Z:/src/RedSalamander/RedSalamander/Ui/AlertOverlayWindow.cpp:911) dispatches them against the current model/callbacks.

The application-scoped alert [reuses one overlay](Z:/src/RedSalamander/RedSalamander/HostServices.cpp:2011) and [Hide retains it](Z:/src/RedSalamander/RedSalamander/HostServices.cpp:2045). A retained close provider from alert A can invoke after A was hidden and B was shown, closing B. A queued A command can also be delivered after presentation replacement. This route requires ordinary overlay reuse, not speculative numeric HWND recycling. The guaranteed application-alert example concerns dismissal; this review did not demonstrate a stale destructive prompt acceptance through this provider.

**Fix/test:** give presentations a generation and invalidate old child providers and queued commands on Hide/model replacement. Retain a provider and delay a posted command across A→B, then assert neither invokes B. Also test destroy/recreate and changed button/option identities. Existing live [dialog UIA tests](Z:/src/RedSalamander/RedSalamander/SelfTest/Commands/Commands.SelfTest.Dialogs.cpp:863) do exercise worker-thread enumeration and invocation; they do not test replacement lifetimes.

**Attribution:** existing robustness defect, not an I26 regression. Source-confirmed routing; no live AT replacement reproduction was run.

### 7. P2 — Partial fallback-window creation leaves an unrecoverable empty surface

[EnterFailureSurface](Z:/src/RedSalamander/RedSalamander/FolderWindow.FileOperations.Popup.cpp:11796) sets `_failureSurface=true` before creating text, Cancel all and Close children. If any creation fails, it [returns false](Z:/src/RedSalamander/RedSalamander/FolderWindow.FileOperations.Popup.cpp:11867) without resetting that flag or destroying partial children. Those children were created at zero size. [Layout](Z:/src/RedSalamander/RedSalamander/FolderWindow.FileOperations.Popup.cpp:11900) requires all three, so none are placed. Later EnterFailureSurface calls immediately return true because the flag is already set.

Under a transient Win32 resource failure, the supposedly graphics-independent fallback can remain blank/incomplete until destroyed and reopened. The code attempts a separate host alert; that does not repair the failed surface. I did not inject a native CreateWindowEx failure in this review.

**Fix/test:** build the child set transactionally with owned handles, publish the success state after all required children exist, and roll back partial creation. Distinguish “fallback mode selected” from “fallback UI successfully created” if recovery is separate. Inject failure at each child creation and verify rollback, retry, teardown, keyboard access, and UIA. Current fallback tests inject graphics/host-attach failures [here](Z:/src/RedSalamander/RedSalamander/SelfTest/Commands/Commands.SelfTest.FileOps.cpp:15958) and assume native child creation succeeds.

**Attribution:** pre-existing failure-path flaw. Confirmed statically; rare fault-injection scenario.

## Coverage assessment

The suite has substantial useful coverage. I would retain it, not replace it with new source-shape assertions.

| Area | What existing evidence establishes | Missing adversarial scenario |
|---|---|---|
| Conflict actions/defaults | Current-prompt action policy, two different pending task IDs, Enter/Escape, Recycle safe default | Same task, next prompt generation; duplicate submits; delayed signal tail; destructive stale UIA/menu input |
| Inspection | Exact provider/path navigation, stale inspection context, focus return, missing targets, teardown | Inspection coverage is not consent-generation coverage |
| Queue destination retargeting | Immutable selected admission views, reprepare/publication checks, durable destination replacement | Combine retargeting with rapid cancel/close and provider callbacks that reenter the owner; keep existing deterministic cases |
| Typed results | Exact rename carrier, malformed/contradictory receipt handling, unknown success, guarded mutation truth | Cancellation with absent receipt or omitted item completion after mutation became possible |
| Path detail | Unicode/long text, read-only selection/copy, keyboard traversal, same HWND, UIA text | Focus loss followed by background tree reconstruction |
| Issues | Stable IDs, typed row correlation, selection, detail, eviction, persistence, visible scale | Hidden/minimized projection and wakeups for a duration exceeding the timer period |
| Alert scrolling | Pure layouts at 320/480/920 DIP widths and 80 DIP height, label fit, focus reveal, footer reachability, grow/clamp; live Home/End | Live wheel/fractional deltas, thumb drag, PageUp/PageDown/Tab, resize while focused/scrolled, clipped UIA bounds/point lookup and offscreen Invoke |
| Alert accessibility | Live provider enumeration and visible button/option invocation from a worker thread | Provider/queued-command reuse across a replacement alert |
| Failure surface | Graphics/host attachment faults, expected three children, keyboard/UIA, reopen/churn | Individual native child allocation failures and recovery |
| Search/index changes | Regex case/rejection round trips, readiness/live fallback, partial SQLite failure/dedup/progress | Cancel after partial indexed results; callback failure during final fallback progress |

The 16-row / 62-witness / 39-unique-native-ID acceptance map is useful traceability, not proof of every behavioral conjunction. Its “stale/default/Enter/Escape” facet conflates distinct cases. The hidden resource fixture observes rendering; it does not establish absence of hidden row work. Add state-correlated behavioral assertions at these boundaries rather than adding more successful layout screenshots.

The plan honestly excludes native 192 DPI, external screen-reader execution, physical injected input, consumer ARM64/ASan and release handoff, and separately owned I18/H4 work. Those are limits, not hidden passes. This review does not claim to have executed those gates. The latest broad Debug FileOps evidence is carried V55/V56 evidence rather than a fresh final full Debug product run; final focused Debug checks establish their narrower scope.

## Architecture and simplification

1. **Extract one conflict broker.** Popup primary, overflow, Escape, default/cancel routing and engine submission independently reread pieces of current policy. A single immutable generation-bearing binding and one atomic accept/consume transition make the consent rules reviewable. Preserve existing provider guards and repeat-scope semantics.
2. **Extract a pure typed-result reducer.** Preparing, possible mutation, valid known receipt, explicit unknown, absent/malformed truth and host-proved partial outcomes should be reduced through one table with meaningful tests. The generic finalizer and qualified per-item paths currently encode overlapping rules; the canceled-missing-receipt inconsistency is the practical cost.
3. **Separate Issues projection from its HWND.** Keep row identity/correlation/sort/view state in a pure projection, update from model revisions, and let the host own only visibility/input/persistence/render scheduling. This simplifies both idle behavior and tests.
4. **Make logical and native focus distinct operations.** New consumer code uses a non-text focus flag on a text control. A deliberate logical restoration policy is clearer than scattered boolean calls whose behavior changes by control type.
5. **Keep alert accessibility as one provider, but centralize presentation identity.** The existing alert painter/model and hand-written UIA mirror need the same lifetime fencing. Do not add a second provider or rebuild this into an unrelated framework merely to hide the issue.
6. **Split the large translation units by ownership.** State.cpp is about 21,000 lines and Popup.cpp about 14,000; the command self-test file is about 20,000. Conflict state, result reduction, immutable presentation projection and host lifetime are useful boundaries. Moving code alone is not a fix; each extraction should remove duplicate rules and expose invariants to tests.

The new immutable selected-admission model, durable retarget commit discipline, exact provider-relative path handling, bounded diagnostic retention, and length-prefixed legacy issue keys are defensible complexity. I found no substantiated reason to simplify those protections away.

## Evidence and review limits

I reviewed the Done claim, current owning contracts, the production diff from `1425d567`, critical callers and existing engine paths, new/adapted tests, the current pin and relevant adopted DxUi input/accessibility/layout behavior. Three cheaper adversarial agents covered engine/result/lifetime, UI/accessibility/layout, and coverage/smaller production changes. I checked primary findings myself and compiled the three characterization probes. The review includes the search/index repairs carried in the final candidate.

No additional confirmed defect emerged in the reviewed destination-retarget transaction, inspection freshness/focus ticket, footer/completed layout helpers, settings/localization changes, deferred plugin alias bookkeeping, or search/index repair paths. This is a review result, not a guarantee that every runtime interleaving is correct.

Rejected or narrowed suspicions:

- A search progress count looked stale after coalescing; tracing the actual consumer showed authoritative QueryStats applied before its forced completed progress callback. No product bug was established.
- Clipboard-ready falls back by task ID for an unresolved nonzero payload token, unlike the neighboring batch-rename receiver. This violates the shared helper's stated receiver contract, but I could not establish a normal same-live-window route or prove the hypothesized Windows message/HWND-reuse delivery. It is a follow-up test/consistency concern, not a confirmed release-blocking finding here.
- Initial claims of no Alert UIA tests were rejected: dialog tests do invoke the live provider from a worker. Their gap is presentation replacement, not basic Invoke coverage.
- Event cancellation ordering does not show a lost cancellation wakeup; the confirmed crossover requires a duplicate submission's delayed signal tail as reproduced above.
- The prior focus-stealing investigation was only a lead. The path-detail finding uses the current consumer and exact current pinned source, not that older hypothesis.

Source identity: the retained V65 Release executable at `C:\RsI26Adopt\.build\x64\Release\RedSalamander.exe` matches SHA-256 `1287c08658659dc62d9cd4c63eb7ee187a3ffa1485fd76fee586a08c895e2872` and build receipt `7442adedcaf17d954a58b726f49756312beabb2c2426589e4fd71af42584fdac`. The frozen source manifest was checked against its retained checkout; targeted production/test hashes also match the reviewed Z: checkout. The old executable under the primary checkout was not used as current validation.

The retained packet's Release Full/Fresh result is **2211 passed / 0 failed / 46 skipped**; it was inspected, not rerun by this review. The gallery and paired performance evidence are genuine retained evidence; they do not cover the failures above.

Current review executions used the verified V65 binary with `--selftest-no-activate`:

- `FileOps_BatchRenameExactResultCarrier`: **3 passed / 0 failed / 0 skipped**, including Setup and Cleanup; one selected behavioral case. [Runner result](C:/RedSalamander.Perf/evidence/i26-code-review-20261001-e860022a0/20261001T195140Z-96200-6dd1c8a76d6e453fb55a818b431e720f-runner-results.json).
- `FileOps_QueuedAdmissionContracts`: **3 passed / 0 failed / 0 skipped**, including Setup and Cleanup; one selected behavioral case. [Runner result](C:/RedSalamander.Perf/evidence/i26-code-review-20261001-e860022a0/queued-admission-runner-results.json), [native result](C:/RedSalamander.Perf/evidence/i26-code-review-20261001-e860022a0/queued-admission-native-results.json). The case was rerun to preserve its receipt after normal runner retention removed the earlier run.
- Conflict characterization: stale action acceptance, duplicate replacement and cross-prompt signal cancellation reproduced.
- Cancellation reducer characterization: false Retained truth after possible mutation reproduced; Preparing control behaved correctly.
- Pinned focus characterization: `SetFocusControl(false)` on read-only text reaches SetFocus; non-text control preserves native focus.

All three probes compile as x64 with `/W4 /WX /permissive- /std:c++latest`; their scripts discover Visual Studio. Source-extracted probes isolate the faulty state transitions but do not replace integrated application/provider regression tests. [Conflict script](C:/RedSalamander.Perf/evidence/i26-code-review-20261001-e860022a0/build-conflict-probe.ps1), [cancellation script](C:/RedSalamander.Perf/evidence/i26-code-review-20261001-e860022a0/build-cancellation-probe.ps1), [focus script](C:/RedSalamander.Perf/evidence/i26-code-review-20261001-e860022a0/build-focus-probe.ps1).

The live runner also reported 260 existing unregistered run directories under the shared test root, all classified `unexpected-test-run-dir`. I did not delete those directories or claim a clean global sandbox audit/performance qualification from these two functional runs. No production source was changed, no manual desktop input or screenshots were used, and no new whole-product, sanitizer, ARM64, external-AT or physical-input qualification is claimed.
