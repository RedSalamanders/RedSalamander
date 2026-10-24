# Terminal Renderer, Native Control, And Accessibility

> **Status — Completed 2026-08-03.** The structured Ghostty snapshot renderer, D2D/DWrite control, Kitty image layers,
> keyboard/mouse/IME/clipboard interactions, selection, and immutable UI Automation text provider are implemented.
> Durable behavior is in `../../Terminal/Terminal_EmbeddedPlugin.md`; advanced UIA mutation/history is outside v1.
> Later WIP/checkpoint language below is historical.

> **Authority:** Execution plan 3 of 6 for `Terminal_EmbeddedConPtyLibGhosttyVt_IntegrationPlan_2026-07-13.md`. The master and approved engine decision are normative.

## Status and entry gate

- **State:** WIP / blocked by plans 1–2
- **Drift baseline:** `fbf6af76f373bf253488862426f4ea70a9ff25ba`
- Enter only when selected-engine model/snapshot fixtures, queue/resource limits, service ownership, and ConPTY lifecycle tests pass on x64 and build on ARM64.
- Consume the exact plan-2 handoff: `Specs/Plans/Done/Terminal_CoreConPtyRuntimeAndModel_2026-07-22.md`, its recorded Git blob SHA and implementation commit, its recorded TerminalEngine finalization-manifest and `External/terminal-engine.lock.json` digests, successful native x64/ARM64 `Tools/Verify-TerminalEngine.ps1` results, all five recorded x64 Debug/Release/`ASan Debug` and native ARM64 Debug/Release `TerminalNative` manifest/companion paths and digests, and the exact `Specs/TestRuns/<MachineHash>/Commands/<RunId>/` core-model run path plus recorded SHA-256 digests for canonical `results.json`, `trace.txt`, `run-all-tests-results.json`, and `perf/perf_metrics.jsonl`. Rerun `Tools/Verify-TerminalNativeEvidence.ps1` on those five exact manifests with their recorded repository/engine-lock identities. Missing files, digest mismatch, a non-passing core/native case, or implicit latest-run discovery blocks entry.
- Re-run drift checks for `Plugins/Terminal`, `Common/PlugInterfaces/Terminal.h`, plugin module/child-HWND ownership, `Common/DxUi`, `Common/WindowMessages.h`, `Common/Helpers.h`, `RedSalamanderMonitor/ColorTextView*`, UIA test support, theme code, and D2D/DWrite guidance.
- Compare those paths and the master/plan-2 durable contracts against the baseline before implementation and again before closeout. Any material drift in ownership, ABI/snapshot shape, resource limits, message payload rules, renderer/device ownership, UIA teardown, test archive format, or build/toolchain assumptions is a STOP: reconcile the master and affected child plans, review the change, and advance the baseline deliberately before continuing.

## Deliverables

- `TerminalRenderSnapshot`, `TerminalRenderer`, `TerminalControl`, the immutable terminal-document/control provider seam in `TerminalAccessibility`, and render/input perf hooks, all compiled into `Plugins/Terminal/Terminal.dll`.
- Plugin-owned native child HWND with D2D/DWrite/D3D11/DXGI resources, WARP-test seam, UIA immutable snapshots, and exact `ITerminal::GetChildWindow`/parent/ownership-marker behavior.
- Selection/copy/paste, keyboard/IME/mouse/focus/scroll, hyperlink and context-command behavior.
- Complete selected-engine style/color/cursor/static-Kitty ledger coverage without a second parser.

## Renderer contract

- Consume only `shared_ptr<const TerminalRenderSnapshot>` publications; UI never walks live engine state or borrowed pointers.
- Device-independent and device-dependent ownership follows existing WIL/D2D patterns. Recreate cleanly on device loss, DPI, font, or cell-metric changes.
- Group compatible cells but break at styles, selection, cursor, hyperlinks, images, and wide-cell boundaries. Normal output paints dirty rows/rects; full paint only for documented invalidations.
- Shape and fallback complete grapheme clusters, combining marks, surrogate pairs, CJK wide cells, emoji/color glyphs, and optional ligatures. Clip every run to its allocated fixed cells.
- V1 preserves engine fixed-cell order and does not claim full Unicode BiDi paragraph reordering. RTL/Arabic fixtures prove shaping, clipping, caret/UIA truth, and documented limitation.
- Draw order is background; Kitty below-cell-background; cell/selection backgrounds; Kitty below text; glyphs/decorations; cursor/hyperlink hover; Kitty above text; transient overlays.
- Runtime OSC palette overrides remain session-local over theme defaults. High contrast preserves readable foreground/cursor/selection and may force opaque background.

## Kitty upload/cache

- Consume only engine-owned decoded representation or the Gate-0-approved bounded embedder decoder. Never decode twice.
- Checked reservations precede decode, staging, BGRA conversion, and D2D upload. Cache key includes session, image identity, engine generation, and device generation.
- Process/session ledgers come from plan 2. Across all image IDs, at most one Kitty conversion generation is queued/running for the entire terminal session; it converts changed images sequentially under reservations. Device loss releases GPU reservations; hidden/unreferenced LRU then obsolete generations are evicted; current publication/paint remains stable for the frame.
- Unsupported animation, glyph/text-sizing, file/temp/shared-memory, Sixel/iTerm, and notification paths remain unadvertised/gated.

## Native input/control

- The plugin-owned child HWND initializes posted-payload ownership during create and drains during `WM_NCDESTROY`; use queue-head-safe contiguous coalescing and stable terminal-instance/session/control generations. Payloads contain no snapshots/content. The host holds a non-owning validated HWND, never calls `DestroyWindow`, and never renders terminal content.
- Native virtual key/scan/repeat/modifier/text goes through selected engine encoding. Committed IME UTF-8 only; composition remains local until commit.
- `Ctrl+C` copies a nonempty local selection, otherwise reaches shell; `Ctrl+Shift+C/V` are explicit copy/paste. Safe/bracketed paste uses all-or-nothing queue admission and async localized warning where required.
- Mouse reporting wins when requested; Shift forces local selection. Wheel, focus reporting, cursor blink, selection autoscroll, hidden-tab timers, and context actions are deterministic and generation-checked.
- Hyperlinks are Ctrl+click with `http`, `https`, `mailto` allowlist only. No automatic URL/file/process/notification side effect.
- Control close requests the plugin service retirement; its destructor never joins or waits. Host tab close uses `ITerminal::RequestClose`/`TerminalCloseApproved`, callback drain, and idempotent `Close`; it never bypasses the plugin's running-close policy.

## Base terminal-control accessibility

- Build immutable UI-thread accessibility snapshots and the base terminal-document/control provider seam for document text, visible/range metadata, selection, caret/focus, scroll state, and native-control status. Provider/RPC threads never walk live control/model/engine trees. Every provider, cloned range, provider-owned copied text/range buffer, and selection/copy export owns or aliases an explicit RAII reservation token from plan 2's 64 MiB session/512 MiB process text ledger; aliases share one token and every physically distinct copy reserves separately before allocation.
- Provider-map membership controls discoverability only. `WM_NCDESTROY`/view retirement atomically removes the HWND/control-generation entry and marks existing providers retired, but never frees an immutable snapshot or ledger token still held by an outstanding COM provider/range/RPC call. A call that captured its immutable snapshot/token before retirement may complete from that capture. Every call beginning after retirement and every owner-thread mutation that loses HWND/session/control/snapshot revalidation returns exactly `UIA_E_ELEMENTNOTAVAILABLE`; final COM/reference release destroys the last snapshot/token.
- Marshal mutations to owner UI thread and revalidate HWND, service session, control generation, and snapshot generation. A UIA request that cannot reserve its complete provider-owned output returns exactly `E_OUTOFMEMORY`, initializes all out parameters to null/empty, publishes no partial BSTR/SAFEARRAY/range, changes no selection/scroll/focus, and remains retryable. A local selection/copy request that cannot reserve returns `HRESULT_FROM_WIN32(ERROR_NOT_ENOUGH_MEMORY)`, leaves clipboard/selection unchanged, and posts one coalesced localized `Terminal text unavailable (resource limit)` status. Successful publication transfers the complete reservation with the owning immutable object; cancellation and failed marshaling release it through RAII.
- Terminal text is exposed only through the terminal document provider, never app history. Nonces, integration messages, hidden password input, clipboard payloads, and sanitized control characters are excluded.
- Retire raw provider-map discoverability at existing DxUi teardown boundaries before HWND reuse; never place a raw snapshot pointer in that map and never revoke an outstanding provider's immutable snapshot/token early.
- This plan deliberately does not own tab, chooser, history, context-menu, unsafe-paste, or exit/restart accessibility because those surfaces do not exist at its entry gate. Plans 4–6 add and test their providers/actions; final application wiring remains Phase 7 of the master.

## Reentrant callback rule

Every context/tab/control/plugin callback may destroy the retained control, replace root, close HWND, release `ITerminal`, or run a nested loop. Capture terminal-instance/lifetime token/stable ID before app code; perform terminal local state transitions first where possible; re-resolve and validate before any later dereference/post/focus write. The host never calls `SetCallback(nullptr,nullptr)` from inside `ITerminalCallback`.

## Tests

- WARP logical draw layers, all colors/styles/underlines/cursor/selection, dirty/full behavior, static images, cap rejection, DPI/theme/high contrast/device loss.
- Grapheme/CJK/combining/emoji/color glyph/Arabic-RTL fixed-cell cases and strict no-cell-overdraw.
- Key modes, Kitty keyboard, IME commit/cancel, every mouse protocol, focus, selection, copy, atomic paste, queue rejection, hyperlink schemes.
- Single-flight publication race, failed post retry, stale token/HWND/session generation, hidden-tab zero product paints/timers, select-hidden full snapshot.
- Base terminal-document UIA worker-thread immutable reads, selection/scroll/focus/native status, teardown/HWND reuse, and password/nonce redaction. Hold providers/ranges/RPC calls across publication and teardown and prove their snapshots/tokens remain charged through final release, while new post-retirement calls return `UIA_E_ELEMENTNOTAVAILABLE`. Force provider/range/copy reservation and allocation failure and prove exact `E_OUTOFMEMORY`/`HRESULT_FROM_WIN32(ERROR_NOT_ENOUGH_MEMORY)`, null/empty out parameters, no partial disclosure, unchanged clipboard/selection/mutation state, retry success after release, and no double charge for aliases. Future pane/history/exit actions are excluded here.
- Reentrant callbacks that destroy control/root/HWND from input, context, focus, and close paths.
- DLL boundary cases load only `Plugins\Terminal.dll`, create through `IID_ITerminal`, reject `IID_IViewer`, verify child ownership/focus, hold UIA/provider/control references across tab close, and prove the module quiet point remains false until every control/provider/callback/render worker/device resource can no longer execute plugin code.

## Performance

Instrument snapshot lock/build, paint/present/input-visible, dirty/full rows, glyph/fallback/color runs, image copy/upload/CPU/GPU/eviction/rejection, hidden paints, and private-byte deltas. Metrics contain no content/path/title/URL. Add one product Commands case, `terminal_perf_renderer_control_baseline`, scoped to the native renderer/control available in this plan. Build test-enabled x64 Release and archive that exact ungated same-machine case in the canonical `Commands` area; execution plan 6 writes reviewed budgets and runs the later end-to-end cases.

## Verification

```powershell
$expectedRepositoryCommit = (& git rev-parse HEAD).Trim()
if ($LASTEXITCODE -ne 0 -or $expectedRepositoryCommit -cnotmatch '^[0-9a-f]{40}$') { throw 'Cannot resolve exact repository commit.' }
$expectedEngineLockSha256 = (Get-FileHash -LiteralPath '.\External\terminal-engine.lock.json' -Algorithm SHA256).Hash.ToLowerInvariant()
.\build.ps1 -ProjectName TerminalTests -Platform x64 -Configuration Debug
$rendererNativeManifest = .\Tools\Run-TerminalNativeTests.ps1 -Slice Renderer -Platform x64 -Configuration Debug -EvidenceRoot .\Specs\TestRuns -PassThruManifest
.\Tools\Verify-TerminalNativeEvidence.ps1 -Manifest $rendererNativeManifest -ExpectedSlice Renderer -ExpectedRepositoryCommit $expectedRepositoryCommit -ExpectedEngineLockSha256 $expectedEngineLockSha256
.\build.ps1 -ProjectName RedSalamander -Platform x64 -Configuration Debug
try {
    $env:RSBuildEnableTests = 'true'
    .\build.ps1 -ProjectName RedSalamander -Platform x64 -Configuration Release
} finally {
    Remove-Item Env:RSBuildEnableTests -ErrorAction SilentlyContinue
}
$rendererCommandsRun = .\Tools\Run-TerminalCommandsEvidence.ps1 -Scenario RendererControl -Platform x64 -Configuration Release -RepositoryCommit $expectedRepositoryCommit -EvidenceRoot .\Specs\TestRuns -PassThruRunPath
.\Tools\Test-TestRunArchive.ps1 -RunPath $rendererCommandsRun
```

Expected: the exact Renderer `TerminalNative` manifest verifies with nonzero unskipped renderer/input/accessibility cases; deterministic/WARP/base-document-UIA cases pass; hidden controls do no product paint work; callbacks survive destructive reentrancy; and the exact ungated Release baseline is archived under `Specs/TestRuns/<MachineHash>/Commands/<RunId>/` for plan 6. No future-surface accessibility claim or budget is invented in this plan.

## Child completion and move

- **Completion record:** `TERMINAL-HANDOFF-UNRESOLVED` — replace this line with the exact plan-2 closeout-commit/blob pair, evidence repository commits, engine/core identities, the Renderer native manifest/companion identities, and renderer Commands archive path/file SHA-256 values required below.

This child is complete only when all of the following are true:

1. Record/revalidate the exact plan-2 closeout commit, that commit's Done-plan Git blob object ID, every predecessor evidence `repositoryCommit`, and all handoff paths/digests named in the entry gate. Resolve the blob as `<closeout-commit>:Specs/Plans/Done/Terminal_CoreConPtyRuntimeAndModel_2026-07-22.md` and require the current Done file to hash to the same object. Run `Tools/Verify-TerminalEngine.ps1` against the locked x64 and ARM64 artifacts and `Tools/Verify-TerminalNativeEvidence.ps1` against all five exact predecessor native manifests used by this implementation; do not substitute a newer lock, build, Commands run, native manifest, or archive by directory discovery.
2. Create the x64 Release renderer Commands evidence only with `Run-TerminalCommandsEvidence.ps1 -Scenario RendererControl -Platform x64 -Configuration Release -RepositoryCommit <exact-commit> -EvidenceRoot .\Specs\TestRuns -PassThruRunPath`, validate the exact returned path with `Tools/Test-TestRunArchive.ps1 -RunPath`, and retain the exact Renderer `TerminalNative` manifest/companion pair emitted above. Retain canonical content-free Commands `results.json`, `trace.txt`, `run-all-tests-results.json`, and required `perf/perf_metrics.jsonl`; run `Tools/Verify-TerminalNativeEvidence.ps1`; record every path and lowercase SHA-256 in this plan's closeout review. Every plugin-DLL child ownership, IID rejection, module quiet, renderer, input, base-document UIA, reentrancy, hidden-control, quota-failure, and outstanding-provider teardown case must be present, passed, and unskipped.
3. Merge the plugin-owned renderer, input, immutable snapshot, native-control, and base terminal-document UIA contracts—including host non-ownership of the child HWND, provider/range ledger-token ownership, discoverability teardown, outstanding-RPC/module lifetime, exact HRESULTs, and no-partial-copy behavior—into `Specs/Terminal/Terminal_EmbeddedPane.md`; merge the applicable DxUi ownership, thread-marshalling, reentrancy, teardown, and HWND-reuse rules into `Specs/UI/UI_DxUiWinUIDesign.md`. No normative requirement owned by this child may remain only here.
4. In the same reviewed closeout change, replace this file's state line with exactly `- **State:** Done`, run the exact move below, and update the N3 next-step cell in `Specs/Plans/WIP/README.md` to exactly: `Start only with Terminal_PaneTabsProfilesAndLifecycle_2026-07-22.md; plans 1–3 are complete under Specs/Plans/Done/.`

```powershell
$source = 'Specs/Plans/WIP/Terminal_RendererControlAndAccessibility_2026-07-22.md'
$destination = 'Specs/Plans/Done/Terminal_RendererControlAndAccessibility_2026-07-22.md'
if (-not (Select-String -LiteralPath $source -SimpleMatch '- **State:** Done' -Quiet)) { throw 'Set the renderer child state to Done before moving it.' }
if (Test-Path -LiteralPath $destination) { throw "Destination already exists: $destination" }
git mv -- $source $destination
if (-not (Select-String -LiteralPath 'Specs/Plans/WIP/README.md' -SimpleMatch 'Start only with Terminal_PaneTabsProfilesAndLifecycle_2026-07-22.md; plans 1–3 are complete under Specs/Plans/Done/.' -Quiet)) { throw 'Update WIP README N3 to plan 4.' }
```

The moved plan is the plan-4 predecessor handoff. Its closeout review must expose the exact engine-lock/finalization digests, Renderer `TerminalNative` manifest/companion paths and digests, and exact renderer Commands run path/file digests consumed by plan 4.

## STOP conditions

- Rendering/input/UIA requires live or borrowed engine access from UI/provider threads.
- Full BiDi reordering becomes required without a separately reviewed scope/engine contract.
- Image accounting cannot cover all live/transient/GPU copies before allocation.
- A UIA provider/range/copy cannot retain and release the exact text-ledger token through outstanding RPC/reference lifetime, quota failure cannot return the frozen HRESULT without partial output, a stale HWND/token/generation can reacquire a view, or a control destructor can block on service work.
- The renderer/control/UIA implementation or owning HWND must leave `Terminal.dll`, the host would have to destroy/render the plugin child, or module unload cannot wait for outstanding controls/providers/callbacks without blocking UI teardown.
