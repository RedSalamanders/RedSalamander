# July 21 to August 4 Independent Code Review and Remediation Plan

**Status:** Done — findings disposition complete; bounded remainders transferred
**Review window:** 2026-07-21 00:00:00 CEST through frozen repository `HEAD` on 2026-08-04 09:28:08 CEST
**Frozen baseline:** `148f270dec3189884cde1a99bca53377e8f634ff..6aecfde8e7c09e64d48f0c2a86b6d189b67e0e7f`
**First included commit:** `52c6220867a92369dca97d7d3a6c6c411169858d`
**Branch at review:** `master`; frozen `HEAD` matched `origin/master`
**Authority:** non-normative. Current contracts under `Specs/<Domain>/`, repository guidance, and public platform contracts remain authoritative.

## 1. Outcome

The July 21-inclusive review found **41 actionable issues**: **1 P0, 14 P1,
23 P2, and 3 P3**. Nineteen packages are newly owned here: eighteen Terminal
packages and one Microsoft Drive defect discovered in surrounding code.
Twenty-two already-owned findings were revalidated and are explicitly routed to
their existing Emberglass and/or Floodgate owners; their detail is repeated here
so this audit remains independently reviewable, but they must not be implemented
twice.

The immediate release blocker is `J21-TERM-01`: `ITerminal::Close()` can call
`ClosePseudoConsole` on the UI thread while a reader/process cycle prevents it
from returning on supported pre-26100 Windows builds. Terminal v1 therefore
cannot remain classified as fully closed. The next gates are output-message
coalescing, Kitty image work off the UI paint path, DLL window-class retirement,
UIA text-unit correctness, readable/accessible security confirmations, and an
honest green Full closeout.

This audit changed no production source. It creates only this WIP record and
updates the WIP routing index.

### Priority meanings

- **P0:** release blocker with a credible application hang or equally severe
  failure on a supported configuration.
- **P1:** data integrity, security, unload safety, major responsiveness,
  accessibility, or false release-gate issue.
- **P2:** correctness, boundedness, maintainability, reproducibility, or test
  fidelity issue that should be scheduled after P0/P1 work.
- **P3:** bounded improvement or currently unreachable scale/efficiency issue.

## 2. Exact scope and repository state

The range contains 41 commits. The union of paths touched by those commits is
566; the final base-to-HEAD delta contains 551 paths, 164,927 insertions, and
3,358 deletions. After excluding archived test evidence and imported upstream
candidate patch bodies, 259 primary first-party code, build, configuration, and
test paths matched the reproducible filter below. Six changed `.rc` resource
files were reviewed separately under the localization/resource contract, for a
265-path primary-plus-resource inventory.

```powershell
$reviewPaths = @(
    git log --format= --name-only `
        '148f270dec3189884cde1a99bca53377e8f634ff..6aecfde8e7c09e64d48f0c2a86b6d189b67e0e7f' |
        Where-Object { $_ } |
        Sort-Object -Unique |
        Where-Object {
            $_ -notlike 'Specs/TestRuns/*' -and
            $_ -notlike 'External/TerminalEngine/CandidatePatches/*' -and
            [IO.Path]::GetExtension($_).ToLowerInvariant() -in @(
                '.cpp', '.h', '.ps1', '.psm1', '.yml', '.json',
                '.vcxproj', '.filters', '.proj', '.sln', '.props',
                '.targets', '.toml'
            )
        }
)
if ($reviewPaths.Count -ne 259) { throw "Review inventory drift: $($reviewPaths.Count)" }
$resourcePaths = @(
    git log --format= --name-only `
        '148f270dec3189884cde1a99bca53377e8f634ff..6aecfde8e7c09e64d48f0c2a86b6d189b67e0e7f' |
        Where-Object { $_ -like '*.rc' } |
        Sort-Object -Unique
)
if ($resourcePaths.Count -ne 6) { throw "Resource inventory drift: $($resourcePaths.Count)" }
```

Coverage included:

1. Every base-to-HEAD first-party code/build/config/test path and the commits
   that introduced it.
2. Direct callers and consumers, public plugin ABI records, process and COM
   ownership, worker/callback teardown, UI-thread boundaries, filesystem
   publication and source/destination identity, settings recovery, and
   localization.
3. Runtime and source-contract tests, the Full-suite policy, authoritative
   domain specs, WIP/Done ownership, committed `Specs/TestRuns` evidence, CI,
   release, and Winget workflows.
4. Surrounding unchanged code when it controlled an edited path or could turn
   a local change into data loss, stale state, an unload hazard, or false proof.

`git diff --check` on first-party code was clean after excluding imported
`External/TerminalEngine/CandidatePatches/*.patch` bodies and archived Done
prose. Trailing whitespace in those imported/review-history artifacts is not a
runtime finding.

### Concurrent worktree changes

The worktree was clean when the review baseline was frozen. During the review,
another actor modified these files:

- `Plugins/FileSystemS3/Factory.cpp`
- `Plugins/FileSystemS3/FileSystemS3.Directory.cpp`
- `Plugins/FileSystemS3/FileSystemS3.IO.cpp`
- `Plugins/FileSystemS3/FileSystemS3.Internal.h`
- `Plugins/FileSystemS3/FileSystemS3.S3.cpp`
- `Tools/Build-GhosttyTerminalRuntime.ps1`
- `Tools/Tests/TerminalPluginSourceContracts.Tests.ps1`

Those changes are preserved and excluded from the frozen finding evidence.
References in this document are to committed `HEAD` unless explicitly stated.
The S3 edits appear to be an in-progress source-revision repair and must receive
a separate stabilized-diff review and build/test pass before they can close
`EG-S3-02`. The concurrent seven-file set grew from roughly +520/-49 to
+988/-77 while this audit was being finalized, so no intermediate state is
treated as reviewed or complete. The two Terminal tool edits do not implement
the exact export-set check in `J21-TERM-12`. At final observation their new
source-contract assertion required the builder not to contain
`$sourceStatus.Count -ne $allowedStatus.Count`, while the dirty builder still
contained that condition; this transient dirty state is internally red and must
not be merged without its own test run.

## 3. Review method and proof standard

The audit used independent subsystem reviews followed by a second evidence pass.
A candidate was promoted only when the current call graph supported a concrete
failure scenario. Preliminary concerns that depended only on naming, a single
regex, or an unobserved timing assumption were rejected or retained as leads in
section 9.

No build, selftest, generator, network mutation, or source edit was run. That is
deliberate: this is a static review and executable remediation plan, not a claim
that the repairs are complete. Each work package starts with a deterministic RED
proof, makes the smallest boundary-local repair, runs focused Debug/Release
coverage, and then runs the repository gate required by its risk. Performance
packages must also archive before/after Release evidence under `Specs/TestRuns/`.

## 4. Simplicity and architecture rules

The repair architecture must remain smaller than the problem:

1. Keep one Terminal session owner. Add only four narrow mechanisms where the
   current object needs them: a ConPTY close state machine, one generation-based
   output notifier, one bounded input drain, and immutable render/accessibility
   snapshots. Do not add a service locator, event bus, generic task framework,
   or second lifetime graph.
2. Keep `WindowProc` as routing. Mouse, keyboard/IME, confirmation, output, and
   destroy behavior belong in named handlers with directly testable state.
3. Reuse a shared Common UIA text-unit boundary helper instead of maintaining a
   second word/line implementation in Terminal.
4. At storage/network boundaries, carry one exact identity record through the
   operation. Do not repair races with an extra pathname restat, retry loop, or
   provider-specific test bypass.
5. Preserve one owner per existing requirement. Rows marked **ROUTED** below are
   evidence refreshes, not parallel implementation queues.
6. Remove or reserve unused ABI promises before adding abstractions to support
   states that v1 never produces.

## 5. Finding and ownership matrix

| ID | Priority | Window classification | Finding | Ownership |
| --- | --- | --- | --- | --- |
| J21-TERM-01 | P0 | Introduced | UI-facing close can deadlock in `ClosePseudoConsole` on supported Windows | OWNED here |
| J21-TERM-02 | P1 | Introduced | Per-read output posts cause a UI/UIA message storm and can lose the final update | OWNED here |
| J21-TERM-03 | P1 | Introduced | Legal Kitty images are copied/converted/uploaded from UI paint under the engine lock | OWNED here |
| J21-TERM-04 | P1 | Introduced | DLL-owned Terminal window class survives module unload | OWNED here |
| J21-TERM-05 | P1 | Introduced | UIA text ranges ignore requested `TextUnit` semantics | OWNED here |
| J21-TERM-06 | P1 | Introduced | Security confirmations are unreadable in light/high-contrast themes and absent from the UIA snapshot | OWNED here |
| J21-TERM-07 | P1 | Introduced | Terminal was archived as complete with a red Full aggregate and no quarantine | OWNED here; coordinate Atlas |
| J21-TERM-08 | P2 | Introduced | Reader-thread construction failure leaves a resumed hidden shell alive | OWNED here |
| J21-TERM-09 | P2 | Introduced | Input hot path allocates/pins unnecessarily and admits 3,841 active-plus-queued requests | OWNED here |
| J21-TERM-10 | P2 | Introduced | Shell resolution trusts ambient current-directory/PATH search | OWNED here |
| J21-TERM-11 | P2 | Introduced | Public location records receive only shape validation in `Open` | OWNED here |
| J21-TERM-12 | P2 | Introduced | Ghostty builder does not enforce the documented closed export set | OWNED here |
| J21-TERM-13 | P2 | Introduced | Final Terminal evidence paths cited by the authoritative spec are absent | OWNED here |
| J21-TERM-14 | P2 | Introduced | Preferences schema, loader diagnostics, and UIA name bypass resources | OWNED here |
| J21-TERM-15 | P2 | Introduced | Regex/source checks certify behavior they do not execute | OWNED here |
| J21-TERM-16 | P2 | Introduced | 444-line Terminal `WindowProc` interleaves unrelated state machines | OWNED here |
| J21-TERM-17 | P2 | Introduced | First public Terminal ABI exposes states/fields with no producer or consumer | OWNED here |
| J21-TERM-18 | P3 | Introduced/surrounding CI | Clean CI repeatedly downloads/builds the pinned Terminal toolchain/runtime | OWNED here |
| EG-SET-02 | P1 | Introduced | Settings backup restats then renames a pathname, not the diagnosed file | ROUTED to Emberglass |
| EG-S3-02 | P1 | Pre-existing in touched scope | S3 COPY/MOVE does not pin the source revision through copy and delete | ROUTED to Emberglass |
| EG-CURL-01 | P1 | Pre-existing in touched scope | Returned Curl writers retain a raw filesystem pointer | ROUTED to Emberglass |
| EG-MSD-02 | P1 | Pre-existing in touched scope | Microsoft Drive accepts unvalidated or unbounded ranged responses | ROUTED to Emberglass; review escalates its existing P2 label |
| EG-ABI-01 | P2 | Pre-existing in touched scope | Packed `FileInfo` traversal reads fields before proving the fixed header/name is in bounds | ROUTED to Emberglass |
| FG-P0-3-FALLBACK-SIZE-ZERO | P1 | Pre-existing, contrary evidence newly proven | Curl can promote an upload whose remote size could not be proved | ROUTED to Floodgate; reopen its proof-depth row |
| EG-SEC-01 | P1 | Pre-existing in touched workflow | Winget event data is interpolated into PowerShell source before validation | ROUTED to Emberglass |
| EG-SEC-02 | P1 | Pre-existing in touched workflow | Build/test jobs receive persisted repository-write credentials | ROUTED to Emberglass |
| EG-SET-01 | P2 | Introduced | Early settings read failures erase the fact that a source exists | ROUTED to Emberglass |
| EG-UI-01 | P2 | Introduced | File Operations focus/UIA traversal differs from visual order | ROUTED to Emberglass |
| EG-UI-02 | P2 | Introduced | File Operations accessibility announcement history never prunes | ROUTED to Emberglass |
| EG-DOC-01 | P2 | Introduced | Changed selftest guidance contradicts itself and review routing names a nonexistent WIP plan | ROUTED to Emberglass |
| R4-UI-01 | P2 | Introduced | Theme-overlay paint and hit testing use different geometry during animation | ROUTED to Emberglass |
| R4-S3-02 | P1 | Pre-existing in touched scope | Multipart-abort retries do not schedule themselves | ROUTED to Floodgate implementation; Emberglass ledger |
| R4-TEST-02 | P2 | Pre-existing in touched scope | Bridge test decorator drops optional writer capabilities | ROUTED to Floodgate implementation; Emberglass ledger |
| J21-MSD-01 | P2 | Pre-existing, newly found | Microsoft Drive seek performs unchecked signed addition | OWNED here; coordinate Floodgate |
| EG-SEC-03 | P2 | Pre-existing in touched workflow | WingetCreate is mutable at the credential boundary | ROUTED to Emberglass |
| EG-CI-01 | P2 | Pre-existing in touched workflow | Token-created release event may never invoke Winget publication | ROUTED to Emberglass |
| EG-CI-02 | P2 | Pre-existing in touched workflow | Release assets are published in two non-atomic steps | ROUTED to Emberglass |
| R4-CI-01 | P2 | Pre-existing in touched workflow | Winget PR submission is neither resumable nor serialized | ROUTED to Emberglass |
| EG-BUILD-01 | P2 | Pre-existing in touched workflow | Official compiler patch identity is mutable and unrecorded | ROUTED to Emberglass |
| EG-CURL-02 | P3 | Pre-existing in touched scope | One slow endpoint's delete limit serializes independent Curl connections | ROUTED to Emberglass |
| R4-S3-03 | P3 | Pre-existing in touched scope | Known-unsupported giant S3 uploads fail only after hundreds of GiB | ROUTED to Floodgate implementation; Emberglass ledger |

## 6. Newly owned Terminal work packages

### J21-TERM-01 — Move all potentially blocking ConPTY teardown off the UI thread

**Priority:** P0
**Effort / repair risk / confidence:** L / high / high

**Evidence and failure.** `Plugins/Terminal/Terminal.cpp:4237-4249` calls
`prepareClose()` synchronously before submitting close work. The reader checks
stop/closing at `:1965-1969` separately from its following synchronous
`ReadFile`. `prepareClose()` requests stop, ignores the result of
`CancelSynchronousIo`, and resets the pseudoconsole at `:4274-4283`; the
kill-on-close job is not reset until `finishClose()` at `:4320-4329`.
`Common/MinimumOsVersion.h:19-22` supports Windows build 22000. Microsoft
documents that before build 26100 `ClosePseudoConsole` may wait indefinitely
unless its output pipe is closed/drained, while `CancelSynchronousIo` does not
cover I/O that is not yet pending and does not wait for cancellation completion:
[ClosePseudoConsole](https://learn.microsoft.com/en-us/windows/console/closepseudoconsole),
[CancelSynchronousIo](https://learn.microsoft.com/en-us/windows/win32/fileio/cancelsynchronousio-func).

A reader can pass the predicate, the UI can cancel before `ReadFile` becomes
pending, and the reader can then block. The UI enters `ClosePseudoConsole`, the
child remains alive because the job is still open, and the job closure needed to
break the cycle is unreachable.

**Smallest safe repair.** The UI path only transitions to Closing, stops new
posts/input, retires UIA, removes the child HWND, and submits one module-pinned
worker. That worker owns a documented, bounded sequence: stop producers, close
input, terminate/close the job when graceful shutdown exceeds its deadline,
ensure the output reader closes its handle, call `ClosePseudoConsole`, join, and
free runtime state. Keep one state machine; do not add a detached thread.

**Required proof.** Add a seam that pauses the reader after its predicate and
before `ReadFile`, an uncooperative child, a simulated pre-26100 close order,
repeated Close, host destruction, retained UIA provider, and unload/reload. The
UI-facing call must return within a fixed bound in every case.

**Done / STOP.** Done only with deterministic RED-before/GREEN-after teardown
coverage and a green Full run. STOP if any proposed sequence calls a potentially
blocking console/process API on the UI thread or can unload while a callback,
window, reader, or provider survives.

### J21-TERM-02 — Coalesce output notification and publish the newest generation

**Priority:** P1
**Effort / repair risk / confidence:** M / medium / high

**Evidence and failure.** `readerMain()` posts one
`kTerminalOutputReadyMessage` for every 16 KiB read at
`Terminal.cpp:1961-1983`; both normal and final `PostMessageW` results are
ignored at `:1982` and `:1990`. The UI handler at `:4809-4812` calls
`publishAccessibilitySnapshot()`, which calls `formatScreen()` at `:3850-3857`.
That routine locks `_terminalMutex`, allocates, exports the entire formatted
buffer, and can copy up to 32 MiB at `:3205-3234`. A burst can therefore fill the
posted-message queue with redundant whole-screen UIA work, block the reader on
the same mutex, and lose the final refresh if posting fails.

**Smallest safe repair.** Use one atomic pending flag plus a monotonically
increasing output generation. Post only on false-to-true. The UI consumes the
newest generation, clears the flag, then rechecks/reposts so output arriving in
the handoff cannot be lost. If the initial post fails, the posting thread must
roll back only the flag value it owns, recheck the generation, and establish a
bounded guaranteed wakeup; the simplest fallback is to invalidate the HWND and
let `WM_PAINT`/`WM_GETOBJECT` consume the same newest-generation routine while
later output can retry the post. Never leave `pending == true` without either a
queued message or another guaranteed UI wake. Publish UIA only when clients are
listening or on a bounded cadence.

**Required proof.** Deterministic multi-megabyte output storm, forced post
failure/queue saturation, output arriving during flag clear, teardown, and final
screen/UIA equality. Record bytes, reads, posts, coalesced reads, maximum pending
count, formatter bytes/time, paint latency, and input-to-visible latency in
archived Release evidence.

**Done / STOP.** Done when posted work stays O(1), newest content is never stale,
and reader/UI responsiveness meets an accepted bound. STOP if a flag-clear race
can lose the last generation.

### J21-TERM-03 — Prepare Kitty images outside UI paint and outside the reader lock

**Priority:** P1
**Effort / repair risk / confidence:** L / high / high

**Evidence and failure.** `renderStructuredScreen()` calls
`snapshotKittyGraphicsLocked()` from `WM_PAINT` while holding `_terminalMutex` at
`Terminal.cpp:3522-3538`. The snapshot loop at `:3237-3460` enumerates
placements, allocates/copies images, and may convert a legal 16-million-pixel,
64-MiB image one pixel at a time. `ensureKittyBitmap()` at `:3463-3481` then
uploads the bitmap on the UI path. A terminal-controlled legal frame can freeze
the application and stop the reader behind the same mutex.

**Smallest safe repair.** Keep Ghostty access on one engine side and publish an
immutable, generation-tagged converted image snapshot by swap. The UI performs
only size- and time-budgeted bitmap creation/draw for the newest snapshot;
stale generations are discarded. Define an accepted per-frame upload limit and
defer/chunk larger legal snapshots or lower the accepted image cap so a single
64-MiB `CreateBitmap` cannot consume an unbounded frame interval. Do not create
a general render graph or a second terminal owner.

**Required proof.** Maximum permitted raw and PNG fixtures, repeated generation
replacement, device loss, close during conversion, and stale-result rejection.
Archive decode, conversion, lock-hold, upload, total-frame, memory-high-water,
and input/output latency before and after.

**Done / STOP.** Done when terminal-controlled input cannot make one UI paint
exceed the accepted conversion/upload frame-latency bound. STOP if a worker
borrows Ghostty-owned memory after releasing the engine lock.

### J21-TERM-04 — Retire the DLL-owned window class at the unload quiet point

**Priority:** P1
**Effort / repair risk / confidence:** S / medium / high

**Evidence and failure.** `Terminal.cpp:1459-1473` registers a class whose
WndProc is inside `Terminal.dll`. `Plugins/Terminal/Factory.cpp:96-104` clears
metadata on shutdown, and `RedSalamander/PluginModuleLifecycle.cpp:55-78` can
subsequently free the module. No `UnregisterClassW` exists. Microsoft states
that DLL-registered classes are not automatically unregistered when the DLL
unloads: [UnregisterClassW](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-unregisterclassw).
A refresh/reload can retain a stale WndProc or encounter an existing class tied
to the old module.

**Smallest safe repair.** Give module scope one atom/registration owner. After
all Terminal HWNDs have been destroyed and producers stopped, unregister it;
`CanUnloadNow` must remain false while it cannot be retired. Registration and
unregistration must be idempotent.

**Required proof.** Create/close/shutdown/unload/reload/create cycles, including
a different module base/version and a retained provider. Assert no old class or
WndProc remains.

**Done / STOP.** Done when module unload is impossible while the class or any
window survives. STOP if shutdown unregisters while a child HWND still exists.

### J21-TERM-05 — Honor UIA `TextUnit` semantics with one shared boundary helper

**Priority:** P1
**Effort / repair risk / confidence:** M / medium / high

**Evidence and failure.** `TerminalAccessibility.cpp:432-452` ignores the unit
passed to `Move`; `:455-485` does the same for `MoveEndpointByUnit`.
`ExpandToEnclosingUnit` at `:275-314` treats unsupported `Format` and `Page` as a
single UTF-16 code unit instead of the next supported larger unit. This also
splits surrogate pairs. Microsoft requires movement in the requested unit, or
the next larger supported unit:
[ITextRangeProvider::Move](https://learn.microsoft.com/en-us/windows/win32/api/uiautomationcore/nf-uiautomationcore-itextrangeprovider-move).
Narrator word/line navigation and returned moved-unit counts are therefore
incorrect.

**Smallest safe repair.** Extract the already-tested character/word/line logic
around `Common/DxUi/DxUi.Accessibility.cpp:2307` into a pure Common helper and
use it from both providers. Support surrogate-safe Character, Word,
Line/Paragraph, and Document; map unsupported units to the next supported
larger unit.

**Required proof.** Direct provider tests for forward/backward character,
surrogate/grapheme, word, CRLF/line, paragraph/page fallback, document,
collapsed/noncollapsed ranges, endpoint crossing, and returned counts.

**Done / STOP.** Done when Terminal and DxUi share the same boundary rules and
UIA conformance cases pass. STOP if Terminal keeps a copied local variant.

### J21-TERM-06 — Make security confirmations themed and screen-reader accessible

**Priority:** P1
**Effort / repair risk / confidence:** M / medium / high

**Evidence and failure.** The host supplies light/dark/high-contrast state at
`FolderWindow.Viewers.cpp:531-545`, but `Terminal::SetTheme()` ignores `dark`
and `highContrast` at `Terminal.cpp:3935-3954`. `drawConfirmation()` always
fills a `#202020` panel at `:2427-2437` and draws normal theme foreground over
it; the built-in light theme supplies black normal text. The same surface asks
users to approve unsafe paste, hyperlinks, and OSC52 clipboard writes. UIA
publication at `:3850-3857` contains only `formatScreen()` output, so the prompt
and its Accept/Cancel actions are absent from the accessibility snapshot while
input is suppressed behind it.

**Smallest safe repair.** Keep one `ConfirmationState` for rendering, pointer,
keyboard, and accessibility. Derive panel/text/button colors from the supplied
theme, with an explicit high-contrast branch and selection foreground. Expose
the prompt text and both actions through standard accessible controls or the
smallest equivalent UIA children; do not maintain a second prompt model.

**Required proof.** All three confirmation kinds in built-in light, dark,
custom, and Windows high-contrast themes; automated minimum-contrast checks;
keyboard, pointer, Narrator/UIA discovery and invoke; focus restoration;
accept/reject/close scrubbing; no content disclosure for OSC52.

**Done / STOP.** Done when a keyboard-only or screen-reader user can perceive
and safely accept/reject every security prompt. STOP if Enter can accept a
prompt that UIA cannot identify.

### J21-TERM-07 — Reopen Terminal closeout until the authoritative aggregate is green

**Priority:** P1
**Effort / repair risk / confidence:** M / low / high

**Evidence and failure.** `Specs/Terminal/Terminal_EmbeddedPlugin.md:457-477`
records the final Full run as 1,174 passed and 5 failed, then classifies the
failures as order sensitivity after isolated reruns. Repository policy in
`Specs/Testing/Testing_SelfTests.md:110-112` and the WIP closeout rules keeps
every non-PASSED result blocking unless it has an approved quarantine; the
quarantine ledger is empty. Nevertheless the frozen-HEAD
`Specs/Plans/WIP/README.md` blob at `6aecfde8e:3-5` and `:196-199` says Terminal
v1 is complete and its plans are archived. The earlier
1,180/0/53 Full run predates later Terminal implementation commits and cannot
certify final code.

**Smallest safe repair.** Reopen only the Terminal closeout gate, not all 13
archived implementation histories. Reproduce the five failures in Full order
and fix their shared-state/isolation cause. If immediate repair is impossible,
record an explicit policy-compliant blocking quarantine with owner, expiry, and
evidence; quarantine tracks the blocker but does not satisfy closeout. Then run
Full on the final binary.

**Required proof.** One final green Full aggregate on the exact closing commit,
plus focused Terminal gates. Repeated isolated passes are diagnostic evidence,
not a substitute for the aggregate.

**Done / STOP.** Done when the WIP index, Terminal spec, test policy, and actual
aggregate agree. STOP if a failure is relabeled without the repository's formal
quarantine process.

### J21-TERM-08 — Roll back a resumed shell when reader startup fails

**Priority:** P2
**Effort / repair risk / confidence:** S / medium / high

**Evidence and failure.** The shell resumes at `Terminal.cpp:1775-1779`; job and
process handles move into members at `:1780-1781`; reader `std::jthread`
construction can throw at `:1783-1799`. The suspended-process guard has already
been released, and `Open()` converts the HRESULT into a retained Failed tab at
`:1443-1449`. Thread-resource exhaustion can leave a hidden shell/job/HPCON
alive with no reader until the tab is explicitly closed.

**Smallest safe repair.** Keep job/process local and a post-resume rollback guard
armed until both workers are constructed and installed. On failure, stop the
integration worker and terminate/reset the job, process, pipes, and
pseudoconsole.

**Required proof.** Inject reader-thread construction failure after resume and
assert prompt process exit, no owned session handles, Failed lifecycle, and
unload eligibility.

**Done / STOP.** Done when every partially started session has one cleanup owner.

### J21-TERM-09 — Schedule one input drain without per-request allocation and count active work

**Priority:** P2
**Effort / repair risk / confidence:** S / medium / high

**Evidence and failure.** `writeInput()` acquires a module pin and allocates
`InputDrainContext` before the queue lock at `Terminal.cpp:2091-2100`; if a
drain is already scheduled it immediately returns at `:2110-2113` and destroys
that unused context/pin. The active request is popped at `:2150-2158`, but
admission at `:2102-2108` counts only queued nodes. While one write blocks, the
documented aggregate cap of 3,840 can be 3,841.

**Smallest safe repair.** Enqueue and update queued-plus-active accounting under
the existing mutex. Allocate/pin only on the false-to-true scheduled transition;
roll back the exact newly enqueued request if callback submission fails.

**Required proof.** Burst keyboard/mouse/reply traffic proving one context per
drain epoch, blocked-writer boundaries at 3,839/3,840/3,841, submission failure,
close/drain scrubbing, and input latency/queue-depth metrics.

**Done / STOP.** Done when cap and allocation counts match the contract exactly.

### J21-TERM-10 — Resolve inbox shells without ambient executable search

**Priority:** P2
**Effort / repair risk / confidence:** M / medium / high

**Evidence and failure.** `FindExecutable()` uses
`SearchPathW(nullptr, ...)` at `Terminal.cpp:473-481`; `resolveShell()` uses it
for `wsl.exe`, `pwsh.exe`, `powershell.exe`, and `cmd.exe` at `:1567-1626`.
Microsoft documents that the default search can include the calling
RedSalamander process's current directory:
[SearchPathW](https://learn.microsoft.com/en-us/windows/win32/api/processenv/nf-processenv-searchpathw).
An attacker-controlled process CWD or PATH can select a same-name program. The
later child working directory passed to `CreateProcessW` is a separate value and
is not the search input claimed here.

**Smallest safe repair.** Resolve inbox `cmd`, `wsl`, and Windows PowerShell from
validated Windows/System32 paths. Give PowerShell 7 a separate documented trust
policy: registered install/App Paths or an explicit validated user setting,
with current-directory search excluded.

**Required proof.** Hostile current directory and PATH entries for every shell,
plus supported installed/portable `pwsh` cases. Record the product decision if
portable discovery changes.

**Done / STOP.** Done when inbox binaries never depend on ambient CWD search.

### J21-TERM-11 — Apply one kind-aware validator at every Terminal ABI entry

**Priority:** P2
**Effort / repair risk / confidence:** M / medium / high

**Evidence and failure.** `validateLocation()` at `Terminal.cpp:1361-1365`
checks only record size and span pointer shape. `Open()` accepts that at
`:1373-1440`, and a Windows value becomes `CreateProcessW`'s working directory
at `:1722-1735`. Later Update/Insert paths add partial kind/absolute checks,
creating inconsistent public semantics. Relative, control-containing,
conflicting-kind, or enormous spans can cross a `noexcept` ABI; current host
callers happen to construct valid records.

**Smallest safe repair.** One pure kind-aware validator with bounded spans,
absolute Windows/UNC rules, WSL distribution/path rules, and required/forbidden
fields. Reuse it in Open, Update, and Insert.

**Required proof.** Every enum value, undersized-record rejection, exact and
forward-larger `sizeBytes` acceptance, null/non-null span combinations,
over-limit span rejection, relative/control values, conflicting fields, valid
UNC/extended paths, and WSL edges.

**Done / STOP.** Done when no entry point carries its own weaker variant.

### J21-TERM-12 — Enforce an exact Ghostty runtime export set

**Priority:** P2
**Effort / repair risk / confidence:** S / low / high

**Evidence and failure.** Committed
`Tools/Build-GhosttyTerminalRuntime.ps1:335-412` checks that each required
export exists, but unlike imports it never rejects additional exports.
`Tools/Tests/TerminalPluginSourceContracts.Tests.ps1:266-303` checks lexical
name presence only. `Specs/Terminal/Terminal_EmbeddedPlugin.md:278-280`
promises a closed required-export set. A deliberate runtime upgrade updates the
fixed hash; an unnoticed extra ABI/attack-surface export can then pass the
advertised gate.

**Smallest safe repair.** Compare sorted unique observed exports to one exact
expected array; fail on missing, duplicate, or additional entries. Reuse the
same expected identity in builder and loader validation rather than adding a
second hand-maintained list.

**Required proof.** Synthetic exact, missing, duplicate, and extra export sets,
plus the real x64/ARM64 pinned runtime check.

**Done / STOP.** Done when the test executes comparison behavior rather than
matching source text.

### J21-TERM-13 — Persist every final Terminal evidence path cited by the spec

**Priority:** P2
**Effort / repair risk / confidence:** S-M / low / high

**Evidence and failure.** `Terminal_EmbeddedPlugin.md:448-456` cites
`Specs/TestRuns/7d3a1247382a/Commands/2026-08-03_214720/` and
`2026-08-04_004604/`. Neither exists in committed `HEAD`; the only tracked
Terminal run is the earlier `2026-08-02_005913/`. `Specs/TestRuns/README.md`
requires curated closeout evidence that a fresh clone can verify.

**Smallest safe repair.** Recover and curate compact final results/metric rows
and force-add them. An alternative publication is acceptable only if it is
immutable and retention-guaranteed for the lifetime of the authoritative claim;
ordinary expiring CI artifacts are not durable evidence. Otherwise remove or
qualify the unsupported claims. Add a repository-relative link-existence check
for authoritative `Specs/TestRuns` citations.

**Required proof.** Fresh-clone link check and `Tools/Test-TestRunArchive.ps1`
over each curated archive; no secrets, machine paths, or oversized logs.

**Done / STOP.** Done only when every authoritative evidence link resolves.

### J21-TERM-14 — Localize all Terminal user-facing schema and diagnostics

**Priority:** P2
**Effort / repair risk / confidence:** M / low / high

**Evidence and failure.** `Terminal.cpp:365-471` embeds the complete Preferences
title, labels, descriptions, and options in English. User-visible loader errors
are built in English at `TerminalRuntimeLoader.cpp:302-363`, `:395-443`, and
`:518-550`; `TerminalAccessibility.cpp:686-691` hard-codes the accessible name.
`TerminalResources.rc` already supplies the module's correct resource path.

**Smallest safe repair.** Build display text from resource strings while keeping
configuration keys and option values invariant. Put formatted diagnostics and
the UIA name in resources with positional placeholders.

**Required proof.** Resource/source placeholder parity, one non-English
satellite smoke test, stable schema keys/values, and a reviewed source-contract
allowlist that rejects new user-facing literals while permitting invariant
configuration keys, protocol tokens, automation IDs, and runtime symbol names.

**Done / STOP.** Done when translated builds have no split English Terminal
surface. STOP if stable stored values are translated.

### J21-TERM-15 — Replace behavioral regex certification with executable seams

**Priority:** P2
**Effort / repair risk / confidence:** L / low / high

**Evidence and failure.** The source test at
`TerminalPluginSourceContracts.Tests.ps1:51-74` inspects only lexical bodies, so
`Close()` passes despite blocking work in `prepareClose()` and queue constants
pass despite active-count behavior. The pane lifecycle at
`Commands.SelfTest.Navigation.cpp:9649-9659` invokes path insertion but asserts
only tab/HWND reuse, not emitted bytes or screen state. Native Terminal tests
starting near `Terminal.cpp:665` exercise runtime objects, not the real window,
transport, queue, close, theme, or UIA range implementation.

**Smallest safe repair.** Add narrow fake/injectable transport, post scheduler,
thread-start, and render/provider seams around the existing object. Keep source
checks only for static ABI/export/build constraints that cannot be exercised.

**Required proof.** The RED tests from `J21-TERM-01` through `-12` must fail on
the frozen implementation and pass after each repair. Path insertion must
assert exact bytes and resulting screen state.

**Done / STOP.** Done when no runtime behavior is certified solely by regex.

### J21-TERM-16 — Reduce `WindowProc` to explicit routing

**Priority:** P2
**Effort / repair risk / confidence:** M / medium / high

**Evidence and failure.** `Terminal.cpp:4394-4837` is a roughly 444-line
WndProc. Mouse selection, application mouse, capture, hyperlink, confirmation,
autoscroll, keyboard, IME, surrogate buffering, clipboard, output, UIA, and
destroy policies share one switch and flags. Event sequences cannot be unit
tested without driving the whole HWND, and a local change requires reasoning
across unrelated state.

**Smallest safe repair.** After behavior tests exist, extract named
`OnMouse*`, `OnKeyboard`, `OnIme`, `OnOutputReady`, `OnClipboard`, and
`OnDestroy` handlers. Keep small explicit confirmation, selection, and
application-mouse state; each WndProc case routes and returns.

**Required proof.** Capture loss, double-click/drag, application mouse,
confirmation interception, IME/surrogates, output during destroy, and default
processing event sequences before and after extraction.

**Done / STOP.** Done when the switch expresses dispatch, not policy. STOP if
extraction creates a generic event framework or changes behavior without a RED
case.

### J21-TERM-17 — Shrink or reserve unimplemented v1 ABI promises

**Priority:** P2
**Effort / repair risk / confidence:** S before release, L after / medium / high for non-use

**Evidence and failure.** `Common/PlugInterfaces/Terminal.h:16-27` publishes
`DrainingAfterRootExit`, `Retiring`, and `Quarantined`, but production only
names them in a status switch. `TerminalActivitySnapshot` fields
`passwordInputActive`, `alternateScreenActive`, and `tuiInputOwner` and
`TerminalViewState` exit/final fields are never populated. Theme ANSI fields
are neither populated by the host nor consumed. The spec at
`Terminal_EmbeddedPlugin.md:479-488` says several future semantics are not
latent v1 promises.

**Smallest safe repair.** Before first external publication, remove impossible
states/fields while retaining reserved tail space. If compatibility already
exists, specify them as reserved/always-zero and add real fields only alongside
implementation and behavioral tests.

**Required proof.** Search all producers/consumers, ABI size/layout tests for
x64/ARM64, plugin compatibility smoke, and a release-owner decision recording
whether v1 is already frozen.

**Done / STOP.** Done when the public record says only what v1 can produce.
STOP if external compatibility status is unknown.

### J21-TERM-18 — Cache pinned Terminal build inputs without weakening validation

**Priority:** P3
**Effort / repair risk / confidence:** S / low / high

**Evidence and failure.** `Directory.Build.targets:105` invokes the Terminal
builder for Terminal builds. On a clean job the script downloads/expands Zig,
fetches Ghostty, and builds the runtime at
`Build-GhosttyTerminalRuntime.ps1:136-294`. `build-reusable.yml:153` caches
vcpkg but not `.build/TerminalEngine`, so independent clean platform/config jobs
repeat deterministic network/build work.

**Smallest safe repair.** Cache downloads/source/toolchain/product by platform
and hashes of builder, pinned patch, archive/runtime identity, and notice inputs.
Always rerun commit, SHA-256, machine, license, import, export, and feature
validation after restore; the cache is not a trust boundary.

**Required proof.** Cold and warm CI timings, deliberate stale/wrong cache
entries rejected, x64/ARM64 identity checks, and no credential in the cache.

**Done / STOP.** Done when warm jobs avoid redundant acquisition/build while
all security checks still execute.

## 7. Revalidated surrounding work packages

The packages in this section are deliberately self-contained. Existing review
rows are reconciled in Emberglass; where Emberglass names Floodgate as the
implementation owner, Floodgate remains the single execution queue. Update the
existing owner's status and evidence; do not create a second implementation
branch from this document.

### EG-SET-02 — Rename the exact diagnosed settings file

**Priority / ownership:** P1 / ROUTED to Emberglass
**Effort / repair risk / confidence:** M / high / high

**Evidence and failure.** `Common/Common/SettingsStore.cpp:557-620` opens and
stamps a pathname, closes the identity-bearing handle, restats the path, then
calls pathname-based `MoveFileExW`. A noncooperating writer can replace bad A
with valid B after the restat; recovery then moves B to `.bad` and leaves the
canonical path empty.

**Smallest safe repair.** Retain a delete-capable handle for the exact diagnosed
file, compare identity through that handle, and rename that handle with
`SetFileInformationByHandle`. Fail closed if exact identity cannot be retained;
do not add another pathname restat or backup-only retry loop.

**Proof and closeout.** Barrier-replace the pathname after diagnosis and before
quarantine; B must remain canonical. Cover cooperating writers, sharing
violations, deletion/recreation, backup collision, and restart. Close together
with reopened `R4-SET-01` and update `Specs/Core/Core_SettingsStore.md`.

### EG-S3-02 — Pin the exact S3 source revision through COPY and MOVE

**Priority / ownership:** P1 / ROUTED to Emberglass
**Effort / repair risk / confidence:** L / high / high

**Evidence and failure.** At committed HEAD,
`Plugins/FileSystemS3/FileSystemS3.Directory.cpp:156-171` and `:1007-1047`
plan key/size without VersionId or ETag. Small copy and multipart copy in
`FileSystemS3.S3.cpp:641-689` and `:912-960` have no source precondition; relay
GET at `FileSystemS3.Directory.cpp:1213` and MOVE deletion at `:2127-2150` are
also unpinned. Replacing same-size A with B can publish B, mix multipart
revisions, or delete B after reporting a MOVE of A. AWS exposes source
preconditions for
[CopyObject](https://docs.aws.amazon.com/AmazonS3/latest/API/API_CopyObject.html)
and
[UploadPartCopy](https://docs.aws.amazon.com/AmazonS3/latest/API/API_UploadPartCopy.html).

**Smallest safe repair.** Carry one `S3ObjectRevision` containing VersionId when
available and ETag otherwise through planning, every server-side copy/part,
relay GET, and deletion. A MOVE that cannot condition exact deletion must retain
the source and report partial failure.

**Proof and closeout.** Replace the source before small copy, between multipart
parts, before relay read, and before delete; cover same/different size and
versioned/unversioned buckets. Record conditional request counts and destination
rollback. The concurrent dirty S3 patch must not close this row until production
AWS calls consume the revision and focused Debug/Release tests pass.

### EG-CURL-01 — Give every returned Curl writer a strong COM owner

**Priority / ownership:** P1 / ROUTED to Emberglass
**Effort / repair risk / confidence:** S / low / high

**Evidence and failure.** `FileSystemCurl.DirectoryOps.cpp:149-319` stores and
later dereferences a raw `FileSystemCurl*` in `TempFileWriter`.
`CurlStreamingWriter` repeats the pattern at `:851-855`, `:1054-1061`, and
`:1208`; `FileSystemCurl.cpp:1497` returns writers independently. A caller can
legally release the filesystem after obtaining a writer, then `Commit()` uses
freed memory. A DLL pin does not keep the COM instance alive.

**Smallest safe repair.** Store `wil::com_ptr<FileSystemCurl>` in both writers
and use normal assignment, not manual AddRef/attach.

**Proof and closeout.** Release factory/filesystem before Write, Commit, cancel,
failure, and writer destruction for both implementations; verify module unload
only after the last writer dies. Update the Curl ownership contract.

### EG-MSD-02 — Validate and bound every Microsoft Drive ranged response

**Priority / ownership:** P1 in this review (raise from Emberglass P2) / ROUTED to Emberglass
**Effort / repair risk / confidence:** M / medium / high

**Evidence and failure.** `FileSystemMicrosoftDrive.cpp:4830-4912` sends a
bounded Range/If-Match request but accepts 200 or 206 without validating
`Content-Range`, start/end/total, exact body length, or representation length.
The response record at `:100-108` carries no range metadata, and transport
collection at `:2190-2222` has no corresponding response-body cap. A shifted
206 is exposed at the requested offset, corrupting copies; an ignored Range at
offset zero can allocate the complete object.

**Smallest safe repair.** Capture `Content-Range` and `Content-Length` at the
shared HTTP boundary. Require an exact compatible start/end/total and body size
for 206. Accept 200 only for a complete representation from zero with the
expected total, under the configured cap.

**Proof and closeout.** Correct, missing, malformed, shifted, truncated,
overlong, changed-total, ignored-range, and oversized 200/206 fixtures; no byte
may be exposed before validation. Reconcile Emberglass's priority count when
updating its owner row.

### EG-ABI-01 — Validate packed `FileInfo` records before the first field read

**Priority / ownership:** P2 / ROUTED to Emberglass
**Effort / repair risk / confidence:** S / low / high

**Evidence and failure.** `Common/PlugInterfaces/FileSystem.h:883-932` obtains a
plugin-owned pointer, record count, and used byte size. Its traversal at
`:962-975` dereferences `entry->FileNameSize` before proving that even
`offsetof(FileInfo, FileName)` remains. Later checks at `:995-1016` do not prove
minimum record length, alignment, or declared-name containment, and the second
pass copies the name at `:1088-1092`. A malformed or buggy plugin can cause an
out-of-bounds host read.

**Smallest safe repair.** Put one structural packed-record iterator/validator in
the lowest Common dependency layer. Before any field/payload read require the
fixed header, even/in-bounds name size, aligned nonterminal offset covering the
record, in-bounds terminal payload, count/termination agreement, and checked
arithmetic. Keep caller policy such as skip/warn outside the validator; do not
create a fourth traversal variant or change ABI layout.

**Proof and closeout.** Malformed corpus covering truncated first/later headers,
odd/oversized names, small/misaligned/overflowing/overlapping offsets, terminal
overrun, early terminal/count mismatch, plus valid single/multi records. Update
the shared-helper catalog and virtual-filesystem consumer contract.

### FG-P0-3-FALLBACK-SIZE-ZERO — Fail closed when remote upload size cannot be proved

**Priority / ownership:** P1 / ROUTED to Floodgate; reopen the existing proof-depth row
**Effort / repair risk / confidence:** S / medium / high

**Evidence and failure.** The real short-success fixture at
`FileSystemCurl.SelfTest.Ftp.cpp:1026` proves a server can consume all client
input while retaining only a prefix. Production accepts a successful probe with
`sizeKnown == false` at `FileSystemCurl.CopyMove.cpp:1236`, or a size-less
listing at `:1260-1273`, before promoting the staged object.
`FileSystemCurl.Shared.cpp:2742` explicitly permits FTP dialects without usable
`SIZE`. A truncated stage can be promoted and MOVE can delete the intact source.

**Smallest safe repair.** Require exact post-upload remote byte-count proof
before publication. If neither targeted stat nor listing can prove the size,
delete the staged sibling and fail closed. Do not infer success from bytes read
from the source or accepted by the transport.

**Proof and closeout.** Extend the loopback server so STOR reports success but
retains a prefix, `SIZE` is rejected, and LIST omits size. Assert destination and
source remain unchanged for COPY/MOVE, cleanup converges, and a normal
size-capable dialect still succeeds.

### EG-SEC-01 — Keep Winget event data out of generated PowerShell

**Priority / ownership:** P1 / ROUTED to Emberglass
**Effort / repair risk / confidence:** S / low / high

**Evidence and failure.** `.github/workflows/winget-release.yml:31-40`
interpolates `github.event.inputs.version` and `release.tag_name` directly into
double-quoted PowerShell before validation. Later steps expose `WINGET_TOKEN`.
GitHub documents direct expression substitution into inline scripts as a script
injection risk:
[Script injections](https://docs.github.com/en/actions/concepts/security/script-injections).
A crafted value can execute early and plant environment/PATH state for a later
credential-bearing step.

**Smallest safe repair.** Pass event values through step `env`, read the raw
environment value, and validate it before deriving outputs. Keep validation in
a credential-free job/step and never generate shell source from event text.

**Proof and closeout.** Workflow/parser fixture with quotes, semicolons,
backticks, subexpressions, newlines, Unicode, and empty/overlong versions; input
must remain inert and no secret-bearing step may run on rejection.

### EG-SEC-02 — Scope repository write credentials to the exact publisher

**Priority / ownership:** P1 / ROUTED to Emberglass
**Effort / repair risk / confidence:** S / medium / high

**Evidence and failure.** `.github/workflows/ci.yml:9-10` grants
`contents: write` workflow-wide and ordinary checkouts retain credentials;
`.github/workflows/release.yml:22` similarly grants write to build/package
jobs. Repository-controlled build scripts and dependency tooling therefore run
with unnecessary write-capable credentials.

**Smallest safe repair.** Default every workflow to `contents: read`, set
`persist-credentials: false` on nonpublisher checkouts, and grant write only to
the exact publish job/step. Coordinate the format auto-commit race with its
existing plan rather than widening this package.

**Proof and closeout.** Parse every workflow and assert the permission/checkout
matrix; run representative PR, branch, tag, and release paths with least
privilege. A build job must be unable to push.

### EG-SET-01 — Preserve source existence across early settings read failure

**Priority / ownership:** P2 / ROUTED to Emberglass
**Effort / repair risk / confidence:** S / medium / high

**Evidence and failure.** `SettingsStore.cpp:377-448` clears the optional stamp,
then size/query/read/short-read failures return before a stamp is captured.
Recovery at `:5525-5572` treats stamp presence as source presence and can load
defaults with ordinary-save permission. An oversized, unreadable, or short-read
existing file is therefore conflated with Missing; later CAS/save cannot
converge safely.

**Smallest safe repair.** Model `Missing`, `PresentStamped`, and
`PresentUnstampable` explicitly. Capture existence immediately after open; any
present or unknown source blocks automatic replacement until recovery is
explicitly resolved.

**Proof and closeout.** Inject size-query, oversize, read, short-read, and stamp
failures plus truly missing input; assert no existing file is overwritten and
save-block state is user-visible/recoverable.

### EG-UI-01 — Build hosted focus/UIA order from visual order

**Priority / ownership:** P2 / ROUTED to Emberglass
**Effort / repair risk / confidence:** S / low / high

**Evidence and failure.** `FolderWindow.FileOperations.Popup.cpp:5341` appends
footer buttons before task controls. Queued-row controls at `:6323-6359` are
appended Collapse, Move Down, Move Up despite visible left-to-right Move Up,
Move Down, Collapse. `DxUi.WindowHost.cpp:982` and
`DxUi.Accessibility.cpp:748` consume child order for Tab and semantic traversal.

**Smallest safe repair.** Produce one stable visual descriptor order:
top-to-bottom, then flow-direction x, and create hosted/UIA children from it.
Do not add a second hard-coded tab-ID list.

**Proof and closeout.** Exact Tab/Shift+Tab/UIA sequence for mixed task states,
dynamic insertion/removal, LTR, and RTL.

### EG-UI-02 — Prune File Operations accessibility announcement state

**Priority / ownership:** P2 / ROUTED to Emberglass
**Effort / repair risk / confidence:** S / low / high

**Evidence and failure.** `FolderWindow.FileOperations.Popup.cpp:5046` inserts
every task into `_announcedTaskStatuses`; the live-task cleanup at `:2458-2495`
prunes collapse/animation state but not this map, declared at
`FolderWindow.FileOperations.Popup.h:682`. A long-lived popup retains one entry
per task ever shown.

**Smallest safe repair.** Prune the map in the existing live-task cleanup pass.

**Proof and closeout.** Cycle thousands of completed/removed tasks and assert
map size never exceeds live task count; announcements for reused/new identities
remain correct.

### EG-DOC-01 — Repair changed review routing and selftest guidance

**Priority / ownership:** P2 / ROUTED to Emberglass
**Effort / repair risk / confidence:** S / low / high

**Evidence and failure.** Frozen `README.md:176-178` labels all selftests
Debug-only, while `:256-258` correctly documents opt-in test-facing Release via
`RSBuildEnableTests=true`. `Specs/Reviews/README.md:24` routes Curl short-success
work to nonexistent
`Specs/Plans/WIP/Operation_Curl_Short2xxStagedUploadProof_2026-07-21.md`; the
historical plan is under Done. Contributors can follow a broken owner or infer
that Release selftest proof is impossible.

**Smallest safe repair.** Rename the README section to “Debug by default;
opt-in Release,” keep production Release exclusion explicit, and route the
review entry to the completed historical plan plus the live Floodgate follow-up.
Add one small path checker that covers both Markdown links and reviewed
backticked repository paths.

**Proof and closeout.** Broken-path fixture, real review/WIP/Done routing scan,
documented Debug command, opt-in Release test build, and production Release
without test hooks. Keep `Specs/Testing/Testing_SelfTests.md` consistent.

### R4-UI-01 — Use one animated geometry for overlay paint and input

**Priority / ownership:** P2 / ROUTED to Emberglass
**Effort / repair risk / confidence:** S-M / low / high

**Evidence and failure.** `ThemeCycleOverlayWindow.cpp:326-391` paints using
`_surfaceScale` and `_surfaceTranslateYDip`; retained mouse and Win32 hit tests
at `:420-469` and `:1437-1476` use final `_surfaceRect`. During animation,
transparent final-rect pixels can capture input while painted shifted pixels can
click through.

**Smallest safe repair.** Centralize current transformed rounded-surface
geometry and inverse-point mapping for paint, hit test, cursor, and routing—or
remove the interactive transform.

**Proof and closeout.** Edge/corner/center points at appear/disappear start,
midpoint, and end, under multiple DPI values and RTL where relevant.

### R4-S3-02 — Give multipart-abort retry one self-converging scheduler

**Priority / ownership:** P1 / ROUTED to Floodgate for implementation; Emberglass is the review ledger
**Effort / repair risk / confidence:** M / medium / high

**Evidence and failure.** At committed HEAD,
`FileSystemS3.IO.cpp:1048-1247` stores `_nextAttempt`; `Schedule()` declines work
before that time, and a failed pass sets a later retry without arming a timer or
new worker. `WaitUntilDrained` polls Schedule in tests, while production is
woken mainly by shutdown/unload checks. A transient abort can remain queued
indefinitely.

**Smallest safe repair.** One tracked threadpool timer/work state machine owns
the queue and rearms itself for the next due attempt. Transfer the plugin module
pin at callback entry and keep one bounded shutdown policy; do not poll from a
detached thread.

**Proof and closeout.** First abort fails and second succeeds without a test
poll, unload query, factory creation, or shutdown trigger; cover terminal 404,
queue cap, concurrent enqueue, shutdown during delay, and unload quiet point.

### R4-TEST-02 — Preserve optional writer capabilities through the bridge decorator

**Priority / ownership:** P2 / ROUTED to Floodgate for implementation; Emberglass is the review ledger
**Effort / repair risk / confidence:** S / low / high

**Evidence and failure.** The test decorator at
`FolderWindow.FileOperations.State.cpp:505` exposes only `IFileWriter` and the
commit-size proof, and wraps all test destination writers at `:692-729`.
Production queries `IFileWriterExpectedSize` at `:9349-9357`; Microsoft Drive
implements it at `FileSystemMicrosoftDrive.cpp:4928`. Tests silently take a
different upload-planning path.

**Smallest safe repair.** Transparently query/advertise/forward
`IFileWriterExpectedSize` only when the inner writer supports it, preserving COM
identity and call order. Do not add provider knowledge to the decorator.

**Proof and closeout.** Supporting and nonsupporting inner writers,
QueryInterface identity, SetExpectedSize before first Write, exact forwarding,
and the production bridge size route.

### J21-MSD-01 — Make Microsoft Drive seek arithmetic overflow-safe

**Priority / ownership:** P2 / OWNED here; coordinate Floodgate
**Effort / repair risk / confidence:** S / low / high

**Evidence and failure.** `FileSystemMicrosoftDrive.cpp:4707-4734` converts a
base to signed form and evaluates `signedBase + offset` without a checked add.
For example, seeking to 1 then adding `INT64_MAX` invokes signed-overflow
undefined behavior or wraps to an invalid position.

**Smallest safe repair.** Use one checked signed/unsigned addition helper and
return the repository's explicit overflow/invalid-seek HRESULT before mutating
position.

**Proof and closeout.** `INT64_MIN`, `INT64_MAX`, base above `INT64_MAX`, exact
boundary successes, one-byte overflows, and all seek origins. Position must be
unchanged on failure.

### EG-SEC-03 — Pin and verify WingetCreate before exposing credentials

**Priority / ownership:** P2 / ROUTED to Emberglass
**Effort / repair risk / confidence:** S / low / high

**Evidence and failure.** `winget-release.yml:98-103` installs the latest
`Microsoft.WingetCreate` without a version constraint; `:209-248` later executes
it with `WINGET_TOKEN`. Publication behavior and the credential-bearing binary
can change without a repository review.

**Smallest safe repair.** Pin a reviewed version/package identity, verify the
resolved executable and version before the secret-bearing step, and prefer an
environment-backed credential if supported.

**Proof and closeout.** Wrong/missing version and executable-path substitutions
must fail before secret exposure; capture final argv and ensure no token is
printed or persisted.

### EG-CI-01 — Invoke Winget publication directly after release publication

**Priority / ownership:** P2 / ROUTED to Emberglass
**Effort / repair risk / confidence:** S-M / medium / high

**Evidence and failure.** `release.yml:399` creates a release with the ordinary
`github.token`; `winget-release.yml:3` depends on the resulting
`release: published` event. GitHub documents that events created by
`GITHUB_TOKEN` generally do not create new workflow runs:
[GITHUB_TOKEN event behavior](https://docs.github.com/en/actions/concepts/security/github_token).
A successful release can silently skip Winget publication.

**Smallest safe repair.** Call a reusable Winget workflow or explicit dependent
job after release creation. Do not use a broader PAT merely to make event
indirection fire.

**Proof and closeout.** A workflow integration test must trace the real release
path into Winget validation/submission, including retry and failure reporting.

### EG-CI-02 — Publish a complete release asset set atomically

**Priority / ownership:** P2 / ROUTED to Emberglass
**Effort / repair risk / confidence:** S / medium / high

**Evidence and failure.** `release.yml:368` validates the full requested asset
matrix, but `:399-416` publishes ZIP/checksum assets first and MSIX in a second
action. Failure of the second action leaves a public release whose checksum set
describes a missing asset.

**Smallest safe repair.** Use one publication invocation for the complete
matrix, or keep the release draft until every upload and final matrix check
succeeds.

**Proof and closeout.** Inject failure on the last asset: no incomplete release
may be public, and rerun must converge without duplicate assets.

### R4-CI-01 — Make Winget PR submission resumable and serialized

**Priority / ownership:** P2 / ROUTED to Emberglass
**Effort / repair risk / confidence:** S-M / medium / high

**Evidence and failure.** `winget-release.yml:231-248` aborts on any matching
open PR. A failure after PR creation but before URL parsing/update poisons
ordinary reruns, and no package/version concurrency group prevents two runs
from racing after both observe no PR.

**Smallest safe repair.** Add package/version concurrency and a three-state
operation: create when absent, resume exactly one automation-owned PR, reject
external or ambiguous matches.

**Proof and closeout.** Crash after creation, malformed tool output, rerun
adoption, two concurrent attempts, an external same-version PR, and duplicate
exact PRs.

### EG-BUILD-01 — Pin and record the official compiler identity

**Priority / ownership:** P2 / ROUTED to Emberglass
**Effort / repair risk / confidence:** M / medium / high

**Evidence and failure.** `build-reusable.yml:38` and `release.yml:144` install
mutable Visual Studio Chocolatey packages. `Directory.Build.props:65`
constrains SDK/toolset families, not exact compiler/component patch versions.
Official binaries can change with no repository diff.

**Smallest safe repair.** Pin the installer/package/channel identity as tightly
as the runner permits and archive `cl /Bv`, MSBuild, Windows SDK, resolved vcpkg
baseline, and image identity with release evidence.

**Proof and closeout.** CI policy rejects floating official toolchains; release
logs and manifest expose the exact resolved identity and a rerun can compare it.

### EG-CURL-02 — Keep Curl delete concurrency scoped per connection

**Priority / ownership:** P3 / ROUTED to Emberglass
**Effort / repair risk / confidence:** M / medium / high

**Evidence and failure.** `FileSystemCurl.CopyMove.cpp:3312-3349` folds every
resolved connection's `effectiveDeleteMaxConcurrency` into one global minimum,
then uses that value as the entire job worker count at `:3386-3393`. Actual
delete/remove calls already use `ConnectionConcurrencyLimiter` keyed by
connection at `:571-620`. One endpoint capped at 1 therefore serializes
unrelated endpoints that can safely overlap.

**Smallest safe repair.** Size the top-level scheduler from the configured
global job cap and task count; leave endpoint safety to the existing keyed
limiter. Preserve one bounded shared budget for recursive/nested work. Do not
add a second limiter or per-directory thread pool.

**Proof and closeout.** Two fixed-latency limiter keys with unequal caps; each
key stays within its cap while the faster key overlaps the slower. Cover mixed
files/directories, cancellation while waiting, first failure,
continue-on-error, and recursive work. Archive per-key high-water and elapsed
before/after Release evidence.

### R4-S3-03 — Reject or plan giant S3 uploads before transferring bytes

**Priority / ownership:** P3 / ROUTED to Floodgate for implementation; Emberglass is the review ledger
**Effort / repair risk / confidence:** M / medium / high

**Evidence and failure.** `FileSystemS3.Internal.h:375` fixes a 64-MiB part;
`FileSystemS3.IO.cpp:1390` does not expose expected-size negotiation and only
rejects at part 10,001 near `:1816-1824`, after roughly 625 GiB has transferred.

**Smallest safe repair.** Accept expected size before the first byte, select one
legal adaptive part size for the whole upload, and reject sizes that cannot meet
S3 part-count/memory constraints. Do not change part size after upload begins.

**Proof and closeout.** Model boundary sizes without allocating them; assert
selected part size/count, memory bound, retry/resume stability, and zero network
activity for rejected sizes.

## 8. Execution order and dependency graph

### Wave 0 — Stabilize ownership and create RED proofs

1. Preserve and separately review the concurrent dirty S3/Terminal-tool diff.
2. Add the deterministic RED seams for `J21-TERM-01`, `-02`, `-03`, `-04`,
   `-05`, `-06`, and `-08`; do not begin architectural extraction first.
3. Reconcile `J21-TERM-07` with Operation Atlas and reopen only the exact
   Terminal closeout row.
4. Add RED proof for `J21-MSD-01`, and route the newly contrary Curl evidence to
   `FG-P0-3-FALLBACK-SIZE-ZERO`; route every existing ID without copying
   ownership.

### Wave 1 — Release blockers and destructive boundaries

1. `J21-TERM-01` ConPTY close.
2. `J21-TERM-04` class retirement/unload.
3. `EG-SET-02`, `EG-S3-02`, `EG-CURL-01`, `EG-MSD-02`, and
   `FG-P0-3-FALLBACK-SIZE-ZERO` data/ownership repairs.
4. `R4-S3-02` abort convergence in its Floodgate implementation owner.
5. `EG-SEC-01` and `EG-SEC-02` workflow trust boundaries.

### Wave 2 — Responsiveness, accessibility, and exact boundaries

1. `J21-TERM-02`, then `J21-TERM-03` with archived performance evidence.
2. `J21-TERM-05` and `J21-TERM-06`.
3. `J21-TERM-08` through `J21-TERM-14`.
4. Remaining P2 storage/UI/workflow packages in their existing owners.

### Wave 3 — Simplification and honest closeout

1. Build behavioral coverage in `J21-TERM-15` before reshaping code.
2. Execute the bounded `WindowProc` extraction in `J21-TERM-16`.
3. Decide the ABI freeze point and execute `J21-TERM-17`.
4. Recover evidence in `J21-TERM-13`, run the exact focused matrix, then run a
   final green Full aggregate for `J21-TERM-07`.
5. Apply the optional CI cache and giant-upload improvements last.

## 9. Reviewed leads that were rejected or not promoted

These checks are recorded to prevent a later review from reopening them without
new evidence:

- The private Terminal runtime loader locks a non-reparse file, validates
  machine/size/SHA-256, resolves the canonical locked path, and uses restricted
  DLL search flags. No current path-swap/load-order defect was found.
- Initial ConPTY launch correctly creates the root suspended, assigns the
  kill-on-close job, then resumes. The defect is later close ordering and the
  reader-start rollback hole, not initial containment.
- WSL launch/insertion uses strict distro/path checks and argument quoting;
  insertion never appends Enter. No command-injection finding was established.
- OSC52 uses bounded opaque posted-payload ownership, initializes/drains the
  payload window, validates UTF-8 on the UI thread, and scrubs retained text.
- Terminal has no detached worker and correctly transfers threadpool module
  pins at callback entry.
- A possible close stall behind the synchronous input writer was not promoted:
  closing ConPTY may unblock it. Retain a saturated-pipe runtime case in the
  close test matrix before claiming a second deadlock.
- Integration-listener readiness before process resume was not promoted; the
  bootstrap retry/OnIdle path provides recovery and no deterministic failure
  was shown.
- `GetViewState` can tag a later locked snapshot with an earlier generation,
  but no production host consumer currently makes decisions from that
  generation. Reconsider when a consumer appears; do not complicate v1 now.
- UIA selection/caret mutation and range geometry are explicit passive-v1
  non-goals. `J21-TERM-05` repairs required read-only unit navigation and
  `J21-TERM-06` exposes security-prompt semantics; neither package authorizes a
  full editable-text or geometry provider.
- Full-grid per-cell rendering may be expensive, but no separate measured
  regression was established. `J21-TERM-02/-03` own the missing output/render
  instrumentation needed to decide.
- Microsoft Drive mutation-time overwrite identity is fixed; the stable item
  identity is revalidated at mutation. Do not reopen `EG-MSD-01` without a new
  interleaving.
- S3 destination no-overwrite is fixed with conditional publication. It is
  distinct from source revision pinning in `EG-S3-02`.
- Cached Keep Both eligibility is fixed by the centralized predicate and focused
  sequential/parallel COPY/MOVE proof. Do not reopen `EG-FOPS-01`.
- SearchService predictable global-event ownership is fixed by the store-scoped
  exclusive lease and its contention/crash tests.
- Known-size Curl short-success is detected when exact remote size is available.
  `FG-P0-3-FALLBACK-SIZE-ZERO` owns the contrary unknown-size proof gap.
- S3 abort cleanup retains the module; a permanent module-unload UAF rationale
  was rejected. Only self-scheduling retry convergence remains in `R4-S3-02`.
- Theme-overlay collapsed Status-root/UIA navigation was internally consistent;
  only transformed hit geometry remains.
- File Operations accessibility snapshot rebuild fanout is a plausible
  performance lead, but no measurement established a material independent
  defect.
- The checkout action update is pinned to an immutable commit SHA.
- Imported Ghostty candidate-patch formatting is upstream/review material, not a
  first-party runtime style defect.

## 10. Verification matrix

| Package group | Minimum deterministic proof | Required closeout |
| --- | --- | --- |
| Terminal close/unload (`01`, `04`, `08`) | Reader cancellation race, uncooperative child, partial start, retained provider, unload/reload | Focused Debug/Release, repeated lifecycle stress, final green Full |
| Terminal output/render (`02`, `03`, `09`) | Output storm, post failure, maximum Kitty image, blocked input, close races | Metrics and archived before/after Release evidence plus Full |
| Terminal accessibility/security (`05`, `06`, `14`) | UIA unit conformance, prompt discovery/invoke, contrast/theme matrix, satellite smoke | Direct provider/UI tests, accessibility review, Full |
| Terminal boundary/dependency (`10`-`13`, `17`) | Hostile CWD/PATH/ABI inputs, exact export set, evidence links, ABI layout | x64/ARM64 focused builds/contracts and fresh-clone checks |
| Storage data integrity and packed ABI | Deterministic source/destination mutation, short/shifted response, unknown-size upload, seek overflow, malformed packed records | Backend focused Debug/Release, bridge COPY/MOVE, PluginContractTests, Full; perf archive where hot-path behavior changes |
| Settings | Injected read/stamp failures and external replacement barrier | Focused process/cross-process cases, authoritative spec update, Full |
| Workflows/security | Parsed permissions, inert metacharacter corpus, pinned tool identity, publication failure injection | Real dry-run or accepted GitHub test path; no secret in logs/argv |
| UI ordering/retention | Exact LTR/RTL focus/UIA order, thousands of task lifecycles, animated edge hit tests | Focused DxUi/Commands coverage and Full |
| Documentation/routing | Broken link/backticked-path fixture plus Debug, opt-in Release, and production Release command checks | Route checker green and README/Testing contracts agree |

## 11. Final definition of done

This plan may move to `Specs/Plans/Done/` only when all of the following are
true:

1. Every OWNED package is closed with a linked RED-before/GREEN-after proof, or
   is transferred to one named live owner with no duplicate queue.
2. Every ROUTED package has its status/evidence updated in its named owner,
   Emberglass and/or Floodgate; this plan does not independently claim those
   fixes.
3. The concurrent worktree changes have a stabilized diff review and do not
   remain an unverified implicit repair.
4. P0/P1 changes have focused x64 Debug and Release proof, relevant ARM64 build
   proof, lifecycle/unload stress where applicable, and a final green Full
   aggregate on the exact closing commit.
5. Responsiveness-sensitive work includes scenario definitions, counters,
   deterministic selftests, and curated before/after Release evidence under
   `Specs/TestRuns/`.
6. Durable behavior and validation rules are merged into the authoritative
   Terminal, Testing, Settings, FileSystem, UI, Installer, and release/workflow
   specs. Normative requirements do not remain only in WIP.
7. Every authoritative evidence link resolves from a fresh clone; no cited run
   is local-only unless labeled as such.
8. The final source remains simpler: one Terminal owner, one close state
   machine, one output notifier, one input drain, shared UIA boundaries, routed
   WndProc cases, and no speculative framework introduced to close a local bug.

## 12. Audited disposition and closeout — 2026-08-04

Every finding was checked against the cited source and surrounding contract. All were
accurate at the reviewed revision except `EG-S3-02`, whose exact source-revision repair
was already present at the starting `HEAD` and was revalidated rather than reimplemented.
The safety, correctness, ownership, accessibility-text, and workflow trust-boundary
findings were repaired in this change. Work requiring a larger product/architecture
decision or official Release/performance evidence was transferred without weakening its
acceptance criteria to the sole residual owner
`../WIP/CodeReview_July21ToAugust4_ResidualDecisionsAndEvidence_2026-08-04.md`.

| Disposition | Finding IDs |
|---|---|
| Repaired and covered here | `J21-TERM-01`, `J21-TERM-02`, `J21-TERM-04`, `J21-TERM-05`, `J21-TERM-08`, `J21-TERM-09`, `J21-TERM-10`, `J21-TERM-11`, `J21-TERM-12`, `EG-SET-01`, `EG-SET-02`, `EG-CURL-01`, `EG-CURL-02`, `EG-MSD-02`, `EG-ABI-01`, `FG-P0-3-FALLBACK-SIZE-ZERO`, `EG-SEC-01`, `EG-SEC-02`, `EG-SEC-03`, `EG-CI-01`, `EG-CI-02`, `EG-BUILD-01`, `EG-UI-01`, `EG-UI-02`, `EG-DOC-01`, `R4-UI-01`, `R4-S3-02`, `R4-TEST-02`, `J21-MSD-01`, `R4-S3-03` |
| Revalidated as already repaired at starting `HEAD` | `EG-S3-02` |
| Repaired core defect; bounded remainder transferred | `J21-TERM-06` (theming/localized UIA snapshot fixed; richer UIA actions require a product decision), `J21-TERM-18` (validated cache implemented; official cold/warm evidence remains) |
| Transferred to the residual owner | `J21-TERM-03`, `J21-TERM-13`, `J21-TERM-14`, `J21-TERM-15`, `J21-TERM-16`, `J21-TERM-17`, `R4-CI-01` adoption/resume remainder |
| Closeout aggregate | `J21-TERM-07` closed by the green canonical Full run recorded below |

The durable behavior is merged into the Terminal, SettingsStore, shared-helper, Virtual
File System, Winget, build-toolchain, and test-coverage specifications. Emberglass and
Floodgate record the completed routed rows; the WIP index names the residual file as the
only unfinished owner.

### Verification record

- Clean Debug x64 solution build: `build.ps1 -Clean -Configuration Debug`, 0 warnings,
  0 errors; log `.build/logs/msbuild-20260804_155723_827.log`.
- Terminal dependency/runtime, plugin source, evidence, network-isolation, workflow,
  and exact-export focused contracts: 207 passed, 0 failed, 5 intentionally skipped.
- Test-harness source contracts: 154 passed, 0 failed.
- Packed `FileInfo` source contracts: 3 passed, 0 failed.
- Deep SettingsStore sandbox suite, including extended-length retained-handle recovery:
  green.
- Final Full-gate Debug x64 build: 0 warnings, 0 errors; log
  `.build/logs/msbuild-20260805_014645_384.log`.
- Seven focused SearchService regressions covering long-path writer ownership,
  SQLite currentness classification, ProgramData roots, warmup failure status, live
  progress, and impersonation warnings: 7 passed, 0 failed.
- SearchService repeated ownership/currentness stress after aggregate triage: 20 passed,
  0 failed; run `20260804T232858Z-9640-4af3b42633f9496d8551db53be5724c1`.
- The four focus/navigation cases implicated by broad-order churn passed together before
  the desktop was locked; run
  `20260804T214641Z-15604-41374af4648e41a7a73edc4400da583a`.
  With Windows `LockApp` subsequently owning the foreground and the OS denying
  `OpenClipboard` with error 5, the 16 affected Commands cases reported explicit
  environmental skips rather than false failures; focused run
  `20260804T234052Z-50764-cb986283e9244f2d8c5fb0ca19b3629f`.
- Archived Monitor burst proof under
  `Specs/TestRuns/7d3a1247382a/Monitor/2026-08-04_235118/`: 60/60 events drained;
  append-to-visible p95 26,193 us, batch-drain p95 1,941 us, frame-total p95
  21,748 us, and present p95 12,566 us.
- `PerformanceTests2`: 14 passed, 0 failed, including the large-folder and plugin
  lifecycle scenarios.
- Canonical Full aggregate: **PASSED**, 1,232 total, 1,153 passed, 0 failed,
  79 explicit skips; zero flaky, regression, isolation-suspect, unclassified, or
  quarantine results; clean TestSandbox disk audit. Run
  `20260804T234641Z-49108-eea3978802d6421580a1f73c78b22d8b`, archived locally at
  `.build/TestSandbox/runs/20260804T234641Z-49108-eea3978802d6421580a1f73c78b22d8b/artifacts/selftest/last_run/run-all-tests-results.json`.

No performance claim without the required same-machine Release archive is closed here;
those evidence rows remain explicit in the residual owner. This bounded transfer is the
user-authorized decision path and does not make the archived review a live parallel plan.
