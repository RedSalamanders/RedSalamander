# Advisor Plan 019 - ViewerWeb async-load snapshot

> **NON-NORMATIVE COMPLETE RECORD.** Archived from root `plans/019-viewerweb-async-load-data-race.md` on 2026-08-25. **No remaining action.** Do not execute this file. Select current work only through `Specs/Plans/WIP/README.md`.

## Archive status

- **State:** COMPLETE
- **Archived:** 2026-08-25
- **Original path:** `plans/019-viewerweb-async-load-data-race.md`
- **Live owner:** Specs/Plans/Done/Operation_Farsight_ViewerPluginsRemediation_ExportSafetyDecodeReliabilityWebSecurityAndTextGeometry_2026-06-16.md
- **No remaining action:** yes

> **Frozen history: do not resume.** Every executor instruction, checkbox, STOP condition, and "next action" below is 2026-06 archive text. **No remaining action.** If work continues, use the live owner named in Archive status.

---
# Plan 019: ViewerWeb — snapshot member state on the UI thread before the async-load worker reads it (fix data race / COM refcount corruption)

> **Frozen historical instructions (do not execute)**: Follow step by step; run every verification before continuing.
> Historical only. Root plans/ was deleted; do not execute.
>
> **Drift check (run first)**: `git diff --stat b274022d9..HEAD -- Plugins/ViewerWeb/ViewerWeb.cpp Plugins/ViewerWeb/ViewerWeb.h`
> If changed, compare excerpts to live code first; mismatch ⇒ STOP.

## Status

- **Priority**: P1 (UAF / double-free of the IFileSystem plugin under a common user sequence)
- **Effort**: S–M
- **Risk**: LOW (mirrors the snapshot pattern the other six viewers already use)
- **Depends on**: none
- **Category**: bug (concurrency / lifetime)
- **Planned at**: commit `b274022d9`, 2026-06-16
- **Reviewed at**: 2026-06-21 — drift check **clean** (`ViewerWeb.{cpp,h}` byte-identical to
  `b274022d9`); all original excerpts verified accurate; **one correction folded in**: the worker
  also reads `self->_metaId` (`:4040`) + `self->_metaName` (`:4046`) — now snapshotted too (see the
  review-correction note under "Current state").

## Why this matters

Unlike the other six viewers (which snapshot everything the worker needs **on the UI thread**
before submitting), `ViewerWeb::StartAsyncLoad` captures only `{viewer, hwnd, requestId, path}`
and the worker reads live members off-thread. Most dangerously,
`wil::com_ptr<IFileSystem> fileSystem = self->_fileSystem;` copies a COM pointer on a threadpool
thread while the UI thread can be executing `_fileSystem = context->fileSystem;` in `Open()`
(Release-old + store + AddRef-new, non-atomic). If the user navigates to another file or refreshes
while a previous (slow/remote) load is still in flight, the worker's `AddRef` races the UI thread's
`Release` on the same object — **a use-after-free / double-free of the active filesystem plugin and
a crash**. The same lambda also copies `_config`, `_theme`/`_hasTheme`, and `_markdownShowSource`
non-atomically while `SetConfiguration`/`SetTheme` mutate them (torn reads → wrong rendering). The
`_openRequestId` guard only filters *delivery* on the UI thread; it does not prevent the concurrent
*reads* the worker already did.

## Current state

```cpp
// ViewerWeb.cpp:2321  (Open(), UI thread) — reassigns the COM pointer:
_fileSystem = context->fileSystem;
```

```cpp
// ViewerWeb.cpp:3900-3977  (StartAsyncLoad, UI thread) — captures only path/hwnd/requestId:
payload->viewer    = this;
payload->hwnd      = hwnd;
payload->requestId = requestId;
payload->path      = path;
// (no snapshot of _fileSystem/_config/_theme/_hasTheme/_markdownShowSource)
```

```cpp
// ViewerWeb.cpp:3979-3996  (AsyncLoadProc, THREADPOOL thread) — reads live members:
ViewerWeb* self  = result->viewer;
auto releaseSelf = wil::scope_exit([&] { self->Release(); });
const ViewerWebKind kind      = self->_kind;
const ViewerWebConfig config  = self->_config;          // torn read race
const bool hasTheme           = self->_hasTheme;        // torn read race
const ViewerTheme theme       = self->_theme;           // torn read race (~25-field POD)
const bool markdownShowSource = self->_markdownShowSource;
wil::com_ptr<IFileSystem> fileSystem = self->_fileSystem;   // <-- COM refcount RACE (UAF/double-free)
```

```cpp
// ViewerWeb.cpp:4040 and :4046  (AsyncLoadProc, same worker) — TWO MORE live reads:
... hasTheme ? ResolveAccentColor(theme, result->path.empty() ? self->_metaId : ...) ...   // :4040
const std::wstring titleW = leafW.empty() ? self->_metaName                                  // :4046
                                          : std::format(L"{} - {}", leafW, self->_metaName);
```

> **Review correction (2026-06-21):** the original plan listed only the six reads above. The
> worker *also* dereferences `self->_metaId` (`:4040`) and `self->_metaName` (`:4046`, twice).
> These two are `std::wstring` set **only in the ViewerWeb constructor** (from the fixed `_kind`,
> `ViewerWeb.cpp:1988-2008`) and never reassigned, so they are **not a live data race**. But they
> are still live `self->_member` reads on the worker thread: leaving them in place would (a) fail
> this plan's own done-criterion #1 grep and (b) violate the stated end-state ("workers read only a
> UI-thread snapshot, never live members"). **Therefore snapshot them too** — they are cheap
> `std::wstring` copies. After the fix, `grep "self->_"` inside `AsyncLoadProc` returns nothing
> (the only remaining `self->` use is `self->Release()`).

```cpp
// ViewerWeb.h:94-107  (AsyncLoadResult — does NOT carry a snapshot today)
struct AsyncLoadResult
{
    ViewerWeb* viewer  = nullptr;
    HWND hwnd          = nullptr;
    uint64_t requestId = 0;
    HRESULT hr         = E_FAIL;
    std::wstring path;
    std::wstring title;
    std::string utf8;
    std::wstring statusMessage;
    std::optional<std::filesystem::path> extractedWin32Path;
    bool jsonExpandCollapseAvailable = false;
    bool offerTextViewerFallback     = false;
};
```

**Exemplar to follow** — ViewerText snapshots on the UI thread before submitting:
`ViewerText.cpp:4959-4965` snapshots config fields and `wil::com_ptr<IFileSystem> fileSystem = _fileSystem;`,
then captures them **by value** into the work lambda (`:4982-4996`). Do the same for ViewerWeb.

## Commands you will need

| Purpose | Command | Expected |
|---|---|---|
| Build plugin | `.\build.ps1 -ProjectName ViewerWeb` | exit 0, 0 warnings/errors |
| Full build | `.\build.ps1` | exit 0 |
| Tests | `.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild` | all pass |

## Scope

**In scope**: `Plugins/ViewerWeb/ViewerWeb.cpp`, `Plugins/ViewerWeb/ViewerWeb.h`
**Out of scope**: introducing locks (not needed — the fix is snapshot-by-value); the script
injection fix (plan 018); WebView2 init lifecycle.

## Git workflow

- Branch: `advisor/019-viewerweb-async-load-race`
- Message: `fix(ViewerWeb): snapshot fileSystem/config/theme on UI thread for async load`
- Do NOT push/PR unless instructed.

## Steps

### Step 1: Add snapshot fields to the work payload/item

Add the snapshot to `AsyncLoadResult` (or, cleaner, to the `AsyncLoadWorkItem` that already wraps
it — pick whichever the worker can read; `AsyncLoadResult` is simplest since `AsyncLoadProc`
receives it). Add:

```cpp
// in AsyncLoadResult (ViewerWeb.h)
ViewerWebKind         kindSnapshot{};
ViewerWebConfig       configSnapshot{};
bool                  hasThemeSnapshot = false;
ViewerTheme           themeSnapshot{};
bool                  markdownShowSourceSnapshot = false;
wil::com_ptr<IFileSystem> fileSystemSnapshot;     // AddRef taken on the UI thread
std::wstring          metaIdSnapshot;             // constructor-immutable, but keep the worker self-contained
std::wstring          metaNameSnapshot;
```

(Confirm `ViewerWebKind`, `ViewerWebConfig`, `ViewerTheme`, `std::wstring`, and `<wil/com.h>` are
all visible in `ViewerWeb.h` — verified 2026-06-21: `ViewerWebKind` enum and `ViewerWebConfig` are
declared in the header, `ViewerTheme` comes via `PlugInterfaces/Viewer.h`, `std::wstring` via
`<string>`, `wil::com_ptr` via `<wil/com.h>`, `IFileSystem` via `PlugInterfaces/FileSystem.h`. No
new include is required. `_theme`/`_hasTheme` are protected members inherited from
`EmbeddedViewerBase<ViewerWeb>` — accessible from `ViewerWeb::AsyncLoadProc`/`StartAsyncLoad`.)

### Step 2: Populate the snapshot in `StartAsyncLoad` (UI thread)

In `StartAsyncLoad`, before `AddRef()`/submit, fill the snapshot from members on the UI thread:

```cpp
payload->kindSnapshot               = _kind;
payload->configSnapshot             = _config;
payload->hasThemeSnapshot           = _hasTheme;
payload->themeSnapshot              = _theme;
payload->markdownShowSourceSnapshot = _markdownShowSource;
payload->fileSystemSnapshot         = _fileSystem;   // AddRef happens here, on the UI thread, safely
payload->metaIdSnapshot             = _metaId;
payload->metaNameSnapshot           = _metaName;
```

The `wil::com_ptr` copy here is safe because it runs on the same (UI) thread that mutates
`_fileSystem` in `Open()` — there is no concurrency on the UI thread.

### Step 3: Make `AsyncLoadProc` read ONLY from the snapshot

In `AsyncLoadProc`, replace every `self->_kind / _config / _hasTheme / _theme /
_markdownShowSource / _fileSystem / _metaId / _metaName` read with the corresponding
`result->*Snapshot` (the two `_metaName` reads at `:4046` and the `_metaId` read at `:4040` are
easy to miss — replace both occurrences). `self` must be used **only** for the final `Release()`
(keep the `releaseSelf` scope_exit). The `requestId` comparison is done on the UI thread in
`OnAsyncLoadComplete` (unchanged). After the edit, the worker must not dereference any
`self->_member` at all.

Confirm: `grep -n "self->_" Plugins/ViewerWeb/ViewerWeb.cpp` returns **no lines inside
`AsyncLoadProc`** (`self->Release()` is `self->Release`, not `self->_`, so it does not match; the
only `self->` left in the worker is that `Release()`). `self->AddRef` is not in the worker.

**Verify**: `.\build.ps1 -ProjectName ViewerWeb` → exit 0, then `.\build.ps1` → exit 0.

## Test plan

- This race needs a slow/remote filesystem to widen the window; deterministically reproducing it
  from the harness is impractical. The high-value verification is **structural**: confirm by
  inspection (and the `grep` above) that the worker reads no mutable member directly.
- Ensure existing ViewerWeb open/navigate cases in `Tests/ViewerPETests/ViewerPETests.cpp` still
  pass (the snapshot must not change rendered output for the single-open case).
- Optional stronger test if the harness supports it: open a document, then immediately call
  `Open()` again with a different file (navigate) before the first load completes, and assert no
  crash and the final rendered file is the second one. Add only if the harness already has a way
  to sequence two opens without waiting for the first to finish.
- Verification: `.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild` → all pass.

## Done criteria

ALL must hold:

- (archived, not live)  `AsyncLoadProc` dereferences no `self->_member` at all; `grep "self->_"` returns nothing
      inside the worker (the only `self->` use is `self->Release()`).
- (archived, not live)  All worker inputs (`kind`, `config`, `hasTheme`, `theme`, `markdownShowSource`, `fileSystem`,
      `metaId`, `metaName`) come from a snapshot populated on the UI thread in `StartAsyncLoad`.
- (archived, not live)  The `fileSystem` snapshot is a `wil::com_ptr` copied on the UI thread (AddRef on UI thread).
- (archived, not live)  No new mutex/lock was introduced (snapshot-by-value is the mechanism).
- (archived, not live)  `.\build.ps1 -ProjectName ViewerWeb` and `.\build.ps1` exit 0, 0 warnings.
- (archived, not live)  `.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild` passes.
- (archived, not live)  Only `ViewerWeb.cpp` and `ViewerWeb.h` modified.
- (archived, not live)  `plans/README.md` updated.

## STOP conditions

- Excerpts don't match live code (drift).
- A worker code path genuinely needs a *live* (not snapshot) view of a member (e.g. it must see a
  config change made after submission) — report; that would change the design intent and is not in
  scope.
- `ViewerTheme`/`ViewerWebConfig` is not trivially copyable in a way that breaks the snapshot —
  report (they should be plain value types).

## Maintenance notes

- All seven viewers should follow the same rule: **workers read only a UI-thread snapshot, never
  live members.** ViewerWeb was the lone exception; after this it matches ViewerText/ViewerImgRaw/
  ViewerPE.
- Reviewer: confirm the `fileSystem` AddRef now happens on the UI thread and that the worker's
  `releaseSelf` still balances the dispatcher's `AddRef` (plan 017 territory — do not double-release).
