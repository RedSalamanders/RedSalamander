# Advisor Plan 017 - Async work-item dispatch hardening

> **NON-NORMATIVE COMPLETE RECORD.** Archived from root `plans/017-async-workitem-dispatch-hardening.md` on 2026-08-25. **No remaining action.** Do not execute this file. Select current work only through `Specs/Plans/WIP/README.md`.

## Archive status

- **State:** COMPLETE
- **Archived:** 2026-08-25
- **Original path:** `plans/017-async-workitem-dispatch-hardening.md`
- **Live owner:** Specs/Plans/Done/Operation_Farsight_ViewerPluginsRemediation_ExportSafetyDecodeReliabilityWebSecurityAndTextGeometry_2026-06-16.md
- **No remaining action:** yes

> **Frozen history: do not resume.** Every executor instruction, checkbox, STOP condition, and "next action" below is 2026-06 archive text. **No remaining action.** If work continues, use the live owner named in Archive status.

---
# Plan 017: Harden async work-item dispatch in ViewerImgRaw + ViewerText (fix COM self-ref leak on submit failure and null-deref on OOM)

> **Frozen historical instructions (do not execute)**: Follow step by step; run every verification before continuing.
> Historical only. Root plans/ was deleted; do not execute.
>
> **Drift check (run first)**: `git diff --stat c3e89e580..HEAD -- Plugins/ViewerImgRaw/ViewerImgRaw.Decode.cpp Plugins/ViewerImgRaw/ViewerImgRaw.Export.cpp Plugins/ViewerText/ViewerText.cpp`
> If any changed, compare excerpts to live code first; mismatch ⇒ STOP.

## Status

- **Priority**: P1 (permanent object/COM leak + hard crash on the OOM path)
- **Effort**: S
- **Risk**: LOW (purely additive guards on already-error paths)
- **Depends on**: none (plans 014/016 — which touched the same `Export.cpp` function — are now
  merged to master, so this lands cleanly on top; no parallel-worktree overlap remains)
- **Category**: bug (RAII / lifetime)
- **Planned at**: commit `c3e89e580`, 2026-06-21 (anchors refreshed during reconcile; finding intact
  — originally planned against `b274022d9`, 2026-06-16, before plans 014/015/016 merged)

## Why this matters

Five async dispatchers across two plugins follow the same `AddRef(); build work item;
TrySubmitThreadpoolCallback(...)` shape, where the balancing `Release()` lives in a
`wil::scope_exit` **inside the work lambda** (so it only runs if the lambda runs). Two defects:

1. **Leak on submit failure**: when `TrySubmitThreadpoolCallback` returns 0 (threadpool
   exhaustion / low memory), the lambda never runs, so the `Release()` never fires — the object's
   refcount never returns to zero and the whole viewer (its D2D device, DWrite formats, file
   readers, decoded-image cache) leaks for the life of the process. In ViewerImgRaw's three
   sites it is **worse**: `ctx.release()` is called unconditionally, so the work-item heap
   allocation (including captured `fileSystem` com_ptr and, for export, the moved-in pixel
   buffer) leaks too.
2. **Null-deref on OOM**: the work item is allocated with `new (std::nothrow)` and then
   immediately dereferenced (`ctx->moduleKeepAlive = ...`) with no null check, so genuine memory
   exhaustion crashes the whole host instead of degrading gracefully.

**The correct pattern already exists** in ViewerWeb and ViewerSqlite. ViewerWeb:

```cpp
// ViewerWeb.cpp:3944-3949 and 3969-3973  (the template)
auto ctx = std::unique_ptr<AsyncLoadWorkItem>(new (std::nothrow) AsyncLoadWorkItem{});
if (! ctx) { Release(); return E_OUTOFMEMORY; }           // null check balances AddRef
// ...
const BOOL queued = TrySubmitThreadpoolCallback(...);
if (queued == 0) { Release(); return E_FAIL; }            // balance AddRef, let ctx dtor free the item
ctx.release();                                            // only on success
```

## Current state — the 5 sites (all with `AddRef();` then unchecked `ctx->` then unconditional/early-return without `Release()`)

1. `ViewerImgRaw.Decode.cpp` — `StartAsyncOpen` (main open):
```cpp
// :2361   AddRef();
// :2373   auto ctx = std::unique_ptr<AsyncOpenWorkItem>(new (std::nothrow) AsyncOpenWorkItem{});
// :2375   ctx->moduleKeepAlive = ...;      // <-- null-deref if ctx == nullptr
// :2817   if (queued == 0) { Debug::Error(L"...async open..."); }
// :2822   ctx.release();                   // <-- unconditional: leaks work item + AddRef on failure
```
2. `ViewerImgRaw.Decode.cpp` — `StartPrefetchNeighbors`:
```cpp
// :2020   AddRef();
// :2031   auto ctx = std::unique_ptr<PrefetchWorkItem>(new (std::nothrow) PrefetchWorkItem{});
// :2033   ctx->moduleKeepAlive = ...;      // <-- null-deref
// :2241   if (queued == 0) { Debug::Error(L"...prefetch..."); }
// :2246   ctx.release();                   // <-- unconditional leak
```
3. `ViewerImgRaw.Export.cpp` — `BeginExport` export worker:
```cpp
// :767    AddRef();
// :778    auto ctx = std::unique_ptr<AsyncExportWorkItem>(new (std::nothrow) AsyncExportWorkItem{});
// :780    ctx->moduleKeepAlive = ...;      // <-- null-deref
// :840    if (queued == 0) { Debug::Error(L"...export..."); }
// :845    ctx.release();                   // <-- unconditional leak
```
4. `ViewerText.cpp` — `StartAsyncOpen`:
```cpp
// :4967   AddRef();
// :4979   auto ctx = std::unique_ptr<AsyncOpenWorkItem>(new (std::nothrow) AsyncOpenWorkItem{});
// :4981   ctx->moduleKeepAlive = ...;      // <-- null-deref
// :5684   if (queued == 0) { Debug::Error(L"...async open..."); return; }   // frees item but LEAKS the AddRef
// :5690   ctx.release();
```

In every case the balancing `Release()` is `auto releaseSelf = wil::scope_exit([&] { Release(); });`
declared **inside** `ctx->work` (e.g. `Decode.cpp:2036` prefetch / `:2389` main-open, `Export.cpp:783`,
`ViewerText.cpp:4998`), so it never runs when the lambda doesn't run.

## Commands you will need

| Purpose | Command | Expected |
|---|---|---|
| Build ViewerImgRaw | `.\build.ps1 -ProjectName ViewerImgRaw` | exit 0, 0 warnings/errors |
| Build ViewerText | `.\build.ps1 -ProjectName ViewerText` | exit 0, 0 warnings/errors |
| Full build | `.\build.ps1` | exit 0 |
| Tests | `.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild` | all pass |

## Scope

**In scope**:
- `Plugins/ViewerImgRaw/ViewerImgRaw.Decode.cpp` (sites 1, 2)
- `Plugins/ViewerImgRaw/ViewerImgRaw.Export.cpp` (site 3)
- `Plugins/ViewerText/ViewerText.cpp` (site 4)

**Out of scope**: ViewerWeb / ViewerSqlite (already correct — use only as reference); the lambda
bodies; everything except the alloc-null-check and the `queued==0` path.

## Git workflow

- Branch: `advisor/017-async-dispatch-hardening`
- Message: `fix(viewers): balance AddRef and free work item on async submit failure / OOM`
- Do NOT push/PR unless instructed.

## Steps

Apply the **same two edits** at each of the 5 sites:

### Step 1: Null-check the work item right after allocation

Immediately after the `auto ctx = std::unique_ptr<...>(new (std::nothrow) ...{});` line and
**before** the first `ctx->...` dereference, insert:

```cpp
if (! ctx)
{
    Release();   // balance the AddRef() above
    return;      // (ViewerImgRaw/ViewerText dispatchers are void; ViewerText::StartAsyncOpen is void too)
}
```

(Confirm each enclosing function's return type — all four/five are `void`, so `return;` is
correct. If any returns `HRESULT`, return `E_OUTOFMEMORY` instead.)

### Step 2: On submit failure, release and return before `ctx.release()`

Change each `queued == 0` block so it balances the `AddRef` and returns, letting the `ctx`
`unique_ptr` destructor free the work item:

```cpp
if (queued == 0)
{
    Debug::Error(L"...");   // keep the existing message
    Release();              // balance the AddRef()
    return;                 // ctx's destructor frees the work item; do NOT call ctx.release()
}
ctx.release();              // reached only on success — threadpool now owns the item
```

For the three ViewerImgRaw sites this means moving `ctx.release()` so it is **after** the
`if (queued == 0) { ... return; }` block (it currently sits unconditionally after it — the added
`return` makes it success-only). For ViewerText, the block already `return`s; just add `Release();`
before that `return`.

**Verify after each file**:
- `.\build.ps1 -ProjectName ViewerImgRaw` → exit 0 (after editing Decode.cpp + Export.cpp)
- `.\build.ps1 -ProjectName ViewerText` → exit 0 (after editing ViewerText.cpp)

### Step 3: Cross-check the invariant

For each site confirm by reading: exactly one `AddRef()` is matched by exactly one `Release()`
on **every** path out of the dispatcher (null-alloc, submit-fail, success-where-lambda-runs).
On the success path the lambda's `releaseSelf` scope_exit still owns the balancing Release — do
NOT add a second Release there.

**Verify**: `.\build.ps1` → exit 0.

## Test plan

- These paths only trigger under allocation/threadpool failure, which is impractical to force
  from the integration harness. The correctness here is verifiable by inspection (the
  AddRef/Release balance table per site) plus a clean build.
- Add no new test unless the harness already has a memory-injection hook; instead, ensure the
  existing viewer open/export cases in `Tests/ViewerPETests/ViewerPETests.cpp` still pass (the
  happy path must be unchanged: `ctx.release()` still runs on success so the threadpool owns the
  item and the lambda's `releaseSelf` still balances the ref).
- Verification: `.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild` → all pass.

## Done criteria

ALL must hold:

- (archived, not live)  All 5 sites have `if (! ctx) { Release(); return; }` immediately after the `nothrow` alloc.
- (archived, not live)  All 5 `queued == 0` blocks call `Release();` and `return;` before any `ctx.release()`.
- (archived, not live)  No site calls `ctx.release()` on a failure path.
- (archived, not live)  `.\build.ps1 -ProjectName ViewerImgRaw` and `-ProjectName ViewerText` each exit 0, 0 warnings.
- (archived, not live)  `.\build.ps1` exits 0.
- (archived, not live)  `.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild` passes (happy-path viewer open/export
      unaffected).
- (archived, not live)  Only the three in-scope files modified.
- (archived, not live)  `plans/README.md` updated.

## STOP conditions

- Excerpts don't match live code (drift).
- Any enclosing dispatcher is not `void` and you're unsure of the correct error return — report.
- You find an additional `AddRef`/work-item dispatcher in these files not listed here — report it
  (it likely needs the same fix) rather than guessing.

## Maintenance notes

- The robust shape is: `AddRef()` → `nothrow` alloc → `if(!ctx){Release();return;}` → build →
  submit → `if(!queued){Release();return;}` → `ctx.release()`. Any new async dispatcher in these
  plugins should copy ViewerWeb's `StartAsyncLoad` (`ViewerWeb.cpp:3900`) verbatim.
- A future cleanup (plan 030) could hoist this into a shared `SubmitViewerWork(this, lambda)`
  helper in Common so the balance is impossible to get wrong; out of scope here.
- Reviewer: trace the refcount on all three exit paths per site.
