# Advisor Plan 005 - De-duplicate plugin Factory.cpp files

> **NON-NORMATIVE COMPLETE RECORD.** Archived from root `plans/005-plugin-factory-dedup.md` on 2026-08-25. **No remaining action.** Do not execute this file. Select current work only through `Specs/Plans/WIP/README.md`.

## Archive status

- **State:** COMPLETE
- **Archived:** 2026-08-25
- **Original path:** `plans/005-plugin-factory-dedup.md`
- **Live owner:** none (implemented as Common/PlugInterfaces/FactoryImpl.h)
- **No remaining action:** yes

> **Frozen history: do not resume.** Every executor instruction, checkbox, STOP condition, and "next action" below is 2026-06 archive text. **No remaining action.** If work continues, use the live owner named in Archive status.

---
# Plan 005: De-duplicate the 14 plugin Factory.cpp files into a shared factory implementation

> **Frozen historical instructions (do not execute)**: This plan is archived; do not follow these steps. Run every
> verification command and confirm the expected result before moving to the
> next step. If anything in the "STOP conditions" section occurs, stop and
> report — do not improvise. Historical: status-row updates no longer apply for this plan
> in `plans/README.md`.
>
> **Drift check (run first)**: `git diff --stat a72512919..HEAD -- Plugins/*/Factory.cpp Common/PlugInterfaces/Factory.h`
> If any in-scope file changed since this plan was written, compare the
> "Current state" excerpts against the live code before proceeding; on a
> mismatch, treat it as a STOP condition.
> NOTE: scope this check to the **Factory.cpp files + Factory.h only** (above) —
> the broader `Plugins/` tree has churned a lot for unrelated reasons (fileops,
> search, cloud safety), but none of that touches the factory entry points this
> plan dedups. A bare `git diff -- Plugins` will show noise; ignore it.

## Status

- **Priority**: P2
- **Effort**: M
- **Risk**: MED (touches every plugin's entry points; mitigated by plan 004's contract test)
- **Depends on**: plans/004-plugin-capabilities-contract-test.md (MUST be green before starting)
- **Category**: tech-debt
- **Planned at**: commit `d6bfccc42`, 2026-06-11 — **drift re-checked at `a72512919`, 2026-06-21** (reconcile): zero drift on the real targets (`git diff d6bfccc42..a72512919 -- Plugins/*/Factory.cpp Common/PlugInterfaces/Factory.h` is empty — all 14 `Factory.cpp` and `Factory.h` are byte-identical to plan time). Still gated on **plan 004** (now DONE/merged at `fe1630d05`), so the safety net is in place.
- **Reviewed**: 2026-06-21 (review-plan pass against `a72512919`). The whole 14-factory inventory was read against source and several load-bearing facts were corrected — most importantly the **empty-`pluginId` semantics differ between single- and multi-plugin DLLs** (the original "empty ⇒ first entry" rule would have silently regressed the 4 multi-id factories from `E_INVALIDARG` to a live instance). The optional `RedSalamanderPluginShutdown`/`…RetainModuleUntilProcessExit` exports and the near-universal `REDSAL_DEFINE_TRACE_PROVIDER` are now catalogued as expected, not anomalies. Done-criteria line gate relaxed (it was unachievable for the multi-id DLLs). **Per maintainer direction, the `const FactoryOptions*` parameter is a kept future-extension hook** (spec-reserved, host already passes a live `&options`); the shared helper therefore *threads it through* into each `createInstance` thunk rather than dropping it — a repo-wide census (17 host call sites, zero current readers) and a design review confirmed this keeps the export ABI byte-identical and makes a future consumer a single-`Factory.cpp` change. See the corrected "Current state" below.
- **Step 6 tightened 2026-06-21 (post-execution):** the first execution ran the repo-wide `.\format-all.ps1`, which reformatted ~140 unrelated files and polluted the worktree. Step 6 now formats ONLY the 15 in-scope files via `clang-format --style=file` and adds a `git status` gate that fails if any other file is dirtied. (The executed branch is unaffected — its 3 commits were already clean; this fixes re-runnability.)

## Why this matters

Each of the 14 plugin DLLs carries a `Factory.cpp` (125–229 lines) whose three exported entry points are near-identical boilerplate: argument validation, IID gating, plugin-id matching, and delegation. When the factory contract changes, all 14 must be edited in lockstep — exactly what commit `6fe57868d` ("Remove legacy GetCapabilities fallback; enforce mandatory contract") had to do. A shared implementation makes the next contract change a one-file edit and removes the drift that already exists between copies (e.g. FileSystemDummy's `CreatePluginInstance` doesn't null-check `result`; ViewerText's does).

## Current state

All facts below were re-verified against source at `a72512919` (review-plan pass). The drift check above must pass before trusting them; if a Factory.cpp changed, re-read it.

**The three exported entry points (declared in `Common\PlugInterfaces\Factory.h`, plugins compile with `#define PLUGFACTORY_EXPORTS`)** — these are the ONLY functions this plan deduplicates. Exact signatures (keep verbatim in every migrated file):
```cpp
HRESULT __stdcall RedSalamanderEnumeratePlugins(REFIID riid, const PluginMetaData** metaData, unsigned int* count);
HRESULT __stdcall RedSalamanderCreate(REFIID riid, const FactoryOptions* factoryOptions, IHost* host, const wchar_t* pluginId, void** result);
HRESULT __stdcall RedSalamanderGetConfigurationSchema(REFIID riid, const wchar_t* pluginId, const char** schemaJsonUtf8);
```
- ⚠️ **`RedSalamanderCreate` takes a `const FactoryOptions*` second parameter** (a struct holding one `DebugLevel` field, `Factory.h:30-33`). It is a **deliberate future-extension hook**, not dead weight — the spec reserves it explicitly (`Specs/Plugins/Plugins_VirtualFileSystem.md:117-118` "the parameter remains part of the ABI for future plugins"; `:1918` "Log detailed diagnostics internally (use DebugLevel from FactoryOptions)"), and the host already constructs a real, addressable `FactoryOptions` and passes `&options` at every shipping call site (currently hardcoded `DEBUG_LEVEL_NONE` — e.g. `FileSystemPluginManager.cpp:1082-1086`, the five `ViewerPluginManager.cpp` create sites, `CompareDirectoriesWindow.cpp:162-171`; `PluginContractTests.cpp:301` is the lone `nullptr` caller). **No code reads it yet** — all 14 factories name it `/*factoryOptions*/` and ignore it. Because it is a kept hook, the shared helper must **thread it through** (Step 3): the export forwards it into `FactoryCreate<TInterface>`, which forwards it to the matched entry's `createInstance` thunk, where it sits named-but-unused — so a future plugin that wants to read it edits ONLY its own `Factory.cpp` thunk, never `FactoryImpl.h` or the other 13 files. The export's byte-for-byte signature is unchanged. Step 1 still re-confirms no factory *reads* it today (a reader would be live behavior the dedup must preserve).

**Single- vs multi-plugin split (this is the critical correctness axis — get it right):**
- **10 single-plugin factories**: `FileSystem`, `FileSystem7z`, `FileSystemDummy`, `FileSystemGoogleDrive`, `ViewerImgRaw`, `ViewerPE`, **`ViewerSpace`**, `ViewerSqlite`, `ViewerText`, `ViewerVLC`. `EnumeratePlugins` returns one `PluginMetaData` with `count == 1`. Empty/null `pluginId` ⇒ **create the only plugin** (the guard is `if (pluginId && pluginId[0] != L'\0' && !EqualsNoCase(pluginId, id)) return ERROR_NOT_FOUND;` — empty falls through to creation). `GetConfigurationSchema` with empty id ⇒ returns the only plugin's schema.
- **4 multi-plugin factories**: `FileSystemCurl` (4 protocol ids), `FileSystemMicrosoftDrive` (3), `FileSystemS3` (2), `ViewerWeb` (web/json/markdown). `EnumeratePlugins` returns a contiguous static array (`std::array<PluginMetaData,N>`) with `count == N`. Empty/null `pluginId` ⇒ **`E_INVALIDARG`** (NOT first-entry) in BOTH `RedSalamanderCreate` and `RedSalamanderGetConfigurationSchema`; an id that matches no entry ⇒ `ERROR_NOT_FOUND`.
- Both kinds: `E_POINTER` on null out-params; out-params zeroed first; `E_NOINTERFACE` if `riid != __uuidof(<the DLL's interface>)`; case-insensitive id match via `OrdinalString::EqualsNoCase` (single-plugin) or an exact-id dispatch (multi-plugin, e.g. ViewerWeb's `KindFromPluginId`).

**Exemplars, read at `a72512919`:**
- `Plugins\FileSystemDummy\Factory.cpp` (125 lines) — single IFileSystem; `CreatePluginInstance(REFIID,IHost*,void**)` does `new (std::nothrow) FileSystemDummy(host)` then `QueryInterface`/`Release`; does NOT re-check `result` for null (relies on the export's `E_POINTER` guard). This is the simplest case.
- `Plugins\ViewerWeb\Factory.cpp` (229 lines) — multi IViewer; metadata is a `std::array<PluginMetaData,3>` built once in a static `LocalizedPluginMetaDataSet`; `CreatePluginInstance(REFIID,IHost*,ViewerWebKind,void**)` takes a **per-entry discriminator** (`ViewerWebKind`), re-checks `result`, and does `new ViewerWeb(kind)` + `instance->SetHost(host)`. Empty id ⇒ `E_INVALIDARG`. Implements the optional shutdown/retain exports (below). This is the hardest case — design the shared helper against it.
- Instance-creation idioms vary across the 14: `ctor(host)` (Dummy), default-ctor + `SetHost(host)` (ViewerText), `ctor(discriminator)` + `SetHost(host)` (ViewerWeb). The shared design must NOT hardcode one — each entry carries its own `createInstance` thunk.

**Two per-DLL elements that are NOT factory logic and MUST be preserved verbatim (do not treat as anomalies):**
- `#define REDSAL_DEFINE_TRACE_PROVIDER` before `#include "Helpers.h"` — defines the DLL's ETW provider; present in **13 of 14** (all except FileSystemDummy). Exactly one TU per DLL must define it; the Factory.cpp is that TU today — leave it there.
- The optional module-quiet-point exports `RedSalamanderPluginShutdown()` / `RedSalamanderPluginRetainModuleUntilProcessExit()` (declared in `Factory.h`, discovered by the host via `GetProcAddress`) — implemented in **3 factories**: `FileSystemS3`, `ViewerSpace`, `ViewerWeb` (e.g. ViewerWeb's `RedSalamanderPluginShutdown` calls `ResetSharedEnvironment()`). These are unrelated to the three deduplicated functions — keep them exactly as-is in their Factory.cpp.

- `OrdinalString::EqualsNoCase` is declared in **`Common\Helpers.h`** (namespace `OrdinalString`). ⚠️ The shared header therefore cannot both "use EqualsNoCase" and "avoid including Helpers.h" — see Step 3 for the resolution (the Factory.cpp TUs already include Helpers.h anyway, for the trace provider and `LoadStringResource`).
- Conventions: C++23, `noexcept` where the existing code uses it, no exceptions across the C ABI. Formatting uses the repo `.clang-format` (`clang-format --style=file`); note `.\format-all.ps1` formats the WHOLE repo, so for this plan format only your changed files — see Step 6.
- Safety net: `PluginContractTests.exe` from plan 004 validates exports, enumeration, schema, capabilities, and bogus-id behavior for every DLL.

## Commands you will need

| Purpose | Command | Expected on success |
|---------|---------|---------------------|
| Build all | `.\build.ps1` | exit 0 |
| Contract test | `.\.build\x64\Debug\PluginContractTests.exe` | exit 0 |
| Full selftests | `$p = Start-Process .\.build\x64\Debug\RedSalamander.exe -ArgumentList '--selftest','--selftest-timeout-multiplier=2.0' -PassThru -Wait; $p.ExitCode` | `0` |
| Format (changed files only — see Step 6) | `clang-format -i --style=file <each in-scope file>` | exit 0; only in-scope files dirtied |

## Scope

**In scope**:
- `Common\PlugInterfaces\FactoryImpl.h` (create — the shared implementation header)
- All 14 `Plugins\*\Factory.cpp` files (reduce the three exports to one-line delegations; keep each file's metadata/schema/creator definitions, its entry array, and any trace-provider macro + optional shutdown exports)

**Out of scope** (do NOT touch):
- `Common\PlugInterfaces\Factory.h` ABI declarations — the exported signatures must remain byte-identical; the HOST side (`RedSalamander\FileSystemPluginManager.cpp`, `ViewerPluginManager.cpp`) must not change at all.
- Plugin class implementations (`FileSystemDummy.h/.cpp` etc.) — only their factory TUs.
- Behavior changes of any kind: identical HRESULTs for identical inputs is the acceptance bar.

## Git workflow

- Branch: `advisor/005-plugin-factory-dedup`
- Commit per phase: "Add shared plugin factory implementation header", "Migrate filesystem plugin factories", "Migrate viewer plugin factories".
- Do NOT push or open a PR unless the operator instructed it.

## Steps

### Step 1: Inventory all 14 factories

Read every `Plugins\*\Factory.cpp`. Build a table (include it in your final report) with one row per file and these columns:
1. plugin name + interface IID (`IFileSystem` or `IViewer`).
2. **single vs multi** (count of enumerated plugins; the "Current state" split says single=10, multi=4 {Curl, MSDrive, S3, Web} — confirm it).
3. **empty/null `pluginId` behavior** in both `RedSalamanderCreate` AND `RedSalamanderGetConfigurationSchema` — record the exact return (single→creates/returns the only entry; multi→`E_INVALIDARG`). This is the column that drives the shared helper's policy; do not skip it.
4. instance-creation idiom: `ctor(host)` / default-ctor+`SetHost` / `ctor(discriminator)+SetHost`, and the discriminator type if any (e.g. `ViewerWebKind`).
5. presence of `REDSAL_DEFINE_TRACE_PROVIDER` (expected in 13/14).
6. presence of the optional exports `RedSalamanderPluginShutdown` / `RedSalamanderPluginRetainModuleUntilProcessExit` (expected in S3/Space/Web), what `Shutdown` does, and the exact `Retain…` return value (it differs by design — S3 returns `TRUE`, ViewerWeb returns `FALSE` — preserve each verbatim).
7. `factoryOptions` usage — confirm each factory currently IGNORES it (named `/*factoryOptions*/`, body never reads it; expected everywhere). The shared design FORWARDS it (so it stays a live future hook), so a factory that merely ignores it is fine; STOP only if a factory actually READS a `factoryOptions` field today — that is live behavior the abstraction must preserve.
8. anything else.

**Verify**: the table covers 14 files and matches the splits in "Current state" (single=10/multi=4, trace-provider=13, optional-exports=3, factoryOptions-used=0). The items in columns 5–7 are **expected and stay in their Factory.cpp verbatim** — they are NOT reasons to stop. STOP and report ONLY if column 8 turns up genuinely novel factory behavior: conditional/stateful registration, environment checks that gate enumeration or creation, a `factoryOptions` field actually being read, or an export beyond the five named in `Factory.h`. Those would mean the abstraction below could erase behavior.

### Step 2: Record the baseline

`.\build.ps1` then run `PluginContractTests.exe` and the full `--selftest` (the Commands-table row). All must pass BEFORE any change. Then capture the contract-test output for later diffing — ⚠️ the test writes `[ OK ]` lines to **stdout** but `[ FAILED ]` lines to **stderr** (`PluginContractTests.cpp:74,78`), so a bare `>` would silently drop failures. Capture **both streams** (PowerShell all-streams redirect), from the output dir `.\.build\x64\Debug`:
```powershell
.\PluginContractTests.exe *> ..\..\..\plans\005-baseline.txt   # *> captures stdout+stderr
```
This is temporary scaffolding — Step 6 deletes it; commit it on the plan branch only so the after-state can be diffed.

Also record the pre-change line counts so the Done-criteria "no file grew" check is exact:
```powershell
Get-ChildItem Plugins\*\Factory.cpp | ForEach-Object { "{0}`t{1}" -f $_.FullName.Split('\')[-2], (Get-Content $_ | Measure-Object -Line).Lines } > plans\005-baseline-linecounts.txt
```
(also temporary; delete in Step 6).

**Verify**: `PluginContractTests.exe` and `--selftest` exit codes both 0; `plans\005-baseline.txt` contains the per-DLL `[ OK ]` lines and ends with `PluginContractTests passed.`

### Step 3: Write Common\PlugInterfaces\FactoryImpl.h

Design (template-over-traits; a macro wrapper is acceptable if it reads better, but the logic lives once):

```cpp
// Each plugin's Factory.cpp provides a descriptor:
struct PluginFactoryEntry
{
    const PluginMetaData* metaData;                                  // static lifetime
    const char* (*getSchema)() noexcept;                             // may be nullptr (no schema)
    HRESULT (*createInstance)(const FactoryOptions* factoryOptions,  // future-extension hook, forwarded
                              IHost* host, void** result) noexcept;  // returns the DLL's interface, already QI'd
};
// One entry per logical plugin. Both function pointers are THIN PER-ENTRY THUNKS that assume the
// shared helper has already validated riid/result and matched the id — they do NOT re-check riid or
// null and do NOT re-match the id. Each thunk wraps the DLL's EXISTING creation/schema function,
// whatever shape it has today; do not assume a uniform discriminator. createInstance's first parameter
// is the FactoryOptions future-extension hook (see "Current state"): the shared FactoryCreate forwards
// it here, but leave it UNNAMED in every thunk today (it is unused, and naming it trips C4100 under /W4);
// a future plugin that consumes it names it in its own Factory.cpp only. The three existing creator shapes:
//   * single-plugin (e.g. FileSystemDummy): CreatePluginInstance(REFIID, IHost*, void**)
//       → thunk: [](const FactoryOptions*, IHost* host, void** r) noexcept { return CreatePluginInstance(__uuidof(IFileSystem), host, r); }
//   * multi by enum (ViewerWeb): CreatePluginInstance(REFIID, IHost*, ViewerWebKind, void**)
//       → one thunk per entry, baking in its kind: { return CreatePluginInstance(__uuidof(IViewer), host, ViewerWebKind::Web, r); }
//   * multi by id-string (FileSystemS3 / FileSystemCurl / FileSystemMicrosoftDrive):
//       CreatePluginInstance(REFIID, IHost*, std::wstring_view pluginId, void**) — dispatches internally
//       via ModeFromPluginId/ProtocolFromPluginId → one thunk per entry, baking in its OWN id literal:
//       { return CreatePluginInstance(__uuidof(IFileSystem), host, L"builtin/file-system-s3", r); }
// getSchema is a per-entry ZERO-ARG thunk returning THIS entry's schema (it takes no FactoryOptions —
// the RedSalamanderGetConfigurationSchema export has none). The existing getters are
// GetPluginSchema(std::wstring_view pluginId) and do their own id-match — drop that parameter (the
// shared helper now owns id matching) so the thunk just returns the entry's schema (single: call the
// static schema getter directly; multi: bake in the kind/mode, e.g. GetViewerWebStaticConfigurationSchema(kind)).
// Because these are plain function pointers they cannot capture — use named free functions or
// non-capturing lambdas, never a captured discriminator. The shared FactoryCreate passes
// __uuidof(TInterface) internally, so the descriptor signature need not carry riid.

// Shared implementation, parameterized by interface type:
template <typename TInterface>
HRESULT FactoryEnumeratePlugins(std::span<const PluginFactoryEntry> entries, REFIID riid, const PluginMetaData** metaData, unsigned int* count) noexcept;
template <typename TInterface>
HRESULT FactoryCreate(std::span<const PluginFactoryEntry> entries, REFIID riid, const FactoryOptions* factoryOptions, IHost* host, const wchar_t* pluginId, void** result) noexcept;
template <typename TInterface>
HRESULT FactoryGetConfigurationSchema(std::span<const PluginFactoryEntry> entries, REFIID riid, const wchar_t* pluginId, const char** schemaJsonUtf8) noexcept;
```

Semantics must replicate the existing files exactly. The one place single- and multi-plugin DLLs **diverge** is the empty-`pluginId` policy, so the shared helper must branch on `entries.size()` (or carry an explicit policy flag):
- `EnumeratePlugins`: `E_POINTER` on null out-params; zero out-params; `E_NOINTERFACE` if `riid != __uuidof(TInterface)`; else `*metaData = entries.data()`-equivalent and `*count = entries.size()`. Single-plugin DLLs are just the `entries.size()==1` case (one contiguous static `PluginMetaData`); multi-plugin DLLs return their `std::array<PluginMetaData,N>`.
- `Create`: `E_POINTER` on null `result`; `*result = nullptr`; IID gate (`E_NOINTERFACE`). Then the **count-dependent** empty-id policy — this is the regression-prone part, replicate it exactly:
  - **single entry** (`entries.size()==1`): empty/null `pluginId` ⇒ create that one entry (matches `FileSystemDummy\Factory.cpp:95` — empty falls through to creation).
  - **multiple entries**: empty/null `pluginId` ⇒ **`E_INVALIDARG`** (matches `ViewerWeb\Factory.cpp:180-183`, and Curl/MSDrive/S3).
  - non-empty id: `OrdinalString::EqualsNoCase` match across entries; `HRESULT_FROM_WIN32(ERROR_NOT_FOUND)` when none matches (uniform across single and multi). On match, delegate to `entry.createInstance(factoryOptions, host, result)`, forwarding `factoryOptions` unchanged. The shared helper must **forward only** — never dereference or validate `factoryOptions` (it may be null in principle, though the shipping host always passes a valid `&options`; null-checking is a future consumer's job in its own thunk).
- `GetConfigurationSchema`: `E_POINTER` on null `schemaJsonUtf8`; `*schemaJsonUtf8 = nullptr`; IID gate. Same count-dependent empty-id policy as `Create`: single ⇒ the only entry's schema; multiple ⇒ `E_INVALIDARG` (matches `ViewerWeb\Factory.cpp:206-209`). Non-empty id matches an entry; `ERROR_NOT_FOUND` when no entry matches or the matched entry has no schema.

⚠️ Do NOT collapse the empty-id case to "first entry" — that silently changes all 4 multi-id DLLs from `E_INVALIDARG` to a live instance/schema and is exactly the behavior shift the contract test exists to catch.

**`OrdinalString::EqualsNoCase` dependency (resolve, don't hand-wave):** it is declared in `Common\Helpers.h` — there is no lighter header for it. The earlier idea of keeping `FactoryImpl.h` free of `Helpers.h` is therefore infeasible without moving `OrdinalString`, which is out of scope. **Default: have `FactoryImpl.h` `#include "Helpers.h"` itself** — it is the robust choice and adds no new weight (every Factory.cpp TU already includes `Helpers.h` for the trace provider and `LoadStringResource`). If — and only if — self-including `Helpers.h` collides with the `REDSAL_DEFINE_TRACE_PROVIDER` one-definition rule (the macro must be defined in exactly one TU; including the header that consumes it from a shared header could double-instantiate), fall back to NOT including it and instead relying on the includer having `Helpers.h` in scope first, and document that include-order requirement at the top of `FactoryImpl.h`. The header MUST include `PlugInterfaces/Factory.h` (it declares `FactoryOptions`, `PluginMetaData`, and the export prototypes the templates mirror) and `<span>`; pull in nothing heavier than what Factory.cpp already includes.

**Verify**: `FactoryImpl.h` exists and is syntactically self-consistent. Note: `FactoryImpl.h` is **header-only** and this plan adds no Common `.cpp` that includes it, so `.\build.ps1 -ProjectName Common` will pass WITHOUT compiling it — that is not a real gate. The first true compile of the header happens in Step 4 (via FileSystemDummy). Do not treat a green Common build here as proof the header is correct.

### Step 4: Migrate ONE plugin (FileSystemDummy)

Rewrite `Plugins\FileSystemDummy\Factory.cpp` to: keep `GetPluginMetaData()`/schema getter/instance creator as local functions, define a `static const PluginFactoryEntry` array, and implement the three `extern "C"` exports as one-line delegations to the shared templates. The exported signatures stay verbatim. Keep the file's includes minimal but identical in effect.

**Verify**: `.\build.ps1` → exit 0; `PluginContractTests.exe` → exit 0. Re-capture both streams and compare to the Step 2 baseline — at this point 13 DLLs are untouched so the output must be **byte-identical** (the test iterates DLLs in a fixed array order; there is no legitimate reordering):
```powershell
.\PluginContractTests.exe *> after.txt
Compare-Object (Get-Content ..\..\..\plans\005-baseline.txt) (Get-Content after.txt)   # expect: NO output
```
Any non-empty `Compare-Object` result is a STOP condition.

### Step 5: Migrate the remaining 13

Mechanically, one commit-worthy batch for filesystem plugins, one for viewers. Preserve per-file elements that are NOT factory logic (do not lose them):
- `REDSAL_DEFINE_TRACE_PROVIDER` (13/14 files) — stays exactly where it is.
- The optional `RedSalamanderPluginShutdown` / `RedSalamanderPluginRetainModuleUntilProcessExit` exports — present in `FileSystemS3`, `ViewerSpace`, `ViewerWeb`; copy them across unchanged. They are separate from the three deduplicated functions.
- Multi-plugin DLLs (`FileSystemCurl`, `FileSystemMicrosoftDrive`, `FileSystemS3`, `ViewerWeb`) provide multi-entry arrays + per-entry discriminator thunks (per Step 1 / the descriptor note in Step 3); their empty-id policy is `E_INVALIDARG`, not first-entry.

**Verify** after each batch: full `.\build.ps1` → exit 0; `PluginContractTests.exe` exit 0; `Compare-Object` of a fresh `*>`-capture against the Step 2 baseline → **NO output** (byte-identical; the DLL iteration order is fixed, so any diff is a real behavior change → STOP).

### Step 6: Full verification + format

⚠️ **Do NOT run `.\format-all.ps1` here.** It takes no arguments and reformats *every* `*.cpp`/`*.h` in the repo, which dirties ~140 unrelated files and pollutes your diff (this happened on the first execution of this plan). Format **only the files you changed**, using the repo's `.clang-format` (`--style=file`), then prove nothing else was touched. From the repo root:

```powershell
# Resolve clang-format the same way format-all.ps1 does: PATH → Program Files LLVM → VS LLVM tools.
$cf = (Get-Command clang-format -ErrorAction SilentlyContinue).Source
if (-not $cf -or -not (Test-Path $cf)) { $cf = "${env:ProgramFiles}\LLVM\bin\clang-format.exe" }
if (-not (Test-Path $cf)) {
    $vs = & "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe" -latest -prerelease -products * -property installationPath | Select-Object -First 1
    $cf = "$vs\VC\Tools\Llvm\x64\bin\clang-format.exe"
}
if (-not (Test-Path $cf)) { throw "clang-format not found" }

# Format ONLY this plan's in-scope files:
$files = @('Common\PlugInterfaces\FactoryImpl.h') + (Get-ChildItem Plugins\*\Factory.cpp | ForEach-Object FullName)
$files | ForEach-Object { & $cf -i --style=file $_ }
```

**Verify**: `git status --porcelain` shows ONLY `Common/PlugInterfaces/FactoryImpl.h` and `Plugins/*/Factory.cpp` (plus the two `plans\005-baseline*.txt` scaffolding files until you delete them) — NO other source file may appear. If any other file is dirty, you ran the wrong formatter; revert those (`git checkout -- <file>`) before continuing.

Then: rebuild (`.\build.ps1` → exit 0); `PluginContractTests.exe` → exit 0 + `Compare-Object` vs baseline empty; full `--selftest`; optionally `.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild` (the one-command runner — there is no `test-all.ps1`). Delete BOTH scaffolding files in the final commit: `plans\005-baseline.txt` and `plans\005-baseline-linecounts.txt`.

## Test plan

No new tests — plan 004's `PluginContractTests` is the regression net (exports, enumeration, id matching incl. case-insensitivity and bogus-id `ERROR_NOT_FOUND`, schema parse, capabilities fields). The before/after output diff is the strongest gate. The full selftest run covers real host loading.

## Done criteria

- (archived, not live)  `Common\PlugInterfaces\FactoryImpl.h` exists; factory logic appears once
- (archived, not live)  Every `Plugins\*\Factory.cpp` is reduced to metadata/creator definitions + delegating exports (the three factory functions are one-line delegations; no hand-rolled E_POINTER/IID/id-match logic remains in any of them). Line-count gate (a reduction check, NOT an absolute cap — the multi-id files legitimately stay larger because they keep N metadata blocks + N creator/schema thunks + the shutdown exports): each **single-plugin** file < 90 lines, and **every** file is strictly shorter than its recorded `plans\005-baseline-linecounts.txt` value. If any file grew, something went wrong.
- (archived, not live)  `PluginContractTests.exe` exit 0 AND `Compare-Object` of a fresh `*>`-capture vs `plans\005-baseline.txt` produces no output (byte-identical)
- (archived, not live)  Full `--selftest` exit 0
- (archived, not live)  Host-side files unchanged: `git diff --stat -- RedSalamander/ Common/PlugInterfaces/Factory.h` → empty
- (archived, not live)  Scope check: `git status --porcelain` shows only `Common/PlugInterfaces/FactoryImpl.h` and `Plugins/*/Factory.cpp` modified/added (the two `plans\005-baseline*.txt` scaffolding files are deleted by Step 6)
- (archived, not live)  `plans/README.md` status row updated

## STOP conditions

Stop and report back if:

- Step 1 finds a factory with logic outside the catalogued axes — conditional/stateful/environment-dependent registration, a `factoryOptions` field actually being read, or an export beyond the five in `Factory.h`. (The optional shutdown/retain exports, `REDSAL_DEFINE_TRACE_PROVIDER`, multi-entry arrays, and per-entry discriminators are all EXPECTED — they are not stop conditions.)
- Plan 004's test is not present/green at HEAD — do not start without the net.
- Any contract-test output difference vs baseline that you cannot attribute to your own deliberate, behavior-identical restructuring — in particular any change to a multi-id DLL's empty-`pluginId` result (must stay `E_INVALIDARG`) or to which interface a DLL accepts.
- A plugin's instance creation needs runtime inputs beyond `IHost*` plus a compile-time-known per-entry discriminator (i.e. it can't be expressed as a fixed per-entry `createInstance(const FactoryOptions*, IHost*, void**)` thunk) without changing the plugin class itself (out of scope).

## Maintenance notes

- Future factory-contract changes now happen in `FactoryImpl.h` + the contract test together. Reviewers of plugin PRs should reject new hand-rolled factory exports.
- This unlocks a "new plugin" template/checklist (deferred; pairs with the plugin developer guide direction item recorded in plans/README.md).
- **`FactoryOptions` is a kept future hook**: it is forwarded through `FactoryCreate` into each `createInstance` thunk, where it sits named-but-unused. A future plugin that wants to read it (e.g. honor `DebugLevel` per `Specs/Plugins/Plugins_VirtualFileSystem.md:1918`) edits ONLY its own `Factory.cpp` thunk — no change to `FactoryImpl.h` or the other 13 factories. Reviewers should expect the parameter unnamed in every thunk until such a consumer appears.
- Reviewer focus: the **count-dependent empty-`pluginId` policy** (single ⇒ default-to-only; multi ⇒ `E_INVALIDARG`) and the per-entry discriminator thunks for multi-id DLLs — those are the two places behavior could silently shift. A new plugin added later inherits the shared logic automatically, but a single→multi transition (adding a 2nd id to a previously-single DLL) flips its empty-id contract from "default" to "error" — call that out in any such PR.
