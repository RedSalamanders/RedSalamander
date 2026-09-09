# Advisor Plan 026 - ViewerSqlite snapshot and display bounds

> **NON-NORMATIVE COMPLETE RECORD.** Archived from root `plans/026-viewersqlite-perf-and-headers.md` on 2026-08-25. **No remaining action.** Do not execute this file. Select current work only through `Specs/Plans/WIP/README.md`.

## Archive status

- **State:** COMPLETE
- **Archived:** 2026-08-25
- **Original path:** `plans/026-viewersqlite-perf-and-headers.md`
- **Live owner:** Specs/Plans/Done/Operation_Farsight_ViewerPluginsRemediation_ExportSafetyDecodeReliabilityWebSecurityAndTextGeometry_2026-06-16.md
- **No remaining action:** yes

> **Frozen history: do not resume.** Every executor instruction, checkbox, STOP condition, and "next action" below is 2026-06 archive text. **No remaining action.** If work continues, use the live owner named in Archive status.

---
# Plan 026: ViewerSqlite — reuse the read-only connection, page deep tables efficiently, and sanitize header text

## Status

- **Priority**: P3 (perf + minor robustness; correctness is otherwise sound)
- **Effort**: M
- **Risk**: LOW–MED (connection lifetime; must stay read-only and thread-correct)
- **Depends on**: none
- **Category**: perf
- **Planned at**: commit `b274022d9`, 2026-06-16
- **Implementation status (2026-07-12)**: DONE — Steps 1 and 3 are implemented and expanded through Farsight remediation; Step 2 is deliberately deferred to a separately designed cursor API. Focused plugin/engine/UI proof is green and archived at `Specs/TestRuns/4cb089111a23/Viewers/2026-07-11_173930_viewersqlite_fs22_snapshot_bounds/`; consolidated operation proof is green at `Specs/TestRuns/4cb089111a23/Viewers/2026-07-12_105903_farsight_viewer_closeout/`; the repository-wide Full gate is green at `Specs/TestRuns/4cb089111a23/Continuation/2026-07-12_205636_farsight_full_green/`.

### Reconciled implementation decisions

- `DatabaseSource` owns one `SQLITE_OPEN_READONLY | SQLITE_OPEN_NOMUTEX | SQLITE_OPEN_PRIVATECACHE` connection to an immutable private snapshot. A source mutex serializes every connection use; deterministic concurrent-call proof records a maximum concurrency of one.
- Local databases are frozen with SQLite backup, including committed WAL content. Virtual sources use bounded byte copying and reject a main header that advertises WAL because `IFileReader` cannot atomically include sidecars.
- Reopening the same immutable snapshot after a fatal corruption error would not heal corruption, so the plan's proposed in-place fatal reset/reopen is intentionally not claimed. A new viewer/source open creates a fresh snapshot and connection.
- Public paging remains the existing exact `rowOffset` contract. Sequential keyset paging requires new cursor state, duplicate/null ordering rules, `WITHOUT ROWID` identity discovery, and backward-navigation semantics; silently substituting it here would change observable behavior. That optimization is deferred rather than approximated.
- Exact SQLite identifiers are retained separately (with an explicit identity ceiling); only display names and column headers are bounded and sanitized. This preserves query identity while excluding controls and bidi-formatting characters from UI text.

## Why this matters

`DatabaseSource` holds only a file path and **opens a brand-new SQLite connection on every
operation** — `ListTables`, every `LoadTablePage` (each Next/Prev click and each sort), every custom
query, and every validate. `sqlite3_open_v2` re-reads the header, re-parses the schema, and rebuilds a
cold page cache each time (with `SQLITE_OPEN_PRIVATECACHE` nothing is shared). On a large untrusted DB
(often snapshotted to a temp file from a virtual FS), that is a fixed per-click latency tax. Separately,
table previews use `LIMIT … OFFSET …`, and SQLite implements `OFFSET` by stepping and discarding rows —
so paging deep into a million-row table re-scans from row 0 and degrades linearly to multi-second loads.

(Notes confirming what is **already fine**: connections use `SQLITE_OPEN_READONLY | NOMUTEX |
PRIVATECACHE` — read-only, no extension loading enabled; identifiers are quoted via `QuoteIdentifier`
(doubled quotes) so there is no SQL injection. Those are correct and out of scope.)

## Current state

```cpp
// ViewerSqlite.Engine.cpp:172-192  (a fresh connection every call)
HRESULT OpenReadOnlyConnection(const std::filesystem::path& localPath, unique_sqlite3& db, std::wstring& errorText) noexcept
{
    // ...
    const int rc = sqlite3_open_v2(pathUtf8.c_str(), &raw,
                                   SQLITE_OPEN_READONLY | SQLITE_OPEN_NOMUTEX | SQLITE_OPEN_PRIVATECACHE, nullptr);
    // ...
}
```
```cpp
// ViewerSqlite.Engine.cpp:511-521  (ListTables — opens a fresh connection; same in LoadTablePage / ExecuteReadOnlyQuery / ValidateReadOnlyQuery)
unique_sqlite3 db(nullptr, sqlite3_close_v2);
HRESULT hr = OpenReadOnlyConnection(_localPath, db, errorText);
```
```cpp
// ViewerSqlite.Engine.cpp:757-767  (preview SQL uses LIMIT/OFFSET — O(offset) deep paging)
std::wstring sql = std::format(L"SELECT * FROM {}", QuoteIdentifier(tableName));
if (sortDirection != ...) sql.append(std::format(L" ORDER BY {} {}", orderByColumnIndex + 1u, ...));
sql.append(std::format(L" LIMIT {} OFFSET {}", pageSize, rowOffset));
```
Also: cell TEXT goes through `SanitizeCellText` (`:194`, applied at `:264`), but **column header names
and table names do not** (`:286-291`, `:552-555`) — adversarial control characters in identifiers reach
the UI unsanitized.

## Commands you will need

| Purpose | Command | Expected |
|---|---|---|
| Build plugin | `.\build.ps1 -ProjectName ViewerSqlite` | exit 0, 0 warnings/errors |
| Full build | `.\build.ps1` | exit 0 |
| ViewerSqlite tests | `.build\x64\Debug\ViewerSqliteTests.exe` | exit 0 |
| Full suite | `.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild` | all pass |

## Scope

**Reconciled in scope**: the SQLite engine/source, private-snapshot lifetime, UI cancellation and
generation plumbing, shared diagnostics, resource strings, dedicated engine/UI tests, performance
instrumentation, and the authoritative ViewerSqlite spec. **Out of scope**: changing the public
`rowOffset` paging contract or designing a new cursor API; extension loading and unquoted identifiers
remain forbidden.

The original “only the two engine files modified” criterion is obsolete as of 2026-07-12 because the
immutable snapshot, cancellation, localization, tests, metrics, and durable contract are inseparable
from safe connection reuse.

## Git workflow

- Branch: `advisor/026-viewersqlite-perf`
- Message: `perf(ViewerSqlite): cache read-only connection; keyset deep paging; sanitize headers`
- Do NOT push/PR unless instructed.

## Steps

### Step 1: Cache one read-only connection for the DatabaseSource lifetime

Give `DatabaseSource` an owned `unique_sqlite3` opened once (lazily on first use) and reused by
`ListTables`/`LoadTablePage`/`ExecuteReadOnlyQuery`/`ValidateReadOnlyQuery`. Requirements:

- The connection is used only from the engine's worker thread. Confirm the engine serializes its
  operations on a single worker (it dispatches via the viewer's async machinery). If operations can run
  concurrently on multiple threads, either keep `SQLITE_OPEN_NOMUTEX` + serialize via a mutex around the
  connection, or open with `SQLITE_OPEN_FULLMUTEX`. **Verify the threading before sharing the handle** —
  reusing a `NOMUTEX` connection across threads without a lock is a bug. If unsure, add a `std::mutex`
  guarding the cached connection.
- Keep `SQLITE_OPEN_READONLY | SQLITE_OPEN_PRIVATECACHE`; finalize every prepared statement after use
  (already done) so the shared connection has no leaked statements.
- On a fatal error (corruption mid-session), reset the cached connection so the next call reopens.

**Verify**: `.\build.ps1 -ProjectName ViewerSqlite` → exit 0; `.build\x64\Debug\ViewerSqliteTests.exe` → exit 0.

### Step 2: Keyset (seek) pagination for deep table pages — deliberately deferred

Replace `LIMIT … OFFSET …` deep paging with keyset pagination so page N does not re-scan N×pageSize
rows. Approach:

- Page forward/back using a `WHERE <orderKey> > :lastSeen ORDER BY <orderKey> LIMIT :pageSize`
  predicate, where `<orderKey>` is the active sort column (or `rowid` when no sort is applied — SQLite
  tables have an implicit `rowid` unless `WITHOUT ROWID`). Carry the last-seen key (and rowid tiebreaker)
  from the previous page instead of a numeric offset.
- Add a stable tiebreaker (`, rowid`) to make ordering deterministic across pages even with duplicate
  sort keys.
- Handle `WITHOUT ROWID` tables (no implicit rowid): fall back to ordering by the primary key, or, if
  none is discoverable, retain `LIMIT/OFFSET` for that table and document the fallback.
- "Jump to arbitrary page/offset" (if the UI offers it) cannot be O(1) with keyset; keep `LIMIT/OFFSET`
  for explicit random jumps but use keyset for sequential Next/Prev (the common case).

This is the larger/riskier half. Live-code reconciliation confirmed that the existing public API is
an exact numeric `rowOffset` contract. Correct keyset paging requires durable cursor identity,
duplicate/null ordering, `WITHOUT ROWID` identity discovery, and reverse-navigation semantics. It is
therefore deferred to a new cursor API rather than approximated or silently substituted. Connection
reuse and private snapshots remove the dominant reopen/cold-cache cost without changing displayed rows.

**Verify**: `.\build.ps1 -ProjectName ViewerSqlite` → exit 0; ViewerSqliteTests green.

### Step 3: Sanitize header/table-name text

Run column header names (`:286-291`) and table names (`:552-555`) through `SanitizeCellText` (the same
function already applied to cell TEXT at `:264`) so adversarial control characters in identifiers cannot
reach the UI.

**Verify**: `.\build.ps1` → exit 0.

## Test plan

- `TestSnapshotConnectionBoundsCancellationAndSanitization` covers serialized cached connection use,
  exact-vs-display identifiers, display caps/sanitization, invalid sort/offset, cancellation, VM/result
  budgets, and concurrent-call diagnostics.
- `TestLocalWalSnapshotVirtualLimitsAndStaleScavenging` covers WAL-inclusive local backup/isolation,
  local/virtual byte ceilings, virtual WAL refusal, exact seek/read behavior, overlong identity,
  deterministic deletion, closed stale-artifact scavenging, and live-snapshot delete denial.
- Existing paging/sort, DX UI, and repeated open/close cases preserve the public offset behavior.
- Focused SQLite proof and consolidated operation proof are archived.

## Done criteria

ALL must hold:

- [x] `DatabaseSource` reuses a single read-only connection (no `sqlite3_open_v2` per page/query);
      thread-safety of the shared handle is verified or guarded.
- [x] Public `rowOffset` paging remains exact; keyset paging is explicitly deferred to a cursor API
      rather than changing observable behavior.
- [x] Column headers and table display names are bounded/sanitized while exact identifiers remain queryable.
- [x] Paging shows correct, non-duplicated, non-skipped rows (focused engine and UI paging/sort tests assert it).
- [x] Scoped `ViewerSqlite`/`ViewerSqliteTests` builds exit 0 with 0 warnings and focused scenarios pass.
- [x] The original two-file scope is explicitly superseded by required UI cancellation, localization, test, performance, and authoritative-spec changes.
- [x] `plans/README.md` reflects the reconciled implementation.
- [x] The consolidated Farsight archive is complete.
- [x] The repository-wide `.\Tools\Run-AllTests.ps1 -Suite Full` passes in the final closeout workspace (`Specs/TestRuns/4cb089111a23/Continuation/2026-07-12_205636_farsight_full_green/`).

## STOP conditions

- Excerpts don't match live code (drift).
- The engine turns out to run operations on multiple threads and you cannot establish a correct
  locking/threading model for a shared connection — keep per-call connections for now, land Step 3, and
  report Step 1 as blocked (do NOT share a `NOMUTEX` handle across threads unguarded).
- Keyset paging cannot reproduce the exact displayed-row set for some table shape — defer Step 2.

## Maintenance notes

- Keep the connection read-only; never enable extension loading or `SQLITE_OPEN_URI` (the current
  rejection of those is correct and is what keeps opening untrusted DBs safe).
- Reviewer: focus on (1) connection thread-safety, (2) keyset paging row-set equivalence and the
  tiebreaker, (3) that finalize/close still happen on all paths with the shared handle.
