# Advisor Plan 002 - Remove duplicate sqlite3 and refresh structure docs

> **NON-NORMATIVE COMPLETE RECORD.** Archived from root `plans/002-vcpkg-dup-and-docs-refresh.md` on 2026-08-25. **No remaining action.** Do not execute this file. Select current work only through `Specs/Plans/WIP/README.md`.

## Archive status

- **State:** COMPLETE
- **Archived:** 2026-08-25
- **Original path:** `plans/002-vcpkg-dup-and-docs-refresh.md`
- **Live owner:** none (implemented in vcpkg.json and contributor docs)
- **No remaining action:** yes

> **Frozen history: do not resume.** Every executor instruction, checkbox, STOP condition, and "next action" below is 2026-06 archive text. **No remaining action.** If work continues, use the live owner named in Archive status.

---
# Plan 002: Remove the duplicate sqlite3 vcpkg entry and fix stale structure docs

> **Frozen historical instructions (do not execute)**: This plan is archived; do not follow these steps. Run every
> verification command and confirm the expected result before moving to the
> next step. If anything in the "STOP conditions" section occurs, stop and
> report — do not improvise. Historical: status-row updates no longer apply for this plan
> in `plans/README.md`.
>
> **Drift check (run first)**: `git diff --stat 29ac6670c..HEAD -- vcpkg.json CLAUDE.md README.md`
> If any in-scope file changed since this plan was written, compare the
> "Current state" excerpts against the live code before proceeding; on a
> mismatch, treat it as a STOP condition.

## Status

- **Priority**: P1
- **Effort**: S
- **Risk**: LOW
- **Depends on**: none
- **Category**: deps + docs
- **Planned at**: commit `29ac6670c`, 2026-06-12 (refresh of the original `d6bfccc42` plan — both findings re-verified live at this SHA; line anchors updated after plan 001's doc merge `1bb1fb5c6` touched CLAUDE.md/README.md)

## Why this matters

`vcpkg.json` declares `sqlite3` twice with different minimum versions — silent ambiguity that confuses dependabot baseline bumps and human audits. Separately, the two entry-point docs (CLAUDE.md, README.md) describe a repo layout that no longer exists: CLAUDE.md's architecture tree lists plugins at the repo root ("18 projects"), but all 14 plugins actually live under `Plugins\`, and the tree omits `RedConfigure/`, `RedLauncher/`, `RedSalamanderSearchService/`, `Tests/`, `Tools/`, and `Installer/`. This repo's workflow leans on AI agents that bootstrap from CLAUDE.md, so a wrong map has outsized cost: agents search wrong paths and mis-state the architecture.

## Current state

- `vcpkg.json:19-22` — first sqlite3 entry (verified at `29ac6670c`):
  ```json
  {
    "name": "sqlite3",
    "version>=": "3.53.2"
  },
  ```
- `vcpkg.json:57-60` — duplicate entry, lower floor (sits between the `ftxui` and `aws-sdk-cpp` entries):
  ```json
  {
    "name": "sqlite3",
    "version>=": "3.50.4"
  },
  ```
- `CLAUDE.md:33-52` — the stale tree, verbatim at `29ac6670c`:
  ```text
  ### Solution Structure (18 projects)

  ```text
  RedSalamander/          # Main file manager application
  RedSalamanderMonitor/   # Monitoring/debug tool with ColorTextView
  Common/                 # Shared library (utilities, settings, plugin interfaces)
    └── PlugInterfaces/   # COM-style plugin interfaces (IFileSystem, IViewer, etc.)

  FileSystem/             # Standard Win32 file system plugin
  FileSystem7z/           # 7-Zip archive plugin
  FileSystemDummy/        # Test/stub plugin

  ViewerText/             # Text/hex viewer plugin
  ViewerSpace/            # Disk space visualization
  ViewerImgRaw/           # Raw image viewer (libraw)
  ViewerVLC/              # Video player (VLC backend)
  ViewerPE/               # PE executable viewer

  PoC/                    # Proof-of-concept projects
  ```
  ```
  Reality (verified): all 14 plugins live under `Plugins\` (`FileSystem`, `FileSystem7z`, `FileSystemCurl`, `FileSystemDummy`, `FileSystemGoogleDrive`, `FileSystemMicrosoftDrive`, `FileSystemS3`, `ViewerImgRaw`, `ViewerPE`, `ViewerSpace`, `ViewerSqlite`, `ViewerText`, `ViewerVLC`, `ViewerWeb`), and the tree omits `RedConfigure/`, `RedLauncher/`, `RedSalamanderSearchService/`, `Tests/` (8 test projects), `Tools/` (+ Pester tests in `Tools/Tests`), `Installer/`, `docs/`, `Specs/`. The solution also contains ~90 per-language localization satellite resource projects, so raw .sln project counts mislead.
- **Nearby content you must NOT touch** (added by plan 001, merged as `1bb1fb5c6`): `CLAUDE.md:29` — the line "Run the full local test suite with `.\Tools\Run-AllTests.ps1 -Suite Full` …" just above `## Architecture`. Leave it exactly as is.
- `README.md:136-147` — "### Solution Structure" list. Accurate on project *names* but doesn't say the plugins live under `Plugins\`, and doesn't mention the localization satellites; a first-time reader opening the .sln sees ~111 projects and assumes the README is stale.
- **Nearby README content you must NOT touch**: the "### Run all tests with one command" subsection (~`README.md:196-213`, also from plan 001).
- `CLAUDE.md:135` already contains `Specs/Plugins/...` doc-path references — which is why the verification gates below use **line-anchored** patterns (`^Plugins/`), not a bare `Plugins/` substring (a bare match passes before any work is done).
- CLAUDE.md is otherwise authoritative project guidance — change ONLY the structural/layout facts, not the development guidelines.

## Commands you will need

| Purpose | Command | Expected on success |
|---------|---------|---------------------|
| JSON validity + dup check | `(Get-Content vcpkg.json -Raw | ConvertFrom-Json).dependencies | Where-Object { $_.name -eq 'sqlite3' } | Measure-Object | Select-Object -ExpandProperty Count` | `1` |
| Optional full dep restore | `.\vcpkg-install.ps1` | exit 0 (slow; optional) |
| Doc grep gates | see Done criteria | as stated |

## Scope

**In scope** (the only files you should modify):
- `vcpkg.json`
- `CLAUDE.md` (the "Solution Structure" tree and any other structural/path facts that reference plugin locations — structural facts only)
- `README.md` (the "Solution Structure" list — add `Plugins\` context and a one-line note about localization satellite projects)

**Out of scope** (do NOT touch):
- The Run-AllTests lines/subsections in CLAUDE.md (`:29`) and README.md (~`:196-213`) — freshly landed by plan 001.
- `AGENTS.md`, `.github/skills/**`, `docs/**`, `Specs/**` — other docs may also drift, but auditing them is not this plan.
- Any `.vcxproj`/`.sln` — no build-system changes.
- The `builtin-baseline` value or any other vcpkg dependency/version — only the duplicate sqlite3 entry is removed.

## Git workflow

- Branch: `advisor/002-vcpkg-docs-refresh`
- One commit is fine; subject style like "Remove duplicate sqlite3 manifest entry and refresh structure docs".
- Do NOT push or open a PR unless the operator instructed it.

## Steps

### Step 1: Remove the duplicate sqlite3 entry

In `vcpkg.json`, delete the second sqlite3 object (the one with `"version>=": "3.50.4"` at `:57-60`, between `ftxui` and `aws-sdk-cpp`). Keep the first (`3.53.2` — the higher floor, matching the dependabot bump merged as `f2f75e3f3`). Watch trailing-comma validity.

**Verify**: the dup-check command above returns `1`, and `Get-Content vcpkg.json -Raw | ConvertFrom-Json` parses without error.

### Step 2: Fix the CLAUDE.md structure section

Replace the tree at `CLAUDE.md:33-52` (excerpted above) with one that reflects reality. Target shape (adjust wording freely, facts must match):

```text
### Solution Structure (~21 core projects + localization satellites)

RedSalamander/             # Main file manager application (incl. SelfTest\ suites)
RedSalamanderMonitor/      # ETW monitoring/debug tool with ColorTextView
RedSalamanderSearchService/# Background search/index Windows service (named pipe + SQLite)
RedConfigure/              # Standalone configuration tool
RedLauncher/               # Launcher app
Common/                    # Shared library (utilities, settings, DxUi framework)
  └── PlugInterfaces/      # COM-style plugin interfaces (IFileSystem, IViewer, ...)
Plugins/                   # All plugin DLLs
  ├── FileSystem, FileSystem7z, FileSystemCurl, FileSystemS3,
  │   FileSystemGoogleDrive, FileSystemMicrosoftDrive, FileSystemDummy
  └── ViewerText, ViewerSqlite, ViewerSpace, ViewerImgRaw,
      ViewerVLC, ViewerPE, ViewerWeb
Tests/                     # 8 standalone test projects (DxUiTests, PerformanceTests2, ...)
Tools/                     # PowerShell tooling (+ Pester tests in Tools/Tests)
Installer/                 # MSIX + MSI packaging
PoC/                       # Proof-of-concept projects
```

Add one sentence after the tree: the solution additionally contains ~90 per-language localization satellite resource projects, so Solution Explorer shows far more than the core list. Then scan the rest of CLAUDE.md for other root-level plugin path references and fix any path that no longer exists (the "Key Components" table at `:61-70` is already correct; `Specs/Plugins/...` references at `:135` are doc paths, not directories — leave them).

**Verify**:
- `Select-String -Path CLAUDE.md -Pattern '^(FileSystem|Viewer)'` → **no matches** (no root-level plugin tree entries remain).
- `Select-String -Path CLAUDE.md -Pattern '^Plugins/'` → **≥1 match** (the new tree entry; note the anchored `^` — a bare `Plugins/` would false-pass on the `Specs/Plugins/` references at `:135`).
- `Select-String -Path CLAUDE.md -Pattern '18 projects'` → no matches.

### Step 3: Clarify README solution structure

In `README.md` "### Solution Structure" (`:136-147`): state that file-system and viewer plugins live under `Plugins\`, and add one sentence: the solution also contains per-language localization satellite projects, so Solution Explorer shows far more than the core projects.

**Verify**: `Select-String -Path README.md -Pattern 'Plugins\\' | Where-Object LineNumber -lt 160` → ≥1 match in that section, and `Select-String -Path README.md -Pattern 'satellite'` → ≥1 match.

### Step 4 (optional, recommended if time permits): restore check

Run `.\vcpkg-install.ps1` to prove the manifest still resolves. This is slow (vcpkg restore); skip if the environment lacks vcpkg, and say so in the report.

## Test plan

No code tests — verification is the JSON parse gate plus the grep gates above. If Step 4 is run, a successful restore is the strongest gate.

## Done criteria

- (archived, not live)  Exactly one `sqlite3` entry in `vcpkg.json` (PowerShell gate above returns `1`), JSON parses
- (archived, not live)  `Select-String -Path CLAUDE.md -Pattern '^(FileSystem|Viewer)'` → no matches; `'^Plugins/'` → ≥1 match; `'18 projects'` → no matches
- (archived, not live)  README structure section mentions `Plugins\` (before line 160) and the localization satellites
- (archived, not live)  The Run-AllTests content from plan 001 is byte-identical (`git diff` on CLAUDE.md shows no change at `:29`; README's "Run all tests with one command" subsection untouched)
- (archived, not live)  `git status` shows only the three in-scope files modified
- (archived, not live)  `plans/README.md` status row updated

## STOP conditions

Stop and report back if:

- `vcpkg.json` at HEAD no longer contains two sqlite3 entries (someone already fixed it) — skip Step 1, do the docs steps, note it.
- You find evidence the duplicate is intentional (e.g. a comment, or a feature-conditional dependency structure) — vcpkg manifests don't support per-feature version floors this way, but if anything looks deliberate, report instead of deleting.
- CLAUDE.md's structure section no longer matches the `:33-52` excerpt above — reconcile against the live file rather than pasting blindly, and report the drift.

## Maintenance notes

- The project inventory will drift again; the durable fix would be a Pester check comparing README/CLAUDE.md claims to the .sln (deferred — noted in plans/README.md).
- Reviewer focus: that no vcpkg version other than the duplicate sqlite3 entry changed, and that plan 001's freshly-merged doc lines were not disturbed.
