# Plan 004: Re-verify HRESULT, Win32, Google Drive, Astrolabe, Rosetta, FO stress, and CI format owners

> **Executor instructions**: Follow this plan step by step. Run every
> verification command and confirm the expected result before moving to the
> next step. If anything in the "STOP conditions" section occurs, stop and
> report — do not improvise. When done, update the status row for this plan
> in `advisor-plans/README.md`.
>
> **Drift check (run first)**: `git diff --stat b5c3b0eab..HEAD -- Specs/Plans/WIP/Core_HResultStatusFormattingCleanup_2026-06-19.md Specs/Plans/WIP/Win32_Inventory.md Specs/Plans/WIP/FileSystem_GoogleDrivePluginPlan.md Specs/Plans/WIP/Operation_Astrolabe_TestContractArchitectureMigration_2026-07-17.md Specs/Plans/WIP/Operation_RosettaLantern_RedConfigureSatelliteLocalizationReview_2026-07-17.md Specs/Plans/WIP/FileOperations_RemainingStressValidation_2026-08-04.md Specs/Plans/WIP/CI_FormatAutocommitRace_2026-08-25.md Specs/Plans/WIP/README.md Specs/FileSystem/FileSystem_GoogleDrive.md RedSalamander/FolderViewInternal.h RedSalamander/FolderWindow.FileOperations.IssuesPane.cpp RedSalamander/RedSalamander.cpp Plugins/ViewerText/ViewerText.cpp Plugins/FileSystemGoogleDrive .github/workflows/ci.yml RedConfigure/Lang`

## Status

- **Priority**: P1
- **Effort**: M
- **Risk**: LOW for docs; HIGH if the executor “helpfully” implements CHK-1, PKCE, or HRESULT inlining while updating the plans
- **Depends on**: `advisor-plans/001-wip-lifecycle-and-status-truth.md`
- **Category**: docs
- **Planned at**: commit `b5c3b0eab`, 2026-08-25

## Why this matters

These ACTIVE owners still describe real unfinished work, but several **anchors and next-actions are false** at `b5c3b0eab`. Executors following them will patch the wrong lines, re-do CHK-1 that already uses `_hWnd.reset()`, or treat Rosetta as “copy English into satellites” after Czech/Japanese/Slovak UI strings already exist.

This plan is a **re-verification and document update**. It does not implement the remaining product work.

## Current state (verified at `b5c3b0eab`; re-run the greps)

### HRESULT (I7) — aliases still real; line numbers stale

`Core_HResultStatusFormattingCleanup_2026-06-19.md:8-16` claims `FormatHResult` at `FolderViewInternal.h:307` and `FormatStatusText` at `IssuesPane.cpp:51`.

Live:

```165:168:RedSalamander/FolderViewInternal.h
std::wstring FormatHResult(HRESULT hr)
{
    return FormatHResultMessage(hr);
}
```

```52:55:RedSalamander/FolderWindow.FileOperations.IssuesPane.cpp
[[nodiscard]] std::wstring FormatStatusText(HRESULT hr) noexcept
{
    return FormatHResultMessage(hr);
}
```

Callers still `FolderView.ErrorOverlay.cpp:593` and `:1191`. Test mirror `Tests/PerformanceTests2/Performance.Tests.FolderViewInternal.Access.h` still has `FormatHResult` (~:306). All five work-item checkboxes are still `[ ]`. FO coordination banner (`:1-3`) is still valid: do not break typed File Operations publication rows.

### Win32 (I8) — CHK-1 cited sites already `reset()`

Plan `Win32_Inventory.md:36-42` says NEXT ACTION is `DestroyWindow(_hWnd.get())` at ViewerText `:592/:602` and RedSalamander `:760/:1152`.

Live ViewerText confirm/cancel already:

```594:607:Plugins/ViewerText/ViewerText.cpp
        if (_hWnd && IsWindow(_hWnd.get()) != FALSE)
        {
            _hWnd.reset();
        }
```

About/Fatal dialogs already `_hWnd.reset()` at `RedSalamander.cpp:809-810` and `:1149-1150`. Remaining `DestroyWindow(_hWnd.get())` production grep at plan time hit **test** code `Tests/PerformanceTests2/Performance.Tests.FolderViewInternal.Access.h:772,:782` plus docs/skills.

CHK-3 (`DxUi.Accessibility.cpp` manual `Release()`, count 18 at plan time — live count still ~18) and CHK-6 (DPI AFTERPARENT audit) remain plausible. CHK-4 “27 CoUninitialize / 18 files” is stale; live `CoUninitialize(` hits are fewer and mixed with tests/plugins. WarpDrive Tasks 4/9 exclusion for CHK-4 is obsolete (those tasks are done).

Parallax PAR-3 still points at CHK-1 (`Operation_Parallax_…:146-147`) with stale RedSalamander line numbers — note it in the Win32 file; do not decide PAR-3 (plan 005).

### Google Drive (I0) — still a real product owner; dates/lines stale

`FileSystem_GoogleDrivePluginPlan.md` next action (`:16`) is still PKCE + download. Live:

- Capabilities still `"copy": false`, `"move": false`, `"delete": false` (`FileSystemGoogleDrive.cpp` ~`:137-140`).
- Mutations return `ERROR_NOT_SUPPORTED` (~`:1230-1390`).
- `rg pkce|code_verifier|LaunchInteractiveAuth` under `Plugins/FileSystemGoogleDrive` → no matches.
- Authoritative spec `Specs/FileSystem/FileSystem_GoogleDrive.md:16-21,191` already says reads/uploads not implemented and no PKCE UX yet.

Open product decisions (`:457-463`) are still unresolved (built-in client ID, Docs export, duplicate-name suffix). The plan still reads as if Phase 1 host OAuth work is the thing to do; it is done.

FO banner (`:1-3`) is load-bearing: mutation milestones need capability v2 + PluginContractTests for every advertised `true`.

### Astrolabe (I4) — still a real architecture owner

`Operation_Astrolabe_TestContractArchitectureMigration_2026-07-17.md` requires reviewed dispositions (`LexicalPolicy`, `GraphOrOwnership`, `CompanionWiring`, `ReplaceWithBehavior`, `Retire`). Live inventory still exposes heuristic `ReplacementCandidates` from `BehavioralShadow`/`MixedSourceShape` (`Tools/Modules/Testing/TestInventory.psm1:152-220`). No per-case disposition enum has landed. Keep ACTIVE. Add planned-at/STOP if 001 did not. Add verification commands. Do not treat the 133-candidate queue as approved rewrites (the plan already says this — keep that sentence).

### Rosetta Lantern (I6) — translation work has started; ledger is false

Plan (`:12-20`, `:41-45`) still implies satellites are English copies and all ledger boxes are `[ ]`.

Live `RedConfigure/Lang/cs-CZ/RedConfigure-cs-CZ.rc`, `ja-JP`, `sk-SK` already contain translated ordinary UI (e.g. Japanese `IDS_REDCONFIGURE_ERROR_CREATE_WINDOW`, Slovak `IDS_REDCONFIGURE_MODE_START_TITLE "Štart"`). `IDS_REDCONFIGURE_APP_TITLE "RedConfigure"` is a plausible language-neutral brand.

Remaining work is **inventory refresh vs current `RedConfigure.rc`, native-speaker review, allowlist, four-page smoke, spec merge** — not “translate from zero.”

### FO stress (I3) — structure good; still five unproven outcomes

`FileOperations_RemainingStressValidation_2026-08-04.md:30-45` lists popup render, long-name caps, concurrency evidence, parallel delete, Queue/Parallel switching. `Specs/Testing/Testing_TestCoverage.md:90` still points here. Re-grep `Get-TestInventory.ps1` / FileOps selftests for any outcome that already has a named case; if none, leave all five open. Do not invent case names.

### CI format (I9) — still accurate

Keep I9 as sole owner (`CI_FormatAutocommitRace_2026-08-25.md`). The 2026-08-04 write-up is frozen in Done. Do **not** revive root `plans/012-format-autocommit-race.md` (archived to `Specs/Plans/Done/AdvisorPlan_012-format-autocommit-race_2026-06-11.md` on 2026-08-25).

## Commands you will need

Run these **read-only** greps first and paste results into each WIP file’s “Current evidence” section:

```powershell
rg -n "FormatHResult\(|FormatStatusText\(" RedSalamander Tests/PerformanceTests2 -g "*.h" -g "*.cpp"
rg -n "DestroyWindow\(_hWnd\.get\(\)\)" Common Plugins RedConfigure RedSalamander RedSalamanderMonitor Tests
rg -n "->Release\(" Common/DxUi/DxUi.Accessibility.cpp
rg -n "CoUninitialize\(" Common Plugins RedConfigure RedSalamander RedSalamanderMonitor
rg -n "WM_DPICHANGED_AFTERPARENT" Common Plugins RedConfigure RedSalamander RedSalamanderMonitor
rg -n "ERROR_NOT_SUPPORTED|\"copy\": false" Plugins/FileSystemGoogleDrive/FileSystemGoogleDrive.cpp
rg -n "pkce|code_verifier|LaunchInteractiveAuth" Plugins/FileSystemGoogleDrive
rg -n "LexicalPolicy|ReplacementCandidates" Tools/Modules/Testing/TestInventory.psm1
Select-String -Path RedConfigure\Lang\*\RedConfigure-*.rc -Pattern 'IDS_REDCONFIGURE_' | Measure-Object
.\Tools\Get-SpecInventory.ps1 -FailOnFindings
git diff --check
```

Do not implement fixes. Do not run Full.

## Scope

**In scope:** the seven WIP files in the drift check, plus `Specs/Plans/WIP/README.md` next-action cells for I0/I4/I6/I7/I8 (and I3/I9 if a sentence is wrong).

**Out of scope:**
- Any production C++/RC/workflow change
- I11
- Archived `Specs/Plans/Done/AdvisorPlan_012-format-autocommit-race_2026-06-11.md` (do not revive)
- Translating Rosetta strings
- Adding Astrolabe dispositions in Pester (product work stays on I4)
- Implementing Google Drive PKCE
- Parallax decision outcomes (plan 005)

## Git workflow

- Branch: `advisor/004-refresh-stale-active-owners`
- One commit per owner file is OK, or one commit with a message `Refresh stale WIP owner evidence`.
- Do NOT push unless asked.

## Steps

### Step 1: HRESULT

Update the audit snapshot date and line numbers to the live `FormatHResult` / `FormatStatusText` / callers / test mirror. Keep all work items `[ ]` unless a grep shows an alias gone. Keep FO coordination. Next action may stay WI-1 **after** confirming WarpDrive/FolderView owners are not mid-edit on `FolderViewInternal.h`.

**Verify**: plan no longer cites `:307` for FormatHResult.

### Step 2: Win32

- Mark CHK-1 **DONE** for the four cited unique_hwnd owner sites **only if** the Step-0 grep shows zero `DestroyWindow(_hWnd.get())` in `Plugins/ViewerText` and `RedSalamander/RedSalamander.cpp`.
- Record the remaining test-harness hits in `Performance.Tests.FolderViewInternal.Access.h` as either (a) out of scope / test-double, or (b) a new CHK row — do not silently expand CHK-1 to rewrite PerformanceTests2 unless you add a remaining-outcome sentence. Prefer (a) with a one-line rationale (test stub vs production owner).
- Refresh CHK-3 count from grep (do not “fix” Release sites).
- Refresh CHK-4: replace “27 calls / 18 files” with the live file list; remove “EXCLUDE WarpDrive Tasks 4/9” or replace with “coordinate with FolderView owners; Tasks 4/9 are complete.”
- CHK-6: keep open; update the 6→11 file list only if you re-grep `WM_DPICHANGED_AFTERPARENT`.
- Point Parallax PAR-3 at this refreshed CHK-1 status (one sentence). Do not fill Parallax outcomes.

**Verify**: Win32 “NEXT ACTION” is no longer ViewerText:592 DestroyWindow.

### Step 3: Google Drive

- Keep ACTIVE. Refresh “Last updated” and ERROR_NOT_SUPPORTED line range from grep.
- Checkboxes: keep download/write/shortcut/remote tests `[ ]`; keep enumeration/hardening `[x]`.
- Rewrite “Next action” to: Phase 2 remainder = interactive PKCE + read stream, **blocked on open product decisions** (client ID, Docs export, duplicate-name UX) until a maintainer records accept/reject in this file or Atlas.
- Do not resolve those product decisions here.

**Verify**: plan does not claim “no Drive feature commit since 2026-04-25” unless `git log -- Plugins/FileSystemGoogleDrive` still supports it; replace with the actual latest commit date from `git log -1 --format=%ci -- Plugins/FileSystemGoogleDrive`.

### Step 4: Astrolabe

Add verification:

```powershell
.\Tools\Get-TestInventory.ps1
Invoke-Pester .\Tools\Tests\TestInventory.Tests.ps1 -PassThru
```

State that `ReplacementCandidates` remaining > 0 is expected until dispositions exist. STOP: do not delete regex tests without a named behavioral companion (already in the file — keep it).

### Step 5: Rosetta

- Recompute source-identical string count vs `RedConfigure.rc` (do not paste a 77-item table unless you generate it). Record: translations exist; native review and allowlist still open.
- Check inventory-refreshed boxes only if you actually produced the inventory in this change (a markdown table of remaining English-identical IDs is in scope). Leave Native review `[ ]`.
- Do not machine-translate.

**Verify**: the “Why this operation exists” paragraph no longer says satellites are byte-for-byte English.

### Step 6: FO stress and CI format

- FO stress: if you find a named selftest that fully covers one of the five outcomes, check that outcome and cite the case name. Otherwise leave all five open and say “re-verified YYYY-MM-DD, still missing.”
- CI format: confirm `git add -A` and `head_ref` still present; if the job changed to check-only, STOP (plan already lists that STOP).

### Step 7: README next actions

I7: WI-1 aliases at the live lines. I8: CHK-3/4/6 remain; CHK-1 production sites done. I0: PKCE + download; product decisions still open. I6: native review, not from-scratch translation. I4 unchanged intent.

## Test plan

- No new product tests.
- `Get-SpecInventory.ps1 -FailOnFindings` must pass.
- Optional: `git log -1 -- Plugins/FileSystemGoogleDrive` for the date stamp.

## Done criteria

- [ ] HRESULT line numbers match live code
- [ ] Win32 CHK-1 matches live `reset()` vs remaining `DestroyWindow(_hWnd.get())`
- [ ] Google Drive next-action and capability facts match live plugin + `FileSystem_GoogleDrive.md`
- [ ] Rosetta no longer claims English-only satellites
- [ ] Astrolabe still forbids treating heuristic ReplacementCandidates as a rewrite queue
- [ ] I9 still sole CI format owner; root `plans/` stays deleted (archived to Done 2026-08-25)
- [ ] No production code changes
- [ ] `advisor-plans/README.md` row 004 updated

## STOP conditions

- CHK-1 production `DestroyWindow(_hWnd.get())` **has returned** on unique_hwnd owners — then CHK-1 is not DONE; update evidence, do not implement in this plan.
- Google Drive already implements PKCE or download — then I0 remaining outcomes must be rewritten from code, not from this excerpt.
- Rosetta satellites are English again (revert) — then keep the original from-scratch wording.
- Implementing any remaining WI/CHK/Phase item looks “small.” Stop; that is the ACTIVE owner’s product work, not this survey.

## Maintenance notes

- Reviewer: this branch should be markdown (+ maybe README). Reject C++.
- After CHK-1 is marked done, Parallax PAR-3 should be considered for SUPERSEDED in plan 005 — do not do it here if 005 has not run; leave a one-line pointer.
- Google Drive open decisions should not silently become “do PKCE with a shipped client ID.”
