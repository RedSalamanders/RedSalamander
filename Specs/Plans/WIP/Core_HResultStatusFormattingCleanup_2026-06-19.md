# Core HRESULT/Status Formatting Cleanup

Last updated: 2026-06-19

Status: WIP

Source split:

- `Specs/Plans/Done/Core_HardeningImprovementPlan_2026-03-19.md`

## Purpose

This plan owns the residual HRESULT/system-message and status text cleanup split out of the Core Hardening umbrella. The callback drain and quiet-point hardening work is already authoritative in public headers/specs; this plan is only about keeping user-visible failure/status text consistent, localized, and backed by the shared HRESULT helpers.

## Source Of Truth

- `Common/Helpers.h`
  - `FormatHResultMessage(HRESULT hr)`
  - `FormatHResultMessageWithCode(HRESULT hr)`
- `Specs/Core/Core_Localization.md`
- `.github/skills/compiler-warnings/SKILL.md`
- `.github/skills/error-handling/SKILL.md`

## Scope

Candidate surfaces from the original hardening plan:

- Connection Manager validation and connection-failure details.
- Find Files status hinting.
- File Operations issue/popup status surfaces.
- FolderView error overlays and clipboard/file-operation overlays.
- FolderWindow command failure alerts and archive operation status.

Non-goals:

- Do not rewrite diagnostic-only `Debug::Error(...)` / `Debug::Warning(...)` messages that already intentionally include numeric HRESULT codes.
- Do not change binary/technical viewers such as PE inspection rows where hexadecimal values are domain data.
- Do not refactor unrelated UI or error-flow behavior.

## Work Rules

- Classify every candidate hit before changing it: user-visible status text, localized resource string, diagnostic log, domain data, or dead code.
- Prefer `FormatHResultMessageWithCode(...)` when a user-visible failure benefits from both the stable code and the OS-provided message.
- Prefer `FormatHResultMessage(...)` only where the surrounding UI already carries enough context and the numeric code adds noise.
- Keep file-local wrappers only when the wrapper name encodes domain/UI intent. Remove aliases that only rename `FormatHResultMessage(...)`.
- Keep diagnostics on `Debug::Error(...)`, `Debug::Warning(...)`, or `Debug::ErrorWithLastError(...)`; do not spend UI string work on diagnostic-only paths.
- Do not touch unrelated control flow, retry behavior, dialog layout, or operation semantics while changing status text.

## Current Audit Snapshot - 2026-06-19

Audit command:

```powershell
rg -n "FormatMessageW|FormatHResult\(|FormatStatusText\(|sprintf_s|swprintf_s|C4774" RedSalamander Common Plugins -g "*.cpp" -g "*.h" -g "*.rc" -g "*.vcxproj"
```

Findings:

- `Common/Helpers.h` is the only production source using `FormatMessageW` for shared HRESULT/system-message text.
- `RedSalamander/FolderViewInternal.h` still has a trivial `FormatHResult(HRESULT)` alias over `FormatHResultMessage(...)`.
- `RedSalamander/FolderWindow.FileOperations.IssuesPane.cpp` still has a trivial `FormatStatusText(HRESULT)` alias over `FormatHResultMessage(...)`.
- No `sprintf_s`, `swprintf_s`, or `C4774` hits were found in the audited production paths.

The broader numeric-code audit has many intentional diagnostic and domain-data hits. Future cleanup must classify user-visible status text separately from diagnostics before changing strings.

Additional triage command for the implementation pass:

```powershell
rg -n "static_cast<unsigned long>\(.*hr|0x\{:08X\}|HRESULT 0x|FormatHResultMessage\(|FormatHResultMessageWithCode\(" RedSalamander Common Plugins -g "*.cpp" -g "*.h"
```

## Surface Triage

| State | Surface | Initial classification | Expected action |
|-------|---------|------------------------|-----------------|
| [ ] | `RedSalamander/FolderViewInternal.h` `FormatHResult(HRESULT)` | Trivial alias | Inline/remove unless a caller-specific name is useful. |
| [ ] | `RedSalamander/FolderWindow.FileOperations.IssuesPane.cpp` `FormatStatusText(HRESULT)` | Trivial alias | Inline/remove or rename only if issue-pane semantics are added. |
| [ ] | FolderView error overlays | User-visible failure details | Prefer shared helpers; preserve code where support/debugging needs it. |
| [ ] | File Operations issue/popup status surfaces | User-visible operation status | Prefer shared helpers; add focused FileOps/Commands evidence if wording changes. |
| [ ] | Connection Manager validation/failure details | User-visible connection failure status | Prefer shared helpers; keep localized resource placeholders positional. |
| [ ] | Find Files status hinting | User-visible status/hint text | Change only if current output is numeric-only or inconsistent. |
| [ ] | FolderWindow command/archive failure alerts | User-visible failure status | Prefer shared helpers; preserve actionable context. |
| [ ] | Diagnostic-only `Debug::*` call sites | Diagnostics | Leave numeric codes when intentional; no user-visible wording obligation. |
| [ ] | PE/binary/technical viewer rows | Domain data | Leave hexadecimal/domain values unchanged unless incorrectly presented as an error message. |

## Work Items

| State | Item | Required proof |
|-------|------|----------------|
| [ ] | Remove or inline trivial local wrappers that add no domain meaning. | Grep shows no trivial `FormatHResult(...)` / `FormatStatusText(...)` aliases in production code unless they encode UI semantics. |
| [ ] | Audit candidate user-visible surfaces for numeric-only HRESULT text. | Each changed surface has before/after notes and uses `FormatHResultMessageWithCode(...)` when both code and system text matter. |
| [ ] | Keep localized resource format strings positional. | `Tools/Tests` localization/resource guard remains green. |
| [ ] | Preserve diagnostics policy. | Diagnostic-only logs still use `Debug::*` and do not allocate extra UI strings unnecessarily. |
| [ ] | Add/update deterministic assertions for any user-visible status wording changed by this plan. | Focused Commands/FileOps/Compare/Viewer test passes and is archived. |

## Validation

Minimum validation:

```powershell
.\build.ps1 -ProjectName RedSalamander
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter "folderView|fileops|find|connection" -TimeoutMultiplier 2.0
```

If a change touches File Operations status/issue surfaces, also run the relevant FileOps focused suite. If a change touches resource format strings, run the `Tools\Tests` Pester resource/localization guards or `.\Tools\Run-AllTests.ps1 -Suite Full` before closeout.

Archive evidence under `Specs/TestRuns/<machine>/Core/<timestamp>/HResultStatusFormattingCleanup/` or the owning subsystem folder when the touched surface has a more specific test archive convention.

## Closeout Criteria

- [ ] Every changed surface has a short before/after note in this plan or its successor closeout note.
- [ ] Every user-visible string change has focused deterministic coverage or an explicit reason why existing coverage is sufficient.
- [ ] Resource placeholders remain positional and translator-safe.
- [ ] Production code has no `sprintf_s`, `swprintf_s`, or `C4774` regressions.
- [ ] The final audit commands above have been rerun and summarized.
- [ ] Validation evidence is archived under `Specs/TestRuns/`.
- [ ] Move this file to `Specs/Plans/Done/` only after the checklist and validation evidence are complete.
