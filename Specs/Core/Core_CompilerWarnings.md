# Compiler Warnings Specification

## Overview

RedSalamander first-party builds default to **`/Wall`** (`<WarningLevel>EnableAllWarnings</WarningLevel>`) through `Directory.Build.props` to keep a high bar for code quality.

RedSalamander first-party VC++ builds also inject **`/FS`** through `Directory.Build.props`. That keeps shared compiler PDB access stable under repo-wide `/MP` and solution-graph builds, instead of relying on ad-hoc per-project fixes.

Proof-of-concept projects inherit the deliberate exception through `PoC/Directory.Build.props`: they default to `Level3` plus the shared `C4710/C4711` suppression so experiments are still consistent without weakening the repo-wide default. Any extra warning deviations beyond that shared PoC baseline should stay local to the specific PoC project file.

Some MSVC warnings are **optimizer/inlining heuristics** or generated-header/codegen noise that become extremely noisy under `/Wall` (especially in template-heavy code paths) and do not represent correctness issues. This spec documents which ones we suppress and why.

`PerformanceTests2` is a documented local exception: it imports Windows SDK-, WIL-, STL-, and product-source-heavy translation units into a `/Wall` native unit-test harness, so only those imported source files carry a local suppression list for expected header/inlining noise (`C4061`, `C4100`, `C4191`, `C4263`, `C4264`, `C4355`, `C4365`, `C4464`, `C4619`, `C4625`, `C4626`, `C4668`, `C4710`, `C4711`, `C4820`, `C4865`, `C5026`, `C5027`, `C5039`, `C5045`, `C5204`, `C5220`, `C5246`). The project’s own test files remain under the repo-wide `/Wall` default.

## Suppression Review Routine

Every non-shared warning suppression MUST stay local to the narrowest project, file, or pragma scope that needs it. Each suppression MUST document:

- rationale: why the warning is not currently actionable,
- owner: the project/component that owns the exception,
- expiry/review task: the condition or follow-up that should remove or narrow it.

Production/runtime projects MUST prefer fixing root causes over adding project-wide suppressions. Test projects MAY carry broader suppressions only when they intentionally aggregate SDK, WIL, STL, or product-source-heavy translation units, and the project file documents the rationale/owner/expiry. `C5039` is especially sensitive: production callback exception-spec warnings must be fixed or suppressed at the exact callback boundary with a local rationale, not hidden at project scope.

## Suppressed Warnings

### C4710 — function not inlined

MSVC warning C4710 (“function not inlined”) is emitted when the optimizer decides not to inline a function it considered.

Under `/Wall` in **Release** builds, this can produce a large volume of warnings where the reported file/line is often in:
- the STL headers (e.g. `...\VC\Tools\MSVC\...\include\algorithm`), and/or
- other headers instantiated by our translation units.

This is expected and is typically **not actionable** without a specific performance investigation.

### C4711 — selected for automatic inline expansion

MSVC warning C4711 (“function selected for automatic inline expansion”) is similarly an optimizer note.

With `/Wall`, C4711 tends to appear (or disappear) based on small changes in codegen, compiler version, PGO/LTO, and other optimization settings. It frequently does not correlate with real-world performance issues and creates noise in build output.

### C4514, C4820, and C5045 — shared noise baseline

The shared production baseline may also suppress warnings that are dominated by SDK, STL, WIL, or generated-header noise and are not useful as per-change quality gates:

- `C4514`: unreferenced inline function removed.
- `C4820`: padding added after a data member.
- `C5045`: compiler will insert Spectre mitigation for memory load if `/Qspectre` switch specified.

These are allowed only as a documented shared baseline or as narrow project/file exceptions with a rationale. They must not be used to hide actionable production callback, ABI, lifetime, or data-layout warnings.

## Rationale

We suppress C4710/C4711 and the shared noisy-warning baseline because:
- They are **not correctness warnings**.
- They are heavily dependent on optimizer heuristics, SDK headers, generated headers, and compiler version, so they are **unstable** across toolset updates.
- They are disproportionately noisy with `/Wall` and templates, drowning out actionable warnings.
- When performance matters, the correct workflow is to **profile first** and then make targeted changes (algorithm/data layout/`__forceinline` where justified), not to chase bulk inlining warnings.

## Project File Policy

All C++ projects keep `/Wall` enabled and disable the shared warning baseline via MSBuild. Current project files express that baseline with `DisableSpecificWarnings` plus `/wd` options, for example:

```xml
<AdditionalOptions>/wd5045 /wd4820 %(AdditionalOptions)</AdditionalOptions>
<DisableSpecificWarnings>4710;4711;4514;%(DisableSpecificWarnings)</DisableSpecificWarnings>
```

Notes:
- We keep `%(DisableSpecificWarnings)` to preserve any existing warning suppressions from imported property sheets.
- `/external:*` and `ExternalWarningLevel` are not sufficient for C4710/C4711 because these warnings are produced during optimization of our translation units even when the file/line points into a header.
- Any suppression beyond this shared optimizer-noise baseline must follow the suppression review routine above.

