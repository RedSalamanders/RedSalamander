# Compiler Warnings Specification

## Overview

RedSalamander builds with **`/Wall`** (`<WarningLevel>EnableAllWarnings</WarningLevel>`) to keep a high bar for code quality.

Some MSVC warnings are **optimizer/inlining heuristics** that become extremely noisy under `/Wall` (especially in template-heavy code paths) and do not represent correctness issues. This spec documents which ones we suppress and why.

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

## Rationale

We suppress C4710/C4711 because:
- They are **not correctness warnings**.
- They are heavily dependent on optimizer heuristics and compiler version, so they are **unstable** across toolset updates.
- They are disproportionately noisy with `/Wall` and templates, drowning out actionable warnings.
- When performance matters, the correct workflow is to **profile first** and then make targeted changes (algorithm/data layout/`__forceinline` where justified), not to chase bulk inlining warnings.

## Project File Policy

All C++ projects keep `/Wall` enabled and disable these warnings via MSBuild:

```xml
<DisableSpecificWarnings>4710;4711;%(DisableSpecificWarnings)</DisableSpecificWarnings>
```

Notes:
- We keep `%(DisableSpecificWarnings)` to preserve any existing warning suppressions from imported property sheets.
- `/external:*` and `ExternalWarningLevel` are not sufficient for C4710/C4711 because these warnings are produced during optimization of our translation units even when the file/line points into a header.

