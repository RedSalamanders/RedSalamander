# I59 libghostty-vt 1.0 migration, ABI review, and terminal opportunities

Status: **DECISION — proposal for owner review; implementation has not started.**
Priority: P1 before the next Terminal runtime promotion.
Owner: Terminal runtime maintainer. Requested reviewer: Eric.
Planned: 2026-11-07 at RedSalamander `7e7c4e8c68f5b02de7592aeb0e52f6607ce2be81`,
branch `ej/review-remediation-20261031`; the checkout was clean before this document.

Integration update, 2026-11-07: renumbered from I58 to I59 because current
`main` already assigns I58 to planning-lifecycle reconciliation. The integration
base is `c60bee829d6a325c2019f664c32bb797f7ad194c`; its changes since the
planned-at commit do not change the reviewed Terminal source, runtime lock,
overlay, or owning Terminal contract. The original review baseline is retained.

This is a non-normative implementation proposal. The request authorizes this
document and its static review, not a dependency change, feature implementation,
build, release, or publication. Current requirements remain in the owning specs.

## Reading guide

- [1. Outcome and important findings](#1-outcome-and-important-findings)
- [2. Sources, identities, and review limits](#2-sources-identities-and-review-limits)
- [3. Existing architecture and requirements to preserve](#3-existing-architecture-and-requirements-to-preserve)
- [4. ABI migration review](#4-abi-migration-review)
- [5. What we patch today, why, and the 1.0 disposition](#5-what-we-patch-today-why-and-the-10-disposition)
- [6. Ordered implementation work](#6-ordered-implementation-work)
- [7. Unused interface and possible terminal improvements](#7-unused-interface-and-possible-terminal-improvements)
- [8. Validation and acceptance matrix](#8-validation-and-acceptance-matrix)
- [9. Risks, stop conditions, and rollback](#9-risks-stop-conditions-and-rollback)
- [10. Decisions requested during document review](#10-decisions-requested-during-document-review)
- [11. Completion checklist and review evidence](#11-completion-checklist-and-review-evidence)
- [Appendix A: every target function](#appendix-a-complete-target-function-inventory)
- [Appendix B: every incumbent loader requirement](#appendix-b-every-incumbent-loader-requirement)
- [Appendix C: all function declaration changes](#appendix-c-complete-function-declaration-changes)
- [Appendix D: all enum changes](#appendix-d-complete-enum-change-review)
- [Appendix E: all public typedefs and callbacks](#appendix-e-complete-public-typedef-and-aggregate-review)
- [Appendix F: public macro changes](#appendix-f-public-macro-changes)
- [Appendix G: complete header coverage](#appendix-g-complete-header-coverage)
- [Appendix H: option/data selector adoption](#appendix-h-option-and-data-selector-adoption)

## 1. Outcome and important findings

Migrate the embedded and floating Terminal plugin from the pre-stable
libghostty-vt source in **Ghostty application 1.4.0** to the independently
released **libghostty-vt 1.0.0**, tag `vt-1.0.0`. Preserve today's terminal
behavior, resource bounds, security policy, performance envelope, plugin ABI,
and portable packaging. Optional experience improvements are a separate backlog
in section 7; adopting 1.0 must not silently implement them.

This is a source/API/ABI migration, not a version-string or DLL replacement.

| Finding | Evidence and consequence | Required response |
|---|---|---|
| Private option collision | Our overlay assigns placement limits to terminal option **48**. Upstream 1.0 assigns **48** to `RESIZE_WINDOW`; upstream inserted `INVALID=0` and shifted all 48 previous terminal options. | Rederive the overlay. Subject to the frozen target remaining unchanged, allocate placement limits **49** and Kitty clipboard-write enable **50**. Change C and Zig together and prove the ranges do not overlap. |
| Four currently required symbols disappear | Selection formatting and three encoder setters in our 71-symbol loader no longer exist under their old names. | Replace the encoder exports and rebuild selection copy around the formatter API. A header-only change cannot fix runtime lookup. |
| Data layout changes behind unchanged names | `GhosttyGridRef`, `GhosttySelection`, `GhosttyStyle`, allocator vtables, and several callback payloads change layout. | Recompile every consumer; remove obsolete `.size` initialization only from types that lost it; update renderer and capture canonicalization. Validate native layouts on x64 and ARM64. |
| Kitty pixel getter changes output type | Image data key 7 changes from a pointer output to `GhosttyString`; key 8 becomes expected pixel length. | Update the batched getter's output storage and pending-image handling. Reusing the old pointer/length addresses is unsafe even though the generic function prototype is unchanged. |
| Existing update observation reports the application version | Candidate/status tooling reads `.version` from `build.zig.zon`, rather than the independent library release. | Carry explicit library release identity through the lock, descriptors, status, runtime identity, and workflows. Do not compare app 1.4 with library 1.0 as one version sequence. |
| Stable symbols do not prove enabled features | 1.0 preserves exports for feature-disabled builds and can return `GHOSTTY_UNSUPPORTED`. | Validate required build features and handle unsupported optional capabilities explicitly. Keep exact plugin/runtime pairing. |

These are **confirmed static migration incompatibilities**, not claims that the
currently pinned product is broken: its public terminal range ends at 47, so
its current private allocation is valid for its current source.
ABI/source adaptation is L effort (multiple days including native validation),
HIGH fix risk because typed `void*` getters can corrupt memory when their output
storage is wrong, and HIGH confidence at the header/source level. The build
patch's removability and performance outcomes remain unverified.

The findings above have the following source anchors and fix estimates. They
describe future migration work, not reproduced defects in the current product.

| ID / priority | Exact evidence | Effort / fix risk / confidence |
|---|---|---|
| M1 / P0 admission prerequisite | `Tools/TerminalEngine/Patches/Ghostty-WindowsPrivateRuntime.patch:68`; target `include/ghostty/vt/terminal.h:2653` assigns `RESIZE_WINDOW=48` | M / HIGH: wrong generic option interprets a different input type / HIGH static confidence |
| M2 / P0 load prerequisite | `Plugins/Terminal/TerminalRuntimeLoader.cpp:470-537`; all four removed required names are mapped in Appendix B and the target headers | M / MED: load table, formatter ownership, event handling / HIGH static confidence |
| M3 / P0 native ABI prerequisite | Target `include/ghostty/vt/grid_ref.h:109`, `style.h` and selection/callback definitions; current renderer fields at `Plugins/Terminal/TerminalVt.cpp:2037` | L / HIGH: unchanged type names with changed storage / HIGH static confidence, native layout unverified |
| M4 / P0 image correctness prerequisite | `Plugins/Terminal/TerminalVt.cpp:1415`; target `include/ghostty/vt/kitty_graphics.h:355` changes key-7 output to `GhosttyString` | M / HIGH: old getter output storage is too small / HIGH static confidence |
| M5 / P1 release identity | `Tools/TerminalEngine/GhosttyCandidateReview.psm1:407-410`; `GhosttyRuntimeStatus.psm1:203`; current lock schema has no distinct library release identity | M/L / MED: schema, receipt and cache compatibility / HIGH static confidence |
| M6 / P1 capability admission | `Plugins/Terminal/TerminalRuntimeLoader.cpp:548-551`; target `include/ghostty/vt/build_info.h:153` and subsequent feature selectors | M / MED: required feature checks and rollback / HIGH static confidence |

Fix estimates overlap; they must not be added as an elapsed-time commitment.

## 2. Sources, identities, and review limits

### 2.1 Exact current and target inputs

| Item | Current product | Reviewed target |
|---|---|---|
| Product being versioned | Ghostty application source containing libghostty-vt | Independent libghostty-vt release |
| Version/tag | Application `v1.4.0` | Library `1.0.0`, tag `vt-1.0.0` |
| Commit | `83fb95624b0439ee660c2679b43b38c8b63769cd` | `0fa3d577a5614150ff8e2c152de5fb9d5c6d2dec` |
| Tree | `4dbd485a3b72c6ced65ad7ea447644a5967e382e` | `f46e4660b796da41ab5bda9501fc51afc90d264d` |
| Annotated release tag object | Existing application release | `6b38d55f04e66ebecc472d8be19946edcc6d2e9d` |
| Zig | Locked 0.17.0, native archive per build-host architecture | `minimum_zig_version` remains 0.17.0; keep exact existing archives unless a separately justified issue requires a change |
| Native public functions | 205 declarations; locked DLL exports also include two CRT helpers, for 207 total | 209 names in upstream's committed native symbol baseline; actual Windows exports still require inspection |
| Public headers | 35 existing headers examined | 36 headers, including new `version.h` |

Current raw lock SHA-256:
`dfd2d6bdf1b9006615b25e9ee75a50320da521ed2c252653d5febbd489b8b147`.
Current normalized overlay SHA-256:
`995cc2b9d39907dfb4b6dac179e844b3f7a7c4279a81df502a693ff7834c8a1d`.

The proposed source is the [official 1.0.0 source tarball](https://release.files.ghostty.org/libghostty-vt/1.0.0/libghostty-vt-1.0.0.tar.gz),
with its matching [dependencies tarball](https://release.files.ghostty.org/libghostty-vt/1.0.0/libghostty-vt-1.0.0-deps.tar.gz).
Published [SHA256SUMS](https://release.files.ghostty.org/libghostty-vt/1.0.0/SHA256SUMS):

| Artifact | Published SHA-256 | This review |
|---|---|---|
| `libghostty-vt-1.0.0.tar.gz` | `4a6ef758be721923e1969e7b81b443dcc92d8dd3fb5a39600c03a8817a753676` | Downloaded into memory; all 4,487,852 bytes hash to this value; contents inspected without extraction or execution. |
| `libghostty-vt-1.0.0-deps.tar.gz` | `4d5ea5e448e4c7ea8fb55ccbaa8668f2a97d44f438cee831e2b7ecd162d85c29` | Checksum read; archive not downloaded, inspected, or authenticated in this review. |

GitHub's tag and commit metadata both report signature verification `valid`.
This is an observed hosting-service assertion, not local signature verification.
The source tarball's `.minisig`, signed checksum list, dependency signature, and
independent repeat download still require verification at the provenance gate.
The three most consequential target headers (`terminal.h`, `style.h`,
`kitty_graphics.h`) and `build.zig.zon` match their immutable commit files exactly.

### 2.2 Primary upstream references

- [Release announcement, November 6](https://mitchellh.com/writing/libghostty-1-0): independent stable library release.
- [Immutable release source](https://github.com/ghostty-org/ghostty/tree/0fa3d577a5614150ff8e2c152de5fb9d5c6d2dec).
- [Version-matched packaging contract](https://github.com/ghostty-org/ghostty/blob/0fa3d577a5614150ff8e2c152de5fb9d5c6d2dec/PACKAGING-LIBGHOSTTY.md).
- [1.0.0 API reference](https://lib.ghostty.org/vt/1.0.0/) and its [stability contract](https://lib.ghostty.org/vt/1.0.0/stability.html).
- [Native public symbol baseline](https://github.com/ghostty-org/ghostty/blob/0fa3d577a5614150ff8e2c152de5fb9d5c6d2dec/src/terminal/c/abi/symbols.txt)
  and [ABI checker](https://github.com/ghostty-org/ghostty/blob/0fa3d577a5614150ff8e2c152de5fb9d5c6d2dec/src/terminal/c/abi/check.zig).

Live `main` already advertises `1.1.0-dev`; it is not the migration target.
The compare API reports the target 266 commits ahead of our pin, with no
divergence. Its file list hit GitHub's 300-file limit. That response is **not**
a complete changed-path review; execution must derive the full path set from
the authenticated Git objects and inspect every path in the governed VT scope.

### 2.3 What this review establishes

All target public headers, native symbol names, current required loader names,
normalized C function declarations, aggregate definitions, enum members and
values, callback typedefs, and public macros were compared. Implicit old enum
values were accounted for: making key-code values explicit does not itself
renumber those keys. Generic getter/setter documentation was inspected for the
changed types that our consumer uses. The current seven-file overlay and its
host consumers were read.

This is static source review. It does not establish MSVC/Zig binary layout,
actual exports/imports, symbol calling behavior, absence of parser regressions,
native ARM64 operation, package integrity, signature authentication, or paired
performance. No build, native test, fixture-producing test, desktop capture,
dependency install, lock edit, workflow dispatch, commit, or push was performed.
The appendices are a review inventory, not a substitute for native ABI probes.

### 2.4 Drift and other plan owners

Before execution, record dirty and untracked work, re-open the cited sources,
and compare the actual affected path history:

```powershell
git status --short --untracked-files=all
git diff 7e7c4e8c68f5b02de7592aeb0e52f6607ce2be81 -- External/TerminalEngine Plugins/Terminal Common/PlugInterfaces/Terminal.h Tools/TerminalEngine Tools/Build-GhosttyTerminalRuntime.ps1 Tools/Get-GhosttyTerminalRuntimeStatus.ps1 Tools/New-GhosttyTerminalRuntimeCandidateReview.ps1 Tools/New-GhosttyTerminalRuntimeCandidateOverlayReview.ps1 Specs/Terminal RuntimeDependencies.props Directory.Build.targets .github/workflows
git log --oneline 7e7c4e8c68f5b02de7592aeb0e52f6607ce2be81..HEAD -- External/TerminalEngine Plugins/Terminal Tools/TerminalEngine Specs/Terminal
```

The public `ITerminal` boundary is declared in
`Common/PlugInterfaces/Terminal.h`; do not infer a new ABI from this migration.
Recompute lock and overlay hashes.
If the pin or source changes, refresh the inventory before reusing conclusions.

Coordinate with [I51](Terminal_Ghostty140_Integration_2026-10-28.md), which retains
its 1.4 qualification residuals, and [I15](Terminal_GhosttyUpgradeReviewRemediation_2026-08-28.md),
which retains older paired-evidence remediation. Do not close either using this
proposal. [I56](CI_HostedAcceptance_2026-10-31.md) owns shared hosted acceptance;
this plan owns the new library's source-bound native acceptance. Historical
permissions for timed desktop work are not fresh permission for a future run.

## 3. Existing architecture and requirements to preserve

The public boundary is RedSalamander's `ITerminal` plugin interface.
`Terminal.dll` owns the child HWND, ConPTY session, input handling, Direct2D/
DirectWrite renderer, UI Automation snapshot, and all Ghostty objects. Ghostty
types remain private to the plugin; do not expose them in host headers/settings.
The plugin privately loads `Plugins/TerminalRuntime/ghostty-vt.dll` after exact
PE machine, size, and raw hash checks. The generated pairing header, runtime
identity, DLL, and notices travel together.

Owning contracts and entrypoints:

- [Embedded Terminal](../../Terminal/Terminal_EmbeddedPlugin.md), especially
  dependency identity, upgrade admission, clipboard, placement, and VT corpus rules.
- [Engine decision](../../Terminal/TerminalEngineDecision.md): private runtime
  choice and maintenance constraints.
- [Runtime lock](../../../External/TerminalEngine/GhosttyRuntimeLock.v1.json),
  [loader](../../../Plugins/Terminal/TerminalRuntimeLoader.cpp), and
  [loader declarations](../../../Plugins/Terminal/TerminalRuntimeLoader.h).
- [Runtime tooling](../../../Tools/TerminalEngine/README.md),
  [tooling governance](../../Testing/Testing_ToolingGovernance.md),
  [validation evidence](../../Testing/Testing_ValidationEvidence.md),
  [selftests](../../Testing/Testing_SelfTests.md),
  [portable ZIP](../../Installer/Installer_PortableZip.md), and
  [unified CI](../../Build/Build_UnifiedCI.md).

Preserve these invariants throughout the upgrade:

1. No system-library search, downloaded prebuilt DLL substitution, public
   Ghostty plugin, or independent runtime update. Stable upstream ABI does not
   waive our private extensions or exact binary pairing.
2. No synchronous ConPTY write or unbounded formatting/parse/decode on the UI
   thread. Preserve the terminal mutex, reader/worker ownership, generation
   checks, bounded queues, and nonblocking UI close.
3. Stop producers, cancel/join owned work, drain typed posted payloads, release
   all Ghostty handles and callback users, then unload. Borrowed pointers must
   not outlive the protected engine state.
4. Clipboard reads remain disabled. OSC 52/1337 write policy remains
   deny/ask/allow with one synchronous engine reply and asynchronous user work.
   Kitty OSC 5522 writes remain rejected before buffering. The decoded write
   cap remains 256 KiB.
5. Kitty external file, temp-file, and shared-memory media stay disabled.
   Preserve image/APC/pixel bounds, 4,096 retained placements, and 1 MiB of exact
   placement-map allocation; display limits alone do not bound engine retention.
6. Existing copy/selection, search, mouse, key/IME, hyperlink confirmation,
   scrollbar, shell/path authentication, pane routing, live tabs, accessibility,
   and theme behavior remain acceptance requirements.
7. Keep baseline CPU policy and SIMD disabled for this migration. Evaluating
   SIMD, feature pruning, a GNU/MSVC target switch, or debug-symbol shipment is
   a separately measured decision, not an automatic consequence of packaging docs.

## 4. ABI migration review

### 4.1 Compatibility boundary and complete inventory

1.0 stability applies to upstream's 1.x public ABI. It does not make our older
pre-stable headers compatible with 1.0, nor stabilize our private enum additions.
Windows also has no ELF SONAME check to protect this consumer. Keep exact-pair
loading and add a descriptive library-version/feature check.

The comparison found **10 added, 6 removed, and 30 changed textual function
declarations**. Three changes only rename an output parameter. The other changes
require signature or source-contract review. Of the 71 functions our loader
requires, **4 are removed** and **16 have changed declaration text** (one of the
16 is only an output-parameter rename). Appendices distinguish those cases.

The target has 209 native public symbols and five additional Wasm-only helper
declarations. Do not compare 214 header declarations with a Windows DLL and call
the five Wasm omissions failures. Conversely, the old two CRT export helpers
are not public terminal capabilities. Recompute actual export count and digest
after the candidate build; neither 207 nor 209 alone is the new DLL allowlist.

### 4.2 Required symbol and signature adaptations

| Area | Current usage | 1.0 change and implementation work |
|---|---|---|
| Keyboard setup | `ghostty_key_encoder_setopt_from_terminal` | Resolve `ghostty_key_encoder_set_from_terminal`; retain terminal-mode refresh before encoding. The unused generic `setopt` also becomes `set` if later used. |
| Mouse setup | `ghostty_mouse_encoder_setopt`, `ghostty_mouse_encoder_setopt_from_terminal` | Resolve `ghostty_mouse_encoder_set` and `_set_from_terminal`; use new enum values and check results. |
| Selection copy | `ghostty_terminal_selection_format_buf` with `GhosttyTerminalSelectionFormatOptions` | Both selection-format functions/options disappear. Configure `ghostty_formatter_set` with plain output, unwrap/trim, selection, and explicit extras; use buffer formatting. Reset selection options afterward or use a separate owned formatter so Find never inherits a copy selection. Keep two-pass required-size handling and byte caps. |
| Formatter creation | Constructor takes options by value | New constructor takes allocator, output handle, and terminal only; apply every required option with checked setters. Add setter-failure rollback stages and destructor coverage. Preserve current Find snapshot trimming and wrapping. |
| Key and mouse byte output; paste encoding | Caller-owned `char` buffers | Encoders now accept `uint8_t*`. Update buffer storage/spans at all call sites, including tests and overflow retries, without introducing extra copies or changing UTF-8 ownership. |
| Event setters | Most return `void` | Setters now return `GhosttyResult`; validate each result before encoding. Do not send partially updated reusable events after failure. |
| Key text | `set_utf8(event, const char*, length)` | Now accepts `GhosttyString`; getter also returns `GhosttyString`. Keep backing bytes alive through encoding and test empty/non-ASCII text. |
| Viewport scroll | Returns `void` | Returns `GhosttyResult`; handle failure in wheel, scrollbar, Enter-follow, and test code. Scroll tags are renamed. |
| Other public functions | Not currently resolved | Focus/color-scheme/mode/size-report encoders, OSC result handling, `select_output`, and SGR attribute access also change. Record them in the inventory so future adoption uses 1.0 declarations. |

`TerminalRuntimeLoader.h/.cpp` must change together: `decltype` declarations,
all required lookups, new entries, failed-load cleanup, `Unload()` nulling, and
the source-contract expected symbol set. A build succeeding against new headers
does not prove that runtime string lookup resolves the same names.

### 4.3 Layout and initialization changes

| Type/family | Verified source change | Consumer obligation |
|---|---|---|
| `GhosttyGridRef` | Removes leading `size` | Remove obsolete assignments; check tracking/selection, hyperlink lookup, snapshots, and copied references under resize/pruning. |
| `GhosttySelection` | Removes leading `size`; embeds the changed grid refs | Update character/rectangle/word/all selection, format configuration, anchor lifetimes, and teardown. Never reuse old byte layouts. |
| `GhosttyStyleColor` | Tag becomes `uint8_t` instead of enum-sized storage | Check union/tag decoding and native offsets. |
| `GhosttyStyleColorValue` | Removes the old `uint64_t _padding` union member | Do not assume the old union size/alignment survives; verify the containing color/style layouts. |
| `GhosttyStyle` | Removes `size` and boolean fields; adds `GhosttyStyleFlags` plus reserved storage | Replace every bold/italic/faint/blink/inverse/invisible/strike/overline read with the named flags and use `GHOSTTY_STYLE_UNDERLINE`. Update render and test snapshot logic together, preserving semantic hashes. |
| Cursor/layer types | Unifies into `GhosttyCursorStyle` and `GhosttyImageLayer` | Replace removed type/enumerator names. Keep layer ordering and background/text/image composition unchanged. |
| Notification/progress/program-status/semantic-prompt payloads | Remove leading `size` | No callbacks are currently registered for these options. Any later adopter must use new payload layouts and copy borrowed strings before returning. |
| `GhosttyAllocatorVtable` | Four named callback typedefs and four reserved pointers | Zero reserved storage if a custom allocator is introduced; test WIC callback allocations against the engine allocator. Current default allocation is not permission to change allocator policy. |
| Formatter options/extras structs | Removed; replaced with enum-keyed options and a `uint32_t` extras mask | Avoid compatibility copies of deleted types. Explicitly set output format, wrapping, trim, selection, and extras. |
| Sized types that remain | New helper `GHOSTTY_INIT_SIZED`; native runtime helper also available | Use the macro for actual sized types; it adds no runtime dependency in C++. Do not apply it to types whose `size` field was removed. |

Appendices enumerate every changed aggregate and every enum/macro delta, including
unused interface. Verify `sizeof`, `alignof`, `offsetof`, enum width/signedness,
callback conventions, pointer/`size_t` widths, and flag masks using MSVC and the
candidate runtime on **both** native architectures. Upstream's committed type
goldens cover Linux x64, macOS ARM64, and Wasm, not our two Windows layouts.
Use its `ghostty_type_json()` diagnostic manifest as input where applicable;
do not copy a Linux layout or claim cross-compilation proves native execution.

### 4.4 Typed generic getters, enums, and callbacks

Function-pointer types alone cannot validate enum-keyed `void*` output storage.
Review every used option/data key with its candidate documentation and Zig
`InType`/`OutType`, including each element of all `*_get_multi` arrays.

- Terminal options move from 0–47 to 1–48; `INVALID=0` is new. Encoder,
  system, search, selection-gesture, OSC, and snapshot option/data enums also
  gain invalid sentinels or other naming/value changes. Never preserve numeric
  casts or hard-coded enum aliases through the migration.
- `GHOSTTY_RENDER_STATE_OPTION_*` becomes `..._OPT_*`; row options follow the
  same convention. Cursor and scroll constants change names. Named modes are
  renamed for readability; their encoded values must be checked against the
  appendix rather than inferred from names.
- Kitty image `...DATA_PIXELS` returns a **`GhosttyString`**, and
  `...DATA_EXPECTED_LEN` returns a **`size_t`**. An incomplete transmission may
  report expected storage before `PIXELS` has a value. Retrieve valid storage
  types, check per-element success/failure and bounds, then copy under the
  engine lock before asynchronous conversion. Do not retain engine pixel pointers.
- Clipboard, PTY, PNG, and other existing named callback function typedefs have
  unchanged declaration shapes in this comparison. That does not preserve the
  layout of their pointees, allocator vtables, callback lifetime, or reentrancy
  rules. Preserve synchronous clipboard replies and shared PNG callback-user
  ownership through the last runtime instance.
- Treat `GHOSTTY_UNSUPPORTED` separately from `NO_VALUE`, buffer growth, and
  actual failure. Missing required renderer/input/formatter/selection/grid/Kitty
  capability must fail initialization; optional feature UI must not appear usable
  when its implementation is compiled out.

### 4.5 Version and capability admission

Recommended runtime admission is still exact raw hash + PE + size pairing,
followed by resolving the required 1.0 symbols, checking `ghostty_version()` and
the numeric build-info version, and checking required feature flags. Check for
the reviewed 1.0.0 identity, including release/pre-release status, rather than
accepting an arbitrary 1.x DLL. Future 1.x upgrades still use this review process.

Make library release identity explicit in machine contracts. The recommended
shape is a typed library-release identity (`product=libghostty-vt`,
`version=1.0.0`, `tag=vt-1.0.0`) alongside the existing repository commit/tree
and source-version metadata. Do not silently reinterpret an application's
`upstream.version` field. Version any incompatible lock/identity/descriptor
schema change, retain explicit historical readers as needed, and reject old
evidence as admission for new inputs. Section 6 assigns the full consumer work.

## 5. What we patch today, why, and the 1.0 disposition

The single tracked [Windows/private-runtime patch](../../../Tools/TerminalEngine/Patches/Ghostty-WindowsPrivateRuntime.patch)
has **five concerns across seven upstream files**, recorded as 372 insertions
and 27 deletions in the current contract. It is a private derivative, not an
unmodified upstream 1.0 DLL. Patches belong in this one reviewed overlay; do not
grow an independent runtime fork or duplicate changes in download/build scripts.

### 5.1 Build/link policy — `build.windows.reproducible-link`

**What:** in `src/build/GhosttyLibVt.zig`, the Windows library forces LLVM,
LLD, and stripping at the library root module. The existing MSVC-related UBSan
handling around that insertion is upstream code, not our addition.

**Why:** the private product does not currently ship a Ghostty PDB. This pins
the intended linker/debug-information behavior and keeps the production layout
consistent with our verification and size envelopes. It does **not** guarantee
identical DLL bytes across clean builds; raw hash pairing records each real build.

**1.0 disposition:** investigate removing the hunk using supported build flags
only if clean x64 and ARM64 builds prove equivalent behavior. 1.0's Windows
branch still does not contain our three forced settings; LLVM/LLD selection in
its ELF branch does not satisfy Windows. Compare actual strip/PDB/output/import
behavior. If flags do not reach the library root, retain the minimal hunk.
Do not switch Windows target ABI or start shipping PDBs as an incidental fix.

**Proof required:** clean and warm cache builds, offline forced rebuilds,
header/runtime identity, actual exports/imports/machine/size, plugin load/unload,
and documented disposition with no claim of bit-for-bit reproducibility.

### 5.2 Kitty clipboard protocol gate — `clipboard.kitty-write-disable`

**What:** `terminal.h` and `src/terminal/c/terminal.zig` add a boolean private
option. `stream_terminal.zig` checks it before buffering a Kitty OSC 5522 write
and responds with `ENOSYS` when disabled. Its runtime default is enabled;
RedSalamander explicitly disables it before admitting any input. OSC 52 and
1337 use the normal host policy and remain supported.

**Why:** Kitty write completion expects a synchronous truthful result. The host
may need asynchronous user consent and clipboard I/O. Some Kitty requests have
metadata indistinguishable from legacy writes at the current callback, so
rejecting them later in the host cannot implement this protocol policy reliably.
The public 256 KiB write cap limits decoded data; it does not replace the gate.

**1.0 disposition:** retain the concern unless a public 1.0 mechanism can reject
this exact protocol before buffering while preserving legacy writes and reply
timing. No equivalent boolean option appears in 1.0's complete public terminal
range. Rederive it as private option **50**, subject to the frozen-range check.
Keep public read callback null and the decoded-data cap as independent controls.

**Proof required:** disabled transaction never buffers or calls the host; one
correct rejection response; legacy writes still reach the callback; null/reset
semantics; initialization fails if the gate/cap/read-policy setup fails; no
inner consent/retry loop; cancellation/teardown with a queued host prompt.

### 5.3 Exact placement accounting — `kitty.placement-accounting`

**What:** `src/terminal/kitty/graphics_storage.zig` computes the retained
placement hash map's allocated bytes, including header, metadata, keys, values,
alignment, and allocated capacity rather than logical entry size alone.

**Why:** a small number of large-capacity maps can retain more memory than a
`count * sizeof(value)` estimate shows. We need exact engine retention evidence,
not merely the number of placements the renderer chooses to display.

**1.0 disposition:** retain or replace with an equivalent upstream accounting
primitive only after checking the target's map representation and Zig 0.17's
allocator/hash-map layout. The source review did not find the private accounting
methods in 1.0. Revalidate the mirrored map-header and capacity calculation;
never carry the formula forward on source resemblance alone.

**Proof required:** allocator-observed allocated bytes match exposed accounting
across growth boundaries and both architectures; empty/retained-capacity states;
rejection leaves allocation unchanged; no overflow in capacity arithmetic.

### 5.4 Admission and memory limits — `kitty.placement-capacity`

**What:** the same storage implementation adds hard defaults of 4,096 entries
and 1 MiB; validates count/allocation admission before mutation; reserves before
garbage collection; preserves explicit replacement and pin ownership; and keeps
configured limits when image storage is disabled/reset. Terminal option handling
validates all initialized screens before applying a limit change to any screen.

**Why:** image-byte limits do not independently bound placement metadata. A
failed allocation/admission must not partially mutate IDs, generations, counts,
map allocation, or tracked pins. The caller retains a rejected placement/pin.

**1.0 disposition:** retain this product requirement. Upstream 1.0's placement
implementation has changed, including parent/relative-placement handling; port
the invariant into that implementation rather than replacing its newer logic
with the old patched function. Confirm regular, virtual, relative, and alternate
screen paths cannot bypass the same retention bound.

**Proof required:** below/at/above count and byte limits; exact-ID replacement
at capacity; allocator failure; rejected caller pin; generation/internal-ID
stability; pruned-parent and relative-placement behavior; deletion/reset;
two-screen all-or-nothing limit changes; long-running retention/teardown.

### 5.5 Private observation/control ABI — `kitty.private-placement-abi`

**What:** `terminal.h` adds the sized `GhosttyKittyPlacementLimits` structure
and private terminal option. `kitty_graphics.h` and both C binding Zig files
expose count, allocated bytes, count limit, and byte limit through Kitty data
keys 3–6. C/Zig compile-time guards protect the allocated numbers.

**Why:** the host must lower/verify the limits and record actual engine
retention. Renderer-local counters cannot prove the runtime obeys its bounds.

**1.0 disposition:** rederive placement option as **49**; keep Kitty data keys
**3–6** only after verifying upstream's entire Kitty data range still ends at
2. These are separate enum spaces: terminal option 49 says nothing about Kitty
data key 3. Recheck the sized private struct against the Windows native ABI and
keep exact-pair admission. Add guards that would reject a future upstream collision.

**Proof required:** candidate C/Zig equality and uniqueness, old number-48
collision rejection, wrong-size/zero/over-hard-limit requests, valid lower limits,
all four readbacks on both screens, and the actual production capture path.

### 5.6 Things we do in the host but do not patch upstream

- ConPTY, Direct2D/DirectWrite, WIC PNG decoding callback, pixel conversion,
  upload/cache scheduling, and UI Automation are our plugin integration.
- Clipboard consent, unsafe paste, hyperlink confirmation, and authenticated
  shell/path handling are host policy. The overlay does not implement their UI.
- Runtime PE/import/export validation, raw hash pairing, notice staging, native
  Zig selection, leases, and child-process containment live in build/tooling code.
- Current public upstream support for clipboard read/write limits, Kitty image
  limits/media restrictions, hyperlink URI lookup, input encoding, and render
  snapshots is not a reason to describe those APIs as private patches.
- Historical ARM64 test-runner transport adaptations are qualification machinery,
  not shipped runtime changes. Re-evaluate whether they are still required and
  prove removal afterward; do not port obsolete workarounds automatically.

The overlay review must map every old concern to **retain**, **replace with
public equivalent**, or **remove with proof**, and state affected files, reason,
verification, and maintenance owner. Preserve the existing cap of five concerns,
16 upstream files, and 2,500 total changed lines. Exceeding a cap requires a
separate decision; renaming or merging concerns must not hide extra behavior.

## 6. Ordered implementation work

Each phase delivers code, meaningful coverage, and current owning specs together.
Unchecked work is intentional: this request is for review of the proposal.

### P0 — Freeze the incumbent and the accepted migration scope

- [ ] Recheck section 2 drift; coordinate I51/I15/I56; create an isolated `ej/`
      work branch/worktree when implementation is authorized.
- [ ] Freeze current source, lock, overlay, runtime identity, full dependency/
      notice closure, build receipts, DLLs, and a rollback manifest before a build
      can replace any output. Do not terminate independently launched programs.
- [ ] Capture the current source-bound, test-enabled Release VT corpus and its
      control hashes **before** adapting host code. Preserve existing failed runs.
- [ ] Confirm the actual supported profiles/packaging gates and select already
      initialized, marked `X:\RedSalamander.Perf` roots. Do not initialize a new
      drive-root path or infer authorization from historical E:/C: evidence.
- [ ] Freeze upgrade-only scope; optional section-7 features stay out. Freeze the
      corpus, semantic serialization, and performance limits before measuring.

Exit: reproducible baseline and rollback packet; scope and evidence identities
recorded. No canonical pin change.

### P1 — Authenticate and model the independent release

- [ ] Resolve `vt-1.0.0` tag/object/commit/tree and ancestry locally; verify release
      signatures using the trusted upstream signing key through the governed
      archive/provenance workflow, including signed checksum and dependency archive.
- [ ] Repeat independent source downloads; compare hashes. Preserve both the
      pristine authenticated Git source needed by candidate review and the
      official pruned release tarball; explain their intentional packaging differences.
- [ ] Add explicit independent-library release/version identity to the lock and
      each bound descriptor/status/runtime identity schema and validator. Retain
      the repository commit/tree and source version separately. Decide schema
      revisions and historical readers explicitly; reject mixed old/new evidence.
- [ ] Extend source review scope to packaging/version/ABI inputs newly relevant to
      1.0: `LIB_VERSION`, `PACKAGING-LIBGHOSTTY.md`, `deps/libghostty-vt*`,
      `src/build/Config.zig`, `src/terminal/build_options.zig`,
      `src/terminal/c/abi/**`, version/header generation, and their imported build
      dependencies, in addition to the existing governed VT scope. Derive the
      actual closure; do not rely on the truncated compare API file list.
- [ ] Adapt candidate-freeze and overlay-review provenance so verified Git trees
      and release tarballs remain linked without weakening clean-source or
      repeat-archive requirements. Do not pretend a pruned tarball is a Git checkout.
- [ ] Update status observation/release preflight to identify stable `vt-*`
      releases separately from application `v*` tags and development branch head.
      Preserve Locked's zero-network mode, exact Candidate no-substitution,
      OSV commit queries, time bounds, and read-only workflow permissions.
- [ ] Cover wrong release product/tag, app/library version confusion, prerelease,
      moved tag, commit mismatch, invalid signature, stale descriptor, source/deps
      mix, missing `LIB_VERSION`, no-network Locked, and oversized/truncated
      metadata. No automatic lock edit, branch, issue, or PR creation.

Files include the `GhosttyRuntime*` and `GhosttyCandidate*` modules, owning
schemas in `Specs/Terminal`, their entrypoints/tests, and the status/native/
product qualification workflows. Include every changed dependency in cache and
workflow identity keys. Read tooling governance before editing `Tools/`.

Exit: authenticated, exact, library-aware candidate review and descriptor.

### P2 — Reconcile build inputs, dependencies, and notices

- [ ] Recommend building the signed, library-specific source tarball with the
      matching offline dependency set. Keep authenticated Git sources for diff/
      overlay review. If the current builder cannot support the pruned archive,
      implement that support first rather than falling back silently to `main`.
- [ ] Handle the release's `LIB_VERSION=1.0.0` correctly. Its `build.zig`
      fallback still says `1.0.0-dev`; do not edit it into a fake release or read
      only the application's `.version`. Verify the built runtime reports 1.0.0.
- [ ] Inspect actual target build options. Existing lock spelling includes
      `-Doptimize=fast`; 1.0 packaging documents `ReleaseFast`, while the exact Zig
      source uses `standardOptimizeOption`. Verify accepted spelling using the
      pinned 0.17.0 toolchain, record it once in the lock, and require the runtime
      to report ReleaseFast. Do not infer accepted flags from documentation alone.
- [ ] Retain native aarch64 Zig on ARM64 hosts, x86-64 Zig on x64 hosts, baseline
      CPU, SIMD-disabled policy, explicit feature configuration, clean cache
      isolation, bounded contained subprocesses, repository runtime lease, and
      current non-destructive output ownership rules.
- [ ] Reconcile the release's dependency manifest with the **actual** selected
      build graph. It lists Aro, Highway, simdutf, translate_c 2.0.0, uucode 0.2.0,
      and Wuffs; build-only and feature-disabled inputs do not all become shipped
      code. Do not copy the incumbent Aro/Pixels/uucode/translate_c closure or
      its eight notices unchanged.
- [ ] For every reached package and vendored source, record transport/content
      hashes, license identity, purpose, feature condition, and shipped notices.
      Inspect embedded simdutf even if SIMD stays disabled; prove why an unused
      component is excluded. Remove a legacy notice only after proving no
      shipped code still requires it. Keep a complete, ordered lock closure.
- [ ] Verify no new C++ runtime or external DLL import is introduced. Keep import
      policy strict; classify a new import rather than widening its allowlist.
- [ ] Update archive layout/root detection, offline materialization, stale cache
      invalidation, runtime outputs/header locations, PE size envelopes, notices,
      and generated identity as required. Keep candidate mode out of Product.
- [ ] Finish the build-hunk retain/remove proof from section 5.1. Record fresh
      Windows export sets and per-platform size bounds from actual verified builds.

Exit: isolated x64/ARM64 candidate build definitions and complete lawful package
closure; no canonical admission.

### P3 — Rederive and verify the five-concern overlay

- [ ] Port each concern using section 5's invariants and current 1.0 source.
      Preserve upstream parent/relative-placement and decoder fixes.
- [ ] Reallocate terminal private options to 49/50 and revalidate Kitty data 3–6.
      Update C/Zig guards, documented allocations, source-contract fixtures, and
      the negative control for the specific old option-48 collision.
- [ ] Exercise engine-level clipboard and placement behavioral controls before
      relying on production host tests. Check exact map bytes against allocation
      observations, not a duplicate of the accounting formula.
- [ ] Produce pristine/applied worktrees, a single normalized full-index patch,
      concern dispositions, schema-valid overlay review, and immutable P4
      descriptor. Repeated review uses a fresh immutable snapshot when its
      non-overlay inputs changed; never overwrite or alias mutable descriptors.
- [ ] Respect the five-concern/file/line caps and fail closed on drift.

Exit: independently reproducible overlay diff and approved concern ledger.

### P4 — Adapt the plugin, loader, and exact Windows ABI

- [ ] Update required runtime symbol resolution, member signatures, cleanup,
      diagnostic strings/satellites if needed, version/feature admission, and
      required-symbol source contracts.
- [ ] Complete every change in section 4 and all used entries in the appendices.
      Include `Terminal.cpp`'s test-enabled runtime checks, `TerminalVt.cpp`,
      `TerminalSecurity.cpp`, `TerminalInternal.h/.cpp`, and VT corpus code.
- [ ] Rebuild selection formatting with checked formatter configuration; restore
      state after copy and verify Find/UIA snapshots contain the full intended
      document, not the last selection. Preserve privacy and clearing behavior.
- [ ] Convert style field reads/semantic serialization, cursor/layer/scroll/mode
      names, byte buffers, result handling, and Kitty pixel getter output types.
- [ ] Review every used data/option key against C docs and Zig types. Exercise
      wrong-type hazards with valid typed probe storage, not deliberate writes
      through undersized buffers.
- [ ] Extend rollback injection to new formatter/configuration/admission stages.
      Assert no live handles, callback users, queues, workers, or HWND payloads
      remain after each failure and close/reopen loop.
- [ ] Run a real MSVC-to-Zig C ABI probe on native x64 and ARM64. Compare types
      and call behavior, including callbacks, strings/slices, return values,
      private limits, buffer query/retry, and `*_get_multi` error reporting.
- [ ] Do not change `ITerminal` UUID/vtable or settings solely for this upgrade.
      If compatibility cannot be preserved, stop for a separate explicit design.

Exit: new plugin/runtime pair passes native ABI and focused behavioral coverage.

### P5 — Qualify correctness and performance before admission

- [ ] Execute upstream lib-vt C/Zig/example/header tests on native x64 and ARM64
      through the governed runner. Inspect actual 1.0 target names/availability
      before reusing the old three-target workflow. Unsupported upstream ABI
      golden targets remain declared separately from our Windows probe.
- [ ] Run focused Terminal contracts plus the normal-order/predecessor/repeat
      coverage for affected Commands. Preserve readiness, shutdown, timeout,
      identity, desktop-restoration, and no-activation assertions.
- [ ] Rebuild test-enabled Release corpora against the correct old/new headers
      in isolated lanes. Keep the same scenario and semantic canonicalization.
      A temporary compatibility adapter must not change either measured product
      algorithm or silently update expected behavior.
- [ ] Collect exactly 200 raw samples each for VT ingest, snapshot, and hosted
      render on the same machine/profile, along with control digests and all
      existing Kitty/process memory counters. Run the authoritative paired
      validator; require candidate p95 at most `ceil(baseline p95 * 1.10)`.
- [ ] Preserve the currently normative x64 paired gate. The proposed ARM64
      acceptance in D6 also requires native paired measurements: explicitly
      extend the corpus/receipt contract if its current x64-only route cannot
      capture ARM64, then compare old/new ARM64 on the same ARM64 machine with
      the same 200-sample protocol and 1.10 ceiling. Never compare x64 timings
      with ARM64 or count a cross-built executable as a native measurement.
- [ ] Investigate every semantic/PTY/control difference. The existing named
      `dcs-high-byte-preservation` allowance is not a blanket approval for other
      1.0 changes; both current 1.4 and target may already share that correction.
      Any genuinely intended new behavior needs predeclared, reviewed corpus/
      normative-spec changes before a new claim-bearing comparison.
- [ ] Complete the full suite required by the Terminal upgrade contract and
      relevant native profiles; scoped/PrePush/PR passes alone do not close it.
      Record missing interactive time or native equipment as pending gates.

Exit: accepting source-bound correctness, native-platform, and paired performance
evidence, with original failures preserved.

### P6 — Atomic product admission, packaging, and closeout

- [ ] Promote only the exact reviewed source/tool/overlay/dependency tuple.
      Rebuild Product via its canonical builder, generate the pairing header
      before compiling `Terminal.dll`, and verify non-force cache reuse plus
      forced offline rebuild behavior on both architectures.
- [ ] Re-run focused tests against the **admitted Product** bytes; isolated
      candidate evidence alone does not qualify a newly built product pair.
- [ ] Update `RuntimeDependencies.props` for actual added/removed notices or
      outputs. Keep staging/portable packaging driven by that same manifest.
      Verify clean extraction, no stale DLL/notices, exact binary identity, load,
      and terminal lifecycle for production x64/ARM64 ZIPs.
- [ ] Satisfy current unified CI: compiler-free PR checks do not waive local
      tests; both native Release test branches pass before production ZIP jobs.
      Preserve the shared-engine pin and explicit publication controls.
- [ ] Update normative Terminal/engine-decision, build/identity, testing, and
      portable contracts for durable changes; update tooling inventory/help/
      callers/cache closure and a dated change fragment. Never hand-edit the
      generated changelog or fold a feature's own fragment.
- [ ] Keep optional features as reviewed separate work. Reconcile I51/I15/I56
      boundaries; move this record to Done only after acceptance and the exact
      reviewed Done-path admission, updating its WIP index entry accordingly.

Exit: accepted version-1.0 product with a tested rollback tuple and all required
native/runtime/package/performance gates closed. Release publication remains a
separate user action unless explicitly authorized during implementation.

## 7. Unused interface and possible terminal improvements

### 7.1 What “unused” means here

The appendix inventories **every** target native symbol. Of 209 names, 67 still
match our existing required loader names; 142 are not currently loaded under
their 1.0 names. Some of those 142 are mandatory migration replacements, not new
features. Four old required names disappear entirely and are listed separately.
All five Wasm-only helpers are outside the Windows product boundary.

A loaded generic setter/getter does not mean every option is used. The appendix
also maps all terminal options and public data/option enum selectors. For example,
our plugin uses terminal getters for scroll state but registers none of the bell,
notification, progress, program-status, semantic-prompt, resize-window, or
render-hold callbacks at the reviewed source. Conversely, Terminal Find,
selection, history suggestions, bracketed paste, hyperlinks, Kitty still images,
and UI Automation already exist; do not present them as entirely missing.

The opportunities below are **optional design proposals**, not bugs and not
upgrade acceptance requirements. Effort is rough including behavior/perf tests:
S = hours, M = days, L = several days or more. Value/risk must be confirmed before
activation. Every accepted feature needs its own product behavior, shortcuts,
accessibility/localization, performance, privacy, and teardown decisions.

### 7.2 Grounded extension options

| Option | Public capability currently unused | Existing experience and possible improvement | Work, limits, and recommendation |
|---|---|---|---|
| E1 Native search model | `ghostty_search_*`, including next/previous selection, match and viewport getters | Find currently formats a bounded plain snapshot and filters off-thread. Engine search could navigate/highlight matches in scrollback using engine coordinates. | L, MED/HIGH risk. 1.0 search folds ASCII case only; current Find uses Windows ordinal case-insensitive matching. Resolve that Unicode behavior difference before replacement, then compare wrapping, live updates, cancellation, and retention. Keep its UI/privacy contract. Recommended first design spike after upgrade. |
| E2 Rich selection gestures | `ghostty_selection_gesture_*`, line/output selection and selection adjust/order/contains helpers | Character/rectangle/word/all selection already exists. Shared gesture handling could add consistent triple-click line selection, extension, and platform repeat-distance behavior. | M/L, MED risk. Validate Windows double-click timing, drag threshold, autoscroll, high DPI, rectangles, and UIA mapping; avoid two competing gesture state machines. |
| E3 Prompt/output navigation | Semantic-prompt callback, prompt-delta scrolling, `ghostty_terminal_select_output` | Existing authenticated shell integration follows cwd and supplies history/path actions. Optional prompt markers could support previous/next command, selecting command output, and completion status. | L, HIGH risk. Preserve authenticated cwd/path trust; terminal-reported strings/exit states are untrusted content. Add shell/protocol support deliberately and bound retained markers. |
| E4 Progress and completion indicators | Progress/program-status callbacks, desktop notification and bell callbacks | Surface long-running command progress in a terminal tab or passive notification without requiring the user to inspect output. | M/L, MED risk. Define opt-in/coalescing, inactive-tab behavior, rate limits, notification ownership, and privacy. Never let child output activate windows or impersonate trusted app notices. |
| E5 Fewer repeated getter calls | Render/cell/row/terminal `*_get_multi` functions and `row_iterator_next_dirty` | Current renderer traverses the viewport and performs multiple per-cell calls. Batched reads or dirty-row preparation may reduce CPU. | M/L, HIGH rendering risk. Benchmark first; Direct2D redraw/device-loss, selection, cursor blink, theme and Kitty composition can require painting clean text. Preserve `begin_update`/`end_update` snapshot ownership. No promised speedup without paired evidence. |
| E6 Unified render-state image stream | `ghostty_render_state_image_*` iterator/getters and shared image layers | Existing Kitty path walks graphics placement/image objects and maintains a host conversion/upload cache. Render-state images may simplify coherent frame capture and layer ordering. | L, HIGH risk. Prove equivalent placement, clipping, pending/stale generation, memory, virtual/relative placements, and teardown behavior before replacing the working pipeline. |
| E7 Animated Kitty graphics | `ghostty_kitty_graphics_animation_tick` plus generation tracking | Current host loads still-image/placement APIs and does not drive animation ticks. It could support animation without a separate decoder model. | L, HIGH memory/CPU risk. Add an owned timer/scheduler, hidden-tab suspension, frame/cache bounds, deterministic timing, cancellation, and native GPU/perf coverage. Not enabled by upgrading the DLL. |
| E8 Better key composition metadata | Consumed-modifier and composing setters, event getters | Current key encoding already handles UTF-8, modifiers and IME in host code. Additional metadata may improve AltGr/dead-key/Kitty-keyboard edge cases. | M, MED risk. First characterize current layouts and IME; adopt only when it fixes a demonstrated mismatch. Do not claim the present IME path is missing. |
| E9 Protocol-aware focus, theme, and size | Focus/color-scheme/size-report encoders and size/color callbacks | TUI applications could receive correct focus transitions, light/dark answers, and pixel-size reports. | M, MED risk. Gate on terminal-requested modes and actual focus/theme state. Avoid duplicate replies or output on unsupported modes; theme changes must respect explicit application colors and high contrast. |
| E10 Lower inactive-session memory | Compression activity/compress APIs | Many open terminal tabs can retain scrollback. Background compression could reduce private bytes. | L, HIGH latency risk. Measure compression/resume latency, memory, tracked refs, cancellation, UIA/Find snapshot interaction, and active typing. No UI-thread compression. |
| E11 Export/copy richer content | Formatter VT/HTML modes, extras, writer/allocated formatting | Copy currently produces bounded plain text; formatted export could preserve colors/hyperlinks or create an explicit terminal transcript. | M/L, MED/HIGH risk. Explicit user action and destination, safe HTML/URI handling, bounded streaming, clear sensitive content policy. Plain copy remains unchanged. |
| E12 Session snapshots and continuation | Snapshot encoder/decoder and continuation APIs | Could support explicit save/reopen or diagnostic replay of terminal state. | L, HIGH risk. Review decoder limits/version compatibility and shell/process semantics; restoring screen bytes does not restore a trusted live process. No automatic persistence of terminal/clipboard/secrets. Defer until product/privacy design. |
| E13 Unicode/cell geometry tools | Codepoint/grapheme width, direct grid/style/row getters | Could help UIA navigation, selection geometry, or renderer diagnostics agree with engine cell widths. | M/L, MED risk. Reuse engine widths without replacing DirectWrite glyph shaping. Cover combining clusters, wide cells, emoji, RTL limitations, and DPI. |
| E14 Palette utilities | Color parsing/default/generation/luminance/contrast APIs | Could help terminal palette import or readability validation of custom terminal themes. | M, MED risk. Shared theme helpers remain authoritative; do not create a second application color engine or bypass high contrast. Requires a specific user-facing palette workflow first. |
| E15 Host-provided paste integration | `ghostty_terminal_paste` and MIME-aware readers | Current unsafe-paste confirmation and bracketed encoding work. Public paste orchestration could simplify future typed paste/MIME support. | L, HIGH clipboard risk. Preserve consent, size limits, asynchronous ownership, input ordering, and forbidden read policy; never add a second unreviewed clipboard route. |
| E16 Diagnostics and protocol building blocks | Standalone OSC/SGR parsers, system log hook, type manifest, reset/write-until | Useful for bounded developer diagnostics, ABI probes, or explicit transcript processing. | S/M for diagnostic probes, larger for product use. Full terminal ingestion already parses OSC/SGR; do not parse every stream twice or log payloads. `ghostty_type_json` is a strong fit for validation, not a terminal feature. |

For a later selected feature, record its owning files/tests, p95 and retained
memory metrics, host trust boundaries, and teardown proof before implementation.
Suggested review order is E1, E2/E3, then E4; evaluate E5 with profiling. E7,
E10, E11, E12 and E15 deserve separate product decisions rather than inclusion
in the 1.0 migration. Wasm helpers, Unix packaging/SONAMEs, Apple frameworks,
Android distribution, and a Ghostty GUI replacement are out of scope.

## 8. Validation and acceptance matrix

### 8.1 Required evidence by boundary

| Gate | Evidence required | Acceptance / negative controls |
|---|---|---|
| Provenance/release | Verified tag/commit/tree, repeat downloads, signed source/checksum/deps, complete diff scope | Correct `vt-1.0.0` only; invalid signatures, moved refs, app tags, wrong archive/deps and stale descriptors rejected. |
| Header/API inventory | Regenerated function/type/enum/macro/selectors inventory and reviewed used data types | All 71 incumbent requirements disposed; removed names absent from new lookup table; no private/public collisions. |
| Native ABI | MSVC caller plus actual Zig DLL, native x64 and ARM64 | Sizes/alignments/offsets, enums/flags, callback calls, strings, alloc/free ownership, getters, buffer query/retry, new return values. No Windows pass inferred from Linux goldens. |
| Build/runtime identity | Clean, reuse and offline forced builds; actual PE/import/export/size/hash identities | Zero actionable warnings/errors in required first-party profiles; no unknown imports/exports; exact generated pair; mismatched or missing/corrupt DLL fails safely. |
| Feature admission | Version/pre-release and `build_info` feature queries | Required disabled feature fails initialization even when symbol exists; optional unsupported feature stays unavailable. |
| Formatter/selection/render | Real styled/Unicode/selection snapshots, plain copy, Find, UIA, scrollbar/cursor/image output | Deleted layouts gone; new flags preserve appearance; copy cannot narrow subsequent Find/UIA; snapshots show the newest completed generation. |
| Key/mouse/paste | Existing terminal contracts plus failure paths | Application modes, UTF-8/IME, modifier keys, repeats/releases, mouse reports, bracketed paste, unsafe paste, buffer bounds, scroll status. |
| Clipboard policy | Engine gate plus production callback/consent tests | No reads; Kitty rejects before buffering; OSC 52/1337 policy intact; 256 KiB boundaries; one reply; rollback/close with pending prompt. |
| Kitty retention/decoder | Allocator-observed map bytes and real production capture/conversion | Count/byte admission, pin ownership, replacement, relative-parent correctness, pending pixels, max PNG/RGBA, generation/cache invalidation, queue cancellation and teardown. |
| Lifecycle/concurrency | Initialization-stage fault injection, repeated close/reopen, active output/PNG callback, late payloads | All objects/workers/callback users released at quiet point; no deadlock, module use after unload, leaked pin, stale HWND delivery, or independently killed process. |
| Performance | Same-machine test-enabled Release pair, exactly 200 samples per existing metric | Recomputed p95 <= baseline * 1.10; corpus/controls match; memory/high-water counters included; no payload text archived. |
| Product integration | Affected iteration, PrePush obligations, required Full, native Debug/Release and applicable ASan | Original failed evidence retained; missing foreground/hardware/native cases pending; source/build/scope receipts exact. |
| Portable packages | Production x64/ARM64 ZIPs, manifest/hash/notice checks, clean extraction/lifecycle | No stale previous-runtime output; exact pair; required notices present; both native Release branches accepted before production packaging. |

Any optional improvement later activated adds its own behavioral and performance
gate; its absence is not a failed upgrade.

### 8.2 Existing commands to reuse

These are **future execution commands**, not results. Resolve tool versions,
candidate paths, and the already marked sandbox before use. Candidate build
input/output locations stay repository-local; runtime test archives stay under
the selected marked sandbox and are published with the existing archive tooling.

```powershell
# Read-only planning; does not authorize a build or native test during this proposal.
.\Test-Changes.ps1 -Explain

# Once exact library-aware candidate observation/review support is implemented:
.\Tools\Get-GhosttyTerminalRuntimeStatus.ps1 -Mode Candidate -CandidateCommit 0fa3d577a5614150ff8e2c152de5fb9d5c6d2dec -PassThru

# Explicit isolated candidate inputs; never a fallback to Product.
.\Tools\Build-GhosttyTerminalRuntime.ps1 -Platform x64 -CandidateLockPath .build\TerminalEngineUpgrade\0fa3d577a5614150ff8e2c152de5fb9d5c6d2dec\candidate.lock.json -CandidateOutputRoot .build\TerminalEngineUpgrade\0fa3d577a5614150ff8e2c152de5fb9d5c6d2dec\x64 -Force -PassThru
.\Tools\Build-GhosttyTerminalRuntime.ps1 -Platform ARM64 -CandidateLockPath .build\TerminalEngineUpgrade\0fa3d577a5614150ff8e2c152de5fb9d5c6d2dec\candidate.lock.json -CandidateOutputRoot .build\TerminalEngineUpgrade\0fa3d577a5614150ff8e2c152de5fb9d5c6d2dec\ARM64 -Force -PassThru

# After implementation: execute the affected plan, then account for PrePush.
.\Test-Changes.ps1
.\Test-Changes.ps1 -Mode PrePush

# Full is required by the Terminal upgrade contract. Foreground coverage needs
# agreement to the run time and the runner's Interactive option, not inferred consent.
.\Tools\Run-AllTests.ps1 -Suite Full -ValidationMode Fresh

# Documentation/tooling integrity; these do not qualify native behavior.
.\Tools\Get-SpecInventory.ps1 -FailOnFindings
.\Tools\Get-ToolInventory.ps1 -Validate
.\Tools\Test-Changelog.ps1
```

Use `New-GhosttyTerminalRuntimeCandidateReview.ps1` and
`New-GhosttyTerminalRuntimeCandidateOverlayReview.ps1` with their documented
required paths/descriptor/concern arguments; do not bypass them with a hand-edited
candidate lock. `PluginContractTests.exe --terminal-selftests` is the existing
focused plugin route. `--terminal-vt-upgrade-corpus <archive-directory>` is the
test-enabled Release corpus route. The paired acceptance command is
`Tools/Test-TestRunArchive.ps1 -BaselineRunPath <product-run> -CandidateRunPath <candidate-run>`.
Use actual owned sandbox run paths; placeholders here are not invented run IDs.

The existing tooling suite includes `TerminalPluginSourceContracts.Tests.ps1`,
`TerminalRuntimeExportIdentity.Tests.ps1`, `TerminalEngineRuntimeImportPolicy.Tests.ps1`,
and the current Ghostty tests discovered under `Tools/Tests`. Select tests from
the governed impact plan and changed module contracts; do not execute sealed
historical engine-selection tooling as if it qualified the current product.

### 8.3 Evidence packet and limitations

Archive under `Specs/TestRuns/<MachineHash>/Terminal/<RunId>/` through the owning
publisher: source/lock/overlay/dependency digests; candidate and immutable P4
descriptors; signature verification record; all declaration/type/enum review
dispositions; native ABI layouts/call results; exact receipts and runtime pairs;
upstream/C-example outcomes; raw performance samples/controls; resource counters;
portable manifests and smoke results. Keep complete failure records and explicit
skips. Store synthetic fixture content only, not terminal queries/output,
clipboard contents, personal paths, accounts, hostnames, or process command lines.

Upstream acceptance of Windows x64 does not replace native Windows ARM64
qualification. The announcement describes ARM64 build coverage, while the
packaging platform table is more conservative about its tier. Our own native
ARM64 acceptance remains required under either description.

## 9. Risks, stop conditions, and rollback

| Risk | Consequence | Control |
|---|---|---|
| Old/new DLL/header mix | Symbol failures or type confusion behind generic getters | Exact raw pair, complete rebuild, version/features and negative load tests. |
| Private enum collision | A placement value may be interpreted as a resize callback | Allocate after the full public ranges; C/Zig and behavioral guards. |
| Mechanical overlay rebase | Loss of newer relative-placement/ownership fixes | Rederive invariants against 1.0; preserve upstream changes; allocator/pin tests. |
| Copy formatter state leaks | Find/UIA/export sees only a previous selection | Restore options or use separate owned formatter; sequential interaction tests. |
| Resource-limit regression | Unbounded placement/frame/map retention | Exact allocator-backed counters, all entry paths, high-water and teardown tests. |
| Release/version ambiguity | Dev/app source mislabeled as stable 1.0 | Signed source + LIB_VERSION + typed release identity + runtime check. |
| Archive/build-system drift | Wrong graph, dependencies, flags, notices or stale cache | Fresh isolated builds and complete dependency/cache closure. |
| Performance comparison drift | Apparent gain hides changed formatting/capture work | Frozen corpus/control serialization, compatible adapters, same environment and raw samples. |

Stop admission when provenance/identity is unresolved; a required feature is
unsupported; a private value collides; overlay caps are exceeded; typed getter
storage is unreviewed; a layout/call probe fails; a bound, semantic control,
native acceptance, required Full, or paired performance gate fails; or package
identity/notices are incomplete. Missing user agreement to foreground timing
stops that phase, not independent document or no-activation work.

Rollback restores the previous **complete tuple**: source/tool/dependency lock,
overlay, built `Terminal.dll`, private runtime DLL, identity, generated pairing
header, and the tuple's own notice manifest. Do not swap only the runtime DLL,
keep a 1.0 plugin with a 1.4-source runtime, or restore a hard-coded “eight
notices” assumption after the closure changes. Preserve candidate evidence for
diagnosis. Verify old-pair load/lifecycle and clean portable staging after rollback.

## 10. Decisions requested during document review

| Decision | Recommendation | Effect |
|---|---|---|
| D1 Migration scope | Upgrade and preserve existing experience; keep E1–E16 optional | Keeps acceptance attributable to the dependency migration. |
| D2 Source distribution | Signed 1.0 source/dependency tarballs for builds, authenticated Git source for review | Requires builder/provenance adaptation; avoids building the entire app distribution as the library package. |
| D3 Private controls | Retain placement/clipboard invariants; reallocate 49/50; remove build hunk only with proof | Maintains current behavior without claiming the private derivative is upstream ABI-only. |
| D4 Capability/CPU policy | Keep current SIMD-disabled baseline and full intended feature set; assert required capabilities | Avoids bundling a throughput/size tradeoff into ABI migration. |
| D5 Future experience | First investigate E1 search, E2/E3 selection/prompt navigation, E4 passive progress; profile E5 | This is a suggested review order, not authorization to implement these features. |
| D6 Native acceptance | Require real x64 and ARM64 product/ABI/performance/package evidence | Cross-build/PR green is insufficient; foreground tests wait for an agreed run time. |

If a recommendation changes, revise this document's phases and acceptance
criteria before implementation. No answer is needed to finish the document.

## 11. Completion checklist and review evidence

### Proposal deliverable

- [x] Resolve the current pin and exact 1.0 library release independently of app versioning.
- [x] Compare the complete public header/native symbol surface and current loader.
- [x] Identify concrete ABI/source incompatibilities and the private option collision.
- [x] Explain each existing patch concern and evidence needed for its disposition.
- [x] Inventory unused APIs and bounded, separate experience opportunities.
- [x] Define ordered work, validation, stop/rollback and normative closeout requirements.
- [ ] Owner reviews the proposal and selects/adjusts recommendations.

### Implementation acceptance — all pending

- [ ] P0 baseline and rollback packet.
- [ ] P1 authenticated library-aware release/descriptor/status model.
- [ ] P2 correct build/dependency/notice closure.
- [ ] P3 reviewed five-concern overlay with collision-free private ABI.
- [ ] P4 adapted plugin and native x64/ARM64 ABI proof.
- [ ] P5 source-bound behavioral/Full/performance acceptance.
- [ ] P6 admitted-product/package acceptance, enduring spec updates, and reviewed Done admission.

Document checks performed while authoring are recorded at the end of this
document. They establish documentation/inventory integrity only. They do not
check off any implementation acceptance item.

## Appendix A. Complete target function inventory

This inventory contains all 209 native public names from the release symbol baseline and all five Wasm-only helpers. “Loaded” means the current lookup table requires the name; it does not imply every selector or capability is used. `ghostty_alloc` is used through a copied callback function pointer, so a search for direct member calls alone would incorrectly classify it as unused. No Wasm helper is part of the Windows export expectation.

| Target native function | Current adoption / migration disposition | Public header |
|---|---|---|
| `ghostty_alloc` | Loaded; declaration unchanged | `allocator.h` |
| `ghostty_build_info` | Loaded; declaration unchanged | `build_info.h` |
| `ghostty_cell_get` | Loaded; declaration unchanged | `screen.h` |
| `ghostty_cell_get_multi` | Not loaded; optional or validation-only | `screen.h` |
| `ghostty_color_contrast` | Not loaded; optional or validation-only | `color.h` |
| `ghostty_color_luminance` | Not loaded; optional or validation-only | `color.h` |
| `ghostty_color_palette_default` | Not loaded; optional or validation-only | `color.h` |
| `ghostty_color_palette_generate` | Not loaded; optional or validation-only | `color.h` |
| `ghostty_color_parse` | Not loaded; optional or validation-only | `color.h` |
| `ghostty_color_parse_palette_entry` | Not loaded; optional or validation-only | `color.h` |
| `ghostty_color_parse_x11` | Not loaded; optional or validation-only | `color.h` |
| `ghostty_color_perceived_luminance` | Not loaded; optional or validation-only | `color.h` |
| `ghostty_color_rgb_get` | Not loaded; optional or validation-only | `color.h` |
| `ghostty_color_scheme_report_encode` | Not loaded; optional or validation-only | `color_scheme.h` |
| `ghostty_color_x11_name_count` | Not loaded; optional or validation-only | `color.h` |
| `ghostty_color_x11_names` | Not loaded; optional or validation-only | `color.h` |
| `ghostty_focus_encode` | Not loaded; optional or validation-only | `focus.h` |
| `ghostty_formatter_format` | Not loaded; optional or validation-only | `formatter.h` |
| `ghostty_formatter_format_alloc` | Not loaded; optional or validation-only | `formatter.h` |
| `ghostty_formatter_format_buf` | Loaded; declaration unchanged | `formatter.h` |
| `ghostty_formatter_free` | Loaded; declaration unchanged | `formatter.h` |
| `ghostty_formatter_get` | Not loaded; optional or validation-only | `formatter.h` |
| `ghostty_formatter_set` | Not loaded; migration addition | `formatter.h` |
| `ghostty_formatter_terminal_new` | Loaded; declaration changes | `formatter.h` |
| `ghostty_free` | Loaded; declaration unchanged | `allocator.h` |
| `ghostty_grid_ref_cell` | Not loaded; optional or validation-only | `grid_ref.h` |
| `ghostty_grid_ref_graphemes` | Not loaded; optional or validation-only | `grid_ref.h` |
| `ghostty_grid_ref_hyperlink_uri` | Loaded; parameter-name-only change | `grid_ref.h` |
| `ghostty_grid_ref_row` | Not loaded; optional or validation-only | `grid_ref.h` |
| `ghostty_grid_ref_style` | Not loaded; optional or validation-only | `grid_ref.h` |
| `ghostty_key_encoder_encode` | Loaded; declaration changes | `key/encoder.h` |
| `ghostty_key_encoder_free` | Loaded; declaration unchanged | `key/encoder.h` |
| `ghostty_key_encoder_new` | Loaded; declaration unchanged | `key/encoder.h` |
| `ghostty_key_encoder_set` | Not loaded; optional or validation-only | `key/encoder.h` |
| `ghostty_key_encoder_set_from_terminal` | Not loaded; migration addition | `key/encoder.h` |
| `ghostty_key_event_free` | Loaded; declaration unchanged | `key/event.h` |
| `ghostty_key_event_get_action` | Not loaded; optional or validation-only | `key/event.h` |
| `ghostty_key_event_get_composing` | Not loaded; optional or validation-only | `key/event.h` |
| `ghostty_key_event_get_consumed_mods` | Not loaded; optional or validation-only | `key/event.h` |
| `ghostty_key_event_get_key` | Not loaded; optional or validation-only | `key/event.h` |
| `ghostty_key_event_get_mods` | Not loaded; optional or validation-only | `key/event.h` |
| `ghostty_key_event_get_unshifted_codepoint` | Not loaded; optional or validation-only | `key/event.h` |
| `ghostty_key_event_get_utf8` | Not loaded; optional or validation-only | `key/event.h` |
| `ghostty_key_event_new` | Loaded; declaration unchanged | `key/event.h` |
| `ghostty_key_event_set_action` | Loaded; declaration changes | `key/event.h` |
| `ghostty_key_event_set_composing` | Not loaded; optional or validation-only | `key/event.h` |
| `ghostty_key_event_set_consumed_mods` | Not loaded; optional or validation-only | `key/event.h` |
| `ghostty_key_event_set_key` | Loaded; declaration changes | `key/event.h` |
| `ghostty_key_event_set_mods` | Loaded; declaration changes | `key/event.h` |
| `ghostty_key_event_set_unshifted_codepoint` | Loaded; declaration changes | `key/event.h` |
| `ghostty_key_event_set_utf8` | Loaded; declaration changes | `key/event.h` |
| `ghostty_kitty_graphics_animation_tick` | Not loaded; optional or validation-only | `kitty_graphics.h` |
| `ghostty_kitty_graphics_get` | Loaded; declaration unchanged | `kitty_graphics.h` |
| `ghostty_kitty_graphics_image` | Loaded; declaration unchanged | `kitty_graphics.h` |
| `ghostty_kitty_graphics_image_get` | Not loaded; optional or validation-only | `kitty_graphics.h` |
| `ghostty_kitty_graphics_image_get_multi` | Loaded; declaration unchanged | `kitty_graphics.h` |
| `ghostty_kitty_graphics_placement_get` | Loaded; declaration unchanged | `kitty_graphics.h` |
| `ghostty_kitty_graphics_placement_get_multi` | Not loaded; optional or validation-only | `kitty_graphics.h` |
| `ghostty_kitty_graphics_placement_grid_size` | Not loaded; optional or validation-only | `kitty_graphics.h` |
| `ghostty_kitty_graphics_placement_iterator_free` | Loaded; declaration unchanged | `kitty_graphics.h` |
| `ghostty_kitty_graphics_placement_iterator_new` | Loaded; declaration unchanged | `kitty_graphics.h` |
| `ghostty_kitty_graphics_placement_iterator_set` | Not loaded; optional or validation-only | `kitty_graphics.h` |
| `ghostty_kitty_graphics_placement_next` | Loaded; declaration unchanged | `kitty_graphics.h` |
| `ghostty_kitty_graphics_placement_pixel_size` | Not loaded; optional or validation-only | `kitty_graphics.h` |
| `ghostty_kitty_graphics_placement_rect` | Not loaded; optional or validation-only | `kitty_graphics.h` |
| `ghostty_kitty_graphics_placement_render_info` | Loaded; declaration unchanged | `kitty_graphics.h` |
| `ghostty_kitty_graphics_placement_source_rect` | Not loaded; optional or validation-only | `kitty_graphics.h` |
| `ghostty_kitty_graphics_placement_viewport_pos` | Not loaded; optional or validation-only | `kitty_graphics.h` |
| `ghostty_mode_report_encode` | Not loaded; optional or validation-only | `modes.h` |
| `ghostty_mouse_encoder_encode` | Loaded; declaration changes | `mouse/encoder.h` |
| `ghostty_mouse_encoder_free` | Loaded; declaration unchanged | `mouse/encoder.h` |
| `ghostty_mouse_encoder_new` | Loaded; declaration unchanged | `mouse/encoder.h` |
| `ghostty_mouse_encoder_reset` | Not loaded; optional or validation-only | `mouse/encoder.h` |
| `ghostty_mouse_encoder_set` | Not loaded; migration addition | `mouse/encoder.h` |
| `ghostty_mouse_encoder_set_from_terminal` | Not loaded; migration addition | `mouse/encoder.h` |
| `ghostty_mouse_event_clear_button` | Loaded; declaration changes | `mouse/event.h` |
| `ghostty_mouse_event_free` | Loaded; declaration unchanged | `mouse/event.h` |
| `ghostty_mouse_event_get_action` | Not loaded; optional or validation-only | `mouse/event.h` |
| `ghostty_mouse_event_get_button` | Not loaded; optional or validation-only | `mouse/event.h` |
| `ghostty_mouse_event_get_mods` | Not loaded; optional or validation-only | `mouse/event.h` |
| `ghostty_mouse_event_get_position` | Not loaded; optional or validation-only | `mouse/event.h` |
| `ghostty_mouse_event_new` | Loaded; declaration unchanged | `mouse/event.h` |
| `ghostty_mouse_event_set_action` | Loaded; declaration changes | `mouse/event.h` |
| `ghostty_mouse_event_set_button` | Loaded; declaration changes | `mouse/event.h` |
| `ghostty_mouse_event_set_mods` | Loaded; declaration changes | `mouse/event.h` |
| `ghostty_mouse_event_set_position` | Loaded; declaration changes | `mouse/event.h` |
| `ghostty_osc_command_data` | Not loaded; optional or validation-only | `osc.h` |
| `ghostty_osc_command_type` | Not loaded; optional or validation-only | `osc.h` |
| `ghostty_osc_end` | Not loaded; optional or validation-only | `osc.h` |
| `ghostty_osc_free` | Not loaded; optional or validation-only | `osc.h` |
| `ghostty_osc_new` | Not loaded; optional or validation-only | `osc.h` |
| `ghostty_osc_next` | Not loaded; optional or validation-only | `osc.h` |
| `ghostty_osc_reset` | Not loaded; optional or validation-only | `osc.h` |
| `ghostty_osc_set` | Not loaded; optional or validation-only | `osc.h` |
| `ghostty_paste_encode` | Loaded; declaration changes | `paste.h` |
| `ghostty_paste_is_safe` | Loaded; declaration unchanged | `paste.h` |
| `ghostty_render_state_begin_update` | Loaded; declaration unchanged | `render.h` |
| `ghostty_render_state_clean` | Not loaded; optional or validation-only | `render.h` |
| `ghostty_render_state_end_update` | Loaded; declaration unchanged | `render.h` |
| `ghostty_render_state_free` | Loaded; declaration unchanged | `render.h` |
| `ghostty_render_state_get` | Loaded; declaration unchanged | `render.h` |
| `ghostty_render_state_get_multi` | Not loaded; optional or validation-only | `render.h` |
| `ghostty_render_state_image_get` | Not loaded; optional or validation-only | `render.h` |
| `ghostty_render_state_image_get_multi` | Not loaded; optional or validation-only | `render.h` |
| `ghostty_render_state_image_iterator_free` | Not loaded; optional or validation-only | `render.h` |
| `ghostty_render_state_image_iterator_new` | Not loaded; optional or validation-only | `render.h` |
| `ghostty_render_state_image_iterator_next` | Not loaded; optional or validation-only | `render.h` |
| `ghostty_render_state_new` | Loaded; declaration unchanged | `render.h` |
| `ghostty_render_state_row_cells_free` | Loaded; declaration unchanged | `render.h` |
| `ghostty_render_state_row_cells_get` | Loaded; declaration unchanged | `render.h` |
| `ghostty_render_state_row_cells_get_multi` | Not loaded; optional or validation-only | `render.h` |
| `ghostty_render_state_row_cells_new` | Loaded; declaration unchanged | `render.h` |
| `ghostty_render_state_row_cells_next` | Loaded; declaration unchanged | `render.h` |
| `ghostty_render_state_row_cells_select` | Not loaded; optional or validation-only | `render.h` |
| `ghostty_render_state_row_get` | Loaded; declaration unchanged | `render.h` |
| `ghostty_render_state_row_get_multi` | Not loaded; optional or validation-only | `render.h` |
| `ghostty_render_state_row_iterator_free` | Loaded; declaration unchanged | `render.h` |
| `ghostty_render_state_row_iterator_new` | Loaded; declaration unchanged | `render.h` |
| `ghostty_render_state_row_iterator_next` | Loaded; declaration unchanged | `render.h` |
| `ghostty_render_state_row_iterator_next_dirty` | Not loaded; optional or validation-only | `render.h` |
| `ghostty_render_state_row_set` | Loaded; declaration unchanged | `render.h` |
| `ghostty_render_state_set` | Loaded; declaration unchanged | `render.h` |
| `ghostty_render_state_update` | Not loaded; optional or validation-only | `render.h` |
| `ghostty_row_get` | Not loaded; optional or validation-only | `screen.h` |
| `ghostty_row_get_multi` | Not loaded; optional or validation-only | `screen.h` |
| `ghostty_search_feed` | Not loaded; optional or validation-only | `search.h` |
| `ghostty_search_free` | Not loaded; optional or validation-only | `search.h` |
| `ghostty_search_get` | Not loaded; optional or validation-only | `search.h` |
| `ghostty_search_get_multi` | Not loaded; optional or validation-only | `search.h` |
| `ghostty_search_new` | Not loaded; optional or validation-only | `search.h` |
| `ghostty_search_run` | Not loaded; optional or validation-only | `search.h` |
| `ghostty_search_select_next` | Not loaded; optional or validation-only | `search.h` |
| `ghostty_search_select_prev` | Not loaded; optional or validation-only | `search.h` |
| `ghostty_search_set` | Not loaded; optional or validation-only | `search.h` |
| `ghostty_search_tick` | Not loaded; optional or validation-only | `search.h` |
| `ghostty_selection_gesture_event` | Not loaded; optional or validation-only | `selection.h` |
| `ghostty_selection_gesture_event_free` | Not loaded; optional or validation-only | `selection.h` |
| `ghostty_selection_gesture_event_new` | Not loaded; optional or validation-only | `selection.h` |
| `ghostty_selection_gesture_event_set` | Not loaded; optional or validation-only | `selection.h` |
| `ghostty_selection_gesture_free` | Not loaded; optional or validation-only | `selection.h` |
| `ghostty_selection_gesture_get` | Not loaded; optional or validation-only | `selection.h` |
| `ghostty_selection_gesture_get_multi` | Not loaded; optional or validation-only | `selection.h` |
| `ghostty_selection_gesture_new` | Not loaded; optional or validation-only | `selection.h` |
| `ghostty_selection_gesture_reset` | Not loaded; optional or validation-only | `selection.h` |
| `ghostty_sgr_attribute_tag` | Not loaded; optional or validation-only | `sgr.h` |
| `ghostty_sgr_attribute_value` | Not loaded; optional or validation-only | `sgr.h` |
| `ghostty_sgr_free` | Not loaded; optional or validation-only | `sgr.h` |
| `ghostty_sgr_new` | Not loaded; optional or validation-only | `sgr.h` |
| `ghostty_sgr_next` | Not loaded; optional or validation-only | `sgr.h` |
| `ghostty_sgr_reset` | Not loaded; optional or validation-only | `sgr.h` |
| `ghostty_sgr_set_params` | Not loaded; optional or validation-only | `sgr.h` |
| `ghostty_sgr_unknown_full` | Not loaded; optional or validation-only | `sgr.h` |
| `ghostty_sgr_unknown_partial` | Not loaded; optional or validation-only | `sgr.h` |
| `ghostty_size_report_encode` | Not loaded; optional or validation-only | `size_report.h` |
| `ghostty_sized_init` | Not loaded; optional or validation-only | `types.h` |
| `ghostty_snapshot_decoder_decode` | Not loaded; optional or validation-only | `snapshot.h` |
| `ghostty_snapshot_decoder_free` | Not loaded; optional or validation-only | `snapshot.h` |
| `ghostty_snapshot_decoder_get` | Not loaded; optional or validation-only | `snapshot.h` |
| `ghostty_snapshot_decoder_get_multi` | Not loaded; optional or validation-only | `snapshot.h` |
| `ghostty_snapshot_decoder_new` | Not loaded; optional or validation-only | `snapshot.h` |
| `ghostty_snapshot_decoder_new_buf` | Not loaded; optional or validation-only | `snapshot.h` |
| `ghostty_snapshot_decoder_next` | Not loaded; optional or validation-only | `snapshot.h` |
| `ghostty_snapshot_decoder_ready` | Not loaded; optional or validation-only | `snapshot.h` |
| `ghostty_snapshot_decoder_set` | Not loaded; optional or validation-only | `snapshot.h` |
| `ghostty_snapshot_encode` | Not loaded; optional or validation-only | `snapshot.h` |
| `ghostty_snapshot_encode_alloc` | Not loaded; optional or validation-only | `snapshot.h` |
| `ghostty_snapshot_encode_buf` | Not loaded; optional or validation-only | `snapshot.h` |
| `ghostty_style_default` | Not loaded; optional or validation-only | `style.h` |
| `ghostty_style_is_default` | Not loaded; optional or validation-only | `style.h` |
| `ghostty_sys_log_stderr` | Not loaded; optional or validation-only | `sys.h` |
| `ghostty_sys_set` | Loaded; declaration unchanged | `sys.h` |
| `ghostty_terminal_compress` | Not loaded; optional or validation-only | `terminal.h` |
| `ghostty_terminal_compression_activity` | Not loaded; optional or validation-only | `terminal.h` |
| `ghostty_terminal_continuation_alloc` | Not loaded; optional or validation-only | `terminal.h` |
| `ghostty_terminal_continuation_buf` | Not loaded; optional or validation-only | `terminal.h` |
| `ghostty_terminal_continuation_write` | Not loaded; optional or validation-only | `terminal.h` |
| `ghostty_terminal_free` | Loaded; declaration unchanged | `terminal.h` |
| `ghostty_terminal_get` | Loaded; declaration unchanged | `terminal.h` |
| `ghostty_terminal_get_multi` | Not loaded; optional or validation-only | `terminal.h` |
| `ghostty_terminal_grid_ref` | Loaded; declaration unchanged | `terminal.h` |
| `ghostty_terminal_grid_ref_track` | Loaded; declaration unchanged | `terminal.h` |
| `ghostty_terminal_new` | Loaded; declaration unchanged | `terminal.h` |
| `ghostty_terminal_paste` | Not loaded; optional or validation-only | `paste.h` |
| `ghostty_terminal_point_from_grid_ref` | Not loaded; optional or validation-only | `terminal.h` |
| `ghostty_terminal_reset` | Not loaded; optional or validation-only | `terminal.h` |
| `ghostty_terminal_resize` | Loaded; declaration unchanged | `terminal.h` |
| `ghostty_terminal_scroll_viewport` | Loaded; declaration changes | `terminal.h` |
| `ghostty_terminal_select_all` | Loaded; declaration unchanged | `selection.h` |
| `ghostty_terminal_select_line` | Not loaded; optional or validation-only | `selection.h` |
| `ghostty_terminal_select_output` | Not loaded; optional or validation-only | `selection.h` |
| `ghostty_terminal_select_word` | Loaded; declaration unchanged | `selection.h` |
| `ghostty_terminal_select_word_between` | Not loaded; optional or validation-only | `selection.h` |
| `ghostty_terminal_selection_adjust` | Not loaded; optional or validation-only | `selection.h` |
| `ghostty_terminal_selection_contains` | Not loaded; optional or validation-only | `selection.h` |
| `ghostty_terminal_selection_equal` | Not loaded; optional or validation-only | `selection.h` |
| `ghostty_terminal_selection_order` | Not loaded; optional or validation-only | `selection.h` |
| `ghostty_terminal_selection_ordered` | Not loaded; optional or validation-only | `selection.h` |
| `ghostty_terminal_set` | Loaded; declaration unchanged | `terminal.h` |
| `ghostty_terminal_vt_write` | Loaded; declaration unchanged | `terminal.h` |
| `ghostty_terminal_vt_write_until` | Not loaded; optional or validation-only | `terminal.h` |
| `ghostty_tracked_grid_ref_free` | Loaded; declaration unchanged | `grid_ref_tracked.h` |
| `ghostty_tracked_grid_ref_has_value` | Not loaded; optional or validation-only | `grid_ref_tracked.h` |
| `ghostty_tracked_grid_ref_point` | Not loaded; optional or validation-only | `grid_ref_tracked.h` |
| `ghostty_tracked_grid_ref_set` | Not loaded; optional or validation-only | `grid_ref_tracked.h` |
| `ghostty_tracked_grid_ref_snapshot` | Loaded; declaration unchanged | `grid_ref_tracked.h` |
| `ghostty_type_json` | Not loaded; optional or validation-only | `types.h` |
| `ghostty_unicode_codepoint_width` | Not loaded; optional or validation-only | `unicode.h` |
| `ghostty_unicode_grapheme_width` | Not loaded; optional or validation-only | `unicode.h` |
| `ghostty_version` | Not loaded; migration addition | `build_info.h` |

Wasm-only declarations, all outside the native product:

- `ghostty_wasm_alloc`
- `ghostty_wasm_free`
- `ghostty_wasm_alloc_opaque`
- `ghostty_wasm_free_opaque`
- `ghostty_wasm_take_opaque`

## Appendix B. Every incumbent loader requirement

All 71 current required symbols are mapped below. References identify representative real call sites, including test-enabled diagnostics embedded in `Terminal.cpp`; they do not claim that every listed occurrence is a shipping call. Loader lookup locations are the primary evidence of requirement.

| Current required export | Representative consumer or indirection | Target action |
|---|---|---|
| `ghostty_build_info` | `TerminalRuntimeLoader.cpp:548` | Keep; recompile and validate selector/pointee contracts |
| `ghostty_terminal_new` | `Terminal.cpp:1538` | Keep; recompile and validate selector/pointee contracts |
| `ghostty_terminal_free` | `Terminal.cpp:1528` | Keep; recompile and validate selector/pointee contracts |
| `ghostty_terminal_resize` | `Terminal.cpp:788` | Keep; recompile and validate selector/pointee contracts |
| `ghostty_terminal_set` | `Terminal.cpp:1667` | Keep; recompile and validate selector/pointee contracts |
| `ghostty_terminal_get` | `Terminal.cpp:801` | Keep; recompile and validate selector/pointee contracts |
| `ghostty_terminal_grid_ref` | `Terminal.cpp:1630` | Keep; recompile and validate selector/pointee contracts |
| `ghostty_grid_ref_hyperlink_uri` | `Terminal.cpp:1631` | Keep; output parameter name only |
| `ghostty_terminal_grid_ref_track` | `Terminal.cpp:1764` | Keep; recompile and validate selector/pointee contracts |
| `ghostty_tracked_grid_ref_free` | `Terminal.cpp:1500` | Keep; recompile and validate selector/pointee contracts |
| `ghostty_tracked_grid_ref_snapshot` | `Terminal.cpp:1765` | Keep; recompile and validate selector/pointee contracts |
| `ghostty_terminal_vt_write` | `Terminal.cpp:269` | Keep; recompile and validate selector/pointee contracts |
| `ghostty_terminal_select_word` | `TerminalVt.cpp:135` | Keep; recompile and validate selector/pointee contracts |
| `ghostty_terminal_select_all` | `Terminal.cpp:1744` | Keep; recompile and validate selector/pointee contracts |
| `ghostty_terminal_selection_format_buf` | `Terminal.cpp:1755` | Replace with `ghostty_formatter_set + ghostty_formatter_format_buf` |
| `ghostty_terminal_scroll_viewport` | `Terminal.cpp:808` | Adapt declaration and callers (Appendix C) |
| `ghostty_paste_is_safe` | `Terminal.cpp:1737` | Keep; recompile and validate selector/pointee contracts |
| `ghostty_paste_encode` | `Terminal.cpp:1738` | Adapt declaration and callers (Appendix C) |
| `ghostty_render_state_new` | `Terminal.cpp:1539` | Keep; recompile and validate selector/pointee contracts |
| `ghostty_render_state_free` | `Terminal.cpp:1524` | Keep; recompile and validate selector/pointee contracts |
| `ghostty_render_state_begin_update` | `Terminal.cpp:1575` | Keep; recompile and validate selector/pointee contracts |
| `ghostty_render_state_end_update` | `Terminal.cpp:1575` | Keep; recompile and validate selector/pointee contracts |
| `ghostty_render_state_get` | `Terminal.cpp:1577` | Keep; recompile and validate selector/pointee contracts |
| `ghostty_render_state_set` | `TerminalVt.cpp:2160` | Keep; recompile and validate selector/pointee contracts |
| `ghostty_render_state_row_iterator_new` | `Terminal.cpp:1540` | Keep; recompile and validate selector/pointee contracts |
| `ghostty_render_state_row_iterator_free` | `Terminal.cpp:1520` | Keep; recompile and validate selector/pointee contracts |
| `ghostty_render_state_row_iterator_next` | `Terminal.cpp:1578` | Keep; recompile and validate selector/pointee contracts |
| `ghostty_render_state_row_get` | `Terminal.cpp:1578` | Keep; recompile and validate selector/pointee contracts |
| `ghostty_render_state_row_set` | `TerminalVt.cpp:2157` | Keep; recompile and validate selector/pointee contracts |
| `ghostty_render_state_row_cells_new` | `Terminal.cpp:1540` | Keep; recompile and validate selector/pointee contracts |
| `ghostty_render_state_row_cells_free` | `Terminal.cpp:1516` | Keep; recompile and validate selector/pointee contracts |
| `ghostty_render_state_row_cells_next` | `Terminal.cpp:1579` | Keep; recompile and validate selector/pointee contracts |
| `ghostty_render_state_row_cells_get` | `Terminal.cpp:1585` | Keep; recompile and validate selector/pointee contracts |
| `ghostty_cell_get` | `TerminalVt.cpp:1924` | Keep; recompile and validate selector/pointee contracts |
| `ghostty_key_encoder_new` | `Terminal.cpp:1541` | Keep; recompile and validate selector/pointee contracts |
| `ghostty_key_encoder_free` | `Terminal.cpp:1512` | Keep; recompile and validate selector/pointee contracts |
| `ghostty_key_encoder_setopt_from_terminal` | `Terminal.cpp:1720` | Replace with `ghostty_key_encoder_set_from_terminal` |
| `ghostty_key_encoder_encode` | `Terminal.cpp:1723` | Adapt declaration and callers (Appendix C) |
| `ghostty_key_event_new` | `Terminal.cpp:1542` | Keep; recompile and validate selector/pointee contracts |
| `ghostty_key_event_free` | `Terminal.cpp:1508` | Keep; recompile and validate selector/pointee contracts |
| `ghostty_key_event_set_action` | `Terminal.cpp:1715` | Adapt declaration and callers (Appendix C) |
| `ghostty_key_event_set_key` | `Terminal.cpp:1716` | Adapt declaration and callers (Appendix C) |
| `ghostty_key_event_set_mods` | `Terminal.cpp:1717` | Adapt declaration and callers (Appendix C) |
| `ghostty_key_event_set_utf8` | `Terminal.cpp:1718` | Adapt declaration and callers (Appendix C) |
| `ghostty_key_event_set_unshifted_codepoint` | `Terminal.cpp:1719` | Adapt declaration and callers (Appendix C) |
| `ghostty_mouse_encoder_new` | `Terminal.cpp:1542` | Keep; recompile and validate selector/pointee contracts |
| `ghostty_mouse_encoder_free` | `Terminal.cpp:1504` | Keep; recompile and validate selector/pointee contracts |
| `ghostty_mouse_encoder_setopt` | `Terminal.cpp:1777` | Replace with `ghostty_mouse_encoder_set` |
| `ghostty_mouse_encoder_setopt_from_terminal` | `Terminal.cpp:1770` | Replace with `ghostty_mouse_encoder_set_from_terminal` |
| `ghostty_mouse_encoder_encode` | `Terminal.cpp:1787` | Adapt declaration and callers (Appendix C) |
| `ghostty_mouse_event_new` | `Terminal.cpp:1543` | Keep; recompile and validate selector/pointee contracts |
| `ghostty_mouse_event_free` | `Terminal.cpp:1492` | Keep; recompile and validate selector/pointee contracts |
| `ghostty_mouse_event_set_action` | `Terminal.cpp:1781` | Adapt declaration and callers (Appendix C) |
| `ghostty_mouse_event_set_button` | `Terminal.cpp:1782` | Adapt declaration and callers (Appendix C) |
| `ghostty_mouse_event_clear_button` | `TerminalVt.cpp:317` | Adapt declaration and callers (Appendix C) |
| `ghostty_mouse_event_set_mods` | `Terminal.cpp:1783` | Adapt declaration and callers (Appendix C) |
| `ghostty_mouse_event_set_position` | `Terminal.cpp:1784` | Adapt declaration and callers (Appendix C) |
| `ghostty_alloc` | `TerminalRuntimeLoader.cpp:171 (copied allocator callback)` | Keep; recompile and validate selector/pointee contracts |
| `ghostty_free` | `Terminal.cpp:2426` | Keep; recompile and validate selector/pointee contracts |
| `ghostty_sys_set` | `TerminalRuntimeLoader.cpp:565` | Keep; recompile and validate selector/pointee contracts |
| `ghostty_kitty_graphics_get` | `Terminal.cpp:1851` | Keep; recompile and validate selector/pointee contracts |
| `ghostty_kitty_graphics_image` | `Terminal.cpp:1857` | Keep; recompile and validate selector/pointee contracts |
| `ghostty_kitty_graphics_image_get_multi` | `Terminal.cpp:1864` | Keep; recompile and validate selector/pointee contracts |
| `ghostty_kitty_graphics_placement_iterator_new` | `Terminal.cpp:1544` | Keep; recompile and validate selector/pointee contracts |
| `ghostty_kitty_graphics_placement_iterator_free` | `Terminal.cpp:1496` | Keep; recompile and validate selector/pointee contracts |
| `ghostty_kitty_graphics_placement_next` | `Terminal.cpp:1852` | Keep; recompile and validate selector/pointee contracts |
| `ghostty_kitty_graphics_placement_get` | `Terminal.cpp:1856` | Keep; recompile and validate selector/pointee contracts |
| `ghostty_kitty_graphics_placement_render_info` | `TerminalVt.cpp:1299` | Keep; recompile and validate selector/pointee contracts |
| `ghostty_formatter_terminal_new` | `Terminal.cpp:2758` | Adapt declaration and callers (Appendix C) |
| `ghostty_formatter_format_buf` | `TerminalVt.cpp:1134` | Keep; recompile and validate selector/pointee contracts |
| `ghostty_formatter_free` | `Terminal.cpp:2910` | Keep; recompile and validate selector/pointee contracts |

## Appendix C. Complete function declaration changes

All 10 additions, six removals and 30 textual changes are included. Parameter renames are labeled separately; unchanged function text is not evidence that enum-keyed pointee contracts or layouts are unchanged.

### C.1 Added native functions

| Function | 1.0 declaration |
|---|---|
| `ghostty_version` | `const char *ghostty_version(void)` |
| `ghostty_formatter_set` | `GhosttyResult ghostty_formatter_set(GhosttyFormatter formatter, GhosttyFormatterOption option, const void* value)` |
| `ghostty_formatter_get` | `GhosttyResult ghostty_formatter_get(GhosttyFormatter formatter, GhosttyFormatterOption option, void* out)` |
| `ghostty_key_encoder_set` | `GhosttyResult ghostty_key_encoder_set(GhosttyKeyEncoder encoder, GhosttyKeyEncoderOption option, const void *value)` |
| `ghostty_key_encoder_set_from_terminal` | `GhosttyResult ghostty_key_encoder_set_from_terminal(GhosttyKeyEncoder encoder, GhosttyTerminal terminal)` |
| `ghostty_mouse_encoder_set` | `GhosttyResult ghostty_mouse_encoder_set(GhosttyMouseEncoder encoder, GhosttyMouseEncoderOption option, const void *value)` |
| `ghostty_mouse_encoder_set_from_terminal` | `GhosttyResult ghostty_mouse_encoder_set_from_terminal(GhosttyMouseEncoder encoder, GhosttyTerminal terminal)` |
| `ghostty_search_select_next` | `GhosttyResult ghostty_search_select_next(GhosttySearch search)` |
| `ghostty_search_select_prev` | `GhosttyResult ghostty_search_select_prev(GhosttySearch search)` |
| `ghostty_sized_init` | `void ghostty_sized_init(void* ptr, size_t size)` |

### C.2 Removed native functions

| Function | Replacement / disposition |
|---|---|
| `ghostty_key_encoder_setopt` | `ghostty_key_encoder_set` |
| `ghostty_key_encoder_setopt_from_terminal` | `ghostty_key_encoder_set_from_terminal` |
| `ghostty_mouse_encoder_setopt` | `ghostty_mouse_encoder_set` |
| `ghostty_mouse_encoder_setopt_from_terminal` | `ghostty_mouse_encoder_set_from_terminal` |
| `ghostty_terminal_selection_format_buf` | `ghostty_formatter_set + ghostty_formatter_format_buf` |
| `ghostty_terminal_selection_format_alloc` | `ghostty_formatter_set + ghostty_formatter_format_alloc` |

### C.3 Changed declarations

| Function | Current declaration | 1.0 declaration | Review classification |
|---|---|---|---|
| `ghostty_color_scheme_report_encode` | `GhosttyResult ghostty_color_scheme_report_encode( GhosttyColorScheme scheme, char* buf, size_t buf_len, size_t* out_written)` | `GhosttyResult ghostty_color_scheme_report_encode( GhosttyColorScheme scheme, uint8_t* buf, size_t buf_len, size_t* out_written)` | Unused today: public contract changed |
| `ghostty_focus_encode` | `GhosttyResult ghostty_focus_encode( GhosttyFocusEvent event, char* buf, size_t buf_len, size_t* out_written)` | `GhosttyResult ghostty_focus_encode( GhosttyFocusEvent event, uint8_t* buf, size_t buf_len, size_t* out_written)` | Unused today: public contract changed |
| `ghostty_formatter_terminal_new` | `GhosttyResult ghostty_formatter_terminal_new( const GhosttyAllocator* allocator, GhosttyFormatter* formatter, GhosttyTerminal terminal, GhosttyFormatterTerminalOptions options)` | `GhosttyResult ghostty_formatter_terminal_new( const GhosttyAllocator* allocator, GhosttyFormatter* formatter, GhosttyTerminal terminal)` | Loaded: must adapt |
| `ghostty_grid_ref_graphemes` | `GhosttyResult ghostty_grid_ref_graphemes(const GhosttyGridRef *ref, uint32_t *buf, size_t buf_len, size_t *out_len)` | `GhosttyResult ghostty_grid_ref_graphemes(const GhosttyGridRef *ref, uint32_t *buf, size_t buf_len, size_t *out_written)` | Parameter name only |
| `ghostty_grid_ref_hyperlink_uri` | `GhosttyResult ghostty_grid_ref_hyperlink_uri( const GhosttyGridRef *ref, uint8_t *buf, size_t buf_len, size_t *out_len)` | `GhosttyResult ghostty_grid_ref_hyperlink_uri( const GhosttyGridRef *ref, uint8_t *buf, size_t buf_len, size_t *out_written)` | Parameter name only |
| `ghostty_key_encoder_encode` | `GhosttyResult ghostty_key_encoder_encode(GhosttyKeyEncoder encoder, GhosttyKeyEvent event, char *out_buf, size_t out_buf_size, size_t *out_len)` | `GhosttyResult ghostty_key_encoder_encode(GhosttyKeyEncoder encoder, GhosttyKeyEvent event, uint8_t *buf, size_t buf_len, size_t *out_written)` | Loaded: must adapt |
| `ghostty_key_event_set_action` | `void ghostty_key_event_set_action(GhosttyKeyEvent event, GhosttyKeyAction action)` | `GhosttyResult ghostty_key_event_set_action(GhosttyKeyEvent event, GhosttyKeyAction action)` | Loaded: must adapt |
| `ghostty_key_event_set_key` | `void ghostty_key_event_set_key(GhosttyKeyEvent event, GhosttyKey key)` | `GhosttyResult ghostty_key_event_set_key(GhosttyKeyEvent event, GhosttyKey key)` | Loaded: must adapt |
| `ghostty_key_event_set_mods` | `void ghostty_key_event_set_mods(GhosttyKeyEvent event, GhosttyMods mods)` | `GhosttyResult ghostty_key_event_set_mods(GhosttyKeyEvent event, GhosttyMods mods)` | Loaded: must adapt |
| `ghostty_key_event_set_consumed_mods` | `void ghostty_key_event_set_consumed_mods(GhosttyKeyEvent event, GhosttyMods consumed_mods)` | `GhosttyResult ghostty_key_event_set_consumed_mods(GhosttyKeyEvent event, GhosttyMods consumed_mods)` | Unused today: public contract changed |
| `ghostty_key_event_set_composing` | `void ghostty_key_event_set_composing(GhosttyKeyEvent event, bool composing)` | `GhosttyResult ghostty_key_event_set_composing(GhosttyKeyEvent event, bool composing)` | Unused today: public contract changed |
| `ghostty_key_event_set_utf8` | `void ghostty_key_event_set_utf8(GhosttyKeyEvent event, const char *utf8, size_t len)` | `GhosttyResult ghostty_key_event_set_utf8(GhosttyKeyEvent event, GhosttyString utf8)` | Loaded: must adapt |
| `ghostty_key_event_get_utf8` | `const char *ghostty_key_event_get_utf8(GhosttyKeyEvent event, size_t *len)` | `GhosttyString ghostty_key_event_get_utf8(GhosttyKeyEvent event)` | Unused today: public contract changed |
| `ghostty_key_event_set_unshifted_codepoint` | `void ghostty_key_event_set_unshifted_codepoint(GhosttyKeyEvent event, uint32_t codepoint)` | `GhosttyResult ghostty_key_event_set_unshifted_codepoint(GhosttyKeyEvent event, uint32_t codepoint)` | Loaded: must adapt |
| `ghostty_mode_report_encode` | `GhosttyResult ghostty_mode_report_encode( GhosttyMode mode, GhosttyModeReportState state, char* buf, size_t buf_len, size_t* out_written)` | `GhosttyResult ghostty_mode_report_encode( GhosttyMode mode, GhosttyModeReportState state, uint8_t* buf, size_t buf_len, size_t* out_written)` | Unused today: public contract changed |
| `ghostty_mouse_encoder_reset` | `void ghostty_mouse_encoder_reset(GhosttyMouseEncoder encoder)` | `GhosttyResult ghostty_mouse_encoder_reset(GhosttyMouseEncoder encoder)` | Unused today: public contract changed |
| `ghostty_mouse_encoder_encode` | `GhosttyResult ghostty_mouse_encoder_encode(GhosttyMouseEncoder encoder, GhosttyMouseEvent event, char *out_buf, size_t out_buf_size, size_t *out_len)` | `GhosttyResult ghostty_mouse_encoder_encode(GhosttyMouseEncoder encoder, GhosttyMouseEvent event, uint8_t *buf, size_t buf_len, size_t *out_written)` | Loaded: must adapt |
| `ghostty_mouse_event_set_action` | `void ghostty_mouse_event_set_action(GhosttyMouseEvent event, GhosttyMouseAction action)` | `GhosttyResult ghostty_mouse_event_set_action(GhosttyMouseEvent event, GhosttyMouseAction action)` | Loaded: must adapt |
| `ghostty_mouse_event_set_button` | `void ghostty_mouse_event_set_button(GhosttyMouseEvent event, GhosttyMouseButton button)` | `GhosttyResult ghostty_mouse_event_set_button(GhosttyMouseEvent event, GhosttyMouseButton button)` | Loaded: must adapt |
| `ghostty_mouse_event_clear_button` | `void ghostty_mouse_event_clear_button(GhosttyMouseEvent event)` | `GhosttyResult ghostty_mouse_event_clear_button(GhosttyMouseEvent event)` | Loaded: must adapt |
| `ghostty_mouse_event_get_button` | `bool ghostty_mouse_event_get_button(GhosttyMouseEvent event, GhosttyMouseButton *out_button)` | `GhosttyResult ghostty_mouse_event_get_button(GhosttyMouseEvent event, GhosttyMouseButton *out_button)` | Unused today: public contract changed |
| `ghostty_mouse_event_set_mods` | `void ghostty_mouse_event_set_mods(GhosttyMouseEvent event, GhosttyMods mods)` | `GhosttyResult ghostty_mouse_event_set_mods(GhosttyMouseEvent event, GhosttyMods mods)` | Loaded: must adapt |
| `ghostty_mouse_event_set_position` | `void ghostty_mouse_event_set_position(GhosttyMouseEvent event, GhosttyMousePosition position)` | `GhosttyResult ghostty_mouse_event_set_position(GhosttyMouseEvent event, GhosttyMousePosition position)` | Loaded: must adapt |
| `ghostty_osc_command_data` | `bool ghostty_osc_command_data(GhosttyOscCommand command, GhosttyOscCommandData data, void *out)` | `GhosttyResult ghostty_osc_command_data(GhosttyOscCommand command, GhosttyOscCommandData data, void *out)` | Unused today: public contract changed |
| `ghostty_terminal_paste` | `GhosttyResult ghostty_terminal_paste( GhosttyTerminal terminal, const GhosttyPaste* paste, bool* out_written)` | `GhosttyResult ghostty_terminal_paste( GhosttyTerminal terminal, const GhosttyPaste* paste, bool* out_pasted)` | Parameter name only |
| `ghostty_paste_encode` | `GhosttyResult ghostty_paste_encode( char* data, size_t data_len, bool bracketed, char* buf, size_t buf_len, size_t* out_written)` | `GhosttyResult ghostty_paste_encode( char* data, size_t data_len, bool bracketed, uint8_t* buf, size_t buf_len, size_t* out_written)` | Loaded: must adapt |
| `ghostty_terminal_select_output` | `GhosttyResult ghostty_terminal_select_output( GhosttyTerminal terminal, GhosttyGridRef ref, GhosttySelection* out_selection)` | `GhosttyResult ghostty_terminal_select_output( GhosttyTerminal terminal, const GhosttyGridRef* ref, GhosttySelection* out_selection)` | Unused today: public contract changed |
| `ghostty_sgr_attribute_tag` | `GhosttySgrAttributeTag ghostty_sgr_attribute_tag(GhosttySgrAttribute attr)` | `GhosttySgrAttributeTag ghostty_sgr_attribute_tag( const GhosttySgrAttribute* attr)` | Unused today: public contract changed |
| `ghostty_size_report_encode` | `GhosttyResult ghostty_size_report_encode( GhosttySizeReportStyle style, GhosttySizeReportSize size, char* buf, size_t buf_len, size_t* out_written)` | `GhosttyResult ghostty_size_report_encode( GhosttySizeReportStyle style, const GhosttySizeReportSize* size, uint8_t* buf, size_t buf_len, size_t* out_written)` | Unused today: public contract changed |
| `ghostty_terminal_scroll_viewport` | `void ghostty_terminal_scroll_viewport(GhosttyTerminal terminal, GhosttyTerminalScrollViewport behavior)` | `GhosttyResult ghostty_terminal_scroll_viewport( GhosttyTerminal terminal, GhosttyTerminalScrollViewport behavior)` | Loaded: must adapt |

## Appendix D. Complete enum change review

Enums were compared by member name and effective value, including old implicit values. Every changed existing enum group and all added/removed enum types are recorded. The `INT_MAX` sentinel is shown as 2147483647; it is a type-width guard, not an allocated usable option number.

### GhosttyBuildInfo

| Member | Current | 1.0 | Change |
|---|---:|---:|---|
| `GHOSTTY_BUILD_INFO_FEATURE_SNAPSHOT` | — | `11` | Added/renamed |
| `GHOSTTY_BUILD_INFO_FEATURE_FORMATTER` | — | `12` | Added/renamed |
| `GHOSTTY_BUILD_INFO_FEATURE_SELECTION` | — | `13` | Added/renamed |
| `GHOSTTY_BUILD_INFO_FEATURE_SEARCH` | — | `14` | Added/renamed |
| `GHOSTTY_BUILD_INFO_FEATURE_RENDER_STATE` | — | `15` | Added/renamed |
| `GHOSTTY_BUILD_INFO_FEATURE_INPUT_ENCODE` | — | `16` | Added/renamed |
| `GHOSTTY_BUILD_INFO_FEATURE_COLOR` | — | `17` | Added/renamed |
| `GHOSTTY_BUILD_INFO_FEATURE_GRID_INTROSPECTION` | — | `18` | Added/renamed |
| `GHOSTTY_BUILD_INFO_FEATURE_GLYPH_PROTOCOL` | — | `19` | Added/renamed |

### GhosttyKeyEncoderOption

| Member | Current | 1.0 | Change |
|---|---:|---:|---|
| `GHOSTTY_KEY_ENCODER_OPT_CURSOR_KEY_APPLICATION` | `0` | `1` | Value changed |
| `GHOSTTY_KEY_ENCODER_OPT_KEYPAD_KEY_APPLICATION` | `1` | `2` | Value changed |
| `GHOSTTY_KEY_ENCODER_OPT_IGNORE_KEYPAD_WITH_NUMLOCK` | `2` | `3` | Value changed |
| `GHOSTTY_KEY_ENCODER_OPT_ALT_ESC_PREFIX` | `3` | `4` | Value changed |
| `GHOSTTY_KEY_ENCODER_OPT_MODIFY_OTHER_KEYS_STATE_2` | `4` | `5` | Value changed |
| `GHOSTTY_KEY_ENCODER_OPT_KITTY_FLAGS` | `5` | `6` | Value changed |
| `GHOSTTY_KEY_ENCODER_OPT_MACOS_OPTION_AS_ALT` | `6` | `7` | Value changed |
| `GHOSTTY_KEY_ENCODER_OPT_BACKARROW_KEY_MODE` | `7` | `8` | Value changed |
| `GHOSTTY_KEY_ENCODER_OPT_INVALID` | — | `0` | Added/renamed |

### GhosttyKittyGraphicsPlacementIteratorOption

| Member | Current | 1.0 | Change |
|---|---:|---:|---|
| `GHOSTTY_KITTY_GRAPHICS_PLACEMENT_ITERATOR_OPTION_LAYER` | `0` | — | Removed/renamed |
| `GHOSTTY_KITTY_GRAPHICS_PLACEMENT_ITERATOR_OPTION_MAX_VALUE` | `2147483647` | — | Removed/renamed |
| `GHOSTTY_KITTY_GRAPHICS_PLACEMENT_ITERATOR_OPT_INVALID` | — | `0` | Added/renamed |
| `GHOSTTY_KITTY_GRAPHICS_PLACEMENT_ITERATOR_OPT_LAYER` | — | `1` | Added/renamed |
| `GHOSTTY_KITTY_GRAPHICS_PLACEMENT_ITERATOR_OPT_MAX_VALUE` | — | `2147483647` | Added/renamed |

### GhosttyKittyGraphicsImageData

| Member | Current | 1.0 | Change |
|---|---:|---:|---|
| `GHOSTTY_KITTY_IMAGE_DATA_INVALID` | `0` | — | Removed/renamed |
| `GHOSTTY_KITTY_IMAGE_DATA_ID` | `1` | — | Removed/renamed |
| `GHOSTTY_KITTY_IMAGE_DATA_NUMBER` | `2` | — | Removed/renamed |
| `GHOSTTY_KITTY_IMAGE_DATA_WIDTH` | `3` | — | Removed/renamed |
| `GHOSTTY_KITTY_IMAGE_DATA_HEIGHT` | `4` | — | Removed/renamed |
| `GHOSTTY_KITTY_IMAGE_DATA_FORMAT` | `5` | — | Removed/renamed |
| `GHOSTTY_KITTY_IMAGE_DATA_COMPRESSION` | `6` | — | Removed/renamed |
| `GHOSTTY_KITTY_IMAGE_DATA_DATA_PTR` | `7` | — | Removed/renamed |
| `GHOSTTY_KITTY_IMAGE_DATA_DATA_LEN` | `8` | — | Removed/renamed |
| `GHOSTTY_KITTY_IMAGE_DATA_GENERATION` | `9` | — | Removed/renamed |
| `GHOSTTY_KITTY_IMAGE_DATA_MAX_VALUE` | `2147483647` | — | Removed/renamed |
| `GHOSTTY_KITTY_GRAPHICS_IMAGE_DATA_INVALID` | — | `0` | Added/renamed |
| `GHOSTTY_KITTY_GRAPHICS_IMAGE_DATA_ID` | — | `1` | Added/renamed |
| `GHOSTTY_KITTY_GRAPHICS_IMAGE_DATA_NUMBER` | — | `2` | Added/renamed |
| `GHOSTTY_KITTY_GRAPHICS_IMAGE_DATA_WIDTH` | — | `3` | Added/renamed |
| `GHOSTTY_KITTY_GRAPHICS_IMAGE_DATA_HEIGHT` | — | `4` | Added/renamed |
| `GHOSTTY_KITTY_GRAPHICS_IMAGE_DATA_FORMAT` | — | `5` | Added/renamed |
| `GHOSTTY_KITTY_GRAPHICS_IMAGE_DATA_COMPRESSION` | — | `6` | Added/renamed |
| `GHOSTTY_KITTY_GRAPHICS_IMAGE_DATA_PIXELS` | — | `7` | Added/renamed |
| `GHOSTTY_KITTY_GRAPHICS_IMAGE_DATA_EXPECTED_LEN` | — | `8` | Added/renamed |
| `GHOSTTY_KITTY_GRAPHICS_IMAGE_DATA_GENERATION` | — | `9` | Added/renamed |
| `GHOSTTY_KITTY_GRAPHICS_IMAGE_DATA_MAX_VALUE` | — | `2147483647` | Added/renamed |

### GhosttyMouseEncoderOption

| Member | Current | 1.0 | Change |
|---|---:|---:|---|
| `GHOSTTY_MOUSE_ENCODER_OPT_EVENT` | `0` | `1` | Value changed |
| `GHOSTTY_MOUSE_ENCODER_OPT_FORMAT` | `1` | `2` | Value changed |
| `GHOSTTY_MOUSE_ENCODER_OPT_SIZE` | `2` | `3` | Value changed |
| `GHOSTTY_MOUSE_ENCODER_OPT_ANY_BUTTON_PRESSED` | `3` | `4` | Value changed |
| `GHOSTTY_MOUSE_ENCODER_OPT_TRACK_LAST_CELL` | `4` | `5` | Value changed |
| `GHOSTTY_MOUSE_ENCODER_OPT_INVALID` | — | `0` | Added/renamed |

### GhosttyOscOption

| Member | Current | 1.0 | Change |
|---|---:|---:|---|
| `GHOSTTY_OSC_OPT_UNKNOWN_MAX_BYTES` | `0` | `1` | Value changed |
| `GHOSTTY_OSC_OPT_INVALID` | — | `0` | Added/renamed |

### GhosttyRenderStateOption

| Member | Current | 1.0 | Change |
|---|---:|---:|---|
| `GHOSTTY_RENDER_STATE_OPTION_DIRTY` | `0` | — | Removed/renamed |
| `GHOSTTY_RENDER_STATE_OPTION_OVERSCAN` | `1` | — | Removed/renamed |
| `GHOSTTY_RENDER_STATE_OPTION_MAX_VALUE` | `2147483647` | — | Removed/renamed |
| `GHOSTTY_RENDER_STATE_OPT_INVALID` | — | `0` | Added/renamed |
| `GHOSTTY_RENDER_STATE_OPT_DIRTY` | — | `1` | Added/renamed |
| `GHOSTTY_RENDER_STATE_OPT_OVERSCAN` | — | `2` | Added/renamed |
| `GHOSTTY_RENDER_STATE_OPT_MAX_VALUE` | — | `2147483647` | Added/renamed |

### GhosttyRenderStateRowOption

| Member | Current | 1.0 | Change |
|---|---:|---:|---|
| `GHOSTTY_RENDER_STATE_ROW_OPTION_DIRTY` | `0` | — | Removed/renamed |
| `GHOSTTY_RENDER_STATE_ROW_OPTION_MAX_VALUE` | `2147483647` | — | Removed/renamed |
| `GHOSTTY_RENDER_STATE_ROW_OPT_INVALID` | — | `0` | Added/renamed |
| `GHOSTTY_RENDER_STATE_ROW_OPT_DIRTY` | — | `1` | Added/renamed |
| `GHOSTTY_RENDER_STATE_ROW_OPT_MAX_VALUE` | — | `2147483647` | Added/renamed |

### GhosttySearchData

| Member | Current | 1.0 | Change |
|---|---:|---:|---|
| `GHOSTTY_SEARCH_DATA_STATUS` | `0` | `1` | Value changed |
| `GHOSTTY_SEARCH_DATA_NEEDLE` | `1` | `2` | Value changed |
| `GHOSTTY_SEARCH_DATA_TOTAL_MATCHES` | `2` | `3` | Value changed |
| `GHOSTTY_SEARCH_DATA_SELECTED_INDEX` | `3` | `4` | Value changed |
| `GHOSTTY_SEARCH_DATA_SELECTED_MATCH` | `4` | `5` | Value changed |
| `GHOSTTY_SEARCH_DATA_MATCHES` | `5` | `6` | Value changed |
| `GHOSTTY_SEARCH_DATA_VIEWPORT_MATCHES` | `6` | `7` | Value changed |
| `GHOSTTY_SEARCH_DATA_SELECT_SCROLL` | `7` | `8` | Value changed |
| `GHOSTTY_SEARCH_DATA_INVALID` | — | `0` | Added/renamed |

### GhosttySearchOption

| Member | Current | 1.0 | Change |
|---|---:|---:|---|
| `GHOSTTY_SEARCH_OPT_NEEDLE` | `0` | `1` | Value changed |
| `GHOSTTY_SEARCH_OPT_SELECT_SCROLL` | `3` | `4` | Value changed |
| `GHOSTTY_SEARCH_OPT_SELECT_NEXT` | `1` | — | Removed/renamed |
| `GHOSTTY_SEARCH_OPT_SELECT_PREV` | `2` | — | Removed/renamed |
| `GHOSTTY_SEARCH_OPT_INVALID` | — | `0` | Added/renamed |

### GhosttySelectionGestureData

| Member | Current | 1.0 | Change |
|---|---:|---:|---|
| `GHOSTTY_SELECTION_GESTURE_DATA_CLICK_COUNT` | `0` | `1` | Value changed |
| `GHOSTTY_SELECTION_GESTURE_DATA_DRAGGED` | `1` | `2` | Value changed |
| `GHOSTTY_SELECTION_GESTURE_DATA_AUTOSCROLL` | `2` | `3` | Value changed |
| `GHOSTTY_SELECTION_GESTURE_DATA_BEHAVIOR` | `3` | `4` | Value changed |
| `GHOSTTY_SELECTION_GESTURE_DATA_ANCHOR` | `4` | `5` | Value changed |
| `GHOSTTY_SELECTION_GESTURE_DATA_INVALID` | — | `0` | Added/renamed |

### GhosttySelectionGestureEventOption

| Member | Current | 1.0 | Change |
|---|---:|---:|---|
| `GHOSTTY_SELECTION_GESTURE_EVENT_OPT_REF` | `0` | `1` | Value changed |
| `GHOSTTY_SELECTION_GESTURE_EVENT_OPT_POSITION` | `1` | `2` | Value changed |
| `GHOSTTY_SELECTION_GESTURE_EVENT_OPT_REPEAT_DISTANCE` | `2` | `3` | Value changed |
| `GHOSTTY_SELECTION_GESTURE_EVENT_OPT_TIME_NS` | `3` | `4` | Value changed |
| `GHOSTTY_SELECTION_GESTURE_EVENT_OPT_REPEAT_INTERVAL_NS` | `4` | `5` | Value changed |
| `GHOSTTY_SELECTION_GESTURE_EVENT_OPT_WORD_BOUNDARY_CODEPOINTS` | `5` | `6` | Value changed |
| `GHOSTTY_SELECTION_GESTURE_EVENT_OPT_BEHAVIORS` | `6` | `7` | Value changed |
| `GHOSTTY_SELECTION_GESTURE_EVENT_OPT_RECTANGLE` | `7` | `8` | Value changed |
| `GHOSTTY_SELECTION_GESTURE_EVENT_OPT_GEOMETRY` | `8` | `9` | Value changed |
| `GHOSTTY_SELECTION_GESTURE_EVENT_OPT_VIEWPORT` | `9` | `10` | Value changed |
| `GHOSTTY_SELECTION_GESTURE_EVENT_OPT_INVALID` | — | `0` | Added/renamed |

### GhosttySnapshotDecoderOption

| Member | Current | 1.0 | Change |
|---|---:|---:|---|
| `GHOSTTY_SNAPSHOT_DECODER_OPT_MAX_CONTINUATION_BYTES` | `0` | `1` | Value changed |
| `GHOSTTY_SNAPSHOT_DECODER_OPT_RETAIN_CONTINUATION` | `1` | `2` | Value changed |
| `GHOSTTY_SNAPSHOT_DECODER_OPT_COMPRESS_HISTORY` | `2` | `3` | Value changed |
| `GHOSTTY_SNAPSHOT_DECODER_OPT_KITTY_IMAGE_STORAGE_LIMIT` | `3` | `4` | Value changed |
| `GHOSTTY_SNAPSHOT_DECODER_OPT_INVALID` | — | `0` | Added/renamed |

### GhosttySysOption

| Member | Current | 1.0 | Change |
|---|---:|---:|---|
| `GHOSTTY_SYS_OPT_USERDATA` | `0` | `1` | Value changed |
| `GHOSTTY_SYS_OPT_DECODE_PNG` | `1` | `2` | Value changed |
| `GHOSTTY_SYS_OPT_LOG` | `2` | `3` | Value changed |
| `GHOSTTY_SYS_OPT_RANDOM_SECURE` | `3` | `4` | Value changed |
| `GHOSTTY_SYS_OPT_INVALID` | — | `0` | Added/renamed |

### GhosttyTerminalScrollViewportTag

| Member | Current | 1.0 | Change |
|---|---:|---:|---|
| `GHOSTTY_SCROLL_VIEWPORT_TOP` | `0` | — | Removed/renamed |
| `GHOSTTY_SCROLL_VIEWPORT_BOTTOM` | `1` | — | Removed/renamed |
| `GHOSTTY_SCROLL_VIEWPORT_DELTA` | `2` | — | Removed/renamed |
| `GHOSTTY_SCROLL_VIEWPORT_ROW` | `3` | — | Removed/renamed |
| `GHOSTTY_SCROLL_VIEWPORT_DELTA_PROMPT` | `4` | — | Removed/renamed |
| `GHOSTTY_SCROLL_VIEWPORT_MAX_VALUE` | `2147483647` | — | Removed/renamed |
| `GHOSTTY_TERMINAL_SCROLL_VIEWPORT_TOP` | — | `0` | Added/renamed |
| `GHOSTTY_TERMINAL_SCROLL_VIEWPORT_BOTTOM` | — | `1` | Added/renamed |
| `GHOSTTY_TERMINAL_SCROLL_VIEWPORT_DELTA` | — | `2` | Added/renamed |
| `GHOSTTY_TERMINAL_SCROLL_VIEWPORT_ROW` | — | `3` | Added/renamed |
| `GHOSTTY_TERMINAL_SCROLL_VIEWPORT_DELTA_PROMPT` | — | `4` | Added/renamed |
| `GHOSTTY_TERMINAL_SCROLL_VIEWPORT_MAX_VALUE` | — | `2147483647` | Added/renamed |

### GhosttyTerminalOption

| Member | Current | 1.0 | Change |
|---|---:|---:|---|
| `GHOSTTY_TERMINAL_OPT_USERDATA` | `0` | `1` | Value changed |
| `GHOSTTY_TERMINAL_OPT_WRITE_PTY` | `1` | `2` | Value changed |
| `GHOSTTY_TERMINAL_OPT_BELL` | `2` | `3` | Value changed |
| `GHOSTTY_TERMINAL_OPT_ENQUIRY` | `3` | `4` | Value changed |
| `GHOSTTY_TERMINAL_OPT_XTVERSION` | `4` | `5` | Value changed |
| `GHOSTTY_TERMINAL_OPT_TITLE_CHANGED` | `5` | `6` | Value changed |
| `GHOSTTY_TERMINAL_OPT_SIZE` | `6` | `7` | Value changed |
| `GHOSTTY_TERMINAL_OPT_COLOR_SCHEME` | `7` | `8` | Value changed |
| `GHOSTTY_TERMINAL_OPT_DEVICE_ATTRIBUTES` | `8` | `9` | Value changed |
| `GHOSTTY_TERMINAL_OPT_TITLE` | `9` | `10` | Value changed |
| `GHOSTTY_TERMINAL_OPT_PWD` | `10` | `11` | Value changed |
| `GHOSTTY_TERMINAL_OPT_COLOR_FOREGROUND` | `11` | `12` | Value changed |
| `GHOSTTY_TERMINAL_OPT_COLOR_BACKGROUND` | `12` | `13` | Value changed |
| `GHOSTTY_TERMINAL_OPT_COLOR_CURSOR` | `13` | `14` | Value changed |
| `GHOSTTY_TERMINAL_OPT_COLOR_PALETTE` | `14` | `15` | Value changed |
| `GHOSTTY_TERMINAL_OPT_KITTY_IMAGE_STORAGE_LIMIT` | `15` | `16` | Value changed |
| `GHOSTTY_TERMINAL_OPT_KITTY_IMAGE_MEDIUM_FILE` | `16` | `17` | Value changed |
| `GHOSTTY_TERMINAL_OPT_KITTY_IMAGE_MEDIUM_TEMP_FILE` | `17` | `18` | Value changed |
| `GHOSTTY_TERMINAL_OPT_KITTY_IMAGE_MEDIUM_SHARED_MEM` | `18` | `19` | Value changed |
| `GHOSTTY_TERMINAL_OPT_APC_MAX_BYTES` | `19` | `20` | Value changed |
| `GHOSTTY_TERMINAL_OPT_APC_MAX_BYTES_KITTY` | `20` | `21` | Value changed |
| `GHOSTTY_TERMINAL_OPT_SELECTION` | `21` | `22` | Value changed |
| `GHOSTTY_TERMINAL_OPT_DEFAULT_CURSOR_STYLE` | `22` | `23` | Value changed |
| `GHOSTTY_TERMINAL_OPT_DEFAULT_CURSOR_BLINK` | `23` | `24` | Value changed |
| `GHOSTTY_TERMINAL_OPT_GLYPH_PROTOCOL` | `24` | `25` | Value changed |
| `GHOSTTY_TERMINAL_OPT_PWD_CHANGED` | `25` | `26` | Value changed |
| `GHOSTTY_TERMINAL_OPT_CLIPBOARD_WRITE` | `26` | `27` | Value changed |
| `GHOSTTY_TERMINAL_OPT_SCROLLBACK_MAX_BYTES` | `27` | `28` | Value changed |
| `GHOSTTY_TERMINAL_OPT_SCROLLBACK_MAX_LINES` | `28` | `29` | Value changed |
| `GHOSTTY_TERMINAL_OPT_DESKTOP_NOTIFICATION` | `29` | `30` | Value changed |
| `GHOSTTY_TERMINAL_OPT_PROGRESS_REPORT` | `30` | `31` | Value changed |
| `GHOSTTY_TERMINAL_OPT_CONTINUATION_MAX_BYTES` | `31` | `32` | Value changed |
| `GHOSTTY_TERMINAL_OPT_TITLE_REPORT` | `32` | `33` | Value changed |
| `GHOSTTY_TERMINAL_OPT_MODE_DEFAULT` | `33` | `34` | Value changed |
| `GHOSTTY_TERMINAL_OPT_MODE` | `34` | `35` | Value changed |
| `GHOSTTY_TERMINAL_OPT_UNKNOWN_SEQUENCE` | `35` | `36` | Value changed |
| `GHOSTTY_TERMINAL_OPT_UNKNOWN_MAX_BYTES` | `36` | `37` | Value changed |
| `GHOSTTY_TERMINAL_OPT_TERMINFO_NAME` | `37` | `38` | Value changed |
| `GHOSTTY_TERMINAL_OPT_CLIPBOARD_READ` | `38` | `39` | Value changed |
| `GHOSTTY_TERMINAL_OPT_CLIPBOARD_WRITE_MAX_BYTES` | `39` | `40` | Value changed |
| `GHOSTTY_TERMINAL_OPT_RESIZE_PULL_SCROLLBACK` | `40` | `41` | Value changed |
| `GHOSTTY_TERMINAL_OPT_RENDER_HOLD` | `41` | `42` | Value changed |
| `GHOSTTY_TERMINAL_OPT_SEMANTIC_PROMPT` | `42` | `43` | Value changed |
| `GHOSTTY_TERMINAL_OPT_RESET` | `43` | `44` | Value changed |
| `GHOSTTY_TERMINAL_OPT_XT_CHECKSUM_REPORT` | `44` | `45` | Value changed |
| `GHOSTTY_TERMINAL_OPT_XT_CHECKSUM_EXTENSION` | `45` | `46` | Value changed |
| `GHOSTTY_TERMINAL_OPT_PROGRAM_STATUS` | `46` | `47` | Value changed |
| `GHOSTTY_TERMINAL_OPT_RESIZE_WINDOW` | `47` | `48` | Value changed |
| `GHOSTTY_TERMINAL_OPT_INVALID` | — | `0` | Added/renamed |

### GhosttyTerminalData

| Member | Current | 1.0 | Change |
|---|---:|---:|---|
| `GHOSTTY_TERMINAL_DATA_CURSOR_STYLE` | `10` | — | Removed/renamed |
| `GHOSTTY_TERMINAL_DATA_CURSOR_SGR_STYLE` | — | `10` | Added/renamed |

### GhosttyResult

| Member | Current | 1.0 | Change |
|---|---:|---:|---|
| `GHOSTTY_UNSUPPORTED` | — | `-8` | Added/renamed |

### Added enum types

- `GhosttyFormatterOption`: `GHOSTTY_FORMATTER_OPT_INVALID`=`0`; `GHOSTTY_FORMATTER_OPT_EMIT`=`1`; `GHOSTTY_FORMATTER_OPT_UNWRAP`=`2`; `GHOSTTY_FORMATTER_OPT_TRIM`=`3`; `GHOSTTY_FORMATTER_OPT_SELECTION`=`4`; `GHOSTTY_FORMATTER_OPT_EXTRAS`=`5`; `GHOSTTY_FORMATTER_OPT_MAX_VALUE`=`2147483647`.
- `GhosttyCursorStyle`: `GHOSTTY_CURSOR_STYLE_BAR`=`0`; `GHOSTTY_CURSOR_STYLE_BLOCK`=`1`; `GHOSTTY_CURSOR_STYLE_UNDERLINE`=`2`; `GHOSTTY_CURSOR_STYLE_BLOCK_HOLLOW`=`3`; `GHOSTTY_CURSOR_STYLE_MAX_VALUE`=`2147483647`.
- `GhosttyImageLayer`: `GHOSTTY_IMAGE_LAYER_BELOW_BACKGROUND`=`0`; `GHOSTTY_IMAGE_LAYER_BELOW_TEXT`=`1`; `GHOSTTY_IMAGE_LAYER_ABOVE_TEXT`=`2`; `GHOSTTY_IMAGE_LAYER_MAX_VALUE`=`2147483647`.

### Removed enum types

- `GhosttyKittyPlacementLayer`: `GHOSTTY_KITTY_PLACEMENT_LAYER_ALL`=`0`; `GHOSTTY_KITTY_PLACEMENT_LAYER_BELOW_BG`=`1`; `GHOSTTY_KITTY_PLACEMENT_LAYER_BELOW_TEXT`=`2`; `GHOSTTY_KITTY_PLACEMENT_LAYER_ABOVE_TEXT`=`3`; `GHOSTTY_KITTY_PLACEMENT_LAYER_MAX_VALUE`=`2147483647`.
- `GhosttyRenderStateCursorVisualStyle`: `GHOSTTY_RENDER_STATE_CURSOR_VISUAL_STYLE_BAR`=`0`; `GHOSTTY_RENDER_STATE_CURSOR_VISUAL_STYLE_BLOCK`=`1`; `GHOSTTY_RENDER_STATE_CURSOR_VISUAL_STYLE_UNDERLINE`=`2`; `GHOSTTY_RENDER_STATE_CURSOR_VISUAL_STYLE_BLOCK_HOLLOW`=`3`; `GHOSTTY_RENDER_STATE_CURSOR_VISUAL_STYLE_MAX_VALUE`=`2147483647`.
- `GhosttyRenderStateImageLayer`: `GHOSTTY_RENDER_STATE_IMAGE_LAYER_BELOW_BACKGROUND`=`0`; `GHOSTTY_RENDER_STATE_IMAGE_LAYER_BELOW_TEXT`=`1`; `GHOSTTY_RENDER_STATE_IMAGE_LAYER_ABOVE_TEXT`=`2`; `GHOSTTY_RENDER_STATE_IMAGE_LAYER_MAX_VALUE`=`2147483647`.
- `GhosttyTerminalCursorStyle`: `GHOSTTY_TERMINAL_CURSOR_STYLE_BAR`=`0`; `GHOSTTY_TERMINAL_CURSOR_STYLE_BLOCK`=`1`; `GHOSTTY_TERMINAL_CURSOR_STYLE_UNDERLINE`=`2`; `GHOSTTY_TERMINAL_CURSOR_STYLE_BLOCK_HOLLOW`=`3`; `GHOSTTY_TERMINAL_CURSOR_STYLE_MAX_VALUE`=`2147483647`.

### Existing enum groups with unchanged names and values

`GhosttyOptimizeMode`, `GhosttyColorScheme`, `GhosttyFocusEvent`, `GhosttyOptionAsAlt`, `GhosttyKeyAction`, `GhosttyKey`, `GhosttyKittyGraphicsData`, `GhosttyKittyGraphicsPlacementData`, `GhosttyKittyImageFormat`, `GhosttyKittyImageCompression`, `GhosttyModeReportState`, `GhosttyMouseShape`, `GhosttyMouseTrackingMode`, `GhosttyMouseFormat`, `GhosttyMouseAction`, `GhosttyMouseButton`, `GhosttyOscCommandType`, `GhosttyOscTerminator`, `GhosttyOscCommandData`, `GhosttyPasteSource`, `GhosttyPointTag`, `GhosttyRenderStateDirty`, `GhosttyRenderStateData`, `GhosttyImageProtocol`, `GhosttyRenderStateKittyPlacementKind`, `GhosttyRenderStateImageData`, `GhosttyRenderStateRowData`, `GhosttyRenderStateRowCellsData`, `GhosttyCellContentTag`, `GhosttyCellWide`, `GhosttyCellSemanticContent`, `GhosttyCellData`, `GhosttyRowSemanticPrompt`, `GhosttyRowData`, `GhosttySearchStatus`, `GhosttySearchScroll`, `GhosttySelectionOrder`, `GhosttySelectionAdjust`, `GhosttySelectionGestureBehavior`, `GhosttySelectionGestureAutoscroll`, `GhosttySelectionGestureEventType`, `GhosttySgrAttributeTag`, `GhosttySgrUnderline`, `GhosttySizeReportStyle`, `GhosttySnapshotDecoderProgressKind`, `GhosttySnapshotDecoderData`, `GhosttyStyleColorTag`, `GhosttySysLogLevel`, `GhosttyTerminalCompressionMode`, `GhosttyTerminalCompressionResult`, `GhosttyTerminalScreen`, `GhosttyTerminalBoundary`, `GhosttyTerminalStringKind`, `GhosttyTerminalStringIntroducer`, `GhosttyTerminalStringEnd`, `GhosttyTerminalUnknownSequenceTag`, `GhosttyClipboardLocation`, `GhosttyClipboardWriteResult`, `GhosttyClipboardReadResult`, `GhosttyTerminalProgressState`, `GhosttyProgramStatusState`, `GhosttyProgramStatusKind`, `GhosttySemanticPromptKind`, `GhosttySemanticPromptPromptKind`, `GhosttyTerminalResizeWindowTag`, `GhosttyFormatterFormat`.

`GhosttyKey` retains all 177 entries including sentinels; its new explicit assignments preserve the old implicit values. `GhosttyFormatterFormat` retains its values while moving to the common types header. The private extensions in section 5 are compared separately from the unmodified public headers.


## Appendix E. Complete public typedef and aggregate review

All 216 target public typedefs are listed, plus the eight removed typedefs. A named callback declaration can remain unchanged while an argument pointee changes layout. “No declaration change” therefore still requires the native checks for every used ABI boundary.

| Target typedef | Declaration review | Mentioned by current Terminal sources | Public header |
|---|---|---|---|
| `GhosttyAllocator` | No declaration change | Yes | `allocator.h` |
| `GhosttyAllocatorAllocFn` | Added | No | `allocator.h` |
| `GhosttyAllocatorFreeFn` | Added | No | `allocator.h` |
| `GhosttyAllocatorRemapFn` | Added | No | `allocator.h` |
| `GhosttyAllocatorResizeFn` | Added | No | `allocator.h` |
| `GhosttyAllocatorVtable` | Changed (see section 4 / enum inventory) | No | `allocator.h` |
| `GhosttyBuffer` | No declaration change | Yes | `types.h` |
| `GhosttyBuildInfo` | Changed (see section 4 / enum inventory) | No | `build_info.h` |
| `GhosttyCell` | No declaration change | Yes | `screen.h` |
| `GhosttyCellContentTag` | No declaration change | No | `screen.h` |
| `GhosttyCellData` | No declaration change | No | `screen.h` |
| `GhosttyCellSemanticContent` | No declaration change | No | `screen.h` |
| `GhosttyCellWide` | No declaration change | Yes | `screen.h` |
| `GhosttyCellsView` | No declaration change | No | `screen.h` |
| `GhosttyClipboardContent` | No declaration change | Yes | `terminal.h` |
| `GhosttyClipboardLocation` | No declaration change | Yes | `terminal.h` |
| `GhosttyClipboardRead` | No declaration change | No | `terminal.h` |
| `GhosttyClipboardReadReply` | No declaration change | No | `terminal.h` |
| `GhosttyClipboardReadReplyFn` | No declaration change | No | `terminal.h` |
| `GhosttyClipboardReadResult` | No declaration change | No | `terminal.h` |
| `GhosttyClipboardWrite` | No declaration change | Yes | `terminal.h` |
| `GhosttyClipboardWriteReply` | No declaration change | Yes | `terminal.h` |
| `GhosttyClipboardWriteReplyFn` | No declaration change | No | `terminal.h` |
| `GhosttyClipboardWriteResult` | No declaration change | Yes | `terminal.h` |
| `GhosttyCodepointBuffer` | Added | No | `types.h` |
| `GhosttyCodepoints` | No declaration change | No | `types.h` |
| `GhosttyColorPaletteIndex` | No declaration change | No | `color.h` |
| `GhosttyColorPaletteMask` | No declaration change | No | `color.h` |
| `GhosttyColorRgb` | No declaration change | Yes | `color.h` |
| `GhosttyColorScheme` | No declaration change | No | `device.h` |
| `GhosttyColorX11Entry` | No declaration change | No | `color.h` |
| `GhosttyCursorStyle` | Added | No | `types.h` |
| `GhosttyDeviceAttributes` | No declaration change | No | `device.h` |
| `GhosttyDeviceAttributesPrimary` | No declaration change | No | `device.h` |
| `GhosttyDeviceAttributesSecondary` | No declaration change | No | `device.h` |
| `GhosttyDeviceAttributesTertiary` | No declaration change | No | `device.h` |
| `GhosttyFocusEvent` | No declaration change | No | `focus.h` |
| `GhosttyFormatter` | No declaration change | Yes | `types.h` |
| `GhosttyFormatterExtras` | Added | No | `formatter.h` |
| `GhosttyFormatterFormat` | Text changed; enum values preserved | No | `types.h` |
| `GhosttyFormatterOption` | Added | No | `formatter.h` |
| `GhosttyGridRef` | Changed (see section 4 / enum inventory) | Yes | `grid_ref.h` |
| `GhosttyImageLayer` | Added | No | `types.h` |
| `GhosttyImageProtocol` | No declaration change | No | `render.h` |
| `GhosttyKey` | Text changed; enum values preserved | Yes | `key/event.h` |
| `GhosttyKeyAction` | No declaration change | Yes | `key/event.h` |
| `GhosttyKeyEncoder` | No declaration change | Yes | `key/encoder.h` |
| `GhosttyKeyEncoderOption` | Changed (see section 4 / enum inventory) | No | `key/encoder.h` |
| `GhosttyKeyEvent` | No declaration change | Yes | `key/event.h` |
| `GhosttyKittyGraphics` | No declaration change | Yes | `types.h` |
| `GhosttyKittyGraphicsData` | No declaration change | No | `kitty_graphics.h` |
| `GhosttyKittyGraphicsImage` | No declaration change | Yes | `types.h` |
| `GhosttyKittyGraphicsImageData` | Changed (see section 4 / enum inventory) | Yes | `kitty_graphics.h` |
| `GhosttyKittyGraphicsPlacementData` | No declaration change | No | `kitty_graphics.h` |
| `GhosttyKittyGraphicsPlacementIterator` | No declaration change | Yes | `types.h` |
| `GhosttyKittyGraphicsPlacementIteratorOption` | Changed (see section 4 / enum inventory) | No | `kitty_graphics.h` |
| `GhosttyKittyGraphicsPlacementRenderInfo` | No declaration change | Yes | `kitty_graphics.h` |
| `GhosttyKittyImageCompression` | No declaration change | Yes | `kitty_graphics.h` |
| `GhosttyKittyImageFormat` | No declaration change | Yes | `kitty_graphics.h` |
| `GhosttyKittyKeyFlags` | No declaration change | No | `key/encoder.h` |
| `GhosttyMimeReader` | No declaration change | No | `io.h` |
| `GhosttyMimeReaderFn` | No declaration change | No | `io.h` |
| `GhosttyMode` | No declaration change | Yes | `modes.h` |
| `GhosttyModeReportState` | No declaration change | No | `modes.h` |
| `GhosttyMods` | No declaration change | Yes | `key/event.h` |
| `GhosttyMouseAction` | No declaration change | Yes | `mouse/event.h` |
| `GhosttyMouseButton` | No declaration change | Yes | `mouse/event.h` |
| `GhosttyMouseEncoder` | No declaration change | Yes | `mouse/encoder.h` |
| `GhosttyMouseEncoderOption` | Changed (see section 4 / enum inventory) | No | `mouse/encoder.h` |
| `GhosttyMouseEncoderSize` | No declaration change | Yes | `mouse/encoder.h` |
| `GhosttyMouseEvent` | No declaration change | Yes | `mouse/event.h` |
| `GhosttyMouseFormat` | No declaration change | No | `mouse/encoder.h` |
| `GhosttyMousePosition` | No declaration change | Yes | `mouse/event.h` |
| `GhosttyMouseShape` | No declaration change | No | `mouse.h` |
| `GhosttyMouseTrackingMode` | No declaration change | No | `mouse/encoder.h` |
| `GhosttyOptimizeMode` | No declaration change | Yes | `build_info.h` |
| `GhosttyOptionAsAlt` | No declaration change | No | `key/encoder.h` |
| `GhosttyOscCommand` | No declaration change | No | `types.h` |
| `GhosttyOscCommandData` | No declaration change | No | `osc.h` |
| `GhosttyOscCommandType` | No declaration change | No | `osc.h` |
| `GhosttyOscOption` | Changed (see section 4 / enum inventory) | No | `osc.h` |
| `GhosttyOscParser` | No declaration change | No | `types.h` |
| `GhosttyOscTerminator` | No declaration change | No | `osc.h` |
| `GhosttyPaste` | No declaration change | No | `paste.h` |
| `GhosttyPasteSource` | No declaration change | No | `paste.h` |
| `GhosttyPoint` | No declaration change | Yes | `point.h` |
| `GhosttyPointCoordinate` | No declaration change | No | `point.h` |
| `GhosttyPointTag` | No declaration change | No | `point.h` |
| `GhosttyPointValue` | No declaration change | No | `point.h` |
| `GhosttyProgramStatusKind` | No declaration change | No | `terminal.h` |
| `GhosttyProgramStatusState` | No declaration change | No | `terminal.h` |
| `GhosttyReader` | No declaration change | No | `io.h` |
| `GhosttyReaderFn` | No declaration change | No | `io.h` |
| `GhosttyRenderState` | No declaration change | Yes | `types.h` |
| `GhosttyRenderStateColors` | No declaration change | No | `render.h` |
| `GhosttyRenderStateCursor` | Changed (see section 4 / enum inventory) | No | `render.h` |
| `GhosttyRenderStateData` | No declaration change | No | `render.h` |
| `GhosttyRenderStateDirty` | No declaration change | Yes | `render.h` |
| `GhosttyRenderStateImageData` | No declaration change | No | `render.h` |
| `GhosttyRenderStateImageIterator` | No declaration change | No | `types.h` |
| `GhosttyRenderStateKittyPlacementKind` | No declaration change | No | `render.h` |
| `GhosttyRenderStateOption` | Changed (see section 4 / enum inventory) | No | `render.h` |
| `GhosttyRenderStateOverscan` | No declaration change | No | `render.h` |
| `GhosttyRenderStateRowCells` | No declaration change | Yes | `types.h` |
| `GhosttyRenderStateRowCellsData` | No declaration change | No | `render.h` |
| `GhosttyRenderStateRowData` | No declaration change | No | `render.h` |
| `GhosttyRenderStateRowId` | No declaration change | No | `render.h` |
| `GhosttyRenderStateRowIterator` | No declaration change | Yes | `types.h` |
| `GhosttyRenderStateRowOption` | Changed (see section 4 / enum inventory) | No | `render.h` |
| `GhosttyRenderStateRowSelection` | No declaration change | No | `render.h` |
| `GhosttyResult` | Changed (see section 4 / enum inventory) | Yes | `types.h` |
| `GhosttyRow` | No declaration change | No | `screen.h` |
| `GhosttyRowData` | No declaration change | No | `screen.h` |
| `GhosttyRowSemanticPrompt` | No declaration change | No | `screen.h` |
| `GhosttySearch` | No declaration change | No | `types.h` |
| `GhosttySearchData` | Changed (see section 4 / enum inventory) | No | `search.h` |
| `GhosttySearchOption` | Changed (see section 4 / enum inventory) | No | `search.h` |
| `GhosttySearchScroll` | No declaration change | No | `search.h` |
| `GhosttySearchStatus` | No declaration change | No | `search.h` |
| `GhosttySelection` | Changed (see section 4 / enum inventory) | Yes | `selection.h` |
| `GhosttySelectionAdjust` | No declaration change | No | `selection.h` |
| `GhosttySelectionBuffer` | No declaration change | No | `selection.h` |
| `GhosttySelectionGesture` | No declaration change | No | `selection.h` |
| `GhosttySelectionGestureAutoscroll` | No declaration change | No | `selection.h` |
| `GhosttySelectionGestureBehavior` | No declaration change | No | `selection.h` |
| `GhosttySelectionGestureBehaviors` | No declaration change | No | `selection.h` |
| `GhosttySelectionGestureData` | Changed (see section 4 / enum inventory) | No | `selection.h` |
| `GhosttySelectionGestureEvent` | No declaration change | No | `selection.h` |
| `GhosttySelectionGestureEventOption` | Changed (see section 4 / enum inventory) | No | `selection.h` |
| `GhosttySelectionGestureEventType` | No declaration change | No | `selection.h` |
| `GhosttySelectionGestureGeometry` | No declaration change | No | `selection.h` |
| `GhosttySelectionOrder` | No declaration change | No | `selection.h` |
| `GhosttySemanticPromptKind` | No declaration change | No | `terminal.h` |
| `GhosttySemanticPromptPromptKind` | No declaration change | No | `terminal.h` |
| `GhosttySgrAttribute` | No declaration change | No | `sgr.h` |
| `GhosttySgrAttributeTag` | No declaration change | No | `sgr.h` |
| `GhosttySgrAttributeValue` | No declaration change | No | `sgr.h` |
| `GhosttySgrParser` | No declaration change | No | `types.h` |
| `GhosttySgrUnderline` | No declaration change | No | `sgr.h` |
| `GhosttySgrUnknown` | No declaration change | No | `sgr.h` |
| `GhosttySizeReportSize` | No declaration change | No | `size_report.h` |
| `GhosttySizeReportStyle` | No declaration change | No | `size_report.h` |
| `GhosttySnapshotDecoder` | No declaration change | No | `types.h` |
| `GhosttySnapshotDecoderData` | No declaration change | No | `snapshot.h` |
| `GhosttySnapshotDecoderOption` | Changed (see section 4 / enum inventory) | No | `snapshot.h` |
| `GhosttySnapshotDecoderProgressKind` | No declaration change | No | `snapshot.h` |
| `GhosttyString` | No declaration change | No | `types.h` |
| `GhosttyStyle` | Changed (see section 4 / enum inventory) | Yes | `style.h` |
| `GhosttyStyleColor` | Changed (see section 4 / enum inventory) | No | `style.h` |
| `GhosttyStyleColorTag` | No declaration change | No | `style.h` |
| `GhosttyStyleColorValue` | Changed (see section 4 / enum inventory) | No | `style.h` |
| `GhosttyStyleFlags` | Added | No | `style.h` |
| `GhosttyStyleId` | No declaration change | No | `style.h` |
| `GhosttySurfacePosition` | No declaration change | No | `types.h` |
| `GhosttySysDecodePngFn` | No declaration change | Yes | `sys.h` |
| `GhosttySysImage` | No declaration change | Yes | `sys.h` |
| `GhosttySysLogFn` | No declaration change | No | `sys.h` |
| `GhosttySysLogLevel` | No declaration change | No | `sys.h` |
| `GhosttySysOption` | Changed (see section 4 / enum inventory) | No | `sys.h` |
| `GhosttySysRandomSecureFn` | No declaration change | No | `sys.h` |
| `GhosttyTerminal` | No declaration change | Yes | `types.h` |
| `GhosttyTerminalBellFn` | No declaration change | No | `terminal.h` |
| `GhosttyTerminalBoundary` | No declaration change | No | `terminal.h` |
| `GhosttyTerminalClipboardReadFn` | No declaration change | No | `terminal.h` |
| `GhosttyTerminalClipboardWriteFn` | No declaration change | No | `terminal.h` |
| `GhosttyTerminalColorSchemeFn` | No declaration change | No | `terminal.h` |
| `GhosttyTerminalCompressionMode` | No declaration change | No | `terminal.h` |
| `GhosttyTerminalCompressionResult` | No declaration change | No | `terminal.h` |
| `GhosttyTerminalData` | Changed (see section 4 / enum inventory) | No | `terminal.h` |
| `GhosttyTerminalDesktopNotification` | Changed (see section 4 / enum inventory) | No | `terminal.h` |
| `GhosttyTerminalDesktopNotificationFn` | No declaration change | No | `terminal.h` |
| `GhosttyTerminalDeviceAttributesFn` | No declaration change | No | `terminal.h` |
| `GhosttyTerminalEnquiryFn` | No declaration change | No | `terminal.h` |
| `GhosttyTerminalMemoryUsage` | No declaration change | No | `terminal.h` |
| `GhosttyTerminalModeConfig` | No declaration change | Yes | `terminal.h` |
| `GhosttyTerminalOption` | Changed (see section 4 / enum inventory) | No | `terminal.h` |
| `GhosttyTerminalProgramStatus` | Changed (see section 4 / enum inventory) | No | `terminal.h` |
| `GhosttyTerminalProgramStatusFn` | No declaration change | No | `terminal.h` |
| `GhosttyTerminalProgressReport` | Changed (see section 4 / enum inventory) | No | `terminal.h` |
| `GhosttyTerminalProgressReportFn` | No declaration change | No | `terminal.h` |
| `GhosttyTerminalProgressState` | No declaration change | No | `terminal.h` |
| `GhosttyTerminalPwdChangedFn` | No declaration change | No | `terminal.h` |
| `GhosttyTerminalRenderHoldFn` | No declaration change | No | `terminal.h` |
| `GhosttyTerminalResetFn` | No declaration change | No | `terminal.h` |
| `GhosttyTerminalResizeWindow` | No declaration change | No | `terminal.h` |
| `GhosttyTerminalResizeWindowFn` | No declaration change | No | `terminal.h` |
| `GhosttyTerminalResizeWindowGrid` | No declaration change | No | `terminal.h` |
| `GhosttyTerminalResizeWindowTag` | No declaration change | No | `terminal.h` |
| `GhosttyTerminalResizeWindowValue` | No declaration change | No | `terminal.h` |
| `GhosttyTerminalScreen` | No declaration change | No | `terminal.h` |
| `GhosttyTerminalScrollViewport` | No declaration change | Yes | `terminal.h` |
| `GhosttyTerminalScrollViewportTag` | Changed (see section 4 / enum inventory) | No | `terminal.h` |
| `GhosttyTerminalScrollViewportValue` | No declaration change | No | `terminal.h` |
| `GhosttyTerminalScrollbar` | No declaration change | Yes | `terminal.h` |
| `GhosttyTerminalSelectLineOptions` | No declaration change | No | `selection.h` |
| `GhosttyTerminalSelectWordBetweenOptions` | No declaration change | No | `selection.h` |
| `GhosttyTerminalSelectWordOptions` | No declaration change | Yes | `selection.h` |
| `GhosttyTerminalSemanticPrompt` | Changed (see section 4 / enum inventory) | No | `terminal.h` |
| `GhosttyTerminalSemanticPromptFn` | No declaration change | No | `terminal.h` |
| `GhosttyTerminalSizeFn` | No declaration change | No | `terminal.h` |
| `GhosttyTerminalStringEnd` | No declaration change | No | `terminal.h` |
| `GhosttyTerminalStringIntroducer` | No declaration change | No | `terminal.h` |
| `GhosttyTerminalStringKind` | No declaration change | No | `terminal.h` |
| `GhosttyTerminalTitleChangedFn` | No declaration change | No | `terminal.h` |
| `GhosttyTerminalUnknownOscSequence` | No declaration change | No | `terminal.h` |
| `GhosttyTerminalUnknownSequence` | No declaration change | No | `terminal.h` |
| `GhosttyTerminalUnknownSequenceFn` | No declaration change | No | `terminal.h` |
| `GhosttyTerminalUnknownSequenceTag` | No declaration change | No | `terminal.h` |
| `GhosttyTerminalUnknownSequenceValue` | No declaration change | No | `terminal.h` |
| `GhosttyTerminalUnknownStringSequence` | No declaration change | No | `terminal.h` |
| `GhosttyTerminalWritePtyFn` | No declaration change | No | `terminal.h` |
| `GhosttyTerminalWriteStop` | No declaration change | No | `terminal.h` |
| `GhosttyTerminalXtversionFn` | No declaration change | No | `terminal.h` |
| `GhosttyTrackedGridRef` | No declaration change | Yes | `types.h` |
| `GhosttyWriter` | No declaration change | No | `io.h` |
| `GhosttyWriterFn` | No declaration change | No | `io.h` |

Removed typedefs: `GhosttyFormatterScreenExtra`, `GhosttyFormatterTerminalExtra`, `GhosttyFormatterTerminalOptions`, `GhosttyKittyPlacementLayer`, `GhosttyRenderStateCursorVisualStyle`, `GhosttyRenderStateImageLayer`, `GhosttyTerminalSelectionFormatOptions`, `GhosttyTerminalCursorStyle`.

### E.1 Changed non-enum aggregate declarations

| Type | Current declaration | 1.0 declaration |
|---|---|---|
| `GhosttyAllocatorVtable` | `typedef struct { void* (*alloc)(void *ctx, size_t len, uint8_t alignment, uintptr_t ret_addr); bool (*resize)(void *ctx, void *memory, size_t memory_len, uint8_t alignment, size_t new_len, uintptr_t ret_addr); void* (*remap)(void *ctx, void *memory, size_t memory_len, uint8_t alignment, size_t new_len, uintptr_t ret_addr); void (*free)(void *ctx, void *memory, size_t memory_len, uint8_t alignment, uintptr_t ret_addr); } GhosttyAllocatorVtable;` | `typedef struct { GhosttyAllocatorAllocFn alloc; GhosttyAllocatorResizeFn resize; GhosttyAllocatorRemapFn remap; GhosttyAllocatorFreeFn free; void* reserved[4]; } GhosttyAllocatorVtable;` |
| `GhosttyGridRef` | `typedef struct { size_t size; void *node; uint16_t x; uint16_t y; } GhosttyGridRef;` | `typedef struct { void *node; uint16_t x; uint16_t y; } GhosttyGridRef;` |
| `GhosttyRenderStateCursor` | `typedef struct { size_t size; bool viewport_has_value; uint16_t viewport_x; uint16_t viewport_y; bool wide_tail; bool visible; bool blinking; bool password_input; GhosttyRenderStateCursorVisualStyle visual_style; } GhosttyRenderStateCursor;` | `typedef struct { size_t size; bool viewport_has_value; uint16_t viewport_x; uint16_t viewport_y; bool wide_tail; bool visible; bool blinking; bool password_input; GhosttyCursorStyle visual_style; } GhosttyRenderStateCursor;` |
| `GhosttySelection` | `typedef struct { size_t size; GhosttyGridRef start; GhosttyGridRef end; bool rectangle; } GhosttySelection;` | `typedef struct { GhosttyGridRef start; GhosttyGridRef end; bool rectangle; } GhosttySelection;` |
| `GhosttyStyleColorValue` | `typedef union { GhosttyColorPaletteIndex palette; GhosttyColorRgb rgb; uint64_t _padding; } GhosttyStyleColorValue;` | `typedef union { GhosttyColorPaletteIndex palette; GhosttyColorRgb rgb; } GhosttyStyleColorValue;` |
| `GhosttyStyleColor` | `typedef struct { GhosttyStyleColorTag tag; GhosttyStyleColorValue value; } GhosttyStyleColor;` | `typedef struct { uint8_t tag; GhosttyStyleColorValue value; } GhosttyStyleColor;` |
| `GhosttyStyle` | `typedef struct { size_t size; GhosttyStyleColor fg_color; GhosttyStyleColor bg_color; GhosttyStyleColor underline_color; bool bold; bool italic; bool faint; bool blink; bool inverse; bool invisible; bool strikethrough; bool overline; int underline; } GhosttyStyle;` | `typedef struct { GhosttyStyleColor fg_color; GhosttyStyleColor bg_color; GhosttyStyleColor underline_color; GhosttyStyleFlags flags; uint16_t reserved; } GhosttyStyle;` |
| `GhosttyTerminalDesktopNotification` | `typedef struct { size_t size; GhosttyString title; GhosttyString body; } GhosttyTerminalDesktopNotification;` | `typedef struct { GhosttyString title; GhosttyString body; } GhosttyTerminalDesktopNotification;` |
| `GhosttyTerminalProgressReport` | `typedef struct { size_t size; GhosttyTerminalProgressState state; int8_t progress; } GhosttyTerminalProgressReport;` | `typedef struct { GhosttyTerminalProgressState state; int8_t progress; } GhosttyTerminalProgressReport;` |
| `GhosttyTerminalProgramStatus` | `typedef struct { size_t size; GhosttyProgramStatusState state; GhosttyProgramStatusKind kind; int8_t progress; GhosttyString id; GhosttyString app; GhosttyString title; GhosttyString message; } GhosttyTerminalProgramStatus;` | `typedef struct { GhosttyProgramStatusState state; GhosttyProgramStatusKind kind; int8_t progress; GhosttyString id; GhosttyString app; GhosttyString title; GhosttyString message; } GhosttyTerminalProgramStatus;` |
| `GhosttyTerminalSemanticPrompt` | `typedef struct { size_t size; GhosttySemanticPromptKind kind; GhosttySemanticPromptPromptKind prompt_kind; bool has_exit_code; int32_t exit_code; GhosttyString command; GhosttyString error; } GhosttyTerminalSemanticPrompt;` | `typedef struct { GhosttySemanticPromptKind kind; GhosttySemanticPromptPromptKind prompt_kind; bool has_exit_code; int32_t exit_code; GhosttyString command; GhosttyString error; } GhosttyTerminalSemanticPrompt;` |

### E.2 New non-enum declarations

| Type | 1.0 declaration |
|---|---|
| `GhosttyAllocatorAllocFn` | `typedef void* (*GhosttyAllocatorAllocFn)(void *ctx, size_t len, uint8_t alignment, uintptr_t ret_addr);` |
| `GhosttyAllocatorResizeFn` | `typedef bool (*GhosttyAllocatorResizeFn)(void *ctx, void *memory, size_t memory_len, uint8_t alignment, size_t new_len, uintptr_t ret_addr);` |
| `GhosttyAllocatorRemapFn` | `typedef void* (*GhosttyAllocatorRemapFn)(void *ctx, void *memory, size_t memory_len, uint8_t alignment, size_t new_len, uintptr_t ret_addr);` |
| `GhosttyAllocatorFreeFn` | `typedef void (*GhosttyAllocatorFreeFn)(void *ctx, void *memory, size_t memory_len, uint8_t alignment, uintptr_t ret_addr);` |
| `GhosttyFormatterExtras` | `typedef uint32_t GhosttyFormatterExtras;` |
| `GhosttyStyleFlags` | `typedef uint16_t GhosttyStyleFlags;` |
| `GhosttyCodepointBuffer` | `typedef struct { uint32_t* ptr; size_t cap; size_t len; } GhosttyCodepointBuffer;` |

### E.3 Callback audit

All 27 existing named callback typedefs retain their normalized declaration text. Four named allocator callback typedefs are added. The complete callback-name set is:

`GhosttyAllocatorAllocFn`, `GhosttyAllocatorFreeFn`, `GhosttyAllocatorRemapFn`, `GhosttyAllocatorResizeFn`, `GhosttyClipboardReadReplyFn`, `GhosttyClipboardWriteReplyFn`, `GhosttyMimeReaderFn`, `GhosttyReaderFn`, `GhosttySysDecodePngFn`, `GhosttySysLogFn`, `GhosttySysRandomSecureFn`, `GhosttyTerminalBellFn`, `GhosttyTerminalClipboardReadFn`, `GhosttyTerminalClipboardWriteFn`, `GhosttyTerminalColorSchemeFn`, `GhosttyTerminalDesktopNotificationFn`, `GhosttyTerminalDeviceAttributesFn`, `GhosttyTerminalEnquiryFn`, `GhosttyTerminalProgramStatusFn`, `GhosttyTerminalProgressReportFn`, `GhosttyTerminalPwdChangedFn`, `GhosttyTerminalRenderHoldFn`, `GhosttyTerminalResetFn`, `GhosttyTerminalResizeWindowFn`, `GhosttyTerminalSemanticPromptFn`, `GhosttyTerminalSizeFn`, `GhosttyTerminalTitleChangedFn`, `GhosttyTerminalUnknownSequenceFn`, `GhosttyTerminalWritePtyFn`, `GhosttyTerminalXtversionFn`, `GhosttyWriterFn`.

Before admission, confirm callback pointees, no-throw behavior, calling conventions, synchronous reply obligations, userdata lifetime, reentrancy and unload barriers with real Windows calls. Default allocator vtable reserved slots and borrowed payload strings are part of that proof.


## Appendix F. Public macro changes

Existing macro replacements were matched by expansion, rather than guessed from their English names. Where macro values are unchanged, only the source name changes. New feature/style/formatter macros are not new exported functions.

| Removed macro | Same-expansion target macro (or explicit review) | Current Terminal source mention |
|---|---|---|
| `GHOSTTY_MODE_KAM` | `GHOSTTY_MODE_DISABLE_KEYBOARD` | No |
| `GHOSTTY_MODE_SRM` | `GHOSTTY_MODE_SEND_RECEIVE_MODE` | No |
| `GHOSTTY_MODE_DECCKM` | `GHOSTTY_MODE_CURSOR_KEYS` | Yes |
| `GHOSTTY_MODE_X10_MOUSE` | `GHOSTTY_MODE_MOUSE_EVENT_X10` | No |
| `GHOSTTY_MODE_LEFT_RIGHT_MARGIN` | `GHOSTTY_MODE_ENABLE_LEFT_AND_RIGHT_MARGIN` | No |
| `GHOSTTY_MODE_NORMAL_MOUSE` | `GHOSTTY_MODE_MOUSE_EVENT_NORMAL` | Yes |
| `GHOSTTY_MODE_BUTTON_MOUSE` | `GHOSTTY_MODE_MOUSE_EVENT_BUTTON` | No |
| `GHOSTTY_MODE_ANY_MOUSE` | `GHOSTTY_MODE_MOUSE_EVENT_ANY` | No |
| `GHOSTTY_MODE_UTF8_MOUSE` | `GHOSTTY_MODE_MOUSE_FORMAT_UTF8` | No |
| `GHOSTTY_MODE_SGR_MOUSE` | `GHOSTTY_MODE_MOUSE_FORMAT_SGR` | Yes |
| `GHOSTTY_MODE_ALT_SCROLL` | `GHOSTTY_MODE_MOUSE_ALTERNATE_SCROLL` | No |
| `GHOSTTY_MODE_URXVT_MOUSE` | `GHOSTTY_MODE_MOUSE_FORMAT_URXVT` | No |
| `GHOSTTY_MODE_SGR_PIXELS_MOUSE` | `GHOSTTY_MODE_MOUSE_FORMAT_SGR_PIXELS` | No |
| `GHOSTTY_MODE_NUMLOCK_KEYPAD` | `GHOSTTY_MODE_IGNORE_KEYPAD_WITH_NUMLOCK` | No |
| `GHOSTTY_MODE_ALT_SENDS_ESC` | `GHOSTTY_MODE_ALT_SENDS_ESCAPE` | No |
| `GHOSTTY_MODE_REVERSE_WRAP_EXT` | `GHOSTTY_MODE_REVERSE_WRAP_EXTENDED` | No |
| `GHOSTTY_MODE_ALT_SCREEN_SAVE` | `GHOSTTY_MODE_ALT_SCREEN_SAVE_CURSOR_CLEAR_ENTER` | Yes |
| `GHOSTTY_MODE_SYNC_OUTPUT` | `GHOSTTY_MODE_SYNCHRONIZED_OUTPUT` | No |
| `GHOSTTY_MODE_COLOR_SCHEME_REPORT` | `GHOSTTY_MODE_REPORT_COLOR_SCHEME` | No |
| `GHOSTTY_MODE_VISIBILITY_REPORT` | `GHOSTTY_MODE_REPORT_VISIBILITY` | No |
| `GHOSTTY_MODE_IN_BAND_RESIZE` | `GHOSTTY_MODE_IN_BAND_SIZE_REPORTS` | No |
| `GHOSTTY_MODE_PASTE_EVENTS` | `GHOSTTY_MODE_KITTY_PASTE_EVENTS` | No |

Added macros not already accounted for by the rename table:

- `GHOSTTY_FORMATTER_EXTRA_PALETTE`: `(1u << 0)`
- `GHOSTTY_FORMATTER_EXTRA_MODES`: `(1u << 1)`
- `GHOSTTY_FORMATTER_EXTRA_SCROLLING_REGION`: `(1u << 2)`
- `GHOSTTY_FORMATTER_EXTRA_TABSTOPS`: `(1u << 3)`
- `GHOSTTY_FORMATTER_EXTRA_PWD`: `(1u << 4)`
- `GHOSTTY_FORMATTER_EXTRA_KEYBOARD`: `(1u << 5)`
- `GHOSTTY_FORMATTER_EXTRA_CURSOR`: `(1u << 6)`
- `GHOSTTY_FORMATTER_EXTRA_STYLE`: `(1u << 7)`
- `GHOSTTY_FORMATTER_EXTRA_HYPERLINK`: `(1u << 8)`
- `GHOSTTY_FORMATTER_EXTRA_PROTECTION`: `(1u << 9)`
- `GHOSTTY_FORMATTER_EXTRA_KITTY_KEYBOARD`: `(1u << 10)`
- `GHOSTTY_FORMATTER_EXTRA_CHARSETS`: `(1u << 11)`
- `GHOSTTY_FORMATTER_EXTRAS_NONE`: `0u`
- `GHOSTTY_FORMATTER_EXTRAS_ALL`: `0xFFFu`
- `GHOSTTY_FORMATTER_EXTRAS_DEFAULT`: `(GHOSTTY_FORMATTER_EXTRA_PALETTE \| GHOSTTY_FORMATTER_EXTRA_STYLE \| GHOSTTY_FORMATTER_EXTRA_HYPERLINK)`
- `GHOSTTY_STYLE_FLAG_BOLD`: `((GhosttyStyleFlags)(1u << 0))`
- `GHOSTTY_STYLE_FLAG_ITALIC`: `((GhosttyStyleFlags)(1u << 1))`
- `GHOSTTY_STYLE_FLAG_FAINT`: `((GhosttyStyleFlags)(1u << 2))`
- `GHOSTTY_STYLE_FLAG_BLINK`: `((GhosttyStyleFlags)(1u << 3))`
- `GHOSTTY_STYLE_FLAG_INVERSE`: `((GhosttyStyleFlags)(1u << 4))`
- `GHOSTTY_STYLE_FLAG_INVISIBLE`: `((GhosttyStyleFlags)(1u << 5))`
- `GHOSTTY_STYLE_FLAG_STRIKETHROUGH`: `((GhosttyStyleFlags)(1u << 6))`
- `GHOSTTY_STYLE_FLAG_OVERLINE`: `((GhosttyStyleFlags)(1u << 7))`
- `GHOSTTY_STYLE_UNDERLINE_SHIFT`: `8`
- `GHOSTTY_STYLE_UNDERLINE_MASK`: `((GhosttyStyleFlags)(0x7u << GHOSTTY_STYLE_UNDERLINE_SHIFT))`
- `GHOSTTY_STYLE_UNDERLINE`: `(style) ((int)(((style).flags & GHOSTTY_STYLE_UNDERLINE_MASK) >> GHOSTTY_STYLE_UNDERLINE_SHIFT))`
- `GHOSTTY_DEPRECATED`: `(msg)`
- `GHOSTTY_IMAGE_LAYER_MASK_BELOW_BACKGROUND`: `(1u << GHOSTTY_IMAGE_LAYER_BELOW_BACKGROUND)`
- `GHOSTTY_IMAGE_LAYER_MASK_BELOW_TEXT`: `(1u << GHOSTTY_IMAGE_LAYER_BELOW_TEXT)`
- `GHOSTTY_IMAGE_LAYER_MASK_ABOVE_TEXT`: `(1u << GHOSTTY_IMAGE_LAYER_ABOVE_TEXT)`
- `GHOSTTY_VT_VERSION_H`: `(header guard)`
- `GHOSTTY_VT_VERSION_MAJOR`: `1`
- `GHOSTTY_VT_VERSION_MINOR`: `0`
- `GHOSTTY_VT_VERSION_PATCH`: `0`
- `GHOSTTY_VT_VERSION`: `"1.0.0"`
- `GHOSTTY_VT_VERSION_AT_LEAST`: `(maj, min, pat) (GHOSTTY_VT_VERSION_MAJOR > (maj) \|\| (GHOSTTY_VT_VERSION_MAJOR == (maj) && (GHOSTTY_VT_VERSION_MINOR > (min) \|\| (GHOSTTY_VT_VERSION_MINOR == (min) && GHOSTTY_VT_VERSION_PATCH >= (pat)))))`

`GHOSTTY_ENUM_TYPED` changes its conditional definition branches: the old header also used fixed-enum extensions in older Clang/GCC C modes, while 1.0 keeps fixed `int` enums for C++11+/supported MSVC C++ and C23, and uses the fallback in older C modes. The MSVC C++ branch is unchanged. All other surviving macro names retain their normalized definitions in this comparison. Check all conditional definitions, not just the final `#define` seen for a name; MSVC C/C++ header probes and enum layout checks remain required.

The new `GHOSTTY_DEPRECATED(msg)` has compiler-dependent definitions: MSVC uses
`__declspec(deprecated(msg))`, GCC/Clang use `__attribute__((deprecated(msg)))`,
and only the fallback is empty. The list above shows that fallback; it is not
permission to suppress deprecation diagnostics in this consumer.


## Appendix G. Complete header coverage

Every current header is represented: 35 current headers, 36 target headers, and no removed header. The five Wasm declarations are excluded from the native function baseline, not silently omitted from header review.

| Public header | Current / target declared functions | Target source |
|---|---:|---|
| `include/ghostty/vt.h` | 0 / 0 | [vt.h](https://github.com/ghostty-org/ghostty/blob/0fa3d577a5614150ff8e2c152de5fb9d5c6d2dec/include/ghostty/vt.h) |
| `include/ghostty/vt/allocator.h` | 2 / 2 | [vt/allocator.h](https://github.com/ghostty-org/ghostty/blob/0fa3d577a5614150ff8e2c152de5fb9d5c6d2dec/include/ghostty/vt/allocator.h) |
| `include/ghostty/vt/build_info.h` | 1 / 2 | [vt/build_info.h](https://github.com/ghostty-org/ghostty/blob/0fa3d577a5614150ff8e2c152de5fb9d5c6d2dec/include/ghostty/vt/build_info.h) |
| `include/ghostty/vt/color.h` | 11 / 11 | [vt/color.h](https://github.com/ghostty-org/ghostty/blob/0fa3d577a5614150ff8e2c152de5fb9d5c6d2dec/include/ghostty/vt/color.h) |
| `include/ghostty/vt/color_scheme.h` | 1 / 1 | [vt/color_scheme.h](https://github.com/ghostty-org/ghostty/blob/0fa3d577a5614150ff8e2c152de5fb9d5c6d2dec/include/ghostty/vt/color_scheme.h) |
| `include/ghostty/vt/device.h` | 0 / 0 | [vt/device.h](https://github.com/ghostty-org/ghostty/blob/0fa3d577a5614150ff8e2c152de5fb9d5c6d2dec/include/ghostty/vt/device.h) |
| `include/ghostty/vt/focus.h` | 1 / 1 | [vt/focus.h](https://github.com/ghostty-org/ghostty/blob/0fa3d577a5614150ff8e2c152de5fb9d5c6d2dec/include/ghostty/vt/focus.h) |
| `include/ghostty/vt/formatter.h` | 5 / 7 | [vt/formatter.h](https://github.com/ghostty-org/ghostty/blob/0fa3d577a5614150ff8e2c152de5fb9d5c6d2dec/include/ghostty/vt/formatter.h) |
| `include/ghostty/vt/grid_ref.h` | 5 / 5 | [vt/grid_ref.h](https://github.com/ghostty-org/ghostty/blob/0fa3d577a5614150ff8e2c152de5fb9d5c6d2dec/include/ghostty/vt/grid_ref.h) |
| `include/ghostty/vt/grid_ref_tracked.h` | 5 / 5 | [vt/grid_ref_tracked.h](https://github.com/ghostty-org/ghostty/blob/0fa3d577a5614150ff8e2c152de5fb9d5c6d2dec/include/ghostty/vt/grid_ref_tracked.h) |
| `include/ghostty/vt/io.h` | 0 / 0 | [vt/io.h](https://github.com/ghostty-org/ghostty/blob/0fa3d577a5614150ff8e2c152de5fb9d5c6d2dec/include/ghostty/vt/io.h) |
| `include/ghostty/vt/key.h` | 0 / 0 | [vt/key.h](https://github.com/ghostty-org/ghostty/blob/0fa3d577a5614150ff8e2c152de5fb9d5c6d2dec/include/ghostty/vt/key.h) |
| `include/ghostty/vt/key/encoder.h` | 5 / 5 | [vt/key/encoder.h](https://github.com/ghostty-org/ghostty/blob/0fa3d577a5614150ff8e2c152de5fb9d5c6d2dec/include/ghostty/vt/key/encoder.h) |
| `include/ghostty/vt/key/event.h` | 16 / 16 | [vt/key/event.h](https://github.com/ghostty-org/ghostty/blob/0fa3d577a5614150ff8e2c152de5fb9d5c6d2dec/include/ghostty/vt/key/event.h) |
| `include/ghostty/vt/kitty_graphics.h` | 17 / 17 | [vt/kitty_graphics.h](https://github.com/ghostty-org/ghostty/blob/0fa3d577a5614150ff8e2c152de5fb9d5c6d2dec/include/ghostty/vt/kitty_graphics.h) |
| `include/ghostty/vt/modes.h` | 1 / 1 | [vt/modes.h](https://github.com/ghostty-org/ghostty/blob/0fa3d577a5614150ff8e2c152de5fb9d5c6d2dec/include/ghostty/vt/modes.h) |
| `include/ghostty/vt/mouse.h` | 0 / 0 | [vt/mouse.h](https://github.com/ghostty-org/ghostty/blob/0fa3d577a5614150ff8e2c152de5fb9d5c6d2dec/include/ghostty/vt/mouse.h) |
| `include/ghostty/vt/mouse/encoder.h` | 6 / 6 | [vt/mouse/encoder.h](https://github.com/ghostty-org/ghostty/blob/0fa3d577a5614150ff8e2c152de5fb9d5c6d2dec/include/ghostty/vt/mouse/encoder.h) |
| `include/ghostty/vt/mouse/event.h` | 11 / 11 | [vt/mouse/event.h](https://github.com/ghostty-org/ghostty/blob/0fa3d577a5614150ff8e2c152de5fb9d5c6d2dec/include/ghostty/vt/mouse/event.h) |
| `include/ghostty/vt/osc.h` | 8 / 8 | [vt/osc.h](https://github.com/ghostty-org/ghostty/blob/0fa3d577a5614150ff8e2c152de5fb9d5c6d2dec/include/ghostty/vt/osc.h) |
| `include/ghostty/vt/paste.h` | 3 / 3 | [vt/paste.h](https://github.com/ghostty-org/ghostty/blob/0fa3d577a5614150ff8e2c152de5fb9d5c6d2dec/include/ghostty/vt/paste.h) |
| `include/ghostty/vt/point.h` | 0 / 0 | [vt/point.h](https://github.com/ghostty-org/ghostty/blob/0fa3d577a5614150ff8e2c152de5fb9d5c6d2dec/include/ghostty/vt/point.h) |
| `include/ghostty/vt/render.h` | 27 / 27 | [vt/render.h](https://github.com/ghostty-org/ghostty/blob/0fa3d577a5614150ff8e2c152de5fb9d5c6d2dec/include/ghostty/vt/render.h) |
| `include/ghostty/vt/screen.h` | 4 / 4 | [vt/screen.h](https://github.com/ghostty-org/ghostty/blob/0fa3d577a5614150ff8e2c152de5fb9d5c6d2dec/include/ghostty/vt/screen.h) |
| `include/ghostty/vt/search.h` | 8 / 10 | [vt/search.h](https://github.com/ghostty-org/ghostty/blob/0fa3d577a5614150ff8e2c152de5fb9d5c6d2dec/include/ghostty/vt/search.h) |
| `include/ghostty/vt/selection.h` | 21 / 19 | [vt/selection.h](https://github.com/ghostty-org/ghostty/blob/0fa3d577a5614150ff8e2c152de5fb9d5c6d2dec/include/ghostty/vt/selection.h) |
| `include/ghostty/vt/sgr.h` | 9 / 9 | [vt/sgr.h](https://github.com/ghostty-org/ghostty/blob/0fa3d577a5614150ff8e2c152de5fb9d5c6d2dec/include/ghostty/vt/sgr.h) |
| `include/ghostty/vt/size_report.h` | 1 / 1 | [vt/size_report.h](https://github.com/ghostty-org/ghostty/blob/0fa3d577a5614150ff8e2c152de5fb9d5c6d2dec/include/ghostty/vt/size_report.h) |
| `include/ghostty/vt/snapshot.h` | 12 / 12 | [vt/snapshot.h](https://github.com/ghostty-org/ghostty/blob/0fa3d577a5614150ff8e2c152de5fb9d5c6d2dec/include/ghostty/vt/snapshot.h) |
| `include/ghostty/vt/style.h` | 2 / 2 | [vt/style.h](https://github.com/ghostty-org/ghostty/blob/0fa3d577a5614150ff8e2c152de5fb9d5c6d2dec/include/ghostty/vt/style.h) |
| `include/ghostty/vt/sys.h` | 2 / 2 | [vt/sys.h](https://github.com/ghostty-org/ghostty/blob/0fa3d577a5614150ff8e2c152de5fb9d5c6d2dec/include/ghostty/vt/sys.h) |
| `include/ghostty/vt/terminal.h` | 18 / 18 | [vt/terminal.h](https://github.com/ghostty-org/ghostty/blob/0fa3d577a5614150ff8e2c152de5fb9d5c6d2dec/include/ghostty/vt/terminal.h) |
| `include/ghostty/vt/types.h` | 1 / 2 | [vt/types.h](https://github.com/ghostty-org/ghostty/blob/0fa3d577a5614150ff8e2c152de5fb9d5c6d2dec/include/ghostty/vt/types.h) |
| `include/ghostty/vt/unicode.h` | 2 / 2 | [vt/unicode.h](https://github.com/ghostty-org/ghostty/blob/0fa3d577a5614150ff8e2c152de5fb9d5c6d2dec/include/ghostty/vt/unicode.h) |
| `include/ghostty/vt/version.h` | 0 / 0 | [vt/version.h](https://github.com/ghostty-org/ghostty/blob/0fa3d577a5614150ff8e2c152de5fb9d5c6d2dec/include/ghostty/vt/version.h) |
| `include/ghostty/vt/wasm.h` | 5 / 5 | [vt/wasm.h](https://github.com/ghostty-org/ghostty/blob/0fa3d577a5614150ff8e2c152de5fb9d5c6d2dec/include/ghostty/vt/wasm.h) |

## Appendix H. Option and data selector adoption

Every target `*Option`, `*Data`, and build-info enum is covered. The source-mention column includes production code and test-enabled diagnostics in `Plugins/Terminal`; it is an inventory lead, not proof that a callback is registered or that all behavior behind a selector is supported. Section 7 distinguishes current host behavior from extension proposals. A missing mention does not mean the engine does not use a feature internally. The all-values sentinel is excluded because it is not a usable selector.

Storage annotations reproduce the candidate type contract, not a standalone calling example: enum-keyed setters/getters still take an address of the documented value (or the documented function-pointer form), and iterator-population getters require an initialized owned handle. Null/reset semantics, borrowed lifetimes, per-key return codes, and alignment must be checked in the owning header and Zig binding before adoption.

### GhosttyBuildInfo

| Selector | 1.0 value | Candidate documented storage type | Current Terminal source mention |
|---|---:|---|---|
| `GHOSTTY_BUILD_INFO_INVALID` | `0` | `Invalid sentinel; no value storage` | Not referenced |
| `GHOSTTY_BUILD_INFO_SIMD` | `1` | `Output type: bool *` | Referenced (`TerminalRuntimeLoader.cpp`) |
| `GHOSTTY_BUILD_INFO_KITTY_GRAPHICS` | `2` | `Output type: bool *` | Referenced (`TerminalRuntimeLoader.cpp`) |
| `GHOSTTY_BUILD_INFO_TMUX_CONTROL_MODE` | `3` | `Output type: bool *` | Not referenced |
| `GHOSTTY_BUILD_INFO_OPTIMIZE` | `4` | `Output type: GhosttyOptimizeMode *` | Referenced (`TerminalRuntimeLoader.cpp`) |
| `GHOSTTY_BUILD_INFO_VERSION_STRING` | `5` | `Output type: GhosttyString *` | Not referenced |
| `GHOSTTY_BUILD_INFO_VERSION_MAJOR` | `6` | `Output type: size_t *` | Not referenced |
| `GHOSTTY_BUILD_INFO_VERSION_MINOR` | `7` | `Output type: size_t *` | Not referenced |
| `GHOSTTY_BUILD_INFO_VERSION_PATCH` | `8` | `Output type: size_t *` | Not referenced |
| `GHOSTTY_BUILD_INFO_VERSION_PRE` | `9` | `Output type: GhosttyString *` | Not referenced |
| `GHOSTTY_BUILD_INFO_VERSION_BUILD` | `10` | `Output type: GhosttyString *` | Not referenced |
| `GHOSTTY_BUILD_INFO_FEATURE_SNAPSHOT` | `11` | `Output type: bool *` | Not referenced |
| `GHOSTTY_BUILD_INFO_FEATURE_FORMATTER` | `12` | `Output type: bool *` | Not referenced |
| `GHOSTTY_BUILD_INFO_FEATURE_SELECTION` | `13` | `Output type: bool *` | Not referenced |
| `GHOSTTY_BUILD_INFO_FEATURE_SEARCH` | `14` | `Output type: bool *` | Not referenced |
| `GHOSTTY_BUILD_INFO_FEATURE_RENDER_STATE` | `15` | `Output type: bool *` | Not referenced |
| `GHOSTTY_BUILD_INFO_FEATURE_INPUT_ENCODE` | `16` | `Output type: bool *` | Not referenced |
| `GHOSTTY_BUILD_INFO_FEATURE_COLOR` | `17` | `Output type: bool *` | Not referenced |
| `GHOSTTY_BUILD_INFO_FEATURE_GRID_INTROSPECTION` | `18` | `Output type: bool *` | Not referenced |
| `GHOSTTY_BUILD_INFO_FEATURE_GLYPH_PROTOCOL` | `19` | `Output type: bool *` | Not referenced |

### GhosttyFormatterOption

| Selector | 1.0 value | Candidate documented storage type | Current Terminal source mention |
|---|---:|---|---|
| `GHOSTTY_FORMATTER_OPT_INVALID` | `0` | `Invalid sentinel; no value storage` | Not referenced |
| `GHOSTTY_FORMATTER_OPT_EMIT` | `1` | `Value type: GhosttyFormatterFormat` | Not referenced |
| `GHOSTTY_FORMATTER_OPT_UNWRAP` | `2` | `Value type: bool` | Not referenced |
| `GHOSTTY_FORMATTER_OPT_TRIM` | `3` | `Value type: bool` | Not referenced |
| `GHOSTTY_FORMATTER_OPT_SELECTION` | `4` | `Value type: GhosttySelection` | Not referenced |
| `GHOSTTY_FORMATTER_OPT_EXTRAS` | `5` | `Value type: GhosttyFormatterExtras` | Not referenced |

### GhosttyKeyEncoderOption

| Selector | 1.0 value | Candidate documented storage type | Current Terminal source mention |
|---|---:|---|---|
| `GHOSTTY_KEY_ENCODER_OPT_INVALID` | `0` | `Invalid sentinel; no value storage` | Not referenced |
| `GHOSTTY_KEY_ENCODER_OPT_CURSOR_KEY_APPLICATION` | `1` | `Value type: bool` | Not referenced |
| `GHOSTTY_KEY_ENCODER_OPT_KEYPAD_KEY_APPLICATION` | `2` | `Value type: bool` | Not referenced |
| `GHOSTTY_KEY_ENCODER_OPT_IGNORE_KEYPAD_WITH_NUMLOCK` | `3` | `Value type: bool` | Not referenced |
| `GHOSTTY_KEY_ENCODER_OPT_ALT_ESC_PREFIX` | `4` | `Value type: bool` | Not referenced |
| `GHOSTTY_KEY_ENCODER_OPT_MODIFY_OTHER_KEYS_STATE_2` | `5` | `Value type: bool` | Not referenced |
| `GHOSTTY_KEY_ENCODER_OPT_KITTY_FLAGS` | `6` | `Value type: GhosttyKittyKeyFlags bitmask` | Not referenced |
| `GHOSTTY_KEY_ENCODER_OPT_MACOS_OPTION_AS_ALT` | `7` | `Value type: GhosttyOptionAsAlt` | Not referenced |
| `GHOSTTY_KEY_ENCODER_OPT_BACKARROW_KEY_MODE` | `8` | `Value type: bool` | Not referenced |

### GhosttyKittyGraphicsData

| Selector | 1.0 value | Candidate documented storage type | Current Terminal source mention |
|---|---:|---|---|
| `GHOSTTY_KITTY_GRAPHICS_DATA_INVALID` | `0` | `Invalid sentinel; no value storage` | Not referenced |
| `GHOSTTY_KITTY_GRAPHICS_DATA_PLACEMENT_ITERATOR` | `1` | `Output type: GhosttyKittyGraphicsPlacementIterator *` | Referenced (`Terminal.cpp`, `Terminal.Tests.VtUpgrade.cpp`, `TerminalVt.cpp`) |
| `GHOSTTY_KITTY_GRAPHICS_DATA_GENERATION` | `2` | `Output type: uint64_t *` | Referenced (`TerminalVt.cpp`) |

### GhosttyKittyGraphicsPlacementData

| Selector | 1.0 value | Candidate documented storage type | Current Terminal source mention |
|---|---:|---|---|
| `GHOSTTY_KITTY_GRAPHICS_PLACEMENT_DATA_INVALID` | `0` | `Invalid sentinel; no value storage` | Not referenced |
| `GHOSTTY_KITTY_GRAPHICS_PLACEMENT_DATA_IMAGE_ID` | `1` | `Output type: uint32_t *` | Referenced (`Terminal.cpp`, `Terminal.Tests.VtUpgrade.cpp`, `TerminalVt.cpp`) |
| `GHOSTTY_KITTY_GRAPHICS_PLACEMENT_DATA_PLACEMENT_ID` | `2` | `Output type: uint32_t *` | Referenced (`Terminal.Tests.VtUpgrade.cpp`) |
| `GHOSTTY_KITTY_GRAPHICS_PLACEMENT_DATA_IS_VIRTUAL` | `3` | `Output type: bool *` | Referenced (`Terminal.Tests.VtUpgrade.cpp`, `TerminalVt.cpp`) |
| `GHOSTTY_KITTY_GRAPHICS_PLACEMENT_DATA_X_OFFSET` | `4` | `Output type: uint32_t *` | Referenced (`Terminal.Tests.VtUpgrade.cpp`, `TerminalVt.cpp`) |
| `GHOSTTY_KITTY_GRAPHICS_PLACEMENT_DATA_Y_OFFSET` | `5` | `Output type: uint32_t *` | Referenced (`Terminal.Tests.VtUpgrade.cpp`, `TerminalVt.cpp`) |
| `GHOSTTY_KITTY_GRAPHICS_PLACEMENT_DATA_SOURCE_X` | `6` | `Output type: uint32_t *` | Referenced (`Terminal.Tests.VtUpgrade.cpp`) |
| `GHOSTTY_KITTY_GRAPHICS_PLACEMENT_DATA_SOURCE_Y` | `7` | `Output type: uint32_t *` | Referenced (`Terminal.Tests.VtUpgrade.cpp`) |
| `GHOSTTY_KITTY_GRAPHICS_PLACEMENT_DATA_SOURCE_WIDTH` | `8` | `Output type: uint32_t *` | Referenced (`Terminal.Tests.VtUpgrade.cpp`) |
| `GHOSTTY_KITTY_GRAPHICS_PLACEMENT_DATA_SOURCE_HEIGHT` | `9` | `Output type: uint32_t *` | Referenced (`Terminal.Tests.VtUpgrade.cpp`) |
| `GHOSTTY_KITTY_GRAPHICS_PLACEMENT_DATA_COLUMNS` | `10` | `Output type: uint32_t *` | Referenced (`Terminal.Tests.VtUpgrade.cpp`) |
| `GHOSTTY_KITTY_GRAPHICS_PLACEMENT_DATA_ROWS` | `11` | `Output type: uint32_t *` | Referenced (`Terminal.Tests.VtUpgrade.cpp`) |
| `GHOSTTY_KITTY_GRAPHICS_PLACEMENT_DATA_Z` | `12` | `Output type: int32_t *` | Referenced (`Terminal.Tests.VtUpgrade.cpp`, `TerminalVt.cpp`) |

### GhosttyKittyGraphicsPlacementIteratorOption

| Selector | 1.0 value | Candidate documented storage type | Current Terminal source mention |
|---|---:|---|---|
| `GHOSTTY_KITTY_GRAPHICS_PLACEMENT_ITERATOR_OPT_INVALID` | `0` | `Invalid sentinel; no value storage` | Not referenced |
| `GHOSTTY_KITTY_GRAPHICS_PLACEMENT_ITERATOR_OPT_LAYER` | `1` | `Input type: const uint32_t *` | Not referenced |

### GhosttyKittyGraphicsImageData

| Selector | 1.0 value | Candidate documented storage type | Current Terminal source mention |
|---|---:|---|---|
| `GHOSTTY_KITTY_GRAPHICS_IMAGE_DATA_INVALID` | `0` | `Invalid sentinel; no value storage` | Not referenced |
| `GHOSTTY_KITTY_GRAPHICS_IMAGE_DATA_ID` | `1` | `Output type: uint32_t *` | Not referenced |
| `GHOSTTY_KITTY_GRAPHICS_IMAGE_DATA_NUMBER` | `2` | `Output type: uint32_t *` | Not referenced |
| `GHOSTTY_KITTY_GRAPHICS_IMAGE_DATA_WIDTH` | `3` | `Output type: uint32_t *` | Not referenced |
| `GHOSTTY_KITTY_GRAPHICS_IMAGE_DATA_HEIGHT` | `4` | `Output type: uint32_t *` | Not referenced |
| `GHOSTTY_KITTY_GRAPHICS_IMAGE_DATA_FORMAT` | `5` | `Output type: GhosttyKittyImageFormat *` | Not referenced |
| `GHOSTTY_KITTY_GRAPHICS_IMAGE_DATA_COMPRESSION` | `6` | `Output type: GhosttyKittyImageCompression *` | Not referenced |
| `GHOSTTY_KITTY_GRAPHICS_IMAGE_DATA_PIXELS` | `7` | `Output type: GhosttyString *` | Not referenced |
| `GHOSTTY_KITTY_GRAPHICS_IMAGE_DATA_EXPECTED_LEN` | `8` | `Output type: size_t *` | Not referenced |
| `GHOSTTY_KITTY_GRAPHICS_IMAGE_DATA_GENERATION` | `9` | `Output type: uint64_t *` | Not referenced |

### GhosttyMouseEncoderOption

| Selector | 1.0 value | Candidate documented storage type | Current Terminal source mention |
|---|---:|---|---|
| `GHOSTTY_MOUSE_ENCODER_OPT_INVALID` | `0` | `Invalid sentinel; no value storage` | Not referenced |
| `GHOSTTY_MOUSE_ENCODER_OPT_EVENT` | `1` | `Value type: GhosttyMouseTrackingMode` | Not referenced |
| `GHOSTTY_MOUSE_ENCODER_OPT_FORMAT` | `2` | `Value type: GhosttyMouseFormat` | Not referenced |
| `GHOSTTY_MOUSE_ENCODER_OPT_SIZE` | `3` | `Value type: GhosttyMouseEncoderSize` | Referenced (`Terminal.cpp`, `TerminalVt.cpp`) |
| `GHOSTTY_MOUSE_ENCODER_OPT_ANY_BUTTON_PRESSED` | `4` | `Value type: bool` | Referenced (`Terminal.cpp`, `TerminalVt.cpp`) |
| `GHOSTTY_MOUSE_ENCODER_OPT_TRACK_LAST_CELL` | `5` | `Value type: bool` | Not referenced |

### GhosttyOscOption

| Selector | 1.0 value | Candidate documented storage type | Current Terminal source mention |
|---|---:|---|---|
| `GHOSTTY_OSC_OPT_INVALID` | `0` | `Invalid sentinel; no value storage` | Not referenced |
| `GHOSTTY_OSC_OPT_UNKNOWN_MAX_BYTES` | `1` | `Input type: size_t*` | Not referenced |

### GhosttyOscCommandData

| Selector | 1.0 value | Candidate documented storage type | Current Terminal source mention |
|---|---:|---|---|
| `GHOSTTY_OSC_DATA_INVALID` | `0` | `Invalid sentinel; no value storage` | Not referenced |
| `GHOSTTY_OSC_DATA_CHANGE_WINDOW_TITLE_STR` | `1` | `Output type: GhosttyString *` | Not referenced |
| `GHOSTTY_OSC_DATA_UNKNOWN_CONTENT` | `2` | `Output type: GhosttyString * Lifetime: Valid until the next call to any ghostty_osc_* function with the same parser instance` | Not referenced |
| `GHOSTTY_OSC_DATA_UNKNOWN_TRUNCATED` | `3` | `Output type: bool *` | Not referenced |
| `GHOSTTY_OSC_DATA_UNKNOWN_TERMINATOR` | `4` | `Output type: GhosttyOscTerminator *` | Not referenced |

### GhosttyRenderStateData

| Selector | 1.0 value | Candidate documented storage type | Current Terminal source mention |
|---|---:|---|---|
| `GHOSTTY_RENDER_STATE_DATA_INVALID` | `0` | `Invalid sentinel; no value storage` | Not referenced |
| `GHOSTTY_RENDER_STATE_DATA_COLS` | `1` | `Storage type: uint16_t` | Referenced (`Terminal.Tests.VtUpgrade.cpp`, `TerminalVt.cpp`) |
| `GHOSTTY_RENDER_STATE_DATA_ROWS` | `2` | `Storage type: uint16_t` | Referenced (`Terminal.Tests.VtUpgrade.cpp`, `TerminalVt.cpp`) |
| `GHOSTTY_RENDER_STATE_DATA_DIRTY` | `3` | `Storage type: GhosttyRenderStateDirty` | Not referenced |
| `GHOSTTY_RENDER_STATE_DATA_ROW_ITERATOR` | `4` | `Storage type: GhosttyRenderStateRowIterator` | Referenced (`Terminal.cpp`, `Terminal.Tests.VtUpgrade.cpp`, `TerminalVt.cpp`) |
| `GHOSTTY_RENDER_STATE_DATA_COLOR_BACKGROUND` | `5` | `Storage type: GhosttyColorRgb` | Not referenced |
| `GHOSTTY_RENDER_STATE_DATA_COLOR_FOREGROUND` | `6` | `Storage type: GhosttyColorRgb` | Not referenced |
| `GHOSTTY_RENDER_STATE_DATA_COLOR_CURSOR` | `7` | `Storage type: GhosttyColorRgb` | Not referenced |
| `GHOSTTY_RENDER_STATE_DATA_COLOR_CURSOR_HAS_VALUE` | `8` | `Storage type: bool` | Not referenced |
| `GHOSTTY_RENDER_STATE_DATA_COLOR_PALETTE` | `9` | `Storage type: GhosttyColorRgb[256]` | Not referenced |
| `GHOSTTY_RENDER_STATE_DATA_CURSOR_VISUAL_STYLE` | `10` | `Storage type: GhosttyCursorStyle` | Referenced (`Terminal.Tests.VtUpgrade.cpp`, `TerminalVt.cpp`) |
| `GHOSTTY_RENDER_STATE_DATA_CURSOR_VISIBLE` | `11` | `Storage type: bool` | Referenced (`Terminal.Tests.VtUpgrade.cpp`, `TerminalVt.cpp`) |
| `GHOSTTY_RENDER_STATE_DATA_CURSOR_BLINKING` | `12` | `Storage type: bool` | Not referenced |
| `GHOSTTY_RENDER_STATE_DATA_CURSOR_PASSWORD_INPUT` | `13` | `Storage type: bool` | Not referenced |
| `GHOSTTY_RENDER_STATE_DATA_CURSOR_VIEWPORT_HAS_VALUE` | `14` | `Storage type: bool` | Referenced (`Terminal.Tests.VtUpgrade.cpp`, `TerminalVt.cpp`) |
| `GHOSTTY_RENDER_STATE_DATA_CURSOR_VIEWPORT_X` | `15` | `Storage type: uint16_t` | Referenced (`Terminal.Tests.VtUpgrade.cpp`, `TerminalVt.cpp`) |
| `GHOSTTY_RENDER_STATE_DATA_CURSOR_VIEWPORT_Y` | `16` | `Storage type: uint16_t` | Referenced (`Terminal.Tests.VtUpgrade.cpp`, `TerminalVt.cpp`) |
| `GHOSTTY_RENDER_STATE_DATA_CURSOR_VIEWPORT_WIDE_TAIL` | `17` | `Storage type: bool` | Not referenced |
| `GHOSTTY_RENDER_STATE_DATA_CURSOR` | `18` | `Storage type: GhosttyRenderStateCursor` | Not referenced |
| `GHOSTTY_RENDER_STATE_DATA_COLORS` | `19` | `Storage type: GhosttyRenderStateColors` | Not referenced |
| `GHOSTTY_RENDER_STATE_DATA_OVERSCAN` | `20` | `Storage type: GhosttyRenderStateOverscan` | Not referenced |
| `GHOSTTY_RENDER_STATE_DATA_OVERSCAN_REQUEST` | `21` | `Storage type: GhosttyRenderStateOverscan` | Not referenced |
| `GHOSTTY_RENDER_STATE_DATA_IMAGE_ITERATOR` | `22` | `Input/output: GhosttyRenderStateImageIterator * (existing owned handle required)` | Not referenced |
| `GHOSTTY_RENDER_STATE_DATA_IMAGE_PLACEMENTS_GENERATION` | `23` | `Storage type: uint64_t` | Not referenced |

### GhosttyRenderStateOption

| Selector | 1.0 value | Candidate documented storage type | Current Terminal source mention |
|---|---:|---|---|
| `GHOSTTY_RENDER_STATE_OPT_INVALID` | `0` | `Invalid sentinel; no value storage` | Not referenced |
| `GHOSTTY_RENDER_STATE_OPT_DIRTY` | `1` | `Storage type: GhosttyRenderStateDirty` | Not referenced |
| `GHOSTTY_RENDER_STATE_OPT_OVERSCAN` | `2` | `Storage type: GhosttyRenderStateOverscan` | Not referenced |

### GhosttyRenderStateImageData

| Selector | 1.0 value | Candidate documented storage type | Current Terminal source mention |
|---|---:|---|---|
| `GHOSTTY_RENDER_STATE_IMAGE_DATA_INVALID` | `0` | `Invalid sentinel; no value storage` | Not referenced |
| `GHOSTTY_RENDER_STATE_IMAGE_DATA_PROTOCOL` | `1` | `Storage type: GhosttyImageProtocol` | Not referenced |
| `GHOSTTY_RENDER_STATE_IMAGE_DATA_IMAGE_ID` | `2` | `Storage type: uint32_t` | Not referenced |
| `GHOSTTY_RENDER_STATE_IMAGE_DATA_IMAGE_GENERATION` | `3` | `Storage type: uint64_t` | Not referenced |
| `GHOSTTY_RENDER_STATE_IMAGE_DATA_LAYER` | `4` | `Storage type: GhosttyImageLayer` | Not referenced |
| `GHOSTTY_RENDER_STATE_IMAGE_DATA_Z` | `5` | `Storage type: int32_t` | Not referenced |
| `GHOSTTY_RENDER_STATE_IMAGE_DATA_VIEWPORT_COL` | `6` | `Storage type: int32_t` | Not referenced |
| `GHOSTTY_RENDER_STATE_IMAGE_DATA_VIEWPORT_ROW` | `7` | `Storage type: int32_t` | Not referenced |
| `GHOSTTY_RENDER_STATE_IMAGE_DATA_GRID_COLS` | `8` | `Storage type: uint32_t` | Not referenced |
| `GHOSTTY_RENDER_STATE_IMAGE_DATA_GRID_ROWS` | `9` | `Storage type: uint32_t` | Not referenced |
| `GHOSTTY_RENDER_STATE_IMAGE_DATA_CELL_OFFSET_X` | `10` | `Storage type: uint32_t` | Not referenced |
| `GHOSTTY_RENDER_STATE_IMAGE_DATA_CELL_OFFSET_Y` | `11` | `Storage type: uint32_t` | Not referenced |
| `GHOSTTY_RENDER_STATE_IMAGE_DATA_PIXEL_WIDTH` | `12` | `Storage type: uint32_t` | Not referenced |
| `GHOSTTY_RENDER_STATE_IMAGE_DATA_PIXEL_HEIGHT` | `13` | `Storage type: uint32_t` | Not referenced |
| `GHOSTTY_RENDER_STATE_IMAGE_DATA_SOURCE_X` | `14` | `Storage type: uint32_t` | Not referenced |
| `GHOSTTY_RENDER_STATE_IMAGE_DATA_SOURCE_Y` | `15` | `Storage type: uint32_t` | Not referenced |
| `GHOSTTY_RENDER_STATE_IMAGE_DATA_SOURCE_WIDTH` | `16` | `Storage type: uint32_t` | Not referenced |
| `GHOSTTY_RENDER_STATE_IMAGE_DATA_SOURCE_HEIGHT` | `17` | `Storage type: uint32_t` | Not referenced |
| `GHOSTTY_RENDER_STATE_IMAGE_DATA_KITTY_PLACEMENT_ID` | `18` | `Storage type: uint32_t` | Not referenced |
| `GHOSTTY_RENDER_STATE_IMAGE_DATA_KITTY_KIND` | `19` | `Storage type: GhosttyRenderStateKittyPlacementKind` | Not referenced |

### GhosttyRenderStateRowData

| Selector | 1.0 value | Candidate documented storage type | Current Terminal source mention |
|---|---:|---|---|
| `GHOSTTY_RENDER_STATE_ROW_DATA_INVALID` | `0` | `Invalid sentinel; no value storage` | Not referenced |
| `GHOSTTY_RENDER_STATE_ROW_DATA_DIRTY` | `1` | `Storage type: bool` | Not referenced |
| `GHOSTTY_RENDER_STATE_ROW_DATA_RAW` | `2` | `Storage type: GhosttyRow` | Not referenced |
| `GHOSTTY_RENDER_STATE_ROW_DATA_CELLS` | `3` | `Storage type: GhosttyRenderStateRowCells` | Referenced (`Terminal.cpp`, `Terminal.Tests.VtUpgrade.cpp`, `TerminalVt.cpp`) |
| `GHOSTTY_RENDER_STATE_ROW_DATA_SELECTION` | `4` | `Storage type: GhosttyRenderStateRowSelection` | Not referenced |
| `GHOSTTY_RENDER_STATE_ROW_DATA_CELLS_RAW` | `5` | `Storage type: GhosttyCellsView` | Not referenced |
| `GHOSTTY_RENDER_STATE_ROW_DATA_VIEWPORT_Y` | `6` | `Storage type: int32_t` | Not referenced |
| `GHOSTTY_RENDER_STATE_ROW_DATA_ID` | `7` | `Storage type: GhosttyRenderStateRowId` | Not referenced |

### GhosttyRenderStateRowOption

| Selector | 1.0 value | Candidate documented storage type | Current Terminal source mention |
|---|---:|---|---|
| `GHOSTTY_RENDER_STATE_ROW_OPT_INVALID` | `0` | `Invalid sentinel; no value storage` | Not referenced |
| `GHOSTTY_RENDER_STATE_ROW_OPT_DIRTY` | `1` | `Storage type: bool` | Not referenced |

### GhosttyRenderStateRowCellsData

| Selector | 1.0 value | Candidate documented storage type | Current Terminal source mention |
|---|---:|---|---|
| `GHOSTTY_RENDER_STATE_ROW_CELLS_DATA_INVALID` | `0` | `Invalid sentinel; no value storage` | Not referenced |
| `GHOSTTY_RENDER_STATE_ROW_CELLS_DATA_RAW` | `1` | `Storage type: GhosttyCell` | Referenced (`Terminal.Tests.VtUpgrade.cpp`, `TerminalVt.cpp`) |
| `GHOSTTY_RENDER_STATE_ROW_CELLS_DATA_STYLE` | `2` | `Storage type: GhosttyStyle` | Referenced (`Terminal.cpp`, `Terminal.Tests.VtUpgrade.cpp`, `TerminalVt.cpp`) |
| `GHOSTTY_RENDER_STATE_ROW_CELLS_DATA_GRAPHEMES_LEN` | `3` | `Storage type: uint32_t` | Not referenced |
| `GHOSTTY_RENDER_STATE_ROW_CELLS_DATA_GRAPHEMES_BUF` | `4` | `Storage type: GhosttyCodepointBuffer` | Not referenced |
| `GHOSTTY_RENDER_STATE_ROW_CELLS_DATA_BG_COLOR` | `5` | `Storage type: GhosttyColorRgb` | Referenced (`Terminal.cpp`, `Terminal.Tests.VtUpgrade.cpp`, `TerminalVt.cpp`) |
| `GHOSTTY_RENDER_STATE_ROW_CELLS_DATA_FG_COLOR` | `6` | `Storage type: GhosttyColorRgb` | Referenced (`Terminal.cpp`, `Terminal.Tests.VtUpgrade.cpp`, `TerminalVt.cpp`) |
| `GHOSTTY_RENDER_STATE_ROW_CELLS_DATA_SELECTED` | `7` | `Storage type: bool` | Referenced (`Terminal.Tests.VtUpgrade.cpp`, `TerminalVt.cpp`) |
| `GHOSTTY_RENDER_STATE_ROW_CELLS_DATA_HAS_STYLING` | `8` | `Storage type: bool` | Referenced (`Terminal.Tests.VtUpgrade.cpp`, `TerminalVt.cpp`) |
| `GHOSTTY_RENDER_STATE_ROW_CELLS_DATA_GRAPHEMES_UTF8` | `9` | `Storage type: GhosttyBuffer` | Referenced (`Terminal.cpp`, `Terminal.Tests.VtUpgrade.cpp`, `TerminalVt.cpp`) |

### GhosttyCellData

| Selector | 1.0 value | Candidate documented storage type | Current Terminal source mention |
|---|---:|---|---|
| `GHOSTTY_CELL_DATA_INVALID` | `0` | `Invalid sentinel; no value storage` | Not referenced |
| `GHOSTTY_CELL_DATA_CODEPOINT` | `1` | `Output type: uint32_t *` | Not referenced |
| `GHOSTTY_CELL_DATA_CONTENT_TAG` | `2` | `Output type: GhosttyCellContentTag *` | Not referenced |
| `GHOSTTY_CELL_DATA_WIDE` | `3` | `Output type: GhosttyCellWide *` | Referenced (`Terminal.Tests.VtUpgrade.cpp`, `TerminalVt.cpp`) |
| `GHOSTTY_CELL_DATA_HAS_TEXT` | `4` | `Output type: bool *` | Not referenced |
| `GHOSTTY_CELL_DATA_HAS_STYLING` | `5` | `Output type: bool *` | Not referenced |
| `GHOSTTY_CELL_DATA_STYLE_ID` | `6` | `Output type: GhosttyStyleId * (uint16_t *)` | Not referenced |
| `GHOSTTY_CELL_DATA_HAS_HYPERLINK` | `7` | `Output type: bool *` | Referenced (`Terminal.Tests.VtUpgrade.cpp`, `TerminalVt.cpp`) |
| `GHOSTTY_CELL_DATA_PROTECTED` | `8` | `Output type: bool *` | Not referenced |
| `GHOSTTY_CELL_DATA_SEMANTIC_CONTENT` | `9` | `Output type: GhosttyCellSemanticContent *` | Not referenced |
| `GHOSTTY_CELL_DATA_COLOR_PALETTE` | `10` | `Output type: GhosttyColorPaletteIndex *` | Not referenced |
| `GHOSTTY_CELL_DATA_COLOR_RGB` | `11` | `Output type: GhosttyColorRgb *` | Not referenced |

### GhosttyRowData

| Selector | 1.0 value | Candidate documented storage type | Current Terminal source mention |
|---|---:|---|---|
| `GHOSTTY_ROW_DATA_INVALID` | `0` | `Invalid sentinel; no value storage` | Not referenced |
| `GHOSTTY_ROW_DATA_WRAP` | `1` | `Output type: bool *` | Not referenced |
| `GHOSTTY_ROW_DATA_WRAP_CONTINUATION` | `2` | `Output type: bool *` | Not referenced |
| `GHOSTTY_ROW_DATA_GRAPHEME` | `3` | `Output type: bool *` | Not referenced |
| `GHOSTTY_ROW_DATA_STYLED` | `4` | `Output type: bool *` | Not referenced |
| `GHOSTTY_ROW_DATA_HYPERLINK` | `5` | `Output type: bool *` | Not referenced |
| `GHOSTTY_ROW_DATA_SEMANTIC_PROMPT` | `6` | `Output type: GhosttyRowSemanticPrompt *` | Not referenced |
| `GHOSTTY_ROW_DATA_KITTY_VIRTUAL_PLACEHOLDER` | `7` | `Output type: bool *` | Not referenced |
| `GHOSTTY_ROW_DATA_DIRTY` | `8` | `Output type: bool *` | Not referenced |

### GhosttySearchData

| Selector | 1.0 value | Candidate documented storage type | Current Terminal source mention |
|---|---:|---|---|
| `GHOSTTY_SEARCH_DATA_INVALID` | `0` | `Invalid sentinel; no value storage` | Not referenced |
| `GHOSTTY_SEARCH_DATA_STATUS` | `1` | `Storage type: GhosttySearchStatus*` | Not referenced |
| `GHOSTTY_SEARCH_DATA_NEEDLE` | `2` | `Storage type: GhosttyString*` | Not referenced |
| `GHOSTTY_SEARCH_DATA_TOTAL_MATCHES` | `3` | `Storage type: size_t*` | Not referenced |
| `GHOSTTY_SEARCH_DATA_SELECTED_INDEX` | `4` | `Storage type: size_t*` | Not referenced |
| `GHOSTTY_SEARCH_DATA_SELECTED_MATCH` | `5` | `Storage type: GhosttySelection*` | Not referenced |
| `GHOSTTY_SEARCH_DATA_MATCHES` | `6` | `Storage type: GhosttySelectionBuffer*` | Not referenced |
| `GHOSTTY_SEARCH_DATA_VIEWPORT_MATCHES` | `7` | `Storage type: GhosttySelectionBuffer*` | Not referenced |
| `GHOSTTY_SEARCH_DATA_SELECT_SCROLL` | `8` | `Storage type: GhosttySearchScroll*` | Not referenced |

### GhosttySearchOption

| Selector | 1.0 value | Candidate documented storage type | Current Terminal source mention |
|---|---:|---|---|
| `GHOSTTY_SEARCH_OPT_INVALID` | `0` | `Invalid sentinel; no value storage` | Not referenced |
| `GHOSTTY_SEARCH_OPT_NEEDLE` | `1` | `Storage type: const GhosttyString*` | Not referenced |
| `GHOSTTY_SEARCH_OPT_SELECT_SCROLL` | `4` | `Storage type: const GhosttySearchScroll*` | Not referenced |

### GhosttySelectionGestureData

| Selector | 1.0 value | Candidate documented storage type | Current Terminal source mention |
|---|---:|---|---|
| `GHOSTTY_SELECTION_GESTURE_DATA_INVALID` | `0` | `Invalid sentinel; no value storage` | Not referenced |
| `GHOSTTY_SELECTION_GESTURE_DATA_CLICK_COUNT` | `1` | `Storage type: uint8_t*` | Not referenced |
| `GHOSTTY_SELECTION_GESTURE_DATA_DRAGGED` | `2` | `Storage type: bool*` | Not referenced |
| `GHOSTTY_SELECTION_GESTURE_DATA_AUTOSCROLL` | `3` | `Storage type: GhosttySelectionGestureAutoscroll*` | Not referenced |
| `GHOSTTY_SELECTION_GESTURE_DATA_BEHAVIOR` | `4` | `Storage type: GhosttySelectionGestureBehavior*` | Not referenced |
| `GHOSTTY_SELECTION_GESTURE_DATA_ANCHOR` | `5` | `Storage type: GhosttyGridRef*` | Not referenced |

### GhosttySelectionGestureEventOption

| Selector | 1.0 value | Candidate documented storage type | Current Terminal source mention |
|---|---:|---|---|
| `GHOSTTY_SELECTION_GESTURE_EVENT_OPT_INVALID` | `0` | `Invalid sentinel; no value storage` | Not referenced |
| `GHOSTTY_SELECTION_GESTURE_EVENT_OPT_REF` | `1` | `Storage type: GhosttyGridRef*` | Not referenced |
| `GHOSTTY_SELECTION_GESTURE_EVENT_OPT_POSITION` | `2` | `Storage type: GhosttySurfacePosition*` | Not referenced |
| `GHOSTTY_SELECTION_GESTURE_EVENT_OPT_REPEAT_DISTANCE` | `3` | `Storage type: double*` | Not referenced |
| `GHOSTTY_SELECTION_GESTURE_EVENT_OPT_TIME_NS` | `4` | `Storage type: uint64_t*` | Not referenced |
| `GHOSTTY_SELECTION_GESTURE_EVENT_OPT_REPEAT_INTERVAL_NS` | `5` | `Storage type: uint64_t*` | Not referenced |
| `GHOSTTY_SELECTION_GESTURE_EVENT_OPT_WORD_BOUNDARY_CODEPOINTS` | `6` | `Storage type: GhosttyCodepoints*` | Not referenced |
| `GHOSTTY_SELECTION_GESTURE_EVENT_OPT_BEHAVIORS` | `7` | `Storage type: GhosttySelectionGestureBehaviors*` | Not referenced |
| `GHOSTTY_SELECTION_GESTURE_EVENT_OPT_RECTANGLE` | `8` | `Storage type: bool*` | Not referenced |
| `GHOSTTY_SELECTION_GESTURE_EVENT_OPT_GEOMETRY` | `9` | `Storage type: GhosttySelectionGestureGeometry*` | Not referenced |
| `GHOSTTY_SELECTION_GESTURE_EVENT_OPT_VIEWPORT` | `10` | `Storage type: GhosttyPointCoordinate*` | Not referenced |

### GhosttySnapshotDecoderOption

| Selector | 1.0 value | Candidate documented storage type | Current Terminal source mention |
|---|---:|---|---|
| `GHOSTTY_SNAPSHOT_DECODER_OPT_INVALID` | `0` | `Invalid sentinel; no value storage` | Not referenced |
| `GHOSTTY_SNAPSHOT_DECODER_OPT_MAX_CONTINUATION_BYTES` | `1` | `Input type: size_t *` | Not referenced |
| `GHOSTTY_SNAPSHOT_DECODER_OPT_RETAIN_CONTINUATION` | `2` | `Input type: bool *` | Not referenced |
| `GHOSTTY_SNAPSHOT_DECODER_OPT_COMPRESS_HISTORY` | `3` | `Input type: bool *` | Not referenced |
| `GHOSTTY_SNAPSHOT_DECODER_OPT_KITTY_IMAGE_STORAGE_LIMIT` | `4` | `Input type: uint64_t *` | Not referenced |

### GhosttySnapshotDecoderData

| Selector | 1.0 value | Candidate documented storage type | Current Terminal source mention |
|---|---:|---|---|
| `GHOSTTY_SNAPSHOT_DECODER_DATA_INVALID` | `0` | `Invalid sentinel; no value storage` | Not referenced |
| `GHOSTTY_SNAPSHOT_DECODER_DATA_MAX_CONTINUATION_BYTES` | `1` | `Output type: size_t *` | Not referenced |
| `GHOSTTY_SNAPSHOT_DECODER_DATA_SOURCE_OFFSET` | `2` | `Output type: size_t *` | Not referenced |
| `GHOSTTY_SNAPSHOT_DECODER_DATA_HISTORY_ROWS_PRIMARY` | `3` | `Output type: uint64_t *` | Not referenced |
| `GHOSTTY_SNAPSHOT_DECODER_DATA_HISTORY_ROWS_ALTERNATE` | `4` | `Output type: uint64_t *` | Not referenced |
| `GHOSTTY_SNAPSHOT_DECODER_DATA_PROGRESS_SCREEN` | `5` | `Output type: GhosttyTerminalScreen *` | Not referenced |
| `GHOSTTY_SNAPSHOT_DECODER_DATA_PROGRESS_ROWS` | `6` | `Output type: size_t *` | Not referenced |
| `GHOSTTY_SNAPSHOT_DECODER_DATA_PROGRESS_REMAINING` | `7` | `Output type: uint32_t *` | Not referenced |
| `GHOSTTY_SNAPSHOT_DECODER_DATA_RETAIN_CONTINUATION` | `8` | `Output type: bool *` | Not referenced |
| `GHOSTTY_SNAPSHOT_DECODER_DATA_COMPRESS_HISTORY` | `9` | `Output type: bool *` | Not referenced |
| `GHOSTTY_SNAPSHOT_DECODER_DATA_KITTY_IMAGE_STORAGE_LIMIT` | `10` | `Output type: uint64_t *` | Not referenced |
| `GHOSTTY_SNAPSHOT_DECODER_DATA_PROGRESS_KIND` | `11` | `Output type: GhosttySnapshotDecoderProgressKind *` | Not referenced |
| `GHOSTTY_SNAPSHOT_DECODER_DATA_PROGRESS_IMAGE_ID` | `12` | `Output type: uint32_t *` | Not referenced |
| `GHOSTTY_SNAPSHOT_DECODER_DATA_PROGRESS_IMAGE_FRAME` | `13` | `Output type: uint32_t *` | Not referenced |
| `GHOSTTY_SNAPSHOT_DECODER_DATA_PROGRESS_IMAGE_BYTES` | `14` | `Output type: size_t *` | Not referenced |
| `GHOSTTY_SNAPSHOT_DECODER_DATA_KITTY_IMAGES_PRIMARY` | `15` | `Output type: uint32_t *` | Not referenced |
| `GHOSTTY_SNAPSHOT_DECODER_DATA_KITTY_IMAGES_ALTERNATE` | `16` | `Output type: uint32_t *` | Not referenced |
| `GHOSTTY_SNAPSHOT_DECODER_DATA_KITTY_PENDING_BYTES_PRIMARY` | `17` | `Output type: uint64_t *` | Not referenced |
| `GHOSTTY_SNAPSHOT_DECODER_DATA_KITTY_PENDING_BYTES_ALTERNATE` | `18` | `Output type: uint64_t *` | Not referenced |

### GhosttySysOption

| Selector | 1.0 value | Candidate documented storage type | Current Terminal source mention |
|---|---:|---|---|
| `GHOSTTY_SYS_OPT_INVALID` | `0` | `Invalid sentinel; no value storage` | Not referenced |
| `GHOSTTY_SYS_OPT_USERDATA` | `1` | `Input type: void* (or NULL)` | Not referenced |
| `GHOSTTY_SYS_OPT_DECODE_PNG` | `2` | `Input type: GhosttySysDecodePngFn (function pointer, or NULL)` | Referenced (`TerminalRuntimeLoader.cpp`) |
| `GHOSTTY_SYS_OPT_LOG` | `3` | `Input type: GhosttySysLogFn (function pointer, or NULL)` | Not referenced |
| `GHOSTTY_SYS_OPT_RANDOM_SECURE` | `4` | `Input type: GhosttySysRandomSecureFn (function pointer, or NULL)` | Not referenced |

### GhosttyTerminalOption

| Selector | 1.0 value | Candidate documented storage type | Current Terminal source mention |
|---|---:|---|---|
| `GHOSTTY_TERMINAL_OPT_INVALID` | `0` | `Invalid sentinel; no value storage` | Not referenced |
| `GHOSTTY_TERMINAL_OPT_USERDATA` | `1` | `Input type: void*` | Referenced (`Terminal.cpp`) |
| `GHOSTTY_TERMINAL_OPT_WRITE_PTY` | `2` | `Input type: GhosttyTerminalWritePtyFn` | Referenced (`Terminal.cpp`) |
| `GHOSTTY_TERMINAL_OPT_BELL` | `3` | `Input type: GhosttyTerminalBellFn` | Not referenced |
| `GHOSTTY_TERMINAL_OPT_ENQUIRY` | `4` | `Input type: GhosttyTerminalEnquiryFn` | Not referenced |
| `GHOSTTY_TERMINAL_OPT_XTVERSION` | `5` | `Input type: GhosttyTerminalXtversionFn` | Not referenced |
| `GHOSTTY_TERMINAL_OPT_TITLE_CHANGED` | `6` | `Input type: GhosttyTerminalTitleChangedFn` | Not referenced |
| `GHOSTTY_TERMINAL_OPT_SIZE` | `7` | `Input type: GhosttyTerminalSizeFn` | Not referenced |
| `GHOSTTY_TERMINAL_OPT_COLOR_SCHEME` | `8` | `Input type: GhosttyTerminalColorSchemeFn` | Not referenced |
| `GHOSTTY_TERMINAL_OPT_DEVICE_ATTRIBUTES` | `9` | `Input type: GhosttyTerminalDeviceAttributesFn` | Not referenced |
| `GHOSTTY_TERMINAL_OPT_TITLE` | `10` | `Input type: GhosttyString*` | Not referenced |
| `GHOSTTY_TERMINAL_OPT_PWD` | `11` | `Input type: GhosttyString*` | Not referenced |
| `GHOSTTY_TERMINAL_OPT_COLOR_FOREGROUND` | `12` | `Input type: GhosttyColorRgb*` | Not referenced |
| `GHOSTTY_TERMINAL_OPT_COLOR_BACKGROUND` | `13` | `Input type: GhosttyColorRgb*` | Not referenced |
| `GHOSTTY_TERMINAL_OPT_COLOR_CURSOR` | `14` | `Input type: GhosttyColorRgb*` | Not referenced |
| `GHOSTTY_TERMINAL_OPT_COLOR_PALETTE` | `15` | `Input type: GhosttyColorRgb[256]*` | Not referenced |
| `GHOSTTY_TERMINAL_OPT_KITTY_IMAGE_STORAGE_LIMIT` | `16` | `Input type: uint64_t*` | Referenced (`Terminal.cpp`) |
| `GHOSTTY_TERMINAL_OPT_KITTY_IMAGE_MEDIUM_FILE` | `17` | `Input type: bool*` | Referenced (`Terminal.cpp`) |
| `GHOSTTY_TERMINAL_OPT_KITTY_IMAGE_MEDIUM_TEMP_FILE` | `18` | `Input type: GhosttyString*` | Referenced (`Terminal.cpp`) |
| `GHOSTTY_TERMINAL_OPT_KITTY_IMAGE_MEDIUM_SHARED_MEM` | `19` | `Input type: bool*` | Referenced (`Terminal.cpp`) |
| `GHOSTTY_TERMINAL_OPT_APC_MAX_BYTES` | `20` | `Input type: size_t*` | Not referenced |
| `GHOSTTY_TERMINAL_OPT_APC_MAX_BYTES_KITTY` | `21` | `Input type: size_t*` | Referenced (`Terminal.cpp`) |
| `GHOSTTY_TERMINAL_OPT_SELECTION` | `22` | `Input type: GhosttySelection*` | Referenced (`Terminal.cpp`, `TerminalVt.cpp`) |
| `GHOSTTY_TERMINAL_OPT_DEFAULT_CURSOR_STYLE` | `23` | `Input type: GhosttyCursorStyle*` | Not referenced |
| `GHOSTTY_TERMINAL_OPT_DEFAULT_CURSOR_BLINK` | `24` | `Input type: bool*` | Not referenced |
| `GHOSTTY_TERMINAL_OPT_GLYPH_PROTOCOL` | `25` | `Input type: bool*` | Not referenced |
| `GHOSTTY_TERMINAL_OPT_PWD_CHANGED` | `26` | `Input type: GhosttyTerminalPwdChangedFn` | Not referenced |
| `GHOSTTY_TERMINAL_OPT_CLIPBOARD_WRITE` | `27` | `Input type: GhosttyTerminalClipboardWriteFn` | Referenced (`Terminal.cpp`) |
| `GHOSTTY_TERMINAL_OPT_SCROLLBACK_MAX_BYTES` | `28` | `Input type: size_t*` | Not referenced |
| `GHOSTTY_TERMINAL_OPT_SCROLLBACK_MAX_LINES` | `29` | `Input type: size_t*` | Not referenced |
| `GHOSTTY_TERMINAL_OPT_DESKTOP_NOTIFICATION` | `30` | `Input type: GhosttyTerminalDesktopNotificationFn` | Not referenced |
| `GHOSTTY_TERMINAL_OPT_PROGRESS_REPORT` | `31` | `Input type: GhosttyTerminalProgressReportFn` | Not referenced |
| `GHOSTTY_TERMINAL_OPT_CONTINUATION_MAX_BYTES` | `32` | `Input type: size_t*` | Not referenced |
| `GHOSTTY_TERMINAL_OPT_TITLE_REPORT` | `33` | `Input type: bool*` | Not referenced |
| `GHOSTTY_TERMINAL_OPT_MODE_DEFAULT` | `34` | `Input type: GhosttyTerminalModeConfig*` | Not referenced |
| `GHOSTTY_TERMINAL_OPT_MODE` | `35` | `Input type: GhosttyTerminalModeConfig*` | Not referenced |
| `GHOSTTY_TERMINAL_OPT_UNKNOWN_SEQUENCE` | `36` | `Input type: GhosttyTerminalUnknownSequenceFn` | Not referenced |
| `GHOSTTY_TERMINAL_OPT_UNKNOWN_MAX_BYTES` | `37` | `Input type: size_t*` | Not referenced |
| `GHOSTTY_TERMINAL_OPT_TERMINFO_NAME` | `38` | `Input type: GhosttyString*` | Not referenced |
| `GHOSTTY_TERMINAL_OPT_CLIPBOARD_READ` | `39` | `Input type: GhosttyTerminalClipboardReadFn` | Referenced (`Terminal.cpp`) |
| `GHOSTTY_TERMINAL_OPT_CLIPBOARD_WRITE_MAX_BYTES` | `40` | `Input type: size_t*` | Referenced (`Terminal.cpp`) |
| `GHOSTTY_TERMINAL_OPT_RESIZE_PULL_SCROLLBACK` | `41` | `Input type: bool*` | Not referenced |
| `GHOSTTY_TERMINAL_OPT_RENDER_HOLD` | `42` | `Input type: GhosttyTerminalRenderHoldFn` | Not referenced |
| `GHOSTTY_TERMINAL_OPT_SEMANTIC_PROMPT` | `43` | `Input type: GhosttyTerminalSemanticPromptFn` | Not referenced |
| `GHOSTTY_TERMINAL_OPT_RESET` | `44` | `Input type: GhosttyTerminalResetFn` | Not referenced |
| `GHOSTTY_TERMINAL_OPT_XT_CHECKSUM_REPORT` | `45` | `Input type: bool*` | Not referenced |
| `GHOSTTY_TERMINAL_OPT_XT_CHECKSUM_EXTENSION` | `46` | `Input type: uint8_t*` | Not referenced |
| `GHOSTTY_TERMINAL_OPT_PROGRAM_STATUS` | `47` | `Input type: GhosttyTerminalProgramStatusFn` | Not referenced |
| `GHOSTTY_TERMINAL_OPT_RESIZE_WINDOW` | `48` | `Input type: GhosttyTerminalResizeWindowFn` | Not referenced |

### GhosttyTerminalData

| Selector | 1.0 value | Candidate documented storage type | Current Terminal source mention |
|---|---:|---|---|
| `GHOSTTY_TERMINAL_DATA_INVALID` | `0` | `Invalid sentinel; no value storage` | Not referenced |
| `GHOSTTY_TERMINAL_DATA_COLS` | `1` | `Output type: uint16_t *` | Not referenced |
| `GHOSTTY_TERMINAL_DATA_ROWS` | `2` | `Output type: uint16_t *` | Not referenced |
| `GHOSTTY_TERMINAL_DATA_CURSOR_X` | `3` | `Output type: uint16_t *` | Not referenced |
| `GHOSTTY_TERMINAL_DATA_CURSOR_Y` | `4` | `Output type: uint16_t *` | Not referenced |
| `GHOSTTY_TERMINAL_DATA_CURSOR_PENDING_WRAP` | `5` | `Output type: bool *` | Not referenced |
| `GHOSTTY_TERMINAL_DATA_ACTIVE_SCREEN` | `6` | `Output type: GhosttyTerminalScreen *` | Not referenced |
| `GHOSTTY_TERMINAL_DATA_CURSOR_VISIBLE` | `7` | `Output type: bool *` | Not referenced |
| `GHOSTTY_TERMINAL_DATA_KITTY_KEYBOARD_FLAGS` | `8` | `Output type: GhosttyKittyKeyFlags * (uint8_t *)` | Not referenced |
| `GHOSTTY_TERMINAL_DATA_SCROLLBAR` | `9` | `Output type: GhosttyTerminalScrollbar *` | Referenced (`Terminal.cpp`, `TerminalVt.cpp`) |
| `GHOSTTY_TERMINAL_DATA_CURSOR_SGR_STYLE` | `10` | `Output type: GhosttyStyle *` | Not referenced |
| `GHOSTTY_TERMINAL_DATA_MOUSE_TRACKING` | `11` | `Output type: bool *` | Not referenced |
| `GHOSTTY_TERMINAL_DATA_TITLE` | `12` | `Output type: GhosttyString *` | Not referenced |
| `GHOSTTY_TERMINAL_DATA_PWD` | `13` | `Output type: GhosttyString *` | Not referenced |
| `GHOSTTY_TERMINAL_DATA_TOTAL_ROWS` | `14` | `Output type: size_t *` | Not referenced |
| `GHOSTTY_TERMINAL_DATA_SCROLLBACK_ROWS` | `15` | `Output type: size_t *` | Not referenced |
| `GHOSTTY_TERMINAL_DATA_WIDTH_PX` | `16` | `Output type: uint32_t *` | Not referenced |
| `GHOSTTY_TERMINAL_DATA_HEIGHT_PX` | `17` | `Output type: uint32_t *` | Not referenced |
| `GHOSTTY_TERMINAL_DATA_COLOR_FOREGROUND` | `18` | `Output type: GhosttyColorRgb *` | Not referenced |
| `GHOSTTY_TERMINAL_DATA_COLOR_BACKGROUND` | `19` | `Output type: GhosttyColorRgb *` | Not referenced |
| `GHOSTTY_TERMINAL_DATA_COLOR_CURSOR` | `20` | `Output type: GhosttyColorRgb *` | Not referenced |
| `GHOSTTY_TERMINAL_DATA_COLOR_PALETTE` | `21` | `Output type: GhosttyColorRgb[256] *` | Not referenced |
| `GHOSTTY_TERMINAL_DATA_COLOR_FOREGROUND_DEFAULT` | `22` | `Output type: GhosttyColorRgb *` | Not referenced |
| `GHOSTTY_TERMINAL_DATA_COLOR_BACKGROUND_DEFAULT` | `23` | `Output type: GhosttyColorRgb *` | Not referenced |
| `GHOSTTY_TERMINAL_DATA_COLOR_CURSOR_DEFAULT` | `24` | `Output type: GhosttyColorRgb *` | Not referenced |
| `GHOSTTY_TERMINAL_DATA_COLOR_PALETTE_DEFAULT` | `25` | `Output type: GhosttyColorRgb[256] *` | Not referenced |
| `GHOSTTY_TERMINAL_DATA_KITTY_IMAGE_STORAGE_LIMIT` | `26` | `Output type: uint64_t *` | Not referenced |
| `GHOSTTY_TERMINAL_DATA_KITTY_IMAGE_MEDIUM_FILE` | `27` | `Output type: bool *` | Not referenced |
| `GHOSTTY_TERMINAL_DATA_KITTY_IMAGE_MEDIUM_TEMP_FILE` | `28` | `Output type: GhosttyString *` | Not referenced |
| `GHOSTTY_TERMINAL_DATA_KITTY_IMAGE_MEDIUM_SHARED_MEM` | `29` | `Output type: bool *` | Not referenced |
| `GHOSTTY_TERMINAL_DATA_KITTY_GRAPHICS` | `30` | `Output type: GhosttyKittyGraphics *` | Referenced (`Terminal.cpp`, `Terminal.Tests.VtUpgrade.cpp`, `TerminalVt.cpp`) |
| `GHOSTTY_TERMINAL_DATA_SELECTION` | `31` | `Output type: GhosttySelection *` | Not referenced |
| `GHOSTTY_TERMINAL_DATA_VIEWPORT_ACTIVE` | `32` | `Output type: bool *` | Not referenced |
| `GHOSTTY_TERMINAL_DATA_VT_PROCESSING_ERROR` | `33` | `Output type: bool *` | Not referenced |
| `GHOSTTY_TERMINAL_DATA_SCROLLBACK_MAX_BYTES` | `34` | `Output type: size_t *` | Not referenced |
| `GHOSTTY_TERMINAL_DATA_SCROLLBACK_MAX_LINES` | `35` | `Output type: size_t *` | Not referenced |
| `GHOSTTY_TERMINAL_DATA_CONTINUATION_MAX_BYTES` | `36` | `Output type: size_t *` | Not referenced |
| `GHOSTTY_TERMINAL_DATA_MODE` | `37` | `Input/output type: GhosttyTerminalModeConfig *` | Referenced (`TerminalRuntimeLoader.cpp`) |
| `GHOSTTY_TERMINAL_DATA_VT_GROUND` | `38` | `Output type: bool *` | Not referenced |
| `GHOSTTY_TERMINAL_DATA_CURSOR_AT_PROMPT` | `39` | `Output type: bool *` | Not referenced |
| `GHOSTTY_TERMINAL_DATA_CLIPBOARD_WRITE_MAX_BYTES` | `40` | `Output type: size_t *` | Not referenced |
| `GHOSTTY_TERMINAL_DATA_MOUSE_SHAPE` | `41` | `Output type: GhosttyMouseShape *` | Not referenced |
| `GHOSTTY_TERMINAL_DATA_MEMORY_USAGE` | `42` | `Output type: GhosttyTerminalMemoryUsage *` | Not referenced |

Private placement selectors and the clipboard gate do not occur in these unmodified upstream headers. Their independent current/new allocations and host uses are reviewed in section 5; never interpret their absence here as permission to remove their controls.

## Authoring checks — 2026-11-07

- `Tools/Get-SpecInventory.ps1 -FailOnFindings`: completed successfully during authoring with **0 blocking findings**. It validates local links, heading anchors, WIP ownership, and current specification inventory; it does not validate native ABI behavior.
- `Tools/Test-Changelog.ps1`: passed; nine valid pending fragments at this checkpoint. No generated changelog edit.
- `git diff --check`: passed for the tracked index change. The new document was separately checked for final newline/trailing whitespace; no violations.
- Inventory completeness checks: 209 unique native function rows, all 71 incumbent loader requirements, 216 target typedef rows plus eight removed typedefs, 314 option/data selectors with documented storage annotations, all 36 target headers, and no omitted changed enum member. The 27 pre-existing named callback declarations and four added allocator callbacks are accounted for. This is source/metadata checking, not native layout or call validation.
- Version-matched API/stability/immutable packaging links returned HTTP 200. Source archive checksum matches the published checksum; local signature verification remains pending.
- `Test-Changes.ps1 -Explain`: emitted a planning-only decision for exactly the three documentation/index/fragment paths. The command returned exit **1**, selected an advisory full build with `SOURCE_INPUT_CHANGED` entries, and listed four interactive groups as pending. No native build or test was executed and no passing test verdict is inferred. The decision was written to the already marked sandbox, `Z:\RedSalamander.Perf\evidence\explain-plan.json`; it is an authoring checkpoint, not qualification of future implementation bytes.

Runtime, dependency, overlay, build scripts, workflows, settings and package outputs are unchanged by this proposal. No implementation acceptance checkbox is complete.
