# August 8 Correctness, Data Safety, Tooling, and Simplification Remediation

> **ARCHIVED 2026-08-09 — SUPERSEDED AS AN EXECUTION QUEUE.** Every confirmed row,
> dependency, proof requirement, and contrary-evidence correction is transferred to
> `../WIP/Operation_Cinderstar_UnifiedRepositoryRiskRemediation_2026-08-09.md` or to the
> sole Atlas/CI owner named there. This file is historical audit evidence only.

> **NON-NORMATIVE HISTORICAL REVIEW.** This file preserves the confirmed findings from the
> review of `0f473e67bd547bc33a72cc9c3d7f39c99e8ca81a..68d2b3bc7b6f8f4e09163735124bd9d35c45890a`.
> Current behavior remains owned by the authoritative domain specifications, public ABI headers,
> workflows, and machine-consumed schemas. Live execution and closeout now occur only through
> Operation Cinderstar and its explicitly routed owners.

## Status and ownership

- **State:** ARCHIVED — live remainder transferred to Operation Cinderstar
- **Created:** 2026-08-08
- **Priority:** P0 data safety, followed by P1 correctness/build/release blockers and P2 hardening/simplification
- **Planned at:** `68d2b3bc7b6f8f4e09163735124bd9d35c45890a`
- **Review base:** `0f473e67bd547bc33a72cc9c3d7f39c99e8ca81a`
- **Review head:** `68d2b3bc7b6f8f4e09163735124bd9d35c45890a`
- **Primary owner:** this plan for every executable `A8-*` row except `A8-CI-01`, which is
  routed to `CI_FormatAutocommitRace_2026-08-25.md`. `A8-ABI-01` is evidence-only until
  `Operation_Atlas_RemainingSpecificationDecisions_2026-08-04.md` resolves the supported
  compatibility boundary; neither routed row creates a second implementation queue.
- **Contrary-evidence boundary:** `A8-S3-01` through `A8-S3-03` narrow the earlier EG-S3-02
  closeout. Exact source-revision capture and conditional reads remain valid, but the old
  rollback, version-specific MOVE-delete, and fake-only direct-copy conclusions are not
  completion evidence for these new rows.
- **Drift check:**

```powershell
git diff 68d2b3bc7b6f8f4e09163735124bd9d35c45890a..HEAD -- `
  Common/PlugInterfaces Common/DxUi Plugins/Terminal Plugins/FileSystemS3 `
  Plugins/FileSystemCurl RedSalamander/HostServices.cpp `
  RedSalamander/FolderWindow.cpp RedSalamander/FolderWindow.Layout.cpp `
  RedSalamander/FolderWindow.FileSystem.Navigation.cpp RedSalamander/RedSalamander.cpp `
  Tools Directory.Build.targets .gitattributes .github/workflows Specs
```

## Review conclusion and captured validation

The reviewed range is not ready to merge or release. The review confirmed one critical
cross-client S3 data-loss route, multiple source/destination integrity failures, host ABI and
marshaling defects, terminal input/location defects, a clean-build blocker, checkout- and
locale-dependent tooling, and an inconsistent release workflow.

No production or test source was changed during the review. The following evidence was captured:

- `build.ps1 -Rebuild -Configuration Debug -Platform x64` failed in
  `Tools/Build-GhosttyTerminalRuntime.ps1` because the downloaded uucode archive SHA-256 was
  `7e76fc7fab1e7ac728c52b35bbb3e5b8c639841abfc7fe1a4bcb13050594bc9e`, not the pinned value.
- `DxUiTests`, `FileSystemCurlTests`, `CrashHandlingTests`, `LocalizationTests`,
  `SettingsSchemaTests`, `ViewerPETests`, `ViewerSqliteTests`, and `RedConfigureTests` passed.
- `PluginContractTests` reached the expected failure caused by absent `Terminal.dll` after the
  blocked build; the reported non-Terminal checks passed.
- Terminal tooling Pester results were 245 passed, 10 failed, and 5 skipped. The failures were
  attributable to raw-byte hashes under CRLF checkout conversion and localized tar output.
- The source worktree was clean after review. The abandoned-artifact contamination marker remains
  until a successful full Debug/x64 rebuild clears it through the repository build protocol.

## Finding ledger

| ID | Priority | Provenance | Confirmed defect | Required implementation outcome | Required focused proof |
|---|---:|---|---|---|---|
| A8-S3-01 | P0 | Pre-existing, adjacent to exact-revision work | `TransferJournal` retains destination key strings but not the revision it published. Rollback reads the current destination to restore a deleted source and then deletes the current destination. A concurrent writer can therefore have its bytes copied into the source and its destination deleted. | Publication returns and journals an immutable destination identity. Rollback reads/deletes only that identity. If identity cannot be proven, preserve all potentially authoritative data and return a partial result. Versioned rollback removes only the operation's exact published revision/delete marker; it never deletes an unrelated current version. | Deterministically replace `dest/a` after publication and before rollback; assert the external replacement survives and the original source bytes, never replacement bytes, are restored. Cover COPY cancellation, MOVE rollback, overwrite backups, and versioned/unversioned buckets. |
| A8-S3-02 | P1 | Introduced by `e14ba96e3` | Versioned MOVE passes the captured `VersionId` to `DeleteObject`. Deleting current `v2` by version makes older `v1` current, so MOVE returns success while the source path remains visible with stale data. Existing fake tests encode this incorrect policy. | Pin `VersionId`/ETag for reads and copies, but perform a key-level conditional source removal that hides the source with a delete marker only when the captured revision is still current. Preserve a newer replacement. Journal the created delete marker so rollback can remove only that marker. | Seed old `v1` plus current `v2`; after MOVE, ordinary HEAD/List must not expose the source. Inject a newer replacement before delete and prove it is retained. Roll back a versioned MOVE and prove only the operation's marker is removed. Update fake and live-compatible endpoint tests. |
| A8-S3-03 | P2 | Base bug pre-existing; version query widened by `e14ba96e3` | `UrlEncodeS3CopySource` percent-encodes the value before `SetCopySource`; AWS SDK 1.11.800 encodes it again. Server-side copy receives `%25` escapes and falls back to local relay. | Pass the logical unescaped SDK input in the SDK-required source form. Remove the stale declaration and keep exactly one copy-source construction helper. | Inspect `GetRequestSpecificHeaders()` for ordinary, reserved-character, and versioned keys; live-compatible copy must complete with zero relay fallbacks for small and multipart objects. |
| A8-CURL-01 | P1 | Pre-existing path touched by `2f82f01d6` | `CreateFileWriter(ALLOW_OVERWRITE)` deletes the existing remote file before returning the writer. `CurlStreamingWriter` starts uploading from its worker before `Commit` and its destructor only stops the worker, so abandonment can leave a partial new object. The temp-backed path also pre-deletes before its final upload. Release without `Commit`, write failure, or upload failure can therefore lose the old destination or publish partial data. | Public Curl writers stage locally or to a unique remote sibling. Only verified `Commit` may backup/promote to the final name; destruction removes staging only. | Loopback FTP and applicable protocol tests cover streaming and temp-backed writers, a seeded overwrite sentinel, and a previously nonexistent destination. Abandon before Commit and inject failed Write/Commit; the sentinel must remain byte-identical, a new-path abort must leave no final object, and staging must be absent. |
| A8-HOST-01 | P1 | Pre-existing in touched host-services scope | Cross-thread Prompt and Connection Manager release a raw pointer into synchronous `SendMessageW`; receivers call `TakeMessagePayload`, which accepts only opaque registered asynchronous tokens. Calls return `E_POINTER` and leak the allocation. | Introduce an explicit synchronous marshaling helper or retain sender ownership and use a dedicated non-owning synchronous message contract. Raw pointers and posted-payload tokens must never share a receiver protocol. | Worker-thread Prompt and Connection Manager tests prove success/cancel propagation, exact destruction, no registry-token confusion, and safe window teardown. |
| A8-HOST-02 | P1 | Pre-existing in touched ABI normalization | `ShowConnectionManager` validates request size but writes `HostConnectionManagerResult` before validating caller-provided result capacity. An undersized plugin buffer can be overwritten. | Validate the result's required prefix before any write, accept reviewed oversized tails, and leave rejected buffers untouched. | Canary-backed undersized/current/oversized result tests at the real host ABI boundary. |
| A8-HOST-03 | P2 | Pre-existing | `ShowConnectionManagerOnUiThread` initializes `hostWindow` to null and uses it as the fallback owner, leaving the nested modal surface unowned. | Resolve the initialized host/root owner once, validate it, and disable/restore it through the modal lifetime. | Null, invalid, and explicit owner tests prove correct owner selection, disabled state during the dialog, and restoration on every exit. |
| A8-TERM-01 | P1 | Introduced by `c330bf9aeb` | Selected Ctrl+C consumes `WM_KEYDOWN` after copying but does not suppress the translated `WM_CHAR` ETX, so the foreground process can be interrupted after the copy. | Character-generating shortcuts arm exact one-shot character suppression; ordinary unselected Ctrl+C retains terminal interrupt behavior. | Live PTY input test: selected Ctrl+C copies and emits zero bytes; unselected Ctrl+C emits exactly one ETX. Cover Ctrl+Shift C/A/V without relying on later key-state sampling. |
| A8-TERM-02 | P1 | Pre-existing fallback exposed by embedded terminal integration | An unsupported plugin pane with no local backing directory substitutes `GetDefaultFileSystemRoot()` and opens a shell in an unrelated Windows root. | Require a validated local backing path or WSL identity. Unsupported S3/MTP/archive/cloud panes show localized feedback and do not open a terminal elsewhere. | Non-file plugin/no-backing command tests assert no terminal creation and the correct localized error. |
| A8-TERM-03 | P2 | Introduced with terminal location marshaling | `MakeTerminalLocation` classifies every `\\` prefix as UNC, so valid extended drive paths such as `\\?\C:\...` are rejected by the terminal validator. | Reuse `Common::Paths::ClassifyWindowsPath`: drive and extended-drive map to `WindowsLocal`; UNC and extended-UNC map to `WindowsUnc`; reject unsupported classes. | ABI/open/update tests cover drive, UNC, extended drive, extended UNC, WSL, and malformed device paths. |
| A8-TERM-04 | P2 | Introduced with source-follow integration | Pane path change publishes formatted history/display text such as `shortId:context|path` as a Windows-local terminal location. Validation fails and the old local generation/path/follow target remains active. | Derive a typed location from `PaneState`, never from display/history text. Unsupported namespaces send a successful `Detached + Unsupported` update. | Local-to-plugin and plugin-to-local transition tests assert generation monotonicity, detached state, and absence of stale follow behavior. |
| A8-BUILD-01 | P1 | Introduced in Terminal runtime tooling | The uucode archive URI/cache name, pinned archive SHA, and expected internal root describe different identities. The current valid archive fails SHA validation; changing only the hash then fails license extraction. | Freeze URI, archive SHA, cache identity, internal archive root, and license SHA as separately named reviewed constants. | Empty-cache online build downloads/verifies/lists/extracts; warm-cache offline build repeats successfully; mutated archive/root/license cases fail closed. |
| A8-BUILD-02 | P1 | Introduced in Terminal frozen-contract tooling | Repository text files are checked out CRLF under ordinary Windows `core.autocrlf=true`, while evaluator/generator contracts hash raw worktree bytes and expect LF digests. | Force LF for every raw-byte identity input or hash a documented canonical LF/Git-blob representation. Renormalize once without changing semantic content. | Fresh checkout tests with `core.autocrlf=true` and `false` produce identical identities and pass all transition, fixture, ABI, and closeout gates. |
| A8-BUILD-03 | P1 | Introduced in bounded archive preflight | Tar verbose-output parsing accepts English month abbreviations or ISO dates only. Valid French output such as `août 08` is rejected. | Request a guaranteed invariant documented tar format or parse archive headers natively; do not infer size from localized display text. | Mocked English/French/non-ASCII month output plus a real valid tar round trip under the available localized host. Malformed/link/special-file protections remain green. |
| A8-BUILD-04 | P2 | Introduced in cached runtime validation | Cached/final runtime acceptance validates the import list claimed by identity JSON but does not re-derive the DLL's actual imports. | Derive actual imports during cache reuse and final validation and compare the exact ordered/set policy to the reviewed closure. | PE fixture with one extra private import must fail cache reuse and final validation while the valid runtime passes. |
| A8-BUILD-05 | P3 | Introduced in runtime builder | `Copy-TerminalLicenseText` and `Get-TerminalRuntimeIdentityHeaderText` are each defined twice, so later definitions silently replace earlier ones. | Keep one definition of each helper and one focused test surface. | PowerShell parse plus source contract requiring exactly one definition of each named helper. |
| A8-RELEASE-01 | P1 | Introduced by `2f82f01d6` | `release.yml` supports `build_arm64=false` but invokes Winget unconditionally; the reusable Winget workflow always requires the ARM64 portable asset. The GitHub release is created before the required Winget job fails. | Choose and document one coherent contract: require ARM64 releases, skip Winget for x64-only, or generate a supported x64-only manifest. Validate assets before irreversible publication. | Workflow-policy matrix covers ARM64 enabled/disabled, portable/MSIX combinations, release creation ordering, and Winget input assets. Use stubs/test repositories only. |
| A8-ABI-01 | P1 decision routed | Introduced by `31080406d1` | Public request/context layouts changed at the beginning of records while existing COM IIDs remained unchanged. Older plugins can interpret shifted pointer/HWND fields or be rejected without a detectable interface generation. | Resolve `ATLAS-DEC-ABI-01` first. If pre-change binaries remain supported, promote new IIDs/methods and compatibility adapters into a separate executable plan. Otherwise define an explicit breaking ABI generation and reject old modules before interface invocation; do not rely on same-IID record reinterpretation. | Compile frozen old-header plugin fixtures and exercise old-plugin/new-host plus new-plugin/old-host behavior. Pointer fields must never be read under the wrong layout. |
| A8-DXUI-01 | P2 | Introduced by `39151fefd` | Generic tab drag physically reorders the vector, while the FolderWindow preview host treats indices 0/1/2 as permanent Folder/Preview/Terminal identities. Dragging around hidden tabs makes visibility, selection, and close callbacks address the wrong semantic page. | Give tabs stable IDs and use them for visibility/selection/close dispatch, or disable reordering for semantic fixed-page hosts. | Hide Preview, show Terminal, drag across Folder, then show/close/select every page and assert semantic identity and child association. |
| A8-DXUI-02 | P2 | Introduced by `2f82f01d6` | Accessibility normalization maps unsupported Paragraph downward to Line, contrary to the required next-larger supported-unit ordering. | Map Paragraph to the next larger supported unit (Document while Page is unsupported), or implement genuine Paragraph/Page support consistently across Expand/Move/MoveEndpoint. | Shared DxUi and Terminal provider tests cover every `TextUnit` and assert Paragraph never degrades to Line. |
| A8-CI-01 | P1 routed | Pre-existing, already owned | The format job is unsafe for forks, stale-head publication, broad `git add -A`, and concurrent pushes. | Implement only through `CI_FormatAutocommitRace_2026-08-25.md`; this row creates no second queue. | Use that plan's deterministic policy tests and four operator cases. |

## Evidence anchors

Line numbers are review-head anchors. Executors must re-run the drift command and relocate them
before editing.

| ID | Primary implementation evidence | Existing test/spec evidence to correct or extend |
|---|---|---|
| A8-S3-01 | `Plugins/FileSystemS3/FileSystemS3.Directory.cpp:280-285,1745-1756,1778-1785,2550` | `Specs/FileSystem/FileSystem_S3.md`; add fake interleavings rather than relying on final-state-only transfer tests |
| A8-S3-02 | `Plugins/FileSystemS3/FileSystemS3.Directory.cpp:1099-1108,2580-2584` | Incorrect expectations at `FileSystemS3.Directory.cpp:3390-3394,3417-3421`; incorrect current contract at `Specs/FileSystem/FileSystem_S3.md:183-189,226-233` |
| A8-S3-03 | `Plugins/FileSystemS3/FileSystemS3.S3.cpp:154-166,757,1034-1041`; relay masking at `FileSystemS3.Directory.cpp:1554-1585` | Add real AWS request-model header serialization; fake graph calls do not encode headers |
| A8-CURL-01 | `Plugins/FileSystemCurl/FileSystemCurl.DirectoryOps.cpp:1067-1080,1570-1588`; `Plugins/FileSystemCurl/FileSystemCurl.Imap.cpp:3904-3929` | Abort contract at `Common/PlugInterfaces/FileSystem.h:562-567`; current ownership selftest covers neither a sentinel nor abandonment |
| A8-HOST-01 | `RedSalamander/HostServices.cpp:600-601,646-647,1098,1112`; opaque-token contract at `Common/Helpers.h:2625-2639` | Any-thread service contract at `Specs/Plugins/Plugins_PluginAPI.md:142`; add Commands dialog worker-thread cases |
| A8-HOST-02 | `RedSalamander/HostServices.cpp:621-627,1414-1424`; result record at `Common/PlugInterfaces/Host.h:154-165` | Add real ABI-boundary canary tests to `PluginContractTests` or the owning host-service harness |
| A8-HOST-03 | `RedSalamander/HostServices.cpp:1421-1430` | Owner/facade contract at `Specs/Core/Core_ConnectionManager.md:5,13,216` |
| A8-TERM-01 | `Plugins/Terminal/Terminal.cpp:4947-4955,5013-5064`; translation at `RedSalamander/RedSalamander.cpp:8193-8197` | Add executable PTY input coverage; do not certify behavior with regex-only source tests |
| A8-TERM-02 | `RedSalamander/FolderWindow.FileSystem.Navigation.cpp:745-792`, especially `773-776` | No-display-string-guessing contract at `Specs/Terminal/Terminal_EmbeddedPlugin.md:165-167` |
| A8-TERM-03 | `RedSalamander/FolderWindow.FileSystem.Navigation.cpp:73-110`; canonical classifier at `Common/PathUtils.h:71-104`; validator at `Plugins/Terminal/Terminal.cpp:374-388` | Add logical-location ABI/open/update matrix |
| A8-TERM-04 | `RedSalamander/FolderWindow.FileSystem.Navigation.cpp:1002-1042,1191-1219`; display format at `RedSalamander/NavigationLocation.h:645-673` | Distinguish plugin-backed locations with a validated Windows backing path from unsupported plugin locations without one |
| A8-BUILD-01 | `Tools/Build-GhosttyTerminalRuntime.ps1:23-25,233-235,647-651`; invocation at `Directory.Build.targets:105-109` | Add online empty-cache and offline warm-cache paths; revalidate current Terminal qualification claims |
| A8-BUILD-02 | `.gitattributes:4`; `Tools/TerminalEngineEvaluator.psm1:483-500`; `Tools/TerminalEngine/Generate-TerminalEngineFixtures.ps1:881-891` | Transition, fixture-generator, Gate0 ABI, and Round4 closeout Pester cases currently fail under CRLF checkout |
| A8-BUILD-03 | `Tools/TerminalEngine/TerminalEngineArchive.psm1:429-435` | `Tools/Tests/TerminalEngineArchive.Tests.ps1:191-216`; retain archive-safety tests even if verbose-output parsing is removed entirely |
| A8-BUILD-04 | `Tools/Build-GhosttyTerminalRuntime.ps1:383-434,623-625,679-701` | Add a PE fixture with an additional import and prove cache/final rejection |
| A8-BUILD-05 | Duplicate definitions at `Tools/Build-GhosttyTerminalRuntime.ps1:75-108,258-291` | Add an exact-one-definition structural guard only after removing the duplication |
| A8-RELEASE-01 | `.github/workflows/release.yml:16-20,419-427`; `.github/workflows/winget-release.yml:87-90,115-119` | Extend `Tools/Tests/ReleaseWorkflowPolicy.Tests.ps1` and `Tools/Tests/WingetValidation.Tests.ps1` across the combined asset matrix |
| A8-ABI-01 | Leading layout changes in `Common/PlugInterfaces/Viewer.h:21-66` and `Common/PlugInterfaces/Host.h:58-165`; retained IIDs including `Viewer.h:117` | Conflicting policy statements in `Specs/Plugins/Plugins_PluginAPI.md:72-78` and `Specs/Plugins/Plugins_ViewerPlugins.md:9,17-25`; decision owner is `ATLAS-DEC-ABI-01` |
| A8-DXUI-01 | `Common/DxUi/DxUi.Controls.cpp:5837-5892`; fixed semantic indices at `RedSalamander/FolderWindow.cpp:1776-1795` and `FolderWindow.Layout.cpp:1136-1139` | Extend `Tests/DxUiTests/DxUiTests.NewControls.cpp:1060+` and Terminal/Preview Commands navigation cases |
| A8-DXUI-02 | `Common/DxUi/DxUi.AccessibilityTextUnits.h:23-26`; `Common/DxUi/DxUi.Accessibility.cpp:7401-7416` | Correct the wrong Paragraph-to-Line assertion at `Tests/DxUiTests/DxUiTests.Accessibility.cpp:34-41`; cover Expand/Move/MoveEndpoint and Terminal |
| A8-CI-01 | `.github/workflows/ci.yml:41,84-90` | Sole executable owner and complete proof matrix are in `CI_FormatAutocommitRace_2026-08-25.md` |

## Primary source constraints used by the review

- AWS S3 `DeleteObject`: in a versioning-enabled bucket, key-only deletion inserts a delete
  marker; specifying `versionId` permanently deletes that version:
  <https://docs.aws.amazon.com/AmazonS3/latest/API/API_DeleteObject.html>.
- Microsoft UI Automation text units are ordered Character, Format, Word, Line, Paragraph,
  Page, Document, and an unsupported unit defaults to the next larger supported unit:
  <https://learn.microsoft.com/windows/win32/api/uiautomationcore/nf-uiautomationcore-itextrangeprovider-expandtoenclosingunit>.
- The exact bundled AWS SDK serializer is also test authority for `SetCopySource`; tests must
  exercise request-header serialization rather than only fake provider semantics.

## Architecture and simplification requirements

These are constraints on the fixes, not optional cleanup:

1. **Mutation journals carry identities, not names alone.** S3 rollback authority must be tied
   to the exact published revision.
2. **Writers stage, verify, then publish.** Overwrite permission never authorizes destruction
   before `Commit` succeeds.
3. **Terminal locations remain typed end to end.** History/display formatting cannot become an
   ABI or execution input.
4. **Semantic tabs use stable IDs.** Vector position is layout state, not business identity.
5. **Synchronous and posted message ownership are visibly distinct.** A raw synchronous pointer
   and an opaque posted-payload token never enter the same adoption helper.
6. **Frozen build inputs have one manifest of independent identities.** URI, archive bytes,
   internal root, extracted license, generated content, and actual DLL imports are not conflated.
7. **Breaking plugin ABI changes are detectable before dereference.** New IID/generation or a
   loader-level rejection precedes reading a changed record layout.

Before adding any helper, search `Specs/Core/Core_SharedHelpers.md`, `Common/`, and
`Tests/TestSupport/`. Extend a canonical helper when its semantics match; document any intentional
domain-specific difference.

## Authoritative specifications to update during implementation

| Rows | Required durable owner updates |
|---|---|
| A8-S3-01..03 | `Specs/FileSystem/FileSystem_S3.md`, plus `Specs/Plugins/Plugins_VirtualFileSystem.md` where public writer/rollback behavior is affected |
| A8-CURL-01 | `Specs/FileSystem/FileSystem_FtpSftpScp.md`, `Specs/Plugins/Plugins_VirtualFileSystem.md`, and the applicable file-operation contract in `Specs/FileSystem/FileSystem_FileOperations.md` |
| A8-HOST-01..03 | `Specs/Plugins/Plugins_PluginAPI.md`, `Specs/Core/Core_ConnectionManager.md`, and public headers |
| A8-ABI-01 | `Operation_Atlas_RemainingSpecificationDecisions_2026-08-04.md` first; after decision, `Specs/Plugins/Plugins_PluginAPI.md`, `Specs/Plugins/Plugins_ViewerPlugins.md`, public headers, and a separately promoted implementation plan when required |
| A8-TERM-01..04 | `Specs/Terminal/Terminal_EmbeddedPlugin.md`; update `TerminalEngineDecision.md` only if engine/tooling decision evidence changes |
| A8-BUILD-01..05 | `Specs/Build/Build_Toolchain.md`, the dependency-identity sections of `Specs/Terminal/Terminal_EmbeddedPlugin.md` and `Specs/Terminal/TerminalEngineDecision.md`, Terminal engine identity schemas/contracts where fields or canonicalization change, and testing guidance for checkout/locale matrices |
| A8-RELEASE-01 | `Specs/Installer/WingetIntegration.md`, `Specs/Installer/Installer_Msix.md`, and the authoritative portable/release packaging contract |
| A8-DXUI-01..02 | Current DxUi/UI and accessibility contracts, plus shared Terminal accessibility text-unit policy |
| A8-CI-01 | The documentation named by `CI_FormatAutocommitRace_2026-08-25.md` only |

## Execution order

### Wave 0 — failing proofs before production edits

1. Add deterministic red tests for A8-S3-01, A8-S3-02, A8-CURL-01, A8-HOST-01,
   A8-HOST-02, A8-TERM-01, and A8-DXUI-01.
2. Add pure/fixture tests for copy-source serialization, path classification, text-unit
   normalization, workflow asset matrices, checkout EOL invariance, archive-root identity, tar
   locale handling, and actual PE imports.
3. Record exact failures without modifying expected results to match unsafe current behavior.

### Wave 1 — data integrity

1. A8-S3-01
2. A8-S3-02
3. A8-CURL-01
4. A8-S3-03

No build, UI, or architecture cleanup may delay these rows once their deterministic proofs exist.

### Wave 2 — host and terminal correctness

1. A8-HOST-01 and A8-HOST-02
2. A8-TERM-01 and A8-TERM-02
3. A8-TERM-03 and A8-TERM-04 through one typed-location repair
4. A8-HOST-03

### Wave 3 — restore reproducible builds and coherent publication

1. A8-BUILD-01, then rerun the clean-cache build far enough to expose all later blockers
2. A8-BUILD-02 and A8-BUILD-03
3. A8-BUILD-04 and A8-BUILD-05
4. A8-RELEASE-01
5. A8-CI-01 only through its existing owner

### Wave 4 — ABI, UI identity, and accessibility

1. Route A8-ABI-01 to `ATLAS-DEC-ABI-01`; implementation requires an explicit compatibility/product decision and a promoted executable plan when adapters/new IIDs are required.
2. A8-DXUI-01
3. A8-DXUI-02

### Wave 5 — closeout

1. Merge every durable behavior into the authoritative specifications listed above.
2. Run focused tests, a clean Debug/x64 full-solution rebuild, test-enabled Release proof where
   relevant, and `Tools/Run-AllTests.ps1 -Suite Full`.
3. Run Release x64 and ARM64 packaging/release-policy validation without production publication.
4. Archive required performance and live-compatible cloud evidence under `Specs/TestRuns/`.
5. Reconcile older plan claims, clear the artifact contamination marker only through a successful
   repository-authorized rebuild, and move this plan to `../Done/`.

## Verification commands

Commands are starting points; executors must use the repository artifact lock and the exact
focused filters introduced with each row.

```powershell
# Documentation/index consistency
.\Tools\Get-SpecInventory.ps1 -FailOnFindings
Invoke-Pester -Path .\Tools\Tests\SpecInformationArchitecture.Tests.ps1 -PassThru

# Clean build recovery and broad gate
.\build.ps1 -Rebuild -Configuration Debug -Platform x64
.\Tools\Run-AllTests.ps1 -Suite Full

# Static hygiene
git diff --check
git status --short
```

Cloud tests must default to deterministic fakes or local/live-compatible test endpoints. Tests
that can delete remote user data, publish a GitHub release, submit Winget, push a branch, or use
production credentials require explicit operator authorization and a disposable target.

## Definition of done

This plan moves to `../Done/` only when:

- every owned executable `A8-*` row is implemented or explicitly rejected by an authorized
  product decision recorded in the authoritative specification;
- `A8-CI-01` and `A8-ABI-01` are linked to their owners' final dispositions without duplicate
  implementation;
- every unsafe current test expectation is corrected and each focused regression is green;
- authoritative domain specs describe the resulting durable behavior;
- a clean Debug/x64 rebuild, the canonical Full suite, applicable Release/test-enabled gates,
  and packaging policy tests are green;
- required S3/terminal/render/build evidence is archived and linked under repository policy;
- no artifact contamination marker remains; and
- the source worktree is clean except for the intentional reviewed closeout changes.

## STOP conditions

- Stop ABI implementation if support for pre-change external plugin binaries is not explicitly
  decided; do not silently preserve the same IID with incompatible layout semantics.
- Stop S3 live validation if the endpoint/bucket is not disposable, versioning state is unknown,
  or conditional-delete semantics cannot be proved without risking user objects.
- Stop release/Winget validation before any real publication, submission, or credential use unless
  the operator explicitly authorizes that external side effect.
- Stop a proposed shared abstraction if it weakens provider-specific cancellation, rollback,
  ownership, module-unload, or performance guarantees.
- Stop closeout if the Full suite is bypassed merely because the Terminal runtime build remains
  blocked; a documented blocker is not a green gate.
