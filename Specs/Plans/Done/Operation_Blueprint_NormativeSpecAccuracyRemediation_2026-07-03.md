# Operation Blueprint — Normative Spec Accuracy Remediation (2026-07-03)

Tracking plan for the discrepancies found by a full audit of **every normative spec under
`Specs/<Domain>/`** (the 52 `.md` contracts in `Core/`, `FileSystem/`, `Installer/`, `Plugins/`,
`Testing/`, `UI/`) against the working tree at HEAD `275c04034` (dirty; WarpDrive/DxUi work in
flight). This ledger is the **single owner of spec-vs-code drift remediation** — no existing WIP
plan tracks documentation accuracy; the execution ledgers (Granite/Clearwater/Tailwind/IronLedger/
Keystone/Floodgate/Farsight) track code defects. Where a Blueprint item overlaps one of those, it is
routed (see the routing table) and executed there.

> **This plan's job is to close the gap between what the specs say and what the code does — by
> fixing whichever side is wrong.** Every item is tagged `fix-spec` (code is the intended, shipped
> behavior; the doc is stale/wrong), `fix-code` (the spec states a sensible safety/security/parity/
> a11y contract the code violates), or `decide` (genuinely ambiguous; a human picks the contract).

## Method & confidence

- One deep reader per spec: read the whole file, enumerated every checkable normative claim
  (behaviors, invariants, named files/functions, settings keys/defaults, shortcut & capability
  tables, perf budgets, error/threading rules, test names, formats), and verified each against the
  code with grep/read. Completeness (undocumented shipped behavior) and quality (precision,
  internal consistency, dead references) judged per spec.
- **All 16 HIGH-severity findings were independently adversarially verified** (a second agent told
  to refute each). Result: **12 CONFIRMED, 2 ADJUSTED (severity lowered, still valid), 1 REFUTED
  (dropped), 1 REFUTED-in-part** — plus additional MEDIUM verifications where the pipeline reached
  them. Findings marked `[CONFIRMED]` survived refutation; `[unverified]` means evidence-backed by
  the reviewer but not yet through a second pass (verify RED-first before acting).
- Line anchors are planned-at the dirty working tree of 2026-07-02/03. **Re-check drift with
  `git diff <file>` before executing** — this ledger does not auto-track the tree, and several
  anchors sit in files another session is actively editing.

## Headline results

- **52/52 specs audited.** Verdicts: **7 ACCURATE**, **34 MINOR_DRIFT**, **6 MAJOR_DRIFT**,
  **5 NOT_NORMATIVE** (docs that are snapshots/memos/redirects, not contracts).
- **222 findings**: 16 high / 99 medium / 107 low → **167 fix-spec, 27 fix-code, 28 decide.**
  The tree is in good shape: the dominant mode is *docs lagging intended code* (fix-spec), not code
  violating contracts. But a **hard core of real defects and dangerous doc claims** exists — those
  are Tracks A/B below.
- **Quality:** 30 specs `good`, 17 `adequate`, **3 `poor`** (`FtpScpSftpPerformanceComparison`,
  `UI_PreferencesDialog_MigrationHistory`, and the migration-history quality is dragged down by
  staleness, not structure). The 5 NOT_NORMATIVE docs should be relabeled/moved (Track D).

## Execution rules (inherited from Granite/Clearwater, still binding)

- This file is a **triage ledger**. Before executing an item, expand it into a self-contained
  sub-plan: planned-at SHA, drift check, in/out-of-scope files, exact verification command, expected
  output, STOP conditions.
- **`fix-code` items need a RED-on-bug test first** (self-test that fails on current code, passes
  after the fix). Green = `.\Tools\Run-AllTests.ps1 -Suite Full` (`AGENTS.md:134`).
- **`fix-spec` items** are documentation edits: update the spec to describe the shipped behavior,
  keeping its normative structure. Where the code is safety-relevant (S3 delete/transfer, MSI
  service, ABI struct), the corrected spec text must be precise enough to audit against.
- **`decide` items** are blocked on a human product/architecture decision. Do not silently pick a
  side; record the decision, then the resolution becomes a `fix-spec` or `fix-code` task.
- Closeout: when a spec's rows are all resolved, note it in the appendix; when the whole plan is
  done, move to `Specs/Plans/Done/`.

## Routed to existing plans — do NOT duplicate here

| Owner | Blueprint item(s) | Note |
|-------|-------------------|------|
| **RedConfigure_LocalizationThemeManagerPlan** (Done) | BP-C4 (RedConfigure export gate, `fix-code`), BP-D3 (RedConfigure never registers with the localization manager, `decide`), BP-S2 (invalid target-text silently reverted, `decide`) | Completed in `Specs/Plans/Done/RedConfigure_LocalizationThemeManagerPlan.md`; durable contracts live in Core_RedConfigure, UI_RedConfigure, and Core_Localization. |
| **FileSystem_GoogleDrivePluginPlan** (WIP) | BP-B3 (`googleDocsMode` default export-vs-native, `decide`) | The default only matters once the Docs export pipeline (that plan's Phase 3+) ships. Decide the default there; align spec + initializer in the same change. |
| **Fix-to AWS-S3-crash** (WIP) | — | No overlap. That plan is AWS-CRT unmap safety; Blueprint's S3 items are delete/transfer **doc** drift. |
| **Keystone** (WIP, KS-2/KS-3) | — | No overlap. KS covers `caseOnlyRename` capability on Curl/MSDrive; Blueprint's MSDrive item is a timeout-schema mismatch. |
| **DxUi_Uia_ContinuationBaton** | — | No overlap; Blueprint has no a11y-snapshot items. UI_DxUiSharedGrid drift is doc-only. |

Everything below is owned by Blueprint unless routed above.

---

## Track A — `fix-code`: code violates a sensible spec contract (real defects)

Ordered by severity/impact. Each needs a RED-first self-test. `[CONFIRMED]` = survived adversarial
verification.

| ID | Sev | Spec | Defect (anchor) | Fix direction | Verify |
|----|-----|------|-----------------|---------------|--------|
| **BP-A1** | **HIGH** `[CONFIRMED]` | Core_ConnectionManager | **Cloud profiles created in the Connection Manager UI are unusable.** New profiles default `authMode=Password` (`ConnectionManagerWindow.cpp:4507`); the editor never sets `OAuth2Pkce` (protocol-change handler is MTP-only `:2749-2750,569-631`; field-commit toggles only Anonymous/Password `:2650-2658`; oauth2 UI state requires the mode to *already* be set `:2503`). Both plugins hard-reject non-oauth2Pkce (`FileSystemGoogleDrive.cpp:1534`, `FileSystemMicrosoftDrive.cpp:2618-2621`) → a saved Google Drive/OneDrive/SharePoint profile fails at connect with `ERROR_INVALID_DATA`. Only Quick Connect coerces the mode (`ConnectionSecrets.cpp:266-269`). | Coerce `authMode=OAuth2Pkce` (and hide the password editor) when the protocol is a cloud provider — mirror the Quick Connect path into the saved-profile editor. Spec is right. | RED: create+connect a GDrive profile via CM. |
| **BP-A2** | **HIGH** `[CONFIRMED]` | Core_Localization | **Plugin satellite DLLs are unreachable at runtime.** The loader probes `<pluginDir>\Lang\<Owner>-<culture>.dll` (`LocalizationManager.cpp:151-158`, owner registered with the plugin HINSTANCE, plugins load from `<exe>\Plugins`), but every satellite ships to the **shared exe-level `Lang\`** (`Directory.Build.props:78-81`, deployment test `RedSalamanderPluginDeployment.Tests.ps1:178-192`, MSI `build-msi.ps1:76-106`; no `Plugins\Lang` exists). ViewerText's translations (`ViewerText.vcxproj:115`, `ViewerText.cpp:521-525`) are dead weight; fr-FR gets English plugin UI. | Loader should also probe the **host executable's `Lang\`** for plugin owners (or deployment must create `Plugins\Lang`). Add a plugin-owner regression test (existing `LocalizationTests` only cover an exe-shaped owner). | RED: satellite lookup for a plugin owner. |
| **BP-A3** | **HIGH** `[RESOLVED 2026-07-14]` | Core_RedConfigure | The missing export gate and invalid-translation pass-through were fixed by the completed RedConfigure manager plan. | Combined validation now blocks errors and requires explicit warning acknowledgement before atomic export; written files are reparsed before success. | Covered by `RedConfigureTests` validation/export cases. |
| **BP-A4** | **HIGH** `[CONFIRMED]` | Plugins_PluginAPI | **Pane-scope alert cookie routing is not implemented.** `ShowAlertOnUiThread` discards the cookie (`HostServices.cpp:1848`) and routes pane-scope alerts to `GetFocusedPane()` (`:1913-1917`); `ClearAlert` dismisses the focused pane's overlay (`:1947-1956`); prompts ignore the cookie (`:1959`, `:291-314`). A background op on the *non-focused* pane paints its alert on the wrong pane, and `ClearAlert` can dismiss an unrelated (even fatal) alert. The ABI ships cookie params as dead weight. | Implement cookie→pane mapping, **or** (minimum) downgrade the spec's MUST to "focused pane" semantics and stop shipping the cookie param. The cookie design is the safer contract — prefer implementing it. | RED: alert from a non-focused-pane operation. |
| **BP-A5** | MED | Core_ConnectionManager | **`windowsHelloReauthTimeoutMinute` is inert for non-zero values.** Both call sites OR the timed check with the untimed one (`HostServices.cpp:1762-1774`, `ConnectionManagerWindow.cpp:3216-3223`), so once any interactive auth succeeds, Windows Hello is suppressed for the rest of the run regardless of the configured minutes. Only `timeout==0` (always prompt) behaves as documented. | Drop the `|| HasSecretAccessAuthorization(...)` fallback where a timeout is set, so the timer actually re-prompts. Spec contract is right (this is a security control). | RED: reveal secret after timeout elapses. |
| **BP-A6** | MED | Core_ConnectionManager | **Editor accepts `/` and `\` in profile names** despite the `@conn-safety` rule. `ValidateConnectionProfileName` checks only empty/reserved/duplicate (`ConnectionManagerWindow.cpp:687-725`); slash sanitization exists only for auto-generated names and at load time (silent rewrite to `-`, `SettingsStore.cpp:3175-3181`). Until reload, `nav:a/b` resolution splits at the slash and fails "Connection not found". | Reject `/` and `\` at edit-commit time (or sanitize consistently). Spec is right. | RED: save a profile named `a/b`, resolve it. |
| **BP-A7** | MED | Core_SettingsStore | **`connections.allowInsecureTlsInAutomation` is dropped on load** when it is the only non-default value. The keep-section gate `hasNonDefaultGlobals` checks only `bypassWindowsHello` + `windowsHelloReauthTimeoutMinute` (`SettingsStore.cpp:3207-3213`), so a connections object with just this flag is discarded and lost on next save, though the writer persists it when the section survives (`:6562-6565`). | Add `allowInsecureTlsInAutomation` to the keep-section gate. Spec (and `SettingsStore.schema.json:1495-1500`) are right. | RED: round-trip settings with only this flag set. |
| **BP-A8** | MED | Core_SettingsStore | **Settings-file themes bypass shared validation.** The disk loader uses strict `ParseThemeDefinitionJson5` (InvalidId/InvalidColorKey/duplicate rejection), but the settings-file path `ParseTheme` is a separate lenient hand-rolled parser (`SettingsStore.cpp:1140-1229`) accepting any id/color-key/baseThemeId and duplicates — violating the spec's anti-divergence mandate. | Route settings-file themes through the same `ThemeDefinitionIo` validation as disk themes. Spec is right. | RED: settings theme with a bad color key. |
| **BP-A9** | MED | FileSystem_Mtp | **Watchdog cancel runs plugin code with no module keepalive.** On timeout, `QueueBackendCancel` submits a threadpool item whose lambda + `IMtpBackend` vtable live in `FileSystemMtp.dll` (`FileSystemMtp.Core.cpp:1757-1771,1721-1731`) with no `GetModuleHandleEx`/`FreeLibraryWhenCallbackReturns` pin; `CanUnloadFileSystemMtpModule` only tracks quarantined jthreads (`:3636-3642`). Unload during a pending cancel → code executed from an unmapped DLL. | Pin the module across the queued cancel (the spec's unload contract requires it). | RED: unload with a cancel item pending (fault-inject). |
| **BP-A10** | MED | Plugins_PluginAPI | **ViewerImgRaw clears window alerts with a `nullptr` cookie** → `E_INVALIDARG`, HRESULT discarded, so stale decode-failure alerts never clear (`ViewerImgRaw.Decode.cpp:1506,2947`, `ViewerImgRaw.cpp:1917,2704`; host rejects null-cookie WINDOW scope `HostServices.cpp:1931-1937`). ViewerText passes the HWND correctly. | Pass the window HWND as the clear cookie (match ViewerText). | RED: decode-fail then succeed; assert alert cleared. |
| **BP-A11** | MED | Plugins_PluginAPI | **`ViewerPluginManager::Unload` skips the `RedSalamanderPluginCanUnloadNow` deferral** the spec mandates for all host plugin managers (`ViewerPluginManager.cpp:1072-1128` vs the compliant `FileSystemPluginManager.cpp:1241-1358`). Latent for built-ins (no in-tree viewer exports it) but live for third-party viewer DLLs. | Add the CanUnloadNow → defer → sweep path to `ViewerPluginManager::Unload`. Spec is right. | RED: viewer that reports not-unloadable. |
| **BP-A12** | MED | Plugins_ViewerWeb | **`allowExternalNavigation` = "Block" does not block for the Markdown viewer.** The handler consults the setting only when `_kind == Web` (`ViewerWeb.cpp:3218-3225`); Markdown/JSON always cancel + hand http/https to the system browser (`:3233-3238`), even though the Markdown schema exposes Block/Allow (`:1339-1375`). A user who sets Block still gets arbitrary URLs launched. | Honor `allowExternalNavigation` for all kinds. Spec is right (security). | RED: Block set, click a link in a rendered .md. |
| **BP-A13** | MED | FileSystem_S3 | **S3 Table capabilities advertise `write:true` + cross-FS import, but every mutating op returns `ERROR_NOT_SUPPORTED`** (`FileSystemS3.IO.cpp:1303-1306`, `Directory.cpp` mode guards; `kCapabilitiesJsonS3Table` at `FileSystemS3.h:333-365`). The host plans pastes/imports into `s3table:` paths that can only fail at `CreateFileWriter`. | Set the S3-Table capability JSON to read-only (`write:false`, no import). Spec/host contract is right. | RED: paste into an s3table: path. |
| **BP-A14..A27** | LOW | (various) | 14 low-severity `fix-code` items: compiler-warning hygiene (`Core_CompilerWarnings` bare C5039 pragma, redundant DxUi `/FS`), locale-cache invalidation not reaching ViewerSpace, MTP picker PnP-id fallback loss, RedSalamander skips resource-owner registration on settings-load failure, symbols-MSI omits the service PDB, winget placeholder SHA, relayout-churn test misses focus re-assert, self-test coverage silently skipped when results.json missing, Preferences/Compare windows don't register app icons, etc. | Each is a small, safe code fix; batch by area. See `scratchpad/actionable.txt` for full anchors. | per item |

---

## Track B — `decide`: spec and code both defensible; needs a human ruling

These are genuine product/architecture forks. Record the decision; it becomes a fix-spec or fix-code
task. The two HIGH ones gate real behavior.

| ID | Sev | Spec | The fork | Resolve by |
|----|-----|------|----------|-----------|
| **BP-B1** | **HIGH** `[CONFIRMED]` | Core_StartupBootstrap | **Shutdown-abort MUST is inverted.** Spec: if an auxiliary window refuses to close, *abort* shutdown and keep running (`:115,124`). Code: logs a warning and proceeds to teardown unconditionally (`RedSalamander.cpp:10940-10968,11155-11159`). Never-abort is deliberate (avoids trapping users behind a hung window) but the spec's rule exists to prevent D2D teardown breaks. | Pick: rewrite the spec to "best-effort close, warn, proceed after posting deferred finalize", **or** restore abort semantics. Then align the other side. |
| **BP-B2** | MED | Core_CompareDirectories | **"Access denied" per-item detail line never renders.** Spec (and its manual-test checklist) promise the failure reason in the details line; `IDS_COMPARE_DETAILS_ACCESS_DENIED` is localized in every language but referenced by no code — failed folders show generic "content differs" (`CompareDirectoriesWindow.cpp:324-361,2475-2616`; engine never caches failures `Engine.cpp:144-147,2447`). | Wire the failure reason through, **or** drop the claim + the checklist item. The localized-everywhere string suggests it was intended. |
| **BP-B3** | MED | Core_ConnectionManager | **Connect-time prompt-and-persist for `savePassword=true` with no stored password is not implemented** — the shipped flow prompts session-only and never writes WinCred (`ConnectionManagerWindow.cpp:4624-4662`, `HostServices.cpp:2231-2242`). A profile with "Save password" checked but nothing typed re-prompts every run. | Pick: implement CM-side prompt+persist, **or** rewrite the spec bullet to the session-only model and clarify what the checkbox means. |
| **BP-B4** | MED | Core_ConnectionManager | **Per-connection `copyMoveMaxConcurrency`/`deleteMaxConcurrency` honored only by FileSystemCurl**, but the UI shows the fields for S3/GDrive/OneDrive too (`ConnectionManagerWindow.cpp:2509,2560-2563`); those plugins ignore them. Spec says "all file-op capable protocols". | Pick: implement in cloud plugins, **or** scope the spec to curl-backed protocols and hide the UI fields elsewhere. |
| **BP-B5** | MED | Core_RedSalamanderMonitor | **"Save As" saves ALL lines, ignoring the active filter and dropping metadata prefixes** (`Document.cpp:901-931`), but the API comment says "Save visible content". | Pick: "save what I see" (iterate `_visibleLines`, decide prefixes) or "export full raw log" (fix the comment). User-visible data contract. |
| **BP-B6** | MED | FileSystem_S3 | **Documented `//@conn` shorthand authority is unreachable dead code in the plugin** — it only works because the host rewrites `s3://@conn/...` before the plugin sees it; a direct plugin call resolves `@conn` as a literal bucket name (`ResolveAwsContext`). | Pick: fix the unreachable branch (honor the shorthand at the plugin boundary) or document it as a host-level alias normalized before the plugin. |
| **BP-B7** | MED | Plugins_PluginAPI | **`HOST_ALERT_SCOPE_PANE` has no distinct surface** — alerts alias the pane-content overlay, prompts anchor to the whole window (`HostServices.cpp:1888-1917,291-314`). The APPLICATION overlay was built; the PANE one was not, yet the enum value ships in the public ABI. | Pick: finish the PANE overlay per the plan, or document that PANE degrades to PANE_CONTENT/APPLICATION today. |
| **BP-B8** | MED | Plugins_ViewerImgRaw | **`.tif` is treated as a RAW extension for sidecar pairing** (`ViewerImgRaw.Internal.h:110`), so `scan.tif` + `scan.jpg` collapse into one pair and Thumbnail mode shows the JPEG, never the TIFF; the `.tif` raw-only decode path is then dead code. | Pick: document `.tif` pairing explicitly, or restrict pairing to the true RAW list. |
| **BP-B9** | MED | Plugins_ViewerSqlite | **Embedded preview keeps the filename combo + full header** (`ViewerSqlite.cpp` has no `_embeddedMode` layout branch), contradicting the parent `Plugins_ViewerPlugins.md:91` embedded-chrome contract it extends; sibling ViewerPE hides its combo. | Pick: hide the header row in embedded mode for parity, or amend both specs to bless the visible DX header for SQLite. |
| **BP-B10** | MED | UI_DxUiWinUIDesign | **Type-ramp sizes drift from the normative token table** — code has Body=13, Title=24, Icon=12 DIP (`DxUi.Typography.h:138,144,148`); spec table says 14/28 and explicitly "16×16 icon grid". | Pick: raise code to the WinUI-aligned sizes, or correct the table to the shipped values. The doc's own hedging makes intent ambiguous. |
| **BP-B11** | MED | UI_RedConfigure | **Invalid target-text edits are silently reverted** (`RedConfigureSession.cpp:1074-1085`), losing typed work and desyncing the editor from the export preview; spec says edits "update the in-memory model immediately". *(Cluster with BP-A3/BP-B via RedConfigure plan.)* | Pick: commit invalid text (surface via problem-row machinery) or amend the spec to state invalid edits are discarded on selection change. |
| **BP-B12** | MED | UI_TopLevelToolWindows | **Compare Directories uses `Primary` backdrop target, not `Tool`** (`CompareDirectoriesWindow.cpp:1505`), and its domain spec `Core_CompareDirectories.md` defines no alternate target, so the tool-window MUST's escape clause isn't satisfied. | Pick: switch code to `Tool`, or document the `Primary` choice in `Core_CompareDirectories.md`. |
| **BP-B13** | MED | UI_VisibleTypographyAudit | **RedSalamanderMonitor has live GDI menu-font/text measurement** (`ColorTextView.cpp:4866-4924`) — an app-owned themed surface — that the audit tool's `sourceRoots` deliberately exclude, so the spec's "zero hits" claim is only true by scope. | Pick: state Monitor is out of scope (and why), or bring Monitor's find-panel typography onto the DirectWrite path. |
| **BP-B14..B28** | LOW | (various) | 15 low `decide` items: FTPS not exposed by FileSystemCurl (creds travel plaintext over FTP — worth a security decision), splash-screen hardcoded strings, `maxResults` dead persisted state, Monitor recovery-warning scope, ViewerPE Esc-during-parse & modal export errors, ViewerImgRaw zoom-glyph menu text, ViewerSpace spinner count basis, BatchRename pane-match identity profile, MicrosoftDrive timeout schema-vs-runtime min, etc. | See `scratchpad/actionable.txt`. | per item |

---

## Track C — `fix-spec`: update the doc to match intended, shipped code (167 findings)

The bulk. Grouped by spec; **prioritize the HIGH/`MAJOR_DRIFT` docs first** because they actively
mislead an auditor. Each is a documentation edit — rewrite the stale claim to describe the shipped
behavior, preserving the spec's normative structure.

### C.1 — HIGH-priority spec rewrites (spec asserts the *opposite* of shipped behavior)

| ID | Sev | Spec | What the spec wrongly says → what ships | Verify |
|----|-----|------|------------------------------------------|--------|
| **BP-C1** | **HIGH** `[CONFIRMED]` | FileSystem_S3 | "Copy/Move/Rename are **not** server-side ops" → a **full server-side transfer engine ships** (CopyObject/multipart UploadPartCopy ≥64 MiB, conflict prompts, `.rs-bak` backups, rollback journal; `FileSystemS3.Directory.cpp:1326-1696,1923-2393`). A reader skips reviewing the data-safety-critical engine. | Rewrite §Operations line 133. |
| **BP-C2** | **HIGH** `[CONFIRMED]` | FileSystem_S3 | "folder/prefix + recursive deletes are **denied**" → **recursive prefix delete ships** with 1000-key batching under `FILESYSTEM_FLAG_RECURSIVE` (`Directory.cpp:1224-1266,1168-1222`). **Dangerous doc**: an auditor signs off on a data-safety guarantee the code doesn't provide. | Rewrite §Operations line 132; state the recursive-delete contract precisely. |
| **BP-C3** | **HIGH** `[CONFIRMED]` | Installer_Msi | "only RedSalamander.exe + Monitor.exe ship" → the MSI also ships **RedSalamanderSearchService.exe and registers it as an auto-start LocalSystem Windows service** (`Product.wxs:47-72`, opt-out `INSTALLSEARCHSERVICE`). **Security-relevant**: the most privileged MSI action is undocumented and the spec asserts the opposite payload. | Document the service component, default-on behavior, and opt-out. |
| **BP-C4** | **HIGH** `[CONFIRMED]` | Plugins_VirtualFileSystem | The reproduced `FileSystemOptions` struct **omits the third ABI field `copyMoveMaxConcurrency`** (`FileSystem.h:66-78`). A plugin author modeling it from the spec builds a struct 4 bytes short; the exact-`sizeBytes` check fails every options-bearing call with `E_INVALIDARG`. **ABI-breaking doc.** | Add the field + its 0=default/clamp semantics; re-sync the other stale ABI snippets (misnamed `RequestNavigate`→`NavigationMenuRequestNavigate`, missing `IFileReader` section, `reserved` host-extension channel, 2 missing search-warning flags, 5-arg factory call). |
| **BP-C5** | **HIGH** `[CONFIRMED]` | Testing_TestCoverage | §3 **omits 47 of 120 active FileOperations phases** (Clearflow/Riptide/Fairstream/Floodgate families, `FolderWindow.FileOperations.SelfTest.cpp:678-810`) and its "116 = 120 + setup + cleanup" arithmetic is impossible. A reader concludes ~39% of the FileOps suite doesn't exist. | Regenerate §3 from `kFileOpsPhaseOrder`; fix the arithmetic. |
| **BP-C6** | **HIGH** `[CONFIRMED]` | Testing_TestCoverage | §1 **omits the entire Batch Rename family** (69 `cmd_pane_batchRename_*` cases in a 13th included file, `Commands.SelfTest.cpp:470`) and claims only "12 family files". | Add the Batch Rename family + case table; correct the file count to 13. |
| **BP-C7** | **HIGH** `[CONFIRMED]` | UI_PreferencesDialog_MigrationHistory | "Schema-Driven UI Generation (IMPLEMENTED)" — **removed from the live UI** (`.squad/decisions.md:877-878`; `PrefsUi::CreateSchemaControl*` have zero matches; parser survives only as a test fixture). An engineer adding `x-ui-*` gets nothing. *(This whole doc is NOT_NORMATIVE — see Track D; but the false "IMPLEMENTED" claim must be corrected regardless.)* | Past-tense or delete §Schema-Driven UI. |
| **BP-C8** | MED `[ADJ→med]` | Core_CompilerWarnings | Documented shared suppression baseline is only C4710/C4711 → **all 7 production projects also suppress C4514 + pass `/wd5045 /wd4820`**, and both shared directory props add them (`RedSalamander.vcxproj:94,97` et al). *(Verifier lowered HIGH→MED; direction fix-spec holds.)* | Document C4514/C4820/C5045 as part of the shared optimizer-noise baseline; update the Project File Policy XML snippet. |
| **BP-C9** | MED `[ADJ→med]` | UI_VisualStyle | The "Owner-draw arrow gotcha (MUST)" section mandates an `MFT_OWNERDRAW`/`WM_DRAWITEM`/`ExcludeClipRect`/`SM_CXMENUCHECK` menu mechanism the code has **deliberately abandoned** (menus are fully DxUi-rendered). *(Verifier lowered HIGH→MED.)* | Remove/rewrite the MUST to the DxUi menu-rendering reality. |
| **BP-C10** | MED | Plugins_ViewerSpace | Invocation shortcut documented as **Shift+F3**; actual default is **Alt+F10** (`ShortcutDefaults.cpp:212`; Shift+F3 is `openCurrentFolder`). Anyone testing/rebinding hits the wrong chord. *(Originally HIGH; the shortcut error is the same class as the other single-value table errors.)* | Correct line 16 to Alt+F10. |

> **Dropped by verification:** Testing_TestCoverage "every headline inventory count is stale / fails
> the doc-drift lint" was **REFUTED** — the lint passes live and the finding misquoted the header
> numbers. The two *structural* omissions (BP-C5, BP-C6) are the real, confirmed drift. Do **not**
> mass-rewrite the headline counts on the strength of the refuted item; reconcile only C5/C6.

### C.2 — MEDIUM/LOW spec rewrites (per spec)

~150 remaining fix-spec findings, all evidence-backed, mechanical to apply. They cluster by spec —
work them one spec at a time so each doc lands internally consistent. Full anchors live in
`scratchpad/digest.md` (per-spec) and the appendix below gives per-spec counts. Notable clusters:

- **Core_CompareDirectories** (7 fix-spec): stale `InvalidateForAbsolutePath` "ancestors stay
  cached" wording (code erases them intentionally), the "O(n) subtree invalidation" note (the
  `lower_bound` optimization already shipped), stale `DisplayMode` enum snippet (missing
  `Thumbnails`), progress run-id mechanism mis-described, IFileSystem-wrapper responsibilities
  misattributed.
- **Core_SettingsStore / Installer_Msix / Core_RedSalamanderMonitor / UI_NavigationView /
  UI_CommandMenuKeyboard / Testing_SelfTestRemoteCredentials** (5–7 each): setting keys, dialog
  inventories, shortcut/command tables that drifted as features shipped.
- **Plugins_VirtualFileSystem** (7 fix-spec beyond BP-C4): reproduced ABI snippets lagging the real
  `Common/PlugInterfaces/*.h`.

---

## Track D — NOT_NORMATIVE: relabel/relocate (5 docs)

These live in `Specs/<Domain>/` (implying normative contracts) but are memos/snapshots/redirects.
They mislead by *location*, and some carry stale claims presented as current.

| ID | Doc | What it actually is | Action |
|----|-----|--------------------|--------|
| **BP-D1** | UI_PreferencesDialog_MigrationHistory | Historical migration record whose "Current Status (Implemented)" / "kept up-to-date" sections are now substantially wrong (schema UI removed, PrefsInput layer deleted, panes renamed, tree grew 9→14). | Keep as history but strip/date-stamp every "current/next-prompt" framing (or move those sections to `Plans/Done/`); replace deleted symbol names with a pointer to `UI_PreferencesDialog.md`. Apply BP-C7 first. |
| **BP-D2** | FileSystem_FtpScpSftpPerformanceComparison | Library-comparison memo; its "Option A quick win" tuning **already shipped** (TCP_NODELAY, keepalive, 512 KB buffers, handle reuse, 4 concurrent — the doc still says "default 1"). | Add a dated "Option A implemented; B/C not pursued" header, or move to a research-notes area. Quality `poor`. |
| **BP-D3** | UI_VisibleTypographyAudit | Dated closeout snapshot; cited evidence archive no longer exists; app-owned Monitor sits outside its scope (BP-B13). | Relabel as a historical snapshot, or replace with a short normative rule ("app-owned visible surfaces route through DxUi.Typography.h; no GDI text/font APIs") + an explicit scope statement. |
| **BP-D4** | UI_VisibleNativeAudit | Dated inventory snapshot (2026-06-08); factual claims all still hold, one broken archive link. | Add a "snapshot as of <date>, not a contract" banner; fix the dead link. No code implied. |
| **BP-D5** | UI_KeyboardManagement | 6-line redirect stub → `UI_CommandMenuKeyboard.md`. Accurate, harmless. | Leave, or delete after updating the 2 remaining historical references. |

---

## Suggested execution order

1. **Track A HIGH (real defects), RED-first:** BP-A1 (cloud profiles unusable) → BP-A2 (plugin
   satellites unreachable) → BP-A4 (pane cookie routing) → BP-A3 (export gate, via RedConfigure
   plan). These are user-facing correctness/security.
2. **Track B HIGH decisions:** BP-B1 (shutdown-abort contract) — unblocks a spec-or-code edit.
3. **Track C HIGH dangerous docs:** BP-C2 (S3 recursive-delete safety claim) → BP-C3 (MSI service)
   → BP-C4 (VFS ABI struct) → BP-C1 → BP-C5/C6 (test inventory). These mislead auditors/plugin
   authors *today*; they're cheap doc edits with high safety payoff.
4. **Track A/B medium**, batched by area (Connection Manager cluster A5/A6/B3/B4; PluginAPI cluster
   A10/A11/B7; S3 cluster A13/B6).
5. **Track C medium/low** spec sweeps, one spec at a time (appendix order).
6. **Track D** relabels (fast, do alongside anything).

## Appendix — per-spec verdict matrix (all 52)

Columns: verdict · quality · findings (H/M/L) · #fix-code · #decide · #fix-spec.

| Spec | Verdict | Quality | H/M/L | code | decide | spec |
|------|---------|---------|-------|------|--------|------|
| Core_CompareDirectories | MINOR DRIFT | good | 0/2/6 | 0 | 1 | 7 |
| Core_CompilerWarnings | MAJOR DRIFT | adequate | 1/1/3 | 2 | 1 | 2 |
| Core_ConnectionManager | MAJOR DRIFT | good | 1/7/2 | 4 | 2 | 4 |
| Core_DirectoryInfoCache | MINOR DRIFT | adequate | 0/2/2 | 0 | 0 | 4 |
| Core_Localization | MINOR DRIFT | good | 1/3/2 | 2 | 2 | 2 |
| Core_RedConfigure | MINOR DRIFT | good | 1/1/2 | 2 | 0 | 2 |
| Core_RedSalamanderMonitor | MINOR DRIFT | adequate | 0/4/4 | 0 | 1 | 7 |
| Core_Search | MINOR DRIFT | good | 0/1/2 | 0 | 1 | 2 |
| Core_SettingsStore | MINOR DRIFT | good | 0/6/2 | 2 | 1 | 5 |
| Core_StartupBootstrap | MINOR DRIFT | adequate | 1/2/1 | 0 | 1 | 3 |
| FileSystem_FileOperations | MINOR DRIFT | good | 0/1/2 | 0 | 0 | 3 |
| FileSystem_FtpScpSftpPerformanceComparison | NOT NORMATIVE | poor | 0/2/2 | 0 | 1 | 3 |
| FileSystem_FtpSftpScp | MINOR DRIFT | adequate | 0/1/2 | 0 | 0 | 3 |
| FileSystem_GoogleDrive | MINOR DRIFT | adequate | 0/4/3 | 0 | 1 | 6 |
| FileSystem_Imap | ACCURATE | good | 0/0/1 | 0 | 0 | 1 |
| FileSystem_MicrosoftDrive | ACCURATE | good | 0/0/2 | 0 | 1 | 1 |
| FileSystem_Mtp | MINOR DRIFT | good | 0/2/2 | 2 | 0 | 2 |
| FileSystem_S3 | MAJOR DRIFT | adequate | 2/3/0 | 1 | 1 | 3 |
| Installer_Msi | MAJOR DRIFT | adequate | 1/0/2 | 1 | 0 | 2 |
| Installer_Msix | MINOR DRIFT | adequate | 0/4/2 | 0 | 0 | 6 |
| Plugins_PluginAPI | MAJOR DRIFT | adequate | 1/5/0 | 3 | 1 | 2 |
| Plugins_ViewerImgRaw | MINOR DRIFT | good | 0/2/5 | 0 | 2 | 5 |
| Plugins_ViewerPE | ACCURATE | adequate | 0/0/2 | 0 | 2 | 0 |
| Plugins_ViewerPlugins | MINOR DRIFT | good | 0/1/4 | 0 | 0 | 5 |
| Plugins_ViewerSpace | MINOR DRIFT | good | 1/0/5 | 2 | 1 | 3 |
| Plugins_ViewerSqlite | MINOR DRIFT | adequate | 0/1/3 | 0 | 1 | 3 |
| Plugins_ViewerText | ACCURATE | good | 0/0/3 | 0 | 0 | 3 |
| Plugins_ViewerWeb | MINOR DRIFT | good | 0/2/1 | 1 | 0 | 2 |
| Plugins_VirtualFileSystem | MINOR DRIFT | good | 1/3/3 | 0 | 0 | 7 |
| Testing_PerformanceValidation | ACCURATE | good | 0/0/2 | 1 | 0 | 1 |
| Testing_SelfTestRemoteCredentials | MINOR DRIFT | good | 0/3/3 | 0 | 0 | 6 |
| Testing_SelfTests | MINOR DRIFT | good | 0/1/2 | 1 | 0 | 2 |
| Testing_TestCoverage | MAJOR DRIFT | adequate | 3/3/1 | 0 | 0 | 7 |
| UI_BatchRenameWindow | MINOR DRIFT | good | 0/2/2 | 0 | 1 | 3 |
| UI_CommandMenuKeyboard | MINOR DRIFT | good | 0/2/3 | 0 | 0 | 5 |
| UI_DxUiSharedGrid | MINOR DRIFT | adequate | 0/4/1 | 0 | 0 | 5 |
| UI_DxUiWinUIDesign | MINOR DRIFT | good | 0/2/2 | 0 | 1 | 3 |
| UI_FindFilesWindow | ACCURATE | good | 0/0/1 | 0 | 0 | 1 |
| UI_FolderView | MINOR DRIFT | good | 0/0/3 | 0 | 0 | 3 |
| UI_FolderWindow | MINOR DRIFT | good | 0/2/1 | 0 | 0 | 3 |
| UI_KeyboardManagement | NOT NORMATIVE | adequate | 0/0/0 | 0 | 0 | 0 |
| UI_ManagePluginsDialog | ACCURATE | good | 0/0/0 | 0 | 0 | 0 |
| UI_NavigationView | MINOR DRIFT | good | 0/3/3 | 0 | 0 | 6 |
| UI_PreferencesDialog | MINOR DRIFT | good | 0/1/1 | 0 | 0 | 2 |
| UI_PreferencesDialog_MigrationHistory | NOT NORMATIVE | poor | 1/7/3 | 0 | 0 | 11 |
| UI_RedConfigure | MINOR DRIFT | good | 0/1/0 | 0 | 1 | 0 |
| UI_TopLevelToolWindows | MINOR DRIFT | good | 0/4/0 | 2 | 2 | 0 |
| UI_VisibleComctlAudit | MINOR DRIFT | good | 0/0/2 | 0 | 0 | 2 |
| UI_VisibleNativeAudit | NOT NORMATIVE | good | 0/0/1 | 0 | 0 | 1 |
| UI_VisibleTypographyAudit | NOT NORMATIVE | adequate | 0/3/2 | 0 | 1 | 4 |
| UI_VisualStyle | MINOR DRIFT | good | 1/0/3 | 0 | 1 | 3 |
| WingetIntegration | MINOR DRIFT | good | 0/1/1 | 1 | 0 | 1 |

*Full per-finding detail (spec claim, code reality, file:line evidence, verification verdict) is in
the audit digest retained at the session scratchpad `digest.md` / `actionable.txt`. Regenerate from
`final-all.json` if needed.*

---
_Audit performed 2026-07-02/03 via a 145-agent + 50-agent multi-pass workflow (one deep reader per
spec → adversarial verifier per HIGH finding), max effort. HEAD `275c04034`, dirty working tree._

## Closeout review — 2026-07-06

Re-audited this ledger against current HEAD `205736167523eaf11600d8f72d322962ec2e2ebf` and the
working tree on 2026-07-06. The original supporting scratchpad artifacts named by this plan
(`scratchpad/digest.md`, `scratchpad/actionable.txt`, `final-all.json`) are not present in the repo,
so closeout used the explicit rows in this file plus the current source/spec tree. Remaining
low-priority rows that exist only in the missing scratchpad are not independently actionable from
this plan and should be rediscovered by a fresh targeted audit if still important.

Resolved in this closeout:

- BP-A1, BP-A2, BP-A4, BP-A6, BP-A8, BP-A9, BP-A10, BP-A11, BP-A12, and BP-A13 were still true and
  were fixed in code/specs with focused regression coverage or source guards where practical.
- BP-A7 was already fixed before this closeout; the current settings loader preserves
  `connections.allowInsecureTlsInAutomation`.
- BP-A5 was reclassified as stale/superseded: `Core_ConnectionManager.md` and the
  `windows_hello_cache` selftest intentionally document the current reuse model for successful
  interactive auth in long-running background operations.
- BP-A3 was completed by `Specs/Plans/Done/RedConfigure_LocalizationThemeManagerPlan.md` on 2026-07-14.
- BP-B1 was decided in favor of the shipped best-effort shutdown behavior; `Core_StartupBootstrap.md`
  now says close auxiliary top-level windows, warn on a stuck window, and proceed through the
  deferred final-close message.
- BP-C1, BP-C2, BP-C3, BP-C4, BP-C5, BP-C6, BP-C7, BP-C8, BP-C9, and BP-C10 were still true enough
  to act on and were corrected in the authoritative specs/tooling docs.
- BP-D1's false schema-driven Preferences UI status was corrected in
  `UI_PreferencesDialog_MigrationHistory.md`. The remaining Track D entries are archival/scope
  classification work, not blockers for retiring this stale audit ledger.

This plan is retired as an active WIP owner. Current durable contracts now live in the updated
domain specs and tests; remaining product-decision rows (BP-B2..BP-B13 and the low decide batch)
should be handled only if a current owner plan revalidates them against the then-current tree.
