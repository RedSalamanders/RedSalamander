# Shared Helper Reuse Contract

## Status and scope

This is the authoritative repository contract for reusable implementation primitives. It applies to production
code, plugins, RedConfigure, first-party test executables, and in-product selftests. Domain specs remain
authoritative for the behavior of a particular feature; this document identifies the canonical implementation
owner and prevents a consumer from silently recreating an already shared policy.

## Mandatory reuse rule

Before adding a file-local or component-local helper:

1. Search this catalog, `Common/`, and, for test infrastructure, `Tests/TestSupport/`.
2. Reuse the canonical API when its semantics match. A rename, forwarding wrapper, or minor syntactic rewrite
   is still a duplicate and must not be added.
3. If a generally reusable capability is missing at the same dependency layer, extend the canonical helper and
   its focused tests instead of copying it into a consumer.
4. Keep a local variant only when it deliberately differs in policy, dependency layer, ABI, ownership, error
   mapping, or measured hot-path constraints. Its name and definition comment must state the difference, and
   the variant must appear in the relevant source-contract allowlist or focused test.
5. When a new canonical helper is introduced, update this catalog and the authoritative domain spec that owns
   the behavior. Add or extend a source-contract test when textual or structural regression could recreate the
   removed copies.

Similar names do not by themselves establish semantic equivalence. In particular, strict and replacement UTF
conversion, fully absolute and merely drive-qualified paths, JSON coercion policies, stable hash algorithms,
and product-specific color contrast thresholds remain distinct until their contracts are proven identical.

## Canonical implementation catalog

| Concern | Canonical owner | Required reuse boundary |
|---|---|---|
| ARGB to Direct2D color | `RedSalamander::DxUi::ColorFromArgb` in `Common/DxUi/DxUi.h` | All code already depending on DxUi; do not add another `ColorF`/ARGB unpacker there. |
| `ARGB`/`COLORREF`, compositing, and luminance math | `ColorRefFromArgb` and `Common::Colors` in `Common/Helpers.h` | Reuse the numeric conversion, `CompositeArgbOverOpaqueBackground`, and luminance math; keep product contrast thresholds and translucent-backdrop policy explicit at the caller. |
| Application and folder DxUi palettes | `MakeAppThemeDxPalette` / `MakeFolderContentDxPalette` in `RedSalamander/DxUiThemePalette.h` | RedSalamander application surfaces; viewer-only themes use their viewer contract instead. |
| DPI, popup geometry, and owner centering | `Common/WindowSizing.h` | Reuse conversions, bounds, popup placement, minimum-track sizing, and centering at Win32 consumers. |
| UTF conversion | `Common/StringConversion.h` | Choose the named strict or replacement API; never hide the malformed-input policy in a local converter. |
| yyjson reads, ownership, and object rewrites | `Common/YyjsonHelpers.h` | Reuse document owners and policy-parameterized accessors. Use `ParseObjectDocument` for transactional object-root parsing and `WriteObjectWithoutMembers` for migration/scrubbing rewrites; schema-specific missing/null/coercion policy stays explicit. |
| Windows path classification, normalized local containment, and unique siblings | `Common/PathUtils.h` | Select the exact drive-qualified/absolute/UNC/device predicate, normalized local Windows path comparison/containment helper, and shared unique-file primitive. Callers must still choose lexical versus physical normalization explicitly. |
| URI percent encoding | `Common/UriEncoding.h` | Select an explicit slash policy; no local byte encoder. |
| Bounded handle I/O | `Common/HandleIo.h` | Reuse exact read, total write, rewind, and bounded-size helpers, including zero-progress failure behavior. |
| Cloud continuation paging | `Common::Paging::ContinuationGuard` in `Common/PaginationGuard.h` | Cloud/provider pagination must reuse the page/item/byte/deadline/cancellation and repeated/empty-continuation guard. Provider request, token parsing, item semantics, and diagnostics remain local. |
| Local regular-file output transactions | `Common::Files::LocalFileTransaction` in `Common/LocalFileTransaction.h` | Local writers that need unique sibling creation, complete/zero-progress-safe writes, flush, optional size verification, same-volume fail-if-exists or replace promotion, and optional exact finalized-file identity returned from `Commit`. Provider-backed writers retain their provider transaction contract. |
| Connection profile identity | `Common::Settings::CreateConnectionProfileId`, `NormalizeConnectionProfileId`, and `ValidateConnectionProfileIds` in `Common/SettingsStore.h` | Generate canonical lowercase profile GUIDs and enforce reserved/duplicate identity rules before Settings Store load, UI commit, save, WinCred target construction, or authorization-cache use. Do not add another local GUID formatter for connection profiles. |
| Throughput parsing/formatting | `Common/ThroughputParsing.h` | Reuse the shared binary-throughput grammar; broader UI option parsing remains local. |
| File metadata display | `Common/FileMetadataFormatting.h` | Reuse normalized fields and an explicit display profile. |
| Unicode clipboard writes | `Common::Clipboard::TrySetUnicodeText` in `Common/UnicodeClipboard.h` | Reuse the direct WIL-owned Unicode allocation/publication path and explicit empty-text policy. The payload is prepared before `EmptyClipboard` and released only after successful `SetClipboardData`; do not add leaf forwarding wrappers for individual Win32 calls. |
| Viewer file-combo hosting | `Common/ViewerFileComboHost.h` | Reuse the proven popup-opening, keyboard, subclass, and height behavior in compatible viewers. |
| Viewer title-bar theming | `Common/ViewerTitleBarTheme.h` | Reuse viewer DWM attribute and accent resolution; app-level tool windows follow their app theme contract. |
| App modal DxUi shell | `Common/ModalWindowShell.h` | Reuse top-level modal lifetime, nested-loop, owner restoration, and teardown behavior; content remains local. |
| HWND Direct2D target lifecycle | `Common/HwndRenderTargetResources.h` | Reuse the factory/target/brush lifecycle where the exact FunctionBar/StatusBar contract applies. |
| Callback registration and module pins | `RegistrationCallbackState<T>` and `TransferModulePinToCallbackReturn` in `Common/Helpers.h` | Reuse callback generation/drain and callback-return module-pin transfer; cancellation/session policy remains explicit. |
| Packed plugin `FileInfo` buffers | `Common::Plugins::PackedFileInfoBuffer` in `Common/PackedFileInfoBuffer.h` | Buffered plugin facades use the checked/aligned owner; streaming and purpose-built fixture models remain explicit exceptions. |
| Plugin configuration schema/codec | `Common/PluginConfiguration.h` and `Common/Common/PluginConfiguration.cpp` | Reuse the common schema model, validation, parse, and write behavior; provider policy is expressed through the model. |
| Process-wide libcurl lifetime | `Common::CurlRuntime::ProcessLease` in `Common/CurlProcessRuntime.h` | Every independently unloadable DLL that uses libcurl owns one lease. Initialize through the lease, reach the plugin quiet point before releasing it, and let only the final process participant call `curl_global_cleanup`; never add a DLL-local global cleanup. |
| Posted payload ownership/coalescing | `PostMessagePayload`, `TakeMessagePayload`, `InitPostedPayloadWindow`, `DrainPostedPayloadsForWindow`, and `TakeAndCoalesceContiguousPostedPayloads` in `Common/Helpers.h` | A nonzero `lParam` is a process-unique opaque registry token, never a payload pointer. Receivers must use `TakeMessagePayload<T>` and ignore a null result from a nonzero stale/drained/type-mismatched token before message-specific work. Zero is reserved for an explicitly defined payload-less fallback. Registered windows initialize on create and drain on `WM_NCDESTROY`; the drain fences new posts, invalidates tokens, and deletes each registered payload exactly once. Stale queued tokens may survive numeric HWND reuse but cannot reacquire storage. Reuse queue-head-safe keyed contiguous draining; message-specific merge semantics stay in the caller. |
| RedConfigure binary reads | `RedConfigure/RedConfigureBinaryFile.h` | All RedConfigure bounded binary-file consumers; it is intentionally RedConfigure-local because of the dependency boundary. |
| Test sandbox and environment scopes | `Tests/TestSupport/TestSupport.h` and the native selftest `AcquireTestSandbox` APIs | First-party tests must use the shared run/scratch/artifact layout rather than process-temp or ad hoc roots. |
| Test message/snapshot polling | `PumpMessagesUntil` and `WaitForSnapshot` in `Tests/TestSupport/TestSupport.h` | Reuse bounded polling and diagnostics; fixed sleeps are not a replacement for observable readiness. |
| Contained child processes | `RunChildProcess` in `Tests/TestSupport/ChildProcess.h` | Reuse structured arguments, concurrent bounded output drains, timeout handling, and kill-on-close Job containment. |
| Commands directed desktop-input warning | `DirectedSelfTestInputWarning` in `RedSalamander/SelfTest/Commands/Commands.SelfTest.cpp` | Every Commands test fragment that must move the real cursor or send desktop-global input; do not add test-family-local warning windows. Every guarded path must also restore cursor/focus state on all exits. |

## Enforcement

The domain specs cited by each consumer define behavior. `Tools/Tests/*SourceContracts.Tests.ps1` guards the
high-risk migrations, including render-target ownership, modal windows, plugin lifetime/configuration, packed
buffers, posted payload coalescing, and test-support adoption. A source guard complements behavioral tests; it
must not replace them or force semantically different policies into one helper.
