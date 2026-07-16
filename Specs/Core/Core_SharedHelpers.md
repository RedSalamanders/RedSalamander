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
| `ARGB`/`COLORREF` and luminance math | `ColorRefFromArgb` and `Common::Colors` in `Common/Helpers.h` | Reuse the numeric conversion/luminance math; keep product contrast thresholds explicit at the caller. |
| Application and folder DxUi palettes | `MakeAppThemeDxPalette` / `MakeFolderContentDxPalette` in `RedSalamander/DxUiThemePalette.h` | RedSalamander application surfaces; viewer-only themes use their viewer contract instead. |
| DPI, popup geometry, and owner centering | `Common/WindowSizing.h` | Reuse conversions, bounds, popup placement, minimum-track sizing, and centering at Win32 consumers. |
| UTF conversion | `Common/StringConversion.h` | Choose the named strict or replacement API; never hide the malformed-input policy in a local converter. |
| yyjson reads and ownership | `Common/YyjsonHelpers.h` | Reuse document owners and policy-parameterized accessors; schema-specific missing/null/coercion policy stays explicit. |
| Windows path classification and unique siblings | `Common/PathUtils.h` | Select the exact drive-qualified/absolute/UNC/device predicate and shared unique-file primitive. |
| URI percent encoding | `Common/UriEncoding.h` | Select an explicit slash policy; no local byte encoder. |
| Bounded handle I/O | `Common/HandleIo.h` | Reuse exact read, total write, rewind, and bounded-size helpers, including zero-progress failure behavior. |
| Throughput parsing/formatting | `Common/ThroughputParsing.h` | Reuse the shared binary-throughput grammar; broader UI option parsing remains local. |
| File metadata display | `Common/FileMetadataFormatting.h` | Reuse normalized fields and an explicit display profile. |
| Unicode clipboard writes | `Common/UnicodeClipboard.h` | Reuse Unicode ownership/allocation/empty-text policy rather than duplicating Win32 clipboard cleanup. |
| Viewer file-combo hosting | `Common/ViewerFileComboHost.h` | Reuse the proven popup-opening, keyboard, subclass, and height behavior in compatible viewers. |
| Viewer title-bar theming | `Common/ViewerTitleBarTheme.h` | Reuse viewer DWM attribute and accent resolution; app-level tool windows follow their app theme contract. |
| App modal DxUi shell | `Common/ModalWindowShell.h` | Reuse top-level modal lifetime, nested-loop, owner restoration, and teardown behavior; content remains local. |
| HWND Direct2D target lifecycle | `Common/HwndRenderTargetResources.h` | Reuse the factory/target/brush lifecycle where the exact FunctionBar/StatusBar contract applies. |
| Callback registration and module pins | `RegistrationCallbackState<T>` and `TransferModulePinToCallbackReturn` in `Common/Helpers.h` | Reuse callback generation/drain and callback-return module-pin transfer; cancellation/session policy remains explicit. |
| Packed plugin `FileInfo` buffers | `Common::Plugins::PackedFileInfoBuffer` in `Common/PackedFileInfoBuffer.h` | Buffered plugin facades use the checked/aligned owner; streaming and purpose-built fixture models remain explicit exceptions. |
| Plugin configuration schema/codec | `Common/PluginConfiguration.h` and `Common/Common/PluginConfiguration.cpp` | Reuse the common schema model, validation, parse, and write behavior; provider policy is expressed through the model. |
| Posted payload ownership/coalescing | `PostMessagePayload`, `TakeMessagePayload`, and `TakeAndCoalesceContiguousPostedPayloads` in `Common/Helpers.h` | Reuse queue-head-safe ownership and keyed contiguous draining; message-specific merge semantics stay in the caller. |
| RedConfigure binary reads | `RedConfigure/RedConfigureBinaryFile.h` | All RedConfigure bounded binary-file consumers; it is intentionally RedConfigure-local because of the dependency boundary. |
| Test sandbox and environment scopes | `Tests/TestSupport/TestSupport.h` and the native selftest `AcquireTestSandbox` APIs | First-party tests must use the shared run/scratch/artifact layout rather than process-temp or ad hoc roots. |
| Test message/snapshot polling | `PumpMessagesUntil` and `WaitForSnapshot` in `Tests/TestSupport/TestSupport.h` | Reuse bounded polling and diagnostics; fixed sleeps are not a replacement for observable readiness. |
| Contained child processes | `RunChildProcess` in `Tests/TestSupport/ChildProcess.h` | Reuse structured arguments, concurrent bounded output drains, timeout handling, and kill-on-close Job containment. |

## Enforcement

The domain specs cited by each consumer define behavior. `Tools/Tests/*SourceContracts.Tests.ps1` guards the
high-risk migrations, including render-target ownership, modal windows, plugin lifetime/configuration, packed
buffers, posted payload coalescing, and test-support adoption. A source guard complements behavioral tests; it
must not replace them or force semantically different policies into one helper.
