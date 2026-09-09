# Plugin Host Services Contract

This document describes **host-provided services** that plugins can use at runtime.
It complements (and does not replace) the existing per-plugin-type specs:

- Viewer plugins: `Specs/Plugins/Plugins_ViewerPlugins.md`
- Virtual file system plugins: `Specs/Plugins/Plugins_VirtualFileSystem.md`

## Purpose

Plugins surface **user-visible alerts** (error/warning/info/busy) through the host UI because:

- plugins should not need to implement Direct2D/DirectWrite rendering for simple alerts, and
- the host must control alert presentation for consistency (theme, DPI, modality, layout), and
- plugins must be able to choose the **scope** (pane vs app) and **modality** (blocks input vs not).

The host uses `RedSalamander::Ui::AlertOverlay` for pane-content alerts and
host-owned overlay windows for pane, application, and explicit-window scopes.
This spec defines the behavioral contract for those services. Public interface
layouts and IIDs are owned by `Common/PlugInterfaces/Host.h`; the required
factory entry point is owned by `Common/PlugInterfaces/Factory.h`.

## Alert Scopes and Modality

The alert system supports the following scopes (surfaces), each with a **modal** or **modeless** interaction policy:

### Scopes

- **Pane content**: pane **without** the navigation bar (typically the `FolderView` client area).
- **Pane**: pane **with** the navigation bar (navigation + content + status bar).
- **Application**: the main application window (covers both panes).

### Modality

- **Modal**: blocks input to the chosen scope while the alert is visible.
- **Modeless**: does not block input to the chosen scope while the alert is visible.

Important (alerts): “modal” here means **input-blocking overlay**, not a synchronous/nested-message-loop dialog.
Calls into the alert API are **non-blocking** and return immediately.

## Modal Prompts (Confirmations With Result)

Some scenarios require a **modal confirmation prompt** (Yes/No, OK/Cancel, Ok, Cancel) where the caller must receive a result (MessageBox-like).

- Prompts are still presented as an overlay (not a Win32 MessageBox), but the API call is **blocking** until the user chooses a button.
- Prompts always block input to their chosen scope (pane content / pane / application).

## Current implementation

- `RedSalamander/Ui/AlertOverlay.h`: shared D2D/DWrite alert renderer (scrim + card + optional close + optional buttons).
- `RedSalamander/FolderView.*`: pane-content alert integration used for enumeration, rendering, operation errors, and busy state.
- `RedSalamander/HostServices.cpp`: `IHost`, `IHostAlerts`, and `IHostPrompts` implementation, UI-thread marshaling, scope routing, window overlays, and prompt completion.
- `RedSalamander/FolderWindow.cpp`: host-service message dispatch and pane ownership.

Viewer plugins SHOULD surface non-fatal alerts via `IHostAlerts` (host-rendered) rather than implementing their own alert rendering.

## Host UI services

The host exposes **UI services** to plugins (alerts + prompts). Plugins call into these services to show/clear alerts and to run modal prompts; the host owns all rendering and input management.

### How plugins obtain host services

Host services are provided to plugins through the single current factory entry
point. The source-tree host and shipped plugins use the same ABI revision.

`IHost` is a COM object that supports `QueryInterface` for specific service interfaces (alerts, prompts, future services).

### ABI evolution policy

This is the governing rule for future host/plugin interface changes.

The currently supported binary boundary is one source-tree/release generation:
the host and every shipped plugin are built and packaged together from the same
revision. Mixing a plugin binary from an earlier or later release with the
current host is unsupported, including binaries from before the August 2026
leading-`sizeBytes` transitions. `sizeBytes` validates records inside the
supported generation; it is not an implicit cross-release compatibility claim.
For every current host, viewer, and Terminal method, the required prefix is the
full current record unless that method's authoritative specification names a
smaller reviewed prefix.

- Same-GUID COM interfaces have fixed vtable slot order. Do not insert, remove, reorder, or change existing virtual methods on an interface that keeps the same IID.
- The already-shipped source-tree-only `ITerminal::SetCallback` insertion is the sole reviewed exception: it retains the current IID only under the closed lockstep boundary documented by `Specs/Terminal/Terminal_EmbeddedPlugin.md`. Atomic central staging and complete receipt closure of the host, Terminal plugin, private runtime, and satellites are mandatory. This waiver supplies no mixed-generation compatibility and authorizes no further same-IID edit.
- Append-only same-GUID changes are allowed only when old-plugin/new-host compatibility is explicitly covered by a slot-compatibility test fixture.
- If that compatibility is not proven, create a new IID/name such as `IHostConnections2` and expose it through `IHost::QueryInterface`.
- New-plugin/old-host compatibility must probe the newer IID with `QueryInterface`; plugins must not assume appended methods exist on an older host implementation.
- Callback shutdown/order validation and an old-plugin/new-host compatibility fixture are required before treating an interface evolution as complete.
- Extensible in-process request/result records use one leading `uint32_t sizeBytes`; they do not duplicate that information with `version`, `apiVersion`, or `structSize` fields. The caller reports the bytes it actually provides. A consumer rejects a value shorter than the prefix it needs, accepts a larger value, and ignores unknown tail bytes. Compatible growth is optional and append-only; incompatible ownership or semantic changes require a new method or COM IID. Serialized files, settings, clipboard formats, and evidence retain their independent schema/format versions.
- Runtime boundary tests, not frozen `sizeof` assertions, prove rejection of undersized records and acceptance of current/oversized records. Source-tree host, viewer, and terminal plugins are changed together when the current required prefix changes.
- A future release that supports mixed-generation or third-party binary plugins must first add an explicit, testable generation negotiation/rejection boundary before interface invocation. Supporting an older generation also requires new IIDs/adapters where layouts or semantics changed, frozen old-header fixtures in both host/plugin directions, and x64/ARM64 prefix/layout coverage. Until that stronger contract is approved and implemented, no loader or retained IID may infer an older layout from pointer, HWND, version, or shifted field bytes.

### File-system typed route ABI

`IFileSystemRouteCapabilities` (`{1e924d87-2e62-4ab4-9f37-c565d465f25e}`)
is a separate interface queried from `IFileSystem`. It is the sole executable
route-capability authority for the current source-tree/release generation.
`IFileSystem` and `IFileSystemPathCapabilities2` keep their existing IIDs and
vtable layouts; capability JSON is diagnostics/extensions only and cannot enable,
repair, or contradict an executable typed decision.

The route records follow the ABI evolution policy above: one leading
`uint32_t sizeBytes`, exact current-size validation by producers, and
current-size-or-larger prefix validation by consumers. Producers must reject a
record from a different generation, must return only strict `BOOL` and valid enum
values, and must not retain any caller record, input string, or arena. Variable
UTF-16 results use caller-owned `FileSystemArena` storage. A caller may retry once
when `requiredArenaBytes` exceeds the normal 4 KiB arena, but must reject requests
above 64 KiB, arena overruns, pointers outside the used arena, missing NUL
termination, or a required/used-byte mismatch.

The supported generation requires the complete typed interface. Missing QI or
malformed output is a provider contract violation and fails closed. A future
mixed-generation plugin boundary must add explicit negotiation and adapters; the
record prefix rule alone does not create that compatibility.

Current factory signature:

```cpp
extern "C"
{
    PLUGFACTORY_API HRESULT __stdcall RedSalamanderCreate(
        REFIID riid,
        const FactoryOptions* factoryOptions,
        IHost* host,
        const wchar_t* pluginId,
        void** result
    );
}
```

### Common factory module lifetime exports

`Common/PlugInterfaces/Factory.h` also declares optional module quiet-point exports:

```cpp
extern "C"
{
    PLUGFACTORY_API void __stdcall RedSalamanderPluginShutdown() noexcept;
    PLUGFACTORY_API BOOL __stdcall RedSalamanderPluginCanUnloadNow() noexcept;
    PLUGFACTORY_API BOOL __stdcall RedSalamanderPluginRetainModuleUntilProcessExit() noexcept;
}
```

These exports are not a default plugin requirement. A plugin SHOULD omit them unless it owns DLL-global worker threads, schedulers, caches, registered window classes, graphics resources, or driver-backed resources whose lifetime cannot be fully represented by ordinary COM instance shutdown.

`RedSalamanderPluginShutdown()` must be idempotent and non-throwing. After it returns, no DLL-global worker may call host callbacks or touch state that the host can release before `FreeLibrary`. A host manager MUST still release or close live plugin instances through their normal interface contract before unloading a module; the export is a final module quiet point, not a replacement for `Close()`, `SetCallback(nullptr, nullptr)`, cancellation, or COM `Release()`.

`RedSalamanderPluginCanUnloadNow()` is optional and is queried after `RedSalamanderPluginShutdown()` but before runtime resource-owner unregister and `FreeLibrary`. Omitted export means `TRUE`. Returning `FALSE` means DLL-global work is still unwinding and explicit runtime unload is unsafe; the host MUST keep the DLL mapped, keep the resource owner registered, mark the plugin entry as unload-deferred, skip same-path reload during rediscovery, and retry automatically. Recovery MUST NOT depend only on a user-initiated Refresh/Apply path: a host manager that receives `ERROR_BUSY` from a configuration/schema query because an unload-deferred entry still owns the plugin path MUST trigger an unload-deferred sweep on that busy result or schedule a bounded retry timer until the entry unloads or the manager shuts down. Plugins that return `FALSE` MUST do so from a non-blocking state check and MUST eventually return `TRUE` after the outstanding work exits.

`RedSalamanderPluginRetainModuleUntilProcessExit()` is only meaningful during process shutdown. Returning `TRUE` asks the host to run the quiet point but leave the DLL mapped until OS process teardown. This is reserved for modules where explicit process-shutdown `FreeLibrary` is known to race with driver or library teardown after the plugin has already gone quiet.

Host plugin managers that load DLLs through `RedSalamanderCreate` MUST centralize unload through a helper that releases normal COM instances first, calls `RedSalamanderPluginShutdown()` when present, honors runtime `RedSalamanderPluginCanUnloadNow()` deferral before unregistering resources, unregisters resource owners only when unload is allowed, and honors `RedSalamanderPluginRetainModuleUntilProcessExit()` only for process-shutdown unload.

Lifetime/ownership:
- `host` is caller-owned and remains valid for the lifetime of the plugin instance created from this call.
- `pluginId` identifies the logical plugin being created; single-plugin DLLs MAY accept `nullptr` or empty.
- Plugins MAY `AddRef()` `host` (or any queried service interface) if they need to store it beyond the factory call.

### Context routing (cookie)

Pane-scoped alerts must be routed to the correct pane instance. Plugins do this by passing back the **opaque cookie** provided by the host in the relevant call/registration:

- For per-call callbacks (`IFileSystem*` operations, search, etc.), use the `cookie` passed alongside the callback.
- For registration-style callbacks (`INavigationMenu::SetCallback`), use the `cookie` passed at registration time.

Plugins MUST NOT invent cookies. If a cookie is missing (`nullptr`) the plugin can only reliably request **application-scoped** alerts.

### Public ABI declarations

`Common/PlugInterfaces/Host.h` is the single source of truth for `IHost`,
`IHostAlerts`, `IHostPrompts`, their IIDs, request layouts, enum values, and
vtable order. Request records use one leading `sizeBytes` field as required by
the ABI evolution policy above. This document intentionally does not duplicate
those declarations.

### Behavioral contract

- **Threading**: plugins MAY call host UI services from any thread; the host MUST marshal to the UI thread internally.
- **Synchronous marshaling**: blocking, result-bearing UI calls use a dedicated non-owning token registry around `SendMessageW`. The caller thread retains the request/result storage for the full synchronous dispatch; the receiving window validates the opaque nonzero token against both message and payload kind. These tokens are never interpreted as pointers and are never passed to the owning `PostMessagePayload` registry.
- **Lifetime**: the host MUST copy strings before returning; plugins may free inputs immediately after the call.
- **Localization**:
  - host-owned alerts must use `.rc` resources (per project guidelines),
  - plugin-provided strings must be localized by the plugin (the host treats them as already-localized).
- **Alerts**:
  - replacement policy: each `(scope, cookie)` pair has at most **one** active alert; `ShowAlert` replaces it.
  - pane routing:
    - `HOST_ALERT_SCOPE_PANE_CONTENT` and `HOST_ALERT_SCOPE_PANE` callers SHOULD pass
      `reinterpret_cast<void*>(HOST_PANE_COOKIE_LEFT)` or `reinterpret_cast<void*>(HOST_PANE_COOKIE_RIGHT)`,
    - `nullptr` is accepted only for legacy focused-pane routing,
    - unknown non-null pane cookies are treated as legacy focused-pane routing for compatibility.
  - dismissal:
    - if `closable == TRUE`, the host provides a close “X” and Esc dismiss behavior (within that scope),
    - if `closable == FALSE`, dismissal is programmatic (host clears on navigation/path change where appropriate, or plugin calls `ClearAlert`).
- **Prompts**:
  - `ShowPrompt` is blocking until the user responds.
  - pane-scoped prompts follow the same pane-cookie routing rules as alerts.
  - button labels are host-localized (standard button sets only); plugins do not provide button text.
  - `HostPromptRequest::presentation` is optional. `HOST_PROMPT_PRESENTATION_DEFAULT` preserves the severity icon/palette and standard button labels. The Copy, Move, and Delete presentations select distinct operation icons and theme-derived palettes: Copy renders CopyTo, Move renders MoveToFolder, and Delete renders Delete through Segoe Fluent Icons (with shared Unicode fallbacks) rather than hand-drawn geometry. Alert-overlay close actions likewise render the Segoe Fluent Clear glyph with a Unicode multiplication-sign fallback. For `OK`/`OK_CANCEL` button sets, the host also localizes the affirmative label as `Copy`, `Move`, or `Delete` and supplies that operation as the title when the caller omits one.
  - Unknown presentation values fall back to `HOST_PROMPT_PRESENTATION_DEFAULT`.
  - Copy/Move confirmation may attach caller-owned `HostFileOperationPromptOptions` through
    `HostPromptRequest::fileOperationOptions`. The call is synchronous, including host UI-thread
    marshaling, so the pointer stays valid for the complete call. The host validates the exact
    structure version and enum/boolean values, displays the finite Links/Verify/Start/Bandwidth
    choices, and writes values back only after OK/Yes. Cancel leaves the structure unchanged. The
    host and all in-tree plugins compile against this breaking request layout together; there is no
    legacy-size adapter.
  - closing the prompt via window close (if allowed by the host) returns:
    - `HOST_PROMPT_RESULT_OK` for `HOST_PROMPT_BUTTONS_OK`
    - `HOST_PROMPT_RESULT_CANCEL` for all other button sets
- **Connection Manager results**:
  - the current host requires the full current `HostConnectionManagerRequest` and `HostConnectionManagerResult` records; any earlier supported binary prefix remains an ABI-generation decision rather than an inferred compatibility promise,
  - result capacity is validated before the first result write; an undersized result returns `E_INVALIDARG` with the caller's bytes unchanged,
  - exact-size and oversized results are accepted, the produced prefix reports the current result size, and unknown oversized tail bytes remain untouched.
