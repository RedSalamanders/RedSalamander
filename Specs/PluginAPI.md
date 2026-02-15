# Plugin API (Host Services + Alert UI)

This document describes **host-provided services** that plugins can use at runtime.
It complements (and does not replace) the existing per-plugin-type specs:

- Viewer plugins: `Specs/PluginsViewer.md`
- Virtual file system plugins: `Specs/PluginsVirtualFileSystem.md`

## Problem Statement

Plugins sometimes need to surface **user-visible alerts** (error/warning/info/busy) in the host UI, but:

- plugins should not need to implement Direct2D/DirectWrite rendering for simple alerts, and
- the host must control alert presentation for consistency (theme, DPI, modality, layout), and
- plugins must be able to choose the **scope** (pane vs app) and **modality** (blocks input vs not).

The host already has an in-pane alert renderer (`RedSalamander::Ui::AlertOverlay`) and a FolderView integration.
This spec defines a host-facing API so plugins can request alerts using the same alert UI system.

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

## Existing Implementation (Current)

- `RedSalamander/Ui/AlertOverlay.h`: shared D2D/DWrite alert renderer (scrim + card + optional close + optional buttons).
- `RedSalamander/FolderView.*`: pane-content alert overlay (used for enumeration/rendering/operation errors and busy state).
  - Today this covers **pane content only** (not NavigationView / status bar).
  - It already supports an internal `blocksInput` behavior (modal vs modeless at the FolderView level).

Viewer plugins SHOULD surface non-fatal alerts via `IHostAlerts` (host-rendered) rather than implementing their own alert rendering.

## New Host UI Services (Planned ABI)

The host exposes **UI services** to plugins (alerts + prompts). Plugins call into these services to show/clear alerts and to run modal prompts; the host owns all rendering and input management.

### How plugins obtain host services

Host services are provided to plugins via a **single factory entry point**. This is a breaking ABI change: the host and **all plugins** must be updated together.

`IHost` is a COM object that supports `QueryInterface` for specific service interfaces (alerts, prompts, future services).

Proposed factory signature:

```cpp
extern "C"
{
    PLUGFACTORY_API HRESULT __stdcall RedSalamanderCreate(
        REFIID riid,
        const FactoryOptions* factoryOptions,
        IHost* host,
        void** result
    );
}
```

Lifetime/ownership:
- `host` is caller-owned and remains valid for the lifetime of the plugin instance created from this call.
- Plugins MAY `AddRef()` `host` (or any queried service interface) if they need to store it beyond the factory call.

### Context routing (cookie)

Pane-scoped alerts must be routed to the correct pane instance. Plugins do this by passing back the **opaque cookie** provided by the host in the relevant call/registration:

- For per-call callbacks (`IFileSystem*` operations, search, etc.), use the `cookie` passed alongside the callback.
- For registration-style callbacks (`INavigationMenu::SetCallback`), use the `cookie` passed at registration time.

Plugins MUST NOT invent cookies. If a cookie is missing (`nullptr`) the plugin can only reliably request **application-scoped** alerts.

### Reference Types (C++ declarations)

The public ABI should live under `Common/PlugInterfaces/` (file name TBD, e.g. `Host.h`).

```cpp
// Root host services object (extensible via QueryInterface).
// UUID: {c7191bad-276e-4f7b-91ec-4803315413a7}
interface __declspec(uuid("c7191bad-276e-4f7b-91ec-4803315413a7")) __declspec(novtable) IHost : public IUnknown
{
    // No methods yet.
    // QueryInterface for service interfaces such as IHostAlerts and IHostPrompts.
};

enum HostAlertScope : uint32_t
{
    HOST_ALERT_SCOPE_PANE_CONTENT = 1, // pane without navigation bar
    HOST_ALERT_SCOPE_PANE         = 2, // pane with navigation bar
    HOST_ALERT_SCOPE_APPLICATION  = 3, // application window
    HOST_ALERT_SCOPE_WINDOW       = 4, // specific HWND (request.targetWindow)
};

enum HostAlertModality : uint32_t
{
    HOST_ALERT_MODELESS       = 1,
    HOST_ALERT_MODAL          = 2,
};

enum HostAlertSeverity : uint32_t
{
    HOST_ALERT_ERROR          = 1,
    HOST_ALERT_WARNING        = 2,
    HOST_ALERT_INFO           = 3,
    HOST_ALERT_BUSY           = 4,
};

struct HostAlertRequest
{
    // ABI versioning.
    uint32_t version;   // 1
    uint32_t sizeBytes; // sizeof(HostAlertRequest)

    HostAlertScope scope;
    HostAlertModality modality;
    HostAlertSeverity severity;

    // Used only when scope == HOST_ALERT_SCOPE_WINDOW.
    HWND targetWindow;

    // Strings are UTF-16, NUL-terminated, caller-owned, and only valid for the duration of the call.
    const wchar_t* title;   // optional (nullptr/empty allowed)
    const wchar_t* message; // required for user-visible alerts

    // If FALSE, the host does not expose a close “X” and does not dismiss the alert on Esc.
    // Typical use: fatal pane errors that require navigation away, or busy states.
    BOOL closable;

    uint32_t reserved[8];
};

// UUID: {06da0f05-fe31-4273-9029-22037e3b1ea8}
interface __declspec(uuid("06da0f05-fe31-4273-9029-22037e3b1ea8")) __declspec(novtable) IHostAlerts : public IUnknown
{
    // Shows (or replaces) the current alert for the given scope+cookie routing.
    // - For pane scopes, cookie MUST identify the correct pane context.
    // - For application scope, cookie MUST be nullptr.
    virtual HRESULT STDMETHODCALLTYPE ShowAlert(const HostAlertRequest* request, void* cookie) noexcept = 0;

    // Clears the current alert for the given scope+cookie routing.
    // - For HOST_ALERT_SCOPE_WINDOW, cookie MUST be the target HWND (cast to void*).
    virtual HRESULT STDMETHODCALLTYPE ClearAlert(HostAlertScope scope, void* cookie) noexcept = 0;
};

enum HostPromptButtons : uint32_t
{
    HOST_PROMPT_BUTTONS_OK            = 1,
    HOST_PROMPT_BUTTONS_OK_CANCEL     = 2,
    HOST_PROMPT_BUTTONS_YES_NO        = 3,
    HOST_PROMPT_BUTTONS_YES_NO_CANCEL = 4,
};

// Result values intentionally match Win32 MessageBox IDs (IDOK/IDCANCEL/IDYES/IDNO).
enum HostPromptResult : uint32_t
{
    HOST_PROMPT_RESULT_NONE        = 0,
    HOST_PROMPT_RESULT_OK          = 1, // IDOK
    HOST_PROMPT_RESULT_CANCEL      = 2, // IDCANCEL
    HOST_PROMPT_RESULT_YES         = 6, // IDYES
    HOST_PROMPT_RESULT_NO          = 7, // IDNO
};

struct HostPromptRequest
{
    // ABI versioning.
    uint32_t version;   // 1
    uint32_t sizeBytes; // sizeof(HostPromptRequest)

    HostAlertScope scope;
    HostAlertSeverity severity;
    HostPromptButtons buttons;

    // Used only when scope == HOST_ALERT_SCOPE_WINDOW.
    HWND targetWindow;

    // Strings are UTF-16, NUL-terminated, caller-owned, and only valid for the duration of the call.
    const wchar_t* title;   // optional (nullptr/empty allowed)
    const wchar_t* message; // required

    // Optional: if set to a value that exists in `buttons`, the host SHOULD default-focus it.
    // Use HOST_PROMPT_RESULT_NONE to indicate no preference.
    HostPromptResult defaultResult;

    uint32_t reserved[8];
};

// UUID: {afb5a715-1110-41f3-b7bb-133d6ca735fd}
interface __declspec(uuid("afb5a715-1110-41f3-b7bb-133d6ca735fd")) __declspec(novtable) IHostPrompts : public IUnknown
{
    // Displays a modal prompt overlay and blocks until the user responds.
    // - For pane scopes, cookie MUST identify the correct pane context.
    // - For application scope, cookie MUST be nullptr.
    // - `result` MUST be non-null on entry.
    virtual HRESULT STDMETHODCALLTYPE ShowPrompt(const HostPromptRequest* request, void* cookie, HostPromptResult* result) noexcept = 0;
};
```

### Behavioral contract

- **Threading**: plugins MAY call host UI services from any thread; the host MUST marshal to the UI thread internally.
- **Lifetime**: the host MUST copy strings before returning; plugins may free inputs immediately after the call.
- **Localization**:
  - host-owned alerts must use `.rc` resources (per project guidelines),
  - plugin-provided strings must be localized by the plugin (the host treats them as already-localized).
- **Alerts**:
  - replacement policy: each `(scope, cookie)` pair has at most **one** active alert; `ShowAlert` replaces it.
  - dismissal:
    - if `closable == TRUE`, the host provides a close “X” and Esc dismiss behavior (within that scope),
    - if `closable == FALSE`, dismissal is programmatic (host clears on navigation/path change where appropriate, or plugin calls `ClearAlert`).
- **Prompts**:
  - `ShowPrompt` is blocking until the user responds.
  - button labels are host-localized (standard button sets only); plugins do not provide button text.
  - closing the prompt via window close (if allowed by the host) returns:
    - `HOST_PROMPT_RESULT_OK` for `HOST_PROMPT_BUTTONS_OK`
    - `HOST_PROMPT_RESULT_CANCEL` for all other button sets

## Host Implementation Plan (High Level)

1. **Unify alert surfaces**:
   - keep existing `FolderView` pane-content alert overlay as the implementation for `HOST_ALERT_SCOPE_PANE_CONTENT`,
   - add overlay windows for `HOST_ALERT_SCOPE_PANE` and `HOST_ALERT_SCOPE_APPLICATION` using the same `RedSalamander::Ui::AlertOverlay` renderer.
2. **Input policy**:
   - modal: overlay window swallows input for the scope,
   - modeless: overlay window is hit-test transparent outside the alert card (click-through) but still clickable on close/button UI.
3. **Routing**:
   - map `cookie` → pane instance (left/right) for pane scopes,
   - application scope does not use cookies.
4. **Plugin handshake**:
   - update `RedSalamanderCreate` signature to accept `IHost*`,
   - update all plugins (built-in + shipped) to the new signature,
   - expose `IHostAlerts` + `IHostPrompts` from `IHost`,
   - update built-in plugins to consume it (optional).
