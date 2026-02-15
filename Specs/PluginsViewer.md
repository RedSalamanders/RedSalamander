# Viewer Plugin Interface Specification

## Overview

Viewer plugins allow RedSalamander to open a **dedicated viewer window** for specific file formats.
The viewer is opened with **F3** based on a file’s **extension** and a user-configurable association in settings.

Key points:
- COM-based plugin architecture (binary compatible, no STL in public interfaces)
- One viewer window per open request (plugin-defined behavior for reuse is allowed)
- Viewer receives:
  - focused file path
  - current selection (if any)
  - “Other Files” list (all files in the current folder mapped to the same viewer plugin)
- Viewer is theme-aware and is notified when the theme changes

## Built-in viewer plugins

Embedded:
- `builtin/viewer-text`: baseline Text/Hex viewer (reference behavior is specified in this document).
- `builtin/viewer-space`: folder disk-usage treemap (see `Specs/ViewerSpacePluginSpec.md`).
- `builtin/viewer-imgraw`: image viewer using WIC + LibRaw (see `Specs/ViewerImgRawPluginSpec.md`).

Optional (loaded from `<exeDir>\\Plugins` when present):
- `builtin/viewer-web`: WebView2-based HTML/PDF viewer (see `Specs/ViewerWebPluginSpec.md`).
- `builtin/viewer-json`: WebView2-based JSON/JSON5 viewer (see `Specs/ViewerWebPluginSpec.md`).
- `builtin/viewer-markdown`: WebView2-based Markdown viewer (see `Specs/ViewerWebPluginSpec.md`).

## Settings (host responsibility)

Viewer selection is controlled by a settings map:

`settings.extensions.openWithViewerByExtension`:
- Key: lowercase extension with leading dot (e.g. `.txt`)
- Value: viewer plugin id (from `PluginMetaData.id`, e.g. `builtin/viewer-text`)

When F3 is pressed on a file:
1. Host extracts the extension (lowercase).
2. Host looks up the viewer plugin id in `openWithViewerByExtension`.
3. If the plugin is available and enabled, host opens the viewer plugin.
4. If no association is found (or plugin missing/disabled), host falls back to `builtin/viewer-text` (which may auto-select Text vs Hex based on file contents). If that fails, host falls back to default behavior.

## Window behavior (host + plugin contract)

- The viewer is a **standard top-level** window (min/max/restore) using `WS_OVERLAPPEDWINDOW` (avoid `WS_EX_TOOLWINDOW` unless a plugin explicitly wants a tool-style window).
- On first display it is positioned **over** the main application window, using the same size and position
  (host provides a window handle for placement; plugin may use `GetWindowRect(...)`).
- The viewer window MUST avoid a “white flash” (or any high-contrast flash) on first show:
  - Host SHOULD call `IViewer::SetTheme()` before `IViewer::Open()` so the plugin can create the window with the correct initial background.
  - Plugins SHOULD ensure the window class has a valid `WNDCLASSEXW::hbrBackground` consistent with the current `ViewerTheme.backgroundArgb` and allow `DefWindowProcW(WM_ERASEBKGND)` to run at least until the first `WM_PAINT` (after which returning `1` is fine for full-surface Direct2D renderers to avoid redundant clears/flicker).
- The viewer behaves like a normal window after opening (it may be in front of or behind the main window); plugins SHOULD NOT rely on Win32 ownership to keep it above the main window.
- Pressing **Esc** closes the viewer window.
  - If an in-view alert overlay is visible (errors/info), **Esc** dismisses the alert first, then closes on the next press.
- “Other Files” navigation is supported using the `otherFiles` list passed at open time.

## Baseline UX expectations (ViewerText reference)

The built-in `builtin/viewer-text` is used as a reference for viewer UX. Other viewer plugins are encouraged to provide equivalent navigation and theme behavior:

- If `otherFileCount > 1`, the viewer SHOULD expose a **filename dropdown** (combo box) listing the `otherFiles` list and selecting `focusedOtherFileIndex`.
- The header SHOULD expose a clearly clickable **mode toggle button** (`TEXT` / `HEX`) so users can switch view mode without hunting in menus.
- The viewer SHOULD also provide “Other Files” navigation commands (Next/Previous/First/Last) using the same shortcuts described in the menu spec.
- In Hex view, selection SHOULD be **byte-accurate**: selecting a byte in the Hex column highlights the corresponding character in the Text column and vice-versa.
- In Hex view, the Text column header SHOULD be clickable to toggle an alternate Unicode byte rendering mode (for inspecting non-ASCII bytes).
- The status bar SHOULD display encoding + size and MUST include a visible-range indicator:
  - Text mode: visible line range (top-bottom), and SHOULD include the total line count when known (`of N`), otherwise `of unknown`.
    - For large streamed files, the status SHOULD include a prefix like `Streaming view (scroll to load more).`
    - For streamed views, plugins SHOULD NOT scan the full file just to compute `N`; leaving `unknown` is acceptable.
  - Hex mode: visible offset range (top-bottom).
- ViewerText provides quick file navigation shortcuts:
  - `Home`: go to the first line of the file.
  - `End`: go to the last line of the file, positioned so the last line is at the bottom of the viewport.
  - These actions are also available from the **View** menu (`Go to Top` / `Go to Bottom`).
- Non-fatal errors/info SHOULD be surfaced via host-rendered alerts (`IHostAlerts`) rather than modal message boxes.
- Viewers MAY expose an “Encoding” menu; `builtin/viewer-text` supports reloading the file under a selected encoding/codepage and optional “Convert on Save …” modes.
- Viewers SHOULD render chrome (header/status) in a theme-aware way and look good in rainbow mode (use the provided `accentArgb` + `rainbowMode` flag).
- Viewer windows SHOULD set a window icon (caption + Alt-Tab) that matches the plugin identity.

### Encoding menu (ViewerText reference)

The built-in `builtin/viewer-text` implements an **Encoding** menu for selecting how the file is decoded for display and (optionally) converted on save:

- Display selection is a **mutually exclusive check group**:
  - Unicode: `ANSI`, `UTF-8`, `UTF-8 BOM`, `UTF-16 BE BOM`, `UTF-16 LE BOM`, `UTF-32 BE BOM`, `UTF-32 LE BOM`.
  - `Character Set`: regional submenus containing common Windows/OEM/ISO code pages.
- Changing the display selection re-reads the file and updates the displayed text using the newly selected encoding/codepage.
- On open, the viewer detects encoding and checks the corresponding display item:
  - BOM-based Unicode files select the matching BOM item.
  - BOM-less files are probed for valid UTF-8; if valid, selects `UTF-8`, otherwise selects `ANSI`.
- Save selection is a **mutually exclusive check group** applied to Save As:
  - Default `Keep Original Encoding`: byte-to-byte save (no conversion).
  - `Convert on Save ...`: converts the currently displayed text to the selected encoding before writing.
- Unsupported code pages are **pruned at runtime** (e.g. via `IsValidCodePage`) so the Encoding menu only exposes code pages available on the current Windows install.
- The status bar displays both:
  - `Detected`: BOM/guess-derived encoding/codepage
  - `Active`: current Encoding display selection

## Factory Method

Viewer plugins are loaded from dynamic libraries and instantiated via the shared factory function:

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

For viewer plugins:
- `RedSalamanderCreate` MUST support creation of `IID_IViewer` and SHOULD return `E_NOINTERFACE` for other IIDs.

### Optional: Multi-Plugin DLL Support

A single DLL MAY implement **multiple logical viewer plugins** for the same interface type (`IID_IViewer`).

To do so, the DLL exports two additional (optional) entry points:

```cpp
extern "C"
{
    // Returns an array of PluginMetaData entries implemented by the DLL for the requested interface type.
    // The array and all strings are owned by the DLL and remain valid until the DLL is unloaded.
    PLUGFACTORY_API HRESULT __stdcall RedSalamanderEnumeratePlugins(
        REFIID riid,
        const PluginMetaData** metaData,
        unsigned int* count
    );

    // Creates a specific plugin instance identified by pluginId (metaData[i].id).
    PLUGFACTORY_API HRESULT __stdcall RedSalamanderCreateEx(
        REFIID riid,
        const FactoryOptions* factoryOptions,
        IHost* host,
        const wchar_t* pluginId,
        void** result
    );
}
```

**Host behavior:**
- If `RedSalamanderEnumeratePlugins` is present, the host calls it during discovery and registers one plugin entry per returned `PluginMetaData` record.
- When instantiating a plugin entry originating from enumeration, the host calls `RedSalamanderCreateEx` with `pluginId == metaData[i].id`.
- If the optional exports are missing, the host falls back to `RedSalamanderCreate`.

**Plugin behavior:**
- `RedSalamanderEnumeratePlugins` MUST return stable metadata pointers for the lifetime of the loaded DLL.
- `RedSalamanderCreateEx` MUST return `HRESULT_FROM_WIN32(ERROR_NOT_FOUND)` (or `E_INVALIDARG`) for unknown `pluginId` values.

The host obtains `IInformations` via `QueryInterface` on the returned `IViewer` instance.

`host` is a caller-owned COM interface pointer that remains valid for the lifetime of the created plugin instance. Plugins MAY `QueryInterface` it for host services (see `Common/PlugInterfaces/Host.h`).

## Plugin Interfaces

### 0. IInformations (required)

Viewer plugins MUST implement `IInformations` (metadata + configuration) and expose it via `QueryInterface` on the `IViewer` instance.

See `Common/PlugInterfaces/Informations.h`.

### 1. IViewer (required)

`IViewer` is the viewer entry point. Instances are created by `RedSalamanderCreate(IID_IViewer, ...)`.

Key responsibilities:
- Create/show the viewer window on `Open()`
- Close the viewer window on `Close()` (also triggered by Esc)
- Apply theme changes via `SetTheme()`
- Notify the host when the viewer window closes via `SetCallback()`

### 2. IViewerCallback (host callback, non-COM)

`IViewerCallback` is a raw vtable callback (NOT COM):
- plugin must treat it as a weak pointer
- host provides an opaque `cookie` and plugin must pass it back verbatim
- host must call `IViewer::SetCallback(nullptr, nullptr)` before releasing/unloading the plugin

## Theme contract

The host calls `IViewer::SetTheme()`:
- once immediately after creating the viewer (plugin must accept either order), but host SHOULD call it before `Open()` to avoid first-show background flash
- every time the application theme changes

The theme structure contains basic colors and flags (dark/high-contrast/rainbow) plus DPI.
Plugins must not read global app state for theme; they must rely on this notification.

## Reference ABI (current)

The current public ABI is defined in `Common/PlugInterfaces/Viewer.h`:

**Lifetime / ownership**
- The host owns all pointer-backed inputs in `ViewerOpenContext` (`focusedPath`, `selectionPaths`, `otherFiles`, and all array elements).
- All pointers are **ephemeral**: plugins MUST NOT retain them beyond the `IViewer::Open()` call.
- If a plugin needs to use any value later (async work, UI callbacks, etc.), it MUST copy the strings/arrays into plugin-owned storage.
- The same rule applies to `ViewerTheme*` passed to `IViewer::SetTheme()` (copy if needed after the call).
- `ViewerOpenContext.fileSystem` is a caller-owned COM interface pointer that remains valid at least for the duration of the `Open()` call; plugins MUST `AddRef()` it if they need it beyond `Open()`.

```cpp
enum ViewerOpenFlags : uint32_t
{
    VIEWER_OPEN_FLAG_NONE        = 0,
    VIEWER_OPEN_FLAG_START_HEX   = 0x1,
};

struct ViewerOpenContext
{
    HWND ownerWindow;

    // Active filesystem instance for `focusedPath`/`otherFiles` paths.
    // Paths are filesystem-internal and may not be valid Win32 paths (e.g. "file.txt" inside an archive).
    IFileSystem* fileSystem;

    // Localized display name of the active filesystem plugin (UTF-16, NUL-terminated).
    const wchar_t* fileSystemName;

    const wchar_t* focusedPath;

    const wchar_t* const* selectionPaths;
    unsigned long selectionCount;

    const wchar_t* const* otherFiles;
    unsigned long otherFileCount;
    unsigned long focusedOtherFileIndex;

    ViewerOpenFlags flags;
};

struct ViewerTheme
{
    uint32_t version; // 2
    unsigned int dpi;

    uint32_t backgroundArgb;
    uint32_t textArgb;
    uint32_t selectionBackgroundArgb;
    uint32_t selectionTextArgb;
    uint32_t accentArgb;

    uint32_t alertErrorBackgroundArgb;
    uint32_t alertErrorTextArgb;
    uint32_t alertWarningBackgroundArgb;
    uint32_t alertWarningTextArgb;
    uint32_t alertInfoBackgroundArgb;
    uint32_t alertInfoTextArgb;

    BOOL darkMode;
    BOOL highContrast;
    BOOL rainbowMode;
    BOOL darkBase;
};

interface __declspec(novtable) IViewerCallback
{
    virtual HRESULT STDMETHODCALLTYPE ViewerClosed(void* cookie) noexcept = 0;
};

// UUID: {d1da10b7-0d0d-4d5c-9b3c-30c386c9d3c7}
interface __declspec(uuid("d1da10b7-0d0d-4d5c-9b3c-30c386c9d3c7")) __declspec(novtable) IViewer : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE Open(const ViewerOpenContext* context) noexcept = 0;
    virtual HRESULT STDMETHODCALLTYPE Close() noexcept = 0;
    virtual HRESULT STDMETHODCALLTYPE SetTheme(const ViewerTheme* theme) noexcept = 0;
    virtual HRESULT STDMETHODCALLTYPE SetCallback(IViewerCallback* callback, void* cookie) noexcept = 0;
};
```
