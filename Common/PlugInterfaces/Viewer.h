#pragma once

#include <cstdint>
#include <unknwn.h>
#include <wchar.h>
#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

interface IFileSystem;

#pragma warning(push)
#pragma warning(disable : 4820) // padding in data structure
enum ViewerOpenFlags : uint32_t
{
    VIEWER_OPEN_FLAG_NONE      = 0,
    VIEWER_OPEN_FLAG_START_HEX = 0x1,
};

struct ViewerOpenContext
{
    // Lifetime/ownership:
    // - All pointer fields (including `focusedPath`, `selectionPaths`/elements, and `otherFiles`/elements) are caller-owned.
    // - Callers MAY free/modify these buffers immediately after Open() returns.
    // - Plugins MUST copy any inputs they need to keep beyond the Open() call.
    // - `fileSystem` is a caller-owned COM interface pointer that remains valid at least for the duration of the Open() call.
    //   Plugins that need to use it beyond Open() MUST AddRef() it (and Release() when done).

    // Optional host/main window handle (for initial placement/activation).
    // Note: viewers SHOULD remain independent top-level windows; do not assume Win32 ownership.
    HWND ownerWindow;

    // Active filesystem instance for `focusedPath`/`otherFiles` paths.
    // Paths are filesystem-internal and may not be valid Win32 paths (e.g. "file.txt" inside an archive).
    IFileSystem* fileSystem;

    // Localized display name of the active filesystem plugin (UTF-16, NUL-terminated).
    const wchar_t* fileSystemName;

    // Focused item path (UTF-16, NUL-terminated).
    const wchar_t* focusedPath;

    // Current selection (UTF-16, NUL-terminated paths).
    const wchar_t* const* selectionPaths;
    unsigned long selectionCount;

    // Ordered list of “other files” the viewer can navigate to (UTF-16, NUL-terminated paths).
    // The host typically provides all files in the current folder whose extensions are associated
    // with the same viewer plugin id as `focusedPath`.
    const wchar_t* const* otherFiles;
    unsigned long otherFileCount;
    unsigned long focusedOtherFileIndex;

    ViewerOpenFlags flags;
};

struct ViewerTheme
{
    // ABI version for forward compatibility.
    // Current version: 2
    uint32_t version;

    // DPI of the host window at the time of notification.
    unsigned int dpi;

    // Basic colors (ARGB 0xAARRGGBB).
    uint32_t backgroundArgb;
    uint32_t textArgb;
    uint32_t selectionBackgroundArgb;
    uint32_t selectionTextArgb;
    uint32_t accentArgb;

    // Alert colors (ARGB 0xAARRGGBB).
    uint32_t alertErrorBackgroundArgb;
    uint32_t alertErrorTextArgb;
    uint32_t alertWarningBackgroundArgb;
    uint32_t alertWarningTextArgb;
    uint32_t alertInfoBackgroundArgb;
    uint32_t alertInfoTextArgb;

    // Theme flags.
    BOOL darkMode;
    BOOL highContrast;
    BOOL rainbowMode;
    BOOL darkBase;
};
#pragma warning(pop)

// Host callback for viewer lifecycle events.
// Notes:
// - This is NOT a COM interface (no IUnknown inheritance); lifetime is managed by the host.
// - The host must call IViewer::SetCallback(nullptr, nullptr) before releasing/unloading the plugin.
// - The cookie is provided by the host at registration time and must be passed back verbatim by the plugin.
interface __declspec(novtable) IViewerCallback
{
    virtual HRESULT STDMETHODCALLTYPE ViewerClosed(void* cookie) noexcept = 0;
};

interface __declspec(uuid("d1da10b7-0d0d-4d5c-9b3c-30c386c9d3c7")) __declspec(novtable) IViewer : public IUnknown
{
    // Opens the viewer window or updates its content.
    // Plugins MUST copy any input strings they need to keep; callers own the input buffers.
    virtual HRESULT STDMETHODCALLTYPE Open(const ViewerOpenContext* context) noexcept = 0;

    // Closes the viewer window if it is open. Safe to call multiple times.
    virtual HRESULT STDMETHODCALLTYPE Close() noexcept = 0;

    // Applies the current theme. Plugins MUST accept being called before or after Open().
    virtual HRESULT STDMETHODCALLTYPE SetTheme(const ViewerTheme* theme) noexcept = 0;

    virtual HRESULT STDMETHODCALLTYPE SetCallback(IViewerCallback * callback, void* cookie) noexcept = 0;
};
