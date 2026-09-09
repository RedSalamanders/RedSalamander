#pragma once

#include <cstdint>
#include <string_view>
#include <unknwn.h>

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

// Public terminal ABI names intentionally stay unsuffixed. The current supported
// boundary is source-tree/release lockstep, and every method requires each full
// current record. sizeBytes accepts unknown current-generation tails; it is not
// a mixed-release compatibility promise. Future compatible additions are
// optional tail fields, while incompatible ownership or semantic changes require
// a new COM IID/interface.

#pragma pack(push, 8)

enum class TerminalLifecycleState : uint32_t
{
    Created               = 1u,
    Diagnostic            = 2u,
    Starting              = 3u,
    Running               = 4u,
    DrainingAfterRootExit = 5u,
    Closing               = 6u,
    Exited                = 7u,
    Failed                = 8u,
    Retiring              = 9u,
    Quarantined           = 10u,
};

enum class TerminalActivityTrust : uint32_t
{
    Untrusted = 1u,
    Trusted   = 2u,
};

enum class TerminalFollowState : uint32_t
{
    Disabled = 1u,
    Pending  = 2u,
    Paused   = 3u,
    Applying = 4u,
    Applied  = 5u,
    Failed   = 6u,
    Detached = 7u,
};

enum class TerminalSourceDisposition : uint32_t
{
    Active   = 1u,
    Detached = 2u,
};

enum class TerminalLocationKind : uint32_t
{
    WindowsLocal = 1u,
    WindowsUnc   = 2u,
    Wsl          = 3u,
    PluginBacked = 4u,
    Unsupported  = 5u,
};

enum class TerminalPathInsertionMode : uint32_t
{
    ContextualLeafOrFull = 1u,
    AlwaysFull           = 2u,
    CurrentDirectoryFull = 3u,
};

enum TerminalCapabilityFlags : uint64_t
{
    TerminalCapabilityNone           = 0u,
    TerminalCapabilityFollow         = 0x1u,
    TerminalCapabilityPathInsertion  = 0x2u,
    TerminalCapabilityPromptTracking = 0x4u,
};

struct TerminalUtf16Span
{
    const wchar_t* data = nullptr;
    uint32_t length     = 0u;
    uint32_t reserved   = 0u;
};

// The plugin allocates data with CoTaskMemAlloc. The host owns a successful
// result and releases it with CoTaskMemFree.
struct TerminalOwnedUtf16
{
    wchar_t* data     = nullptr;
    uint32_t length   = 0u;
    uint32_t reserved = 0u;
};

struct TerminalInstanceId
{
    uint8_t bytes[16]{};
};

struct TerminalOriginalSourceKey
{
    uint64_t folderWindowInstanceId = 0u;
    uint64_t paneInstanceId         = 0u;
};

struct TerminalLogicalLocation
{
    uint32_t sizeBytes        = 0u;
    TerminalLocationKind kind = TerminalLocationKind::Unsupported;
    uint32_t reserved0        = 0u;
    TerminalUtf16Span windowsPath{};
    TerminalUtf16Span wslDistribution{};
    TerminalUtf16Span wslAbsolutePath{};
    TerminalUtf16Span pluginShortId{};
    TerminalUtf16Span pluginBackingWindowsPath{};
    uint32_t reserved[4]{};
};

struct TerminalOpenContext
{
    uint32_t sizeBytes = 0u;
    HWND parentWindow  = nullptr;
    TerminalInstanceId instanceId{};
    TerminalOriginalSourceKey originalSource{};
    uint64_t sourceGeneration = 0u;
    TerminalLogicalLocation sourceLocation{};
    TerminalLogicalLocation launchLocation{};
    uint32_t reserved[4]{};
};

struct TerminalTheme
{
    uint32_t sizeBytes               = 0u;
    uint32_t dpi                     = USER_DEFAULT_SCREEN_DPI;
    uint32_t dark                    = 0u;
    uint32_t highContrast            = 0u;
    uint32_t backgroundArgb          = 0xFF000000u;
    uint32_t foregroundArgb          = 0xFFFFFFFFu;
    uint32_t cursorArgb              = 0xFFFFFFFFu;
    uint32_t selectionBackgroundArgb = 0xFF264F78u;
    uint32_t selectionForegroundArgb = 0xFFFFFFFFu;
    uint32_t hyperlinkArgb           = 0xFF4EA1FFu;
    uint32_t inactiveStatusArgb      = 0xFF808080u;
    uint32_t ansiArgb[16]{};
    uint32_t reserved[4]{};
};

struct TerminalSourceUpdate
{
    uint32_t sizeBytes = 0u;
    TerminalOriginalSourceKey originalSource{};
    uint64_t sourceGeneration             = 0u;
    TerminalSourceDisposition disposition = TerminalSourceDisposition::Active;
    uint32_t reserved0                    = 0u;
    TerminalLogicalLocation sourceLocation{};
    uint32_t reserved[4]{};
};

struct TerminalPathInsertion
{
    uint32_t sizeBytes = 0u;
    TerminalOriginalSourceKey initiatingSource{};
    uint64_t initiatingSourceGeneration = 0u;
    TerminalLogicalLocation initiatingSourceLocation{};
    uint64_t itemGeneration        = 0u;
    TerminalPathInsertionMode mode = TerminalPathInsertionMode::ContextualLeafOrFull;
    uint32_t reserved0             = 0u;
    TerminalLogicalLocation itemLocation{};
    TerminalLogicalLocation parentLocation{};
    TerminalUtf16Span displayLeaf{};
    uint32_t reserved[4]{};
};

struct TerminalActivitySnapshot
{
    uint32_t sizeBytes                    = 0u;
    TerminalLifecycleState lifecycleState = TerminalLifecycleState::Created;
    TerminalActivityTrust activityTrust   = TerminalActivityTrust::Untrusted;
    uint64_t generation                   = 0u;
    uint32_t hasRunningCommandOrChild     = 0u;
    uint32_t idleAtPrimaryPrompt          = 0u;
    uint32_t hasPendingUserInput          = 0u;
    uint32_t passwordInputActive          = 0u;
    uint32_t alternateScreenActive        = 0u;
    uint32_t tuiInputOwner                = 0u;
    uint32_t reserved[4]{};
};

struct TerminalViewState
{
    uint32_t sizeBytes = 0u;
    TerminalInstanceId instanceId{};
    uint64_t stateGeneration        = 0u;
    uint64_t sessionGeneration      = 0u;
    uint64_t capabilityFlags        = TerminalCapabilityNone;
    TerminalFollowState followState = TerminalFollowState::Disabled;
    uint32_t exitCodePresent        = 0u;
    uint32_t exitCode               = 0u;
    uint32_t finalSnapshotComplete  = 0u;
    TerminalActivitySnapshot activity{};
    TerminalOwnedUtf16 title{};
    TerminalOwnedUtf16 status{};
    uint32_t reserved[4]{};
};

enum class TerminalShortcutRoute : uint32_t
{
    PassThrough       = 1u,
    Handled           = 2u,
    InvokeHostCommand = 3u,
    Blocked           = 4u,
};

enum TerminalShortcutModifierFlags : uint32_t
{
    TerminalShortcutModifierNone       = 0u,
    TerminalShortcutModifierCtrl       = 0x1u,
    TerminalShortcutModifierAlt        = 0x2u,
    TerminalShortcutModifierShift      = 0x4u,
    TerminalShortcutModifierLeftCtrl   = 0x8u,
    TerminalShortcutModifierRightCtrl  = 0x10u,
    TerminalShortcutModifierLeftAlt    = 0x20u,
    TerminalShortcutModifierRightAlt   = 0x40u,
    TerminalShortcutModifierLeftShift  = 0x80u,
    TerminalShortcutModifierRightShift = 0x100u,
};

struct TerminalShortcutRequest
{
    uint32_t sizeBytes = 0u;
    TerminalUtf16Span commandId{};
    uint32_t message       = 0u;
    uint32_t virtualKey    = 0u;
    uint32_t scanCode      = 0u;
    uint32_t extended      = 0u;
    uint32_t systemKey     = 0u;
    uint32_t repeatCount   = 0u;
    uint32_t previousDown  = 0u;
    uint32_t modifierFlags = TerminalShortcutModifierNone;
    TerminalInstanceId instanceId{};
    uint64_t sessionGeneration = 0u;
    uint32_t reserved[4]{};
};

struct TerminalActionRequest
{
    uint32_t sizeBytes = 0u;
    TerminalUtf16Span commandId{};
    TerminalInstanceId instanceId{};
    uint64_t sessionGeneration = 0u;
    uint32_t reserved[4]{};
};

struct TerminalActionState
{
    uint32_t sizeBytes = 0u;
    uint32_t enabled   = 0u;
    uint32_t checked   = 0u;
    uint32_t reserved0 = 0u;
    RECT childAnchor{};
    uint32_t anchorPresent = 0u;
    uint32_t reserved[4]{};
};

enum class TerminalEventKind : uint32_t
{
    RootSessionExited = 1u,
};

struct TerminalEvent
{
    uint32_t sizeBytes     = 0u;
    TerminalEventKind kind = TerminalEventKind::RootSessionExited;
    TerminalInstanceId instanceId{};
    uint64_t sessionGeneration          = 0u;
    uint64_t stateGeneration            = 0u;
    uint32_t exitCodePresent            = 0u;
    uint32_t exitCode                   = 0u;
    uint32_t finalSnapshotComplete      = 0u;
    int64_t rootExitObservedTimestampNs = 0;
    uint32_t reserved[2]{};
};

#pragma pack(pop)

interface ITerminalEventCallback;

interface __declspec(uuid("47E1AC4C-022E-42C4-8E2C-0C2BBC687D31")) __declspec(novtable) ITerminal : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE Open(const TerminalOpenContext* context) noexcept                     = 0;
    virtual HRESULT STDMETHODCALLTYPE GetChildWindow(HWND * childWindow) noexcept                           = 0;
    virtual HRESULT STDMETHODCALLTYPE SetTheme(const TerminalTheme* theme) noexcept                         = 0;
    virtual HRESULT STDMETHODCALLTYPE UpdateSourceLocation(const TerminalSourceUpdate* update) noexcept     = 0;
    virtual HRESULT STDMETHODCALLTYPE InsertPath(const TerminalPathInsertion* insertion) noexcept           = 0;
    virtual HRESULT STDMETHODCALLTYPE GetViewState(TerminalViewState * state) noexcept                      = 0;
    virtual HRESULT STDMETHODCALLTYPE SetCallback(ITerminalEventCallback * callback, void* cookie) noexcept = 0;
    virtual HRESULT STDMETHODCALLTYPE Close() noexcept                                                      = 0;
};

interface __declspec(uuid("842E2028-816A-440A-A906-33D8606C7899")) __declspec(novtable) ITerminalActions : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE RouteShortcut(const TerminalShortcutRequest* request, TerminalShortcutRoute* route) noexcept = 0;
    virtual HRESULT STDMETHODCALLTYPE GetActionState(const TerminalActionRequest* request, TerminalActionState* state) noexcept    = 0;
    virtual HRESULT STDMETHODCALLTYPE ExecuteAction(const TerminalActionRequest* request) noexcept                                 = 0;
};

// Plugin Route/GetActionState/ExecuteAction and both hosts share this ID set.
// Keyboard copy-or-passthrough stays a RouteShortcut fork; palette/host execute
// still query this list so the command is not stuck at E_NOTIMPL.
[[nodiscard]] constexpr bool IsTerminalPluginActionId(std::wstring_view commandId) noexcept
{
    return commandId == L"cmd/terminal/copy" || commandId == L"cmd/terminal/copySelectionOrPassthrough" || commandId == L"cmd/terminal/find" ||
           commandId == L"cmd/terminal/paste" || commandId == L"cmd/terminal/selectAll" || commandId == L"cmd/terminal/suggestions" ||
           commandId.starts_with(L"cmd/terminal/scroll/") || commandId.starts_with(L"cmd/terminal/font/");
}

// A host owns this weak, non-COM callback. SetCallback(nullptr, nullptr) is a
// hard barrier: after it returns, no callback is active or can begin.
interface __declspec(novtable) ITerminalEventCallback
{
    virtual void STDMETHODCALLTYPE OnTerminalEvent(const TerminalEvent* event, void* cookie) noexcept = 0;

protected:
    ~ITerminalEventCallback() = default;
};
