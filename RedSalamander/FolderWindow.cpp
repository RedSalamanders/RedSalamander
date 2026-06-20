#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <winsock2.h>

#include "BatchRenameWindow.h"
#include "DxUiThemePalette.h"
#include "FolderWindowInternal.h"
#include "HostServices.h"
#include "NavigationLocation.h"
#include "SelfTestCommon.h"

#include <netioapi.h>
#include <shlwapi.h>
#include <shobjidl.h>
#include <winioctl.h>

#pragma comment(lib, "Iphlpapi.lib")
#pragma comment(lib, "Shlwapi.lib")

#ifdef ENABLE_TESTS
[[nodiscard]] bool IsFolderWindowSelfTestTracingEnabled() noexcept
{
    return ! SelfTest::GetRunStartedUtcIso().empty();
}
#endif

namespace
{
constexpr UINT_PTR kPreviewPaneRefreshTimerId = 0x7250;

[[nodiscard]] bool EnsureFolderWindowDxHostClass(HINSTANCE instance) noexcept
{
    WNDCLASSW existing{};
    if (GetClassInfoW(instance, kFolderWindowDxHostClassName, &existing) != FALSE)
    {
        return true;
    }

    WNDCLASSW wc{};
    wc.lpfnWndProc   = &FolderWindowDxHostWndProc;
    wc.hInstance     = instance;
    wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
    wc.hbrBackground = nullptr;
    wc.lpszClassName = kFolderWindowDxHostClassName;
    return RegisterClassW(&wc) != 0 || GetLastError() == ERROR_CLASS_ALREADY_EXISTS;
}

LRESULT OnHostServicesMessage(UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    LRESULT result = 0;
    if (TryHandleHostServicesWindowMessage(msg, wp, lp, result))
    {
        return result;
    }
    return 0;
}

[[nodiscard]] COLORREF SplitterGripColor(const AppTheme& theme) noexcept
{
    if (theme.highContrast)
    {
        return theme.menu.text;
    }

    constexpr int kTowardTextWeight = 1;
    constexpr int kDenom            = 4;
    static_assert(kTowardTextWeight > 0 && kTowardTextWeight < kDenom);

    const int baseWeight           = kDenom - kTowardTextWeight;
    const COLORREF baseColor       = theme.menu.separator;
    const COLORREF towardTextColor = theme.menu.text;

    const int r = (static_cast<int>(GetRValue(baseColor)) * baseWeight + static_cast<int>(GetRValue(towardTextColor)) * kTowardTextWeight) / kDenom;
    const int g = (static_cast<int>(GetGValue(baseColor)) * baseWeight + static_cast<int>(GetGValue(towardTextColor)) * kTowardTextWeight) / kDenom;
    const int b = (static_cast<int>(GetBValue(baseColor)) * baseWeight + static_cast<int>(GetBValue(towardTextColor)) * kTowardTextWeight) / kDenom;

    return RGB(static_cast<BYTE>(r), static_cast<BYTE>(g), static_cast<BYTE>(b));
}

[[nodiscard]] std::wstring FileActionTargetText(const std::optional<std::filesystem::path>& targetPath) noexcept
{
    if (! targetPath.has_value() || targetPath.value().empty())
    {
        return {};
    }

    std::wstring target = targetPath.value().filename().wstring();
    if (target.empty())
    {
        target = targetPath.value().wstring();
    }
    return target;
}

void ShowFileActionUnavailableOverlay(FolderWindow& window,
                                      FolderWindow::Pane pane,
                                      bool viewerAction,
                                      std::wstring_view actionId,
                                      const std::optional<std::filesystem::path>& targetPath) noexcept
{
    Debug::Perf::Scope perf(L"fileaction.feedback_us");
    perf.SetDetail(viewerAction ? L"viewer-unavailable" : L"editor-unavailable");
    perf.SetValue0(static_cast<uint64_t>(actionId.size()));

    std::wstring title = LoadStringResource(nullptr, viewerAction ? IDS_FILEACTION_VIEWER_UNAVAILABLE_TITLE : IDS_FILEACTION_EDITOR_UNAVAILABLE_TITLE);
    if (title.empty())
    {
        title = LoadStringResource(nullptr, IDS_CAPTION_WARNING);
    }

    std::wstring target  = FileActionTargetText(targetPath);
    std::wstring message = FormatStringResource(
        nullptr, viewerAction ? IDS_FMT_FILEACTION_VIEWER_NOT_AVAILABLE : IDS_FMT_FILEACTION_EDITOR_NOT_AVAILABLE, std::wstring(actionId), target);
    if (message.empty())
    {
        message = std::wstring(actionId);
    }

    window.ShowPaneAlertOverlay(
        pane, FolderView::ErrorOverlayKind::Operation, FolderView::OverlaySeverity::Warning, std::move(title), std::move(message), S_OK, true, false);
}

void ShowFileActionLaunchFailureOverlay(FolderWindow& window,
                                        FolderWindow::Pane pane,
                                        bool viewerAction,
                                        std::wstring_view actionId,
                                        const std::optional<std::filesystem::path>& targetPath,
                                        HRESULT hr) noexcept
{
    Debug::Perf::Scope perf(L"fileaction.feedback_us");
    perf.SetDetail(viewerAction ? L"viewer-launch-failed" : L"editor-launch-failed");
    perf.SetValue0(static_cast<uint64_t>(actionId.size()));
    perf.SetHr(hr);

    std::wstring title = LoadStringResource(nullptr, viewerAction ? IDS_FILEACTION_VIEWER_UNAVAILABLE_TITLE : IDS_FILEACTION_EDITOR_UNAVAILABLE_TITLE);
    if (title.empty())
    {
        title = LoadStringResource(nullptr, IDS_CAPTION_WARNING);
    }

    std::wstring target  = FileActionTargetText(targetPath);
    std::wstring message = FormatStringResource(nullptr,
                                                viewerAction ? IDS_FMT_FILEACTION_VIEWER_LAUNCH_FAILED : IDS_FMT_FILEACTION_EDITOR_LAUNCH_FAILED,
                                                std::wstring(actionId),
                                                target,
                                                static_cast<unsigned long>(static_cast<uint32_t>(hr)));
    if (message.empty())
    {
        message = std::wstring(actionId);
    }

    window.ShowPaneAlertOverlay(
        pane, FolderView::ErrorOverlayKind::Operation, FolderView::OverlaySeverity::Warning, std::move(title), std::move(message), hr, true, false);
}

void ShowAlternateFileActionUnavailableOverlay(FolderWindow& window,
                                               FolderWindow::Pane pane,
                                               bool viewerAction,
                                               const std::optional<std::filesystem::path>& targetPath) noexcept
{
    Debug::Perf::Scope perf(L"fileaction.feedback_us");
    perf.SetDetail(viewerAction ? L"alternate-viewer-unavailable" : L"alternate-editor-unavailable");

    std::wstring title = LoadStringResource(nullptr, viewerAction ? IDS_FILEACTION_VIEWER_UNAVAILABLE_TITLE : IDS_FILEACTION_EDITOR_UNAVAILABLE_TITLE);
    if (title.empty())
    {
        title = LoadStringResource(nullptr, IDS_CAPTION_WARNING);
    }

    std::wstring target  = FileActionTargetText(targetPath);
    std::wstring message = FormatStringResource(
        nullptr, viewerAction ? IDS_FMT_FILEACTION_ALTERNATE_VIEWER_NOT_CONFIGURED : IDS_FMT_FILEACTION_ALTERNATE_EDITOR_NOT_CONFIGURED, target);
    if (message.empty())
    {
        message = target;
    }

    window.ShowPaneAlertOverlay(
        pane, FolderView::ErrorOverlayKind::Operation, FolderView::OverlaySeverity::Warning, std::move(title), std::move(message), S_OK, true, false);
}

[[nodiscard]] std::wstring ShellActionTargetText(const std::filesystem::path& path) noexcept
{
    std::wstring target = path.filename().wstring();
    if (target.empty())
    {
        target = path.wstring();
    }
    return target;
}

void ShowShellActionUnavailableOverlay(FolderWindow& window, FolderWindow::Pane pane, UINT messageStringId) noexcept
{
    Debug::Perf::Scope perf(L"shell.feedback_us");
    perf.SetDetail(L"unavailable");

    std::wstring title = LoadStringResource(nullptr, IDS_SHELL_ACTION_UNAVAILABLE_TITLE);
    if (title.empty())
    {
        title = LoadStringResource(nullptr, IDS_CAPTION_WARNING);
    }

    std::wstring message = LoadStringResource(nullptr, messageStringId);
    if (message.empty())
    {
        message = title;
    }

    window.ShowPaneAlertOverlay(
        pane, FolderView::ErrorOverlayKind::Operation, FolderView::OverlaySeverity::Warning, std::move(title), std::move(message), S_OK, true, false);
}

void ShowShellActionFailedOverlay(FolderWindow& window, FolderWindow::Pane pane, const std::filesystem::path& path, HRESULT hr) noexcept
{
    Debug::Perf::Scope perf(L"shell.feedback_us");
    perf.SetDetail(L"failed");
    perf.SetHr(hr);

    std::wstring title = LoadStringResource(nullptr, IDS_SHELL_ACTION_UNAVAILABLE_TITLE);
    if (title.empty())
    {
        title = LoadStringResource(nullptr, IDS_CAPTION_WARNING);
    }

    std::wstring message =
        FormatStringResource(nullptr, IDS_FMT_SHELL_ACTION_FAILED, ShellActionTargetText(path), static_cast<unsigned long>(static_cast<uint32_t>(hr)));
    if (message.empty())
    {
        message = ShellActionTargetText(path);
    }

    window.ShowPaneAlertOverlay(
        pane, FolderView::ErrorOverlayKind::Operation, FolderView::OverlaySeverity::Warning, std::move(title), std::move(message), hr, true, false);
}

enum class ShortcutTargetResolutionStatus : uint8_t
{
    Resolved,
    Unsupported,
    Failed
};

struct ShortcutTargetResolution final
{
    ShortcutTargetResolutionStatus status = ShortcutTargetResolutionStatus::Unsupported;
    std::filesystem::path target;
    HRESULT hr = S_OK;
};

struct MountPointReparseDataBuffer final
{
    ULONG ReparseTag            = IO_REPARSE_TAG_MOUNT_POINT;
    USHORT ReparseDataLength    = 0;
    USHORT Reserved             = 0;
    USHORT SubstituteNameOffset = 0;
    USHORT SubstituteNameLength = 0;
    USHORT PrintNameOffset      = 0;
    USHORT PrintNameLength      = 0;
    wchar_t PathBuffer[1]{};
};

struct SymbolicLinkReparseDataBuffer final
{
    ULONG ReparseTag            = IO_REPARSE_TAG_SYMLINK;
    USHORT ReparseDataLength    = 0;
    USHORT Reserved             = 0;
    USHORT SubstituteNameOffset = 0;
    USHORT SubstituteNameLength = 0;
    USHORT PrintNameOffset      = 0;
    USHORT PrintNameLength      = 0;
    ULONG Flags                 = 0;
    wchar_t PathBuffer[1]{};
};

[[nodiscard]] std::wstring TrimShortcutValue(std::wstring_view value)
{
    while (! value.empty() && (value.front() == L' ' || value.front() == L'\t' || value.front() == L'\r' || value.front() == L'\n'))
    {
        value.remove_prefix(1);
    }
    while (! value.empty() && (value.back() == L' ' || value.back() == L'\t' || value.back() == L'\r' || value.back() == L'\n'))
    {
        value.remove_suffix(1);
    }
    return std::wstring(value);
}

[[nodiscard]] bool IsWindowsAbsolutePathText(std::wstring_view text) noexcept
{
    if (text.size() >= 3 && std::iswalpha(static_cast<wint_t>(text[0])) != 0 && text[1] == L':' && (text[2] == L'\\' || text[2] == L'/'))
    {
        return true;
    }

    return text.size() >= 2 && ((text[0] == L'\\' && text[1] == L'\\') || (text[0] == L'/' && text[1] == L'/'));
}

[[nodiscard]] HRESULT ConvertFileUrlToLocalPath(std::wstring_view url, std::filesystem::path& outPath) noexcept
{
    outPath.clear();

    std::wstring urlText(url);
    std::wstring pathText(32768, L'\0');
    DWORD pathCharCount = static_cast<DWORD>(pathText.size());
    const HRESULT hr    = PathCreateFromUrlW(urlText.c_str(), pathText.data(), &pathCharCount, 0);
    if (FAILED(hr))
    {
        return hr;
    }

    const size_t terminator = pathText.find(L'\0');
    if (terminator != std::wstring::npos)
    {
        pathText.resize(terminator);
    }
    else
    {
        pathText.resize(std::min<size_t>(pathCharCount, pathText.size()));
    }

    if (pathText.empty())
    {
        return E_INVALIDARG;
    }

    outPath = std::filesystem::path(pathText);
    return S_OK;
}

[[nodiscard]] bool TryReadInternetShortcutUrl(const std::filesystem::path& shortcutPath, std::wstring& outUrl) noexcept
{
    outUrl.clear();

    std::wstring value(32768, L'\0');
    const DWORD copied = GetPrivateProfileStringW(L"InternetShortcut", L"URL", L"", value.data(), static_cast<DWORD>(value.size()), shortcutPath.c_str());
    if (copied == 0 || copied >= (value.size() - 1u))
    {
        return false;
    }

    value.resize(copied);
    outUrl = TrimShortcutValue(value);
    return ! outUrl.empty();
}

[[nodiscard]] ShortcutTargetResolution ResolveInternetShortcutTarget(const std::filesystem::path& shortcutPath) noexcept
{
    std::wstring url;
    if (! TryReadInternetShortcutUrl(shortcutPath, url))
    {
        return {.status = ShortcutTargetResolutionStatus::Failed, .hr = HRESULT_FROM_WIN32(ERROR_BAD_FORMAT)};
    }

    std::filesystem::path target;
    if (OrdinalString::StartsWithNoCase(url, L"file:"))
    {
        const HRESULT hr = ConvertFileUrlToLocalPath(url, target);
        if (FAILED(hr))
        {
            return {.status = ShortcutTargetResolutionStatus::Failed, .hr = hr};
        }

        return {.status = ShortcutTargetResolutionStatus::Resolved, .target = std::move(target), .hr = S_OK};
    }

    if (IsWindowsAbsolutePathText(url))
    {
        return {.status = ShortcutTargetResolutionStatus::Resolved, .target = std::filesystem::path(url), .hr = S_OK};
    }

    return {.status = ShortcutTargetResolutionStatus::Unsupported};
}

[[nodiscard]] ShortcutTargetResolution ResolveShellLinkTarget(const std::filesystem::path& shortcutPath) noexcept
{
    const HRESULT coHr      = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
    const bool uninitialize = SUCCEEDED(coHr);
    const auto coCleanup    = wil::scope_exit([&]
    {
        if (uninitialize)
        {
            CoUninitialize();
        }
    });
    if (FAILED(coHr) && coHr != RPC_E_CHANGED_MODE)
    {
        return {.status = ShortcutTargetResolutionStatus::Failed, .hr = coHr};
    }

    wil::com_ptr<IShellLinkW> shellLink;
    HRESULT hr = CoCreateInstance(CLSID_ShellLink, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(shellLink.put()));
    if (FAILED(hr) || ! shellLink)
    {
        return {.status = ShortcutTargetResolutionStatus::Failed, .hr = FAILED(hr) ? hr : E_POINTER};
    }

    wil::com_ptr<IPersistFile> persistFile;
    hr = shellLink->QueryInterface(IID_PPV_ARGS(persistFile.put()));
    if (FAILED(hr) || ! persistFile)
    {
        return {.status = ShortcutTargetResolutionStatus::Failed, .hr = FAILED(hr) ? hr : E_NOINTERFACE};
    }

    hr = persistFile->Load(shortcutPath.c_str(), STGM_READ);
    if (FAILED(hr))
    {
        return {.status = ShortcutTargetResolutionStatus::Failed, .hr = hr};
    }

    WIN32_FIND_DATAW findData{};
    std::wstring targetText(32768, L'\0');
    hr = shellLink->GetPath(targetText.data(), static_cast<int>(targetText.size()), &findData, SLGP_UNCPRIORITY);
    if (FAILED(hr))
    {
        return {.status = ShortcutTargetResolutionStatus::Failed, .hr = hr};
    }

    const size_t terminator = targetText.find(L'\0');
    if (terminator != std::wstring::npos)
    {
        targetText.resize(terminator);
    }
    if (targetText.empty())
    {
        return {.status = ShortcutTargetResolutionStatus::Failed, .hr = HRESULT_FROM_WIN32(ERROR_BAD_FORMAT)};
    }

    return {.status = ShortcutTargetResolutionStatus::Resolved, .target = std::filesystem::path(targetText), .hr = S_OK};
}

[[nodiscard]] std::optional<std::wstring> ExtractReparsePathBufferString(const wchar_t* pathBuffer,
                                                                         USHORT substituteNameOffset,
                                                                         USHORT substituteNameLength,
                                                                         size_t pathBufferBytes)
{
    if ((substituteNameOffset % sizeof(wchar_t)) != 0 || (substituteNameLength % sizeof(wchar_t)) != 0)
    {
        return std::nullopt;
    }
    if (static_cast<size_t>(substituteNameOffset) + static_cast<size_t>(substituteNameLength) > pathBufferBytes)
    {
        return std::nullopt;
    }

    const size_t charOffset = substituteNameOffset / sizeof(wchar_t);
    const size_t charLength = substituteNameLength / sizeof(wchar_t);
    if (charLength == 0)
    {
        return std::nullopt;
    }

    return std::wstring(pathBuffer + charOffset, charLength);
}

[[nodiscard]] std::wstring NormalizeReparseSubstituteName(std::wstring_view substituteName)
{
    constexpr std::wstring_view kNtDosPrefix = LR"(\??\)";
    if (OrdinalString::StartsWithNoCase(substituteName, kNtDosPrefix))
    {
        const std::wstring_view tail = substituteName.substr(kNtDosPrefix.size());
        if (OrdinalString::StartsWithNoCase(tail, LR"(UNC\)"))
        {
            return LR"(\\)" + std::wstring(tail.substr(4));
        }
        if (OrdinalString::StartsWithNoCase(tail, L"Volume{"))
        {
            return LR"(\\?\)" + std::wstring(tail);
        }
        return std::wstring(tail);
    }

    constexpr std::wstring_view kMupPrefix = LR"(\Device\Mup\)";
    if (OrdinalString::StartsWithNoCase(substituteName, kMupPrefix))
    {
        return LR"(\\)" + std::wstring(substituteName.substr(kMupPrefix.size()));
    }

    return std::wstring(substituteName);
}

[[nodiscard]] ShortcutTargetResolution ResolveReparsePointTarget(const std::filesystem::path& sourcePath) noexcept
{
    constexpr DWORD kReparseBufferSize = 16u * 1024u;
    std::vector<std::byte> reparseBuffer(kReparseBufferSize);

#pragma warning(push)
#pragma warning(disable : 4625 4626) // WIL unique_hfile copy operations are intentionally deleted.
    wil::unique_hfile reparseHandle(CreateFileW(sourcePath.c_str(),
                                                FILE_READ_ATTRIBUTES,
                                                FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                                nullptr,
                                                OPEN_EXISTING,
                                                FILE_FLAG_OPEN_REPARSE_POINT | FILE_FLAG_BACKUP_SEMANTICS,
                                                nullptr));
#pragma warning(pop)
    if (! reparseHandle)
    {
        return {.status = ShortcutTargetResolutionStatus::Failed, .hr = HRESULT_FROM_WIN32(GetLastError())};
    }

    DWORD bytesReturned = 0;
    if (DeviceIoControl(reparseHandle.get(),
                        FSCTL_GET_REPARSE_POINT,
                        nullptr,
                        0,
                        reparseBuffer.data(),
                        static_cast<DWORD>(reparseBuffer.size()),
                        &bytesReturned,
                        nullptr) == FALSE)
    {
        return {.status = ShortcutTargetResolutionStatus::Failed, .hr = HRESULT_FROM_WIN32(GetLastError())};
    }

    if (bytesReturned < offsetof(MountPointReparseDataBuffer, PathBuffer))
    {
        return {.status = ShortcutTargetResolutionStatus::Failed, .hr = HRESULT_FROM_WIN32(ERROR_BAD_FORMAT)};
    }

    const auto* header = reinterpret_cast<const MountPointReparseDataBuffer*>(reparseBuffer.data());
    std::optional<std::wstring> substituteName;
    if (header->ReparseTag == IO_REPARSE_TAG_MOUNT_POINT)
    {
        const auto* mountPoint       = reinterpret_cast<const MountPointReparseDataBuffer*>(reparseBuffer.data());
        const size_t pathBufferBytes = mountPoint->ReparseDataLength >= (4u * sizeof(USHORT)) ? mountPoint->ReparseDataLength - (4u * sizeof(USHORT)) : 0u;
        substituteName =
            ExtractReparsePathBufferString(mountPoint->PathBuffer, mountPoint->SubstituteNameOffset, mountPoint->SubstituteNameLength, pathBufferBytes);
    }
    else if (header->ReparseTag == IO_REPARSE_TAG_SYMLINK)
    {
        if (bytesReturned < offsetof(SymbolicLinkReparseDataBuffer, PathBuffer))
        {
            return {.status = ShortcutTargetResolutionStatus::Failed, .hr = HRESULT_FROM_WIN32(ERROR_BAD_FORMAT)};
        }

        const auto* symlink = reinterpret_cast<const SymbolicLinkReparseDataBuffer*>(reparseBuffer.data());
        const size_t pathBufferBytes =
            symlink->ReparseDataLength >= ((4u * sizeof(USHORT)) + sizeof(ULONG)) ? symlink->ReparseDataLength - ((4u * sizeof(USHORT)) + sizeof(ULONG)) : 0u;
        substituteName = ExtractReparsePathBufferString(symlink->PathBuffer, symlink->SubstituteNameOffset, symlink->SubstituteNameLength, pathBufferBytes);
    }
    else
    {
        return {.status = ShortcutTargetResolutionStatus::Unsupported};
    }

    if (! substituteName.has_value() || substituteName.value().empty())
    {
        return {.status = ShortcutTargetResolutionStatus::Failed, .hr = HRESULT_FROM_WIN32(ERROR_BAD_FORMAT)};
    }

    std::wstring target = NormalizeReparseSubstituteName(substituteName.value());
    if (target.empty())
    {
        return {.status = ShortcutTargetResolutionStatus::Failed, .hr = HRESULT_FROM_WIN32(ERROR_BAD_FORMAT)};
    }

    return {.status = ShortcutTargetResolutionStatus::Resolved, .target = std::filesystem::path(std::move(target)), .hr = S_OK};
}

[[nodiscard]] ShortcutTargetResolution ResolveShortcutOrLinkTarget(const std::filesystem::path& sourcePath) noexcept
{
    if (OrdinalString::EqualsNoCase(sourcePath.extension().wstring(), L".url"))
    {
        return ResolveInternetShortcutTarget(sourcePath);
    }

    if (OrdinalString::EqualsNoCase(sourcePath.extension().wstring(), L".lnk"))
    {
        return ResolveShellLinkTarget(sourcePath);
    }

    const DWORD attributes = GetFileAttributesW(sourcePath.c_str());
    if (attributes == INVALID_FILE_ATTRIBUTES)
    {
        return {.status = ShortcutTargetResolutionStatus::Failed, .hr = HRESULT_FROM_WIN32(GetLastError())};
    }
    if ((attributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0u)
    {
        return ResolveReparsePointTarget(sourcePath);
    }

    return {.status = ShortcutTargetResolutionStatus::Unsupported};
}

[[nodiscard]] POINT ResolveShellContextMenuPoint(HWND owner) noexcept
{
    POINT point{};
    RECT windowRect{};
    if (owner && IsWindow(owner) != FALSE && GetWindowRect(owner, &windowRect) != FALSE)
    {
        point.x = windowRect.left + ((windowRect.right - windowRect.left) / 2);
        point.y = windowRect.top + ((windowRect.bottom - windowRect.top) / 2);
    }
    return point;
}

[[nodiscard]] HRESULT ShowShellContextMenuForPath(HWND owner, const std::filesystem::path& path) noexcept
{
    PIDLIST_ABSOLUTE rawPidl = nullptr;
    SFGAOF attributes        = 0;
    HRESULT hr               = SHParseDisplayName(path.c_str(), nullptr, &rawPidl, 0, &attributes);
    if (FAILED(hr))
    {
        Debug::Warning(L"Shell context menu: SHParseDisplayName failed for {}: {:#x}", path.wstring(), hr);
        return hr;
    }

    wil::unique_cotaskmem_ptr<ITEMIDLIST_ABSOLUTE> pidl(reinterpret_cast<ITEMIDLIST_ABSOLUTE*>(rawPidl));

    wil::com_ptr<IShellFolder> parentFolder;
    PCUITEMID_CHILD childPidl = nullptr;
    hr                        = SHBindToParent(pidl.get(), IID_IShellFolder, parentFolder.put_void(), &childPidl);
    if (FAILED(hr) || ! parentFolder || ! childPidl)
    {
        Debug::Warning(L"Shell context menu: SHBindToParent failed for {}: {:#x}", path.wstring(), hr);
        return FAILED(hr) ? hr : E_FAIL;
    }

    wil::com_ptr<IContextMenu> contextMenu;
    hr = parentFolder->GetUIObjectOf(owner, 1, &childPidl, IID_IContextMenu, nullptr, contextMenu.put_void());
    if (FAILED(hr) || ! contextMenu)
    {
        Debug::Warning(L"Shell context menu: GetUIObjectOf(IContextMenu) failed for {}: {:#x}", path.wstring(), hr);
        return FAILED(hr) ? hr : E_FAIL;
    }

    wil::unique_hmenu menu(CreatePopupMenu());
    if (! menu)
    {
        const DWORD lastError = Debug::ErrorWithLastError(L"Shell context menu: CreatePopupMenu failed");
        return HRESULT_FROM_WIN32(lastError);
    }

    constexpr UINT kFirstCommandId = 1u;
    constexpr UINT kLastCommandId  = 0x7FFFu;
    hr                             = contextMenu->QueryContextMenu(menu.get(), 0, kFirstCommandId, kLastCommandId, CMF_NORMAL);
    if (FAILED(hr))
    {
        Debug::Warning(L"Shell context menu: QueryContextMenu failed for {}: {:#x}", path.wstring(), hr);
        return hr;
    }

    const POINT point        = ResolveShellContextMenuPoint(owner);
    const UINT chosenCommand = static_cast<UINT>(TrackPopupMenuEx(menu.get(), TPM_RETURNCMD | TPM_RIGHTBUTTON, point.x, point.y, owner, nullptr));
    if (chosenCommand == 0u)
    {
        return S_FALSE;
    }

    CMINVOKECOMMANDINFOEX invoke{};
    invoke.cbSize   = sizeof(invoke);
    invoke.fMask    = CMIC_MASK_UNICODE | CMIC_MASK_PTINVOKE;
    invoke.hwnd     = owner;
    invoke.lpVerb   = MAKEINTRESOURCEA(chosenCommand - kFirstCommandId);
    invoke.lpVerbW  = MAKEINTRESOURCEW(chosenCommand - kFirstCommandId);
    invoke.nShow    = SW_SHOWNORMAL;
    invoke.ptInvoke = point;

    hr = contextMenu->InvokeCommand(reinterpret_cast<LPCMINVOKECOMMANDINFO>(&invoke));
    if (FAILED(hr))
    {
        Debug::Warning(L"Shell context menu: InvokeCommand failed for {}: {:#x}", path.wstring(), hr);
    }
    return hr;
}
} // namespace

void FolderWindow::ClearFileActionFailure() noexcept
{
    _lastFileActionFailure.reset();
}

void FolderWindow::RecordFileActionLaunchFailure(bool viewerAction, std::wstring_view actionId, const std::filesystem::path& targetPath, HRESULT hr) noexcept
{
    FileActionFailure failure{};
    failure.kind           = FileActionFailureKind::LaunchFailed;
    failure.viewerAction   = viewerAction;
    failure.actionId       = std::wstring(actionId);
    failure.targetPath     = targetPath;
    failure.hr             = hr;
    _lastFileActionFailure = std::move(failure);
}

std::optional<FolderWindow::FileActionFailure> FolderWindow::TakeFileActionFailure() noexcept
{
    std::optional<FileActionFailure> failure = std::move(_lastFileActionFailure);
    _lastFileActionFailure.reset();
    return failure;
}

bool FolderWindow::ShowRecordedFileActionFailureOverlay(Pane pane) noexcept
{
    std::optional<FileActionFailure> failure = TakeFileActionFailure();
    if (! failure.has_value())
    {
        return false;
    }

    if (failure.value().kind != FileActionFailureKind::LaunchFailed)
    {
        return false;
    }

    ShowFileActionLaunchFailureOverlay(*this, pane, failure.value().viewerAction, failure.value().actionId, failure.value().targetPath, failure.value().hr);
    return true;
}

class FolderWindow::NetworkChangeSubscription final
{
public:
    explicit NetworkChangeSubscription(HWND hwnd) noexcept : _hwnd(hwnd)
    {
        if (! _hwnd)
        {
            return;
        }

        HANDLE handle                  = nullptr;
        const BOOL initialNotification = FALSE;

        const NTSTATUS status = RedSalamander::Win32Callback::InvokeC5039Suppressed([&]() noexcept
        { return NotifyIpInterfaceChange(AF_UNSPEC, &NetworkChangeSubscription::OnIpInterfaceChanged, this, initialNotification, &handle); });

        if (status != NO_ERROR)
        {
            Debug::Warning(L"FolderWindow: NotifyIpInterfaceChange failed (status={})", status);
            return;
        }

        _handle.reset(handle);
    }

    NetworkChangeSubscription(const NetworkChangeSubscription&)            = delete;
    NetworkChangeSubscription& operator=(const NetworkChangeSubscription&) = delete;
    NetworkChangeSubscription(NetworkChangeSubscription&&)                 = delete;
    NetworkChangeSubscription& operator=(NetworkChangeSubscription&&)      = delete;

    ~NetworkChangeSubscription() = default;

private:
    static void CALLBACK OnIpInterfaceChanged(PVOID callerContext, PMIB_IPINTERFACE_ROW /*row*/, MIB_NOTIFICATION_TYPE notificationType) noexcept
    {
        if (notificationType == MibInitialNotification)
        {
            return;
        }

        auto* self = static_cast<NetworkChangeSubscription*>(callerContext);
        if (! self || ! self->_hwnd)
        {
            return;
        }

        PostMessageW(self->_hwnd, WndMsg::kNetworkConnectivityChanged, 0, 0);
    }

    struct NetworkHandleDeleter
    {
        void operator()(HANDLE handle) const noexcept
        {
            if (handle)
            {
                CancelMibChangeNotify2(handle);
            }
        }
    };

    HWND _hwnd = nullptr;
    std::unique_ptr<void, NetworkHandleDeleter> _handle;
};

FolderWindow::FolderWindow()
{
    _viewerCallback.owner = this;
}

FolderWindow::~FolderWindow()
{
    Destroy();
}

void FolderWindow::SetSettings(Common::Settings::Settings* settings) noexcept
{
#ifdef ENABLE_TESTS
    if (IsFolderWindowSelfTestTracingEnabled())
    {
        SelfTest::AppendSelfTestTrace(L"FolderWindow::SetSettings: begin");
    }
#endif
    _settings = settings;
#ifdef ENABLE_TESTS
    if (IsFolderWindowSelfTestTracingEnabled())
    {
        SelfTest::AppendSelfTestTrace(L"FolderWindow::SetSettings: settings assigned");
    }
#endif
    _leftPane.navigationView.SetSettings(settings);
#ifdef ENABLE_TESTS
    if (IsFolderWindowSelfTestTracingEnabled())
    {
        SelfTest::AppendSelfTestTrace(L"FolderWindow::SetSettings: left navigation set");
    }
#endif
    _rightPane.navigationView.SetSettings(settings);
#ifdef ENABLE_TESTS
    if (IsFolderWindowSelfTestTracingEnabled())
    {
        SelfTest::AppendSelfTestTrace(L"FolderWindow::SetSettings: right navigation set");
    }
#endif
}

void FolderWindow::SetShortcutManager(const ShortcutManager* shortcuts) noexcept
{
#ifdef ENABLE_TESTS
    if (IsFolderWindowSelfTestTracingEnabled())
    {
        SelfTest::AppendSelfTestTrace(L"FolderWindow::SetShortcutManager: begin");
    }
#endif
    _shortcutManager = shortcuts;
#ifdef ENABLE_TESTS
    if (IsFolderWindowSelfTestTracingEnabled())
    {
        SelfTest::AppendSelfTestTrace(L"FolderWindow::SetShortcutManager: shortcut pointer assigned");
    }
#endif
    _functionBar.SetShortcutManager(shortcuts);
#ifdef ENABLE_TESTS
    if (IsFolderWindowSelfTestTracingEnabled())
    {
        SelfTest::AppendSelfTestTrace(L"FolderWindow::SetShortcutManager: function bar set");
    }
#endif
    _leftPane.folderView.SetShortcutManager(shortcuts);
#ifdef ENABLE_TESTS
    if (IsFolderWindowSelfTestTracingEnabled())
    {
        SelfTest::AppendSelfTestTrace(L"FolderWindow::SetShortcutManager: left folder view set");
    }
#endif
    _rightPane.folderView.SetShortcutManager(shortcuts);
#ifdef ENABLE_TESTS
    if (IsFolderWindowSelfTestTracingEnabled())
    {
        SelfTest::AppendSelfTestTrace(L"FolderWindow::SetShortcutManager: right folder view set");
    }
#endif
}

void FolderWindow::SetFunctionBarModifiers(uint32_t modifiers) noexcept
{
    _functionBar.SetModifiers(modifiers);
}

void FolderWindow::SetFunctionBarPressedKey(std::optional<uint32_t> vk) noexcept
{
    _functionBar.SetPressedFunctionKey(vk);
}

void FolderWindow::SetFunctionBarVisible(bool visible) noexcept
{
    const bool changed  = _functionBarVisible != visible;
    _functionBarVisible = visible;

    if (_hWnd)
    {
        CalculateLayout();
        AdjustChildWindows();
    }

    if (const HWND bar = _functionBar.GetHwnd())
    {
        if (_hWnd)
        {
            const int x = _functionBarRect.left;
            const int y = _functionBarRect.top;
            const int w = std::max(0L, _functionBarRect.right - _functionBarRect.left);
            const int h = std::max(0L, _functionBarRect.bottom - _functionBarRect.top);
            SetWindowPos(bar, nullptr, x, y, w, h, SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOOWNERZORDER);
        }

        ShowWindow(bar, _functionBarVisible ? SW_SHOWNA : SW_HIDE);
        if (_functionBarVisible)
        {
            RedrawWindow(bar, nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW);
        }
    }

    if (_hWnd)
    {
        if (changed)
        {
            InvalidateRect(_hWnd.get(), nullptr, FALSE);
        }
    }
}

#ifdef ENABLE_TESTS
bool FolderWindow::DebugGetFunctionBarSnapshot(FolderWindowFunctionBarDebugSnapshot& out) const noexcept
{
    out         = {};
    out.visible = _functionBarVisible;

    const HWND bar = _functionBar.GetHwnd();
    if (! bar || IsWindow(bar) == FALSE)
    {
        return false;
    }

    out.windowVisible              = IsWindowVisible(bar) != FALSE;
    out.rect                       = _functionBarRect;
    out.usesDirectWriteTextMetrics = _functionBar.DebugUsesDirectWriteTextMetrics();
    return true;
}
#endif

void FolderWindow::SetPanePathChangedCallback(PanePathChangedCallback callback)
{
    _panePathChangedCallback = std::move(callback);
}

void FolderWindow::SetPaneEnumerationCompletedCallback(Pane pane, FolderView::EnumerationCompletedCallback callback)
{
    PaneState& state                   = pane == Pane::Left ? _leftPane : _rightPane;
    state.enumerationCompletedCallback = std::move(callback);
    InstallFolderViewEnumerationCompletedCallback(pane);
}

void FolderWindow::InstallFolderViewEnumerationCompletedCallback(Pane pane)
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.SetEnumerationCompletedCallback([this, pane](const std::filesystem::path& folder) { OnFolderViewEnumerationCompleted(pane, folder); });
}

void FolderWindow::OnFolderViewEnumerationCompleted(Pane pane, const std::filesystem::path& folder)
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (state.pendingNavigationToPaintMetric.has_value())
    {
        const FolderWindow::PaneState::PendingNavigationToPaintMetric& pendingRef = state.pendingNavigationToPaintMetric.value();
        const bool folderMatches = OrdinalString::EqualsNoCasePath(folder, pendingRef.targetFolder) || folder.native() == pendingRef.targetFolder.native();
        if (folderMatches)
        {
            const FolderWindow::PaneState::PendingNavigationToPaintMetric pending = std::move(state.pendingNavigationToPaintMetric.value());
            state.pendingNavigationToPaintMetric.reset();
            std::wstring enumeratedMetricName          = pending.metricName;
            constexpr std::wstring_view kToPaintSuffix = L"_to_paint_us";
            if (enumeratedMetricName.size() >= kToPaintSuffix.size() &&
                std::wstring_view(enumeratedMetricName).substr(enumeratedMetricName.size() - kToPaintSuffix.size()) == kToPaintSuffix)
            {
                enumeratedMetricName.replace(enumeratedMetricName.size() - kToPaintSuffix.size(), kToPaintSuffix.size(), L"_enumerated_us");
            }
            else
            {
                enumeratedMetricName.append(L".enumerated_us");
            }
            Debug::Perf::Emit(enumeratedMetricName, pending.detail, Debug::Perf::ElapsedUs(pending.startedAt), pending.value0, pending.value1, S_OK);
            state.folderView.RecordPendingInputToPaintStart(pending.startedAt, pending.metricName, pending.detail, pending.value0, pending.value1);
        }
    }

    if (state.enumerationCompletedCallback)
    {
        state.enumerationCompletedCallback(folder);
    }
}

void FolderWindow::SetPaneDetailsTextProvider(Pane pane, FolderView::DetailsTextProvider provider)
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.SetDetailsTextProvider(std::move(provider));
}

void FolderWindow::SetPaneMetadataTextProvider(Pane pane, FolderView::MetadataTextProvider provider)
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.SetMetadataTextProvider(std::move(provider));
}

void FolderWindow::SetPaneEmptyStateMessage(Pane pane, std::wstring message)
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.SetEmptyStateMessage(std::move(message));
}

void FolderWindow::SetPaneBackgroundWatermark(Pane pane, std::wstring message, bool animated)
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.SetBackgroundWatermark(std::move(message), animated);
}

void FolderWindow::RefreshPaneDetailsText(Pane pane)
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.RefreshDetailsText();
}

void FolderWindow::SetPaneSelectionByDisplayNamePredicate(Pane pane,
                                                          const std::function<bool(std::wstring_view)>& shouldSelect,
                                                          bool clearExistingSelection) noexcept
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.SetSelectionByDisplayNamePredicate(shouldSelect, clearExistingSelection);
}

void FolderWindow::ClearPaneSelectionByDisplayNamePredicate(Pane pane, const std::function<bool(std::wstring_view)>& shouldUnselect) noexcept
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.ClearSelectionByDisplayNamePredicate(shouldUnselect);
}

uint64_t FolderWindow::AddFileOperationCompletedCallback(FileOperationCompletedCallback callback, std::weak_ptr<void> lifetimeGuard)
{
    const uint64_t token = _nextFileOperationCompletedCallbackToken++;
    const std::weak_ptr<void> emptyLifetimeGuard;
    const bool hasLifetimeGuard =
        ! lifetimeGuard.expired() || lifetimeGuard.owner_before(emptyLifetimeGuard) || emptyLifetimeGuard.owner_before(lifetimeGuard);
    _fileOperationCompletedCallbacks.push_back(
        FileOperationCompletedSubscription{.token = token,
                                           .callback = std::move(callback),
                                           .lifetimeGuard = std::move(lifetimeGuard),
                                           .hasLifetimeGuard = hasLifetimeGuard});
    return token;
}

void FolderWindow::RemoveFileOperationCompletedCallback(uint64_t token) noexcept
{
    std::erase_if(_fileOperationCompletedCallbacks,
                  [token](const FileOperationCompletedSubscription& subscription) noexcept { return subscription.token == token; });
}

ATOM FolderWindow::RegisterWndClass(HINSTANCE instance)
{
    static ATOM atom = 0;
    if (atom)
        return atom;

    WNDCLASSEXW wc{};
    wc.cbSize        = sizeof(wc);
    wc.style         = CS_HREDRAW | CS_VREDRAW | CS_DBLCLKS;
    wc.lpfnWndProc   = WndProcThunk;
    wc.hInstance     = instance;
    wc.hCursor       = LoadCursor(nullptr, IDC_ARROW);
    wc.hbrBackground = nullptr; // Custom painting
    wc.lpszClassName = kClassName;

    atom = RegisterClassExW(&wc);
    return atom;
}

HWND FolderWindow::Create(HWND parent, int x, int y, int width, int height)
{
    _hInstance = GetModuleHandle(nullptr);

    if (! RegisterWndClass(_hInstance))
    {
        return nullptr;
    }

    _clientSize = {static_cast<LONG>(width), static_cast<LONG>(height)};

    const HWND hwnd =
        CreateWindowExW(0, kClassName, L"", WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS, x, y, width, height, parent, nullptr, _hInstance, this);
    if (! hwnd)
    {
        return nullptr;
    }

    _backgroundBrush.reset(CreateSolidBrush(_theme.windowBackground));
    _splitterBrush.reset(CreateSolidBrush(_theme.menu.separator));
    _splitterGripBrush.reset(CreateSolidBrush(SplitterGripColor(_theme)));
    _splitterArrowHoverBrush.reset(CreateSolidBrush(_theme.menu.selectionBg));

    return hwnd;
}

void FolderWindow::Destroy()
{
#ifdef ENABLE_TESTS
    if (IsFolderWindowSelfTestTracingEnabled())
    {
        SelfTest::AppendSelfTestTrace(L"FolderWindow::Destroy: begin");
    }
#endif
    ShutdownFileOperations();

    CancelSelectionSizeComputation(Pane::Left);
    CancelSelectionSizeComputation(Pane::Right);

    if (_leftPane.selectionSizeThread.joinable())
    {
        _leftPane.selectionSizeThread.request_stop();
        _leftPane.selectionSizeCv.notify_all();
        _leftPane.selectionSizeThread = std::jthread{};
    }

    if (_rightPane.selectionSizeThread.joinable())
    {
        _rightPane.selectionSizeThread.request_stop();
        _rightPane.selectionSizeCv.notify_all();
        _rightPane.selectionSizeThread = std::jthread{};
    }

    CloseSharedDirectoriesDialog();
    CloseOpenedFilesDialog();
#ifdef ENABLE_TESTS
    if (IsFolderWindowSelfTestTracingEnabled())
    {
        SelfTest::AppendSelfTestTrace(L"FolderWindow::Destroy: dialogs closed");
    }
#endif

    _backgroundBrush.reset();
    _splitterBrush.reset();
    _splitterGripBrush.reset();
    _splitterArrowHoverBrush.reset();

    DestroyCommandLineControls();

    _functionBar.Destroy();
#ifdef ENABLE_TESTS
    if (IsFolderWindowSelfTestTracingEnabled())
    {
        SelfTest::AppendSelfTestTrace(L"FolderWindow::Destroy: function bar destroyed");
    }
#endif

    if (_leftPane.hNavigationView)
    {
        _leftPane.navigationView.Destroy();
        _leftPane.hNavigationView.reset();
    }
#ifdef ENABLE_TESTS
    if (IsFolderWindowSelfTestTracingEnabled())
    {
        SelfTest::AppendSelfTestTrace(L"FolderWindow::Destroy: left navigation destroyed");
    }
#endif

    if (_leftPane.hFolderView)
    {
        _leftPane.folderView.Destroy();
        _leftPane.hFolderView.reset();
    }
#ifdef ENABLE_TESTS
    if (IsFolderWindowSelfTestTracingEnabled())
    {
        SelfTest::AppendSelfTestTrace(L"FolderWindow::Destroy: left folder view destroyed");
    }
#endif

    _leftPane.filterBarHost.Detach();
    _leftPane.filterBarLabel  = nullptr;
    _leftPane.filterBarCombo  = nullptr;
    _leftPane.filterBarToggle = nullptr;
    _leftPane.hFilterBar.reset();
    _leftPane.previewTabsHost.Detach();
    _leftPane.previewTabsControl = nullptr;
    _leftPane.hPreviewTabs.reset();
    _leftPane.previewContentHost.Detach();
    _leftPane.previewContentLabel     = nullptr;
    _leftPane.previewPropertiesScroll = nullptr;
    _leftPane.previewPropertiesSections.clear();
    _leftPane.hPreviewContent.reset();
    _leftPane.hStatusBar.reset();
#ifdef ENABLE_TESTS
    if (IsFolderWindowSelfTestTracingEnabled())
    {
        SelfTest::AppendSelfTestTrace(L"FolderWindow::Destroy: left chrome reset");
    }
#endif

    if (_rightPane.hNavigationView)
    {
        _rightPane.navigationView.Destroy();
        _rightPane.hNavigationView.reset();
    }
#ifdef ENABLE_TESTS
    if (IsFolderWindowSelfTestTracingEnabled())
    {
        SelfTest::AppendSelfTestTrace(L"FolderWindow::Destroy: right navigation destroyed");
    }
#endif

    if (_rightPane.hFolderView)
    {
        _rightPane.folderView.Destroy();
        _rightPane.hFolderView.reset();
    }
#ifdef ENABLE_TESTS
    if (IsFolderWindowSelfTestTracingEnabled())
    {
        SelfTest::AppendSelfTestTrace(L"FolderWindow::Destroy: right folder view destroyed");
    }
#endif

    _rightPane.filterBarHost.Detach();
    _rightPane.filterBarLabel  = nullptr;
    _rightPane.filterBarCombo  = nullptr;
    _rightPane.filterBarToggle = nullptr;
    _rightPane.hFilterBar.reset();
    _rightPane.previewTabsHost.Detach();
    _rightPane.previewTabsControl = nullptr;
    _rightPane.hPreviewTabs.reset();
    _rightPane.previewContentHost.Detach();
    _rightPane.previewContentLabel     = nullptr;
    _rightPane.previewPropertiesScroll = nullptr;
    _rightPane.previewPropertiesSections.clear();
    _rightPane.hPreviewContent.reset();
    _rightPane.hStatusBar.reset();
#ifdef ENABLE_TESTS
    if (IsFolderWindowSelfTestTracingEnabled())
    {
        SelfTest::AppendSelfTestTrace(L"FolderWindow::Destroy: right chrome reset");
    }
#endif

    if (_leftPane.fileSystem)
    {
        DirectoryInfoCache::GetInstance().UnregisterProvider(_leftPane.fileSystem.get());
    }
    if (_rightPane.fileSystem)
    {
        DirectoryInfoCache::GetInstance().UnregisterProvider(_rightPane.fileSystem.get());
    }

    _leftPane.fileSystem = nullptr;
    _leftPane.fileSystemModule.reset();
    _leftPane.pluginId.clear();
    _leftPane.currentPath.reset();
    _leftPane.updatingPath = false;

    _rightPane.fileSystem = nullptr;
    _rightPane.fileSystemModule.reset();
    _rightPane.pluginId.clear();
    _rightPane.currentPath.reset();
    _rightPane.updatingPath = false;

#ifdef ENABLE_TESTS
    if (IsFolderWindowSelfTestTracingEnabled())
    {
        SelfTest::AppendSelfTestTrace(L"FolderWindow::Destroy: resetting hwnd");
    }
#endif
    _hWnd.reset();
    _settings = nullptr;
#ifdef ENABLE_TESTS
    if (IsFolderWindowSelfTestTracingEnabled())
    {
        SelfTest::AppendSelfTestTrace(L"FolderWindow::Destroy: complete");
    }
#endif
}

LRESULT CALLBACK FolderWindow::WndProcThunk(HWND hWindow, UINT msg, WPARAM wp, LPARAM lp)
{
    FolderWindow* self = nullptr;

    if (msg == WM_NCCREATE)
    {
        auto cs = reinterpret_cast<CREATESTRUCTW*>(lp);
        self    = reinterpret_cast<FolderWindow*>(cs->lpCreateParams);
        SetWindowLongPtrW(hWindow, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
        self->_hWnd.reset(hWindow);
        InitPostedPayloadWindow(hWindow);
    }
    else
    {
        self = reinterpret_cast<FolderWindow*>(GetWindowLongPtrW(hWindow, GWLP_USERDATA));
    }

    if (self)
    {
        return self->WndProc(hWindow, msg, wp, lp);
    }

    return DefWindowProcW(hWindow, msg, wp, lp);
}

LRESULT FolderWindow::WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp)
{
    switch (msg)
    {
        case WM_CREATE: return OnCreate(hwnd) ? 0 : -1;
        case WM_DESTROY: OnDestroy(); return 0;
        case WM_NCDESTROY: static_cast<void>(DrainPostedPayloadsForWindow(hwnd)); break;
        case WM_SIZE: OnSize(LOWORD(lp), HIWORD(lp)); return 0;
        case WM_TIMER:
            if (static_cast<UINT_PTR>(wp) == kPreviewPaneRefreshTimerId)
            {
                OnPreviewPaneRefreshTimer();
                return 0;
            }
            break;
        case WM_SETFOCUS: OnSetFocus(); return 0;
        case WM_DEVICECHANGE: return OnDeviceChange(static_cast<UINT>(wp), lp);
        case WndMsg::kNetworkConnectivityChanged: OnNetworkConnectivityChanged(); return 0;
        case WM_ERASEBKGND: return 1; // no erase background
        case WM_PAINT: OnPaint(); return 0;
        case WM_DRAWITEM: return OnDrawItem(reinterpret_cast<DRAWITEMSTRUCT*>(lp));
        case WM_LBUTTONDOWN: OnLButtonDown({GET_X_LPARAM(lp), GET_Y_LPARAM(lp)}); return 0;
        case WM_LBUTTONDBLCLK: OnLButtonDblClk({GET_X_LPARAM(lp), GET_Y_LPARAM(lp)}); return 0;
        case WM_LBUTTONUP: OnLButtonUp(); return 0;
        case WM_MOUSEMOVE: OnMouseMove({GET_X_LPARAM(lp), GET_Y_LPARAM(lp)}); return 0;
        case WM_MOUSELEAVE: OnMouseLeave(); return 0;
        case WM_CAPTURECHANGED: OnCaptureChanged(); return 0;
        case WM_PARENTNOTIFY: OnParentNotify(LOWORD(wp), HIWORD(wp)); return 0;
        case WM_NOTIFY: return OnNotify(lp);
        case WM_SETCURSOR: return OnSetCursor(reinterpret_cast<HWND>(wp), LOWORD(lp), HIWORD(lp));
        case WndMsg::kPaneFocusChanged: UpdatePaneFocusStates(); return 0;
        case WndMsg::kPaneRestoreFolderFocus: static_cast<void>(TryRestoreActivePaneFolderViewFocus()); return 0;
        case WndMsg::kPaneSelectionSizeComputed: return OnPaneSelectionSizeComputed(lp);
        case WndMsg::kPaneSelectionSizeProgress: return OnPaneSelectionSizeProgress(lp);
        case WndMsg::kFileOperationCompleted: return OnFileOperationCompleted(lp);
        case WndMsg::kChangeCaseTaskUpdate: return OnChangeCaseTaskUpdate(lp);
        case WndMsg::kChangeCaseCompleted: return OnChangeCaseCompleted(lp);
        case WndMsg::kChangeAttributesTaskUpdate: return OnChangeAttributesTaskUpdate(lp);
        case WndMsg::kChangeAttributesCompleted: return OnChangeAttributesCompleted(lp);
        case WndMsg::kFolderWindowCloseOpenedFilesDialog: CloseOpenedFilesDialog(); return 0;
        case WndMsg::kFolderWindowCloseSharedDirectoriesDialog: CloseSharedDirectoriesDialog(); return 0;
        case WndMsg::kHostShowAlert: return OnHostServicesMessage(msg, wp, lp);
        case WndMsg::kHostClearAlert: return OnHostServicesMessage(msg, wp, lp);
        case WndMsg::kHostShowPrompt: return OnHostServicesMessage(msg, wp, lp);
        case WndMsg::kHostShowConnectionManager: return OnHostServicesMessage(msg, wp, lp);
        case WndMsg::kHostGetConnectionJsonUtf8: return OnHostServicesMessage(msg, wp, lp);
        case WndMsg::kHostGetConnectionSecret: return OnHostServicesMessage(msg, wp, lp);
        case WndMsg::kHostPromptConnectionSecret: return OnHostServicesMessage(msg, wp, lp);
        case WndMsg::kHostClearCachedConnectionSecret: return OnHostServicesMessage(msg, wp, lp);
        case WndMsg::kHostUpgradeFtpAnonymousToPassword: return OnHostServicesMessage(msg, wp, lp);
        case WndMsg::kHostExecuteInPane: return OnHostServicesMessage(msg, wp, lp);
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

LRESULT FolderWindow::OnDeviceChange(UINT event, LPARAM data) noexcept
{
    if (event != DBT_DEVICEARRIVAL && event != DBT_DEVICEREMOVECOMPLETE)
    {
        return static_cast<LRESULT>(TRUE);
    }

    const auto* hdr = reinterpret_cast<const DEV_BROADCAST_HDR*>(data);
    if (! hdr || hdr->dbch_devicetype != DBT_DEVTYP_VOLUME)
    {
        return static_cast<LRESULT>(TRUE);
    }

    const auto* volume   = reinterpret_cast<const DEV_BROADCAST_VOLUME*>(hdr);
    const DWORD unitmask = volume->dbcv_unitmask;
    if (unitmask == 0)
    {
        return static_cast<LRESULT>(TRUE);
    }

    auto refreshIfAffected = [&](PaneState& pane)
    {
        const std::optional<std::filesystem::path>& current = pane.currentPath;
        if (! current.has_value())
        {
            return;
        }

        const auto driveLetter = NavigationLocation::TryGetWindowsDriveLetter(current.value());
        if (! driveLetter.has_value() || ! NavigationLocation::DriveMaskContainsLetter(static_cast<uint32_t>(unitmask), driveLetter.value()))
        {
            return;
        }

        pane.folderView.ForceRefresh();
    };

    refreshIfAffected(_leftPane);
    refreshIfAffected(_rightPane);
    return static_cast<LRESULT>(TRUE);
}

void FolderWindow::OnNetworkConnectivityChanged() noexcept
{
    const uint64_t now             = GetTickCount64();
    constexpr uint64_t kDebounceMs = 500;
    if (_lastNetworkConnectivityRefreshTick != 0 && now - _lastNetworkConnectivityRefreshTick < kDebounceMs)
    {
        return;
    }
    _lastNetworkConnectivityRefreshTick = now;

    auto refreshIfNetworkPath = [&](PaneState& pane)
    {
        if (! pane.hFolderView)
        {
            return;
        }

        if (! NavigationLocation::IsFilePluginShortId(pane.pluginShortId))
        {
            return;
        }

        if (! pane.currentPath.has_value())
        {
            return;
        }

        const std::wstring_view pathText = pane.currentPath.value().native();
        if (NavigationLocation::LooksLikeUncPath(pathText))
        {
            pane.folderView.ForceRefresh();
            return;
        }

        const auto driveLetter = NavigationLocation::TryGetWindowsDriveLetter(pathText);
        if (! driveLetter.has_value())
        {
            return;
        }

        std::wstring driveRoot;
        driveRoot.push_back(driveLetter.value());
        driveRoot.append(L":\\");

        const UINT driveType = GetDriveTypeW(driveRoot.c_str());
        if (driveType == DRIVE_REMOTE)
        {
            pane.folderView.ForceRefresh();
        }
    };

    refreshIfNetworkPath(_leftPane);
    refreshIfNetworkPath(_rightPane);
}

LRESULT FolderWindow::OnDrawItem(DRAWITEMSTRUCT* dis)
{
    if (! _hWnd)
    {
        return 0;
    }

    const WPARAM controlId = dis ? static_cast<WPARAM>(dis->CtlID) : 0;
    return DefWindowProcW(_hWnd.get(), WM_DRAWITEM, controlId, reinterpret_cast<LPARAM>(dis));
}

LRESULT FolderWindow::OnNotify(LPARAM data)
{
    const auto* header = reinterpret_cast<const FolderWindowNotifyHeader*>(data);
    if (header && header->code == kStatusBarSortClickNotification && (header->idFrom == kLeftStatusBarId || header->idFrom == kRightStatusBarId))
    {
        const auto* mouse = reinterpret_cast<const StatusBarSortClickNotification*>(header);
        if (mouse->part == kStatusBarPartSort && _showSortMenuCallback)
        {
            const Pane pane = header->idFrom == kLeftStatusBarId ? Pane::Left : Pane::Right;
            SetActivePane(pane);

            RECT partRect{};
            POINT screenPoint{};
            if (GetStatusBarPartRect(header->hwndFrom, kStatusBarPartSort, partRect))
            {
                screenPoint = {partRect.right, partRect.top};
            }
            else
            {
                screenPoint = mouse->clientPoint;
            }
            ClientToScreen(header->hwndFrom, &screenPoint);
            _showSortMenuCallback(pane, screenPoint);
            return 0;
        }
    }

    if (! _hWnd)
    {
        return 0;
    }
    const WPARAM controlId = header ? static_cast<WPARAM>(header->idFrom) : 0;
    return DefWindowProcW(_hWnd.get(), WM_NOTIFY, controlId, data);
}

LRESULT FolderWindow::HandlePaneDxHostMessage(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, bool& handled) noexcept
{
    if (_hCommandLineHost && hwnd == _hCommandLineHost.get())
    {
        if (msg == WM_NCDESTROY)
        {
            handled = true;
            _commandLineHost.ReleaseMouseCapture();
            _commandLineHost.Detach();
            _commandLineLabel = nullptr;
            _commandLineField = nullptr;
            _hCommandLineHost.release();
            return 0;
        }

        const LRESULT result = _commandLineHost.HandleMessage(hwnd, msg, wp, lp, handled);
        if (msg == WM_SIZE)
        {
            UpdateCommandLineHostLayout();
        }
        return handled ? result : 0;
    }

    const auto dispatch = [&](Pane pane,
                              wil::unique_hwnd& expectedHwnd,
                              RedSalamander::DxUi::WindowHost& host,
                              RedSalamander::DxUi::Label** label,
                              RedSalamander::DxUi::TabControl** tabs,
                              bool previewContent,
                              bool filterBar) noexcept -> std::optional<LRESULT>
    {
        if (! expectedHwnd || hwnd != expectedHwnd.get())
        {
            return std::nullopt;
        }

        if (msg == WM_NCDESTROY)
        {
            handled = true;
            host.ReleaseMouseCapture();
            host.Detach();
            if (label)
            {
                *label = nullptr;
            }
            if (filterBar)
            {
                PaneState& state      = pane == Pane::Left ? _leftPane : _rightPane;
                state.filterBarCombo  = nullptr;
                state.filterBarToggle = nullptr;
            }
            if (tabs)
            {
                *tabs = nullptr;
            }
            expectedHwnd.release();
            return 0;
        }

        const LRESULT result = host.HandleMessage(hwnd, msg, wp, lp, handled);
        if (msg == WM_SIZE)
        {
            if (previewContent)
            {
                UpdatePreviewContentLayout(pane);
                LayoutEmbeddedPreviewViewer(pane);
            }
            else if (filterBar)
            {
                UpdateFilterBarLayout(pane);
            }
        }
        return handled ? std::optional<LRESULT>{result} : std::nullopt;
    };

    if (const auto result = dispatch(Pane::Left, _leftPane.hFilterBar, _leftPane.filterBarHost, &_leftPane.filterBarLabel, nullptr, false, true))
    {
        return result.value();
    }
    if (const auto result = dispatch(Pane::Right, _rightPane.hFilterBar, _rightPane.filterBarHost, &_rightPane.filterBarLabel, nullptr, false, true))
    {
        return result.value();
    }
    if (const auto result = dispatch(Pane::Left, _leftPane.hPreviewTabs, _leftPane.previewTabsHost, nullptr, &_leftPane.previewTabsControl, false, false))
    {
        return result.value();
    }
    if (const auto result = dispatch(Pane::Right, _rightPane.hPreviewTabs, _rightPane.previewTabsHost, nullptr, &_rightPane.previewTabsControl, false, false))
    {
        return result.value();
    }
    if (const auto result = dispatch(Pane::Left, _leftPane.hPreviewContent, _leftPane.previewContentHost, &_leftPane.previewContentLabel, nullptr, true, false))
    {
        return result.value();
    }
    if (const auto result =
            dispatch(Pane::Right, _rightPane.hPreviewContent, _rightPane.previewContentHost, &_rightPane.previewContentLabel, nullptr, true, false))
    {
        return result.value();
    }

    handled = false;
    return 0;
}

LRESULT CALLBACK FolderWindowDxHostWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    auto* self = reinterpret_cast<FolderWindow*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (msg == WM_NCCREATE)
    {
        const auto* createStruct = reinterpret_cast<const CREATESTRUCTW*>(lp);
        self                     = createStruct ? static_cast<FolderWindow*>(createStruct->lpCreateParams) : nullptr;
        if (self)
        {
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
        }
    }

    if (self)
    {
        bool handled         = false;
        const LRESULT result = self->HandlePaneDxHostMessage(hwnd, msg, wp, lp, handled);
        if (handled)
        {
            if (msg == WM_NCDESTROY)
            {
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
            }
            return result;
        }
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

bool FolderWindow::OnCreate(HWND hwnd) noexcept
{
    // _hWnd not yet initialize
    {
        Debug::Perf::Scope perf(L"FolderWindow.OnCreate.GetDpiForWindow");
        _dpi = GetDpiForWindow(hwnd);
    }

    {
        Debug::Perf::Scope perf(L"FolderWindow.OnCreate.EnsureFileOperations");
        EnsureFileOperations();
    }

    if (FAILED(EnsureFolderWindowStatusBarClass(_hInstance)))
    {
        Debug::Error(L"FolderWindow::OnCreate failed to register status-bar class.");
        return false;
    }
    if (! EnsureFolderWindowDxHostClass(_hInstance))
    {
        Debug::Error(L"FolderWindow::OnCreate failed to register DxUi host class.");
        return false;
    }

    {
        Debug::Perf::Scope perf(L"FolderWindow.OnCreate.CalculateLayout");
        CalculateLayout();
    }

    auto createPane = [&](Pane pane,
                          PaneState& state,
                          const RECT& navRect,
                          const RECT& filterRect,
                          const RECT& folderRect,
                          const RECT& statusRect,
                          const RECT& previewTabsRect,
                          const RECT& previewContentRect,
                          UINT_PTR navId,
                          UINT_PTR filterId,
                          UINT_PTR folderId,
                          UINT_PTR statusId,
                          UINT_PTR previewTabsId,
                          UINT_PTR previewContentId) -> bool
    {
        const std::wstring_view paneName = pane == Pane::Left ? L"Left" : L"Right";

        {
            Debug::Perf::Scope perf(L"FolderWindow.OnCreate.CreatePane.NavigationView.Create");
            perf.SetDetail(paneName);
            state.hNavigationView.reset(
                state.navigationView.Create(hwnd, navRect.left, navRect.top, navRect.right - navRect.left, navRect.bottom - navRect.top));
        }
        if (! state.hNavigationView)
        {
            return false;
        }
        SetWindowLongPtrW(state.hNavigationView.get(), GWLP_ID, static_cast<LONG_PTR>(navId));

        const int filterWidth   = std::max(0L, filterRect.right - filterRect.left);
        const int filterHeight  = std::max(0L, filterRect.bottom - filterRect.top);
        const DWORD filterStyle = WS_CHILD | WS_CLIPCHILDREN | WS_CLIPSIBLINGS;
        {
            Debug::Perf::Scope perf(L"FolderWindow.OnCreate.CreatePane.FilterBar.CreateDxHost");
            perf.SetDetail(paneName);
            state.hFilterBar.reset(CreateWindowExW(0,
                                                   kFolderWindowDxHostClassName,
                                                   nullptr,
                                                   filterStyle,
                                                   filterRect.left,
                                                   filterRect.top,
                                                   filterWidth,
                                                   filterHeight,
                                                   hwnd,
                                                   reinterpret_cast<HMENU>(filterId),
                                                   _hInstance,
                                                   this));
        }
        if (! state.hFilterBar || ! state.filterBarHost.Attach(state.hFilterBar.get()))
        {
            return false;
        }
        {
            auto root = std::make_unique<RedSalamander::DxUi::Panel>();

            state.filterBarCombo = root->AddChild<RedSalamander::DxUi::ComboBox>();
            state.filterBarCombo->SetEditable(true);
            state.filterBarCombo->SetVariant(RedSalamander::DxUi::ComboBoxVariant::Edit);
            state.filterBarCombo->SetAutoOpenOnTextInput(false);
            state.filterBarCombo->SetMaxVisibleItems(MaskSyntax::kWildcardMaskHistoryMaxItems);
            state.filterBarCombo->SetPlaceholder(LoadStringResource(nullptr, IDS_LABEL_PANE_FILTER));
            state.filterBarCombo->SetAccessibleName(LoadStringResource(nullptr, IDS_LABEL_PANE_FILTER));
            state.filterBarCombo->SetOnTextChanged([this, pane](std::wstring_view text) noexcept { OnFilterBarTextChanged(pane, text); });
            state.filterBarCombo->SetOnSelectionChanged([this, pane](size_t) noexcept { OnFilterBarSubmitted(pane); });
            state.filterBarCombo->SetOnSubmitted([this, pane] { OnFilterBarSubmitted(pane); });
            state.filterBarCombo->SetOnPopupRequested([this, pane] { return ShowFilterBarHistoryMenu(pane); });

            state.filterBarToggle = root->AddChild<RedSalamander::DxUi::Toggle>();
            state.filterBarToggle->SetAccessibleName(LoadStringResource(nullptr, IDS_LABEL_PANE_FILTER_USE_FILTER));
            state.filterBarToggle->SetStateLabels(LoadStringResource(nullptr, IDS_PREFS_COMMON_OFF), LoadStringResource(nullptr, IDS_PREFS_COMMON_ON));
            state.filterBarToggle->SetOnToggled([this, pane](bool checked) noexcept { OnFilterBarToggled(pane, checked); });

            state.filterBarHost.SetRoot(std::move(root));
            state.filterBarHost.SetTheme(MakeAppThemeDxPalette(_theme, _theme.windowBackground));
            RefreshFilterBarHistoryItems(pane);
            UpdateFilterBarLayout(pane);
        }

        {
            Debug::Perf::Scope perf(L"FolderWindow.OnCreate.CreatePane.FolderView.Create");
            perf.SetDetail(paneName);
            state.hFolderView.reset(
                state.folderView.Create(hwnd, folderRect.left, folderRect.top, folderRect.right - folderRect.left, folderRect.bottom - folderRect.top));
        }
        if (! state.hFolderView)
        {
            return false;
        }
        SetWindowLongPtrW(state.hFolderView.get(), GWLP_ID, static_cast<LONG_PTR>(folderId));

        const int statusWidth   = std::max(0L, statusRect.right - statusRect.left);
        const int statusHeight  = std::max(0L, statusRect.bottom - statusRect.top);
        const DWORD statusStyle = WS_CHILD | WS_VISIBLE;
        {
            Debug::Perf::Scope perf(L"FolderWindow.OnCreate.CreatePane.StatusBar.CreateWindowExW");
            perf.SetDetail(paneName);
            state.hStatusBar.reset(CreateWindowExW(0,
                                                   kFolderWindowStatusBarClassName,
                                                   nullptr,
                                                   statusStyle,
                                                   statusRect.left,
                                                   statusRect.top,
                                                   statusWidth,
                                                   statusHeight,
                                                   hwnd,
                                                   reinterpret_cast<HMENU>(statusId),
                                                   _hInstance,
                                                   nullptr));
        }
        if (! state.hStatusBar)
        {
            return false;
        }
        {
            Debug::Perf::Scope perf(L"FolderWindow.OnCreate.CreatePane.StatusBar.Initialize");
            perf.SetDetail(paneName);
            SetPropW(state.hStatusBar.get(), kStatusBarOwnerProp, reinterpret_cast<HANDLE>(this));
            SetPropW(state.hStatusBar.get(), kStatusBarSelectionTextProp, reinterpret_cast<HANDLE>(&state.statusSelectionText));
            SetPropW(state.hStatusBar.get(), kStatusBarSecurityTextProp, reinterpret_cast<HANDLE>(&state.statusSecurityText));
            SetPropW(state.hStatusBar.get(), kStatusBarSortTextProp, reinterpret_cast<HANDLE>(&state.statusSortText));
            SetPropW(state.hStatusBar.get(), kStatusBarFocusHueProp, reinterpret_cast<HANDLE>(&state.statusFocusHueDegrees));
        }

        const int previewTabsWidth   = std::max(0L, previewTabsRect.right - previewTabsRect.left);
        const int previewTabsHeight  = std::max(0L, previewTabsRect.bottom - previewTabsRect.top);
        const DWORD previewTabsStyle = WS_CHILD | WS_CLIPCHILDREN | WS_CLIPSIBLINGS | WS_TABSTOP;
        {
            Debug::Perf::Scope perf(L"FolderWindow.OnCreate.CreatePane.PreviewTabs.CreateDxHost");
            perf.SetDetail(paneName);
            state.hPreviewTabs.reset(CreateWindowExW(0,
                                                     kFolderWindowDxHostClassName,
                                                     nullptr,
                                                     previewTabsStyle,
                                                     previewTabsRect.left,
                                                     previewTabsRect.top,
                                                     previewTabsWidth,
                                                     previewTabsHeight,
                                                     hwnd,
                                                     reinterpret_cast<HMENU>(previewTabsId),
                                                     _hInstance,
                                                     this));
        }
        if (! state.hPreviewTabs || ! state.previewTabsHost.Attach(state.hPreviewTabs.get()))
        {
            return false;
        }

        {
            std::wstring folderTabText  = LoadStringResource(nullptr, IDS_PREVIEW_TAB_FOLDER);
            std::wstring previewTabText = LoadStringResource(nullptr, IDS_PREVIEW_TAB_PREVIEW);
            auto tabs                   = std::make_unique<RedSalamander::DxUi::TabControl>();
            state.previewTabsControl    = tabs.get();
            tabs->SetFocusable(false);
            tabs->AddTab<RedSalamander::DxUi::Panel>(std::move(folderTabText));
            tabs->AddTab<RedSalamander::DxUi::Panel>(std::move(previewTabText));
            tabs->SetTabClosable(1u, true);
            tabs->SetSelectedIndex(0u);
            tabs->SetOnSelectionChanged([this, pane](size_t index) noexcept { SetPreviewPaneTab(pane, index == 1u); });
            tabs->SetOnTabCloseRequested([this](size_t index) noexcept
            {
                if (index != 1u)
                {
                    return false;
                }

                ClosePreviewPane();
                return true;
            });
            state.previewTabsHost.SetRoot(std::move(tabs));
            state.previewTabsHost.SetTheme(MakeAppThemeDxPalette(_theme, _theme.windowBackground));
        }

        const int previewContentWidth   = std::max(0L, previewContentRect.right - previewContentRect.left);
        const int previewContentHeight  = std::max(0L, previewContentRect.bottom - previewContentRect.top);
        const DWORD previewContentStyle = WS_CHILD | WS_CLIPCHILDREN | WS_CLIPSIBLINGS;
        {
            Debug::Perf::Scope perf(L"FolderWindow.OnCreate.CreatePane.PreviewContent.CreateDxHost");
            perf.SetDetail(paneName);
            state.hPreviewContent.reset(CreateWindowExW(0,
                                                        kFolderWindowDxHostClassName,
                                                        nullptr,
                                                        previewContentStyle,
                                                        previewContentRect.left,
                                                        previewContentRect.top,
                                                        previewContentWidth,
                                                        previewContentHeight,
                                                        hwnd,
                                                        reinterpret_cast<HMENU>(previewContentId),
                                                        _hInstance,
                                                        this));
        }
        if (! state.hPreviewContent || ! state.previewContentHost.Attach(state.hPreviewContent.get()))
        {
            return false;
        }
        {
            auto root                 = std::make_unique<RedSalamander::DxUi::Panel>();
            state.previewContentLabel = root->AddChild<RedSalamander::DxUi::Label>(LoadStringResource(nullptr, IDS_PREVIEW_EMPTY));
            state.previewContentLabel->SetFontRole(RedSalamander::DxUi::FontRole::Body);
            state.previewContentLabel->SetMultiline(true);
            state.previewPropertiesScroll = root->AddChild<RedSalamander::DxUi::ScrollPanel>();
            state.previewPropertiesScroll->SetScrollStepDip(48.0f);
            state.previewPropertiesScroll->SetVisible(false);
            state.previewContentHost.SetRoot(std::move(root));
            state.previewContentHost.SetTheme(MakeAppThemeDxPalette(_theme, _theme.windowBackground));
            UpdatePreviewContentLayout(pane);
        }

        {
            Debug::Perf::Scope perf(L"FolderWindow.OnCreate.CreatePane.SetFileSystem");
            perf.SetDetail(paneName);
            state.folderView.SetFileSystem(state.fileSystem);
            state.navigationView.SetFileSystem(state.fileSystem);
        }

        {
            Debug::Perf::Scope perf(L"FolderWindow.OnCreate.CreatePane.SetCallbacks");
            perf.SetDetail(paneName);
            state.navigationView.SetPathChangedCallback([this, pane](const std::optional<std::filesystem::path>& path)
            { OnNavigationPathChanged(pane, path); });
            state.navigationView.SetRequestFolderViewFocusCallback([this, pane]
            {
                SetActivePane(pane);
                FocusPaneFolderView(pane);
            });

            state.folderView.SetPathChangedCallback([this, pane](const std::optional<std::filesystem::path>& path)
            {
                OnFolderViewPathChanged(pane, path);
                if (_previewSourcePane.has_value() && _previewSourcePane.value() == pane)
                {
                    RequestPreviewPaneRefresh();
                }
            });
            state.folderView.SetDirectoryImpactCallback([this, pane](const DirectoryInfoCache::DirectoryImpact& impact) noexcept
            { OnFolderViewDirectoryImpact(pane, impact); });
            state.folderView.SetNavigateUpFromRootRequestCallback([this, pane] { OnFolderViewNavigateUpFromRoot(pane); });
            state.folderView.SetOpenFileRequestCallback([this, pane](const std::filesystem::path& path) { return TryOpenFileAsVirtualFileSystem(pane, path); });
            state.folderView.SetViewFileRequestCallback([this, pane](const FolderView::ViewFileRequest& request)
            { return TryViewFileWithViewer(pane, request); });
            state.folderView.SetFileOperationRequestCallback([this, pane](FolderView::FileOperationRequest request) noexcept -> HRESULT
            { return StartFileOperationFromFolderView(pane, std::move(request)); });
            state.folderView.SetPropertiesRequestCallback([this, pane](std::filesystem::path path) noexcept -> HRESULT
            { return ShowItemPropertiesFromFolderView(pane, std::move(path)); });
            state.folderView.SetBatchRenameRequestCallback([this, pane](std::filesystem::path targetPath, bool isDirectoryRoot)
            {
                if (targetPath.empty())
                {
                    // Empty target path = batch-rename the pane's current selection (FolderView VK_F2).
                    CommandBatchRename(pane);
                    return;
                }

                if (isDirectoryRoot)
                {
                    CommandBatchRename(pane, std::move(targetPath));
                    return;
                }

                // File from the rename prompt: seed exactly that item, independent of the live selection.
                std::vector<std::filesystem::path> initialPaths;
                initialPaths.push_back(std::move(targetPath));
                CommandBatchRename(pane, std::nullopt, std::move(initialPaths));
            });
            state.folderView.SetNavigationRequestCallback([this, pane](FolderView::NavigationRequest request)
            {
                PaneState& s = pane == Pane::Left ? _leftPane : _rightPane;
                switch (request)
                {
                    case FolderView::NavigationRequest::FocusNavigationMenu:
                        SetNavigationBarVisible(pane, true);
                        s.navigationView.SetFocusRegion(NavigationView::FocusRegion::Menu);
                        if (s.hNavigationView)
                        {
                            SetFocus(s.hNavigationView.get());
                        }
                        break;
                    case FolderView::NavigationRequest::FocusNavigationDiskInfo:
                        SetNavigationBarVisible(pane, true);
                        s.navigationView.SetFocusRegion(NavigationView::FocusRegion::DiskInfo);
                        if (s.hNavigationView)
                        {
                            SetFocus(s.hNavigationView.get());
                        }
                        break;
                    case FolderView::NavigationRequest::FocusAddressBar:
                        SetNavigationBarVisible(pane, true);
                        s.navigationView.FocusAddressBar();
                        break;
                    case FolderView::NavigationRequest::OpenHistoryDropdown:
                        SetNavigationBarVisible(pane, true);
                        s.navigationView.OpenHistoryDropdownFromKeyboard();
                        break;
                    case FolderView::NavigationRequest::SwitchPane:
                    {
                        const Pane otherPane = pane == Pane::Left ? Pane::Right : Pane::Left;
                        PaneState& other     = otherPane == Pane::Left ? _leftPane : _rightPane;
                        if (other.hFolderView)
                        {
                            SetActivePane(otherPane);
                            SetFocus(other.hFolderView.get());
                        }
                        break;
                    }
                }
            });

            state.folderView.SetSelectionChangedCallback([this, pane](const FolderView::SelectionStats& stats)
            {
                PaneState& s     = pane == Pane::Left ? _leftPane : _rightPane;
                s.selectionStats = stats;
                CancelSelectionSizeComputation(pane);
                UpdatePaneStatusBar(pane);
                if (_previewSourcePane.has_value() && _previewSourcePane.value() == pane)
                {
                    RequestPreviewPaneRefresh();
                }
            });

            state.folderView.SetFocusedItemChangedCallback([this, pane]
            {
                UpdatePaneStatusBar(pane);
                if (_previewSourcePane.has_value() && _previewSourcePane.value() == pane)
                {
                    RequestPreviewPaneRefresh();
                }
            });
            state.folderView.SetIncrementalSearchChangedCallback([this, pane] { UpdatePaneStatusBar(pane); });
            state.folderView.SetSelectionSizeComputationRequestedCallback([this, pane] { RequestSelectionSizeComputation(pane); });
            InstallFolderViewEnumerationCompletedCallback(pane);
        }

        {
            Debug::Perf::Scope perf(L"FolderWindow.OnCreate.CreatePane.StartSelectionSizeWorker");
            perf.SetDetail(paneName);
            StartSelectionSizeWorker(pane);
        }

        return true;
    };

    {
        Debug::Perf::Scope perf(L"FolderWindow.OnCreate.CreatePane");
        perf.SetDetail(L"Left");
        if (! createPane(Pane::Left,
                         _leftPane,
                         _leftNavigationRect,
                         _leftFilterBarRect,
                         _leftFolderViewRect,
                         _leftStatusBarRect,
                         _leftPreviewTabsRect,
                         _leftPreviewContentRect,
                         kLeftNavigationId,
                         kLeftFilterBarId,
                         kLeftFolderViewId,
                         kLeftStatusBarId,
                         kLeftPreviewTabsId,
                         kLeftPreviewContentId))
        {
            Debug::Error(L"FolderWindow::OnCreate failed to create left pane.");
            return false;
        }
    }

    {
        Debug::Perf::Scope perf(L"FolderWindow.OnCreate.CreatePane");
        perf.SetDetail(L"Right");
        if (! createPane(Pane::Right,
                         _rightPane,
                         _rightNavigationRect,
                         _rightFilterBarRect,
                         _rightFolderViewRect,
                         _rightStatusBarRect,
                         _rightPreviewTabsRect,
                         _rightPreviewContentRect,
                         kRightNavigationId,
                         kRightFilterBarId,
                         kRightFolderViewId,
                         kRightStatusBarId,
                         kRightPreviewTabsId,
                         kRightPreviewContentId))
        {
            Debug::Error(L"FolderWindow::OnCreate failed to create right pane.");
            return false;
        }
    }

    {
        Debug::Perf::Scope perf(L"FolderWindow.OnCreate.CommandLine.Create");
        if (! CreateCommandLineControls(hwnd))
        {
            Debug::Error(L"FolderWindow::OnCreate failed to create command-line controls.");
            return false;
        }
    }

    const int functionBarWidth  = std::max(0L, _functionBarRect.right - _functionBarRect.left);
    const int functionBarHeight = std::max(0L, _functionBarRect.bottom - _functionBarRect.top);
    HWND functionBarHwnd        = nullptr;
    {
        Debug::Perf::Scope perf(L"FolderWindow.OnCreate.FunctionBar.Create");
        functionBarHwnd = _functionBar.Create(hwnd, _functionBarRect.left, _functionBarRect.top, functionBarWidth, functionBarHeight);
    }
    if (functionBarHwnd)
    {
        Debug::Perf::Scope perf(L"FolderWindow.OnCreate.FunctionBar.Initialize");
        _functionBar.SetDpi(_dpi);
        _functionBar.SetShortcutManager(_shortcutManager);
        _functionBar.SetTheme(_theme);
        ShowWindow(functionBarHwnd, _functionBarVisible ? SW_SHOWNA : SW_HIDE);
    }

    {
        Debug::Perf::Scope perf(L"FolderWindow.OnCreate.UpdatePaneUI");
        UpdatePaneFilterBar(Pane::Left);
        UpdatePaneFilterBar(Pane::Right);
        UpdatePaneStatusBar(Pane::Left);
        UpdatePaneStatusBar(Pane::Right);
        UpdatePaneFocusStates();
    }

    const std::wstring_view defaultPluginId = FileSystemPluginManager::GetInstance().GetActivePluginId();
    if (! defaultPluginId.empty())
    {
        Debug::Perf::Scope perf(L"FolderWindow.OnCreate.EnsurePaneFileSystems");
        perf.SetDetail(defaultPluginId);
        static_cast<void>(EnsurePaneFileSystem(Pane::Left, defaultPluginId));
        static_cast<void>(EnsurePaneFileSystem(Pane::Right, defaultPluginId));
    }

    {
        Debug::Perf::Scope perf(L"FolderWindow.OnCreate.NetworkChangeSubscription");
        _networkChangeSubscription = std::make_unique<NetworkChangeSubscription>(hwnd);
    }

    {
        Debug::Perf::Scope perf(L"FolderWindow.OnCreate.ApplyTheme");
        ApplyTheme(ResolveAppTheme(ThemeMode::System, L"RedSalamander"));
    }
    return true;
}

void FolderWindow::DebugShowOverlaySample(Pane pane, FolderView::OverlaySeverity severity)
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.DebugShowOverlaySample(severity);
}

void FolderWindow::DebugShowOverlaySampleNonModal(Pane pane, FolderView::OverlaySeverity severity)
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.DebugShowOverlaySample(FolderView::ErrorOverlayKind::Operation, severity, false);
}

void FolderWindow::DebugShowOverlaySampleBusyWithCancel(Pane pane)
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.DebugShowOverlaySample(FolderView::ErrorOverlayKind::Enumeration, FolderView::OverlaySeverity::Busy, true);
}

void FolderWindow::DebugShowOverlaySampleCanceled(Pane pane)
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.DebugShowCanceledOverlaySample();
}

void FolderWindow::DebugHideOverlaySample(Pane pane)
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.DebugHideOverlaySample();
}

#ifdef ENABLE_TESTS
bool FolderWindow::DebugGetPaneAlertSnapshot(Pane pane, FolderView::AlertOverlayDebugSnapshot& out) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.DebugGetAlertOverlaySnapshot(out);
}

FolderView::RenderingDebugSnapshot FolderWindow::DebugGetPaneRenderingSnapshot(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.DebugGetRenderingSnapshot();
}

void FolderWindow::DebugReportPaneRenderingFailureForSelfTest(Pane pane, HRESULT hr) const
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.DebugReportRenderingFailureForSelfTest(hr);
}

void FolderWindow::DebugAgePaneRenderingFailureForSelfTest(Pane pane, uint64_t ageMs) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.DebugAgeRenderingFailureForSelfTest(ageMs);
}

void FolderWindow::DebugClearPaneRenderingFailureForSelfTest(Pane pane) const
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.DebugClearRenderingFailureForSelfTest();
}
#endif

void FolderWindow::ShowPaneAlertOverlay(Pane pane,
                                        FolderView::ErrorOverlayKind kind,
                                        FolderView::OverlaySeverity severity,
                                        std::wstring title,
                                        std::wstring message,
                                        HRESULT hr,
                                        bool closable,
                                        bool blocksInput)
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (! state.hFolderView)
    {
        return;
    }

    state.folderView.ShowAlertOverlay(kind, severity, std::move(title), std::move(message), hr, closable, blocksInput);
}

void FolderWindow::DismissPaneAlertOverlay(Pane pane)
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (! state.hFolderView)
    {
        return;
    }

    state.folderView.DismissAlertOverlay();
}

void FolderWindow::OnDestroy()
{
#ifdef ENABLE_TESTS
    if (IsFolderWindowSelfTestTracingEnabled())
    {
        SelfTest::AppendSelfTestTrace(L"FolderWindow::OnDestroy: begin");
    }
#endif
    CancelPendingPreviewPaneRefresh();
    _networkChangeSubscription.reset();

    ShutdownViewers();
#ifdef ENABLE_TESTS
    if (IsFolderWindowSelfTestTracingEnabled())
    {
        SelfTest::AppendSelfTestTrace(L"FolderWindow::OnDestroy: viewers shutdown");
    }
#endif
    ShutdownFileOperations();
#ifdef ENABLE_TESTS
    if (IsFolderWindowSelfTestTracingEnabled())
    {
        SelfTest::AppendSelfTestTrace(L"FolderWindow::OnDestroy: file ops shutdown");
    }
#endif

    CancelSelectionSizeComputation(Pane::Left);
    CancelSelectionSizeComputation(Pane::Right);

    if (_leftPane.selectionSizeThread.joinable())
    {
        _leftPane.selectionSizeThread.request_stop();
        _leftPane.selectionSizeCv.notify_all();
        _leftPane.selectionSizeThread = std::jthread{};
    }

    if (_rightPane.selectionSizeThread.joinable())
    {
        _rightPane.selectionSizeThread.request_stop();
        _rightPane.selectionSizeCv.notify_all();
        _rightPane.selectionSizeThread = std::jthread{};
    }

    if (_draggingSplitter)
    {
        ReleaseCapture();
        _draggingSplitter = false;
    }

    DestroyCommandLineControls();

    auto destroyPane = [](PaneState& state)
    {
        if (state.hNavigationView)
        {
            state.navigationView.Destroy();
            state.hNavigationView = nullptr;
        }

        if (state.hFolderView)
        {
            state.folderView.Destroy();
            state.hFolderView = nullptr;
        }

        if (state.hFilterBar)
        {
            state.filterBarHost.Detach();
            state.filterBarLabel  = nullptr;
            state.filterBarCombo  = nullptr;
            state.filterBarToggle = nullptr;
            state.hFilterBar      = nullptr;
        }

        if (state.hPreviewTabs)
        {
            state.previewTabsHost.Detach();
            state.previewTabsControl = nullptr;
            state.hPreviewTabs       = nullptr;
        }

        if (state.hPreviewContent)
        {
            state.previewContentHost.Detach();
            state.previewContentLabel     = nullptr;
            state.previewPropertiesScroll = nullptr;
            state.previewPropertiesSections.clear();
            state.hPreviewContent = nullptr;
        }

        if (state.hStatusBar)
        {
            state.hStatusBar = nullptr;
        }

        state.fileSystem = nullptr;
        state.fileSystemModule.reset();
        state.pluginId.clear();
        state.currentPath.reset();
        state.updatingPath = false;
    };

    destroyPane(_leftPane);
    destroyPane(_rightPane);
}

void FolderWindow::CommandRename(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (state.folderView.GetSelectedPaths().size() > 1u)
    {
        CommandBatchRename(pane);
        return;
    }

    state.folderView.CommandRename();
}

void FolderWindow::RefreshPanesAfterBatchRename(const std::wstring_view sourcePluginId,
                                                const std::wstring_view sourceInstanceContext,
                                                const std::span<const std::filesystem::path> sourcePaths,
                                                const std::span<const std::filesystem::path> targetPaths) noexcept
{
    const auto pathTouchesFolder = [](const std::filesystem::path& folder, const std::filesystem::path& path) noexcept
    {
        if (path.empty())
        {
            return false;
        }

        if (OrdinalString::EqualsNoCasePath(folder, path))
        {
            return true;
        }

        const std::filesystem::path parent = path.parent_path();
        return ! parent.empty() && OrdinalString::EqualsNoCasePath(folder, parent);
    };

    const auto spanTouchesFolder = [&](const std::filesystem::path& folder, const std::span<const std::filesystem::path> paths) noexcept
    {
        return std::ranges::any_of(paths, [&](const std::filesystem::path& path) noexcept
        { return pathTouchesFolder(folder, path); });
    };

    // When `source` is a path-segment-aware case-insensitive ancestor of `folder` (strictly: `folder`
    // lives somewhere below the renamed directory), returns the corresponding folder under `target`.
    const auto retargetDescendantFolder =
        [](const std::filesystem::path& folder, const std::filesystem::path& source, const std::filesystem::path& target) noexcept
        -> std::optional<std::filesystem::path>
    {
        if (folder.empty() || source.empty() || target.empty())
        {
            return std::nullopt;
        }

        const std::filesystem::path normalizedSource = source.lexically_normal();
        const std::filesystem::path normalizedFolder = folder.lexically_normal();

        auto folderIt        = normalizedFolder.begin();
        const auto folderEnd = normalizedFolder.end();
        for (auto sourceIt = normalizedSource.begin(); sourceIt != normalizedSource.end(); ++sourceIt)
        {
            if (sourceIt->empty())
            {
                continue; // trailing-separator artifact of path iteration
            }

            while (folderIt != folderEnd && folderIt->empty())
            {
                ++folderIt;
            }

            if (folderIt == folderEnd || ! OrdinalString::EqualsNoCase(sourceIt->native(), folderIt->native()))
            {
                return std::nullopt;
            }

            ++folderIt;
        }

        std::filesystem::path retargeted = target;
        bool hasRemainder                = false;
        for (; folderIt != folderEnd; ++folderIt)
        {
            if (folderIt->empty())
            {
                continue;
            }

            retargeted /= *folderIt;
            hasRemainder = true;
        }

        return hasRemainder ? std::optional<std::filesystem::path>(std::move(retargeted)) : std::nullopt;
    };

    const auto findDescendantRetarget = [&](const std::filesystem::path& folder) noexcept -> std::optional<std::filesystem::path>
    {
        // Apply the rename pairs sequentially (they arrive in execution order, deepest first) so a
        // folder below both a renamed child and a renamed parent ends up under both new names.
        std::filesystem::path current = folder;
        bool rewritten                = false;
        const size_t renameCount      = std::min(sourcePaths.size(), targetPaths.size());
        for (size_t index = 0u; index < renameCount; ++index)
        {
            std::optional<std::filesystem::path> retargeted = retargetDescendantFolder(current, sourcePaths[index], targetPaths[index]);
            if (retargeted.has_value())
            {
                current   = std::move(retargeted).value();
                rewritten = true;
            }
        }

        return rewritten ? std::optional<std::filesystem::path>(std::move(current)) : std::nullopt;
    };

    const auto refreshIfAffected = [&](const Pane pane) noexcept
    {
        PaneState& paneState = pane == Pane::Left ? _leftPane : _rightPane;
        if (! paneState.fileSystem || ! OrdinalString::EqualsNoCase(paneState.pluginId, sourcePluginId) ||
            ! OrdinalString::EqualsNoCase(paneState.instanceContext, sourceInstanceContext))
        {
            return;
        }

        const auto folder = paneState.folderView.GetFolderPath();
        if (! folder.has_value())
        {
            return;
        }

        if (spanTouchesFolder(folder.value(), sourcePaths) || spanTouchesFolder(folder.value(), targetPaths))
        {
            paneState.folderView.ForceRefresh();
            return;
        }

        // A pane showing a folder inside a renamed directory would be left on a dead path; follow the
        // rename instead. Limited to plain local file-system panes, where the rewritten path can be
        // validated and renavigated directly.
        if (! paneState.instanceContext.empty() || ! NavigationLocation::IsFilePluginShortId(paneState.pluginShortId))
        {
            return;
        }

        const std::optional<std::filesystem::path> retargeted = findDescendantRetarget(folder.value());
        if (! retargeted.has_value())
        {
            return;
        }

        // If the rewritten path does not exist (e.g. a deeper segment changed too), fall back to the
        // nearest existing ancestor so the pane never stays on a dead path.
        std::filesystem::path destination = retargeted.value();
        std::error_code existsError;
        while (! destination.empty() && ! std::filesystem::exists(destination, existsError))
        {
            const std::filesystem::path parent = destination.parent_path();
            if (parent.empty() || parent.native() == destination.native())
            {
                break;
            }

            destination = parent;
        }

        if (destination.empty())
        {
            destination = retargeted.value();
        }

        SetFolderPath(pane, destination);
    };

    refreshIfAffected(Pane::Left);
    refreshIfAffected(Pane::Right);
}

bool FolderWindow::RevealBatchRenamePathInPane(const Pane pane, const std::filesystem::path& path) noexcept
{
    if (path.empty())
    {
        return false;
    }

    const std::filesystem::path parent = path.parent_path();
    const std::wstring leaf            = path.filename().native();
    if (parent.empty() || leaf.empty())
    {
        return false;
    }

    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.RememberFocusedItemForFolder(parent, leaf);

    const std::optional<std::filesystem::path> currentFolder = state.folderView.GetFolderPath();
    if (currentFolder.has_value() && OrdinalString::EqualsNoCasePath(currentFolder.value(), parent))
    {
        return state.folderView.PrepareForExternalCommand(leaf);
    }

    SetFolderPath(pane, parent);
    return true;
}

void FolderWindow::CommandBatchRename(Pane pane, std::optional<std::filesystem::path> rootOverride, std::vector<std::filesystem::path> initialPathsOverride)
{
    SetActivePane(pane);
    if (! _settings)
    {
        return;
    }

    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;

    BatchRenamePaneContext context{};
    context.fileSystem      = state.fileSystem;
    context.pluginId        = state.pluginId;
    context.pluginShortId   = state.pluginShortId;
    context.instanceContext = state.instanceContext;
    context.rootPluginPath  = rootOverride.has_value() ? rootOverride.value() : state.currentPath.value_or(std::filesystem::path{});
    if (! rootOverride.has_value())
    {
        context.initialPaths = initialPathsOverride.empty() ? state.folderView.GetSelectedOrFocusedPaths() : std::move(initialPathsOverride);
    }
    // Capture the originating pane's identity now: the Batch Rename window is modeless, so the pane may
    // navigate elsewhere (e.g. into an archive) before renames complete.
    context.onSuccessfulRename = [this, sourcePluginId = state.pluginId, sourceInstanceContext = state.instanceContext](
                                     std::span<const std::filesystem::path> sourcePaths,
                                     std::span<const std::filesystem::path> targetPaths) noexcept
    {
        RefreshPanesAfterBatchRename(sourcePluginId, sourceInstanceContext, sourcePaths, targetPaths);
    };
    context.onRevealPath = [this, pane](const std::filesystem::path& path) noexcept
    { return RevealBatchRenamePathInPane(pane, path); };

    static_cast<void>(ShowBatchRenameWindow(_hWnd.get(), *_settings, _theme, std::move(context)));
}

void FolderWindow::CommandView(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    ClearFileActionFailure();
    state.folderView.CommandView();
    static_cast<void>(ShowRecordedFileActionFailureOverlay(pane));
}

void FolderWindow::CommandAlternateView(Pane pane)
{
    SetActivePane(pane);
    PaneState& state                                       = pane == Pane::Left ? _leftPane : _rightPane;
    const std::optional<std::filesystem::path> focusedPath = state.folderView.GetFocusedPath();
    ClearFileActionFailure();
    if (! state.folderView.CommandAlternateView())
    {
        if (! ShowRecordedFileActionFailureOverlay(pane))
        {
            ShowAlternateFileActionUnavailableOverlay(*this, pane, true, focusedPath);
        }
    }
}

void FolderWindow::CommandViewWith(Pane pane, std::wstring_view actionId)
{
    if (actionId.empty())
    {
        return;
    }

    SetActivePane(pane);
    PaneState& state                                       = pane == Pane::Left ? _leftPane : _rightPane;
    const std::optional<std::filesystem::path> focusedPath = state.folderView.GetFocusedPath();
    ClearFileActionFailure();
    if (! state.folderView.CommandViewWith(actionId))
    {
        if (! ShowRecordedFileActionFailureOverlay(pane))
        {
            ShowFileActionUnavailableOverlay(*this, pane, true, actionId, focusedPath);
        }
    }
}

void FolderWindow::CommandEdit(Pane pane)
{
    SetActivePane(pane);
    ClearFileActionFailure();
    static_cast<void>(TryEditFocusedFileWithEditor(pane, {}, false));
    static_cast<void>(ShowRecordedFileActionFailureOverlay(pane));
}

void FolderWindow::CommandAlternateEdit(Pane pane)
{
    SetActivePane(pane);
    PaneState& state                                       = pane == Pane::Left ? _leftPane : _rightPane;
    const std::optional<std::filesystem::path> focusedPath = state.folderView.GetFocusedPath();
    ClearFileActionFailure();
    if (! TryEditFocusedFileWithEditor(pane, {}, true))
    {
        if (! ShowRecordedFileActionFailureOverlay(pane))
        {
            ShowAlternateFileActionUnavailableOverlay(*this, pane, false, focusedPath);
        }
    }
}

void FolderWindow::CommandEditWith(Pane pane, std::wstring_view actionId)
{
    if (actionId.empty())
    {
        return;
    }

    SetActivePane(pane);
    PaneState& state                                       = pane == Pane::Left ? _leftPane : _rightPane;
    const std::optional<std::filesystem::path> focusedPath = state.folderView.GetFocusedPath();
    ClearFileActionFailure();
    if (! TryEditFocusedFileWithEditor(pane, actionId, false))
    {
        if (! ShowRecordedFileActionFailureOverlay(pane))
        {
            ShowFileActionUnavailableOverlay(*this, pane, false, actionId, focusedPath);
        }
    }
}

void FolderWindow::CommandContextMenuCurrentDirectory(Pane pane)
{
    Debug::Perf::Scope perf(L"shell.context_menu_current_directory_us");
    SetActivePane(pane);

    if (GetFileSystemPluginId(pane) != std::wstring_view(L"builtin/file-system"))
    {
        ShowShellActionUnavailableOverlay(*this, pane, IDS_MSG_SHELL_ACTION_LOCAL_FOLDER_REQUIRED);
        return;
    }

    const std::optional<std::filesystem::path> pathOpt = GetCurrentPluginPath(pane);
    if (! pathOpt.has_value() || pathOpt.value().empty())
    {
        ShowShellActionUnavailableOverlay(*this, pane, IDS_MSG_SHELL_ACTION_LOCAL_FOLDER_REQUIRED);
        return;
    }

    const std::filesystem::path path = pathOpt.value();

#ifdef ENABLE_TESTS
    if (_debugShellActionCallback)
    {
        DebugShellAction action{};
        action.kind = DebugShellActionKind::ContextMenuCurrentDirectory;
        action.pane = pane;
        action.path = path;

        const HRESULT hr = _debugShellActionCallback(action);
        if (FAILED(hr))
        {
            ShowShellActionFailedOverlay(*this, pane, path, hr);
        }
        return;
    }
#endif

    const HRESULT hr = ShowShellContextMenuForPath(_hWnd.get(), path);
    if (FAILED(hr))
    {
        ShowShellActionFailedOverlay(*this, pane, path, hr);
    }
}

void FolderWindow::CommandOpenSecurity(Pane pane)
{
    Debug::Perf::Scope perf(L"shell.open_security_us");
    SetActivePane(pane);

    if (GetFileSystemPluginId(pane) != std::wstring_view(L"builtin/file-system"))
    {
        ShowShellActionUnavailableOverlay(*this, pane, IDS_MSG_SHELL_ACTION_LOCAL_SELECTION_REQUIRED);
        return;
    }

    const std::optional<std::filesystem::path> pathOpt = GetFocusedItemPath(pane);
    if (! pathOpt.has_value() || pathOpt.value().empty())
    {
        ShowShellActionUnavailableOverlay(*this, pane, IDS_MSG_SHELL_ACTION_LOCAL_SELECTION_REQUIRED);
        return;
    }

    const std::filesystem::path path = pathOpt.value();

#ifdef ENABLE_TESTS
    if (_debugShellActionCallback)
    {
        DebugShellAction action{};
        action.kind         = DebugShellActionKind::OpenSecurity;
        action.pane         = pane;
        action.path         = path;
        action.propertyPage = L"Security";

        const HRESULT hr = _debugShellActionCallback(action);
        if (FAILED(hr))
        {
            ShowShellActionFailedOverlay(*this, pane, path, hr);
        }
        return;
    }
#endif

    const BOOL shown = SHObjectProperties(_hWnd.get(), SHOP_FILEPATH, path.c_str(), L"Security");
    if (shown == FALSE)
    {
        const DWORD lastError = GetLastError();
        const HRESULT hr      = lastError != ERROR_SUCCESS ? HRESULT_FROM_WIN32(lastError) : E_FAIL;
        Debug::Warning(L"Open Security: SHObjectProperties failed for {}: {:#x}", path.wstring(), hr);
        ShowShellActionFailedOverlay(*this, pane, path, hr);
    }
}

void FolderWindow::CommandGoToShortcutOrLinkTarget(Pane pane)
{
    Debug::Perf::Scope perf(L"shell.go_to_shortcut_target_us");
    SetActivePane(pane);

    if (GetFileSystemPluginId(pane) != std::wstring_view(L"builtin/file-system"))
    {
        ShowShellActionUnavailableOverlay(*this, pane, IDS_MSG_SHORTCUT_TARGET_SELECTION_REQUIRED);
        return;
    }

    const std::optional<std::filesystem::path> sourcePathOpt = GetFocusedItemPath(pane);
    if (! sourcePathOpt.has_value() || sourcePathOpt.value().empty())
    {
        ShowShellActionUnavailableOverlay(*this, pane, IDS_MSG_SHORTCUT_TARGET_SELECTION_REQUIRED);
        return;
    }

    const std::filesystem::path sourcePath    = sourcePathOpt.value();
    const ShortcutTargetResolution resolution = ResolveShortcutOrLinkTarget(sourcePath);
    if (resolution.status == ShortcutTargetResolutionStatus::Unsupported)
    {
        ShowShellActionUnavailableOverlay(*this, pane, IDS_MSG_SHORTCUT_TARGET_UNSUPPORTED);
        return;
    }
    if (resolution.status == ShortcutTargetResolutionStatus::Failed || resolution.target.empty())
    {
        ShowShellActionFailedOverlay(*this, pane, sourcePath, FAILED(resolution.hr) ? resolution.hr : E_FAIL);
        return;
    }

    const DWORD targetAttributes = GetFileAttributesW(resolution.target.c_str());
    if (targetAttributes == INVALID_FILE_ATTRIBUTES)
    {
        const DWORD lastError = GetLastError();
        ShowShellActionFailedOverlay(*this, pane, resolution.target, lastError != ERROR_SUCCESS ? HRESULT_FROM_WIN32(lastError) : E_FAIL);
        return;
    }

    if ((targetAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0u)
    {
        SetFolderPath(pane, resolution.target);
        return;
    }

    const std::filesystem::path parent = resolution.target.parent_path();
    if (parent.empty())
    {
        ShowShellActionFailedOverlay(*this, pane, resolution.target, E_INVALIDARG);
        return;
    }

    const DWORD parentAttributes = GetFileAttributesW(parent.c_str());
    if (parentAttributes == INVALID_FILE_ATTRIBUTES || (parentAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0u)
    {
        const HRESULT hr = parentAttributes == INVALID_FILE_ATTRIBUTES ? HRESULT_FROM_WIN32(GetLastError()) : E_INVALIDARG;
        ShowShellActionFailedOverlay(*this, pane, parent, FAILED(hr) ? hr : E_FAIL);
        return;
    }

    PaneState& state             = pane == Pane::Left ? _leftPane : _rightPane;
    const std::wstring focusName = resolution.target.filename().wstring();
    if (! focusName.empty())
    {
        state.folderView.RememberFocusedItemForFolder(parent, focusName);
    }

    SetFolderPath(pane, parent);
}

void FolderWindow::CommandViewSpace(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;

    std::filesystem::path targetPath;
    const std::vector<std::filesystem::path> selectedDirs = state.folderView.GetSelectedDirectoryPaths();
    if (selectedDirs.size() == 1)
    {
        targetPath = selectedDirs.front();
    }
    else
    {
        const std::optional<std::filesystem::path> currentFolder = state.folderView.GetFolderPath();
        if (currentFolder.has_value())
        {
            targetPath = currentFolder.value();
        }
    }

    if (targetPath.empty())
    {
        return;
    }

    static_cast<void>(TryViewSpaceWithViewer(pane, targetPath));
}

void FolderWindow::SetShowSortMenuCallback(ShowSortMenuCallback callback)
{
    _showSortMenuCallback = std::move(callback);
}
