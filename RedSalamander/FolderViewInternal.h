#pragma once

// Internal implementation header for FolderView split across multiple .cpp files.
// Keep this header private to the FolderView translation units.

#include <algorithm>
#include <array>
#include <atomic>
#include <cassert>
#include <chrono>
#include <cmath>
#include <cstring>
#include <cwchar>
#include <cwctype>
#include <exception>
#include <execution>
#include <format>
#include <iterator>
#include <limits>
#include <memory>
#include <new>
#include <unordered_map>
#include <unordered_set>

#define WINDOWS_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <commctrl.h>
#include <commdlg.h>
#include <dwmapi.h>
#include <propvarutil.h>
#include <shellapi.h>
#include <shellscalingapi.h>
#include <shlobj_core.h>
#include <shlwapi.h>
#include <shobjidl.h>

#include <sstream>
#include <system_error>

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027)
#include <wil/result.h>
#pragma warning(pop)

#include <wincodec.h>
#include <wincodecsdk.h>

#include "Helpers.h"
#include "PlugInterfaces/FileSystem.h"
#include "PlugInterfaces/Informations.h"
#include "Ui/AlertOverlay.h"

#include "FolderView.h"
#include "Helpers.h"
#include "HostServices.h"
#include "IconCache.h"
#include "ThemedControls.h"
#include "WindowMessages.h"
#include "resource.h"

#ifndef GET_X_LPARAM
#define GET_X_LPARAM(lp) ((int)(short)LOWORD(lp))
#define GET_Y_LPARAM(lp) ((int)(short)HIWORD(lp))
#endif

#ifndef CLSID_WICImagingFactory2
#define CLSID_WICImagingFactory2 CLSID_WICImagingFactory
#endif

#ifndef WICBitmapAlphaChannelOptionUseBitmapAlpha
// Fallback definition for older Windows SDK versions
#pragma warning(push)
#pragma warning(disable : 5264) // C5264: 'const' variable is not used
inline constexpr WICBitmapAlphaChannelOption WICBitmapAlphaChannelOptionUseBitmapAlpha_Fallback = static_cast<WICBitmapAlphaChannelOption>(2);
#pragma warning(pop)
#define WICBitmapAlphaChannelOptionUseBitmapAlpha WICBitmapAlphaChannelOptionUseBitmapAlpha_Fallback
#endif

#pragma comment(lib, "Shell32.lib")
#pragma comment(lib, "Ole32.lib")
#pragma comment(lib, "Dwmapi.lib")

using wil::unique_hbitmap;

#pragma warning(push)
// 5245 : unreferenced function with internal linkage has been removed
#pragma warning(disable : 5245)

namespace
{
constexpr wchar_t kFolderViewClassName[]          = L"RedSalamanderFolderView";
constexpr float kLabelHorizontalPaddingDip        = 12.0f;
constexpr float kLabelVerticalPaddingDip          = 4.0f;
constexpr float kFocusStrokeThicknessDip          = 2.0f;
constexpr float kFocusStrokeThicknessUnfocusedDip = 1.0f;
constexpr float kFocusBorderOpacityUnfocused      = 0.60f;
constexpr float kSelectionCornerRadiusDip         = 2.0f;
constexpr float kIconTextGapDip                   = 12.0f;
constexpr float kColumnSpacingDip                 = 18.0f;
constexpr float kRowSpacingDip                    = 4.0f;
constexpr float kDetailsGapDip                    = 2.0f;
constexpr float kDetailsTextAlpha                 = 0.75f;
constexpr float kMetadataTextAlpha                = 0.55f;
constexpr UINT kSwapChainBufferCount              = 2;
constexpr UINT_PTR kOverlayTimerId                = 1;
constexpr uint64_t kBusyOverlayDelayMs            = 300;

bool ConfirmNonRevertableFileOperation(HWND owner,
                                       [[maybe_unused]] IFileSystem* fileSystem,
                                       FileSystemOperation operation,
                                       const std::vector<std::filesystem::path>& sourcePaths,
                                       const std::filesystem::path& destinationFolder) noexcept
{
    if (operation != FILESYSTEM_COPY && operation != FILESYSTEM_MOVE)
    {
        return true;
    }

    if (sourcePaths.empty())
    {
        return true;
    }

    // Avoid I/O in the confirmation prompt path (plugins may require network access to answer GetAttributes).
    // Best-effort: treat item types as unknown.
    unsigned long long fileCount    = 0;
    unsigned long long folderCount  = 0;
    unsigned long long unknownCount = static_cast<unsigned long long>(sourcePaths.size());
    std::filesystem::path sampleFile;
    bool hasSampleFile = false;

    auto suffixFor = [](unsigned long long count) noexcept -> std::wstring_view { return count == 1ull ? std::wstring_view(L"") : std::wstring_view(L"s"); };

    const unsigned long long itemCount = static_cast<unsigned long long>(sourcePaths.size());
    std::wstring what;
    if (unknownCount > 0)
    {
        const std::wstring_view itemSuffix = suffixFor(itemCount);
        what                               = FormatStringResource(nullptr, IDS_FMT_FILEOPS_COUNT_ITEM, itemCount, itemSuffix);
    }
    else if (fileCount > 0 && folderCount > 0)
    {
        const std::wstring_view fileSuffix   = suffixFor(fileCount);
        const std::wstring_view folderSuffix = suffixFor(folderCount);
        what = FormatStringResource(nullptr, IDS_FMT_FILEOPS_COUNT_FILES_FOLDERS, fileCount, fileSuffix, folderCount, folderSuffix);
    }
    else if (fileCount > 0)
    {
        const std::wstring_view fileSuffix = suffixFor(fileCount);
        what                               = FormatStringResource(nullptr, IDS_FMT_FILEOPS_COUNT_FILE, fileCount, fileSuffix);
    }
    else
    {
        const std::wstring_view folderSuffix = suffixFor(folderCount);
        what                                 = FormatStringResource(nullptr, IDS_FMT_FILEOPS_COUNT_FOLDER, folderCount, folderSuffix);
    }

    auto ensureTrailingSeparator = [](std::wstring text) noexcept -> std::wstring
    {
        if (text.empty())
        {
            return text;
        }

        const wchar_t last = text.back();
        if (last == L'\\' || last == L'/')
        {
            return text;
        }

        text.push_back(L'\\');
        return text;
    };

    auto normalizeSlashes = [](std::wstring& text) noexcept
    {
        for (auto& ch : text)
        {
            if (ch == L'/')
            {
                ch = L'\\';
            }
        }
    };

    std::wstring fromText;
    if (sourcePaths.size() == 1u)
    {
        fromText = sourcePaths.front().wstring();
        if (unknownCount == 0 && folderCount == 1ull && fileCount == 0ull)
        {
            fromText = ensureTrailingSeparator(std::move(fromText));
        }
    }
    else
    {
        std::filesystem::path commonParent = sourcePaths.front().parent_path();
        bool multipleParents               = false;
        for (size_t index = 1; index < sourcePaths.size(); ++index)
        {
            const std::filesystem::path parent = sourcePaths[index].parent_path();
            if (CompareStringOrdinal(commonParent.c_str(), -1, parent.c_str(), -1, TRUE) != CSTR_EQUAL)
            {
                multipleParents = true;
                break;
            }
        }

        if (multipleParents)
        {
            fromText = LoadStringResource(nullptr, IDS_FILEOPS_LOCATION_MULTIPLE);
        }
        else if (unknownCount == 0 && fileCount > 0 && folderCount > 0 && hasSampleFile)
        {
            fromText = sampleFile.wstring();
        }
        else
        {
            fromText = ensureTrailingSeparator(commonParent.wstring());
        }
    }

    std::wstring toText = ensureTrailingSeparator(destinationFolder.wstring());
    normalizeSlashes(fromText);
    normalizeSlashes(toText);

    const UINT messageId = operation == FILESYSTEM_COPY ? static_cast<UINT>(IDS_FMT_FILEOPS_CONFIRM_COPY) : static_cast<UINT>(IDS_FMT_FILEOPS_CONFIRM_MOVE);
    const std::wstring message = FormatStringResource(nullptr, messageId, what, fromText, toText);

    const std::wstring caption = LoadStringResource(nullptr, IDS_CAPTION_CONFIRM);
    HostPromptRequest prompt{};
    prompt.version       = 1;
    prompt.sizeBytes     = sizeof(prompt);
    prompt.scope         = (owner && IsWindow(owner)) ? HOST_ALERT_SCOPE_WINDOW : HOST_ALERT_SCOPE_APPLICATION;
    prompt.severity      = HOST_ALERT_INFO;
    prompt.buttons       = HOST_PROMPT_BUTTONS_OK_CANCEL;
    prompt.targetWindow  = (prompt.scope == HOST_ALERT_SCOPE_WINDOW) ? owner : nullptr;
    prompt.title         = caption.c_str();
    prompt.message       = message.c_str();
    prompt.defaultResult = HOST_PROMPT_RESULT_OK;

    HostPromptResult promptResult = HOST_PROMPT_RESULT_NONE;
    const HRESULT hr              = HostShowPrompt(prompt, nullptr, &promptResult);
    if (FAILED(hr))
    {
        return false;
    }

    return promptResult == HOST_PROMPT_RESULT_OK;
}

bool IsOverlaySampleEnabled() noexcept
{
#if defined(_DEBUG) || defined(DEBUG)
    return true;
#else
    return false;
#endif
}

enum FolderCommands : UINT
{
    CmdOpen                             = IDM_FOLDERVIEW_CONTEXT_OPEN,
    CmdOpenWith                         = IDM_FOLDERVIEW_CONTEXT_OPEN_WITH,
    CmdViewSpace                        = IDM_FOLDERVIEW_CONTEXT_VIEW_SPACE,
    CmdDelete                           = IDM_FOLDERVIEW_CONTEXT_DELETE,
    CmdRename                           = IDM_FOLDERVIEW_CONTEXT_RENAME,
    CmdCopy                             = IDM_FOLDERVIEW_CONTEXT_COPY,
    CmdPaste                            = IDM_FOLDERVIEW_CONTEXT_PASTE,
    CmdSelectAll                        = IDM_FOLDERVIEW_CONTEXT_SELECT_ALL,
    CmdUnselectAll                      = IDM_FOLDERVIEW_CONTEXT_UNSELECT_ALL,
    CmdProperties                       = IDM_FOLDERVIEW_CONTEXT_PROPERTIES,
    CmdMove                             = IDM_FOLDERVIEW_CONTEXT_MOVE,
    CmdOverlaySampleError               = IDM_FOLDERVIEW_CONTEXT_OVERLAY_SAMPLE_ERROR,
    CmdOverlaySampleWarning             = IDM_FOLDERVIEW_CONTEXT_OVERLAY_SAMPLE_WARNING,
    CmdOverlaySampleInformation         = IDM_FOLDERVIEW_CONTEXT_OVERLAY_SAMPLE_INFORMATION,
    CmdOverlaySampleBusy                = IDM_FOLDERVIEW_CONTEXT_OVERLAY_SAMPLE_BUSY,
    CmdOverlaySampleHide                = IDM_FOLDERVIEW_CONTEXT_OVERLAY_SAMPLE_HIDE,
    CmdOverlaySampleErrorNonModal       = IDM_FOLDERVIEW_CONTEXT_OVERLAY_SAMPLE_ERROR_NONMODAL,
    CmdOverlaySampleWarningNonModal     = IDM_FOLDERVIEW_CONTEXT_OVERLAY_SAMPLE_WARNING_NONMODAL,
    CmdOverlaySampleInformationNonModal = IDM_FOLDERVIEW_CONTEXT_OVERLAY_SAMPLE_INFORMATION_NONMODAL,
    CmdOverlaySampleCanceled            = IDM_FOLDERVIEW_CONTEXT_OVERLAY_SAMPLE_CANCELED,
    CmdOverlaySampleBusyWithCancel      = IDM_FOLDERVIEW_CONTEXT_OVERLAY_SAMPLE_BUSY_WITH_CANCEL,
};

struct RenameDialogState
{
    std::wstring currentName;
    std::wstring newName;
    bool isDirectory = false;
};

std::wstring FormatHResult(HRESULT hr)
{
    wil::unique_hlocal_string message;
    if (SUCCEEDED(::FormatMessageW(FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_IGNORE_INSERTS,
                                   nullptr,
                                   static_cast<DWORD>(hr),
                                   MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
                                   reinterpret_cast<LPWSTR>(message.addressof()),
                                   0,
                                   nullptr)))
    {
        return std::wstring(message.get());
    }
    return std::format(L"HRESULT 0x{:08X}", hr);
}

HRESULT HrFromErrorCode(const std::error_code& ec)
{
    if (! ec)
    {
        return S_OK;
    }
    if (&ec.category() == &std::system_category())
    {
        return HRESULT_FROM_WIN32(static_cast<DWORD>(ec.value()));
    }
    return HRESULT_FROM_WIN32(ERROR_GEN_FAILURE);
}

std::wstring FormatLocalTime(int64_t fileTime)
{
    if (fileTime <= 0)
    {
        return {};
    }

    ULARGE_INTEGER uli{};
    uli.QuadPart = static_cast<ULONGLONG>(fileTime);

    FILETIME ft{};
    ft.dwLowDateTime  = uli.LowPart;
    ft.dwHighDateTime = uli.HighPart;

    FILETIME local{};
    SYSTEMTIME st{};
    if (! FileTimeToLocalFileTime(&ft, &local) || ! FileTimeToSystemTime(&local, &st))
    {
        return {};
    }

    return std::format(L"{:04d}-{:02d}-{:02d} {:02d}:{:02d}", st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute);
}

std::wstring FormatFileAttributes(DWORD attrs)
{
    std::wstring result;
    result.reserve(10);

    auto add = [&](DWORD flag, wchar_t ch)
    {
        if ((attrs & flag) != 0)
        {
            result.push_back(ch);
        }
    };

    add(FILE_ATTRIBUTE_READONLY, L'R');
    add(FILE_ATTRIBUTE_HIDDEN, L'H');
    add(FILE_ATTRIBUTE_SYSTEM, L'S');
    add(FILE_ATTRIBUTE_ARCHIVE, L'A');
    add(FILE_ATTRIBUTE_COMPRESSED, L'C');
    add(FILE_ATTRIBUTE_ENCRYPTED, L'E');
    add(FILE_ATTRIBUTE_TEMPORARY, L'T');
    add(FILE_ATTRIBUTE_OFFLINE, L'O');
    add(FILE_ATTRIBUTE_REPARSE_POINT, L'P');

    if (result.empty())
    {
        result = L"-";
    }

    return result;
}

std::wstring FileTypeLabel(std::wstring_view extension, bool isDirectory)
{
    if (isDirectory)
    {
        return LoadStringResource(nullptr, IDS_FOLDERVIEW_TYPE_FOLDER);
    }

    std::wstring type(extension);
    if (! type.empty() && type.front() == L'.')
    {
        type.erase(type.begin());
    }
    if (type.empty())
    {
        return LoadStringResource(nullptr, IDS_FOLDERVIEW_TYPE_FILE);
    }

    for (auto& ch : type)
    {
        ch = static_cast<wchar_t>(towupper(ch));
    }

    return type;
}

std::wstring PadLeftToWidth(std::wstring_view text, size_t width)
{
    if (text.size() >= width)
    {
        return std::wstring(text);
    }

    std::wstring result;
    result.reserve(width);
    result.append(width - text.size(), L' ');
    result.append(text);
    return result;
}

std::wstring BuildDetailsText(bool isDirectory, uint64_t sizeBytes, int64_t lastWriteTime, DWORD fileAttributes, size_t sizeSlotChars)
{
    const std::wstring timeText  = FormatLocalTime(lastWriteTime);
    const std::wstring attrsText = FormatFileAttributes(fileAttributes);

    if (isDirectory)
    {
        return std::format(L"{} • {}", timeText, attrsText);
    }

    std::wstring sizeField;
    if (sizeSlotChars > 0)
    {
        const std::wstring sizeText = FormatBytesCompact(sizeBytes);
        sizeField                   = PadLeftToWidth(sizeText, sizeSlotChars);
    }
    else
    {
        sizeField = FormatBytesCompact(sizeBytes);
    }

    return std::format(L"{} • {} • {}", timeText, sizeField, attrsText);
}

constexpr UINT_PTR kRenameEditSubclassId = 1;

void CenterMultilineEditTextVertically(HWND edit) noexcept
{
    ThemedControls::CenterEditTextVertically(edit);
}

LRESULT OnRenameEditPaste(HWND hwnd, WPARAM wParam, LPARAM lParam)
{
    const LRESULT result = DefSubclassProc(hwnd, WM_PASTE, wParam, lParam);

    const int length = GetWindowTextLengthW(hwnd);
    if (length <= 0)
    {
        return result;
    }

    std::wstring buffer;
    buffer.resize(static_cast<size_t>(length) + 1u);
    GetWindowTextW(hwnd, buffer.data(), length + 1);
    buffer.resize(static_cast<size_t>(length));

    buffer.erase(std::remove(buffer.begin(), buffer.end(), L'\r'), buffer.end());
    buffer.erase(std::remove(buffer.begin(), buffer.end(), L'\n'), buffer.end());
    buffer.erase(std::remove(buffer.begin(), buffer.end(), L'\t'), buffer.end());

    SetWindowTextW(hwnd, buffer.c_str());
    return result;
}

LRESULT CALLBACK RenameEditSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam, UINT_PTR /*uIdSubclass*/, DWORD_PTR /*dwRefData*/)
{
    switch (msg)
    {
        case WM_KEYDOWN:
            if (wParam == VK_RETURN)
            {
                SendMessageW(GetParent(hwnd), WM_COMMAND, IDOK, 0);
                return 0;
            }
            break;
        case WM_CHAR:
            if (wParam == L'\r' || wParam == L'\n')
            {
                return 0;
            }
            break;
        case WM_PASTE: return OnRenameEditPaste(hwnd, wParam, lParam);
    }

    return DefSubclassProc(hwnd, msg, wParam, lParam);
}

INT_PTR OnRenameDialogInit(HWND dlg, RenameDialogState* state)
{
    if (! state)
    {
        return FALSE;
    }

    SetWindowLongPtrW(dlg, DWLP_USER, reinterpret_cast<LONG_PTR>(state));
    const HWND edit = GetDlgItem(dlg, IDC_FOLDERVIEW_RENAME_EDIT);
    if (edit)
    {
        SetWindowTextW(edit, state->currentName.c_str());
        CenterMultilineEditTextVertically(edit);

#pragma warning(push)
#pragma warning(disable : 5039) // C5039: passing potentially-throwing callback to extern "C" Win32 API under -EHc
        SetWindowSubclass(edit, RenameEditSubclassProc, kRenameEditSubclassId, 0);
#pragma warning(pop)

        int selectionEnd = -1;
        if (! state->isDirectory)
        {
            const std::wstring_view nameView(state->currentName);
            const size_t dotPos = nameView.find_last_of(L'.');
            if (dotPos != std::wstring_view::npos && dotPos > 0 && dotPos + 1 < nameView.size())
            {
                if (dotPos <= static_cast<size_t>(std::numeric_limits<int>::max()))
                {
                    selectionEnd = static_cast<int>(dotPos);
                }
            }
        }

        SetFocus(edit);
        SendMessageW(edit, EM_SETSEL, 0, static_cast<LPARAM>(selectionEnd));
        return FALSE;
    }
    return TRUE;
}

INT_PTR OnRenameDialogCommand(HWND dlg, RenameDialogState* state, UINT commandId)
{
    if (commandId == IDCANCEL)
    {
        EndDialog(dlg, IDCANCEL);
        return TRUE;
    }

    if (commandId != IDOK)
    {
        return FALSE;
    }

    if (! state)
    {
        return FALSE;
    }

    wchar_t buffer[MAX_PATH] = {};
    GetDlgItemTextW(dlg, IDC_FOLDERVIEW_RENAME_EDIT, buffer, static_cast<int>(std::size(buffer)));

    std::wstring trimmed(buffer);
    trimmed.erase(trimmed.begin(), std::find_if(trimmed.begin(), trimmed.end(), [](wchar_t ch) { return ! iswspace(ch); }));
    trimmed.erase(std::find_if(trimmed.rbegin(), trimmed.rend(), [](wchar_t ch) { return ! iswspace(ch); }).base(), trimmed.end());

    if (trimmed.empty())
    {
        MessageBeep(MB_ICONWARNING);
        return TRUE;
    }

    state->newName = std::move(trimmed);
    EndDialog(dlg, IDOK);
    return TRUE;
}

INT_PTR CALLBACK RenameDialogProc(HWND dlg, UINT msg, WPARAM wParam, LPARAM lParam)
{
    auto* state = reinterpret_cast<RenameDialogState*>(GetWindowLongPtrW(dlg, DWLP_USER));

    switch (msg)
    {
        case WM_INITDIALOG: return OnRenameDialogInit(dlg, reinterpret_cast<RenameDialogState*>(lParam));
        case WM_COMMAND: return OnRenameDialogCommand(dlg, state, LOWORD(wParam));
    }
    return FALSE;
}

std::optional<std::wstring> PromptForRename(HWND owner, const std::wstring& currentName, bool isDirectory)
{
    RenameDialogState state{};
    state.currentName = currentName;
    state.isDirectory = isDirectory;
#pragma warning(push)
// pointer or reference to potentially throwing function passed to 'extern "C"' function
#pragma warning(disable : 5039)
    INT_PTR result =
        DialogBoxParamW(GetModuleHandleW(nullptr), MAKEINTRESOURCEW(IDD_FOLDERVIEW_RENAME), owner, RenameDialogProc, reinterpret_cast<LPARAM>(&state));
#pragma warning(pop)
    if (result == IDOK && ! state.newName.empty())
    {
        return state.newName;
    }
    return std::nullopt;
}

void AppendMultiSz(std::wstring& buffer, const std::wstring& path)
{
    buffer.append(path);
    buffer.push_back(L'\0');
}

std::wstring BuildMultiSz(const std::vector<std::filesystem::path>& paths)
{
    std::wstring buffer;
    for (const auto& p : paths)
    {
        AppendMultiSz(buffer, p.c_str());
    }
    buffer.push_back(L'\0');
    return buffer;
}

HRESULT
BuildPathArrayArena(const std::vector<std::filesystem::path>& paths, FileSystemArenaOwner& arenaOwner, const wchar_t*** outPaths, unsigned long* outCount)
{
    if (! outPaths || ! outCount)
    {
        return E_POINTER;
    }

    *outPaths = nullptr;
    *outCount = 0;

    if (paths.empty())
    {
        return S_OK;
    }

    const uint64_t count64 = static_cast<uint64_t>(paths.size());
    if (count64 > std::numeric_limits<unsigned long>::max())
    {
        return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
    }

    const uint64_t arrayBytes64 = count64 * static_cast<uint64_t>(sizeof(const wchar_t*));
    if (arrayBytes64 > std::numeric_limits<unsigned long>::max())
    {
        return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
    }

    unsigned long totalBytes = static_cast<unsigned long>(arrayBytes64);

    for (const auto& path : paths)
    {
        const std::wstring& text = path.native();
        const size_t length      = text.size();
        if (length > (std::numeric_limits<unsigned long>::max() / sizeof(wchar_t)) - 1u)
        {
            return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
        }

        const unsigned long bytes = static_cast<unsigned long>((length + 1u) * sizeof(wchar_t));
        if (totalBytes > std::numeric_limits<unsigned long>::max() - bytes)
        {
            return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
        }
        totalBytes += bytes;
    }

    HRESULT hr = arenaOwner.Initialize(totalBytes);
    if (FAILED(hr))
    {
        return hr;
    }

    FileSystemArena* arena = arenaOwner.Get();
    auto* array            = static_cast<const wchar_t**>(
        AllocateFromFileSystemArena(arena, static_cast<unsigned long>(arrayBytes64), static_cast<unsigned long>(alignof(const wchar_t*))));
    if (! array)
    {
        return E_OUTOFMEMORY;
    }

    for (size_t index = 0; index < paths.size(); ++index)
    {
        const std::wstring& text  = paths[index].native();
        const size_t length       = text.size();
        const unsigned long bytes = static_cast<unsigned long>((length + 1u) * sizeof(wchar_t));
        auto* buffer              = static_cast<wchar_t*>(AllocateFromFileSystemArena(arena, bytes, static_cast<unsigned long>(alignof(wchar_t))));
        if (! buffer)
        {
            return E_OUTOFMEMORY;
        }

        if (length > 0)
        {
            ::CopyMemory(buffer, text.data(), length * sizeof(wchar_t));
        }
        buffer[length] = L'\0';
        array[index]   = buffer;
    }

    *outPaths = array;
    *outCount = static_cast<unsigned long>(count64);
    return S_OK;
}

std::filesystem::path GenerateShortcutPath(const std::filesystem::path& folder, const std::filesystem::path& target, int attempt)
{
    std::wstring stem = target.stem().wstring();
    if (stem.empty())
    {
        stem = target.filename().wstring();
    }
    std::wstring suffix;
    if (attempt > 0)
    {
        suffix = std::format(L" ({})", attempt + 1);
    }
    std::wstring candidate = std::format(L"{} - Shortcut{}.lnk", stem, suffix);
    return folder / candidate;
}

void ConfigureLabelLayout(IDWriteTextLayout* layout, IDWriteInlineObject* ellipsisSign, bool enableEllipsisTrimming = true)
{
    if (! layout)
    {
        return;
    }
    layout->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP);
    layout->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);
    layout->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_NEAR);
    DWRITE_TRIMMING trimming{};
    trimming.granularity = enableEllipsisTrimming ? DWRITE_TRIMMING_GRANULARITY_CHARACTER : DWRITE_TRIMMING_GRANULARITY_NONE;
    layout->SetTrimming(&trimming, enableEllipsisTrimming ? ellipsisSign : nullptr);
}

UINT PreferredDropEffectFormat()
{
    static const UINT format = RegisterClipboardFormatW(CFSTR_PREFERREDDROPEFFECT);
    return format;
}

UINT RedSalamanderInternalFileDropFormat()
{
    static const UINT format = RegisterClipboardFormatW(L"RedSalamander.InternalFileDrop.V1");
    return format;
}

class FormatEnumerator final : public IEnumFORMATETC
{
public:
    explicit FormatEnumerator(const std::vector<FORMATETC>& formats) : _refCount(1), _formats(formats)
    {
    }

    FormatEnumerator(const FormatEnumerator& other) : _refCount(1), _formats(other._formats), _index(other._index)
    {
    }

    // Explicitly delete assignment operators (COM objects are not assignable)
    FormatEnumerator& operator=(const FormatEnumerator&) = delete;
    FormatEnumerator(FormatEnumerator&&)                 = delete;
    FormatEnumerator& operator=(FormatEnumerator&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
    {
        if (! ppvObject)
        {
            return E_POINTER;
        }
        if (riid == IID_IUnknown || riid == IID_IEnumFORMATETC)
        {
            *ppvObject = static_cast<IEnumFORMATETC*>(this);
            AddRef();
            return S_OK;
        }
        *ppvObject = nullptr;
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() override
    {
        return ++_refCount;
    }

    ULONG STDMETHODCALLTYPE Release() override
    {
        ULONG remaining = --_refCount;
        if (remaining == 0)
        {
            delete this;
        }
        return remaining;
    }

    HRESULT STDMETHODCALLTYPE Next(ULONG celt, FORMATETC* rgelt, ULONG* pceltFetched) override
    {
        if (! rgelt)
        {
            return E_POINTER;
        }
        ULONG fetched = 0;
        while (fetched < celt && _index < _formats.size())
        {
            rgelt[fetched]     = _formats[_index];
            rgelt[fetched].ptd = nullptr;
            ++_index;
            ++fetched;
        }
        if (pceltFetched)
        {
            *pceltFetched = fetched;
        }
        return fetched == celt ? S_OK : S_FALSE;
    }

    HRESULT STDMETHODCALLTYPE Skip(ULONG celt) override
    {
        const size_t remaining = _formats.size() - std::min(_index, _formats.size());
        if (celt > remaining)
        {
            _index = _formats.size();
            return S_FALSE;
        }
        _index += celt;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Reset() override
    {
        _index = 0;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Clone(IEnumFORMATETC** ppenum) override
    {
        if (! ppenum)
        {
            return E_POINTER;
        }
        auto* clone = new (std::nothrow) FormatEnumerator(*this);
        if (! clone)
        {
            return E_OUTOFMEMORY;
        }
        clone->_index = _index;
        *ppenum       = clone;
        return S_OK;
    }

private:
    std::atomic<ULONG> _refCount;
    std::vector<FORMATETC> _formats;
    size_t _index = 0;
};

class FolderViewDataObject final : public IDataObject
{
public:
    FolderViewDataObject(
        std::vector<std::filesystem::path> paths, std::wstring pluginId, std::wstring instanceContext, DWORD preferredEffect, bool includeHDrop)
        : _refCount(1),
          _paths(std::move(paths)),
          _pluginId(std::move(pluginId)),
          _instanceContext(std::move(instanceContext)),
          _preferredEffect(preferredEffect),
          _includeHDrop(includeHDrop)
    {
    }

    // Explicitly delete copy/move operations (COM objects are not copyable/movable)
    FolderViewDataObject(const FolderViewDataObject&)            = delete;
    FolderViewDataObject(FolderViewDataObject&&)                 = delete;
    FolderViewDataObject& operator=(const FolderViewDataObject&) = delete;
    FolderViewDataObject& operator=(FolderViewDataObject&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
    {
        if (! ppvObject)
        {
            return E_POINTER;
        }
        if (riid == IID_IUnknown || riid == IID_IDataObject)
        {
            *ppvObject = static_cast<IDataObject*>(this);
            AddRef();
            return S_OK;
        }
        *ppvObject = nullptr;
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() override
    {
        return ++_refCount;
    }

    ULONG STDMETHODCALLTYPE Release() override
    {
        ULONG remaining = --_refCount;
        if (remaining == 0)
        {
            delete this;
        }
        return remaining;
    }

    HRESULT STDMETHODCALLTYPE GetData(FORMATETC* format, STGMEDIUM* medium) override
    {
        if (! format || ! medium)
        {
            return E_POINTER;
        }

        if ((format->tymed & TYMED_HGLOBAL) == 0)
        {
            return DV_E_TYMED;
        }

        if (format->cfFormat == static_cast<CLIPFORMAT>(RedSalamanderInternalFileDropFormat()))
        {
            auto data = CreateInternalFileDrop();
            if (! data)
            {
                return E_OUTOFMEMORY;
            }
            medium->tymed          = TYMED_HGLOBAL;
            medium->hGlobal        = data.release();
            medium->pUnkForRelease = nullptr;
            return S_OK;
        }

        if (format->cfFormat == CF_HDROP)
        {
            if (! _includeHDrop)
            {
                return DV_E_FORMATETC;
            }

            auto data = CreateHDrop();
            if (! data)
            {
                return E_OUTOFMEMORY;
            }
            medium->tymed          = TYMED_HGLOBAL;
            medium->hGlobal        = data.release();
            medium->pUnkForRelease = nullptr;
            return S_OK;
        }

        if (format->cfFormat == PreferredDropEffectFormat())
        {
            auto data = CreatePreferredEffect();
            if (! data)
            {
                return E_OUTOFMEMORY;
            }
            medium->tymed          = TYMED_HGLOBAL;
            medium->hGlobal        = data.release();
            medium->pUnkForRelease = nullptr;
            return S_OK;
        }

        return DV_E_FORMATETC;
    }

    HRESULT STDMETHODCALLTYPE GetDataHere(FORMATETC*, STGMEDIUM*) override
    {
        return DATA_E_FORMATETC;
    }

    HRESULT STDMETHODCALLTYPE QueryGetData(FORMATETC* format) override
    {
        if (! format)
        {
            return E_POINTER;
        }
        if ((format->tymed & TYMED_HGLOBAL) == 0)
        {
            return DV_E_TYMED;
        }
        if (format->cfFormat == static_cast<CLIPFORMAT>(RedSalamanderInternalFileDropFormat()) || format->cfFormat == PreferredDropEffectFormat())
        {
            return S_OK;
        }
        if (format->cfFormat == CF_HDROP)
        {
            return _includeHDrop ? S_OK : DV_E_FORMATETC;
        }
        return DV_E_FORMATETC;
    }

    HRESULT STDMETHODCALLTYPE GetCanonicalFormatEtc(FORMATETC*, FORMATETC* result) override
    {
        if (! result)
        {
            return E_POINTER;
        }
        *result = {};
        return DATA_S_SAMEFORMATETC;
    }

    HRESULT STDMETHODCALLTYPE SetData(FORMATETC*, STGMEDIUM*, BOOL) override
    {
        return E_NOTIMPL;
    }

    HRESULT STDMETHODCALLTYPE EnumFormatEtc(DWORD direction, IEnumFORMATETC** enumerator) override
    {
        if (! enumerator)
        {
            return E_POINTER;
        }
        *enumerator = nullptr;
        if (direction != DATADIR_GET)
        {
            return E_NOTIMPL;
        }

        std::vector<FORMATETC> formats;
        FORMATETC hdrop{};
        hdrop.dwAspect = DVASPECT_CONTENT;
        hdrop.lindex   = -1;
        hdrop.ptd      = nullptr;
        hdrop.tymed    = TYMED_HGLOBAL;

        hdrop.cfFormat = static_cast<CLIPFORMAT>(RedSalamanderInternalFileDropFormat());
        formats.push_back(hdrop);

        if (_includeHDrop)
        {
            hdrop.cfFormat = CF_HDROP;
            formats.push_back(hdrop);
        }

        hdrop.cfFormat = static_cast<CLIPFORMAT>(PreferredDropEffectFormat());
        formats.push_back(hdrop);

        auto* enumFormats = new (std::nothrow) FormatEnumerator(formats);
        if (! enumFormats)
        {
            return E_OUTOFMEMORY;
        }
        *enumerator = enumFormats;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE DAdvise(FORMATETC*, DWORD, IAdviseSink*, DWORD*) override
    {
        return OLE_E_ADVISENOTSUPPORTED;
    }

    HRESULT STDMETHODCALLTYPE DUnadvise(DWORD) override
    {
        return OLE_E_ADVISENOTSUPPORTED;
    }

    HRESULT STDMETHODCALLTYPE EnumDAdvise(IEnumSTATDATA**) override
    {
        return OLE_E_ADVISENOTSUPPORTED;
    }

private:
    wil::unique_hglobal CreateInternalFileDrop() const
    {
        struct Header
        {
            uint32_t version              = 1;
            uint32_t pluginIdChars        = 0;
            uint32_t instanceContextChars = 0;
            uint32_t pathCount            = 0;
        };

        const size_t pluginIdChars = _pluginId.size();
        const size_t instanceChars = _instanceContext.size();
        const size_t pathCount     = _paths.size();
        if (pluginIdChars > std::numeric_limits<uint32_t>::max() || instanceChars > std::numeric_limits<uint32_t>::max() ||
            pathCount > std::numeric_limits<uint32_t>::max())
        {
            return nullptr;
        }

        size_t totalBytes = sizeof(Header);
        auto addString    = [&](size_t chars) -> bool
        {
            const size_t add = (chars + 1u) * sizeof(wchar_t);
            if (totalBytes > (std::numeric_limits<size_t>::max)() - add)
            {
                return false;
            }
            totalBytes += add;
            return true;
        };

        if (! addString(pluginIdChars) || ! addString(instanceChars))
        {
            return nullptr;
        }

        for (const auto& path : _paths)
        {
            const std::wstring& text = path.native();
            const size_t chars       = text.size();
            if (chars > std::numeric_limits<uint32_t>::max())
            {
                return nullptr;
            }

            if (totalBytes > (std::numeric_limits<size_t>::max)() - sizeof(uint32_t))
            {
                return nullptr;
            }
            totalBytes += sizeof(uint32_t);

            if (! addString(chars))
            {
                return nullptr;
            }
        }

        wil::unique_hglobal data(GlobalAlloc(GHND, totalBytes));
        if (! data)
        {
            return nullptr;
        }

        void* raw = GlobalLock(data.get());
        if (! raw)
        {
            return nullptr;
        }
        // Copy the handle value now; don’t reference the local `data`.
        HGLOBAL h   = data.get();
        auto unlock = wil::scope_exit(
            [h]()
            {
                if (h)
                    GlobalUnlock(h);
            });

        auto* header                 = static_cast<Header*>(raw);
        header->version              = 1;
        header->pluginIdChars        = static_cast<uint32_t>(pluginIdChars);
        header->instanceContextChars = static_cast<uint32_t>(instanceChars);
        header->pathCount            = static_cast<uint32_t>(pathCount);

        std::byte* cursor = reinterpret_cast<std::byte*>(header + 1u);

        const auto writeString = [&](std::wstring_view text)
        {
            const size_t bytes = (text.size() + 1u) * sizeof(wchar_t);
            memcpy(cursor, text.data(), text.size() * sizeof(wchar_t));
            cursor += text.size() * sizeof(wchar_t);
            *reinterpret_cast<wchar_t*>(cursor) = L'\0';
            cursor += sizeof(wchar_t);
            static_cast<void>(bytes);
        };

        writeString(_pluginId);
        writeString(_instanceContext);

        for (const auto& path : _paths)
        {
            const std::wstring& text = path.native();
            const uint32_t chars32   = static_cast<uint32_t>(text.size());
            memcpy(cursor, &chars32, sizeof(chars32));
            cursor += sizeof(chars32);
            writeString(text);
        }

        return data;
    }

    wil::unique_hglobal CreateHDrop() const
    {
        size_t totalChars = 0;
        for (const auto& path : _paths)
        {
            totalChars += path.native().size() + 1;
        }
        totalChars += 1; // double-null terminator

        const size_t bytes = sizeof(DROPFILES) + totalChars * sizeof(wchar_t);
        wil::unique_hglobal memory(GlobalAlloc(GMEM_MOVEABLE, bytes));
        if (! memory)
        {
            return nullptr;
        }

        auto* dropFiles = static_cast<DROPFILES*>(GlobalLock(memory.get()));
        if (! dropFiles)
        {
            return nullptr;
        }

        dropFiles->pFiles = sizeof(DROPFILES);
        dropFiles->pt     = POINT{};
        dropFiles->fNC    = FALSE;
        dropFiles->fWide  = TRUE;

        auto* buffer = reinterpret_cast<wchar_t*>(reinterpret_cast<BYTE*>(dropFiles) + dropFiles->pFiles);
        for (const auto& path : _paths)
        {
            const std::wstring wide = path.native();
            std::copy(wide.begin(), wide.end(), buffer);
            buffer += wide.size();
            *buffer++ = L'\0';
        }
        *buffer = L'\0';
        GlobalUnlock(memory.get());
        return memory;
    }

    wil::unique_hglobal CreatePreferredEffect() const
    {
        wil::unique_hglobal memory(GlobalAlloc(GMEM_MOVEABLE, sizeof(DWORD)));
        if (! memory)
        {
            return nullptr;
        }
        auto* effect = static_cast<DWORD*>(GlobalLock(memory.get()));
        if (! effect)
        {
            return nullptr;
        }
        *effect = _preferredEffect;
        GlobalUnlock(memory.get());
        return memory;
    }

    std::atomic<ULONG> _refCount;
    std::vector<std::filesystem::path> _paths;
    std::wstring _pluginId;
    std::wstring _instanceContext;
    DWORD _preferredEffect = DROPEFFECT_COPY;
    bool _includeHDrop     = false;
};

class FolderViewDropSource final : public IDropSource
{
public:
    FolderViewDropSource() : _refCount(1)
    {
    }

    // Explicitly delete copy/move operations (COM objects are not copyable/movable)
    FolderViewDropSource(const FolderViewDropSource&)            = delete;
    FolderViewDropSource(FolderViewDropSource&&)                 = delete;
    FolderViewDropSource& operator=(const FolderViewDropSource&) = delete;
    FolderViewDropSource& operator=(FolderViewDropSource&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
    {
        if (! ppvObject)
        {
            return E_POINTER;
        }
        if (riid == IID_IUnknown || riid == IID_IDropSource)
        {
            *ppvObject = static_cast<IDropSource*>(this);
            AddRef();
            return S_OK;
        }
        *ppvObject = nullptr;
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() override
    {
        return ++_refCount;
    }

    ULONG STDMETHODCALLTYPE Release() override
    {
        ULONG remaining = --_refCount;
        if (remaining == 0)
        {
            delete this;
        }
        return remaining;
    }

    HRESULT STDMETHODCALLTYPE QueryContinueDrag(BOOL escapePressed, DWORD keyState) override
    {
        if (escapePressed)
        {
            return DRAGDROP_S_CANCEL;
        }
        if ((keyState & MK_LBUTTON) == 0)
        {
            return DRAGDROP_S_DROP;
        }
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GiveFeedback(DWORD) override
    {
        return DRAGDROP_S_USEDEFAULTCURSORS;
    }

private:
    std::atomic<ULONG> _refCount;
};

} // namespace

#pragma warning(pop)
