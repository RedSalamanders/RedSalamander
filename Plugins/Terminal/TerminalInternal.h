#pragma once

#include "Terminal.h"
#include "TerminalAccessibility.h"

#include <algorithm>
#include <array>
#include <atomic>
#include <chrono>
#include <clocale>
#include <cmath>
#include <cstdlib>
#include <cstring>
#include <format>
#include <limits>
#include <mutex>
#include <new>
#include <numeric>
#include <optional>
#include <ranges>
#include <set>
#include <stop_token>
#include <system_error>
#include <utility>
#include <vector>

#include <windowsx.h>
#include <wincrypt.h>
#include <imm.h>
#include <shlobj.h>
#include <shellapi.h>
#include <UIAutomation.h>
#include <uxtheme.h>

#include <yyjson.h>

#include "Helpers.h"
#include "DxUi/DxUi.h"
#include "PathUtils.h"
#include "PaneVisualState.h"
#include "ProcessCommandLine.h"
#include "StringConversion.h"
#include "UnicodeClipboard.h"
#include "WindowMessages.h"
#include "resource.h"

extern HINSTANCE g_hInstance;

namespace TerminalPluginDetail
{
extern std::atomic_uint32_t g_terminalObjectCount;
extern std::atomic_uint32_t g_terminalWindowCount;
extern std::mutex g_terminalWindowClassMutex;
extern bool g_terminalWindowClassRegistered;
extern std::mutex g_terminalSchemaMutex;
extern std::set<std::string> g_terminalSchemas;

constexpr wchar_t kTerminalWindowClass[] = L"RedSalamander.Terminal.Plugin.Window";
constexpr UINT_PTR kSelectionAutoScrollTimer = 0x366u;
constexpr UINT_PTR kKittyUploadRetryTimer = 0x367u;
constexpr uint8_t kTerminalModuleAnchor = 0u;
constexpr uint32_t kMinimumFormattedBytes = 1u * 1024u * 1024u;
constexpr uint32_t kMaximumFormattedBytes = 32u * 1024u * 1024u;
constexpr size_t kMaximumInputBytes = 8u * 1024u * 1024u;
constexpr size_t kMaximumInputRequests = 3840u;
constexpr uint64_t kKittyImageStorageBytes = TerminalKittyImagePipeline::MaximumSourceBytes;
constexpr uint64_t kMaximumKittyPixels = TerminalKittyImagePipeline::MaximumPixels;
constexpr size_t kMaximumKittyPlacements = 4096u;
constexpr size_t kMaximumKittyPlacementBytes = 1u * 1024u * 1024u;
constexpr size_t kMaximumKittyUploadBytesPerFrame = 64u * 1024u * 1024u;
constexpr uint32_t kMaximumKittyUploadsPerFrame = 1u;
constexpr uint32_t kMaximumKittyTransientUploadAttempts = 3u;
constexpr DWORD kInitialKittyUploadRetryDelayMilliseconds = 16u;
constexpr size_t kMaximumHyperlinkBytes = 8192u;
constexpr uint32_t kMinimumOsc52Bytes = 4096u;
constexpr uint32_t kMaximumOsc52Bytes = 256u * 1024u;
// Every key this build understands. Anything else in a stored configuration belongs to a different
// build and must be copied through SetConfiguration/GetConfiguration untouched.
constexpr std::array<std::string_view, 10u> kTerminalConfigurationKeys{{"defaultShell",
                                                                        "fontFamily",
                                                                        "fontSizeDip",
                                                                        "maxFormattedMiB",
                                                                        "pasteMaxBytes",
                                                                        "warnOnUnsafePaste",
                                                                        "followPathWhenIdle",
                                                                        "hyperlinkPolicy",
                                                                        "osc52Policy",
                                                                        "osc52MaxBytes"}};
constexpr size_t kMaximumFontFamilyChars = 128u;
constexpr float kFontNoticeMarginDip = 8.0f;
constexpr float kFontNoticeHeightDip = 28.0f;
constexpr float kFontNoticeDismissWidthDip = 80.0f;
constexpr float kFontNoticeMinimumWidthDip = 220.0f;
constexpr double kMinimumFontSizeDip = 8.0;
constexpr double kMaximumFontSizeDip = 32.0;
constexpr int64_t kMinimumFormattedMiB = 1;
constexpr int64_t kMaximumFormattedMiB = 32;
constexpr int64_t kMinimumPasteBytes = 4096;
constexpr size_t kGhosttyClipboardWriteMaxBytes = 256u * 1024u;
constexpr size_t kMaximumClipboardRepresentations = 16u;
constexpr size_t kMaximumIntegrationPathBytes = 128u * 1024u;
constexpr DWORD kIntegrationIoPollMilliseconds = 250u;
constexpr size_t kMaximumFindResultRows = 2'000u;
constexpr size_t kMaximumSuggestionRows = 100u;
constexpr size_t kMaximumPersistedHistoryBytes = 4u * 1024u * 1024u;
constexpr size_t kMaximumPersistedHistoryEntries = 10'000u;
constexpr size_t kMaximumSessionHistoryEntries = 1'000u;
constexpr size_t kMaximumCommandSurfaceQueryCharacters = 4'096u;

[[nodiscard]] constexpr bool IsTransientKittyUploadFailure(HRESULT hr) noexcept
{
    return hr == HRESULT_FROM_WIN32(ERROR_RETRY) || hr == D2DERR_DISPLAY_STATE_INVALID;
}

[[nodiscard]] constexpr bool IsTerminalPluginAction(std::wstring_view commandId) noexcept
{
    return IsTerminalPluginActionId(commandId);
}

[[nodiscard]] constexpr bool IsTerminalCommandSurfaceAction(std::wstring_view commandId) noexcept
{
    return commandId == L"cmd/terminal/find" || commandId == L"cmd/terminal/suggestions";
}

struct TerminalCommandSurfaceReadyPayload final
{
    uint64_t generation = 0u;
    uint64_t sessionGeneration = 0u;
    uint64_t snapshotGeneration = 0u;
    bool suggestions = false;
    bool snapshotOnly = false;
    std::shared_ptr<std::wstring> snapshot;
    TerminalCommandSurfaceModel::FindResult find;
    TerminalCommandSurfaceModel::SuggestionResult suggestion;
    std::shared_ptr<std::vector<std::wstring>> persistedHistory;
    uint64_t durationUs = 0u;
    uint64_t snapshotDurationUs = 0u;
    uint64_t historyLoadDurationUs = 0u;
};

struct IntegrationRequest final
{
    bool versioned = false;
    std::wstring currentDirectory;
    std::wstring historyPath;
    std::wstring promptBuffer;
    std::vector<std::wstring> currentSessionHistory;
};

inline void ScrubStrings(std::vector<std::wstring>& values) noexcept
{
    for (std::wstring& value : values)
    {
        SecureZeroMemory(value.data(), value.size() * sizeof(wchar_t));
    }
    values.clear();
}

inline void ScrubFindRows(std::vector<TerminalCommandSurfaceModel::FindMatch>& rows) noexcept
{
    for (auto& row : rows)
    {
        SecureZeroMemory(row.text.data(), row.text.size() * sizeof(wchar_t));
    }
    rows.clear();
}

inline void ScrubSuggestionRows(std::vector<TerminalCommandSurfaceModel::Suggestion>& rows) noexcept
{
    for (auto& row : rows)
    {
        SecureZeroMemory(row.text.data(), row.text.size() * sizeof(wchar_t));
    }
    rows.clear();
}

// Ghostty exposes an absolute row viewport rather than a pixel range, so this
// terminal-local adapter deliberately does not use the pixel-scroll helpers.
[[nodiscard]] inline SCROLLINFO MakeTerminalScrollInfo(const GhosttyTerminalScrollbar& scrollbar) noexcept
{
    const uint64_t total = std::max<uint64_t>(1u, scrollbar.total);
    const uint64_t visible = std::clamp<uint64_t>(scrollbar.len, 1u, total);
    const uint64_t maximumOffset = total - visible;
    const int nativeMaximum = static_cast<int>(std::min<uint64_t>(
        total - 1u, static_cast<uint64_t>((std::numeric_limits<int>::max)())));
    const UINT nativePage = static_cast<UINT>(std::min<uint64_t>(
        visible, static_cast<uint64_t>(nativeMaximum) + 1u));
    const int nativeMaximumOffset = std::max(0, nativeMaximum - static_cast<int>(nativePage - 1u));

    SCROLLINFO info{};
    info.cbSize = sizeof(info);
    info.fMask = SIF_RANGE | SIF_PAGE | SIF_POS | SIF_DISABLENOSCROLL;
    info.nMin = 0;
    info.nMax = nativeMaximum;
    info.nPage = nativePage;
    info.nPos = static_cast<int>(std::min<uint64_t>(
        std::min(scrollbar.offset, maximumOffset), static_cast<uint64_t>(nativeMaximumOffset)));
    return info;
}

struct TerminalWheelScrollResult final
{
    intptr_t rows = 0;
    int remainder = 0;
};

// Ghostty scrolls in terminal rows, so this terminal-local adapter preserves high-resolution
// wheel remainder and converts the configured Windows wheel policy directly to row deltas.
[[nodiscard]] constexpr TerminalWheelScrollResult ResolveTerminalWheelScroll(
    int previousRemainder, int wheelDelta, UINT linesPerDetent, uint16_t viewportRows) noexcept
{
    const int accumulatedDelta = previousRemainder + wheelDelta;
    const int detents          = accumulatedDelta / WHEEL_DELTA;
    const int remainder        = accumulatedDelta % WHEEL_DELTA;
    if (detents == 0 || linesPerDetent == 0u)
    {
        return {.remainder = remainder};
    }

    const uint64_t rowsPerDetent = linesPerDetent == WHEEL_PAGESCROLL
        ? static_cast<uint64_t>(std::max<uint16_t>(viewportRows, 1u))
        : static_cast<uint64_t>(linesPerDetent);
    const int64_t rowDelta = -static_cast<int64_t>(detents) * static_cast<int64_t>(rowsPerDetent);
    return {
        .rows = static_cast<intptr_t>(std::clamp<int64_t>(
            rowDelta,
            static_cast<int64_t>((std::numeric_limits<intptr_t>::min)()),
            static_cast<int64_t>((std::numeric_limits<intptr_t>::max)()))),
        .remainder = remainder,
    };
}


[[nodiscard]] inline int64_t SteadyTimestampNs() noexcept
{
    return std::chrono::duration_cast<std::chrono::nanoseconds>(
               std::chrono::steady_clock::now().time_since_epoch())
        .count();
}

constexpr wchar_t kPowerShellIntegrationScript[] = LR"powershell(
$ErrorActionPreference = 'Stop'
$rsPipe = '__PIPE_NAME__'
$rsLastPublishUtc = [DateTime]::MinValue
$rsUtf8 = [System.Text.UTF8Encoding]::new($false)
$rsEncode = { param([AllowEmptyString()][string]$rsValue) [Convert]::ToBase64String($rsUtf8.GetBytes($rsValue)) }.GetNewClosure()
$rsDecode = { param([AllowEmptyString()][string]$rsValue) $rsUtf8.GetString([Convert]::FromBase64String($rsValue)) }.GetNewClosure()
$rsPublish = {
    $ErrorActionPreference = 'Continue'
    try {
        $rsNow = [DateTime]::UtcNow
        if (($rsNow - $rsLastPublishUtc).TotalMilliseconds -lt 500) { return }
        $rsLastPublishUtc = $rsNow
        $rsPath = (Microsoft.PowerShell.Management\Get-Location).ProviderPath
        if ([string]::IsNullOrEmpty($rsPath) -or -not [System.IO.Path]::IsPathRooted($rsPath)) { return }
        $rsHistoryPath = ''
        $rsPromptBuffer = ''
        $rsSessionHistory = ''
        try {
            $rsOptionCommand = Microsoft.PowerShell.Core\Get-Command -Name Get-PSReadLineOption -ErrorAction SilentlyContinue
            if ($null -ne $rsOptionCommand) {
                $rsHistoryPath = [string]((Get-PSReadLineOption).HistorySavePath)
            }
        } catch [System.Management.Automation.RuntimeException] {
        }
        try {
            if ($null -ne ('Microsoft.PowerShell.PSConsoleReadLine' -as [type])) {
                [int]$rsCursor = 0
                [Microsoft.PowerShell.PSConsoleReadLine]::GetBufferState([ref]$rsPromptBuffer, [ref]$rsCursor)
            }
        } catch [System.Management.Automation.MethodInvocationException] {
            $rsPromptBuffer = ''
        } catch [System.InvalidOperationException] {
            $rsPromptBuffer = ''
        }
        try {
            $rsHistoryValues = [System.Collections.Generic.List[string]]::new()
            foreach ($rsEntry in @(Microsoft.PowerShell.Core\Get-History -Count 256)) {
                $rsCommand = [string]$rsEntry.CommandLine
                if (-not [string]::IsNullOrEmpty($rsCommand) -and $rsCommand.Length -le 32768) {
                    $rsHistoryValues.Add($rsCommand)
                }
            }
            while ($rsHistoryValues.Count -ne 0) {
                $rsSessionHistory = [string]::Join([char]0, $rsHistoryValues)
                if ($rsSessionHistory.Length -le 32768) { break }
                $rsHistoryValues.RemoveAt(0)
            }
        } catch [System.Management.Automation.RuntimeException] {
            $rsSessionHistory = ''
        }
        $rsRecord = 'RS2' + [char]9 + (& $rsEncode $rsPath) + [char]9 +
            (& $rsEncode $rsHistoryPath) + [char]9 + (& $rsEncode $rsPromptBuffer) + [char]9 +
            (& $rsEncode $rsSessionHistory)
        $rsClient = [System.IO.Pipes.NamedPipeClientStream]::new('.', $rsPipe, [System.IO.Pipes.PipeDirection]::InOut)
        try {
            $rsClient.Connect(1000)
            $rsWriter = [System.IO.StreamWriter]::new($rsClient, [System.Text.UTF8Encoding]::new($false), 1024, $true)
            try { $rsWriter.WriteLine($rsRecord); $rsWriter.Flush() } finally { $rsWriter.Dispose() }
            $rsReader = [System.IO.StreamReader]::new($rsClient, [System.Text.UTF8Encoding]::new($false), $false, 1024, $true)
            try { $rsResponse = $rsReader.ReadLine() } finally { $rsReader.Dispose() }
            $rsTarget = ''
            $rsReplacement = ''
            if (-not [string]::IsNullOrEmpty($rsResponse)) {
                $rsFields = $rsResponse.Split([char]9)
                if ($rsFields.Count -eq 3 -and $rsFields[0] -ceq 'RS2') {
                    $rsTarget = & $rsDecode $rsFields[1]
                    $rsReplacement = & $rsDecode $rsFields[2]
                }
            }
            if (-not [string]::IsNullOrEmpty($rsTarget)) {
                Microsoft.PowerShell.Management\Set-Location -LiteralPath $rsTarget
            }
            if (-not [string]::IsNullOrEmpty($rsReplacement)) {
                [Microsoft.PowerShell.PSConsoleReadLine]::RevertLine()
                [Microsoft.PowerShell.PSConsoleReadLine]::Insert($rsReplacement)
            }
        } finally { $rsClient.Dispose() }
    } catch [System.TimeoutException] {
    } catch [System.IO.IOException] {
    } catch [System.UnauthorizedAccessException] {
    } catch [System.FormatException] {
    } catch [System.Management.Automation.ItemNotFoundException] {
    } catch [System.Management.Automation.MethodInvocationException] {
    } catch [System.Management.Automation.ProviderInvocationException] {
    }
}.GetNewClosure()
& $rsPublish
Microsoft.PowerShell.Utility\Register-EngineEvent -SourceIdentifier PowerShell.OnIdle -Action $rsPublish | Microsoft.PowerShell.Core\Out-Null
)powershell";

inline HRESULT EncodePowerShellCommand(std::wstring_view pipeName, std::wstring& encoded) noexcept
{
    encoded.clear();
    if (pipeName.empty() || pipeName.find_first_not_of(L"abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789.-") !=
                                std::wstring_view::npos)
    {
        return E_INVALIDARG;
    }
    std::wstring script = kPowerShellIntegrationScript;
    constexpr std::wstring_view marker = L"__PIPE_NAME__";
    const size_t markerOffset = script.find(marker);
    if (markerOffset == std::wstring::npos)
    {
        return E_UNEXPECTED;
    }
    script.replace(markerOffset, marker.size(), pipeName);
    if (script.size() > (std::numeric_limits<DWORD>::max)() / sizeof(wchar_t))
    {
        return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
    }

    DWORD characterCount = 0u;
    const DWORD byteCount = static_cast<DWORD>(script.size() * sizeof(wchar_t));
    if (CryptBinaryToStringW(reinterpret_cast<const BYTE*>(script.data()),
                             byteCount,
                             CRYPT_STRING_BASE64 | CRYPT_STRING_NOCRLF,
                             nullptr,
                             &characterCount) == FALSE ||
        characterCount == 0u)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }
    encoded.resize(characterCount);
    DWORD writableCount = characterCount;
    if (CryptBinaryToStringW(reinterpret_cast<const BYTE*>(script.data()),
                             byteCount,
                             CRYPT_STRING_BASE64 | CRYPT_STRING_NOCRLF,
                             encoded.data(),
                             &writableCount) == FALSE)
    {
        encoded.clear();
        return HRESULT_FROM_WIN32(GetLastError());
    }
    if (writableCount != 0u && writableCount <= encoded.size() && encoded[writableCount - 1u] == L'\0')
    {
        encoded.resize(writableCount - 1u);
    }
    else
    {
        encoded.resize(writableCount);
    }
    return S_OK;
}

inline HRESULT MakeIntegrationPipeName(std::wstring& leafName, std::wstring& fullName) noexcept
{
    std::array<std::byte, 16u> random{};
    RETURN_IF_FAILED(Common::Crypto::GenerateRandomBytes(random));
    leafName = L"RedSalamander.Terminal.";
    Common::Crypto::AppendHexToken(leafName, random);
    fullName = L"\\\\.\\pipe\\" + leafName;
    return S_OK;
}

inline HRESULT WaitForOverlappedIo(HANDLE pipe, OVERLAPPED& overlapped, std::stop_token stopToken, DWORD& bytes) noexcept
{
    for (;;)
    {
        const DWORD wait = WaitForSingleObject(overlapped.hEvent, kIntegrationIoPollMilliseconds);
        if (wait == WAIT_OBJECT_0)
        {
            return GetOverlappedResult(pipe, &overlapped, &bytes, FALSE) != FALSE
                ? S_OK
                : HRESULT_FROM_WIN32(GetLastError());
        }
        if (wait == WAIT_FAILED)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }
        if (stopToken.stop_requested())
        {
            static_cast<void>(CancelIoEx(pipe, &overlapped));
            static_cast<void>(GetOverlappedResult(pipe, &overlapped, &bytes, TRUE));
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
    }
}

inline HRESULT ConnectIntegrationPipe(HANDLE pipe, std::stop_token stopToken) noexcept
{
    wil::unique_event_nothrow event;
    event.reset(CreateEventW(nullptr, TRUE, FALSE, nullptr));
    if (! event)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }
    OVERLAPPED overlapped{};
    overlapped.hEvent = event.get();
    if (ConnectNamedPipe(pipe, &overlapped) != FALSE)
    {
        return S_OK;
    }
    const DWORD error = GetLastError();
    if (error == ERROR_PIPE_CONNECTED)
    {
        return S_OK;
    }
    if (error != ERROR_IO_PENDING)
    {
        return HRESULT_FROM_WIN32(error);
    }
    DWORD ignored = 0u;
    return WaitForOverlappedIo(pipe, overlapped, stopToken, ignored);
}

inline HRESULT ReadIntegrationPath(HANDLE pipe, std::stop_token stopToken, std::string& utf8) noexcept
{
    utf8.clear();
    std::array<char, 4096u> chunk{};
    while (utf8.size() <= kMaximumIntegrationPathBytes)
    {
        wil::unique_event_nothrow event;
        event.reset(CreateEventW(nullptr, TRUE, FALSE, nullptr));
        if (! event)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }
        OVERLAPPED overlapped{};
        overlapped.hEvent = event.get();
        DWORD bytes = 0u;
        if (ReadFile(pipe, chunk.data(), static_cast<DWORD>(chunk.size()), &bytes, &overlapped) == FALSE)
        {
            const DWORD error = GetLastError();
            if (error != ERROR_IO_PENDING)
            {
                return HRESULT_FROM_WIN32(error);
            }
            RETURN_IF_FAILED(WaitForOverlappedIo(pipe, overlapped, stopToken, bytes));
        }
        if (bytes == 0u || utf8.size() > kMaximumIntegrationPathBytes - bytes)
        {
            return HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW);
        }
        utf8.append(chunk.data(), bytes);
        if (const size_t newline = utf8.find('\n'); newline != std::string::npos)
        {
            if (newline + 1u != utf8.size())
            {
                return E_INVALIDARG;
            }
            utf8.resize(newline);
            if (! utf8.empty() && utf8.back() == '\r')
            {
                utf8.pop_back();
            }
            return utf8.empty() ? E_INVALIDARG : S_OK;
        }
    }
    return HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW);
}

inline HRESULT EncodeIntegrationField(std::wstring_view value, std::string& encoded) noexcept
{
    encoded.clear();
    const std::optional<std::string> utf8 = Common::Strings::TryUtf8FromUtf16Strict(value);
    if (! utf8.has_value() || utf8->size() > static_cast<size_t>((std::numeric_limits<DWORD>::max)()))
    {
        return E_INVALIDARG;
    }
    if (utf8->empty())
    {
        return S_OK;
    }
    DWORD encodedCharacters = 0u;
    if (CryptBinaryToStringA(reinterpret_cast<const BYTE*>(utf8->data()),
                             static_cast<DWORD>(utf8->size()),
                             CRYPT_STRING_BASE64 | CRYPT_STRING_NOCRLF,
                             nullptr,
                             &encodedCharacters) == FALSE ||
        encodedCharacters == 0u)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }
    encoded.resize(encodedCharacters);
    DWORD written = encodedCharacters;
    if (CryptBinaryToStringA(reinterpret_cast<const BYTE*>(utf8->data()),
                             static_cast<DWORD>(utf8->size()),
                             CRYPT_STRING_BASE64 | CRYPT_STRING_NOCRLF,
                             encoded.data(),
                             &written) == FALSE)
    {
        encoded.clear();
        return HRESULT_FROM_WIN32(GetLastError());
    }
    if (written != 0u && written <= encoded.size() && encoded[written - 1u] == '\0')
    {
        --written;
    }
    encoded.resize(written);
    return S_OK;
}

inline HRESULT DecodeIntegrationField(std::string_view encoded, std::wstring& value) noexcept
{
    value.clear();
    if (encoded.empty())
    {
        return S_OK;
    }
    if (encoded.size() > static_cast<size_t>((std::numeric_limits<DWORD>::max)()))
    {
        return E_INVALIDARG;
    }
    DWORD byteCount = 0u;
    if (CryptStringToBinaryA(encoded.data(),
                             static_cast<DWORD>(encoded.size()),
                             CRYPT_STRING_BASE64,
                             nullptr,
                             &byteCount,
                             nullptr,
                             nullptr) == FALSE ||
        byteCount == 0u || byteCount > kMaximumIntegrationPathBytes)
    {
        return E_INVALIDARG;
    }
    std::string utf8(byteCount, '\0');
    DWORD written = byteCount;
    if (CryptStringToBinaryA(encoded.data(),
                             static_cast<DWORD>(encoded.size()),
                             CRYPT_STRING_BASE64,
                             reinterpret_cast<BYTE*>(utf8.data()),
                             &written,
                             nullptr,
                             nullptr) == FALSE)
    {
        return E_INVALIDARG;
    }
    utf8.resize(written);
    std::optional<std::wstring> decoded = Common::Strings::TryUtf16FromUtf8Strict(utf8);
    if (! decoded.has_value())
    {
        return E_INVALIDARG;
    }
    value = std::move(decoded.value());
    SecureZeroMemory(utf8.data(), utf8.size());
    return S_OK;
}

inline HRESULT WriteIntegrationReply(HANDLE pipe,
                              std::stop_token stopToken,
                              bool versioned,
                              std::wstring_view target,
                              std::wstring_view promptReplacement) noexcept
{
    std::string utf8;
    if (versioned)
    {
        std::string encodedTarget;
        std::string encodedReplacement;
        RETURN_IF_FAILED(EncodeIntegrationField(target, encodedTarget));
        RETURN_IF_FAILED(EncodeIntegrationField(promptReplacement, encodedReplacement));
        utf8.reserve(5u + encodedTarget.size() + encodedReplacement.size());
        utf8.append("RS2\t");
        utf8.append(encodedTarget);
        utf8.push_back('\t');
        utf8.append(encodedReplacement);
    }
    else if (! target.empty())
    {
        std::optional<std::string> converted = Common::Strings::TryUtf8FromUtf16Strict(target);
        if (! converted.has_value() || converted->empty())
        {
            return E_INVALIDARG;
        }
        utf8 = std::move(converted.value());
    }
    if (utf8.size() > kMaximumIntegrationPathBytes)
    {
        return HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW);
    }
    utf8.push_back('\n');

    size_t offset = 0u;
    while (offset < utf8.size())
    {
        wil::unique_event_nothrow event;
        event.reset(CreateEventW(nullptr, TRUE, FALSE, nullptr));
        if (! event)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }
        OVERLAPPED overlapped{};
        overlapped.hEvent = event.get();
        DWORD bytes = 0u;
        const DWORD requested = static_cast<DWORD>(utf8.size() - offset);
        if (WriteFile(pipe, utf8.data() + offset, requested, &bytes, &overlapped) == FALSE)
        {
            const DWORD error = GetLastError();
            if (error != ERROR_IO_PENDING)
            {
                return HRESULT_FROM_WIN32(error);
            }
            RETURN_IF_FAILED(WaitForOverlappedIo(pipe, overlapped, stopToken, bytes));
        }
        if (bytes == 0u)
        {
            return HRESULT_FROM_WIN32(ERROR_WRITE_FAULT);
        }
        offset += bytes;
    }
    return S_OK;
}

[[nodiscard]] inline bool IsSpanWellFormed(TerminalUtf16Span span) noexcept
{
    constexpr uint32_t kMaximumTerminalSpanCodeUnits = 32767u;
    return span.length <= kMaximumTerminalSpanCodeUnits && (span.length == 0u || span.data != nullptr);
}

[[nodiscard]] inline bool ContainsCommandControl(std::wstring_view value) noexcept
{
    return std::ranges::any_of(value, [](wchar_t character) noexcept { return character < L' '; });
}

[[nodiscard]] inline bool IsSupportedAbsoluteWindowsPath(std::wstring_view value) noexcept
{
    if (value.empty() || value.size() > 32767u || ContainsCommandControl(value))
    {
        return false;
    }
    const Common::Paths::WindowsPathClass pathClass = Common::Paths::ClassifyWindowsPath(value);
    return pathClass == Common::Paths::WindowsPathClass::DriveAbsolute ||
        pathClass == Common::Paths::WindowsPathClass::Unc ||
        pathClass == Common::Paths::WindowsPathClass::ExtendedDriveAbsolute ||
        pathClass == Common::Paths::WindowsPathClass::ExtendedUnc;
}

inline HRESULT ParseIntegrationRequest(std::string_view record, IntegrationRequest& request) noexcept
{
    request = {};
    if (! record.starts_with("RS2\t"))
    {
        std::optional<std::wstring> path = Common::Strings::TryUtf16FromUtf8Strict(record);
        if (! path.has_value() || ! IsSupportedAbsoluteWindowsPath(path.value()))
        {
            return E_INVALIDARG;
        }
        request.currentDirectory = std::move(path.value());
        return S_OK;
    }

    std::array<std::string_view, 5u> fields{};
    size_t start = 0u;
    for (size_t index = 0u; index < fields.size(); ++index)
    {
        const size_t end = record.find('\t', start);
        if (index + 1u == fields.size())
        {
            if (end != std::string_view::npos)
            {
                return E_INVALIDARG;
            }
            fields[index] = record.substr(start);
        }
        else
        {
            if (end == std::string_view::npos)
            {
                return E_INVALIDARG;
            }
            fields[index] = record.substr(start, end - start);
            start = end + 1u;
        }
    }
    if (fields[0] != "RS2")
    {
        return E_INVALIDARG;
    }

    std::wstring packedHistory;
    RETURN_IF_FAILED(DecodeIntegrationField(fields[1], request.currentDirectory));
    RETURN_IF_FAILED(DecodeIntegrationField(fields[2], request.historyPath));
    RETURN_IF_FAILED(DecodeIntegrationField(fields[3], request.promptBuffer));
    RETURN_IF_FAILED(DecodeIntegrationField(fields[4], packedHistory));
    const auto scrubPackedHistory = wil::scope_exit([&packedHistory]() noexcept
    {
        SecureZeroMemory(packedHistory.data(), packedHistory.size() * sizeof(wchar_t));
    });
    if (! IsSupportedAbsoluteWindowsPath(request.currentDirectory) ||
        (! request.historyPath.empty() && ! IsSupportedAbsoluteWindowsPath(request.historyPath)) ||
        request.promptBuffer.size() > kMaximumCommandSurfaceQueryCharacters * 8u)
    {
        return E_INVALIDARG;
    }

    size_t startCharacter = 0u;
    while (startCharacter <= packedHistory.size())
    {
        const size_t endCharacter = packedHistory.find(L'\0', startCharacter);
        const size_t commandEnd = endCharacter == std::wstring::npos ? packedHistory.size() : endCharacter;
        std::wstring command(packedHistory.data() + startCharacter, commandEnd - startCharacter);
        if (TerminalCommandSurfaceModel::IsSafeSuggestion(command))
        {
            request.currentSessionHistory.push_back(std::move(command));
            if (request.currentSessionHistory.size() > kMaximumSessionHistoryEntries)
            {
                request.currentSessionHistory.erase(request.currentSessionHistory.begin());
            }
        }
        if (endCharacter == std::wstring::npos)
        {
            break;
        }
        startCharacter = endCharacter + 1u;
    }
    request.versioned = true;
    return S_OK;
}

[[nodiscard]] inline bool IsSupportedAbsoluteWslPath(std::wstring_view value) noexcept
{
    return ! value.empty() && value.size() <= 32767u && value.front() == L'/' && ! ContainsCommandControl(value);
}

[[nodiscard]] inline std::wstring CopyTerminalSpan(TerminalUtf16Span span)
{
    return span.data != nullptr && span.length != 0u ? std::wstring(span.data, span.length) : std::wstring{};
}

[[nodiscard]] inline std::wstring CopyTerminalLocationPath(const TerminalLogicalLocation& location)
{
    std::wstring value = CopyTerminalSpan(location.windowsPath);
    if (! value.empty())
    {
        return value;
    }
    value = CopyTerminalSpan(location.wslAbsolutePath);
    return value.empty() ? CopyTerminalSpan(location.pluginBackingWindowsPath) : value;
}

[[nodiscard]] inline bool IsValidWslDistribution(std::wstring_view value) noexcept
{
    return ! value.empty() && value.size() <= 256u && value.front() != L'-' &&
        value.find_first_of(L"\\/") == std::wstring_view::npos && ! ContainsCommandControl(value);
}

[[nodiscard]] inline std::wstring_view ViewTerminalSpan(TerminalUtf16Span span) noexcept
{
    return span.length == 0u ? std::wstring_view{} : std::wstring_view(span.data, span.length);
}

[[nodiscard]] inline bool ValidateTerminalLogicalLocation(const TerminalLogicalLocation& location) noexcept
{
    if (location.sizeBytes < sizeof(TerminalLogicalLocation) || ! IsSpanWellFormed(location.windowsPath) ||
        ! IsSpanWellFormed(location.wslDistribution) || ! IsSpanWellFormed(location.wslAbsolutePath) ||
        ! IsSpanWellFormed(location.pluginShortId) || ! IsSpanWellFormed(location.pluginBackingWindowsPath))
    {
        return false;
    }

    const std::wstring_view windowsPath = ViewTerminalSpan(location.windowsPath);
    const std::wstring_view distribution = ViewTerminalSpan(location.wslDistribution);
    const std::wstring_view wslPath = ViewTerminalSpan(location.wslAbsolutePath);
    const std::wstring_view pluginShortId = ViewTerminalSpan(location.pluginShortId);
    const std::wstring_view pluginBackingPath = ViewTerminalSpan(location.pluginBackingWindowsPath);
    const bool noWindows = windowsPath.empty();
    const bool noWsl = distribution.empty() && wslPath.empty();
    const bool noPlugin = pluginShortId.empty() && pluginBackingPath.empty();

    switch (location.kind)
    {
    case TerminalLocationKind::WindowsLocal:
    {
        const Common::Paths::WindowsPathClass pathClass = Common::Paths::ClassifyWindowsPath(windowsPath);
        return noWsl && noPlugin && ! ContainsCommandControl(windowsPath) &&
            (pathClass == Common::Paths::WindowsPathClass::DriveAbsolute ||
             pathClass == Common::Paths::WindowsPathClass::ExtendedDriveAbsolute);
    }
    case TerminalLocationKind::WindowsUnc:
    {
        const Common::Paths::WindowsPathClass pathClass = Common::Paths::ClassifyWindowsPath(windowsPath);
        return noWsl && noPlugin && ! ContainsCommandControl(windowsPath) &&
            (pathClass == Common::Paths::WindowsPathClass::Unc ||
             pathClass == Common::Paths::WindowsPathClass::ExtendedUnc);
    }
    case TerminalLocationKind::Wsl:
        return noWindows && noPlugin && IsValidWslDistribution(distribution) && IsSupportedAbsoluteWslPath(wslPath);
    case TerminalLocationKind::PluginBacked:
        return noWindows && noWsl && ! pluginShortId.empty() && pluginShortId.size() <= 256u &&
            ! ContainsCommandControl(pluginShortId) &&
            (pluginBackingPath.empty() || IsSupportedAbsoluteWindowsPath(pluginBackingPath));
    case TerminalLocationKind::Unsupported:
        return noWindows && noWsl && noPlugin;
    default:
        return false;
    }
}

[[nodiscard]] inline std::wstring ResolvePathInsertionValue(const TerminalPathInsertion& insertion,
                                                     bool integrationTrusted,
                                                     bool idleAtPrimaryPrompt,
                                                     bool hasPendingUserInput,
                                                     std::wstring_view trustedCurrentDirectory)
{
    if (insertion.mode == TerminalPathInsertionMode::CurrentDirectoryFull)
    {
        return CopyTerminalLocationPath(insertion.initiatingSourceLocation);
    }
    if (insertion.mode == TerminalPathInsertionMode::ContextualLeafOrFull && integrationTrusted &&
        idleAtPrimaryPrompt && ! hasPendingUserInput)
    {
        const std::wstring parent = CopyTerminalSpan(insertion.parentLocation.windowsPath);
        if (! parent.empty() && OrdinalString::EqualsNoCase(trustedCurrentDirectory, std::wstring_view(parent)))
        {
            const std::wstring leaf = CopyTerminalSpan(insertion.displayLeaf);
            if (! leaf.empty())
            {
                return leaf;
            }
        }
    }
    return CopyTerminalLocationPath(insertion.itemLocation);
}

[[nodiscard]] inline bool AddLocalizedSchemaString(
    yyjson_mut_doc* document, yyjson_mut_val* object, const char* key, UINT resourceId) noexcept
{
    const std::wstring localized = LoadStringResource(g_hInstance, resourceId);
    const std::optional<std::string> utf8 = Common::Strings::TryUtf8FromUtf16Strict(localized);
    return utf8.has_value() && ! utf8->empty() &&
        yyjson_mut_obj_add_strncpy(document, object, key, utf8->data(), utf8->size());
}

[[nodiscard]] inline yyjson_mut_val* AddSchemaField(yyjson_mut_doc* document,
                                             yyjson_mut_val* fields,
                                             const char* key,
                                             const char* type,
                                             UINT labelResourceId,
                                             UINT descriptionResourceId) noexcept
{
    yyjson_mut_val* field = yyjson_mut_obj(document);
    if (field == nullptr || ! yyjson_mut_arr_add_val(fields, field) ||
        ! yyjson_mut_obj_add_strcpy(document, field, "key", key) ||
        ! yyjson_mut_obj_add_strcpy(document, field, "type", type) ||
        ! AddLocalizedSchemaString(document, field, "label", labelResourceId) ||
        ! AddLocalizedSchemaString(document, field, "description", descriptionResourceId))
    {
        return nullptr;
    }
    return field;
}

[[nodiscard]] inline bool AddSchemaOption(
    yyjson_mut_doc* document, yyjson_mut_val* options, const char* value, UINT labelResourceId) noexcept
{
    yyjson_mut_val* option = yyjson_mut_obj(document);
    return option != nullptr && yyjson_mut_arr_add_val(options, option) &&
        yyjson_mut_obj_add_strcpy(document, option, "value", value) &&
        AddLocalizedSchemaString(document, option, "label", labelResourceId);
}

[[nodiscard]] inline yyjson_mut_val* AddSchemaOptions(yyjson_mut_doc* document, yyjson_mut_val* field) noexcept
{
    yyjson_mut_val* options = yyjson_mut_arr(document);
    return options != nullptr && yyjson_mut_obj_add_val(document, field, "options", options) ? options : nullptr;
}

[[nodiscard]] inline std::string BuildTerminalConfigurationSchema() noexcept
{
    yyjson_mut_doc* document = yyjson_mut_doc_new(nullptr);
    if (document == nullptr)
    {
        return {};
    }
    const auto freeDocument = wil::scope_exit([&] noexcept { yyjson_mut_doc_free(document); });
    yyjson_mut_val* root = yyjson_mut_obj(document);
    yyjson_mut_val* fields = yyjson_mut_arr(document);
    if (root == nullptr || fields == nullptr)
    {
        return {};
    }
    yyjson_mut_doc_set_root(document, root);
    if (! yyjson_mut_obj_add_uint(document, root, "version", 1u) ||
        ! AddLocalizedSchemaString(document, root, "title", IDS_TERMINAL_SCHEMA_TITLE) ||
        ! yyjson_mut_obj_add_val(document, root, "fields", fields))
    {
        return {};
    }

    yyjson_mut_val* field = AddSchemaField(document,
                                           fields,
                                           "defaultShell",
                                           "option",
                                           IDS_TERMINAL_SCHEMA_DEFAULT_SHELL_LABEL,
                                           IDS_TERMINAL_SCHEMA_DEFAULT_SHELL_DESCRIPTION);
    yyjson_mut_val* options = field == nullptr ? nullptr : AddSchemaOptions(document, field);
    if (field == nullptr || options == nullptr || ! yyjson_mut_obj_add_strcpy(document, field, "default", "auto") ||
        ! AddSchemaOption(document, options, "auto", IDS_TERMINAL_SCHEMA_OPTION_AUTO) ||
        ! AddSchemaOption(document, options, "powershell", IDS_TERMINAL_SCHEMA_OPTION_POWERSHELL) ||
        ! AddSchemaOption(document, options, "cmd", IDS_TERMINAL_SCHEMA_OPTION_COMMAND_PROMPT))
    {
        return {};
    }

    field = AddSchemaField(document,
                           fields,
                           "fontFamily",
                           "text",
                           IDS_TERMINAL_SCHEMA_FONT_FAMILY_LABEL,
                           IDS_TERMINAL_SCHEMA_FONT_FAMILY_DESCRIPTION);
    if (field == nullptr || ! yyjson_mut_obj_add_strcpy(document, field, "default", "Cascadia Mono"))
    {
        return {};
    }

    const auto addValueField = [document, fields](const char* key,
                                                   UINT labelResourceId,
                                                   UINT descriptionResourceId,
                                                   uint64_t defaultValue,
                                                   uint64_t minimum,
                                                   uint64_t maximum) noexcept
    {
        yyjson_mut_val* valueField =
            AddSchemaField(document, fields, key, "value", labelResourceId, descriptionResourceId);
        return valueField != nullptr && yyjson_mut_obj_add_uint(document, valueField, "default", defaultValue) &&
            yyjson_mut_obj_add_uint(document, valueField, "min", minimum) &&
            yyjson_mut_obj_add_uint(document, valueField, "max", maximum);
    };
    if (! addValueField("fontSizeDip", IDS_TERMINAL_SCHEMA_FONT_SIZE_LABEL, IDS_TERMINAL_SCHEMA_FONT_SIZE_DESCRIPTION, 14u, 8u, 32u) ||
        ! addValueField("maxFormattedMiB", IDS_TERMINAL_SCHEMA_FORMAT_MEMORY_LABEL, IDS_TERMINAL_SCHEMA_FORMAT_MEMORY_DESCRIPTION, 8u, 1u, 32u) ||
        ! addValueField("pasteMaxBytes", IDS_TERMINAL_SCHEMA_PASTE_BYTES_LABEL, IDS_TERMINAL_SCHEMA_PASTE_BYTES_DESCRIPTION, 4194304u, 4096u, 8388608u))
    {
        return {};
    }

    const auto addBooleanOptionField = [document, fields](const char* key,
                                                           UINT labelResourceId,
                                                           UINT descriptionResourceId,
                                                           const char* defaultValue) noexcept
    {
        yyjson_mut_val* optionField =
            AddSchemaField(document, fields, key, "option", labelResourceId, descriptionResourceId);
        yyjson_mut_val* optionValues = optionField == nullptr ? nullptr : AddSchemaOptions(document, optionField);
        return optionField != nullptr && optionValues != nullptr &&
            yyjson_mut_obj_add_strcpy(document, optionField, "default", defaultValue) &&
            AddSchemaOption(document, optionValues, "0", IDS_TERMINAL_SCHEMA_OPTION_OFF) &&
            AddSchemaOption(document, optionValues, "1", IDS_TERMINAL_SCHEMA_OPTION_ON);
    };
    if (! addBooleanOptionField("warnOnUnsafePaste",
                                IDS_TERMINAL_SCHEMA_WARN_PASTE_LABEL,
                                IDS_TERMINAL_SCHEMA_WARN_PASTE_DESCRIPTION,
                                "1") ||
        ! addBooleanOptionField("followPathWhenIdle",
                                IDS_TERMINAL_SCHEMA_FOLLOW_PATH_LABEL,
                                IDS_TERMINAL_SCHEMA_FOLLOW_PATH_DESCRIPTION,
                                "0"))
    {
        return {};
    }

    field = AddSchemaField(document,
                           fields,
                           "hyperlinkPolicy",
                           "option",
                           IDS_TERMINAL_SCHEMA_HYPERLINK_LABEL,
                           IDS_TERMINAL_SCHEMA_HYPERLINK_DESCRIPTION);
    options = field == nullptr ? nullptr : AddSchemaOptions(document, field);
    if (field == nullptr || options == nullptr || ! yyjson_mut_obj_add_strcpy(document, field, "default", "ask") ||
        ! AddSchemaOption(document, options, "disabled", IDS_TERMINAL_SCHEMA_OPTION_DISABLED) ||
        ! AddSchemaOption(document, options, "ask", IDS_TERMINAL_SCHEMA_OPTION_ASK_OPEN) ||
        ! AddSchemaOption(document, options, "open", IDS_TERMINAL_SCHEMA_OPTION_OPEN_CTRL_CLICK))
    {
        return {};
    }

    field = AddSchemaField(document,
                           fields,
                           "osc52Policy",
                           "option",
                           IDS_TERMINAL_SCHEMA_OSC52_LABEL,
                           IDS_TERMINAL_SCHEMA_OSC52_DESCRIPTION);
    options = field == nullptr ? nullptr : AddSchemaOptions(document, field);
    if (field == nullptr || options == nullptr || ! yyjson_mut_obj_add_strcpy(document, field, "default", "ask") ||
        ! AddSchemaOption(document, options, "deny", IDS_TERMINAL_SCHEMA_OPTION_DENY) ||
        ! AddSchemaOption(document, options, "ask", IDS_TERMINAL_SCHEMA_OPTION_ASK_WRITE) ||
        ! AddSchemaOption(document, options, "allow", IDS_TERMINAL_SCHEMA_OPTION_ALLOW) ||
        ! addValueField("osc52MaxBytes",
                        IDS_TERMINAL_SCHEMA_OSC52_BYTES_LABEL,
                        IDS_TERMINAL_SCHEMA_OSC52_BYTES_DESCRIPTION,
                        262144u,
                        4096u,
                        kMaximumOsc52Bytes))
    {
        return {};
    }

    char* written = yyjson_mut_write(document, YYJSON_WRITE_NOFLAG, nullptr);
    if (written == nullptr)
    {
        return {};
    }
    const auto freeWritten = wil::scope_exit([&] noexcept { free(written); });
    return std::string(written);
}

[[nodiscard]] inline bool IsExistingExecutable(std::wstring_view path) noexcept
{
    if (! IsSupportedAbsoluteWindowsPath(path))
    {
        return false;
    }
    const DWORD attributes = GetFileAttributesW(std::wstring(path).c_str());
    return attributes != INVALID_FILE_ATTRIBUTES && (attributes & FILE_ATTRIBUTE_DIRECTORY) == 0u;
}

[[nodiscard]] inline std::wstring GetSystemExecutable(std::wstring_view relativePath) noexcept
{
    std::array<wchar_t, MAX_PATH> systemDirectory{};
    const UINT copied = GetSystemDirectoryW(systemDirectory.data(), static_cast<UINT>(systemDirectory.size()));
    if (copied == 0u || copied >= systemDirectory.size())
    {
        return {};
    }
    std::wstring path(systemDirectory.data(), copied);
    path.push_back(L'\\');
    path.append(relativePath);
    return IsExistingExecutable(path) ? path : std::wstring{};
}

[[nodiscard]] inline std::wstring ReadPowerShellAppPath(HKEY root) noexcept
{
    constexpr wchar_t kPwshAppPath[] = L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\pwsh.exe";
    DWORD bytes = 0u;
    if (RegGetValueW(root, kPwshAppPath, nullptr, RRF_RT_REG_SZ | RRF_NOEXPAND, nullptr, nullptr, &bytes) != ERROR_SUCCESS ||
        bytes < sizeof(wchar_t) || bytes > 32768u * sizeof(wchar_t))
    {
        return {};
    }
    std::wstring path(bytes / sizeof(wchar_t), L'\0');
    if (RegGetValueW(root, kPwshAppPath, nullptr, RRF_RT_REG_SZ | RRF_NOEXPAND, nullptr, path.data(), &bytes) != ERROR_SUCCESS)
    {
        return {};
    }
    path.resize(wcsnlen_s(path.c_str(), path.size()));
    return IsExistingExecutable(path) ? path : std::wstring{};
}

[[nodiscard]] inline std::wstring FindPowerShell7() noexcept
{
    if (std::wstring path = ReadPowerShellAppPath(HKEY_CURRENT_USER); ! path.empty())
    {
        return path;
    }
    if (std::wstring path = ReadPowerShellAppPath(HKEY_LOCAL_MACHINE); ! path.empty())
    {
        return path;
    }
    wil::unique_cotaskmem_string programFiles;
    if (FAILED(SHGetKnownFolderPath(FOLDERID_ProgramFiles, KF_FLAG_DEFAULT, nullptr, programFiles.put())))
    {
        return {};
    }
    std::wstring path(programFiles.get());
    path.append(L"\\PowerShell\\7\\pwsh.exe");
    return IsExistingExecutable(path) ? path : std::wstring{};
}

[[nodiscard]] inline bool ReadBoolean(yyjson_val* value, bool& output) noexcept
{
    if (value == nullptr)
    {
        return true;
    }
    if (yyjson_is_bool(value))
    {
        output = yyjson_get_bool(value);
        return true;
    }
    if (! yyjson_is_str(value))
    {
        return false;
    }
    const char* text = yyjson_get_str(value);
    if (text == nullptr)
    {
        return false;
    }
    if (strcmp(text, "1") == 0 || strcmp(text, "true") == 0 || strcmp(text, "on") == 0)
    {
        output = true;
        return true;
    }
    if (strcmp(text, "0") == 0 || strcmp(text, "false") == 0 || strcmp(text, "off") == 0)
    {
        output = false;
        return true;
    }
    return false;
}

// Terminal paste is deliberately plugin-local: unlike general clipboard UI it
// performs one bounded read and rejects unterminated data before any shell
// bytes are admitted.
[[nodiscard]] inline std::optional<std::wstring> ReadBoundedClipboardText(HWND ownerWindow, size_t maximumCodeUnits) noexcept
{
    if (maximumCodeUnits == 0u || OpenClipboard(ownerWindow) == FALSE)
    {
        return std::nullopt;
    }
    const auto closeClipboard = wil::scope_exit([]() noexcept { static_cast<void>(CloseClipboard()); });
    const HANDLE handle = GetClipboardData(CF_UNICODETEXT);
    if (handle == nullptr)
    {
        return std::nullopt;
    }
    const SIZE_T storageBytes = GlobalSize(handle);
    if (storageBytes < sizeof(wchar_t))
    {
        return std::nullopt;
    }
    const auto* text = static_cast<const wchar_t*>(GlobalLock(handle));
    if (text == nullptr)
    {
        return std::nullopt;
    }
    const auto unlock = wil::scope_exit([handle]() noexcept { static_cast<void>(GlobalUnlock(handle)); });
    const size_t boundedCapacity = std::min<size_t>(storageBytes / sizeof(wchar_t), maximumCodeUnits + 1u);
    const wchar_t* terminator = std::char_traits<wchar_t>::find(text, boundedCapacity, L'\0');
    if (terminator == nullptr)
    {
        return std::nullopt;
    }
    return std::wstring(text, static_cast<size_t>(terminator - text));
}

// IME buffers are window-local transient input and have a much smaller hard
// bound than paste or terminal output.
[[nodiscard]] inline std::optional<std::wstring> ReadImeString(HWND window, LPARAM value) noexcept
{
    constexpr LONG kMaximumImeBytes = 64 * 1024;
    HIMC inputContext = ImmGetContext(window);
    if (inputContext == nullptr)
    {
        return std::nullopt;
    }
    const auto releaseContext = wil::scope_exit(
        [window, inputContext]() noexcept { static_cast<void>(ImmReleaseContext(window, inputContext)); });
    const LONG bytes = ImmGetCompositionStringW(inputContext, static_cast<DWORD>(value), nullptr, 0u);
    if (bytes <= 0 || bytes > kMaximumImeBytes || (bytes % static_cast<LONG>(sizeof(wchar_t))) != 0)
    {
        return std::nullopt;
    }
    std::wstring text(static_cast<size_t>(bytes) / sizeof(wchar_t), L'\0');
    return ImmGetCompositionStringW(inputContext, static_cast<DWORD>(value), text.data(), static_cast<DWORD>(bytes)) == bytes
        ? std::optional<std::wstring>(std::move(text))
        : std::nullopt;
}

[[nodiscard]] inline uint32_t ArgbFromGhostty(const GhosttyColorRgb& color, uint8_t alpha = 0xFFu) noexcept
{
    return (static_cast<uint32_t>(alpha) << 24u) | (static_cast<uint32_t>(color.r) << 16u) |
        (static_cast<uint32_t>(color.g) << 8u) | static_cast<uint32_t>(color.b);
}

[[nodiscard]] inline uint32_t WithAlpha(uint32_t argb, uint8_t alpha) noexcept
{
    return (argb & 0x00FFFFFFu) | (static_cast<uint32_t>(alpha) << 24u);
}

[[nodiscard]] inline GhosttyKey GhosttyKeyFromVirtualKey(WPARAM virtualKey) noexcept
{
    switch (virtualKey)
    {
    case VK_RETURN:
        return GHOSTTY_KEY_ENTER;
    case VK_ESCAPE:
        return GHOSTTY_KEY_ESCAPE;
    case VK_UP:
        return GHOSTTY_KEY_ARROW_UP;
    case VK_DOWN:
        return GHOSTTY_KEY_ARROW_DOWN;
    case VK_LEFT:
        return GHOSTTY_KEY_ARROW_LEFT;
    case VK_RIGHT:
        return GHOSTTY_KEY_ARROW_RIGHT;
    case VK_HOME:
        return GHOSTTY_KEY_HOME;
    case VK_END:
        return GHOSTTY_KEY_END;
    case VK_INSERT:
        return GHOSTTY_KEY_INSERT;
    case VK_DELETE:
        return GHOSTTY_KEY_DELETE;
    case VK_PRIOR:
        return GHOSTTY_KEY_PAGE_UP;
    case VK_NEXT:
        return GHOSTTY_KEY_PAGE_DOWN;
    default:
        if (virtualKey >= VK_F1 && virtualKey <= VK_F24)
        {
            return static_cast<GhosttyKey>(
                static_cast<int>(GHOSTTY_KEY_F1) + static_cast<int>(virtualKey - VK_F1));
        }
        return GHOSTTY_KEY_UNIDENTIFIED;
    }
}

[[nodiscard]] constexpr bool VirtualKeyProducesTranslatedCharacter(WPARAM virtualKey) noexcept
{
    if ((virtualKey >= static_cast<WPARAM>('0') && virtualKey <= static_cast<WPARAM>('9')) ||
        (virtualKey >= static_cast<WPARAM>('A') && virtualKey <= static_cast<WPARAM>('Z')) ||
        (virtualKey >= VK_NUMPAD0 && virtualKey <= VK_DIVIDE) ||
        (virtualKey >= VK_OEM_1 && virtualKey <= VK_OEM_3) ||
        (virtualKey >= VK_OEM_4 && virtualKey <= VK_OEM_8))
    {
        return true;
    }

    switch (virtualKey)
    {
    case VK_RETURN:
    case VK_TAB:
    case VK_ESCAPE:
    case VK_BACK:
    case VK_SPACE:
    case VK_OEM_102:
        return true;
    default:
        return false;
    }
}

inline void MarkTranslatedCharacterConsumed(bool& suppressionPending) noexcept
{
    suppressionPending = true;
}

[[nodiscard]] inline bool ConsumeTranslatedCharacterSuppression(bool& suppressionPending) noexcept
{
    return std::exchange(suppressionPending, false);
}

[[nodiscard]] constexpr uint8_t ReportedMouseButtonBit(GhosttyMouseButton button) noexcept
{
    if (button == GHOSTTY_MOUSE_BUTTON_LEFT)
    {
        return 0x1u;
    }
    if (button == GHOSTTY_MOUSE_BUTTON_MIDDLE)
    {
        return 0x2u;
    }
    return button == GHOSTTY_MOUSE_BUTTON_RIGHT ? 0x4u : 0u;
}

[[nodiscard]] inline bool HasReportedMouseButton(uint8_t buttons, GhosttyMouseButton button) noexcept
{
    const uint8_t bit = ReportedMouseButtonBit(button);
    return bit != 0u && (buttons & bit) != 0u;
}

[[nodiscard]] inline bool AddReportedMouseButton(uint8_t& buttons, GhosttyMouseButton button) noexcept
{
    const uint8_t bit = ReportedMouseButtonBit(button);
    if (bit == 0u || (buttons & bit) != 0u)
    {
        return false;
    }
    buttons |= bit;
    return true;
}

[[nodiscard]] inline bool RemoveReportedMouseButton(uint8_t& buttons, GhosttyMouseButton button) noexcept
{
    const uint8_t bit = ReportedMouseButtonBit(button);
    if (bit == 0u || (buttons & bit) == 0u)
    {
        return false;
    }
    buttons &= static_cast<uint8_t>(~bit);
    return true;
}

constexpr std::array<GhosttyMouseButton, 3u> kTrackedMouseButtons{{
    GHOSTTY_MOUSE_BUTTON_LEFT,
    GHOSTTY_MOUSE_BUTTON_MIDDLE,
    GHOSTTY_MOUSE_BUTTON_RIGHT,
}};

struct GhosttyClipboardDataView final
{
    const uint8_t* data = nullptr;
    size_t length = 0u;
};

// Ghostty ABI-local by design: the staged Product and Candidate request tails
// differ, and only callback-lifetime borrowed fields may be inspected here.
[[nodiscard]] constexpr bool GhosttySizedFieldPresent(
    size_t recordSize, size_t fieldOffset, size_t fieldSize) noexcept
{
    return fieldOffset <= recordSize && fieldSize <= recordSize - fieldOffset;
}

[[nodiscard]] inline GhosttyClipboardWriteResult ValidateGhosttyClipboardWrite(
    const GhosttyClipboardWrite* write, size_t maximumBytes, GhosttyClipboardDataView& selectedView) noexcept
{
    selectedView = {};
    constexpr size_t minimumSize = offsetof(GhosttyClipboardWrite, contents_len) +
        sizeof(size_t);
    if (write == nullptr || write->size < minimumSize ||
        write->location != GHOSTTY_CLIPBOARD_LOCATION_STANDARD ||
        write->contents_len > kMaximumClipboardRepresentations ||
        (write->contents_len != 0u && write->contents == nullptr))
    {
        return GHOSTTY_CLIPBOARD_WRITE_RESULT_UNSUPPORTED;
    }

    const GhosttyClipboardContent* selected = nullptr;
    for (size_t index = 0u; index < write->contents_len; ++index)
    {
        const GhosttyClipboardContent& candidate = write->contents[index];
        if ((candidate.mime.ptr == nullptr && candidate.mime.len != 0u) ||
            (candidate.data.ptr == nullptr && candidate.data.len != 0u))
        {
            return GHOSTTY_CLIPBOARD_WRITE_RESULT_INVALID_DATA;
        }
        const std::string_view mime(
            candidate.mime.ptr != nullptr ? reinterpret_cast<const char*>(candidate.mime.ptr) : "", candidate.mime.len);
        if (mime == "text/plain" || mime == "text/plain;charset=utf-8")
        {
            selected = &candidate;
            break;
        }
    }
    if (write->contents_len != 0u && selected == nullptr)
    {
        return GHOSTTY_CLIPBOARD_WRITE_RESULT_UNSUPPORTED;
    }

    selectedView.length = selected != nullptr ? selected->data.len : 0u;
    selectedView.data = selected != nullptr ? selected->data.ptr : nullptr;
    const std::string_view selectedText(
        selectedView.data != nullptr ? reinterpret_cast<const char*>(selectedView.data) : "",
        selectedView.length);
    if (selectedView.length > maximumBytes ||
        (selectedView.data != nullptr && std::memchr(selectedView.data, 0, selectedView.length) != nullptr) ||
        ! Common::Strings::IsValidUtf8Strict(selectedText))
    {
        selectedView = {};
        return GHOSTTY_CLIPBOARD_WRITE_RESULT_INVALID_DATA;
    }
    return GHOSTTY_CLIPBOARD_WRITE_RESULT_SUCCESS;
}

#ifdef ENABLE_TESTS
struct RuntimeClipboardCapture final
{
    bool invoked = false;
    uint32_t callbackCount = 0u;
    uint32_t replyCount = 0u;
    GhosttyClipboardWriteResult replyResult = GHOSTTY_CLIPBOARD_WRITE_RESULT_MAX_VALUE;
    GhosttyClipboardLocation location = GHOSTTY_CLIPBOARD_LOCATION_MAX_VALUE;
    std::array<uint8_t, 32u> data{};
    size_t length = 0u;
    std::array<uint8_t, 256u> ptyData{};
    size_t ptyLength = 0u;
};

inline void CaptureRuntimeClipboard(
    GhosttyTerminal /*terminal*/, void* userData, const GhosttyClipboardWrite* write) noexcept
{
    auto* capture = static_cast<RuntimeClipboardCapture*>(userData);
    GhosttyClipboardWriteResult result = GHOSTTY_CLIPBOARD_WRITE_RESULT_INVALID_DATA;
    if (capture != nullptr)
    {
        ++capture->callbackCount;
    }
    if (capture != nullptr && write != nullptr &&
        write->size >= offsetof(GhosttyClipboardWrite, contents_len) + sizeof(size_t) &&
        write->contents != nullptr && write->contents_len == 1u)
    {
        const GhosttyClipboardContent& content = write->contents[0];
        if (content.data.ptr != nullptr && content.data.len <= capture->data.size())
        {
            capture->invoked = true;
            capture->location = write->location;
            capture->length = content.data.len;
            std::memcpy(capture->data.data(), content.data.ptr, content.data.len);
            result = GHOSTTY_CLIPBOARD_WRITE_RESULT_SUCCESS;
        }
    }
    if (capture != nullptr && write != nullptr &&
        GhosttySizedFieldPresent(write->size, offsetof(GhosttyClipboardWrite, reply), sizeof(write->reply)) &&
        write->reply != nullptr)
    {
        GhosttyClipboardWriteReply reply{};
        reply.size = sizeof(reply);
        reply.result = result;
        reply.remember = false;
        write->reply(write, &reply);
        ++capture->replyCount;
        capture->replyResult = result;
    }
}

inline void CaptureRuntimePty(
    GhosttyTerminal /*terminal*/, void* userData, const uint8_t* data, size_t length) noexcept
{
    auto* capture = static_cast<RuntimeClipboardCapture*>(userData);
    if (capture == nullptr || data == nullptr || length > capture->ptyData.size() - capture->ptyLength)
    {
        return;
    }
    std::memcpy(capture->ptyData.data() + capture->ptyLength, data, length);
    capture->ptyLength += length;
}

inline void CountKittyPipelineNotification(void* context) noexcept
{
    if (auto* count = static_cast<std::atomic_uint32_t*>(context); count != nullptr)
    {
        count->fetch_add(1u, std::memory_order_acq_rel);
    }
}
#endif

} // namespace TerminalPluginDetail
