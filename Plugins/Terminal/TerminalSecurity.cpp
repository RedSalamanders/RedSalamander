#include "Terminal.h"
#include "TerminalAccessibility.h"
#include "TerminalInternal.h"

#include <algorithm>
#include <array>
#include <chrono>
#include <clocale>
#include <cmath>
#include <cstdlib>
#include <cstring>
#include <format>
#include <limits>
#include <new>
#include <numeric>
#include <optional>
#include <ranges>
#include <set>
#include <system_error>
#include <utility>
#include <vector>

#include <UIAutomation.h>
#include <imm.h>
#include <shellapi.h>
#include <shlobj.h>
#include <uxtheme.h>
#include <wincrypt.h>
#include <windowsx.h>

#pragma comment(lib, "crypt32")
#pragma comment(lib, "imm32")
#pragma comment(lib, "advapi32")
#pragma comment(lib, "shell32")
#pragma comment(lib, "uxtheme")

#include <yyjson.h>

#include "Helpers.h"
#include "PaneVisualState.h"
#include "PathUtils.h"
#include "ProcessCommandLine.h"
#include "StringConversion.h"
#include "UnicodeClipboard.h"
#include "WindowMessages.h"
#include "resource.h"
#include <DxUi/DxUi.h>

extern HINSTANCE g_hInstance;

using namespace TerminalPluginDetail;

#include "TerminalSecurity.h"

GhosttyClipboardWriteResult Terminal::handleGhosttyClipboardWrite(const ::GhosttyClipboardWrite* write) noexcept
{
    if (_closing.load(std::memory_order_acquire))
    {
        return GHOSTTY_CLIPBOARD_WRITE_RESULT_DENIED;
    }
    if (_osc52Policy.load(std::memory_order_acquire) == Osc52Policy::Deny)
    {
        return GHOSTTY_CLIPBOARD_WRITE_RESULT_DENIED;
    }
    GhosttyClipboardDataView selected{};
    const GhosttyClipboardWriteResult validation = ValidateGhosttyClipboardWrite(write, _osc52MaxBytes.load(std::memory_order_acquire), selected);
    if (validation != GHOSTTY_CLIPBOARD_WRITE_RESULT_SUCCESS)
    {
        return validation;
    }

    bool expected = false;
    if (! _osc52PostPending.compare_exchange_strong(expected, true, std::memory_order_acq_rel, std::memory_order_acquire))
    {
        return GHOSTTY_CLIPBOARD_WRITE_RESULT_BUSY;
    }

    auto payload = std::unique_ptr<Osc52ClipboardPayload>(new (std::nothrow) Osc52ClipboardPayload());
    if (! payload)
    {
        _osc52PostPending.store(false, std::memory_order_release);
        return GHOSTTY_CLIPBOARD_WRITE_RESULT_IO_ERROR;
    }
    if (selected.length != 0u)
    {
        payload->utf8.assign(reinterpret_cast<const char*>(selected.data), selected.length);
    }
    const HWND hwnd = _windowHandle.load(std::memory_order_acquire);
    if (hwnd == nullptr || ! PostMessagePayload(hwnd, WndMsg::kTerminalClipboardWrite, 0u, std::move(payload)))
    {
        _osc52PostPending.store(false, std::memory_order_release);
        return GHOSTTY_CLIPBOARD_WRITE_RESULT_IO_ERROR;
    }
#if defined(ENABLE_TESTS)
    _debugOsc52PayloadPostCount.fetch_add(1u, std::memory_order_acq_rel);
#endif
    // OSC 52 and OSC 1337 discard this reply, but the ABI still defines
    // SUCCESS as a completed write. The UI-thread post is only queued here, so
    // report BUSY rather than claiming that consent or clipboard I/O succeeded.
    return GHOSTTY_CLIPBOARD_WRITE_RESULT_BUSY;
}

void Terminal::GhosttyClipboardWrite(GhosttyTerminal /*terminal*/, void* userData, const ::GhosttyClipboardWrite* write) noexcept
{
    if (write == nullptr || ! GhosttySizedFieldPresent(write->size, offsetof(::GhosttyClipboardWrite, reply), sizeof(write->reply)) || write->reply == nullptr)
    {
        return;
    }
    auto* self                               = static_cast<Terminal*>(userData);
    const GhosttyClipboardWriteResult result = self != nullptr ? self->handleGhosttyClipboardWrite(write) : GHOSTTY_CLIPBOARD_WRITE_RESULT_DENIED;
#if defined(ENABLE_TESTS)
    if (self != nullptr)
    {
        self->_debugClipboardCallbackCount.fetch_add(1u, std::memory_order_acq_rel);
        self->_debugClipboardReplyResult.store(result, std::memory_order_release);
    }
#endif
    GhosttyClipboardWriteReply reply{};
    reply.size     = sizeof(reply);
    reply.result   = result;
    reply.remember = false;
    write->reply(write, &reply);
#if defined(ENABLE_TESTS)
    if (self != nullptr)
    {
        self->_debugClipboardReplyCount.fetch_add(1u, std::memory_order_acq_rel);
    }
#endif
}

bool Terminal::pasteClipboard() noexcept
{
    const HWND hwnd = _windowHandle.load(std::memory_order_acquire);
    if (hwnd == nullptr || _ghosttyTerminal == nullptr || _runtime.pasteIsSafe == nullptr || _runtime.pasteEncode == nullptr || _runtime.terminalGet == nullptr)
    {
        return false;
    }

    std::optional<std::wstring> clipboard = ReadBoundedClipboardText(hwnd, _config.pasteMaxBytes);
    if (! clipboard.has_value() || clipboard.value().empty())
    {
        return false;
    }

    std::wstring normalized;
    normalized.reserve(clipboard.value().size());
    for (size_t index = 0u; index < clipboard.value().size(); ++index)
    {
        const wchar_t character = clipboard.value()[index];
        if (character == L'\r')
        {
            if (index + 1u < clipboard.value().size() && clipboard.value()[index + 1u] == L'\n')
            {
                ++index;
            }
            normalized.push_back(L'\n');
        }
        else
        {
            normalized.push_back(character);
        }
    }

    std::optional<std::string> converted = Common::Strings::TryUtf8FromUtf16Strict(normalized);
    SecureZeroMemory(clipboard.value().data(), clipboard.value().size() * sizeof(wchar_t));
    clipboard.reset();
    SecureZeroMemory(normalized.data(), normalized.size() * sizeof(wchar_t));
    normalized.clear();
    if (! converted.has_value() || converted.value().empty() || converted.value().size() > _config.pasteMaxBytes)
    {
        return false;
    }

    std::string source = std::move(converted.value());
    converted.reset();
    const auto wipeSource = wil::scope_exit([&source]() noexcept { SecureZeroMemory(source.data(), source.size()); });
    if (_config.warnOnUnsafePaste && ! _runtime.pasteIsSafe(source.data(), source.size()))
    {
        if (confirmationVisible())
        {
            MessageBeep(MB_ICONWARNING);
            return false;
        }
        _pendingUnsafePaste             = std::move(source);
        _unsafePasteConfirmationVisible = true;
        publishAccessibilitySnapshot();
        InvalidateRect(hwnd, nullptr, FALSE);
        return true;
    }

    return encodePasteSource(std::move(source));
}

bool Terminal::encodePasteSource(std::string source) noexcept
{
    const auto wipeSource = wil::scope_exit([&source]() noexcept { SecureZeroMemory(source.data(), source.size()); });
    if (source.empty() || source.size() > _config.pasteMaxBytes)
    {
        return false;
    }

    bool bracketed = false;
    {
        std::scoped_lock lock(_terminalMutex);
        if (_ghosttyTerminal == nullptr || _runtime.GetTerminalMode(_ghosttyTerminal, GHOSTTY_MODE_BRACKETED_PASTE, bracketed) != GHOSTTY_SUCCESS)
        {
            return false;
        }
    }

    constexpr size_t kBracketedPasteWrapperBytes = 12u;
    if (source.size() > _config.pasteMaxBytes - std::min<size_t>(_config.pasteMaxBytes, kBracketedPasteWrapperBytes))
    {
        return false;
    }
    std::vector<char> encoded(source.size() + kBracketedPasteWrapperBytes);
    const auto wipeEncoded     = wil::scope_exit([&encoded]() noexcept { SecureZeroMemory(encoded.data(), encoded.size()); });
    size_t written             = 0u;
    const GhosttyResult result = _runtime.pasteEncode(source.data(), source.size(), bracketed, encoded.data(), encoded.size(), &written);
    if (result != GHOSTTY_SUCCESS || written == 0u || written > encoded.size() || written > _config.pasteMaxBytes)
    {
        return false;
    }
    if (! writeInput(std::string_view(encoded.data(), written)))
    {
        return false;
    }
    markUserInput(source);
    return true;
}

bool Terminal::confirmationVisible() const noexcept
{
    return _unsafePasteConfirmationVisible || _hyperlinkConfirmationVisible || _osc52ConfirmationVisible;
}

std::wstring Terminal::confirmationAccessibilityText() const
{
    if (! confirmationVisible())
    {
        return {};
    }
    const bool hyperlink        = _hyperlinkConfirmationVisible;
    const bool osc52            = _osc52ConfirmationVisible;
    const std::wstring& title   = osc52 ? _osc52Title : hyperlink ? _hyperlinkTitle : _unsafePasteTitle;
    const std::wstring& message = osc52 ? _osc52Message : hyperlink ? _hyperlinkMessage : _unsafePasteMessage;
    const std::wstring& accept  = osc52 ? _osc52Accept : hyperlink ? _hyperlinkAccept : _unsafePasteAccept;
    const std::wstring& cancel  = osc52 ? _osc52Cancel : hyperlink ? _hyperlinkCancel : _unsafePasteCancel;
    // Never expose pending OSC52 contents. The same visible labels become the
    // document's accessible prompt, including the keyboard actions.
    return FormatStringResource(g_hInstance, IDS_TERMINAL_CONFIRMATION_ACCESSIBILITY, title, message, accept, cancel);
}

void Terminal::resolveConfirmation(bool accept) noexcept
{
    if (! confirmationVisible())
    {
        return;
    }
    const bool paste                = _unsafePasteConfirmationVisible;
    const bool hyperlink            = _hyperlinkConfirmationVisible;
    const bool osc52                = _osc52ConfirmationVisible;
    _unsafePasteConfirmationVisible = false;
    _hyperlinkConfirmationVisible   = false;
    _osc52ConfirmationVisible       = false;
    std::string source;
    std::wstring uri;
    std::wstring clipboardText;
    if (paste)
    {
        source = std::move(_pendingUnsafePaste);
        SecureZeroMemory(_pendingUnsafePaste.data(), _pendingUnsafePaste.size());
        _pendingUnsafePaste.clear();
    }
    if (hyperlink)
    {
        uri = std::move(_pendingHyperlink);
        _pendingHyperlink.clear();
    }
    if (osc52)
    {
        clipboardText = std::move(_pendingOsc52Text);
        SecureZeroMemory(_pendingOsc52Text.data(), _pendingOsc52Text.size() * sizeof(wchar_t));
        _pendingOsc52Text.clear();
        _osc52PostPending.store(false, std::memory_order_release);
    }
    if (const HWND hwnd = _windowHandle.load(std::memory_order_acquire); hwnd != nullptr)
    {
        publishAccessibilitySnapshot();
        InvalidateRect(hwnd, nullptr, FALSE);
    }
    if (accept && paste)
    {
        static_cast<void>(encodePasteSource(std::move(source)));
    }
    if (accept && hyperlink)
    {
        static_cast<void>(openHyperlink(uri));
    }
    if (accept && osc52)
    {
        static_cast<void>(commitOsc52Clipboard(clipboardText));
    }
    SecureZeroMemory(source.data(), source.size());
    SecureZeroMemory(clipboardText.data(), clipboardText.size() * sizeof(wchar_t));
}

void Terminal::drawConfirmation(float widthDip, float heightDip) noexcept
{
    if (! confirmationVisible() || ! _renderTarget || ! _foregroundBrush || ! _overlayTextFormat || ! _overlayBoldTextFormat)
    {
        return;
    }
    const bool hyperlink            = _hyperlinkConfirmationVisible;
    const bool osc52                = _osc52ConfirmationVisible;
    const std::wstring& title       = osc52 ? _osc52Title : hyperlink ? _hyperlinkTitle : _unsafePasteTitle;
    const std::wstring& baseMessage = osc52 ? _osc52Message : hyperlink ? _hyperlinkMessage : _unsafePasteMessage;
    const std::wstring& acceptText  = osc52 ? _osc52Accept : hyperlink ? _hyperlinkAccept : _unsafePasteAccept;
    const std::wstring& cancelText  = osc52 ? _osc52Cancel : hyperlink ? _hyperlinkCancel : _unsafePasteCancel;
    std::wstring message            = baseMessage;
    if (hyperlink && ! _pendingHyperlink.empty())
    {
        message.append(L"\n");
        message.append(_pendingHyperlink.substr(0u, 1024u));
    }
    const D2D1_RECT_F bounds = D2D1::RectF(0.0f, 0.0f, widthDip, heightDip);
    _foregroundBrush->SetColor(DxUi::ColorFromArgb(WithAlpha(_backgroundArgb, 0xD0u)));
    _renderTarget->FillRectangle(bounds, _foregroundBrush.get());

    const float panelWidth   = std::min(560.0f, std::max(240.0f, widthDip - 32.0f));
    const float panelHeight  = std::min(190.0f, std::max(150.0f, heightDip - 32.0f));
    const float left         = std::max(0.0f, (widthDip - panelWidth) * 0.5f);
    const float top          = std::max(0.0f, (heightDip - panelHeight) * 0.5f);
    const D2D1_RECT_F panel  = D2D1::RectF(left, top, left + panelWidth, top + panelHeight);
    const uint32_t panelArgb = _themeHighContrast ? _backgroundArgb : (_themeDark ? _backgroundArgb : 0xFFF7F7F7u);
    _foregroundBrush->SetColor(DxUi::ColorFromArgb(panelArgb));
    _renderTarget->FillRectangle(panel, _foregroundBrush.get());
    _foregroundBrush->SetColor(DxUi::ColorFromArgb(_foregroundArgb));
    _renderTarget->DrawRectangle(panel, _foregroundBrush.get(), 1.0f);

    const D2D1_RECT_F titleRect = D2D1::RectF(left + 18.0f, top + 14.0f, panel.right - 18.0f, top + 44.0f);
    _renderTarget->DrawTextW(
        title.data(), static_cast<UINT32>(title.size()), _overlayBoldTextFormat.get(), titleRect, _foregroundBrush.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);
    const D2D1_RECT_F messageRect = D2D1::RectF(left + 18.0f, top + 48.0f, panel.right - 18.0f, panel.bottom - 58.0f);
    _renderTarget->DrawTextW(
        message.data(), static_cast<UINT32>(message.size()), _overlayTextFormat.get(), messageRect, _foregroundBrush.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);

    const D2D1_RECT_F accept = D2D1::RectF(panel.right - 274.0f, panel.bottom - 46.0f, panel.right - 146.0f, panel.bottom - 12.0f);
    const D2D1_RECT_F cancel = D2D1::RectF(panel.right - 138.0f, panel.bottom - 46.0f, panel.right - 10.0f, panel.bottom - 12.0f);
    _foregroundBrush->SetColor(DxUi::ColorFromArgb(_selectionBackgroundArgb));
    _renderTarget->FillRectangle(accept, _foregroundBrush.get());
    _foregroundBrush->SetColor(DxUi::ColorFromArgb(_selectionForegroundArgb));
    _renderTarget->DrawTextW(
        acceptText.data(), static_cast<UINT32>(acceptText.size()), _overlayTextFormat.get(), accept, _foregroundBrush.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);
    _foregroundBrush->SetColor(DxUi::ColorFromArgb(_foregroundArgb));
    _renderTarget->DrawRectangle(cancel, _foregroundBrush.get(), 1.0f);
    _renderTarget->DrawTextW(
        cancelText.data(), static_cast<UINT32>(cancelText.size()), _overlayTextFormat.get(), cancel, _foregroundBrush.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);
}

bool Terminal::handleConfirmationClick(POINT clientPoint) noexcept
{
    if (! confirmationVisible())
    {
        return false;
    }
    RECT client{};
    const HWND hwnd = _windowHandle.load(std::memory_order_acquire);
    if (hwnd == nullptr || GetClientRect(hwnd, &client) == FALSE)
    {
        return true;
    }
    const float scale       = 96.0f / static_cast<float>(std::max<UINT>(_dpi, 1u));
    const float widthDip    = static_cast<float>(client.right - client.left) * scale;
    const float heightDip   = static_cast<float>(client.bottom - client.top) * scale;
    const float panelWidth  = std::min(560.0f, std::max(240.0f, widthDip - 32.0f));
    const float panelHeight = std::min(190.0f, std::max(150.0f, heightDip - 32.0f));
    const float left        = std::max(0.0f, (widthDip - panelWidth) * 0.5f);
    const float top         = std::max(0.0f, (heightDip - panelHeight) * 0.5f);
    const D2D1_POINT_2F point{static_cast<float>(clientPoint.x) * scale, static_cast<float>(clientPoint.y) * scale};
    const D2D1_RECT_F accept = D2D1::RectF(left + panelWidth - 274.0f, top + panelHeight - 46.0f, left + panelWidth - 146.0f, top + panelHeight - 12.0f);
    const D2D1_RECT_F cancel = D2D1::RectF(left + panelWidth - 138.0f, top + panelHeight - 46.0f, left + panelWidth - 10.0f, top + panelHeight - 12.0f);
    if (point.x >= accept.left && point.x <= accept.right && point.y >= accept.top && point.y <= accept.bottom)
    {
        resolveConfirmation(true);
    }
    else if (point.x >= cancel.left && point.x <= cancel.right && point.y >= cancel.top && point.y <= cancel.bottom)
    {
        resolveConfirmation(false);
    }
    return true;
}

bool Terminal::openHyperlink(std::wstring_view uri) noexcept
{
    if (uri.empty() || uri.size() > kMaximumHyperlinkBytes ||
        (! OrdinalString::StartsWithNoCase(uri, L"http://") && ! OrdinalString::StartsWithNoCase(uri, L"https://") &&
         ! OrdinalString::StartsWithNoCase(uri, L"mailto:")))
    {
        return false;
    }
    for (const wchar_t character : uri)
    {
        if (character < 0x20 || character == 0x7F)
        {
            return false;
        }
    }
    const std::wstring owned(uri);
    const HWND hwnd        = _windowHandle.load(std::memory_order_acquire);
    const HINSTANCE result = ShellExecuteW(hwnd, L"open", owned.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
    return reinterpret_cast<INT_PTR>(result) > 32;
}

bool Terminal::commitOsc52Clipboard(std::wstring_view text) noexcept
{
    const HWND hwnd = _windowHandle.load(std::memory_order_acquire);
    return hwnd != nullptr && Common::Clipboard::TrySetUnicodeText(hwnd, text, Common::Clipboard::EmptyUnicodeTextPolicy::Allow);
}

void Terminal::handleOsc52Clipboard(std::unique_ptr<Osc52ClipboardPayload> payload) noexcept
{
    if (! payload || _closing.load(std::memory_order_acquire))
    {
        _osc52PostPending.store(false, std::memory_order_release);
        return;
    }
    std::optional<std::wstring> converted = Common::Strings::TryUtf16FromUtf8Strict(payload->utf8);
    if (! converted.has_value() || converted.value().find(L'\0') != std::wstring::npos)
    {
        _osc52PostPending.store(false, std::memory_order_release);
        MessageBeep(MB_ICONWARNING);
        return;
    }
    std::wstring text        = std::move(converted.value());
    const auto scrubText     = wil::scope_exit([&text]() noexcept { SecureZeroMemory(text.data(), text.size() * sizeof(wchar_t)); });
    const Osc52Policy policy = _osc52Policy.load(std::memory_order_acquire);
    if (policy == Osc52Policy::Deny)
    {
        _osc52PostPending.store(false, std::memory_order_release);
        return;
    }
    if (policy == Osc52Policy::Allow)
    {
        if (! commitOsc52Clipboard(text))
        {
            MessageBeep(MB_ICONWARNING);
        }
        _osc52PostPending.store(false, std::memory_order_release);
        return;
    }
    if (confirmationVisible())
    {
        _osc52PostPending.store(false, std::memory_order_release);
        MessageBeep(MB_ICONWARNING);
        return;
    }
    _pendingOsc52Text         = std::move(text);
    _osc52ConfirmationVisible = true;
    if (const HWND hwnd = _windowHandle.load(std::memory_order_acquire); hwnd != nullptr)
    {
        publishAccessibilitySnapshot();
        InvalidateRect(hwnd, nullptr, FALSE);
    }
}

bool Terminal::activateHyperlinkAt(POINT clientPoint) noexcept
{
    GhosttyGridRef cell{};
    cell.size = sizeof(cell);
    std::vector<uint8_t> bytes;
    size_t required = 0u;
    {
        std::scoped_lock lock(_terminalMutex);
        if (_ghosttyTerminal == nullptr || _runtime.terminalGridRef == nullptr || _runtime.gridRefHyperlinkUri == nullptr ||
            _runtime.terminalGridRef(_ghosttyTerminal, terminalPointFromClient(clientPoint), &cell) != GHOSTTY_SUCCESS)
        {
            return false;
        }
        const GhosttyResult query = _runtime.gridRefHyperlinkUri(&cell, nullptr, 0u, &required);
        if (query == GHOSTTY_SUCCESS && required == 0u)
        {
            return false;
        }
        if (query != GHOSTTY_OUT_OF_SPACE || required == 0u || required > kMaximumHyperlinkBytes)
        {
            MessageBeep(MB_ICONWARNING);
            return true;
        }
        bytes.resize(required);
        size_t written = 0u;
        if (_runtime.gridRefHyperlinkUri(&cell, bytes.data(), bytes.size(), &written) != GHOSTTY_SUCCESS || written == 0u || written > bytes.size())
        {
            return true;
        }
        bytes.resize(written);
    }

    const std::optional<std::wstring> converted =
        Common::Strings::TryUtf16FromUtf8Strict(std::string_view(reinterpret_cast<const char*>(bytes.data()), bytes.size()));
    if (! converted.has_value() || converted.value().empty())
    {
        MessageBeep(MB_ICONWARNING);
        return true;
    }
    const std::wstring& uri = converted.value();
    if (_config.hyperlinkPolicy == L"disabled")
    {
        MessageBeep(MB_ICONWARNING);
        return true;
    }
    if (_config.hyperlinkPolicy == L"ask")
    {
        if (confirmationVisible())
        {
            MessageBeep(MB_ICONWARNING);
            return true;
        }
        _pendingHyperlink             = uri;
        _hyperlinkConfirmationVisible = true;
        if (const HWND hwnd = _windowHandle.load(std::memory_order_acquire); hwnd != nullptr)
        {
            publishAccessibilitySnapshot();
            InvalidateRect(hwnd, nullptr, FALSE);
        }
        return true;
    }
    static_cast<void>(openHyperlink(uri));
    return true;
}

std::wstring Terminal::quotePath(std::wstring_view path) const
{
    if (_shellKind == ShellKind::CommandPrompt)
    {
        std::wstring quoted = L"\"";
        for (const wchar_t character : path)
        {
            if (character == L'%')
            {
                // cmd expands %NAME% even inside quotes. Move each percent
                // outside the quoted segment and caret-escape it; adjacent
                // segments still form one argument for a launched process.
                quoted.append(L"\"^%\"");
            }
            else
            {
                quoted.push_back(character);
            }
        }
        quoted.push_back(L'\"');
        return quoted;
    }
    std::wstring quoted = L"'";
    for (const wchar_t ch : path)
    {
        if (ch == L'\'' && _shellKind == ShellKind::Posix)
        {
            quoted.append(L"'\\''");
        }
        else
        {
            quoted.push_back(ch);
            if (ch == L'\'')
            {
                quoted.push_back(L'\'');
            }
        }
    }
    quoted.push_back(L'\'');
    return quoted;
}

HRESULT STDMETHODCALLTYPE Terminal::InsertPath(const TerminalPathInsertion* insertion) noexcept
{
    if (insertion == nullptr || insertion->sizeBytes < sizeof(TerminalPathInsertion) || ! validateLocation(insertion->initiatingSourceLocation) ||
        ! validateLocation(insertion->itemLocation) || ! validateLocation(insertion->parentLocation) || ! IsSpanWellFormed(insertion->displayLeaf) ||
        (insertion->mode != TerminalPathInsertionMode::ContextualLeafOrFull && insertion->mode != TerminalPathInsertionMode::AlwaysFull &&
         insertion->mode != TerminalPathInsertionMode::CurrentDirectoryFull))
    {
        return E_INVALIDARG;
    }
    const TerminalLifecycleState lifecycle = _lifecycle.load(std::memory_order_acquire);
    if (lifecycle != TerminalLifecycleState::Starting && lifecycle != TerminalLifecycleState::Running)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_STATE);
    }
    std::wstring value;
    {
        std::scoped_lock stateLock(_stateMutex);
        const std::wstring initiatingPath = CopyTerminalLocationPath(insertion->initiatingSourceLocation);
        const bool sourcePathMatches      = _sourceLocationKind == TerminalLocationKind::Wsl
                                                ? initiatingPath == _sourceLocationPath
                                                : OrdinalString::EqualsNoCase(std::wstring_view(initiatingPath), std::wstring_view(_sourceLocationPath));
        const bool sourceNamespaceMatches =
            (_sourceLocationKind != TerminalLocationKind::Wsl ||
             OrdinalString::EqualsNoCase(copySpan(insertion->initiatingSourceLocation.wslDistribution), _launchWslDistribution)) &&
            (_sourceLocationKind != TerminalLocationKind::PluginBacked ||
             OrdinalString::EqualsNoCase(copySpan(insertion->initiatingSourceLocation.pluginShortId), _sourcePluginShortId));
        if (insertion->initiatingSource.folderWindowInstanceId != _originalSource.folderWindowInstanceId ||
            insertion->initiatingSource.paneInstanceId != _originalSource.paneInstanceId || insertion->initiatingSourceGeneration != _sourceGeneration ||
            insertion->initiatingSourceLocation.kind != _sourceLocationKind || ! sourcePathMatches || ! sourceNamespaceMatches)
        {
            return E_INVALIDARG;
        }

        const auto compatibleLocation = [&](const TerminalLogicalLocation& location) noexcept
        {
            if (_shellKind == ShellKind::Posix)
            {
                return location.kind == TerminalLocationKind::Wsl && OrdinalString::EqualsNoCase(copySpan(location.wslDistribution), _launchWslDistribution) &&
                       IsSupportedAbsoluteWslPath(copySpan(location.wslAbsolutePath));
            }
            return (location.kind == TerminalLocationKind::WindowsLocal || location.kind == TerminalLocationKind::WindowsUnc) &&
                   IsSupportedAbsoluteWindowsPath(copySpan(location.windowsPath));
        };
        if (insertion->mode != TerminalPathInsertionMode::CurrentDirectoryFull &&
            (! compatibleLocation(insertion->itemLocation) || ! compatibleLocation(insertion->parentLocation)))
        {
            return E_INVALIDARG;
        }
        value = ResolvePathInsertionValue(*insertion, _integrationTrusted, _idleAtPrimaryPrompt, _hasPendingUserInput, _trustedCurrentDirectory);
    }
    if (value.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }
    if (ContainsCommandControl(value))
    {
        return E_INVALIDARG;
    }
    return writeTextInput(quotePath(value)) ? S_OK : HRESULT_FROM_WIN32(ERROR_NOT_ENOUGH_MEMORY);
}
