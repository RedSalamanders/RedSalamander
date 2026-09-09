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

#include "DxUi/DxUi.h"
#include "Helpers.h"
#include "PaneVisualState.h"
#include "PathUtils.h"
#include "ProcessCommandLine.h"
#include "StringConversion.h"
#include "UnicodeClipboard.h"
#include "WindowMessages.h"
#include "resource.h"

extern HINSTANCE g_hInstance;

using namespace TerminalPluginDetail;

#include "TerminalVt.h"

GhosttyPoint Terminal::terminalPointFromClient(POINT clientPoint) const noexcept
{
    const float scale    = static_cast<float>(std::max<UINT>(_dpi, 1u)) / 96.0f;
    const int cellWidth  = std::max(1, static_cast<int>(std::lround(_cellWidthDip * scale)));
    const int cellHeight = std::max(1, static_cast<int>(std::lround(_cellHeightDip * scale)));
    const int maximumX   = std::max(0, static_cast<int>(_columns) - 1);
    const int maximumY   = std::max(0, static_cast<int>(_rows) - 1);
    GhosttyPoint point{};
    point.tag                = GHOSTTY_POINT_TAG_VIEWPORT;
    point.value.coordinate.x = static_cast<uint16_t>(std::clamp(static_cast<int>(clientPoint.x) / cellWidth, 0, maximumX));
    point.value.coordinate.y = static_cast<uint32_t>(std::clamp(static_cast<int>(clientPoint.y) / cellHeight, 0, maximumY));
    return point;
}

bool Terminal::updateSelection(POINT clientPoint) noexcept
{
    const GhosttyPoint currentPoint = terminalPointFromClient(clientPoint);
    GhosttyGridRef anchor{};
    anchor.size = sizeof(anchor);
    GhosttyGridRef current{};
    current.size = sizeof(current);
    GhosttySelection selection{};
    selection.size = sizeof(selection);
    {
        std::scoped_lock lock(_terminalMutex);
        if (_ghosttyTerminal == nullptr || _selectionAnchor == nullptr || _runtime.terminalGridRef == nullptr || _runtime.trackedGridRefSnapshot == nullptr ||
            _runtime.trackedGridRefSnapshot(_selectionAnchor, &anchor) != GHOSTTY_SUCCESS ||
            _runtime.terminalGridRef(_ghosttyTerminal, currentPoint, &current) != GHOSTTY_SUCCESS)
        {
            return false;
        }
        selection.start     = anchor;
        selection.end       = current;
        selection.rectangle = (GetKeyState(VK_MENU) & 0x8000) != 0;
        if (_runtime.terminalSet(_ghosttyTerminal, GHOSTTY_TERMINAL_OPT_SELECTION, &selection) != GHOSTTY_SUCCESS)
        {
            return false;
        }
    }
    if (const HWND hwnd = _windowHandle.load(std::memory_order_acquire); hwnd != nullptr)
    {
        InvalidateRect(hwnd, nullptr, FALSE);
    }
    return true;
}

void Terminal::stopSelection() noexcept
{
    _selecting               = false;
    _selectionAutoScrollRows = 0;
    if (const HWND hwnd = _windowHandle.load(std::memory_order_acquire); hwnd != nullptr)
    {
        KillTimer(hwnd, kSelectionAutoScrollTimer);
    }
    std::scoped_lock lock(_terminalMutex);
    if (_selectionAnchor != nullptr && _runtime.trackedGridRefFree != nullptr)
    {
        _runtime.trackedGridRefFree(_selectionAnchor);
        _selectionAnchor = nullptr;
    }
}

bool Terminal::selectWordAt(POINT clientPoint) noexcept
{
    stopSelection();
    GhosttyGridRef cell{};
    cell.size = sizeof(cell);
    GhosttySelection selection{};
    selection.size = sizeof(selection);
    GhosttyTerminalSelectWordOptions options{};
    options.size = sizeof(options);
    {
        std::scoped_lock lock(_terminalMutex);
        if (_ghosttyTerminal == nullptr || _runtime.terminalGridRef == nullptr || _runtime.terminalSelectWord == nullptr ||
            _runtime.terminalGridRef(_ghosttyTerminal, terminalPointFromClient(clientPoint), &cell) != GHOSTTY_SUCCESS)
        {
            return false;
        }
        options.ref = cell;
        if (_runtime.terminalSelectWord(_ghosttyTerminal, &options, &selection) != GHOSTTY_SUCCESS ||
            _runtime.terminalSet(_ghosttyTerminal, GHOSTTY_TERMINAL_OPT_SELECTION, &selection) != GHOSTTY_SUCCESS)
        {
            return false;
        }
    }
    if (const HWND hwnd = _windowHandle.load(std::memory_order_acquire); hwnd != nullptr)
    {
        InvalidateRect(hwnd, nullptr, FALSE);
    }
    return true;
}

bool Terminal::selectAll() noexcept
{
    stopSelection();
    GhosttySelection selection{};
    selection.size = sizeof(selection);
    {
        std::scoped_lock lock(_terminalMutex);
        if (_ghosttyTerminal == nullptr || _runtime.terminalSelectAll == nullptr ||
            _runtime.terminalSelectAll(_ghosttyTerminal, &selection) != GHOSTTY_SUCCESS ||
            _runtime.terminalSet(_ghosttyTerminal, GHOSTTY_TERMINAL_OPT_SELECTION, &selection) != GHOSTTY_SUCCESS)
        {
            return false;
        }
    }
    if (const HWND hwnd = _windowHandle.load(std::memory_order_acquire); hwnd != nullptr)
    {
        InvalidateRect(hwnd, nullptr, FALSE);
    }
    return true;
}

bool Terminal::copySelection(bool* selectionPresent) noexcept
{
    if (selectionPresent != nullptr)
    {
        *selectionPresent = false;
    }
    GhosttyTerminalSelectionFormatOptions options{};
    options.size    = sizeof(options);
    options.emit    = GHOSTTY_FORMATTER_FORMAT_PLAIN;
    options.unwrap  = true;
    options.trim    = true;
    size_t required = 0u;
    {
        std::scoped_lock lock(_terminalMutex);
        if (_ghosttyTerminal == nullptr || _runtime.terminalSelectionFormatBuffer == nullptr)
        {
            return false;
        }
        const GhosttyResult query = _runtime.terminalSelectionFormatBuffer(_ghosttyTerminal, options, nullptr, 0u, &required);
        if (query == GHOSTTY_NO_VALUE || required == 0u)
        {
            return false;
        }
        if (query != GHOSTTY_OUT_OF_SPACE || required > _config.maxFormattedBytes)
        {
            return false;
        }
    }

    auto bytes = std::unique_ptr<uint8_t[]>(new (std::nothrow) uint8_t[required]);
    if (! bytes)
    {
        return false;
    }
    size_t written = 0u;
    {
        std::scoped_lock lock(_terminalMutex);
        if (_ghosttyTerminal == nullptr ||
            _runtime.terminalSelectionFormatBuffer(_ghosttyTerminal, options, bytes.get(), required, &written) != GHOSTTY_SUCCESS || written == 0u ||
            written > required)
        {
            SecureZeroMemory(bytes.get(), required);
            return false;
        }
    }
    const auto wipeBytes                   = wil::scope_exit([&bytes, required]() noexcept { SecureZeroMemory(bytes.get(), required); });
    const std::optional<std::wstring> text = Common::Strings::TryUtf16FromUtf8Strict(std::string_view(reinterpret_cast<const char*>(bytes.get()), written));
    if (! text.has_value() || text.value().empty())
    {
        return false;
    }
#if defined(ENABLE_TESTS)
    if (_debugForceNextClipboardCopyFailure.exchange(false, std::memory_order_acq_rel))
    {
        return false;
    }
#endif
    const HWND hwnd = _windowHandle.load(std::memory_order_acquire);
    const bool copied =
        hwnd != nullptr && Common::Clipboard::TrySetUnicodeText(GetAncestor(hwnd, GA_ROOT), text.value(), Common::Clipboard::EmptyUnicodeTextPolicy::Reject);
    if (copied)
    {
        if (selectionPresent != nullptr)
        {
            *selectionPresent = true;
        }
        clearSelection();
    }
    return copied;
}

bool Terminal::hasSelection() noexcept
{
    GhosttyTerminalSelectionFormatOptions options{};
    options.size    = sizeof(options);
    options.emit    = GHOSTTY_FORMATTER_FORMAT_PLAIN;
    options.unwrap  = true;
    options.trim    = true;
    size_t required = 0u;
    std::scoped_lock lock(_terminalMutex);
    return _ghosttyTerminal != nullptr && _runtime.terminalSelectionFormatBuffer != nullptr &&
           _runtime.terminalSelectionFormatBuffer(_ghosttyTerminal, options, nullptr, 0u, &required) == GHOSTTY_OUT_OF_SPACE && required != 0u;
}

void Terminal::clearSelection() noexcept
{
    {
        std::scoped_lock lock(_terminalMutex);
        if (_ghosttyTerminal == nullptr || _runtime.terminalSet == nullptr ||
            _runtime.terminalSet(_ghosttyTerminal, GHOSTTY_TERMINAL_OPT_SELECTION, nullptr) != GHOSTTY_SUCCESS)
        {
            return;
        }
    }
    if (const HWND hwnd = _windowHandle.load(std::memory_order_acquire); hwnd != nullptr)
    {
        InvalidateRect(hwnd, nullptr, FALSE);
    }
}

bool Terminal::encodeMouse(GhosttyMouseAction action, GhosttyMouseButton button, POINT clientPoint, bool anyButtonPressed) noexcept
{
    if (_mouseEncoder == nullptr || _mouseEvent == nullptr)
    {
        return false;
    }

    GhosttyMouseEncoderSize size{};
    size.size = sizeof(size);
    RECT client{};
    const HWND hwnd = _windowHandle.load(std::memory_order_acquire);
    if (hwnd == nullptr || GetClientRect(hwnd, &client) == FALSE)
    {
        return false;
    }
    size.screen_width  = static_cast<uint32_t>(std::max(1L, client.right - client.left));
    size.screen_height = static_cast<uint32_t>(std::max(1L, client.bottom - client.top));
    const float scale  = static_cast<float>(std::max<UINT>(_dpi, 1u)) / 96.0f;
    size.cell_width    = static_cast<uint32_t>(std::max(1, static_cast<int>(std::lround(_cellWidthDip * scale))));
    size.cell_height   = static_cast<uint32_t>(std::max(1, static_cast<int>(std::lround(_cellHeightDip * scale))));

    GhosttyMods modifiers = 0u;
    if ((GetKeyState(VK_SHIFT) & 0x8000) != 0)
    {
        modifiers |= GHOSTTY_MODS_SHIFT;
    }
    if ((GetKeyState(VK_CONTROL) & 0x8000) != 0)
    {
        modifiers |= GHOSTTY_MODS_CTRL;
    }
    if ((GetKeyState(VK_MENU) & 0x8000) != 0)
    {
        modifiers |= GHOSTTY_MODS_ALT;
    }

    {
        std::scoped_lock lock(_terminalMutex);
        if (_ghosttyTerminal == nullptr)
        {
            return false;
        }
        _runtime.mouseEncoderSetFromTerminal(_mouseEncoder, _ghosttyTerminal);
    }
    _runtime.mouseEncoderSetOption(_mouseEncoder, GHOSTTY_MOUSE_ENCODER_OPT_SIZE, &size);
    _runtime.mouseEncoderSetOption(_mouseEncoder, GHOSTTY_MOUSE_ENCODER_OPT_ANY_BUTTON_PRESSED, &anyButtonPressed);
    _runtime.mouseEventSetAction(_mouseEvent, action);
    if (button == GHOSTTY_MOUSE_BUTTON_UNKNOWN)
    {
        _runtime.mouseEventClearButton(_mouseEvent);
    }
    else
    {
        _runtime.mouseEventSetButton(_mouseEvent, button);
    }
    _runtime.mouseEventSetMods(_mouseEvent, modifiers);
    _runtime.mouseEventSetPosition(_mouseEvent, GhosttyMousePosition{static_cast<float>(clientPoint.x), static_cast<float>(clientPoint.y)});

    std::array<char, 128u> encoded{};
    size_t written       = 0u;
    GhosttyResult result = _runtime.mouseEncoderEncode(_mouseEncoder, _mouseEvent, encoded.data(), encoded.size(), &written);
    if (result == GHOSTTY_SUCCESS)
    {
        return written != 0u && writeInput(std::string_view(encoded.data(), written));
    }
    if (result != GHOSTTY_OUT_OF_SPACE || written == 0u || written > 4096u)
    {
        return false;
    }
    std::vector<char> overflow(written);
    size_t overflowWritten = 0u;
    result                 = _runtime.mouseEncoderEncode(_mouseEncoder, _mouseEvent, overflow.data(), overflow.size(), &overflowWritten);
    return result == GHOSTTY_SUCCESS && overflowWritten != 0u && overflowWritten <= overflow.size() &&
           writeInput(std::string_view(overflow.data(), overflowWritten));
}

void Terminal::NotifyKittyReady(void* context) noexcept
{
    auto* self = static_cast<Terminal*>(context);
    if (self == nullptr || self->_closing.load(std::memory_order_acquire))
    {
        return;
    }
    const HWND hwnd = self->_windowHandle.load(std::memory_order_acquire);
    auto payload    = std::unique_ptr<ReadyWakePayload>(new (std::nothrow) ReadyWakePayload());
    if (hwnd == nullptr || self->_closing.load(std::memory_order_acquire) || ! payload)
    {
        return;
    }
    payload->sessionGeneration = self->_sessionGeneration.load(std::memory_order_acquire);
    static_cast<void>(PostMessagePayload(hwnd, WndMsg::kTerminalKittyReady, 0u, std::move(payload)));
}

void Terminal::releaseReportedMouseButtons() noexcept
{
    for (const GhosttyMouseButton button : kTrackedMouseButtons)
    {
        if (RemoveReportedMouseButton(_reportedMouseButtons, button))
        {
            static_cast<void>(encodeMouse(GHOSTTY_MOUSE_ACTION_RELEASE, button, _reportedMouseLastPoint, _reportedMouseButtons != 0u));
        }
    }
}

void Terminal::scrollViewport(intptr_t rows) noexcept
{
    GhosttyTerminalScrollViewport scroll{};
    scroll.tag         = GHOSTTY_SCROLL_VIEWPORT_DELTA;
    scroll.value.delta = rows;
    {
        std::scoped_lock lock(_terminalMutex);
        if (_ghosttyTerminal == nullptr || _runtime.terminalScrollViewport == nullptr)
        {
            return;
        }
        _runtime.terminalScrollViewport(_ghosttyTerminal, scroll);
#if defined(ENABLE_TESTS)
        ++_debugScrollViewportMutationCount;
#endif
    }
    updateScrollbar();
    if (const HWND hwnd = _windowHandle.load(std::memory_order_acquire); hwnd != nullptr)
    {
        InvalidateRect(hwnd, nullptr, FALSE);
    }
}

void Terminal::scrollViewportTo(uint64_t row) noexcept
{
    GhosttyTerminalScrollViewport scroll{};
    scroll.tag       = GHOSTTY_SCROLL_VIEWPORT_ROW;
    scroll.value.row = static_cast<size_t>(std::min<uint64_t>(row, static_cast<uint64_t>((std::numeric_limits<size_t>::max)())));
    {
        std::scoped_lock lock(_terminalMutex);
        if (_ghosttyTerminal == nullptr || _runtime.terminalScrollViewport == nullptr)
        {
            return;
        }
        _runtime.terminalScrollViewport(_ghosttyTerminal, scroll);
    }
    updateScrollbar();
    if (const HWND hwnd = _windowHandle.load(std::memory_order_acquire); hwnd != nullptr)
    {
        InvalidateRect(hwnd, nullptr, FALSE);
    }
}

void Terminal::setLiveFontSize(float fontSizeDip) noexcept
{
    const float clamped = std::clamp(fontSizeDip, 8.0f, 32.0f);
    if (std::abs(clamped - _liveFontSizeDip) < 0.01f)
    {
        return;
    }
    _liveFontSizeDip = clamped;
    discardDeviceResources();
    resizeTerminal();
    if (const HWND hwnd = _windowHandle.load(std::memory_order_acquire); hwnd != nullptr)
    {
        InvalidateRect(hwnd, nullptr, FALSE);
    }
}

void Terminal::handleVerticalScroll(UINT scrollCode) noexcept
{
    GhosttyTerminalScrollbar scrollbar{};
    if (! getScrollbarState(scrollbar))
    {
        return;
    }

    const uint64_t visible       = std::min(scrollbar.len, scrollbar.total);
    const uint64_t maximumOffset = scrollbar.total - visible;
    uint64_t target              = std::min(scrollbar.offset, maximumOffset);
    switch (scrollCode)
    {
        case SB_LINEUP: target = target > 0u ? target - 1u : 0u; break;
        case SB_LINEDOWN:
            if (target < maximumOffset)
            {
                ++target;
            }
            break;
        case SB_PAGEUP: target = target > visible ? target - visible : 0u; break;
        case SB_PAGEDOWN: target += std::min(visible, maximumOffset - target); break;
        case SB_TOP: target = 0u; break;
        case SB_BOTTOM: target = maximumOffset; break;
        case SB_THUMBPOSITION:
        case SB_THUMBTRACK:
        {
            SCROLLINFO info{};
            info.cbSize = sizeof(info);
            info.fMask  = SIF_TRACKPOS;
            if (GetScrollInfo(_windowHandle.load(std::memory_order_acquire), SB_VERT, &info) != FALSE)
            {
                target = static_cast<uint64_t>(std::max(0, info.nTrackPos));
            }
            break;
        }
        case SB_ENDSCROLL:
        default: return;
    }
    scrollViewportTo(std::min(target, maximumOffset));
}

bool Terminal::getScrollbarState(GhosttyTerminalScrollbar& state) noexcept
{
    state = {};
    std::scoped_lock lock(_terminalMutex);
    return _ghosttyTerminal != nullptr && _runtime.terminalGet != nullptr &&
           _runtime.terminalGet(_ghosttyTerminal, GHOSTTY_TERMINAL_DATA_SCROLLBAR, &state) == GHOSTTY_SUCCESS;
}

void Terminal::updateScrollbar() noexcept
{
    const HWND hwnd = _windowHandle.load(std::memory_order_acquire);
    if (hwnd == nullptr)
    {
        return;
    }

    GhosttyTerminalScrollbar scrollbar{};
    if (! getScrollbarState(scrollbar))
    {
        scrollbar.total = _rows;
        scrollbar.len   = _rows;
    }
    const SCROLLINFO info = MakeTerminalScrollInfo(scrollbar);
    SetScrollInfo(hwnd, SB_VERT, &info, TRUE);
}

void Terminal::updateImeCandidatePosition() noexcept
{
    const HWND hwnd = _windowHandle.load(std::memory_order_acquire);
    if (hwnd == nullptr)
    {
        return;
    }
    uint16_t cursorX      = 0u;
    uint16_t cursorY      = 0u;
    bool cursorInViewport = false;
    if (_renderState != nullptr)
    {
        static_cast<void>(_runtime.renderStateGet(_renderState, GHOSTTY_RENDER_STATE_DATA_CURSOR_VIEWPORT_HAS_VALUE, &cursorInViewport));
        if (cursorInViewport)
        {
            static_cast<void>(_runtime.renderStateGet(_renderState, GHOSTTY_RENDER_STATE_DATA_CURSOR_VIEWPORT_X, &cursorX));
            static_cast<void>(_runtime.renderStateGet(_renderState, GHOSTTY_RENDER_STATE_DATA_CURSOR_VIEWPORT_Y, &cursorY));
        }
    }
    const float scale = static_cast<float>(std::max<UINT>(_dpi, 1u)) / 96.0f;
    const LONG x      = static_cast<LONG>(std::lround(static_cast<float>(cursorX) * _cellWidthDip * scale));
    const LONG y      = static_cast<LONG>(std::lround(static_cast<float>(cursorY + 1u) * _cellHeightDip * scale));
    HIMC inputContext = ImmGetContext(hwnd);
    if (inputContext == nullptr)
    {
        return;
    }
    const auto releaseContext = wil::scope_exit([hwnd, inputContext]() noexcept { static_cast<void>(ImmReleaseContext(hwnd, inputContext)); });
    CANDIDATEFORM candidate{};
    candidate.dwStyle      = CFS_CANDIDATEPOS;
    candidate.ptCurrentPos = POINT{x, y};
    static_cast<void>(ImmSetCandidateWindow(inputContext, &candidate));
    COMPOSITIONFORM composition{};
    composition.dwStyle      = CFS_POINT;
    composition.ptCurrentPos = POINT{x, y};
    static_cast<void>(ImmSetCompositionWindow(inputContext, &composition));
}

bool Terminal::encodeSpecialKey(WPARAM virtualKey, LPARAM keyData, bool released) noexcept
{
    const GhosttyKey key = GhosttyKeyFromVirtualKey(virtualKey);
    if (key == GHOSTTY_KEY_UNIDENTIFIED || _keyEncoder == nullptr || _keyEvent == nullptr)
    {
        return false;
    }
    if (virtualKey == VK_F4 && (GetKeyState(VK_MENU) & 0x8000) != 0)
    {
        return false;
    }

    GhosttyMods modifiers = 0u;
    if ((GetKeyState(VK_SHIFT) & 0x8000) != 0)
    {
        modifiers |= GHOSTTY_MODS_SHIFT;
    }
    if ((GetKeyState(VK_CONTROL) & 0x8000) != 0)
    {
        modifiers |= GHOSTTY_MODS_CTRL;
    }
    if ((GetKeyState(VK_MENU) & 0x8000) != 0)
    {
        modifiers |= GHOSTTY_MODS_ALT;
    }
    if ((GetKeyState(VK_LWIN) & 0x8000) != 0 || (GetKeyState(VK_RWIN) & 0x8000) != 0)
    {
        modifiers |= GHOSTTY_MODS_SUPER;
    }
    if ((GetKeyState(VK_CAPITAL) & 1) != 0)
    {
        modifiers |= GHOSTTY_MODS_CAPS_LOCK;
    }
    if ((GetKeyState(VK_NUMLOCK) & 1) != 0)
    {
        modifiers |= GHOSTTY_MODS_NUM_LOCK;
    }

    {
        std::scoped_lock lock(_terminalMutex);
        if (_ghosttyTerminal == nullptr)
        {
            return false;
        }
        _runtime.keyEncoderSetFromTerminal(_keyEncoder, _ghosttyTerminal);
    }
    const GhosttyKeyAction action =
        released ? GHOSTTY_KEY_ACTION_RELEASE : ((keyData & (static_cast<LPARAM>(1u) << 30u)) != 0 ? GHOSTTY_KEY_ACTION_REPEAT : GHOSTTY_KEY_ACTION_PRESS);
    _runtime.keyEventSetAction(_keyEvent, action);
    _runtime.keyEventSetKey(_keyEvent, key);
    _runtime.keyEventSetMods(_keyEvent, modifiers);
    _runtime.keyEventSetUtf8(_keyEvent, nullptr, 0u);
    _runtime.keyEventSetUnshiftedCodepoint(_keyEvent, 0u);

    const auto writeEncodedInput = [&](std::string_view bytes) noexcept
    {
        if (! writeInput(bytes) || released)
        {
            return;
        }
        markUserInput(bytes);
        // TranslateMessage still delivers WM_CHAR after a Ghostty-encoded keydown for
        // keys that produce a character (Return/Escape). Do not arm for arrows/F-keys.
        if (VirtualKeyProducesTranslatedCharacter(virtualKey))
        {
            MarkTranslatedCharacterConsumed(_translatedCharacterSuppressionPending);
        }
        if (virtualKey == VK_RETURN)
        {
            scrollViewportTo((std::numeric_limits<uint64_t>::max)());
        }
    };

    std::array<char, 128u> encoded{};
    size_t length        = 0u;
    GhosttyResult result = _runtime.keyEncoderEncode(_keyEncoder, _keyEvent, encoded.data(), encoded.size(), &length);
    if (result == GHOSTTY_SUCCESS)
    {
        if (length != 0u)
        {
            writeEncodedInput(std::string_view(encoded.data(), length));
        }
        return true;
    }
    if (result != GHOSTTY_OUT_OF_SPACE || length == 0u || length > 4096u)
    {
        return false;
    }
    std::vector<char> overflow(length);
    size_t written = 0u;
    result         = _runtime.keyEncoderEncode(_keyEncoder, _keyEvent, overflow.data(), overflow.size(), &written);
    if (result != GHOSTTY_SUCCESS || written > overflow.size())
    {
        return false;
    }
    writeEncodedInput(std::string_view(overflow.data(), written));
    return true;
}

void Terminal::resizeTerminal() noexcept
{
    const HWND hwnd = _windowHandle.load(std::memory_order_acquire);
    if (hwnd == nullptr)
    {
        return;
    }
    RECT client{};
    GetClientRect(hwnd, &client);
    const int width        = client.right - client.left;
    const int height       = client.bottom - client.top;
    const bool emptyClient = width <= 0 || height <= 0;
    if (emptyClient)
    {
        // Open and hidden/swapped panes have no usable layout yet. Preserve
        // the current grid until the host supplies a nonzero client size.
        return;
    }
    updateCellMetrics();
    const int cellWidth    = std::max(1, static_cast<int>(std::lround(_cellWidthDip * static_cast<float>(_dpi) / 96.0f)));
    const int cellHeight   = std::max(1, static_cast<int>(std::lround(_cellHeightDip * static_cast<float>(_dpi) / 96.0f)));
    const uint16_t columns = static_cast<uint16_t>(std::clamp(width / cellWidth, 1, static_cast<int>((std::numeric_limits<uint16_t>::max)())));
    const uint16_t rows    = static_cast<uint16_t>(std::clamp(height / cellHeight, 1, static_cast<int>((std::numeric_limits<uint16_t>::max)())));
    if (_ptyClientSizeSynced && columns == _columns && rows == _rows)
    {
        return;
    }
    _columns = columns;
    _rows    = rows;

    {
        std::scoped_lock teardownLock(_sessionTeardownMutex);
        if (_pseudoConsole)
        {
            static_cast<void>(ResizePseudoConsole(
                _pseudoConsole.get(),
                COORD{static_cast<SHORT>(std::min<uint16_t>(_columns, SHRT_MAX)), static_cast<SHORT>(std::min<uint16_t>(_rows, SHRT_MAX))}));
            _ptyClientSizeSynced = true;
        }
    }
    std::scoped_lock lock(_terminalMutex);
    if (_ghosttyTerminal != nullptr)
    {
        static_cast<void>(_runtime.terminalResize(_ghosttyTerminal, _columns, _rows, static_cast<uint32_t>(cellWidth), static_cast<uint32_t>(cellHeight)));
    }
}

void Terminal::updateCellMetrics() noexcept
{
    if (_cellMetricsValid)
    {
        return;
    }

    _cellWidthDip     = std::max(1.0f, _liveFontSizeDip * 0.62f);
    _cellHeightDip    = std::max(1.0f, _liveFontSizeDip * 1.40f);
    _cellMetricsValid = true;
    if (FAILED(ensureTextResources()) || ! _dwriteFactory || ! _textFormat)
    {
        return;
    }

    wil::com_ptr<IDWriteTextLayout> layout;
    constexpr wchar_t kCellMeasureText[] = L"M";
    if (FAILED(_dwriteFactory->CreateTextLayout(kCellMeasureText, 1u, _textFormat.get(), _liveFontSizeDip * 4.0f, _liveFontSizeDip * 4.0f, layout.put())))
    {
        return;
    }

    DWRITE_TEXT_METRICS metrics{};
    if (FAILED(layout->GetMetrics(&metrics)) || metrics.widthIncludingTrailingWhitespace <= 0.0f || metrics.height <= 0.0f)
    {
        return;
    }

    const float scale      = static_cast<float>(std::max<UINT>(_dpi, 1u)) / 96.0f;
    const int cellWidthPx  = std::max(1, static_cast<int>(std::lround(metrics.widthIncludingTrailingWhitespace * scale)));
    const int cellHeightPx = std::max(1, static_cast<int>(std::lround(metrics.height * scale)));
    _cellWidthDip          = static_cast<float>(cellWidthPx) / scale;
    _cellHeightDip         = static_cast<float>(cellHeightPx) / scale;
}

HRESULT Terminal::ensureTextResources() noexcept
{
    if (! _dwriteFactory)
    {
        const HRESULT hr = DWriteCreateFactory(DWRITE_FACTORY_TYPE_SHARED, __uuidof(IDWriteFactory), reinterpret_cast<IUnknown**>(_dwriteFactory.put()));
        if (FAILED(hr))
        {
            return hr;
        }
    }
    const auto createTextFormat = [this](DWRITE_FONT_WEIGHT weight, DWRITE_FONT_STYLE style, wil::com_ptr<IDWriteTextFormat>& output) noexcept -> HRESULT
    {
        if (output)
        {
            return S_OK;
        }
        const HRESULT hr = _dwriteFactory->CreateTextFormat(
            _config.fontFamily.c_str(), nullptr, weight, style, DWRITE_FONT_STRETCH_NORMAL, _liveFontSizeDip, L"en-us", output.put());
        if (FAILED(hr))
        {
            return hr;
        }
        static_cast<void>(output->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP));
        static_cast<void>(output->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_LEADING));
        static_cast<void>(output->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_NEAR));
        return S_OK;
    };

    HRESULT hr = createTextFormat(DWRITE_FONT_WEIGHT_NORMAL, DWRITE_FONT_STYLE_NORMAL, _textFormat);
    if (SUCCEEDED(hr))
    {
        hr = createTextFormat(DWRITE_FONT_WEIGHT_BOLD, DWRITE_FONT_STYLE_NORMAL, _boldTextFormat);
    }
    if (SUCCEEDED(hr))
    {
        hr = createTextFormat(DWRITE_FONT_WEIGHT_NORMAL, DWRITE_FONT_STYLE_ITALIC, _italicTextFormat);
    }
    if (SUCCEEDED(hr))
    {
        hr = createTextFormat(DWRITE_FONT_WEIGHT_BOLD, DWRITE_FONT_STYLE_ITALIC, _boldItalicTextFormat);
    }
    if (SUCCEEDED(hr))
    {
        hr = createTextFormat(DWRITE_FONT_WEIGHT_NORMAL, DWRITE_FONT_STYLE_NORMAL, _overlayTextFormat);
    }
    if (SUCCEEDED(hr))
    {
        hr = createTextFormat(DWRITE_FONT_WEIGHT_BOLD, DWRITE_FONT_STYLE_NORMAL, _overlayBoldTextFormat);
    }
    if (FAILED(hr))
    {
        return hr;
    }
    static_cast<void>(_overlayTextFormat->SetWordWrapping(DWRITE_WORD_WRAPPING_WRAP));
    static_cast<void>(_overlayBoldTextFormat->SetWordWrapping(DWRITE_WORD_WRAPPING_WRAP));
    return S_OK;
}

HRESULT Terminal::ensureDeviceResources() noexcept
{
    RETURN_IF_FAILED(ensureTextResources());
    if (! _d2dFactory)
    {
        const HRESULT hr = D2D1CreateFactory(D2D1_FACTORY_TYPE_SINGLE_THREADED, _d2dFactory.put());
        if (FAILED(hr))
        {
            return hr;
        }
    }
    if (! _renderTarget)
    {
        const HWND hwnd = _windowHandle.load(std::memory_order_acquire);
        if (hwnd == nullptr)
        {
            return E_HANDLE;
        }
        RECT client{};
        GetClientRect(hwnd, &client);
        const D2D1_SIZE_U size =
            D2D1::SizeU(static_cast<UINT32>(std::max(0L, client.right - client.left)), static_cast<UINT32>(std::max(0L, client.bottom - client.top)));
        HRESULT hr = _d2dFactory->CreateHwndRenderTarget(D2D1::RenderTargetProperties(), D2D1::HwndRenderTargetProperties(hwnd, size), _renderTarget.put());
        if (FAILED(hr))
        {
            return hr;
        }
        ++_kittyDeviceGeneration;
        if (_kittyDeviceGeneration == 0u)
        {
            ++_kittyDeviceGeneration;
        }
        _renderTarget->SetDpi(static_cast<float>(_dpi), static_cast<float>(_dpi));
        hr = _renderTarget->CreateSolidColorBrush(RedSalamander::DxUi::ColorFromArgb(_foregroundArgb), _foregroundBrush.put());
        if (FAILED(hr))
        {
            discardDeviceResources();
            return hr;
        }
    }
    return S_OK;
}

HRESULT Terminal::resizeRenderTargetToClient() noexcept
{
    if (! _renderTarget)
    {
        return S_FALSE;
    }
    const HWND hwnd = _windowHandle.load(std::memory_order_acquire);
    if (hwnd == nullptr)
    {
        return E_HANDLE;
    }
    RECT client{};
    if (GetClientRect(hwnd, &client) == FALSE)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }
    const D2D1_SIZE_U size =
        D2D1::SizeU(static_cast<UINT32>(std::max(0L, client.right - client.left)), static_cast<UINT32>(std::max(0L, client.bottom - client.top)));
    if (size.width == 0u || size.height == 0u)
    {
        return S_FALSE;
    }
    const D2D1_SIZE_U currentSize = _renderTarget->GetPixelSize();
    if (currentSize.width == size.width && currentSize.height == size.height)
    {
        return S_FALSE;
    }

    const auto startedAt = std::chrono::steady_clock::now();
    const HRESULT hr     = _renderTarget->Resize(size);
    if (SUCCEEDED(hr))
    {
        ++_kittyResizeCount;
        _kittyResizeRetainedBitmapCount = 0u;
        _kittyResizeRetainedBitmapBytes = 0u;
        for (const auto& [imageKey, image] : _kittyImages)
        {
            static_cast<void>(imageKey);
            if (image.bitmap)
            {
                ++_kittyResizeRetainedBitmapCount;
                _kittyResizeRetainedBitmapBytes += image.bgra.size();
            }
        }
    }
    else if (hr == D2DERR_RECREATE_TARGET)
    {
        ++_kittyResizeRecreateCount;
    }
    Debug::Perf::EmitDurationUs(L"terminal.kitty.resize_us",
                                std::max<uint64_t>(1u, Debug::Perf::ElapsedUs(startedAt)),
                                _kittyResizeRetainedBitmapBytes,
                                _kittyResizeRetainedBitmapCount,
                                hr);
    return hr;
}

void Terminal::handleResize(HWND hwnd) noexcept
{
    RECT client{};
    GetClientRect(hwnd, &client);
    if ((client.right - client.left) <= 0 || (client.bottom - client.top) <= 0)
    {
        return;
    }
    if (resizeRenderTargetToClient() == D2DERR_RECREATE_TARGET)
    {
        _kittyRecreateTargetAfterFrame = true;
        static_cast<void>(recoverKittyDeviceAfterFrame(hwnd));
    }
    resizeTerminal();
    updateScrollbar();
    InvalidateRect(hwnd, nullptr, FALSE);
}

void Terminal::discardDeviceResources() noexcept
{
    if (const HWND hwnd = _windowHandle.load(std::memory_order_acquire); hwnd != nullptr)
    {
        KillTimer(hwnd, kKittyUploadRetryTimer);
    }
    _kittyNextUploadRetryAtTick = 0u;
    for (auto& [imageKey, image] : _kittyImages)
    {
        static_cast<void>(imageKey);
        image.bitmap.reset();
    }
    _foregroundBrush.reset();
    _renderTarget.reset();
    _overlayBoldTextFormat.reset();
    _overlayTextFormat.reset();
    _boldItalicTextFormat.reset();
    _italicTextFormat.reset();
    _boldTextFormat.reset();
    _textFormat.reset();
    _cellMetricsValid = false;
}

std::wstring Terminal::formatScreen() noexcept
{
    std::scoped_lock lock(_terminalMutex);
    if (! _diagnosticText.empty())
    {
        return _diagnosticText;
    }
    if (_formatter == nullptr)
    {
        return {};
    }
    size_t required           = 0u;
    const GhosttyResult query = _runtime.formatterFormatBuffer(_formatter, nullptr, 0u, &required);
    if (required == 0u)
    {
        return {};
    }
    const uint32_t limit = std::clamp(_config.maxFormattedBytes, kMinimumFormattedBytes, kMaximumFormattedBytes);
    if (query != GHOSTTY_OUT_OF_SPACE || required > limit)
    {
        return LoadStringResource(g_hInstance, IDS_TERMINAL_FORMAT_LIMIT);
    }
    std::vector<uint8_t> bytes(required);
    size_t written = 0u;
    if (_runtime.formatterFormatBuffer(_formatter, bytes.data(), bytes.size(), &written) != GHOSTTY_SUCCESS || written > bytes.size())
    {
        return LoadStringResource(g_hInstance, IDS_TERMINAL_FORMAT_ERROR);
    }
    return Common::Strings::Utf16FromUtf8ReplacingInvalid(std::string_view(reinterpret_cast<const char*>(bytes.data()), written));
}

bool Terminal::handleKittyImageQueryResult(KittyImageSnapshot& image, GhosttyResult result) noexcept
{
    if (result == GHOSTTY_SUCCESS)
    {
        return false;
    }
    image.state            = result == GHOSTTY_NO_VALUE ? KittyImageState::EnginePending : KittyImageState::Rejected;
    image.pendingRequestId = 0u;
    return true;
}

void Terminal::resetKittyPendingForStorageChange(KittyImageSnapshot& image) noexcept
{
    if (image.state == KittyImageState::ConversionPending || image.state == KittyImageState::EnginePending || image.state == KittyImageState::Rejected)
    {
        image.state            = KittyImageState::Missing;
        image.pendingRequestId = 0u;
    }
}

bool Terminal::needsKittyReplacementWork(const KittyImageSnapshot& image, uint64_t currentRequestId, bool retryDeferred) noexcept
{
    return image.state == KittyImageState::Missing || image.state == KittyImageState::EnginePending || retryDeferred ||
           (image.state == KittyImageState::ConversionPending && image.pendingRequestId != currentRequestId);
}

bool Terminal::isCurrentKittyConversionResult(const KittyImageSnapshot& image, uint64_t requestId) noexcept
{
    return image.state == KittyImageState::ConversionPending && image.pendingRequestId == requestId;
}

bool Terminal::captureKittyGraphicsLocked(TerminalKittyGenerationWork& work, std::vector<KittyPlacementSnapshot>& placements, bool& submitWork) noexcept
{
    struct VisibleImage final
    {
        KittyImageKey key{};
        GhosttyKittyGraphicsImage image = nullptr;
    };

    work = {};
    placements.clear();
    submitWork                          = false;
    _kittyCaptureSourceBytes            = 0u;
    _kittyCapturePinnedBytes            = 0u;
    _kittyCaptureEnginePlacementBytes   = 0u;
    _kittyCaptureEnginePlacementCount   = 0u;
    _kittyCaptureTotalPlacements        = 0u;
    _kittyCaptureVisiblePlacements      = 0u;
    _kittyCaptureOffscreenPlacements    = 0u;
    _kittyCaptureVirtualPlacements      = 0u;
    _kittyCaptureVisibleImageKeys       = 0u;
    _kittyCaptureMissingImageKeys       = 0u;
    _kittyCaptureEnginePendingImageKeys = 0u;
    _kittyCaptureDeferredImageKeys      = 0u;
    _kittyCaptureRejectedImageKeys      = 0u;
    if (_ghosttyTerminal == nullptr || _kittyPlacementIterator == nullptr)
    {
        return true;
    }

    GhosttyKittyGraphics graphics = nullptr;
    uint64_t storageGeneration    = 0u;
    if (_runtime.terminalGet(_ghosttyTerminal, GHOSTTY_TERMINAL_DATA_KITTY_GRAPHICS, &graphics) != GHOSTTY_SUCCESS || graphics == nullptr ||
        _runtime.kittyGraphicsGet(graphics, GHOSTTY_KITTY_GRAPHICS_DATA_GENERATION, &storageGeneration) != GHOSTTY_SUCCESS)
    {
        return false;
    }
    work.storageGeneration = storageGeneration;
    if (storageGeneration == 0u)
    {
        return true;
    }
    size_t placementCountLimit = 0u;
    size_t placementBytesLimit = 0u;
    if (_runtime.kittyGraphicsGet(graphics, GHOSTTY_KITTY_GRAPHICS_DATA_PLACEMENT_COUNT, &_kittyCaptureEnginePlacementCount) != GHOSTTY_SUCCESS ||
        _runtime.kittyGraphicsGet(graphics, GHOSTTY_KITTY_GRAPHICS_DATA_PLACEMENT_ALLOCATED_BYTES, &_kittyCaptureEnginePlacementBytes) != GHOSTTY_SUCCESS ||
        _runtime.kittyGraphicsGet(graphics, GHOSTTY_KITTY_GRAPHICS_DATA_PLACEMENT_COUNT_LIMIT, &placementCountLimit) != GHOSTTY_SUCCESS ||
        _runtime.kittyGraphicsGet(graphics, GHOSTTY_KITTY_GRAPHICS_DATA_PLACEMENT_BYTES_LIMIT, &placementBytesLimit) != GHOSTTY_SUCCESS ||
        _kittyCaptureEnginePlacementCount > kMaximumKittyPlacements || _kittyCaptureEnginePlacementBytes > kMaximumKittyPlacementBytes ||
        placementCountLimit != kMaximumKittyPlacements || placementBytesLimit != kMaximumKittyPlacementBytes)
    {
        return false;
    }
    if (_runtime.kittyGraphicsGet(graphics, GHOSTTY_KITTY_GRAPHICS_DATA_PLACEMENT_ITERATOR, &_kittyPlacementIterator) != GHOSTTY_SUCCESS)
    {
        return false;
    }

    const bool storageGenerationChanged = storageGeneration != _kittyTargetStorageGeneration;
    if (storageGenerationChanged)
    {
        _kittyPipeline.Invalidate(++_kittyRequestId);
        _kittyTargetStorageGeneration = storageGeneration;
        for (auto& [imageKey, image] : _kittyImages)
        {
            static_cast<void>(imageKey);
            resetKittyPendingForStorageChange(image);
        }
    }

    ++_kittyCacheEpoch;
    if (_kittyCacheEpoch == 0u)
    {
        _kittyCacheEpoch = 1u;
        for (auto& [imageKey, image] : _kittyImages)
        {
            static_cast<void>(imageKey);
            image.lastVisibleEpoch = 0u;
        }
    }

    std::vector<VisibleImage> visibleImages;
    visibleImages.reserve(std::min(_kittyCaptureEnginePlacementCount, kMaximumKittyPlacements));
    placements.reserve(std::min(_kittyCaptureEnginePlacementCount, kMaximumKittyPlacements));
    while (_runtime.kittyPlacementNext(_kittyPlacementIterator))
    {
        ++_kittyCaptureTotalPlacements;
        if (_kittyCaptureTotalPlacements > kMaximumKittyPlacements)
        {
            return false;
        }
        uint32_t imageId = 0u;
        uint32_t xOffset = 0u;
        uint32_t yOffset = 0u;
        int32_t z        = 0;
        bool isVirtual   = false;
        if (_runtime.kittyPlacementGet(_kittyPlacementIterator, GHOSTTY_KITTY_GRAPHICS_PLACEMENT_DATA_IMAGE_ID, &imageId) != GHOSTTY_SUCCESS ||
            _runtime.kittyPlacementGet(_kittyPlacementIterator, GHOSTTY_KITTY_GRAPHICS_PLACEMENT_DATA_IS_VIRTUAL, &isVirtual) != GHOSTTY_SUCCESS ||
            _runtime.kittyPlacementGet(_kittyPlacementIterator, GHOSTTY_KITTY_GRAPHICS_PLACEMENT_DATA_X_OFFSET, &xOffset) != GHOSTTY_SUCCESS ||
            _runtime.kittyPlacementGet(_kittyPlacementIterator, GHOSTTY_KITTY_GRAPHICS_PLACEMENT_DATA_Y_OFFSET, &yOffset) != GHOSTTY_SUCCESS ||
            _runtime.kittyPlacementGet(_kittyPlacementIterator, GHOSTTY_KITTY_GRAPHICS_PLACEMENT_DATA_Z, &z) != GHOSTTY_SUCCESS || isVirtual)
        {
            if (isVirtual)
            {
                ++_kittyCaptureVirtualPlacements;
            }
            continue;
        }

        const GhosttyKittyGraphicsImage image = _runtime.kittyGraphicsImage(graphics, imageId);
        if (image == nullptr)
        {
            continue;
        }
        GhosttyKittyGraphicsPlacementRenderInfo renderInfo{};
        renderInfo.size = sizeof(renderInfo);
        if (_runtime.kittyPlacementRenderInfo(_kittyPlacementIterator, image, _ghosttyTerminal, &renderInfo) != GHOSTTY_SUCCESS)
        {
            continue;
        }
        if (! renderInfo.viewport_visible)
        {
            ++_kittyCaptureOffscreenPlacements;
            continue;
        }
        if (renderInfo.pixel_width == 0u || renderInfo.pixel_height == 0u || renderInfo.source_width == 0u || renderInfo.source_height == 0u)
        {
            continue;
        }

        uint64_t imageGeneration = 0u;
        const std::array fields{GHOSTTY_KITTY_IMAGE_DATA_GENERATION};
        std::array<void*, 1u> imageGenerationValues{{&imageGeneration}};
        if (_runtime.kittyGraphicsImageGetMulti(image, fields.size(), fields.data(), imageGenerationValues.data(), nullptr) != GHOSTTY_SUCCESS ||
            imageGeneration == 0u)
        {
            continue;
        }

        const KittyImageKey imageKey{imageId, imageGeneration};
        for (auto imageIterator = _kittyImages.begin(); imageIterator != _kittyImages.end();)
        {
            if (imageIterator->first.imageId != imageId || imageIterator->first.imageGeneration == imageGeneration)
            {
                ++imageIterator;
                continue;
            }
            _kittyCachedBytes -= std::min(_kittyCachedBytes, imageIterator->second.bgra.size());
            imageIterator = _kittyImages.erase(imageIterator);
            ++_kittySupersededImageCount;
        }
        auto [snapshotIterator, inserted] = _kittyImages.try_emplace(imageKey);
        static_cast<void>(inserted);
        KittyImageSnapshot& snapshot = snapshotIterator->second;
        if (snapshot.lastVisibleEpoch != _kittyCacheEpoch)
        {
            snapshot.lastVisibleEpoch = _kittyCacheEpoch;
            visibleImages.push_back(VisibleImage{imageKey, image});
        }

        KittyPlacementSnapshot placement{};
        placement.imageKey      = imageKey;
        placement.renderInfo    = renderInfo;
        placement.xOffsetPixels = xOffset;
        placement.yOffsetPixels = yOffset;
        placement.z             = z;
        placements.push_back(placement);
    }

    _kittyCaptureVisiblePlacements = placements.size();
    _kittyCaptureVisibleImageKeys  = visibleImages.size();
    if (storageGenerationChanged)
    {
        for (auto imageIterator = _kittyImages.begin(); imageIterator != _kittyImages.end();)
        {
            if (imageIterator->second.lastVisibleEpoch == _kittyCacheEpoch)
            {
                ++imageIterator;
                continue;
            }
            _kittyCachedBytes -= std::min(_kittyCachedBytes, imageIterator->second.bgra.size());
            imageIterator = _kittyImages.erase(imageIterator);
            ++_kittyRemovedImageCount;
        }
    }

    for (const VisibleImage& visible : visibleImages)
    {
        const KittyImageSnapshot& snapshot = _kittyImages.find(visible.key)->second;
        if (snapshot.state == KittyImageState::Ready)
        {
            _kittyCapturePinnedBytes += snapshot.bgra.size();
        }
    }

    const auto retryDeferred = [this, visibleKeyCount = visibleImages.size()](const KittyImageSnapshot& image) noexcept
    {
        return image.state == KittyImageState::Deferred &&
               (_kittyCapturePinnedBytes < image.deferredPinnedBytes || visibleKeyCount < image.deferredVisibleKeyCount);
    };
    const bool needsReplacementWork = std::ranges::any_of(visibleImages,
                                                          [this, &retryDeferred](const VisibleImage& visible) noexcept
    {
        const KittyImageSnapshot& image = _kittyImages.find(visible.key)->second;
        return needsKittyReplacementWork(image, _kittyRequestId, retryDeferred(image));
    });

    size_t totalSourceBytes    = 0u;
    size_t totalConvertedBytes = 0u;
    if (needsReplacementWork)
    {
        const size_t availableConvertedBytes =
            TerminalKittyImagePipeline::MaximumConvertedBytes - std::min(_kittyCapturePinnedBytes, TerminalKittyImagePipeline::MaximumConvertedBytes);
        for (const VisibleImage& visible : visibleImages)
        {
            KittyImageSnapshot& snapshot = _kittyImages.find(visible.key)->second;
            if (snapshot.state == KittyImageState::Ready || snapshot.state == KittyImageState::Rejected ||
                (snapshot.state == KittyImageState::Deferred && ! retryDeferred(snapshot)))
            {
                continue;
            }

            uint32_t width                           = 0u;
            uint32_t height                          = 0u;
            GhosttyKittyImageFormat format           = GHOSTTY_KITTY_IMAGE_FORMAT_MAX_VALUE;
            GhosttyKittyImageCompression compression = GHOSTTY_KITTY_IMAGE_COMPRESSION_MAX_VALUE;
            const uint8_t* source                    = nullptr;
            size_t sourceLength                      = 0u;
            const std::array<GhosttyKittyGraphicsImageData, 6u> keys{GHOSTTY_KITTY_IMAGE_DATA_WIDTH,
                                                                     GHOSTTY_KITTY_IMAGE_DATA_HEIGHT,
                                                                     GHOSTTY_KITTY_IMAGE_DATA_FORMAT,
                                                                     GHOSTTY_KITTY_IMAGE_DATA_COMPRESSION,
                                                                     GHOSTTY_KITTY_IMAGE_DATA_DATA_PTR,
                                                                     GHOSTTY_KITTY_IMAGE_DATA_DATA_LEN};
            std::array<void*, 6u> values{&width, &height, &format, &compression, &source, &sourceLength};
            const GhosttyResult imageResult = _runtime.kittyGraphicsImageGetMulti(visible.image, keys.size(), keys.data(), values.data(), nullptr);
            if (handleKittyImageQueryResult(snapshot, imageResult))
            {
                continue;
            }
            if (source == nullptr || width == 0u || height == 0u || compression != GHOSTTY_KITTY_IMAGE_COMPRESSION_NONE)
            {
                snapshot.state            = KittyImageState::Rejected;
                snapshot.pendingRequestId = 0u;
                continue;
            }

            size_t sourceBytesPerPixel           = 0u;
            TerminalKittyPixelFormat pixelFormat = TerminalKittyPixelFormat::Rgba;
            switch (format)
            {
                case GHOSTTY_KITTY_IMAGE_FORMAT_RGB:
                    sourceBytesPerPixel = 3u;
                    pixelFormat         = TerminalKittyPixelFormat::Rgb;
                    break;
                case GHOSTTY_KITTY_IMAGE_FORMAT_RGBA:
                    sourceBytesPerPixel = 4u;
                    pixelFormat         = TerminalKittyPixelFormat::Rgba;
                    break;
                case GHOSTTY_KITTY_IMAGE_FORMAT_GRAY_ALPHA:
                    sourceBytesPerPixel = 2u;
                    pixelFormat         = TerminalKittyPixelFormat::GrayAlpha;
                    break;
                case GHOSTTY_KITTY_IMAGE_FORMAT_GRAY:
                    sourceBytesPerPixel = 1u;
                    pixelFormat         = TerminalKittyPixelFormat::Gray;
                    break;
                case GHOSTTY_KITTY_IMAGE_FORMAT_PNG:
                case GHOSTTY_KITTY_IMAGE_FORMAT_MAX_VALUE:
                    snapshot.state            = KittyImageState::Rejected;
                    snapshot.pendingRequestId = 0u;
                    continue;
            }

            const uint64_t pixelCount = static_cast<uint64_t>(width) * static_cast<uint64_t>(height);
            if (pixelCount > kMaximumKittyPixels)
            {
                snapshot.state            = KittyImageState::Rejected;
                snapshot.pendingRequestId = 0u;
                continue;
            }
            const uint64_t expectedSourceBytes = pixelCount * sourceBytesPerPixel;
            const uint64_t convertedBytes      = pixelCount * 4u;
            if (expectedSourceBytes != sourceLength || expectedSourceBytes > TerminalKittyImagePipeline::MaximumSourceBytes ||
                convertedBytes > TerminalKittyImagePipeline::MaximumConvertedBytes)
            {
                snapshot.state            = KittyImageState::Rejected;
                snapshot.pendingRequestId = 0u;
                continue;
            }
            if (totalSourceBytes > TerminalKittyImagePipeline::MaximumSourceBytes - static_cast<size_t>(expectedSourceBytes) ||
                convertedBytes > availableConvertedBytes || totalConvertedBytes > availableConvertedBytes - static_cast<size_t>(convertedBytes))
            {
                snapshot.state                   = KittyImageState::Deferred;
                snapshot.pendingRequestId        = 0u;
                snapshot.deferredPinnedBytes     = _kittyCapturePinnedBytes;
                snapshot.deferredVisibleKeyCount = visibleImages.size();
                continue;
            }

            TerminalKittyImageWork imageWork{};
            imageWork.imageId         = visible.key.imageId;
            imageWork.width           = width;
            imageWork.height          = height;
            imageWork.imageGeneration = visible.key.imageGeneration;
            imageWork.format          = pixelFormat;
            imageWork.source.assign(source, source + sourceLength);
            totalSourceBytes += sourceLength;
            totalConvertedBytes += static_cast<size_t>(convertedBytes);
            work.images.push_back(std::move(imageWork));
        }
    }

    if (! work.images.empty())
    {
        work.requestId = ++_kittyRequestId;
        submitWork     = true;
        for (const TerminalKittyImageWork& imageWork : work.images)
        {
            const KittyImageKey key{imageWork.imageId, imageWork.imageGeneration};
            KittyImageSnapshot& snapshot = _kittyImages.find(key)->second;
            snapshot.state               = KittyImageState::ConversionPending;
            snapshot.pendingRequestId    = work.requestId;
        }
    }
    _kittyCaptureSourceBytes = totalSourceBytes;
    for (const VisibleImage& visible : visibleImages)
    {
        const KittyImageState state = _kittyImages.find(visible.key)->second.state;
        if (state == KittyImageState::Missing || state == KittyImageState::ConversionPending)
        {
            ++_kittyCaptureMissingImageKeys;
        }
        else if (state == KittyImageState::EnginePending)
        {
            ++_kittyCaptureEnginePendingImageKeys;
        }
        else if (state == KittyImageState::Deferred)
        {
            ++_kittyCaptureDeferredImageKeys;
        }
        else if (state == KittyImageState::Rejected)
        {
            ++_kittyCaptureRejectedImageKeys;
        }
    }
    return true;
}

void Terminal::consumeKittyReady() noexcept
{
    std::optional<TerminalKittyGenerationResult> result = _kittyPipeline.TakeReady(_kittyRequestId);
    if (! result.has_value() || result->storageGeneration != _kittyTargetStorageGeneration)
    {
        return;
    }
    Debug::Perf::EmitValue(L"terminal.kitty.converted_bytes", result->convertedBytes);

    std::vector<KittyImageKey> evictionCandidates;
    evictionCandidates.reserve(_kittyImages.size());
    for (const auto& [imageKey, image] : _kittyImages)
    {
        if (image.state == KittyImageState::Ready && image.lastVisibleEpoch != _kittyCacheEpoch)
        {
            evictionCandidates.push_back(imageKey);
        }
    }
    std::ranges::sort(evictionCandidates, [this](const KittyImageKey& left, const KittyImageKey& right) noexcept {
        return _kittyImages.find(left)->second.lastVisibleEpoch < _kittyImages.find(right)->second.lastVisibleEpoch;
    });
    size_t evictionIndex = 0u;
    for (TerminalKittyConvertedImage& converted : result->images)
    {
        const KittyImageKey imageKey{converted.imageId, converted.imageGeneration};
        const auto imageIterator = _kittyImages.find(imageKey);
        if (imageIterator == _kittyImages.end() || ! isCurrentKittyConversionResult(imageIterator->second, result->requestId))
        {
            continue;
        }
        const size_t convertedBytes = converted.bgra.size();
        while (convertedBytes >
                   TerminalKittyImagePipeline::MaximumConvertedBytes - std::min(_kittyCachedBytes, TerminalKittyImagePipeline::MaximumConvertedBytes) &&
               evictionIndex < evictionCandidates.size())
        {
            const KittyImageKey evictedKey = evictionCandidates[evictionIndex++];
            const auto evictedIterator     = _kittyImages.find(evictedKey);
            if (evictedIterator == _kittyImages.end())
            {
                continue;
            }
            const size_t evictedBytes = evictedIterator->second.bgra.size();
            _kittyCachedBytes -= std::min(_kittyCachedBytes, evictedBytes);
            _kittyEvictedBytes += evictedBytes;
            ++_kittyEvictedImageCount;
            _kittyImages.erase(evictedIterator);
        }

        KittyImageSnapshot& image = imageIterator->second;
        if (convertedBytes > TerminalKittyImagePipeline::MaximumConvertedBytes - std::min(_kittyCachedBytes, TerminalKittyImagePipeline::MaximumConvertedBytes))
        {
            image.state                   = KittyImageState::Deferred;
            image.pendingRequestId        = 0u;
            image.deferredPinnedBytes     = _kittyCapturePinnedBytes;
            image.deferredVisibleKeyCount = _kittyCaptureVisibleImageKeys;
            continue;
        }
        image.width  = converted.width;
        image.height = converted.height;
        image.bgra   = std::move(converted.bgra);
        image.bitmap.reset();
        image.uploadState            = KittyBitmapUploadState::Eligible;
        image.uploadHr               = S_OK;
        image.uploadRetryAtTick      = 0u;
        image.uploadDeviceGeneration = _kittyDeviceGeneration;
        image.uploadAttemptCount     = 0u;
        image.state                  = KittyImageState::Ready;
        image.pendingRequestId       = 0u;
        _kittyCachedBytes += image.bgra.size();
    }
}

void Terminal::requestKittyRepaint(HWND hwnd) noexcept
{
    if (hwnd != nullptr && InvalidateRect(hwnd, nullptr, FALSE) != FALSE)
    {
        ++_kittyRepaintRequestCount;
    }
}

void Terminal::armKittyUploadRetryTimer(ULONGLONG retryAtTick) noexcept
{
    const HWND hwnd = _windowHandle.load(std::memory_order_acquire);
    if (hwnd == nullptr || (_kittyNextUploadRetryAtTick != 0u && _kittyNextUploadRetryAtTick <= retryAtTick))
    {
        return;
    }
    const ULONGLONG now           = GetTickCount64();
    const ULONGLONG remaining     = retryAtTick > now ? retryAtTick - now : 1u;
    const DWORD delayMilliseconds = static_cast<DWORD>(std::min<ULONGLONG>((std::numeric_limits<DWORD>::max)(), std::max<ULONGLONG>(1u, remaining)));
    _kittyNextUploadRetryAtTick   = retryAtTick;
    if (SetTimer(hwnd, kKittyUploadRetryTimer, delayMilliseconds, nullptr) == 0u)
    {
        _kittyNextUploadRetryAtTick = 0u;
        requestKittyRepaint(hwnd);
    }
}

void Terminal::scheduleKittyUploadRetry(KittyImageSnapshot& image, DWORD delayMilliseconds) noexcept
{
    image.uploadRetryAtTick = GetTickCount64() + delayMilliseconds;
    ++_kittyUploadRetryScheduleCount;
    armKittyUploadRetryTimer(image.uploadRetryAtTick);
}

Terminal::KittyBitmapUploadResult Terminal::ensureKittyBitmap(KittyImageSnapshot& image) noexcept
{
    if (image.bitmap)
    {
        return {KittyBitmapUploadDisposition::Ready, S_OK};
    }
    if (! _renderTarget || image.width == 0u || image.height == 0u || image.width > (std::numeric_limits<UINT32>::max)() / 4u ||
        image.bgra.size() != static_cast<size_t>(image.width) * image.height * 4u)
    {
        return {KittyBitmapUploadDisposition::InvalidInput, E_INVALIDARG};
    }
    if (image.uploadDeviceGeneration != _kittyDeviceGeneration)
    {
        image.uploadState            = KittyBitmapUploadState::Eligible;
        image.uploadHr               = S_OK;
        image.uploadRetryAtTick      = 0u;
        image.uploadDeviceGeneration = _kittyDeviceGeneration;
        image.uploadAttemptCount     = 0u;
    }
    if (image.uploadState == KittyBitmapUploadState::PermanentFailure)
    {
        return {KittyBitmapUploadDisposition::PermanentFailure, image.uploadHr};
    }
    if (image.uploadState == KittyBitmapUploadState::RecreatePending)
    {
        return {KittyBitmapUploadDisposition::RecreateTarget, image.uploadHr};
    }
    if (image.uploadState == KittyBitmapUploadState::RetryPending)
    {
        if (GetTickCount64() < image.uploadRetryAtTick)
        {
            armKittyUploadRetryTimer(image.uploadRetryAtTick);
            return {KittyBitmapUploadDisposition::DeferredBackoff, image.uploadHr};
        }
        image.uploadState = KittyBitmapUploadState::Eligible;
    }
    const size_t uploadBytes = image.bgra.size();
    if (_kittyFrameUploadCount >= kMaximumKittyUploadsPerFrame || uploadBytes > kMaximumKittyUploadBytesPerFrame - _kittyFrameUploadBytes)
    {
        _kittyUploadDeferred = true;
        return {KittyBitmapUploadDisposition::DeferredBudget, S_FALSE};
    }
    ++_kittyFrameUploadCount;
    _kittyFrameUploadBytes += uploadBytes;
    ++image.uploadAttemptCount;
    const D2D1_BITMAP_PROPERTIES properties =
        D2D1::BitmapProperties(D2D1::PixelFormat(DXGI_FORMAT_B8G8R8A8_UNORM, D2D1_ALPHA_MODE_PREMULTIPLIED), 96.0f, 96.0f);
    const auto startedAt = std::chrono::steady_clock::now();
    HRESULT hr           = S_OK;
#if defined(ENABLE_TESTS)
    if (! _debugKittyUploadFailures.empty())
    {
        hr = _debugKittyUploadFailures.front();
        _debugKittyUploadFailures.pop_front();
    }
    else
#endif
    {
        hr = _renderTarget->CreateBitmap(D2D1::SizeU(image.width, image.height), image.bgra.data(), image.width * 4u, properties, image.bitmap.put());
    }
    Debug::Perf::EmitDurationUs(L"terminal.kitty.upload_us", Debug::Perf::ElapsedUs(startedAt), uploadBytes, 1u, hr);
    image.uploadHr = hr;
    if (SUCCEEDED(hr))
    {
        image.uploadState       = KittyBitmapUploadState::Eligible;
        image.uploadRetryAtTick = 0u;
        return {KittyBitmapUploadDisposition::Uploaded, hr};
    }
    if (hr == D2DERR_RECREATE_TARGET)
    {
        image.uploadState              = KittyBitmapUploadState::RecreatePending;
        _kittyRecreateTargetAfterFrame = true;
        return {KittyBitmapUploadDisposition::RecreateTarget, hr};
    }
    if (IsTransientKittyUploadFailure(hr) && image.uploadAttemptCount < kMaximumKittyTransientUploadAttempts)
    {
        image.uploadState             = KittyBitmapUploadState::RetryPending;
        const DWORD delayMilliseconds = kInitialKittyUploadRetryDelayMilliseconds << ((image.uploadAttemptCount - 1u) * 2u);
        scheduleKittyUploadRetry(image, delayMilliseconds);
        return {KittyBitmapUploadDisposition::RetryScheduled, hr};
    }
    image.uploadState = KittyBitmapUploadState::PermanentFailure;
    ++_kittyStableUploadFailureCount;
    return {KittyBitmapUploadDisposition::PermanentFailure, hr};
}

bool Terminal::recoverKittyDeviceAfterFrame(HWND hwnd) noexcept
{
    if (! _kittyRecreateTargetAfterFrame)
    {
        return false;
    }
    _kittyRecreateTargetAfterFrame = false;
    discardDeviceResources();
    ++_kittyDeviceRecoveryCount;
    requestKittyRepaint(hwnd);
    return true;
}

void Terminal::drawKittyLayer(GhosttyKittyPlacementLayer layer) noexcept
{
    constexpr int32_t belowBackgroundMaximum = (std::numeric_limits<int32_t>::min)() / 2;
    const float pixelToDip                   = 96.0f / static_cast<float>(std::max<UINT>(_dpi, 1u));
    for (const KittyPlacementSnapshot& placement : _kittyPlacements)
    {
        const bool matches = layer == GHOSTTY_KITTY_PLACEMENT_LAYER_BELOW_BG     ? placement.z < belowBackgroundMaximum
                             : layer == GHOSTTY_KITTY_PLACEMENT_LAYER_BELOW_TEXT ? placement.z >= belowBackgroundMaximum && placement.z < 0
                                                                                 : layer == GHOSTTY_KITTY_PLACEMENT_LAYER_ABOVE_TEXT && placement.z >= 0;
        if (! matches)
        {
            continue;
        }
        ++_kittyFrameImageLookupCount;
        const auto imageIterator = _kittyImages.find(placement.imageKey);
        if (imageIterator == _kittyImages.end() || imageIterator->second.state != KittyImageState::Ready)
        {
            continue;
        }
        KittyImageSnapshot& image            = imageIterator->second;
        const KittyBitmapUploadResult upload = ensureKittyBitmap(image);
        if (upload.disposition != KittyBitmapUploadDisposition::Ready && upload.disposition != KittyBitmapUploadDisposition::Uploaded)
        {
            continue;
        }
        const auto& info = placement.renderInfo;
        const float left = static_cast<float>(info.viewport_col) * _cellWidthDip + static_cast<float>(placement.xOffsetPixels) * pixelToDip;
        const float top  = static_cast<float>(info.viewport_row) * _cellHeightDip + static_cast<float>(placement.yOffsetPixels) * pixelToDip;
        const D2D1_RECT_F destination =
            D2D1::RectF(left, top, left + static_cast<float>(info.pixel_width) * pixelToDip, top + static_cast<float>(info.pixel_height) * pixelToDip);
        const D2D1_RECT_F source = D2D1::RectF(static_cast<float>(info.source_x),
                                               static_cast<float>(info.source_y),
                                               static_cast<float>(info.source_x + info.source_width),
                                               static_cast<float>(info.source_y + info.source_height));
        _renderTarget->DrawBitmap(image.bitmap.get(), destination, 1.0f, D2D1_BITMAP_INTERPOLATION_MODE_LINEAR, source);
    }
}

bool Terminal::renderStructuredScreen(float widthDip, float heightDip) noexcept
{
    if (_renderState == nullptr || _renderRows == nullptr || _renderCells == nullptr)
    {
        return false;
    }

    TerminalKittyGenerationWork kittyWork;
    std::vector<KittyPlacementSnapshot> kittyPlacements;
    bool kittyCaptureReady           = false;
    bool submitKittyWork             = false;
    const auto kittyCaptureStartedAt = std::chrono::steady_clock::now();
    {
        std::scoped_lock lock(_terminalMutex);
        if (! _diagnosticText.empty() || _ghosttyTerminal == nullptr || _runtime.renderStateBeginUpdate(_renderState, _ghosttyTerminal) != GHOSTTY_SUCCESS)
        {
            return false;
        }
        kittyCaptureReady = captureKittyGraphicsLocked(kittyWork, kittyPlacements, submitKittyWork);
    }
    size_t kittySourceBytes = 0u;
    for (const TerminalKittyImageWork& image : kittyWork.images)
    {
        kittySourceBytes += image.source.size();
    }
    Debug::Perf::EmitDurationUs(L"terminal.kitty.capture_lock_us",
                                Debug::Perf::ElapsedUs(kittyCaptureStartedAt),
                                kittySourceBytes,
                                kittyPlacements.size(),
                                kittyCaptureReady ? S_OK : E_FAIL);
    if (_runtime.renderStateEndUpdate(_renderState) != GHOSTTY_SUCCESS)
    {
        return false;
    }

    if (! kittyCaptureReady)
    {
        _kittyPipeline.Invalidate(++_kittyRequestId);
        _kittyPlacements.clear();
        for (auto& [imageKey, image] : _kittyImages)
        {
            static_cast<void>(imageKey);
            if (image.state == KittyImageState::ConversionPending || image.state == KittyImageState::EnginePending)
            {
                image.state            = KittyImageState::Missing;
                image.pendingRequestId = 0u;
            }
        }
    }
    else if (kittyWork.storageGeneration == 0u)
    {
        if (_kittyTargetStorageGeneration != 0u)
        {
            _kittyPipeline.Invalidate(++_kittyRequestId);
        }
        _kittyTargetStorageGeneration = 0u;
        _kittyPlacements.clear();
        _kittyImages.clear();
        _kittyCachedBytes = 0u;
    }
    else
    {
        _kittyPlacements = std::move(kittyPlacements);
        if (submitKittyWork)
        {
            const uint64_t storageGeneration = kittyWork.storageGeneration;
            if (_kittyPipeline.Submit(std::move(kittyWork)))
            {
                _kittyTargetStorageGeneration = storageGeneration;
            }
            else
            {
                _kittyPipeline.Invalidate(_kittyRequestId);
                for (auto& [imageKey, image] : _kittyImages)
                {
                    static_cast<void>(imageKey);
                    if (image.state == KittyImageState::ConversionPending && image.pendingRequestId == _kittyRequestId)
                    {
                        image.state            = KittyImageState::Missing;
                        image.pendingRequestId = 0u;
                    }
                }
            }
        }
        consumeKittyReady();
    }
    const TerminalKittyPipelineStats kittyPipelineStats = _kittyPipeline.GetStats();
    Debug::Perf::EmitValue(L"terminal.kitty.placement_total_count", _kittyCaptureTotalPlacements);
    Debug::Perf::EmitValue(L"terminal.kitty.placement_visible_count", _kittyCaptureVisiblePlacements);
    Debug::Perf::EmitValue(L"terminal.kitty.placement_offscreen_count", _kittyCaptureOffscreenPlacements);
    Debug::Perf::EmitValue(L"terminal.kitty.placement_virtual_count", _kittyCaptureVirtualPlacements);
    Debug::Perf::EmitValue(L"terminal.kitty.visible_image_key_count", _kittyCaptureVisibleImageKeys);
    Debug::Perf::EmitValue(L"terminal.kitty.engine_placement_count", _kittyCaptureEnginePlacementCount);
    Debug::Perf::EmitValue(L"terminal.kitty.engine_placement_bytes", _kittyCaptureEnginePlacementBytes);
    Debug::Perf::EmitValue(L"terminal.kitty.source_bytes", _kittyCaptureSourceBytes);
    Debug::Perf::EmitValue(L"terminal.kitty.pipeline_active_source_bytes", kittyPipelineStats.activeSourceBytes);
    Debug::Perf::EmitValue(L"terminal.kitty.pipeline_active_converted_bytes", kittyPipelineStats.activeConvertedBytes);
    Debug::Perf::EmitValue(L"terminal.kitty.pipeline_ready_bytes", kittyPipelineStats.readyBytes);
    Debug::Perf::EmitValue(L"terminal.kitty.cache_bytes", _kittyCachedBytes);
    Debug::Perf::EmitValue(L"terminal.kitty.pinned_bytes", _kittyCapturePinnedBytes);
    Debug::Perf::EmitValue(L"terminal.kitty.evicted_bytes", _kittyEvictedBytes);
    Debug::Perf::EmitValue(L"terminal.kitty.evicted_image_count", _kittyEvictedImageCount);
    Debug::Perf::EmitValue(L"terminal.kitty.missing_image_key_count", _kittyCaptureMissingImageKeys);
    Debug::Perf::EmitValue(L"terminal.kitty.engine_pending_image_key_count", _kittyCaptureEnginePendingImageKeys);
    Debug::Perf::EmitValue(L"terminal.kitty.deferred_image_key_count", _kittyCaptureDeferredImageKeys);
    Debug::Perf::EmitValue(L"terminal.kitty.rejected_image_key_count", _kittyCaptureRejectedImageKeys);
    Debug::Perf::EmitValue(L"terminal.kitty.superseded_image_count", _kittySupersededImageCount);
    Debug::Perf::EmitValue(L"terminal.kitty.removed_image_count", _kittyRemovedImageCount);
    _kittyFrameUploadBytes         = 0u;
    _kittyFrameUploadCount         = 0u;
    _kittyFrameImageLookupCount    = 0u;
    _kittyUploadDeferred           = false;
    _kittyRecreateTargetAfterFrame = false;

    uint16_t columns = 0u;
    uint16_t rows    = 0u;
    if (_runtime.renderStateGet(_renderState, GHOSTTY_RENDER_STATE_DATA_COLS, &columns) != GHOSTTY_SUCCESS ||
        _runtime.renderStateGet(_renderState, GHOSTTY_RENDER_STATE_DATA_ROWS, &rows) != GHOSTTY_SUCCESS ||
        _runtime.renderStateGet(_renderState, GHOSTTY_RENDER_STATE_DATA_ROW_ITERATOR, &_renderRows) != GHOSTTY_SUCCESS)
    {
        return false;
    }

    drawKittyLayer(GHOSTTY_KITTY_PLACEMENT_LAYER_BELOW_BG);

    {
        const D2D1_ANTIALIAS_MODE previousAntialiasMode = _renderTarget->GetAntialiasMode();
        _renderTarget->SetAntialiasMode(D2D1_ANTIALIAS_MODE_ALIASED);
        const auto restoreAntialiasMode = wil::scope_exit([&]() noexcept { _renderTarget->SetAntialiasMode(previousAntialiasMode); });

        uint16_t backgroundRowIndex = 0u;
        while (backgroundRowIndex < rows && _runtime.renderRowIteratorNext(_renderRows))
        {
            if (_runtime.renderRowGet(_renderRows, GHOSTTY_RENDER_STATE_ROW_DATA_CELLS, &_renderCells) != GHOSTTY_SUCCESS)
            {
                ++backgroundRowIndex;
                continue;
            }
            uint16_t backgroundColumnIndex = 0u;
            while (backgroundColumnIndex < columns && _runtime.renderRowCellsNext(_renderCells))
            {
                GhosttyCell rawCell  = 0u;
                GhosttyCellWide wide = GHOSTTY_CELL_WIDE_NARROW;
                GhosttyStyle style{};
                style.size      = sizeof(style);
                bool selected   = false;
                bool hasStyling = false;
                GhosttyColorRgb background{};
                GhosttyColorRgb foreground{};
                static_cast<void>(_runtime.renderRowCellsGet(_renderCells, GHOSTTY_RENDER_STATE_ROW_CELLS_DATA_RAW, &rawCell));
                static_cast<void>(_runtime.cellGet(rawCell, GHOSTTY_CELL_DATA_WIDE, &wide));
                static_cast<void>(_runtime.renderRowCellsGet(_renderCells, GHOSTTY_RENDER_STATE_ROW_CELLS_DATA_SELECTED, &selected));
                static_cast<void>(_runtime.renderRowCellsGet(_renderCells, GHOSTTY_RENDER_STATE_ROW_CELLS_DATA_HAS_STYLING, &hasStyling));
                if (hasStyling)
                {
                    static_cast<void>(_runtime.renderRowCellsGet(_renderCells, GHOSTTY_RENDER_STATE_ROW_CELLS_DATA_STYLE, &style));
                }
                uint32_t backgroundArgb = _backgroundArgb;
                uint32_t foregroundArgb = _foregroundArgb;
                const bool explicitBackground =
                    _runtime.renderRowCellsGet(_renderCells, GHOSTTY_RENDER_STATE_ROW_CELLS_DATA_BG_COLOR, &background) == GHOSTTY_SUCCESS;
                const bool explicitForeground =
                    _runtime.renderRowCellsGet(_renderCells, GHOSTTY_RENDER_STATE_ROW_CELLS_DATA_FG_COLOR, &foreground) == GHOSTTY_SUCCESS;
                if (explicitBackground)
                {
                    backgroundArgb = ArgbFromGhostty(background);
                }
                if (explicitForeground)
                {
                    foregroundArgb = ArgbFromGhostty(foreground);
                }
                if (style.inverse)
                {
                    std::swap(backgroundArgb, foregroundArgb);
                }
                if (selected)
                {
                    backgroundArgb = _selectionBackgroundArgb;
                }
                if (explicitBackground || selected || style.inverse)
                {
                    const float left           = static_cast<float>(backgroundColumnIndex) * _cellWidthDip;
                    const float top            = static_cast<float>(backgroundRowIndex) * _cellHeightDip;
                    const float cellWidth      = wide == GHOSTTY_CELL_WIDE_WIDE ? _cellWidthDip * 2.0f : _cellWidthDip;
                    const D2D1_RECT_F cellRect = D2D1::RectF(left, top, std::min(widthDip, left + cellWidth), std::min(heightDip, top + _cellHeightDip));
                    _foregroundBrush->SetColor(RedSalamander::DxUi::ColorFromArgb(backgroundArgb));
                    _renderTarget->FillRectangle(cellRect, _foregroundBrush.get());
                }
                ++backgroundColumnIndex;
            }
            ++backgroundRowIndex;
        }
    }
    drawKittyLayer(GHOSTTY_KITTY_PLACEMENT_LAYER_BELOW_TEXT);
    if (_runtime.renderStateGet(_renderState, GHOSTTY_RENDER_STATE_DATA_ROW_ITERATOR, &_renderRows) != GHOSTTY_SUCCESS)
    {
        return false;
    }

    uint16_t rowIndex = 0u;
    while (rowIndex < rows && _runtime.renderRowIteratorNext(_renderRows))
    {
        if (_runtime.renderRowGet(_renderRows, GHOSTTY_RENDER_STATE_ROW_DATA_CELLS, &_renderCells) != GHOSTTY_SUCCESS)
        {
            ++rowIndex;
            continue;
        }

        uint16_t columnIndex = 0u;
        while (columnIndex < columns && _runtime.renderRowCellsNext(_renderCells))
        {
            GhosttyCell rawCell  = 0u;
            GhosttyCellWide wide = GHOSTTY_CELL_WIDE_NARROW;
            GhosttyStyle style{};
            style.size      = sizeof(style);
            bool selected   = false;
            bool hasStyling = false;
            bool hyperlink  = false;
            GhosttyColorRgb background{};
            GhosttyColorRgb foreground{};

            static_cast<void>(_runtime.renderRowCellsGet(_renderCells, GHOSTTY_RENDER_STATE_ROW_CELLS_DATA_RAW, &rawCell));
            static_cast<void>(_runtime.cellGet(rawCell, GHOSTTY_CELL_DATA_WIDE, &wide));
            static_cast<void>(_runtime.cellGet(rawCell, GHOSTTY_CELL_DATA_HAS_HYPERLINK, &hyperlink));
            static_cast<void>(_runtime.renderRowCellsGet(_renderCells, GHOSTTY_RENDER_STATE_ROW_CELLS_DATA_SELECTED, &selected));
            static_cast<void>(_runtime.renderRowCellsGet(_renderCells, GHOSTTY_RENDER_STATE_ROW_CELLS_DATA_HAS_STYLING, &hasStyling));
            if (hasStyling)
            {
                static_cast<void>(_runtime.renderRowCellsGet(_renderCells, GHOSTTY_RENDER_STATE_ROW_CELLS_DATA_STYLE, &style));
            }

            uint32_t backgroundArgb = _backgroundArgb;
            uint32_t foregroundArgb = hyperlink ? _hyperlinkArgb : _foregroundArgb;
            const bool explicitBackground =
                _runtime.renderRowCellsGet(_renderCells, GHOSTTY_RENDER_STATE_ROW_CELLS_DATA_BG_COLOR, &background) == GHOSTTY_SUCCESS;
            const bool explicitForeground =
                _runtime.renderRowCellsGet(_renderCells, GHOSTTY_RENDER_STATE_ROW_CELLS_DATA_FG_COLOR, &foreground) == GHOSTTY_SUCCESS;
            if (explicitBackground)
            {
                backgroundArgb = ArgbFromGhostty(background);
            }
            if (explicitForeground)
            {
                foregroundArgb = ArgbFromGhostty(foreground);
            }
            if (style.inverse)
            {
                std::swap(backgroundArgb, foregroundArgb);
            }
            if (selected)
            {
                backgroundArgb = _selectionBackgroundArgb;
                foregroundArgb = _selectionForegroundArgb;
            }
            if (style.faint)
            {
                foregroundArgb = WithAlpha(foregroundArgb, 0x8Fu);
            }

            const float left           = static_cast<float>(columnIndex) * _cellWidthDip;
            const float top            = static_cast<float>(rowIndex) * _cellHeightDip;
            const float cellWidth      = wide == GHOSTTY_CELL_WIDE_WIDE ? _cellWidthDip * 2.0f : _cellWidthDip;
            const D2D1_RECT_F cellRect = D2D1::RectF(left, top, std::min(widthDip, left + cellWidth), std::min(heightDip, top + _cellHeightDip));
            if (wide != GHOSTTY_CELL_WIDE_SPACER_TAIL && wide != GHOSTTY_CELL_WIDE_SPACER_HEAD && ! style.invisible)
            {
                std::array<uint8_t, 256u> graphemeBytes{};
                GhosttyBuffer grapheme{graphemeBytes.data(), graphemeBytes.size(), 0u};
                if (_runtime.renderRowCellsGet(_renderCells, GHOSTTY_RENDER_STATE_ROW_CELLS_DATA_GRAPHEMES_UTF8, &grapheme) == GHOSTTY_SUCCESS &&
                    grapheme.len != 0u)
                {
                    const std::wstring text =
                        Common::Strings::Utf16FromUtf8ReplacingInvalid(std::string_view(reinterpret_cast<const char*>(grapheme.ptr), grapheme.len));
                    IDWriteTextFormat* textFormat = _textFormat.get();
                    if (style.bold && style.italic)
                    {
                        textFormat = _boldItalicTextFormat.get();
                    }
                    else if (style.bold)
                    {
                        textFormat = _boldTextFormat.get();
                    }
                    else if (style.italic)
                    {
                        textFormat = _italicTextFormat.get();
                    }
                    _foregroundBrush->SetColor(RedSalamander::DxUi::ColorFromArgb(foregroundArgb));
                    _renderTarget->DrawTextW(
                        text.data(), static_cast<UINT32>(text.size()), textFormat, cellRect, _foregroundBrush.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);
                }
            }

            _foregroundBrush->SetColor(RedSalamander::DxUi::ColorFromArgb(foregroundArgb));
            const float lineWidth = std::max(1.0f, 96.0f / static_cast<float>(std::max<UINT>(_dpi, 1u)));
            if (style.underline != GHOSTTY_SGR_UNDERLINE_NONE || hyperlink)
            {
                const float y = std::max(cellRect.top, cellRect.bottom - lineWidth);
                _renderTarget->DrawLine(D2D1::Point2F(cellRect.left, y), D2D1::Point2F(cellRect.right, y), _foregroundBrush.get(), lineWidth);
            }
            if (style.strikethrough)
            {
                const float y = cellRect.top + (cellRect.bottom - cellRect.top) * 0.52f;
                _renderTarget->DrawLine(D2D1::Point2F(cellRect.left, y), D2D1::Point2F(cellRect.right, y), _foregroundBrush.get(), lineWidth);
            }
            if (style.overline)
            {
                _renderTarget->DrawLine(
                    D2D1::Point2F(cellRect.left, cellRect.top), D2D1::Point2F(cellRect.right, cellRect.top), _foregroundBrush.get(), lineWidth);
            }
            ++columnIndex;
        }
        ++rowIndex;
    }

    bool cursorVisible                              = false;
    bool cursorInViewport                           = false;
    uint16_t cursorX                                = 0u;
    uint16_t cursorY                                = 0u;
    GhosttyRenderStateCursorVisualStyle cursorStyle = GHOSTTY_RENDER_STATE_CURSOR_VISUAL_STYLE_BLOCK;
    if (_runtime.renderStateGet(_renderState, GHOSTTY_RENDER_STATE_DATA_CURSOR_VISIBLE, &cursorVisible) == GHOSTTY_SUCCESS &&
        _runtime.renderStateGet(_renderState, GHOSTTY_RENDER_STATE_DATA_CURSOR_VIEWPORT_HAS_VALUE, &cursorInViewport) == GHOSTTY_SUCCESS && cursorVisible &&
        cursorInViewport && _runtime.renderStateGet(_renderState, GHOSTTY_RENDER_STATE_DATA_CURSOR_VIEWPORT_X, &cursorX) == GHOSTTY_SUCCESS &&
        _runtime.renderStateGet(_renderState, GHOSTTY_RENDER_STATE_DATA_CURSOR_VIEWPORT_Y, &cursorY) == GHOSTTY_SUCCESS)
    {
        static_cast<void>(_runtime.renderStateGet(_renderState, GHOSTTY_RENDER_STATE_DATA_CURSOR_VISUAL_STYLE, &cursorStyle));
        const float left       = static_cast<float>(cursorX) * _cellWidthDip;
        const float top        = static_cast<float>(cursorY) * _cellHeightDip;
        D2D1_RECT_F cursorRect = D2D1::RectF(left, top, left + _cellWidthDip, top + _cellHeightDip);
        switch (cursorStyle)
        {
            case GHOSTTY_RENDER_STATE_CURSOR_VISUAL_STYLE_BAR: cursorRect.right = cursorRect.left + std::max(1.0f, _cellWidthDip * 0.16f); break;
            case GHOSTTY_RENDER_STATE_CURSOR_VISUAL_STYLE_UNDERLINE: cursorRect.top = cursorRect.bottom - std::max(1.0f, _cellHeightDip * 0.12f); break;
            case GHOSTTY_RENDER_STATE_CURSOR_VISUAL_STYLE_BLOCK_HOLLOW:
            case GHOSTTY_RENDER_STATE_CURSOR_VISUAL_STYLE_BLOCK:
            case GHOSTTY_RENDER_STATE_CURSOR_VISUAL_STYLE_MAX_VALUE: break;
        }
        _foregroundBrush->SetColor(RedSalamander::DxUi::ColorFromArgb(WithAlpha(_cursorArgb, 0xB0u)));
        if (cursorStyle == GHOSTTY_RENDER_STATE_CURSOR_VISUAL_STYLE_BLOCK_HOLLOW || GetFocus() != _windowHandle.load(std::memory_order_acquire))
        {
            _renderTarget->DrawRectangle(cursorRect, _foregroundBrush.get());
        }
        else
        {
            _renderTarget->FillRectangle(cursorRect, _foregroundBrush.get());
        }
    }
    if (! _imeComposition.empty() && cursorInViewport)
    {
        const float left           = static_cast<float>(cursorX) * _cellWidthDip;
        const float top            = static_cast<float>(cursorY) * _cellHeightDip;
        const float requestedWidth = static_cast<float>(_imeComposition.size()) * _cellWidthDip;
        const D2D1_RECT_F compositionRect =
            D2D1::RectF(left, top, std::min(widthDip, left + std::max(_cellWidthDip, requestedWidth)), std::min(heightDip, top + _cellHeightDip));
        _foregroundBrush->SetColor(RedSalamander::DxUi::ColorFromArgb(_foregroundArgb));
        _renderTarget->DrawTextW(_imeComposition.data(),
                                 static_cast<UINT32>(std::min<size_t>(_imeComposition.size(), (std::numeric_limits<UINT32>::max)())),
                                 _textFormat.get(),
                                 compositionRect,
                                 _foregroundBrush.get(),
                                 D2D1_DRAW_TEXT_OPTIONS_CLIP);
        const float underlineY = std::max(compositionRect.top, compositionRect.bottom - 1.0f);
        _renderTarget->DrawLine(
            D2D1::Point2F(compositionRect.left, underlineY), D2D1::Point2F(compositionRect.right, underlineY), _foregroundBrush.get(), 1.0f);
    }
    drawKittyLayer(GHOSTTY_KITTY_PLACEMENT_LAYER_ABOVE_TEXT);
    return true;
}

void Terminal::clearRenderDirtyState() noexcept
{
    if (_renderState == nullptr || _renderRows == nullptr ||
        _runtime.renderStateGet(_renderState, GHOSTTY_RENDER_STATE_DATA_ROW_ITERATOR, &_renderRows) != GHOSTTY_SUCCESS)
    {
        return;
    }
    const bool clean = false;
    while (_runtime.renderRowIteratorNext(_renderRows))
    {
        static_cast<void>(_runtime.renderRowSet(_renderRows, GHOSTTY_RENDER_STATE_ROW_OPTION_DIRTY, &clean));
    }
    const GhosttyRenderStateDirty dirty = GHOSTTY_RENDER_STATE_DIRTY_FALSE;
    static_cast<void>(_runtime.renderStateSet(_renderState, GHOSTTY_RENDER_STATE_OPTION_DIRTY, &dirty));
}

void Terminal::publishAccessibilitySnapshot() noexcept
{
    if (! _accessibility)
    {
        return;
    }
    const HWND hwnd   = _windowHandle.load(std::memory_order_acquire);
    std::wstring text = formatScreen();
    if (std::wstring confirmation = confirmationAccessibilityText(); ! confirmation.empty())
    {
        if (! text.empty())
        {
            text.append(L"\r\n");
        }
        text.append(confirmation);
    }
    if (std::wstring surface = commandSurfaceAccessibilityText(); ! surface.empty())
    {
        if (! text.empty())
        {
            text.append(L"\r\n");
        }
        text.append(surface);
    }
    _accessibility->Publish(std::move(text), hwnd != nullptr && GetFocus() == hwnd);
}

void Terminal::render() noexcept
{
    Debug::Perf::Scope renderPerf(L"terminal.render.frame_us");
    renderPerf.SetValue0(_columns);
    renderPerf.SetValue1(_rows);
    const auto frameStartedAt   = std::chrono::steady_clock::now();
    _kittyFrameUploadBytes      = 0u;
    _kittyFrameUploadCount      = 0u;
    _kittyFrameImageLookupCount = 0u;
    _kittyUploadDeferred        = false;
    PAINTSTRUCT paint{};
    const HWND hwnd = _windowHandle.load(std::memory_order_acquire);
    if (hwnd == nullptr)
    {
        return;
    }
    const bool paneFocused = GetFocus() == hwnd;
    renderPerf.SetDetail(paneFocused ? L"focused" : L"unfocused");
    const wil::unique_hdc_paint paintDc = wil::BeginPaint(hwnd, &paint);
    const HRESULT resourcesHr           = ensureDeviceResources();
    if (FAILED(resourcesHr))
    {
        renderPerf.SetHr(resourcesHr);
        return;
    }

    const HRESULT resizeHr = resizeRenderTargetToClient();
    if (resizeHr == D2DERR_RECREATE_TARGET)
    {
        _kittyRecreateTargetAfterFrame = true;
        static_cast<void>(recoverKittyDeviceAfterFrame(hwnd));
        renderPerf.SetHr(resizeHr);
        return;
    }
    if (FAILED(resizeHr))
    {
        renderPerf.SetHr(resizeHr);
        return;
    }
    RECT client{};
    GetClientRect(hwnd, &client);
    const D2D1_SIZE_U size =
        D2D1::SizeU(static_cast<UINT32>(std::max(0L, client.right - client.left)), static_cast<UINT32>(std::max(0L, client.bottom - client.top)));

    _renderTarget->BeginDraw();
    _renderTarget->Clear(RedSalamander::DxUi::ColorFromArgb(_backgroundArgb));
    _foregroundBrush->SetColor(RedSalamander::DxUi::ColorFromArgb(_foregroundArgb));
    const float widthDip  = static_cast<float>(size.width) * 96.0f / static_cast<float>(std::max<UINT>(_dpi, 1u));
    const float heightDip = static_cast<float>(size.height) * 96.0f / static_cast<float>(std::max<UINT>(_dpi, 1u));
    const bool structured = renderStructuredScreen(widthDip, heightDip);
    if (! structured)
    {
        const std::wstring text = formatScreen();
        _renderTarget->DrawTextW(text.data(),
                                 static_cast<UINT32>(std::min<size_t>(text.size(), (std::numeric_limits<UINT32>::max)())),
                                 _textFormat.get(),
                                 D2D1::RectF(4.0f, 2.0f, std::max(4.0f, widthDip - 4.0f), std::max(2.0f, heightDip - 2.0f)),
                                 _foregroundBrush.get(),
                                 D2D1_DRAW_TEXT_OPTIONS_CLIP);
    }
    drawCommandSurface(widthDip, heightDip);
    drawConfirmation(widthDip, heightDip);
    drawInactivePaneOverlay(widthDip, heightDip, paneFocused);
    const HRESULT drawHr = _renderTarget->EndDraw();
    renderPerf.SetHr(drawHr);
    if (drawHr == D2DERR_RECREATE_TARGET)
    {
        _kittyRecreateTargetAfterFrame = true;
    }
    const bool recoveredDevice = recoverKittyDeviceAfterFrame(hwnd);
    if (! recoveredDevice && SUCCEEDED(drawHr) && structured)
    {
        clearRenderDirtyState();
    }
    Debug::Perf::EmitDurationUs(L"terminal.frame_total_us", Debug::Perf::ElapsedUs(frameStartedAt), _kittyFrameUploadBytes, _kittyFrameUploadCount, drawHr);
    Debug::Perf::EmitValue(L"terminal.kitty.image_lookup_count", _kittyFrameImageLookupCount);
    Debug::Perf::EmitValue(L"terminal.kitty.frame_upload_bytes", _kittyFrameUploadBytes);
    Debug::Perf::EmitValue(L"terminal.kitty.frame_upload_count", _kittyFrameUploadCount);
    if (SUCCEEDED(drawHr))
    {
        const int64_t nowNs    = SteadyTimestampNs();
        const int64_t inputNs  = _lastInputTimestampNs.exchange(0, std::memory_order_acq_rel);
        const int64_t outputNs = _lastOutputTimestampNs.exchange(0, std::memory_order_acq_rel);
        if (inputNs > 0 && nowNs >= inputNs)
        {
            Debug::Perf::EmitDurationUs(L"terminal.input_to_frame_us", static_cast<uint64_t>(nowNs - inputNs) / 1000u);
        }
        if (outputNs > 0 && nowNs >= outputNs)
        {
            Debug::Perf::EmitDurationUs(L"terminal.output_to_frame_us", static_cast<uint64_t>(nowNs - outputNs) / 1000u);
        }
    }
    if (! recoveredDevice && _kittyUploadDeferred)
    {
        requestKittyRepaint(hwnd);
    }
}

void Terminal::drawInactivePaneOverlay(float widthDip, float heightDip, bool paneFocused) noexcept
{
    const float overlayAlpha = Common::PaneVisualState::ResolveInactiveContentOverlayAlpha(paneFocused, _themeHighContrast);
    if (overlayAlpha <= 0.0f || ! _renderTarget || ! _foregroundBrush)
    {
        return;
    }

    D2D1_COLOR_F overlayColor = RedSalamander::DxUi::ColorFromArgb(_backgroundArgb);
    overlayColor.a            = overlayAlpha;
    _foregroundBrush->SetColor(overlayColor);
    _renderTarget->FillRectangle(D2D1::RectF(0.0f, 0.0f, widthDip, heightDip), _foregroundBrush.get());
}
