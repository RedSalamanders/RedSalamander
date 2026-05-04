#include "FolderWindow.FileOperations.Popup.h"

#include "DxUi/DxUi.Typography.h"
#include "DxUi/DxUi.h"
#include "DxUiThemePalette.h"
#include "FluentIcons.h"
#include "FolderWindow.FileOperationsInternal.h"
#include "HostServices.h"
#include "NavigationLocation.h"
#include "WindowMaximizeBehavior.h"

#include <algorithm>
#include <array>
#include <chrono>
#include <cmath>
#include <d2d1.h>
#include <dwrite.h>
#include <limits>
#include <unordered_map>
#include <windowsx.h>

namespace
{
constexpr wchar_t kFileOperationsPopupClassName[] = L"RedSalamander.FileOperationsPopup";

constexpr UINT_PTR kFileOperationsPopupTimerId                     = 1;
constexpr UINT kFileOperationsPopupTimerIntervalMs                 = 100;
constexpr ULONGLONG kRateSampleBucketMs                            = 100ull;
constexpr UINT kFileOperationsPopupDeferredSpeedLimitPromptMessage = WM_APP + 0x71;
constexpr wchar_t kFileOperationsSpeedLimitPromptClassName[]       = L"RedSalamander.FileOperations.SpeedLimitPrompt";
constexpr std::wstring_view kEllipsisText                          = L"\u2026";

#ifdef ENABLE_TESTS
constexpr UINT kFileOperationsSpeedLimitPromptDebugMessage = WM_APP + 0x73;

[[nodiscard]] bool IsActuallyVisibleChildWindow(HWND hwnd) noexcept
{
    return hwnd && IsWindow(hwnd) != FALSE && IsWindowVisible(hwnd) != FALSE && (GetWindowLongPtrW(hwnd, GWL_STYLE) & WS_CHILD) != 0;
}
#endif

[[nodiscard]] uint64_t PerfNowUs() noexcept
{
    return static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now().time_since_epoch()).count());
}

[[nodiscard]] uint64_t PerfElapsedUs(uint64_t startUs) noexcept
{
    const uint64_t nowUs = PerfNowUs();
    return (nowUs >= startUs) ? (nowUs - startUs) : 0u;
}

float DipsToPixels(float dip, UINT dpi) noexcept
{
    return dip * (static_cast<float>(dpi) / static_cast<float>(USER_DEFAULT_SCREEN_DPI));
}

int DipsToPixels(int dip, UINT dpi) noexcept
{
    return MulDiv(dip, static_cast<int>(dpi), USER_DEFAULT_SCREEN_DPI);
}

float PixelsToDips(float px, UINT dpi) noexcept
{
    if (dpi == 0)
    {
        return px;
    }
    return px * (static_cast<float>(USER_DEFAULT_SCREEN_DPI) / static_cast<float>(dpi));
}

D2D1_RECT_F RectPixelsToDips(const RECT& rc, UINT dpi) noexcept
{
    return D2D1::RectF(PixelsToDips(static_cast<float>(rc.left), dpi),
                       PixelsToDips(static_cast<float>(rc.top), dpi),
                       PixelsToDips(static_cast<float>(rc.right), dpi),
                       PixelsToDips(static_cast<float>(rc.bottom), dpi));
}

[[nodiscard]] bool DirectWriteFormatHasGlyph(IDWriteFactory* factory, IDWriteTextFormat* format, wchar_t glyph) noexcept
{
    if (! factory || ! format || glyph == 0)
    {
        return false;
    }

    const UINT32 familyLength = format->GetFontFamilyNameLength();
    if (familyLength == 0)
    {
        return false;
    }

    std::wstring familyName(familyLength + 1u, L'\0');
    if (FAILED(format->GetFontFamilyName(familyName.data(), familyLength + 1u)))
    {
        return false;
    }
    familyName.resize(familyLength);

    wil::com_ptr<IDWriteFontCollection> collection;
    if (FAILED(factory->GetSystemFontCollection(collection.addressof(), FALSE)) || ! collection)
    {
        return false;
    }

    UINT32 familyIndex = 0;
    BOOL familyExists  = FALSE;
    if (FAILED(collection->FindFamilyName(familyName.c_str(), &familyIndex, &familyExists)) || familyExists == FALSE)
    {
        return false;
    }

    wil::com_ptr<IDWriteFontFamily> fontFamily;
    if (FAILED(collection->GetFontFamily(familyIndex, fontFamily.addressof())) || ! fontFamily)
    {
        return false;
    }

    wil::com_ptr<IDWriteFont> font;
    if (FAILED(fontFamily->GetFirstMatchingFont(format->GetFontWeight(), format->GetFontStretch(), format->GetFontStyle(), font.addressof())) || ! font)
    {
        return false;
    }

    BOOL hasGlyph = FALSE;
    if (FAILED(font->HasCharacter(static_cast<UINT32>(glyph), &hasGlyph)))
    {
        return false;
    }
    return hasGlyph != FALSE;
}

void CenterWindowOnOwnerWindow(HWND hwnd, HWND ownerWindow) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE || ! ownerWindow || IsWindow(ownerWindow) == FALSE)
    {
        return;
    }

    RECT windowRect{};
    RECT ownerRect{};
    if (GetWindowRect(hwnd, &windowRect) == FALSE || GetWindowRect(ownerWindow, &ownerRect) == FALSE)
    {
        return;
    }

    HMONITOR monitor = MonitorFromWindow(ownerWindow, MONITOR_DEFAULTTONEAREST);
    MONITORINFO mi{};
    mi.cbSize = sizeof(mi);
    if (GetMonitorInfoW(monitor, &mi) == FALSE)
    {
        return;
    }

    const LONG width     = windowRect.right - windowRect.left;
    const LONG height    = windowRect.bottom - windowRect.top;
    const LONG centeredX = ownerRect.left + ((ownerRect.right - ownerRect.left - width) / 2);
    const LONG centeredY = ownerRect.top + ((ownerRect.bottom - ownerRect.top - height) / 2);
    const LONG maxX      = std::max(mi.rcWork.left, mi.rcWork.right - width);
    const LONG maxY      = std::max(mi.rcWork.top, mi.rcWork.bottom - height);
    const int x          = std::clamp(centeredX, mi.rcWork.left, maxX);
    const int y          = std::clamp(centeredY, mi.rcWork.top, maxY);

    SetWindowPos(hwnd, nullptr, x, y, 0, 0, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
}

bool IsRectFullyVisible(const RECT& rect) noexcept
{
    if (rect.right <= rect.left || rect.bottom <= rect.top)
    {
        return false;
    }

    HMONITOR monitor = MonitorFromRect(&rect, MONITOR_DEFAULTTONULL);
    if (! monitor)
    {
        return false;
    }

    MONITORINFO mi{};
    mi.cbSize = sizeof(mi);
    if (! GetMonitorInfoW(monitor, &mi))
    {
        return false;
    }

    const RECT& work = mi.rcWork;
    return rect.left >= work.left && rect.top >= work.top && rect.right <= work.right && rect.bottom <= work.bottom;
}

float Clamp01(float v) noexcept
{
    return std::clamp(v, 0.0f, 1.0f);
}

float ComputeFileOperationsTaskCompleteFractionForDisplay(const FileOperationsPopupInternal::TaskSnapshot& task) noexcept
{
    if (task.finished && SUCCEEDED(task.resultHr))
    {
        return 1.0f;
    }

    if (task.operation == FILESYSTEM_DELETE)
    {
        if (task.totalBytes > 0 && task.completedBytes > 0)
        {
            return Clamp01(static_cast<float>(static_cast<double>(task.completedBytes) / static_cast<double>(task.totalBytes)));
        }
        if (task.totalItems > 0)
        {
            return Clamp01(static_cast<float>(static_cast<double>(task.completedItems) / static_cast<double>(task.totalItems)));
        }
        return 0.0f;
    }

    if (task.totalBytes > 0)
    {
        return Clamp01(static_cast<float>(static_cast<double>(task.completedBytes) / static_cast<double>(task.totalBytes)));
    }
    if (task.totalItems > 0)
    {
        return Clamp01(static_cast<float>(static_cast<double>(task.completedItems) / static_cast<double>(task.totalItems)));
    }

    return 0.0f;
}

void NormalizeCompletedTaskSnapshotForDisplay(FileOperationsPopupInternal::TaskSnapshot& task) noexcept
{
    if (! task.finished || FAILED(task.resultHr))
    {
        return;
    }

    if (task.totalItems > 0)
    {
        task.completedItems = task.totalItems;
    }
    if (task.totalBytes > 0)
    {
        task.completedBytes = task.totalBytes;
    }
    if (task.itemTotalBytes > 0)
    {
        task.itemCompletedBytes = task.itemTotalBytes;
    }
}

D2D1_RECT_F ComputeIndeterminateBarFill(const D2D1_RECT_F& bar, ULONGLONG tick) noexcept
{
    const float width = bar.right - bar.left;
    if (width <= 0.0f)
    {
        return bar;
    }

    constexpr ULONGLONG kPeriodMs = 1200ull;
    const float segmentW          = width * 0.28f;

    const ULONGLONG phaseMs = tick % kPeriodMs;
    const float t           = static_cast<float>(phaseMs) / static_cast<float>(kPeriodMs);

    const float travel = width + segmentW;
    const float x      = bar.left + travel * t - segmentW;

    const float left  = std::clamp(x, bar.left, bar.right);
    const float right = std::clamp(x + segmentW, bar.left, bar.right);
    return D2D1::RectF(left, bar.top, right, bar.bottom);
}

float ClampCornerRadius(const D2D1_RECT_F& rc, float desired) noexcept
{
    const float w         = std::max(0.0f, rc.right - rc.left);
    const float h         = std::max(0.0f, rc.bottom - rc.top);
    const float maxRadius = std::min(w, h) * 0.5f;
    return std::clamp(desired, 0.0f, maxRadius);
}

std::wstring FormatDurationHms(uint64_t seconds)
{
    const uint64_t hours64   = seconds / 3600u;
    const uint64_t minutes64 = (seconds % 3600u) / 60u;
    const uint64_t seconds64 = seconds % 60u;

    const unsigned long long hours = static_cast<unsigned long long>(hours64);
    const unsigned int minutes     = static_cast<unsigned int>(minutes64);
    const unsigned int secs        = static_cast<unsigned int>(seconds64);

    if (hours > 0ull)
    {
        return std::format(L"{}:{:02d}:{:02d}", hours, minutes, secs);
    }

    return std::format(L"{:02d}:{:02d}", minutes, secs);
}

bool IsAsciiSpace(wchar_t ch) noexcept
{
    return ch == L' ' || ch == L'\t' || ch == L'\n' || ch == L'\r' || ch == L'\f' || ch == L'\v';
}

std::wstring_view TrimAscii(std::wstring_view text) noexcept
{
    while (! text.empty() && IsAsciiSpace(text.front()))
    {
        text.remove_prefix(1);
    }
    while (! text.empty() && IsAsciiSpace(text.back()))
    {
        text.remove_suffix(1);
    }
    return text;
}

wchar_t FoldAsciiCase(wchar_t ch) noexcept
{
    if (ch >= L'A' && ch <= L'Z')
    {
        return static_cast<wchar_t>(ch - L'A' + L'a');
    }
    return ch;
}

bool EqualsIgnoreAsciiCase(std::wstring_view a, std::wstring_view b) noexcept
{
    if (a.size() != b.size())
    {
        return false;
    }

    for (size_t i = 0; i < a.size(); ++i)
    {
        if (FoldAsciiCase(a[i]) != FoldAsciiCase(b[i]))
        {
            return false;
        }
    }

    return true;
}

bool TryParseThroughputText(std::wstring_view text, uint64_t& outBytesPerSecond) noexcept
{
    constexpr uint64_t kKiB = 1024ull;
    constexpr uint64_t kMiB = 1024ull * 1024ull;
    constexpr uint64_t kGiB = 1024ull * 1024ull * 1024ull;
    constexpr uint64_t kTiB = 1024ull * 1024ull * 1024ull * 1024ull;
    constexpr uint64_t kPiB = 1024ull * 1024ull * 1024ull * 1024ull * 1024ull;

    outBytesPerSecond = 0;

    text = TrimAscii(text);
    if (text.empty())
    {
        return true;
    }

    bool sawDigit          = false;
    bool sawDecimal        = false;
    double number          = 0.0;
    double fractionalScale = 0.1;
    size_t index           = 0;
    for (; index < text.size(); ++index)
    {
        const wchar_t ch = text[index];
        if (ch >= L'0' && ch <= L'9')
        {
            sawDigit                 = true;
            const unsigned int digit = static_cast<unsigned int>(ch - L'0');
            if (! sawDecimal)
            {
                number = number * 10.0 + static_cast<double>(digit);
            }
            else
            {
                number += static_cast<double>(digit) * fractionalScale;
                fractionalScale *= 0.1;
            }
            continue;
        }

        if ((ch == L'.' || ch == L',') && ! sawDecimal)
        {
            sawDecimal = true;
            continue;
        }

        break;
    }

    if (! sawDigit)
    {
        return false;
    }

    std::wstring_view unit = text.substr(index);
    unit                   = TrimAscii(unit);

    if (unit.size() >= 2)
    {
        const wchar_t penultimate = unit[unit.size() - 2];
        const wchar_t last        = unit.back();
        if (penultimate == L'/' && (last == L's' || last == L'S'))
        {
            unit.remove_suffix(2);
            unit = TrimAscii(unit);
        }
    }

    uint64_t multiplier = 0;
    if (unit.empty() || EqualsIgnoreAsciiCase(unit, L"kb") || EqualsIgnoreAsciiCase(unit, L"k") || EqualsIgnoreAsciiCase(unit, L"kib"))
    {
        // Bare numeric strings are interpreted as KiB for user-friendliness.
        multiplier = kKiB;
    }
    else if (EqualsIgnoreAsciiCase(unit, L"b"))
    {
        multiplier = 1;
    }
    else if (EqualsIgnoreAsciiCase(unit, L"mb") || EqualsIgnoreAsciiCase(unit, L"m") || EqualsIgnoreAsciiCase(unit, L"mib"))
    {
        multiplier = kMiB;
    }
    else if (EqualsIgnoreAsciiCase(unit, L"gb") || EqualsIgnoreAsciiCase(unit, L"g") || EqualsIgnoreAsciiCase(unit, L"gib"))
    {
        multiplier = kGiB;
    }
    else if (EqualsIgnoreAsciiCase(unit, L"tb") || EqualsIgnoreAsciiCase(unit, L"t") || EqualsIgnoreAsciiCase(unit, L"tib"))
    {
        multiplier = kTiB;
    }
    else if (EqualsIgnoreAsciiCase(unit, L"pb") || EqualsIgnoreAsciiCase(unit, L"p") || EqualsIgnoreAsciiCase(unit, L"pib"))
    {
        multiplier = kPiB;
    }
    else
    {
        return false;
    }

    const double result = number * static_cast<double>(multiplier);
    if (result <= 0.0)
    {
        outBytesPerSecond = 0;
        return true;
    }

    constexpr double maxValue = static_cast<double>(std::numeric_limits<uint64_t>::max());
    if (result >= maxValue)
    {
        outBytesPerSecond = std::numeric_limits<uint64_t>::max();
        return true;
    }

    outBytesPerSecond = static_cast<uint64_t>(result + 0.5);
    return true;
}

class FileOperationsSpeedLimitPromptWindow final
{
public:
    FileOperationsSpeedLimitPromptWindow(HWND ownerWindow, const AppTheme& theme, uint64_t initialLimitBytesPerSecond) noexcept
        : _ownerWindow(ownerWindow && IsWindow(ownerWindow) != FALSE ? ownerWindow : nullptr),
          _theme(theme),
          _initialLimitBytesPerSecond(initialLimitBytesPerSecond)
    {
    }

    FileOperationsSpeedLimitPromptWindow(const FileOperationsSpeedLimitPromptWindow&)            = delete;
    FileOperationsSpeedLimitPromptWindow& operator=(const FileOperationsSpeedLimitPromptWindow&) = delete;
    FileOperationsSpeedLimitPromptWindow(FileOperationsSpeedLimitPromptWindow&&)                 = delete;
    FileOperationsSpeedLimitPromptWindow& operator=(FileOperationsSpeedLimitPromptWindow&&)      = delete;

    [[nodiscard]] std::optional<uint64_t> ShowModal() noexcept
    {
        const HRESULT classHr = EnsureWindowClass();
        if (FAILED(classHr))
        {
            return std::nullopt;
        }

        const DWORD style   = WS_CAPTION | WS_SYSMENU | WS_POPUP | WS_CLIPCHILDREN | WS_CLIPSIBLINGS;
        const DWORD exStyle = WS_EX_DLGMODALFRAME;
        const UINT dpi      = _ownerWindow && IsWindow(_ownerWindow) != FALSE ? GetDpiForWindow(_ownerWindow) : GetDpiForSystem();

        RECT bounds{0, 0, DipsToPixels(440, dpi), DipsToPixels(212, dpi)};
        if (AdjustWindowRectExForDpi(&bounds, style, FALSE, exStyle, dpi) == FALSE)
        {
            return std::nullopt;
        }

        const HWND hwnd = CreateWindowExW(exStyle,
                                          kFileOperationsSpeedLimitPromptClassName,
                                          LoadStringResource(nullptr, IDS_CAPTION_FILEOP_SPEED_LIMIT_CUSTOM).c_str(),
                                          style,
                                          CW_USEDEFAULT,
                                          CW_USEDEFAULT,
                                          bounds.right - bounds.left,
                                          bounds.bottom - bounds.top,
                                          _ownerWindow,
                                          nullptr,
                                          GetModuleHandleW(nullptr),
                                          this);
        if (! hwnd)
        {
            return std::nullopt;
        }

        if (! _hWnd)
        {
            _hWnd.reset(hwnd);
        }

        if (_ownerWindow && IsWindow(_ownerWindow) != FALSE)
        {
            CenterWindowOnOwnerWindow(_hWnd.get(), _ownerWindow);
        }

        static_cast<void>(_dxHost.PrimeForShow());
        ShowWindow(_hWnd.get(), SW_SHOWNORMAL);
        UpdateWindow(_hWnd.get());
        SetForegroundWindow(_hWnd.get());

        MSG msg{};
        while (! _done)
        {
            const BOOL getMessageResult = GetMessageW(&msg, nullptr, 0, 0);
            if (getMessageResult == -1)
            {
                _result.reset();
                _done = true;
                break;
            }
            if (getMessageResult == 0)
            {
                _result.reset();
                _done = true;
                PostQuitMessage(static_cast<int>(msg.wParam));
                break;
            }

            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }

        return _result;
    }

    static LRESULT CALLBACK WndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
    {
        auto* self = reinterpret_cast<FileOperationsSpeedLimitPromptWindow*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
        if (message == WM_NCCREATE)
        {
            const auto* create = reinterpret_cast<const CREATESTRUCTW*>(lParam);
            self               = create ? reinterpret_cast<FileOperationsSpeedLimitPromptWindow*>(create->lpCreateParams) : nullptr;
            if (! self)
            {
                return FALSE;
            }

            self->_hWnd.reset(hwnd);
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
        }

        if (! self)
        {
            return DefWindowProcW(hwnd, message, wParam, lParam);
        }

        return self->WindowProc(hwnd, message, wParam, lParam);
    }

#ifdef ENABLE_TESTS
public:
    enum class DebugCommand : uintptr_t
    {
        GetSnapshot,
        SetText,
        Confirm,
        Cancel,
    };
#endif

private:
    [[nodiscard]] static HRESULT EnsureWindowClass() noexcept
    {
        static ATOM atom = 0;
        if (atom != 0)
        {
            return S_OK;
        }

        WNDCLASSEXW wc{};
        wc.cbSize        = sizeof(wc);
        wc.lpfnWndProc   = FileOperationsSpeedLimitPromptWindow::WndProc;
        wc.hInstance     = GetModuleHandleW(nullptr);
        wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
        wc.lpszClassName = kFileOperationsSpeedLimitPromptClassName;
        wc.style         = CS_DBLCLKS;

        atom = RegisterClassExW(&wc);
        if (atom != 0)
        {
            return S_OK;
        }

        const DWORD lastError = GetLastError();
        if (lastError == ERROR_CLASS_ALREADY_EXISTS)
        {
            atom = 1;
            return S_OK;
        }

        return HRESULT_FROM_WIN32(lastError);
    }

    [[nodiscard]] bool OnCreate(HWND hwnd) noexcept
    {
        if (! _dxHost.Attach(hwnd))
        {
            return false;
        }

        BuildUi();
        ApplyTheme();
        Layout();
        _dxHost.SetFocusControl(_field);
        _dxHost.SetDefaultButton(_okButton);
        _dxHost.SetCancelButton(_cancelButton);
        return true;
    }

    void BuildUi()
    {
        if (_root != nullptr)
        {
            return;
        }

        using namespace RedSalamander::DxUi;

        _rootStorage = std::make_unique<Panel>();
        _root        = _rootStorage.get();

        _label = _root->AddChild<Label>(LoadStringResource(nullptr, IDS_FILEOP_SPEED_LIMIT_PROMPT_LABEL));
        _label->SetAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);

        _field = _root->AddChild<TextField>(_initialLimitBytesPerSecond == 0 ? L"0" : FormatBytesCompact(_initialLimitBytesPerSecond));
        _field->SetOnTextChanged([this](std::wstring_view) { ClearValidation(); });

        _hintLabel = _root->AddChild<Label>(LoadStringResource(nullptr, IDS_FILEOP_SPEED_LIMIT_PROMPT_HINT));
        _hintLabel->SetMultiline(true);
        _hintLabel->SetAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);
        _hintLabel->SetFontRole(FontRole::Small);

        _okButton = _root->AddChild<Button>(LoadStringResource(nullptr, IDS_BTN_OK));
        _okButton->SetPrimary(true);
        _okButton->SetOnClick([this] { Confirm(); });

        _cancelButton = _root->AddChild<Button>(LoadStringResource(nullptr, IDS_BTN_CANCEL));
        _cancelButton->SetOnClick([this] { Cancel(); });

        _dxHost.SetRoot(std::move(_rootStorage));
    }

    void ApplyTheme() noexcept
    {
        _palette = MakeAppThemeDxPalette(_theme, _theme.windowBackground);
        _dxHost.SetTheme(_palette);
        UpdateHintVisuals();
        if (_hWnd)
        {
            ApplyWindowChromeTheme(_hWnd.get(), _theme, WindowBackdropTarget::Tool, GetActiveWindow() == _hWnd.get());
        }
    }

    void Layout() noexcept
    {
        if (! _root)
        {
            return;
        }

        const D2D1_RECT_F client = _dxHost.GetClientBoundsDip();
        _root->SetBounds(client);

        constexpr float kMarginDip       = 16.0f;
        constexpr float kGapDip          = 10.0f;
        constexpr float kFieldHeightDip  = 32.0f;
        constexpr float kButtonHeightDip = 34.0f;
        constexpr float kButtonWidthDip  = 96.0f;
        constexpr float kLabelHeightDip  = 24.0f;
        constexpr float kHintHeightDip   = 46.0f;

        const float left  = client.left + kMarginDip;
        const float right = client.right - kMarginDip;
        float y           = client.top + kMarginDip;

        _label->SetBounds(D2D1::RectF(left, y, right, y + kLabelHeightDip));
        y += kLabelHeightDip + 4.0f;

        _field->SetBounds(D2D1::RectF(left, y, right, y + kFieldHeightDip));
        y += kFieldHeightDip + kGapDip;

        _hintLabel->SetBounds(D2D1::RectF(left, y, right, y + kHintHeightDip));

        const float buttonsTop = client.bottom - kMarginDip - kButtonHeightDip;
        const float cancelLeft = right - kButtonWidthDip;
        const float okLeft     = cancelLeft - 8.0f - kButtonWidthDip;
        _okButton->SetBounds(D2D1::RectF(okLeft, buttonsTop, okLeft + kButtonWidthDip, buttonsTop + kButtonHeightDip));
        _cancelButton->SetBounds(D2D1::RectF(cancelLeft, buttonsTop, cancelLeft + kButtonWidthDip, buttonsTop + kButtonHeightDip));
    }

    void ClearValidation() noexcept
    {
        if (! _showingValidationError)
        {
            return;
        }

        _showingValidationError = false;
        _validationText.clear();
        UpdateHintVisuals();
        _dxHost.Invalidate();
    }

    void ShowValidation(UINT messageId) noexcept
    {
        _showingValidationError = true;
        _validationText         = LoadStringResource(nullptr, messageId);
        UpdateHintVisuals();
        _dxHost.Invalidate();
        MessageBeep(MB_ICONWARNING);
    }

    void UpdateHintVisuals() noexcept
    {
        if (! _hintLabel)
        {
            return;
        }

        _hintLabel->SetText(_showingValidationError ? _validationText : LoadStringResource(nullptr, IDS_FILEOP_SPEED_LIMIT_PROMPT_HINT));
        _hintLabel->SetTextColor(_showingValidationError ? _palette.errorText : _palette.disabledText);
    }

    void Confirm() noexcept
    {
        ClearValidation();

        const std::wstring text = _field ? std::wstring(_field->GetText()) : std::wstring{};
        uint64_t parsed         = 0;
        if (! TryParseThroughputText(text, parsed))
        {
            ShowValidation(IDS_MSG_FILEOP_SPEED_LIMIT_INVALID);
            _dxHost.SetFocusControl(_field);
            return;
        }

        _result = parsed;
        _done   = true;
        if (_hWnd && IsWindow(_hWnd.get()) != FALSE)
        {
            _hWnd.reset();
        }
    }

    void Cancel() noexcept
    {
        _result.reset();
        _done = true;
        if (_hWnd && IsWindow(_hWnd.get()) != FALSE)
        {
            _hWnd.reset();
        }
    }

#ifdef ENABLE_TESTS
    LRESULT OnDebugCommand(DebugCommand command, LPARAM payload) noexcept
    {
        switch (command)
        {
            case DebugCommand::GetSnapshot:
            {
                auto* snapshot = reinterpret_cast<FileOperationsSpeedLimitPromptDebugSnapshot*>(payload);
                if (! snapshot)
                {
                    return FALSE;
                }

                snapshot->usesDxUiHost            = _dxHost.GetRoot() == _root;
                snapshot->visibleChildWindowCount = 0u;
                if (_hWnd && IsWindow(_hWnd.get()) != FALSE)
                {
                    EnumChildWindows(_hWnd.get(),
                                     [](HWND child, LPARAM cookie) noexcept -> BOOL
                    {
                        if (! IsActuallyVisibleChildWindow(child))
                        {
                            return TRUE;
                        }

                        auto* count = reinterpret_cast<size_t*>(cookie);
                        if (count)
                        {
                            *count += 1u;
                        }
                        return TRUE;
                    },
                                     reinterpret_cast<LPARAM>(&snapshot->visibleChildWindowCount));
                }
                snapshot->initialLimitBytesPerSecond = _initialLimitBytesPerSecond;
                snapshot->text                       = _field ? std::wstring(_field->GetText()) : std::wstring{};
                snapshot->hintText       = _showingValidationError ? std::wstring{} : LoadStringResource(nullptr, IDS_FILEOP_SPEED_LIMIT_PROMPT_HINT);
                snapshot->validationText = _validationText;
                return TRUE;
            }
            case DebugCommand::SetText:
            {
                auto* text = reinterpret_cast<const std::wstring*>(payload);
                if (! text || ! _field)
                {
                    return FALSE;
                }

                _field->SetTextAndNotify(*text);
                ClearValidation();
                _dxHost.SetFocusControl(_field);
                _dxHost.SyncTextInputBridge(_field);
                _dxHost.Invalidate();
                return TRUE;
            }
            case DebugCommand::Confirm: Confirm(); return TRUE;
            case DebugCommand::Cancel: Cancel(); return TRUE;
        }

        return FALSE;
    }
#endif

    LRESULT WindowProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
    {
        bool dxHandled   = false;
        LRESULT dxResult = 0;
        if (message != WM_CREATE)
        {
            dxResult = _dxHost.HandleMessage(hwnd, message, wParam, lParam, dxHandled);
        }
        if (dxHandled)
        {
            if (message == WM_SIZE || message == WM_DPICHANGED)
            {
                Layout();
            }
            if (message == WM_NCDESTROY)
            {
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
                _dxHost.Detach();
                if (_hWnd.get() == hwnd)
                {
                    static_cast<void>(_hWnd.release());
                }
                if (! _done)
                {
                    _result.reset();
                    _done = true;
                }
            }
            return dxResult;
        }

        switch (message)
        {
            case WM_CREATE: return OnCreate(hwnd) ? 0 : -1;
            case WM_SIZE: Layout(); return 0;
            case WM_DPICHANGED:
            {
                if (const auto* suggested = reinterpret_cast<const RECT*>(lParam))
                {
                    SetWindowPos(hwnd,
                                 nullptr,
                                 suggested->left,
                                 suggested->top,
                                 suggested->right - suggested->left,
                                 suggested->bottom - suggested->top,
                                 SWP_NOZORDER | SWP_NOACTIVATE);
                }
                ApplyTheme();
                Layout();
                return 0;
            }
            case WM_NCACTIVATE: ApplyTitleBarTheme(hwnd, _theme, wParam != FALSE); return DefWindowProcW(hwnd, message, wParam, lParam);
            case WM_ERASEBKGND: return 1;
            case WM_CLOSE: Cancel(); return 0;
#ifdef ENABLE_TESTS
            case kFileOperationsSpeedLimitPromptDebugMessage: return OnDebugCommand(static_cast<DebugCommand>(wParam), lParam);
#endif
            case WM_NCDESTROY:
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
                _dxHost.Detach();
                if (_hWnd.get() == hwnd)
                {
                    static_cast<void>(_hWnd.release());
                }
                if (! _done)
                {
                    _result.reset();
                    _done = true;
                }
                return 0;
        }

        return DefWindowProcW(hwnd, message, wParam, lParam);
    }

    HWND _ownerWindow = nullptr;
    AppTheme _theme{};
    uint64_t _initialLimitBytesPerSecond = 0;
    wil::unique_hwnd _hWnd;
    RedSalamander::DxUi::WindowHost _dxHost;
    std::unique_ptr<RedSalamander::DxUi::Panel> _rootStorage;
    RedSalamander::DxUi::Panel* _root = nullptr;
    RedSalamander::DxUi::ThemePalette _palette{};
    RedSalamander::DxUi::Label* _label         = nullptr;
    RedSalamander::DxUi::TextField* _field     = nullptr;
    RedSalamander::DxUi::Label* _hintLabel     = nullptr;
    RedSalamander::DxUi::Button* _okButton     = nullptr;
    RedSalamander::DxUi::Button* _cancelButton = nullptr;
    std::wstring _validationText;
    bool _showingValidationError = false;
    bool _done                   = false;
    std::optional<uint64_t> _result;
};

[[nodiscard]] std::optional<uint64_t> ShowCustomSpeedLimitPrompt(HWND ownerWindow, const AppTheme& theme, uint64_t initialLimitBytesPerSecond) noexcept
{
    FileOperationsSpeedLimitPromptWindow window(ownerWindow, theme, initialLimitBytesPerSecond);
    return window.ShowModal();
}

bool PointInRectF(const D2D1_RECT_F& rc, float x, float y) noexcept
{
    return rc.left <= x && x <= rc.right && rc.top <= y && y <= rc.bottom;
}

float MeasureTextWidth(IDWriteFactory* factory, IDWriteTextFormat* format, std::wstring_view text, float maxWidth, float height) noexcept
{
    if (! factory || ! format || text.empty())
    {
        return 0.0f;
    }

    wil::com_ptr<IDWriteTextLayout> layout;
    const HRESULT hr = factory->CreateTextLayout(text.data(), static_cast<UINT32>(text.size()), format, maxWidth, height, layout.addressof());
    if (FAILED(hr) || ! layout)
    {
        return 0.0f;
    }

    DWRITE_TEXT_METRICS metrics{};
    if (FAILED(layout->GetMetrics(&metrics)))
    {
        return 0.0f;
    }

    return metrics.width;
}

std::wstring TruncateTextMiddleToWidth(IDWriteFactory* factory,
                                       IDWriteTextFormat* format,
                                       std::wstring_view text,
                                       float maxWidth,
                                       float height,
                                       std::wstring_view ellipsisText,
                                       size_t fixedPrefixChars,
                                       size_t minSuffixChars) noexcept
{
    const float fullWidth = MeasureTextWidth(factory, format, text, maxWidth, height);
    if (fullWidth <= maxWidth)
    {
        return std::wstring(text);
    }

    const float dotsWidth = MeasureTextWidth(factory, format, ellipsisText, maxWidth, height);
    if (dotsWidth <= 0.0f || maxWidth <= dotsWidth)
    {
        return std::wstring(ellipsisText);
    }

    fixedPrefixChars = std::min(fixedPrefixChars, text.size());
    minSuffixChars   = std::min(minSuffixChars, text.size());

    if (fixedPrefixChars + minSuffixChars > text.size())
    {
        const size_t overlap = fixedPrefixChars + minSuffixChars - text.size();
        const size_t reduce  = std::min(overlap, fixedPrefixChars);
        fixedPrefixChars -= reduce;
    }

    const std::wstring_view prefix = text.substr(0, fixedPrefixChars);

    const float prefixWidth = MeasureTextWidth(factory, format, prefix, maxWidth, height);
    if (prefixWidth + dotsWidth >= maxWidth)
    {
        return std::wstring(ellipsisText);
    }

    size_t low  = minSuffixChars;
    size_t high = text.size() - fixedPrefixChars;

    while (low < high)
    {
        const size_t mid               = (low + high + 1u) / 2u;
        const std::wstring_view suffix = text.substr(text.size() - mid);

        std::wstring candidate;
        candidate.reserve(prefix.size() + ellipsisText.size() + suffix.size());
        candidate.append(prefix);
        candidate.append(ellipsisText);
        candidate.append(suffix);

        const float w = MeasureTextWidth(factory, format, candidate, maxWidth, height);
        if (w <= maxWidth)
        {
            low = mid;
        }
        else
        {
            high = mid - 1u;
        }
    }

    const std::wstring_view suffix = text.substr(text.size() - low);
    std::wstring result;
    result.reserve(prefix.size() + ellipsisText.size() + suffix.size());
    result.append(prefix);
    result.append(ellipsisText);
    result.append(suffix);
    return result;
}

size_t ComputePathFixedPrefixChars(std::wstring_view path) noexcept
{
    if (path.size() >= 3 && path[1] == L':' && (path[2] == L'\\' || path[2] == L'/'))
    {
        return 3u;
    }

    if (! path.empty() && (path.front() == L'\\' || path.front() == L'/'))
    {
        return 1u;
    }

    return 0u;
}

size_t ComputePathLeafChars(std::wstring_view path) noexcept
{
    std::wstring_view trimmed = path;
    while (! trimmed.empty())
    {
        const wchar_t last = trimmed.back();
        if (last != L'\\' && last != L'/')
        {
            break;
        }
        trimmed.remove_suffix(1);
    }

    const size_t pos = trimmed.find_last_of(L"\\/");
    if (pos == std::wstring_view::npos)
    {
        return trimmed.size();
    }

    if (pos + 1u >= trimmed.size())
    {
        return 0u;
    }

    return trimmed.size() - (pos + 1u);
}

D2D1::ColorF RainbowProgressColor(const AppTheme& theme, std::wstring_view seed) noexcept
{
    if (seed.empty())
    {
        return theme.navigationView.accent;
    }

    const uint32_t hash = StableHash32(seed);
    const float hue     = static_cast<float>(hash % 360u);
    const float sat     = 0.85f;
    const float val     = theme.dark ? 0.80f : 0.90f;
    return ColorFromHSV(hue, sat, val, 1.0f);
}

float RateSampleHue(std::wstring_view sourcePath) noexcept
{
    if (sourcePath.empty())
    {
        return -1.0f;
    }

    const uint32_t pathHash = StableHash32(sourcePath);
    return static_cast<float>(pathHash % 360u);
}

bool IsRateSamplingBlocked(const FileOperationsPopupInternal::RateSnapshot& task) noexcept
{
    return task.paused || task.queuePaused || task.waitingInQueue;
}

[[nodiscard]] double ClampFiniteNonNegative(double value) noexcept
{
    return std::isfinite(value) && value > 0.0 ? value : 0.0;
}

[[nodiscard]] double SmoothFileOperationValue(double previousValue, double sampleValue, ULONGLONG elapsedMs, double tauMs, double maxAlpha) noexcept
{
    previousValue = ClampFiniteNonNegative(previousValue);
    sampleValue   = ClampFiniteNonNegative(sampleValue);
    if (previousValue <= 0.0 || elapsedMs == 0)
    {
        return sampleValue;
    }

    const double elapsed = static_cast<double>(elapsedMs);
    const double alpha   = std::clamp(1.0 - std::exp(-elapsed / tauMs), 0.02, maxAlpha);
    return previousValue + (sampleValue - previousValue) * alpha;
}

[[nodiscard]] double SmoothRateForDisplay(double previousRate, double sampleRate, ULONGLONG elapsedMs) noexcept
{
    constexpr double kRateSmoothingTauMs = 1200.0;
    constexpr double kRateMaxAlpha       = 0.35;
    return SmoothFileOperationValue(previousRate, sampleRate, elapsedMs, kRateSmoothingTauMs, kRateMaxAlpha);
}

[[nodiscard]] double SmoothEtaSecondsForDisplay(double previousEtaSeconds, double sampleEtaSeconds, ULONGLONG elapsedMs) noexcept
{
    constexpr double kEtaSmoothingTauMs = 1500.0;
    constexpr double kEtaMaxAlpha       = 0.45;
    return SmoothFileOperationValue(previousEtaSeconds, sampleEtaSeconds, elapsedMs, kEtaSmoothingTauMs, kEtaMaxAlpha);
}

[[nodiscard]] double DecayRateForCallbackSilence(double smoothedRate, ULONGLONG silenceMs) noexcept
{
    constexpr ULONGLONG kRateSilenceHoldMs = 600ull;
    constexpr double kRateSilenceDecayMs   = 900.0;

    smoothedRate = ClampFiniteNonNegative(smoothedRate);
    if (smoothedRate <= 0.0 || silenceMs <= kRateSilenceHoldMs)
    {
        return smoothedRate;
    }

    const double decayMs = static_cast<double>(silenceMs - kRateSilenceHoldMs);
    return smoothedRate * std::exp(-decayMs / kRateSilenceDecayMs);
}

void AppendRateSample(FileOperationsPopupInternal::RateHistory& history, double sample, float hue) noexcept
{
    const double clampedSample          = ClampFiniteNonNegative(sample);
    history.samples[history.writeIndex] = static_cast<float>(std::min<double>(clampedSample, std::numeric_limits<float>::max()));
    history.hues[history.writeIndex]    = hue;
    history.writeIndex                  = (history.writeIndex + 1u) % FileOperationsPopupInternal::RateHistory::kMaxSamples;
    history.count                       = std::min(FileOperationsPopupInternal::RateHistory::kMaxSamples, history.count + 1u);
}

void ResetPendingRateSample(FileOperationsPopupInternal::RateHistory& history) noexcept
{
    history.pendingBucketMs         = 0;
    history.pendingWeightedSampleMs = 0.0;
    history.pendingHue              = -1.0f;
}

void AppendResampledRateSamples(FileOperationsPopupInternal::RateHistory& history, ULONGLONG elapsedMs, double sample, float hue) noexcept
{
    const ULONGLONG maxResampleMs = static_cast<ULONGLONG>(FileOperationsPopupInternal::RateHistory::kMaxSamples) * kRateSampleBucketMs;
    ULONGLONG remainingMs         = (std::min)(elapsedMs, maxResampleMs);
    while (remainingMs > 0)
    {
        const ULONGLONG bucketRemainingMs = kRateSampleBucketMs - history.pendingBucketMs;
        const ULONGLONG sliceMs           = std::min(remainingMs, bucketRemainingMs);

        history.pendingWeightedSampleMs += static_cast<double>(sample) * static_cast<double>(sliceMs);
        history.pendingBucketMs += sliceMs;
        if (hue >= 0.0f)
        {
            history.pendingHue = hue;
        }

        remainingMs -= sliceMs;
        if (history.pendingBucketMs < kRateSampleBucketMs)
        {
            continue;
        }

        const float bucketHue     = history.pendingHue >= 0.0f ? history.pendingHue : hue;
        const double bucketSample = history.pendingWeightedSampleMs / static_cast<double>(kRateSampleBucketMs);
        AppendRateSample(history, bucketSample, bucketHue);
        ResetPendingRateSample(history);
    }
}

std::wstring TruncatePathMiddleToWidth(IDWriteFactory* factory, IDWriteTextFormat* format, std::wstring_view path, float maxWidth, float height) noexcept
{
    const size_t prefixChars = ComputePathFixedPrefixChars(path);
    const size_t leafChars   = ComputePathLeafChars(path);

    return TruncateTextMiddleToWidth(factory, format, path, maxWidth, height, kEllipsisText, prefixChars, leafChars);
}

ATOM RegisterFileOperationsPopupWndClass(HINSTANCE instance)
{
    static ATOM atom = 0;
    if (atom)
    {
        return atom;
    }

    WNDCLASSEXW wc{};
    wc.cbSize        = sizeof(wc);
    wc.style         = CS_HREDRAW | CS_VREDRAW | CS_DBLCLKS;
    wc.lpfnWndProc   = FileOperationsPopupInternal::FileOperationsPopupState::WndProcThunk;
    wc.hInstance     = instance;
    wc.hCursor       = LoadCursor(nullptr, IDC_ARROW);
    wc.hIcon         = LoadIconW(instance, MAKEINTRESOURCEW(IDI_REDSALAMANDER));
    wc.hIconSm       = LoadIconW(instance, MAKEINTRESOURCEW(IDI_SMALL));
    wc.hbrBackground = nullptr;
    wc.lpszClassName = kFileOperationsPopupClassName;

    atom = RegisterClassExW(&wc);
    return atom;
}

} // namespace

using FileOperationsPopupInternal::PopupButton;
using FileOperationsPopupInternal::PopupHitTest;
using FileOperationsPopupInternal::RateHistory;
using FileOperationsPopupInternal::RateSnapshot;
using FileOperationsPopupInternal::TaskSnapshot;

void FileOperationsPopupInternal::FileOperationsPopupState::ApplyScrollBarTheme(HWND hwnd) const noexcept
{
    if (! hwnd)
    {
        return;
    }

    if (! hostLifetime.lock())
    {
        return;
    }

    if (! folderWindow)
    {
        return;
    }

    const AppTheme& theme = folderWindow->GetTheme();
    if (theme.highContrast)
    {
        SetWindowTheme(hwnd, L"", nullptr);
        return;
    }

    if (theme.dark)
    {
        SetWindowTheme(hwnd, L"DarkMode_Explorer", nullptr);
        return;
    }

    SetWindowTheme(hwnd, L"Explorer", nullptr);
}

bool FileOperationsPopupInternal::FileOperationsPopupState::IsTaskCollapsed(uint64_t taskId) const noexcept
{
    const auto it = _collapsedTasks.find(taskId);
    if (it == _collapsedTasks.end())
    {
        return false;
    }

    return it->second;
}

void FileOperationsPopupInternal::FileOperationsPopupState::ToggleTaskCollapsed(uint64_t taskId) noexcept
{
    const bool next         = ! IsTaskCollapsed(taskId);
    _collapsedTasks[taskId] = next;
}

void FileOperationsPopupInternal::FileOperationsPopupState::CleanupCollapsedTasks(const std::vector<TaskSnapshot>& snapshot) noexcept
{
    std::unordered_map<uint64_t, bool> seen;
    seen.reserve(snapshot.size());
    for (const TaskSnapshot& task : snapshot)
    {
        seen[task.taskId] = true;
    }

    for (auto it = _collapsedTasks.begin(); it != _collapsedTasks.end();)
    {
        if (seen.find(it->first) == seen.end())
        {
            it = _collapsedTasks.erase(it);
            continue;
        }
        ++it;
    }
}

void FileOperationsPopupInternal::FileOperationsPopupState::DiscardDeviceResources() noexcept
{
    _target.reset();
    _captionGlyphTarget.reset();

    _bgBrush.reset();
    _textBrush.reset();
    _subTextBrush.reset();
    _borderBrush.reset();
    _progressBgBrush.reset();
    _progressGlobalBrush.reset();
    _progressItemBrush.reset();
    _checkboxFillBrush.reset();
    _checkboxCheckBrush.reset();
    _statusOkBrush.reset();
    _statusWarningBrush.reset();
    _statusErrorBrush.reset();
    _graphBgBrush.reset();
    _graphGridBrush.reset();
    _graphLimitBrush.reset();
    _graphLineBrush.reset();
    _graphFillBrush.reset();
    _graphDynamicBrush.reset();
    _graphTextShadowBrush.reset();
    _buttonBgBrush.reset();
    _buttonHoverBrush.reset();
    _buttonPressedBrush.reset();
    _captionGlyphBrush.reset();
}

void FileOperationsPopupInternal::FileOperationsPopupState::EnsureFactories() noexcept
{
    if (! _d2dFactory)
    {
        const HRESULT hr = D2D1CreateFactory(D2D1_FACTORY_TYPE_SINGLE_THREADED, _d2dFactory.addressof());
        if (FAILED(hr))
        {
            _d2dFactory.reset();
        }
    }

    if (! _dwriteFactory)
    {
        const HRESULT hr = DWriteCreateFactory(DWRITE_FACTORY_TYPE_SHARED, __uuidof(IDWriteFactory), reinterpret_cast<IUnknown**>(_dwriteFactory.addressof()));
        if (FAILED(hr))
        {
            _dwriteFactory.reset();
        }
    }
}

void FileOperationsPopupInternal::FileOperationsPopupState::EnsureTextFormats() noexcept
{
    if (! _dwriteFactory)
    {
        return;
    }

    if (_headerFormat && _bodyFormat && _smallFormat && _buttonFormat && _buttonSmallFormat && _graphOverlayFormat && _statusIconFallbackFormat)
    {
        return;
    }

    if (! _headerFormat)
    {
        static_cast<void>(RedSalamander::DxUi::Typography::CreateTextFormat(
            _dwriteFactory.get(),
            RedSalamander::DxUi::Typography::MakeUiTextSpec(DipsToPixels(12.0f, _dpi), DWRITE_FONT_WEIGHT_SEMI_BOLD),
            _headerFormat.put(),
            L""));
    }

    if (! _bodyFormat)
    {
        static_cast<void>(RedSalamander::DxUi::Typography::CreateTextFormat(
            _dwriteFactory.get(), RedSalamander::DxUi::Typography::MakeUiTextSpec(DipsToPixels(12.0f, _dpi)), _bodyFormat.put(), L""));
    }

    if (! _smallFormat)
    {
        static_cast<void>(RedSalamander::DxUi::Typography::CreateTextFormat(
            _dwriteFactory.get(), RedSalamander::DxUi::Typography::MakeUiTextSpec(DipsToPixels(11.0f, _dpi)), _smallFormat.put(), L""));
    }

    if (! _buttonFormat)
    {
        static_cast<void>(RedSalamander::DxUi::Typography::CreateTextFormat(
            _dwriteFactory.get(), RedSalamander::DxUi::Typography::MakeUiTextSpec(DipsToPixels(12.0f, _dpi)), _buttonFormat.put(), L""));
    }

    if (! _buttonSmallFormat)
    {
        static_cast<void>(RedSalamander::DxUi::Typography::CreateTextFormat(
            _dwriteFactory.get(), RedSalamander::DxUi::Typography::MakeUiTextSpec(DipsToPixels(11.0f, _dpi)), _buttonSmallFormat.put(), L""));
    }

    if (! _graphOverlayFormat)
    {
        static_cast<void>(RedSalamander::DxUi::Typography::CreateTextFormat(
            _dwriteFactory.get(),
            RedSalamander::DxUi::Typography::MakeUiTextSpec(DipsToPixels(14.0f, _dpi), DWRITE_FONT_WEIGHT_SEMI_BOLD),
            _graphOverlayFormat.put(),
            L""));
    }

    if (! _statusIconFallbackFormat)
    {
        static_cast<void>(RedSalamander::DxUi::Typography::CreateTextFormat(
            _dwriteFactory.get(),
            RedSalamander::DxUi::Typography::MakeUiTextSpec(DipsToPixels(14.0f, _dpi), DWRITE_FONT_WEIGHT_SEMI_BOLD),
            _statusIconFallbackFormat.put(),
            L""));
    }

    if (! _statusIconFormat)
    {
        // Optional: Segoe Fluent Icons. If missing, fallback format draws standard Unicode glyphs.
        const HRESULT hr = RedSalamander::DxUi::Typography::CreateTextFormat(
            _dwriteFactory.get(), RedSalamander::DxUi::Typography::MakeUiIconSpec(DipsToPixels(14.0f, _dpi)), _statusIconFormat.put(), L"");
        if (FAILED(hr))
        {
            _statusIconFormat.reset();
        }
    }

    auto configureLineFormat = [](IDWriteTextFormat* format) noexcept
    {
        if (! format)
        {
            return;
        }
        format->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP);
        format->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);
        format->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
    };

    auto configureButtonFormat = [](IDWriteTextFormat* format) noexcept
    {
        if (! format)
        {
            return;
        }
        format->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP);
        format->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_CENTER);
        format->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
    };

    configureLineFormat(_headerFormat.get());
    configureLineFormat(_bodyFormat.get());
    configureLineFormat(_smallFormat.get());
    configureButtonFormat(_buttonFormat.get());
    configureButtonFormat(_buttonSmallFormat.get());
    configureButtonFormat(_graphOverlayFormat.get());

    configureButtonFormat(_statusIconFormat.get());
    configureButtonFormat(_statusIconFallbackFormat.get());
}

void FileOperationsPopupInternal::FileOperationsPopupState::EnsureCaptionGlyphTextFormats(UINT dpi) noexcept
{
    if (! _dwriteFactory)
    {
        return;
    }

    if (_captionGlyphDpi != dpi)
    {
        _captionGlyphDpi = dpi;
        _captionGlyphFormat.reset();
        _captionGlyphFallbackFormat.reset();
    }

    if (! _captionGlyphFallbackFormat)
    {
        const HRESULT hr = RedSalamander::DxUi::Typography::CreateTextFormat(
            _dwriteFactory.get(), RedSalamander::DxUi::Typography::MakeUiTextSpec(20.0f, DWRITE_FONT_WEIGHT_SEMI_BOLD), _captionGlyphFallbackFormat.put(), L"");
        if (FAILED(hr))
        {
            _captionGlyphFallbackFormat.reset();
        }
    }

    if (! _captionGlyphFormat)
    {
        const HRESULT hr = RedSalamander::DxUi::Typography::CreateTextFormat(
            _dwriteFactory.get(), RedSalamander::DxUi::Typography::MakeUiIconSpec(20.0f), _captionGlyphFormat.put(), L"");
        if (FAILED(hr))
        {
            _captionGlyphFormat.reset();
        }
    }

    auto configure = [](IDWriteTextFormat* format) noexcept
    {
        if (! format)
        {
            return;
        }
        format->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP);
        format->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_CENTER);
        format->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
    };

    configure(_captionGlyphFormat.get());
    configure(_captionGlyphFallbackFormat.get());
}

void FileOperationsPopupInternal::FileOperationsPopupState::EnsureTarget(HWND hwnd) noexcept
{
    EnsureFactories();
    if (! _d2dFactory)
    {
        return;
    }

    if (_target)
    {
        return;
    }

    RECT rc{};
    GetClientRect(hwnd, &rc);
    _clientSize.cx = std::max(0L, rc.right - rc.left);
    _clientSize.cy = std::max(0L, rc.bottom - rc.top);

    _dpi = GetDpiForWindow(hwnd);

    const D2D1_SIZE_U size                             = D2D1::SizeU(static_cast<UINT32>(_clientSize.cx), static_cast<UINT32>(_clientSize.cy));
    D2D1_RENDER_TARGET_PROPERTIES props                = D2D1::RenderTargetProperties();
    props.dpiX                                         = 96.0f;
    props.dpiY                                         = 96.0f;
    const D2D1_HWND_RENDER_TARGET_PROPERTIES hwndProps = D2D1::HwndRenderTargetProperties(hwnd, size);

    wil::com_ptr<ID2D1HwndRenderTarget> rt;
    const HRESULT hr = _d2dFactory->CreateHwndRenderTarget(props, hwndProps, rt.addressof());
    if (FAILED(hr) || ! rt)
    {
        _target.reset();
        return;
    }

    rt->SetTextAntialiasMode(D2D1_TEXT_ANTIALIAS_MODE_CLEARTYPE);
    _target = std::move(rt);
}

bool FileOperationsPopupInternal::FileOperationsPopupState::EnsureCaptionGlyphTarget(UINT dpi) noexcept
{
    EnsureFactories();
    EnsureCaptionGlyphTextFormats(dpi);
    if (! _d2dFactory)
    {
        return false;
    }

    if (! _captionGlyphTarget)
    {
        const D2D1_RENDER_TARGET_PROPERTIES props = D2D1::RenderTargetProperties();
        const HRESULT hr                          = _d2dFactory->CreateDCRenderTarget(&props, _captionGlyphTarget.addressof());
        if (FAILED(hr) || ! _captionGlyphTarget)
        {
            _captionGlyphTarget.reset();
            return false;
        }

        _captionGlyphTarget->SetTextAntialiasMode(D2D1_TEXT_ANTIALIAS_MODE_CLEARTYPE);
    }

    _captionGlyphTarget->SetDpi(static_cast<float>(dpi), static_cast<float>(dpi));

    if (! _captionGlyphBrush)
    {
        const HRESULT hr = _captionGlyphTarget->CreateSolidColorBrush(D2D1::ColorF(0.0f, 0.0f, 0.0f, 1.0f), _captionGlyphBrush.addressof());
        if (FAILED(hr))
        {
            _captionGlyphBrush.reset();
            return false;
        }
    }

    return _captionGlyphTarget && _captionGlyphBrush;
}

void FileOperationsPopupInternal::FileOperationsPopupState::EnsureBrushes() noexcept
{
    if (! _target)
    {
        return;
    }

    if (! hostLifetime.lock())
    {
        return;
    }

    if (! folderWindow)
    {
        return;
    }

    const AppTheme& theme     = folderWindow->GetTheme();
    const D2D1::ColorF bg     = ColorFromCOLORREF(theme.windowBackground);
    const D2D1::ColorF fg     = ColorFromCOLORREF(theme.menu.text);
    const D2D1::ColorF sub    = ColorFromCOLORREF(theme.menu.disabledText);
    const D2D1::ColorF border = ColorFromCOLORREF(theme.menu.border);

    const D2D1::ColorF progressBg     = theme.fileOperations.progressBackground;
    const D2D1::ColorF progressGlobal = theme.fileOperations.progressTotal;
    const D2D1::ColorF progressItem   = theme.fileOperations.progressItem;
    _progressItemBaseColor            = progressItem;

    const D2D1::ColorF okAccent    = theme.accent;
    const D2D1::ColorF warningText = theme.folderView.warningText;
    const D2D1::ColorF errorText   = theme.folderView.errorText;

    const D2D1::ColorF graphBg    = theme.fileOperations.graphBackground;
    const D2D1::ColorF graphGrid  = theme.fileOperations.graphGrid;
    const D2D1::ColorF graphLimit = theme.fileOperations.graphLimit;
    const D2D1::ColorF graphLine  = theme.fileOperations.graphLine;

    if (! _bgBrush)
    {
        _target->CreateSolidColorBrush(bg, _bgBrush.addressof());
    }
    else
    {
        _bgBrush->SetColor(bg);
    }

    if (! _textBrush)
    {
        _target->CreateSolidColorBrush(fg, _textBrush.addressof());
    }
    else
    {
        _textBrush->SetColor(fg);
    }

    if (! _subTextBrush)
    {
        _target->CreateSolidColorBrush(sub, _subTextBrush.addressof());
    }
    else
    {
        _subTextBrush->SetColor(sub);
    }

    if (! _borderBrush)
    {
        _target->CreateSolidColorBrush(border, _borderBrush.addressof());
    }
    else
    {
        _borderBrush->SetColor(border);
    }

    if (! _progressBgBrush)
    {
        _target->CreateSolidColorBrush(progressBg, _progressBgBrush.addressof());
    }
    else
    {
        _progressBgBrush->SetColor(progressBg);
    }

    if (! _progressGlobalBrush)
    {
        _target->CreateSolidColorBrush(progressGlobal, _progressGlobalBrush.addressof());
    }
    else
    {
        _progressGlobalBrush->SetColor(progressGlobal);
    }

    if (! _progressItemBrush)
    {
        _target->CreateSolidColorBrush(progressItem, _progressItemBrush.addressof());
    }
    else
    {
        _progressItemBrush->SetColor(progressItem);
    }

    if (! _statusOkBrush)
    {
        _target->CreateSolidColorBrush(okAccent, _statusOkBrush.addressof());
    }
    else
    {
        _statusOkBrush->SetColor(okAccent);
    }

    if (! _statusWarningBrush)
    {
        _target->CreateSolidColorBrush(warningText, _statusWarningBrush.addressof());
    }
    else
    {
        _statusWarningBrush->SetColor(warningText);
    }

    if (! _statusErrorBrush)
    {
        _target->CreateSolidColorBrush(errorText, _statusErrorBrush.addressof());
    }
    else
    {
        _statusErrorBrush->SetColor(errorText);
    }

    if (! _graphBgBrush)
    {
        _target->CreateSolidColorBrush(graphBg, _graphBgBrush.addressof());
    }
    else
    {
        _graphBgBrush->SetColor(graphBg);
    }

    if (! _graphGridBrush)
    {
        _target->CreateSolidColorBrush(graphGrid, _graphGridBrush.addressof());
    }
    else
    {
        _graphGridBrush->SetColor(graphGrid);
    }

    if (! _graphLimitBrush)
    {
        _target->CreateSolidColorBrush(graphLimit, _graphLimitBrush.addressof());
    }
    else
    {
        _graphLimitBrush->SetColor(graphLimit);
    }

    if (! _graphLineBrush)
    {
        _target->CreateSolidColorBrush(graphLine, _graphLineBrush.addressof());
    }
    else
    {
        _graphLineBrush->SetColor(graphLine);
    }

    const float graphFillAlpha   = theme.dark ? 0.22f : 0.18f;
    const D2D1::ColorF graphFill = D2D1::ColorF(graphLine.r, graphLine.g, graphLine.b, graphFillAlpha);
    _graphFillBaseColor          = graphFill;

    if (! _graphFillBrush)
    {
        _target->CreateSolidColorBrush(graphFill, _graphFillBrush.addressof());
    }
    else
    {
        _graphFillBrush->SetColor(graphFill);
    }

    if (! _graphDynamicBrush)
    {
        _target->CreateSolidColorBrush(graphFill, _graphDynamicBrush.addressof());
    }

    // Shadow brush for overlay text - lighter on light themes for subtlety
    const float shadowAlpha        = theme.dark ? 0.6f : 0.25f;
    const D2D1::ColorF shadowColor = D2D1::ColorF(0.0f, 0.0f, 0.0f, shadowAlpha);
    if (! _graphTextShadowBrush)
    {
        _target->CreateSolidColorBrush(shadowColor, _graphTextShadowBrush.addressof());
    }
    else
    {
        _graphTextShadowBrush->SetColor(shadowColor);
    }

    const D2D1::ColorF btnBg      = ColorFromCOLORREF(theme.menu.background);
    const D2D1::ColorF btnHover   = ColorFromCOLORREF(theme.menu.selectionBg, 0.15f);
    const D2D1::ColorF btnPressed = ColorFromCOLORREF(theme.menu.selectionBg, 0.25f);

    if (! _buttonBgBrush)
    {
        _target->CreateSolidColorBrush(btnBg, _buttonBgBrush.addressof());
    }
    else
    {
        _buttonBgBrush->SetColor(btnBg);
    }

    if (! _buttonHoverBrush)
    {
        _target->CreateSolidColorBrush(btnHover, _buttonHoverBrush.addressof());
    }
    else
    {
        _buttonHoverBrush->SetColor(btnHover);
    }

    if (! _buttonPressedBrush)
    {
        _target->CreateSolidColorBrush(btnPressed, _buttonPressedBrush.addressof());
    }
    else
    {
        _buttonPressedBrush->SetColor(btnPressed);
    }

    const D2D1::ColorF checkboxFill = ColorFromCOLORREF(theme.menu.selectionBg);
    if (! _checkboxFillBrush)
    {
        _target->CreateSolidColorBrush(checkboxFill, _checkboxFillBrush.addressof());
    }
    else
    {
        _checkboxFillBrush->SetColor(checkboxFill);
    }

    const D2D1::ColorF checkMark = ColorFromCOLORREF(theme.menu.selectionText);
    if (! _checkboxCheckBrush)
    {
        _target->CreateSolidColorBrush(checkMark, _checkboxCheckBrush.addressof());
    }
    else
    {
        _checkboxCheckBrush->SetColor(checkMark);
    }
}

std::vector<TaskSnapshot> FileOperationsPopupInternal::FileOperationsPopupState::BuildSnapshot() const
{
    std::vector<FolderWindow::FileOperationState::Task*> tasks;
    std::vector<FolderWindow::InformationalTaskUpdate> informationalTasks;
    std::vector<FolderWindow::FileOperationState::CompletedTaskSummary> completedTasks;
    if (fileOps && hostLifetime.lock())
    {
        fileOps->CollectTasks(tasks);
        fileOps->CollectInformationalTasks(informationalTasks);
        fileOps->CollectCompletedTasks(completedTasks);
    }

    std::vector<TaskSnapshot> result;
    result.reserve(tasks.size() + completedTasks.size() + informationalTasks.size());
    std::unordered_map<uint64_t, bool> activeTaskIds;
    activeTaskIds.reserve(tasks.size());

    for (const auto& info : informationalTasks)
    {
        if (info.taskId == 0)
        {
            continue;
        }

        TaskSnapshot snap{};
        snap.kind                  = TaskSnapshot::Kind::Informational;
        snap.taskId                = info.taskId;
        activeTaskIds[snap.taskId] = true;
        snap.informational         = info;
        snap.started               = true;
        snap.finished              = info.finished;
        snap.resultHr              = info.resultHr;

        result.push_back(std::move(snap));
    }

    for (auto* task : tasks)
    {
        if (! task)
        {
            continue;
        }

        TaskSnapshot snap{};
        snap.taskId                = task->GetId();
        activeTaskIds[snap.taskId] = true;
        snap.operation             = task->GetOperation();

        snap.totalItems         = task->_publishedProgressTotalItems.load(std::memory_order_acquire);
        snap.completedItems     = task->_publishedProgressCompletedItems.load(std::memory_order_acquire);
        snap.totalBytes         = task->_publishedProgressTotalBytes.load(std::memory_order_acquire);
        snap.completedBytes     = task->_publishedProgressCompletedBytes.load(std::memory_order_acquire);
        snap.itemTotalBytes     = task->_publishedProgressItemTotalBytes.load(std::memory_order_acquire);
        snap.itemCompletedBytes = task->_publishedProgressItemCompletedBytes.load(std::memory_order_acquire);
        snap.completedFiles     = task->_publishedCompletedTopLevelFiles.load(std::memory_order_acquire);
        snap.completedFolders   = task->_publishedCompletedTopLevelFolders.load(std::memory_order_acquire);

        {
            std::scoped_lock lock(task->_progressPathMutex);
            snap.currentSourcePath = ! task->_lastProgressCallbackSourcePath.empty() ? task->_lastProgressCallbackSourcePath : task->_progressSourcePath;
            snap.currentDestinationPath =
                ! task->_lastProgressCallbackDestinationPath.empty() ? task->_lastProgressCallbackDestinationPath : task->_progressDestinationPath;
            snap.hasProgressCallbacks     = task->_lastProgressCallbackTick != 0;
            snap.lastProgressCallbackTick = task->_lastProgressCallbackTick;
        }

        {
            std::scoped_lock lock(task->_inFlightFilesMutex);
            snap.inFlightFileCount = std::min(task->_inFlightFileCount, snap.inFlightFiles.size());
            for (size_t i = 0; i < snap.inFlightFileCount; ++i)
            {
                snap.inFlightFiles[i].sourcePath     = task->_inFlightFiles[i].sourcePath;
                snap.inFlightFiles[i].totalBytes     = task->_inFlightFiles[i].totalBytes;
                snap.inFlightFiles[i].completedBytes = task->_inFlightFiles[i].completedBytes;

                // Defensive: for display purposes, avoid showing a misleading "100%" when a plugin reports
                // currentItemCompletedBytes > currentItemTotalBytes (can happen with out-of-order updates or bugs).
                if (snap.inFlightFiles[i].totalBytes > 0 && snap.inFlightFiles[i].completedBytes > snap.inFlightFiles[i].totalBytes)
                {
                    constexpr uint64_t kClampThresholdBytes = 64ull * 1024ull;
                    const uint64_t delta                    = snap.inFlightFiles[i].completedBytes - snap.inFlightFiles[i].totalBytes;
                    if (delta <= kClampThresholdBytes)
                    {
                        snap.inFlightFiles[i].completedBytes = snap.inFlightFiles[i].totalBytes;
                    }
                    else
                    {
                        // Unknown/invalid totals: render as indeterminate.
                        snap.inFlightFiles[i].totalBytes     = 0;
                        snap.inFlightFiles[i].completedBytes = 0;
                    }
                }
                snap.inFlightFiles[i].lastUpdateTick = task->_inFlightFiles[i].lastUpdateTick;
            }
        }

        {
            std::scoped_lock lock(task->_conflictMutex);
            snap.conflict.active            = task->_conflictPrompt.active;
            snap.conflict.bucket            = static_cast<uint8_t>(task->_conflictPrompt.bucket);
            snap.conflict.status            = task->_conflictPrompt.status;
            snap.conflict.sourcePath        = task->_conflictPrompt.sourcePath;
            snap.conflict.destinationPath   = task->_conflictPrompt.destinationPath;
            snap.conflict.applyToAllChecked = task->_conflictPrompt.applyToAllChecked;
            snap.conflict.retryFailed       = task->_conflictPrompt.retryFailed;

            snap.conflict.actionCount = std::min(task->_conflictPrompt.actionCount, snap.conflict.actions.size());
            for (size_t i = 0; i < snap.conflict.actionCount; ++i)
            {
                snap.conflict.actions[i] = static_cast<uint8_t>(task->_conflictPrompt.actions[i]);
            }
        }

        snap.started            = task->HasStarted();
        snap.paused             = task->IsPaused();
        snap.waitingForOthers   = task->IsWaitingForOthers();
        snap.waitingInQueue     = task->IsWaitingInQueue();
        snap.queuePaused        = task->IsQueuePaused();
        snap.plannedItems       = task->GetPlannedItemCount();
        snap.destinationFolder  = task->GetDestinationFolder();
        snap.destinationPane    = task->GetDestinationPane();
        snap.operationStartTick = task->_operationStartTick.load(std::memory_order_acquire);

        snap.desiredSpeedLimitBytesPerSecond       = task->_desiredSpeedLimitBytesPerSecond.load(std::memory_order_acquire);
        snap.effectiveSpeedLimitBytesPerSecond     = task->_effectiveSpeedLimitBytesPerSecond.load(std::memory_order_acquire);
        snap.autoConcurrencyUsed                   = task->_autoConcurrencyUsed.load(std::memory_order_acquire);
        snap.autoConcurrencyStorageKind            = task->_autoConcurrencyStorageKind.load(std::memory_order_acquire);
        snap.autoConcurrencyDestinationStorageKind = task->_autoConcurrencyDestinationStorageKind.load(std::memory_order_acquire);
        snap.autoTunedConcurrency                  = task->_autoTunedConcurrency.load(std::memory_order_acquire);
        snap.effectiveConcurrencyBudget            = task->_effectiveConcurrencyBudget.load(std::memory_order_acquire);

        // Pre-calculation state
        snap.preCalcInProgress     = task->_preCalcInProgress.load(std::memory_order_acquire);
        snap.preCalcSkipped        = task->_preCalcSkipped.load(std::memory_order_acquire);
        snap.preCalcCompleted      = task->_preCalcCompleted.load(std::memory_order_acquire);
        snap.preCalcTotalBytes     = task->_preCalcTotalBytes.load(std::memory_order_acquire);
        snap.preCalcFileCount      = task->_preCalcFileCount.load(std::memory_order_acquire);
        snap.preCalcDirectoryCount = task->_preCalcDirectoryCount.load(std::memory_order_acquire);

        const ULONGLONG startTick = task->_preCalcStartTick.load(std::memory_order_acquire);
        if (snap.preCalcInProgress && startTick > 0)
        {
            const ULONGLONG nowTick = GetTickCount64();
            snap.preCalcElapsedMs   = (nowTick >= startTick) ? (nowTick - startTick) : 0;
        }

        if (snap.totalItems == 0 && snap.operation != FILESYSTEM_DELETE)
        {
            snap.totalItems = snap.plannedItems;
        }

        if (snap.totalItems > 0)
        {
            snap.completedItems = std::min(snap.completedItems, snap.totalItems);
        }
        if (snap.totalBytes > 0)
        {
            snap.completedBytes = std::min(snap.completedBytes, snap.totalBytes);
        }
        if (snap.itemTotalBytes > 0)
        {
            snap.itemCompletedBytes = std::min(snap.itemCompletedBytes, snap.itemTotalBytes);
        }
        NormalizeCompletedTaskSnapshotForDisplay(snap);

        result.push_back(std::move(snap));
    }

    for (const auto& completed : completedTasks)
    {
        if (activeTaskIds.find(completed.taskId) != activeTaskIds.end())
        {
            continue;
        }

        TaskSnapshot snap{};
        snap.taskId                                = completed.taskId;
        snap.operation                             = completed.operation;
        snap.totalItems                            = completed.totalItems;
        snap.completedItems                        = completed.completedItems;
        snap.totalBytes                            = completed.totalBytes;
        snap.completedBytes                        = completed.completedBytes;
        snap.completedFiles                        = completed.completedFiles;
        snap.completedFolders                      = completed.completedFolders;
        snap.currentSourcePath                     = completed.sourcePath;
        snap.currentDestinationPath                = completed.destinationPath;
        snap.destinationFolder                     = completed.destinationFolder;
        snap.destinationPane                       = completed.destinationPane;
        snap.started                               = true;
        snap.finished                              = true;
        snap.autoConcurrencyUsed                   = completed.autoConcurrencyUsed;
        snap.autoConcurrencyStorageKind            = completed.autoConcurrencyStorageKind;
        snap.autoConcurrencyDestinationStorageKind = completed.autoConcurrencyDestinationStorageKind;
        snap.autoTunedConcurrency                  = completed.autoTunedConcurrency;
        snap.effectiveConcurrencyBudget            = completed.effectiveConcurrencyBudget;
        snap.resultHr                              = completed.resultHr;
        snap.warningCount                          = completed.warningCount;
        snap.errorCount                            = completed.errorCount;
        snap.lastDiagnosticMessage                 = completed.lastDiagnosticMessage;
        snap.preCalcSkipped                        = completed.preCalcSkipped;
        snap.hasProgressCallbacks                  = completed.lastProgressCallbackTick != 0;
        snap.lastProgressCallbackTick              = completed.lastProgressCallbackTick;

        if (snap.totalItems > 0)
        {
            snap.completedItems = std::min(snap.completedItems, snap.totalItems);
        }
        if (snap.totalBytes > 0)
        {
            snap.completedBytes = std::min(snap.completedBytes, snap.totalBytes);
        }
        NormalizeCompletedTaskSnapshotForDisplay(snap);

        result.push_back(std::move(snap));
    }

    return result;
}

std::vector<RateSnapshot> FileOperationsPopupInternal::FileOperationsPopupState::BuildRateSnapshot() const
{
    std::vector<FolderWindow::FileOperationState::Task*> tasks;
    std::vector<FolderWindow::FileOperationState::CompletedTaskSummary> completedTasks;
    if (fileOps && hostLifetime.lock())
    {
        fileOps->CollectTasks(tasks);
        fileOps->CollectCompletedTasks(completedTasks);
    }

    std::unordered_map<uint64_t, bool> activeTaskIds;
    activeTaskIds.reserve(tasks.size());

    std::vector<RateSnapshot> result;
    result.reserve(tasks.size() + completedTasks.size());

    for (auto* task : tasks)
    {
        if (! task)
        {
            continue;
        }

        RateSnapshot snap{};
        snap.taskId                = task->GetId();
        snap.operation             = task->GetOperation();
        activeTaskIds[snap.taskId] = true;

        snap.completedItems = task->_publishedProgressCompletedItems.load(std::memory_order_acquire);
        snap.totalBytes     = task->_publishedProgressTotalBytes.load(std::memory_order_acquire);
        snap.completedBytes = task->_publishedProgressCompletedBytes.load(std::memory_order_acquire);

        {
            std::scoped_lock lock(task->_progressPathMutex);
            snap.currentSourcePath        = ! task->_lastProgressCallbackSourcePath.empty() ? task->_lastProgressCallbackSourcePath : task->_progressSourcePath;
            snap.lastProgressCallbackTick = task->_lastProgressCallbackTick;
        }

        snap.progressStateChangeTick = task->_rateSamplingStateChangeTick.load(std::memory_order_acquire);
        snap.started                 = task->HasStarted();
        snap.paused                  = task->IsPaused();
        snap.waitingForOthers        = task->IsWaitingForOthers();
        snap.waitingInQueue          = task->IsWaitingInQueue();
        snap.queuePaused             = task->IsQueuePaused();

        result.push_back(snap);
    }

    for (const auto& completed : completedTasks)
    {
        if (activeTaskIds.find(completed.taskId) != activeTaskIds.end())
        {
            continue;
        }

        RateSnapshot snap{};
        snap.taskId                   = completed.taskId;
        snap.operation                = completed.operation;
        snap.completedItems           = completed.completedItems;
        snap.totalBytes               = completed.totalBytes;
        snap.completedBytes           = completed.completedBytes;
        snap.currentSourcePath        = completed.sourcePath;
        snap.lastProgressCallbackTick = completed.lastProgressCallbackTick;
        snap.started                  = true;
        snap.finished                 = true;

        result.push_back(std::move(snap));
    }

    return result;
}

void FileOperationsPopupInternal::FileOperationsPopupState::UpdateRates() noexcept
{
    const bool capturePerf                   = Debug::Perf::IsCaptureEnabled();
    const uint64_t perfStartUs               = capturePerf ? PerfNowUs() : 0u;
    const ULONGLONG nowTick                  = GetTickCount64();
    const std::vector<RateSnapshot> snapshot = BuildRateSnapshot();

    std::unordered_map<uint64_t, bool> seen;
    seen.reserve(snapshot.size());

    uint64_t maxCallbackSilenceMs = 0;
    uint64_t maxDisplayGapMs      = 0;
    uint64_t silentTaskCount      = 0;
    uint64_t updatedTaskCount     = 0;

    for (const RateSnapshot& task : snapshot)
    {
        seen[task.taskId] = true;

        RateHistory& history = _rates[task.taskId];
        const bool blocked   = IsRateSamplingBlocked(task);
        const bool itemRate  = task.operation == FILESYSTEM_DELETE;

        if (! history.initialized)
        {
            history.initialized              = true;
            history.lastBytes                = task.completedBytes;
            history.lastItems                = task.completedItems;
            history.lastProgressCallbackTick = task.lastProgressCallbackTick;
            history.lastDisplaySampleTick    = nowTick;
            history.lastStateChangeTick      = task.progressStateChangeTick;
            continue;
        }

        history.lastBytes = std::min(history.lastBytes, task.completedBytes);
        history.lastItems = std::min(history.lastItems, task.completedItems);

        if (task.progressStateChangeTick > history.lastStateChangeTick)
        {
            history.lastStateChangeTick   = task.progressStateChangeTick;
            history.resumeTick            = blocked ? 0 : task.progressStateChangeTick;
            history.lastDisplaySampleTick = nowTick;
            ResetPendingRateSample(history);

            if (blocked)
            {
                history.lastBytes                = task.completedBytes;
                history.lastItems                = task.completedItems;
                history.lastProgressCallbackTick = task.lastProgressCallbackTick;
                history.hasSmoothedEta           = false;
                continue;
            }
        }

        if (blocked)
        {
            history.resumeTick               = 0;
            history.lastBytes                = task.completedBytes;
            history.lastItems                = task.completedItems;
            history.lastProgressCallbackTick = task.lastProgressCallbackTick;
            history.lastDisplaySampleTick    = nowTick;
            history.hasSmoothedEta           = false;
            continue;
        }

        const bool hasNewProgressInput = task.lastProgressCallbackTick != 0 && task.lastProgressCallbackTick > history.lastProgressCallbackTick;
        const float hue                = RateSampleHue(task.currentSourcePath);
        ULONGLONG smoothingElapsedMs   = kFileOperationsPopupTimerIntervalMs;

        if (hasNewProgressInput)
        {
            ULONGLONG baselineTick = history.lastProgressCallbackTick;
            if (baselineTick == 0 || baselineTick > task.lastProgressCallbackTick)
            {
                baselineTick = task.lastProgressCallbackTick;
            }
            if (history.resumeTick > baselineTick && history.resumeTick <= task.lastProgressCallbackTick)
            {
                baselineTick = history.resumeTick;
            }

            const ULONGLONG elapsedMs = std::max<ULONGLONG>(1ull, task.lastProgressCallbackTick - baselineTick);
            const double dtSec        = static_cast<double>(elapsedMs) / 1000.0;
            smoothingElapsedMs        = elapsedMs;

            if (itemRate)
            {
                const unsigned long deltaItems = task.completedItems - history.lastItems;
                if (deltaItems > 0 && dtSec > 0.0)
                {
                    const double instItemsPerSec = static_cast<double>(deltaItems) / dtSec;
                    history.smoothedItemsPerSec  = SmoothRateForDisplay(history.smoothedItemsPerSec, instItemsPerSec, elapsedMs);
                    history.displayedItemsPerSec = history.smoothedItemsPerSec;
                    ++updatedTaskCount;
                }

                history.lastItems = task.completedItems;
            }
            else
            {
                const uint64_t deltaBytes = task.completedBytes - history.lastBytes;
                if (deltaBytes > 0 && dtSec > 0.0)
                {
                    const double instBytesPerSec = static_cast<double>(deltaBytes) / dtSec;
                    history.smoothedBytesPerSec  = SmoothRateForDisplay(history.smoothedBytesPerSec, instBytesPerSec, elapsedMs);
                    history.displayedBytesPerSec = history.smoothedBytesPerSec;
                    ++updatedTaskCount;
                }

                history.lastBytes = task.completedBytes;
            }

            history.lastProgressCallbackTick = task.lastProgressCallbackTick;
            history.resumeTick               = 0;
        }
        else if (! task.finished && task.lastProgressCallbackTick != 0 && nowTick >= task.lastProgressCallbackTick)
        {
            const ULONGLONG silenceMs = nowTick - task.lastProgressCallbackTick;
            maxCallbackSilenceMs      = std::max<uint64_t>(maxCallbackSilenceMs, silenceMs);
            if (silenceMs > 0)
            {
                ++silentTaskCount;
            }

            if (itemRate)
            {
                history.displayedItemsPerSec = DecayRateForCallbackSilence(history.smoothedItemsPerSec, silenceMs);
            }
            else
            {
                history.displayedBytesPerSec = DecayRateForCallbackSilence(history.smoothedBytesPerSec, silenceMs);
            }
        }

        if (! itemRate && task.totalBytes > 0 && task.completedBytes <= task.totalBytes && history.displayedBytesPerSec > 0.0)
        {
            const uint64_t remainingBytes = task.totalBytes - task.completedBytes;
            const double rawEtaSeconds    = static_cast<double>(remainingBytes) / history.displayedBytesPerSec;
            history.smoothedEtaSeconds =
                history.hasSmoothedEta ? SmoothEtaSecondsForDisplay(history.smoothedEtaSeconds, rawEtaSeconds, smoothingElapsedMs) : rawEtaSeconds;
            history.hasSmoothedEta = true;
        }
        else if (! itemRate)
        {
            history.hasSmoothedEta = false;
        }

        if (! task.finished)
        {
            if (history.lastDisplaySampleTick == 0 || history.lastDisplaySampleTick > nowTick)
            {
                history.lastDisplaySampleTick = nowTick;
            }
            else
            {
                const ULONGLONG displayElapsedMs = nowTick - history.lastDisplaySampleTick;
                if (displayElapsedMs > 0)
                {
                    maxDisplayGapMs            = std::max<uint64_t>(maxDisplayGapMs, displayElapsedMs);
                    const double displaySample = itemRate ? history.displayedItemsPerSec : history.displayedBytesPerSec;
                    if (displaySample > 0.0 || history.count > 0)
                    {
                        AppendResampledRateSamples(history, displayElapsedMs, displaySample, hue);
                    }
                    history.lastDisplaySampleTick = nowTick;
                }
            }
        }
    }

    for (auto it = _rates.begin(); it != _rates.end();)
    {
        const auto found = seen.find(it->first);
        if (found == seen.end())
        {
            it = _rates.erase(it);
            continue;
        }
        ++it;
    }

    if (capturePerf)
    {
        const std::wstring detail = std::format(L"tasks={} updated={} silent={} rates={}", snapshot.size(), updatedTaskCount, silentTaskCount, _rates.size());
        Debug::Perf::Emit(L"FileOps.Popup.Rate.UpdateUs", detail, PerfElapsedUs(perfStartUs), snapshot.size(), _rates.size(), S_OK);
        Debug::Perf::Emit(L"FileOps.Popup.Rate.MaxCallbackSilenceMs", detail, maxCallbackSilenceMs, silentTaskCount, snapshot.size(), S_OK);
        Debug::Perf::Emit(L"FileOps.Popup.Rate.MaxDisplayGapMs", detail, maxDisplayGapMs, snapshot.size(), _rates.size(), S_OK);
    }
}

void FileOperationsPopupInternal::FileOperationsPopupState::LayoutChrome(float width, float height) noexcept
{
    const float footerH = DipsToPixels(44.0f, _dpi);

    const float footerTop = std::max(0.0f, height - footerH);
    _listViewportRect     = D2D1::RectF(0.0f, 0.0f, width, footerTop);

    const float footerBtnH = DipsToPixels(28.0f, _dpi);
    const float footerBtnY = footerTop + (footerH - footerBtnH) / 2.0f;
    const float footerBtnW = DipsToPixels(120.0f, _dpi);
    const float footerGap  = DipsToPixels(10.0f, _dpi);

    _footerCancelAllRect = D2D1::RectF(DipsToPixels(10.0f, _dpi), footerBtnY, DipsToPixels(10.0f, _dpi) + footerBtnW, footerBtnY + footerBtnH);

    _footerQueueModeRect =
        D2D1::RectF(_footerCancelAllRect.right + footerGap, footerBtnY, _footerCancelAllRect.right + footerGap + footerBtnW, footerBtnY + footerBtnH);
}

void FileOperationsPopupInternal::FileOperationsPopupState::UpdateScrollBar(HWND hwnd, float viewH, float contentH) noexcept
{
    if (! hwnd)
    {
        return;
    }

    const int viewHeight      = std::max(0, static_cast<int>(std::ceil(viewH)));
    const int contentHeightPx = std::max(0, static_cast<int>(std::ceil(contentH)));

    if (! _scrollBarVisible)
    {
        _scrollPos = 0;
    }

    SCROLLINFO si{};
    si.cbSize      = sizeof(si);
    si.fMask       = SIF_RANGE | SIF_PAGE | SIF_POS;
    si.nMin        = 0;
    si.nMax        = std::max(0, contentHeightPx - 1);
    const int page = std::clamp(viewHeight, 1, std::numeric_limits<int>::max());
    si.nPage       = static_cast<UINT>(page);

    const int maxPos = std::max(0, si.nMax - page + 1);
    _scrollPos       = std::clamp(_scrollPos, 0, maxPos);
    si.nPos          = _scrollPos;

    SetScrollInfo(hwnd, SB_VERT, &si, TRUE);
}

void FileOperationsPopupInternal::FileOperationsPopupState::AutoResizeWindow(HWND hwnd, float desiredContentHeight, size_t taskCount) noexcept
{
    if (! hwnd || _inSizeMove)
    {
        return;
    }

    // Only auto-resize if task count or content height changed
    const bool taskCountChanged     = taskCount != _lastTaskCount;
    const bool contentHeightChanged = std::abs(desiredContentHeight - _lastAutoSizedContentHeight) > 1.0f;

    if (! taskCountChanged && ! contentHeightChanged)
    {
        return;
    }

    _lastTaskCount              = taskCount;
    _lastAutoSizedContentHeight = desiredContentHeight;

    // Get current window rect
    RECT windowRc{};
    GetWindowRect(hwnd, &windowRc);

    // Get screen work area (excludes taskbar)
    HMONITOR hMonitor = MonitorFromWindow(hwnd, MONITOR_DEFAULTTONEAREST);
    MONITORINFO mi{};
    mi.cbSize = sizeof(mi);
    if (! GetMonitorInfoW(hMonitor, &mi))
    {
        return;
    }
    const RECT& workArea      = mi.rcWork;
    const int maxScreenHeight = workArea.bottom - workArea.top;

    // Calculate the footer and chrome heights
    const float footerH             = DipsToPixels(44.0f, _dpi);
    const float desiredClientHeight = desiredContentHeight + footerH;

    // Get window style for AdjustWindowRectExForDpi
    const DWORD style   = static_cast<DWORD>(GetWindowLongPtrW(hwnd, GWL_STYLE));
    const DWORD exStyle = static_cast<DWORD>(GetWindowLongPtrW(hwnd, GWL_EXSTYLE));

    // Calculate desired window height from client height
    RECT clientRc{0, 0, windowRc.right - windowRc.left, static_cast<LONG>(std::ceil(desiredClientHeight))};
    AdjustWindowRectExForDpi(&clientRc, style, FALSE, exStyle, _dpi);

    int desiredWindowHeight = clientRc.bottom - clientRc.top;

    // Apply minimum height constraint
    constexpr int kMinClientHeightDip = 320;
    const int minClientH              = DipsToPixels(kMinClientHeightDip, _dpi);
    RECT minRc{0, 0, 0, minClientH};
    AdjustWindowRectExForDpi(&minRc, style, FALSE, exStyle, _dpi);
    const int minWindowHeight = minRc.bottom - minRc.top;
    desiredWindowHeight       = std::max(desiredWindowHeight, minWindowHeight);

    // Clamp to screen height
    desiredWindowHeight = std::min(desiredWindowHeight, maxScreenHeight);

    // Prevent resize "dancing": once the window grows to fit more lines/tasks, don't auto-shrink it again.
    if (_maxAutoSizedWindowHeight > 0)
    {
        desiredWindowHeight = std::max(desiredWindowHeight, _maxAutoSizedWindowHeight);
        desiredWindowHeight = std::min(desiredWindowHeight, maxScreenHeight);
    }

    // Calculate new position - keep top position, adjust bottom
    int newTop    = windowRc.top;
    int newBottom = newTop + desiredWindowHeight;

    // If window would extend below work area, move it up
    if (newBottom > workArea.bottom)
    {
        newBottom = workArea.bottom;
        newTop    = newBottom - desiredWindowHeight;
        // But don't go above work area
        if (newTop < workArea.top)
        {
            newTop    = workArea.top;
            newBottom = newTop + std::min(desiredWindowHeight, maxScreenHeight);
        }
    }

    // Only resize if height actually changed
    const int currentHeight = windowRc.bottom - windowRc.top;
    if (std::abs(desiredWindowHeight - currentHeight) < 2)
    {
        return;
    }

    SetWindowPos(hwnd, nullptr, windowRc.left, newTop, windowRc.right - windowRc.left, newBottom - newTop, SWP_NOZORDER | SWP_NOACTIVATE);

    _maxAutoSizedWindowHeight = std::max(_maxAutoSizedWindowHeight, desiredWindowHeight);
}

void FileOperationsPopupInternal::FileOperationsPopupState::DrawButton(const PopupButton& button, IDWriteTextFormat* format, std::wstring_view text) noexcept
{
    if (! _target || ! _borderBrush)
    {
        return;
    }

    const bool hot       = button.hit.kind == _hotHit.kind && button.hit.taskId == _hotHit.taskId && button.hit.data == _hotHit.data;
    const bool pressed   = button.hit.kind == _pressedHit.kind && button.hit.taskId == _pressedHit.taskId && button.hit.data == _pressedHit.data;
    const D2D1_RECT_F rc = button.bounds;

    if (_buttonBgBrush)
    {
        const float radius = ClampCornerRadius(rc, DipsToPixels(2.0f, _dpi));
        _target->FillRoundedRectangle(D2D1::RoundedRect(rc, radius, radius), _buttonBgBrush.get());
    }

    if (hot && _buttonHoverBrush)
    {
        const float radius = ClampCornerRadius(rc, DipsToPixels(2.0f, _dpi));
        _target->FillRoundedRectangle(D2D1::RoundedRect(rc, radius, radius), _buttonHoverBrush.get());
    }

    if (pressed && _buttonPressedBrush)
    {
        const float radius = ClampCornerRadius(rc, DipsToPixels(2.0f, _dpi));
        _target->FillRoundedRectangle(D2D1::RoundedRect(rc, radius, radius), _buttonPressedBrush.get());
    }

    {
        const float radius = ClampCornerRadius(rc, DipsToPixels(2.0f, _dpi));
        _target->DrawRoundedRectangle(D2D1::RoundedRect(rc, radius, radius), _borderBrush.get(), 1.0f);
    }

    if (format && ! text.empty() && _textBrush)
    {
        const float inset        = DipsToPixels(6.0f, _dpi);
        const D2D1_RECT_F textRc = D2D1::RectF(rc.left + inset, rc.top, rc.right - inset, rc.bottom);
        _target->DrawTextW(text.data(), static_cast<UINT32>(text.size()), format, textRc, _textBrush.get());
    }
}

void FileOperationsPopupInternal::FileOperationsPopupState::DrawMenuButton(const PopupButton& button,
                                                                           IDWriteTextFormat* format,
                                                                           std::wstring_view text) noexcept
{
    if (! _target || ! _borderBrush)
    {
        return;
    }

    const bool hot       = button.hit.kind == _hotHit.kind && button.hit.taskId == _hotHit.taskId && button.hit.data == _hotHit.data;
    const bool pressed   = button.hit.kind == _pressedHit.kind && button.hit.taskId == _pressedHit.taskId && button.hit.data == _pressedHit.data;
    const D2D1_RECT_F rc = button.bounds;

    if (_buttonBgBrush)
    {
        const float radius = ClampCornerRadius(rc, DipsToPixels(2.0f, _dpi));
        _target->FillRoundedRectangle(D2D1::RoundedRect(rc, radius, radius), _buttonBgBrush.get());
    }

    if (hot && _buttonHoverBrush)
    {
        const float radius = ClampCornerRadius(rc, DipsToPixels(2.0f, _dpi));
        _target->FillRoundedRectangle(D2D1::RoundedRect(rc, radius, radius), _buttonHoverBrush.get());
    }

    if (pressed && _buttonPressedBrush)
    {
        const float radius = ClampCornerRadius(rc, DipsToPixels(2.0f, _dpi));
        _target->FillRoundedRectangle(D2D1::RoundedRect(rc, radius, radius), _buttonPressedBrush.get());
    }

    {
        const float radius = ClampCornerRadius(rc, DipsToPixels(2.0f, _dpi));
        _target->DrawRoundedRectangle(D2D1::RoundedRect(rc, radius, radius), _borderBrush.get(), 1.0f);
    }

    const float arrowSectionW = DipsToPixels(22.0f, _dpi);
    const float separatorX    = std::clamp(rc.right - arrowSectionW, rc.left, rc.right);

    if (separatorX > rc.left && separatorX < rc.right)
    {
        const float lineInset = DipsToPixels(2.0f, _dpi);
        _target->DrawLine(D2D1::Point2F(separatorX, rc.top + lineInset), D2D1::Point2F(separatorX, rc.bottom - lineInset), _borderBrush.get(), 1.0f);
    }

    if (_textBrush)
    {
        const float centerX = (separatorX + rc.right) * 0.5f;
        const float centerY = (rc.top + rc.bottom) * 0.5f;

        const float halfW     = DipsToPixels(4.0f, _dpi);
        const float halfH     = DipsToPixels(2.5f, _dpi);
        const float thickness = DipsToPixels(1.5f, _dpi);

        _target->DrawLine(D2D1::Point2F(centerX - halfW, centerY - halfH), D2D1::Point2F(centerX, centerY + halfH), _textBrush.get(), thickness);
        _target->DrawLine(D2D1::Point2F(centerX, centerY + halfH), D2D1::Point2F(centerX + halfW, centerY - halfH), _textBrush.get(), thickness);
    }

    if (format && ! text.empty() && _textBrush)
    {
        const float inset        = DipsToPixels(6.0f, _dpi);
        const float right        = std::max(rc.left + inset, separatorX - inset);
        const D2D1_RECT_F textRc = D2D1::RectF(rc.left + inset, rc.top, right, rc.bottom);
        _target->DrawTextW(text.data(), static_cast<UINT32>(text.size()), format, textRc, _textBrush.get());
    }
}

void FileOperationsPopupInternal::FileOperationsPopupState::DrawCheckboxBox(const D2D1_RECT_F& rect, bool checked) noexcept
{
    if (! _target)
    {
        return;
    }

    const float size = std::max(0.0f, std::min(rect.right - rect.left, rect.bottom - rect.top));
    if (size <= 1.0f)
    {
        return;
    }

    const float left = rect.left + (rect.right - rect.left - size) * 0.5f;
    const float top  = rect.top + (rect.bottom - rect.top - size) * 0.5f;

    const D2D1_RECT_F boxRc = D2D1::RectF(left, top, left + size, top + size);

    ID2D1Brush* base = _buttonBgBrush ? _buttonBgBrush.get() : (_bgBrush ? _bgBrush.get() : nullptr);
    if (base)
    {
        _target->FillRectangle(boxRc, base);
    }

    if (checked && _checkboxFillBrush)
    {
        _target->FillRectangle(boxRc, _checkboxFillBrush.get());
    }

    if (_borderBrush)
    {
        _target->DrawRectangle(boxRc, _borderBrush.get(), 1.0f);
    }

    if (! checked)
    {
        return;
    }

    ID2D1Brush* checkBrush = _checkboxCheckBrush ? _checkboxCheckBrush.get() : (_textBrush ? _textBrush.get() : nullptr);
    if (! checkBrush)
    {
        return;
    }

    const D2D1_POINT_2F p1{left + size * 0.20f, top + size * 0.55f};
    const D2D1_POINT_2F p2{left + size * 0.42f, top + size * 0.75f};
    const D2D1_POINT_2F p3{left + size * 0.80f, top + size * 0.30f};

    const float thickness = DipsToPixels(1.8f, _dpi);
    _target->DrawLine(p1, p2, checkBrush, thickness);
    _target->DrawLine(p2, p3, checkBrush, thickness);
}

void FileOperationsPopupInternal::FileOperationsPopupState::DrawCollapseChevron(const D2D1_RECT_F& rc, bool collapsed) noexcept
{
    if (! _target || ! _textBrush)
    {
        return;
    }

    const float centerX   = (rc.left + rc.right) * 0.5f;
    const float centerY   = (rc.top + rc.bottom) * 0.5f;
    const float halfW     = DipsToPixels(4.0f, _dpi);
    const float halfH     = DipsToPixels(2.5f, _dpi);
    const float thickness = DipsToPixels(1.5f, _dpi);

    if (collapsed)
    {
        // Draw a down chevron (expand).
        _target->DrawLine(D2D1::Point2F(centerX - halfW, centerY - halfH), D2D1::Point2F(centerX, centerY + halfH), _textBrush.get(), thickness);
        _target->DrawLine(D2D1::Point2F(centerX, centerY + halfH), D2D1::Point2F(centerX + halfW, centerY - halfH), _textBrush.get(), thickness);
    }
    else
    {
        // Draw an up chevron (collapse).
        _target->DrawLine(D2D1::Point2F(centerX - halfW, centerY + halfH), D2D1::Point2F(centerX, centerY - halfH), _textBrush.get(), thickness);
        _target->DrawLine(D2D1::Point2F(centerX, centerY - halfH), D2D1::Point2F(centerX + halfW, centerY + halfH), _textBrush.get(), thickness);
    }
}

void FileOperationsPopupInternal::FileOperationsPopupState::DrawBandwidthGraph(const D2D1_RECT_F& rect,
                                                                               const RateHistory& history,
                                                                               uint64_t limitBytesPerSecond,
                                                                               std::wstring_view overlayText,
                                                                               bool showAnimation,
                                                                               bool rainbowMode,
                                                                               ULONGLONG tick) noexcept
{
    if (! _target)
    {
        return;
    }

    const float w = rect.right - rect.left;
    const float h = rect.bottom - rect.top;
    if (w <= 0.0f || h <= 0.0f)
    {
        return;
    }

    if (_graphBgBrush)
    {
        _target->FillRectangle(rect, _graphBgBrush.get());
    }

    const AppTheme* theme  = folderWindow ? &folderWindow->GetTheme() : nullptr;
    const float rainbowSat = 0.85f;
    const float rainbowVal = (theme && theme->dark) ? 0.80f : 0.90f;

    auto sampleColorFromHue = [&](float hue, float alpha) noexcept -> D2D1_COLOR_F
    {
        if (hue < 0.0f)
        {
            D2D1_COLOR_F c = theme ? theme->navigationView.accent : D2D1::ColorF(D2D1::ColorF::DodgerBlue);
            c.a            = alpha;
            return c;
        }
        return ColorFromHSV(hue, rainbowSat, rainbowVal, alpha);
    };

    // Helper to compute rainbow color based on tick
    auto computeRainbowColor = [](ULONGLONG tick, ULONGLONG periodMs, float saturation, float value, float alpha) -> D2D1_COLOR_F
    {
        const float hue = static_cast<float>((tick % periodMs) * 360ull / periodMs);
        return ColorFromHSV(hue, saturation, value, alpha);
    };

    // Draw animation for pre-calculation phase
    if (showAnimation && _graphDynamicBrush)
    {
        // Pulsing background effect
        constexpr ULONGLONG kPulsePeriodMs = 1600ull;
        const ULONGLONG pulsePhase         = tick % kPulsePeriodMs;
        const float pulseT                 = static_cast<float>(pulsePhase) / static_cast<float>(kPulsePeriodMs);
        const float pulseAlpha             = 0.15f + 0.15f * std::sin(pulseT * 2.0f * 3.14159265f);

        D2D1_COLOR_F pulseColor = _graphFillBaseColor;
        if (rainbowMode)
        {
            // Rainbow: cycle through hues for pulse background
            constexpr ULONGLONG kRainbowPeriodMs = 3000ull;
            pulseColor                           = computeRainbowColor(tick, kRainbowPeriodMs, 0.6f, 0.8f, pulseAlpha);
        }
        else
        {
            pulseColor.a = pulseAlpha;
        }

        _graphDynamicBrush->SetColor(pulseColor);
        _target->FillRectangle(rect, _graphDynamicBrush.get());

        // Horizontal sweep line effect
        constexpr ULONGLONG kSweepPeriodMs = 1200ull;
        const ULONGLONG sweepPhase         = tick % kSweepPeriodMs;
        const float sweepT                 = static_cast<float>(sweepPhase) / static_cast<float>(kSweepPeriodMs);
        const float sweepX                 = rect.left + w * sweepT;

        D2D1_COLOR_F sweepColor = _graphFillBaseColor;
        if (rainbowMode)
        {
            // Rainbow: sweep line changes color each sweep
            sweepColor = computeRainbowColor(tick, kSweepPeriodMs, 0.85f, 0.9f, 0.7f);
        }
        else
        {
            sweepColor.a = 0.5f;
        }

        const float sweepWidth = DipsToPixels(2.0f, _dpi);
        _graphDynamicBrush->SetColor(sweepColor);
        _target->DrawLine(D2D1::Point2F(sweepX, rect.top), D2D1::Point2F(sweepX, rect.bottom), _graphDynamicBrush.get(), sweepWidth);

        // Spinner dots effect (3 dots bouncing)
        constexpr ULONGLONG kSpinPeriodMs = 1000ull;
        constexpr int kDotCount           = 3;
        const float centerX               = rect.left + w * 0.5f;
        const float centerY               = rect.bottom - h * 0.35f;
        const float dotSpacing            = DipsToPixels(10.0f, _dpi);

        for (int i = 0; i < kDotCount; ++i)
        {
            const float phaseOffset  = static_cast<float>(i) / static_cast<float>(kDotCount);
            const ULONGLONG dotPhase = (tick + static_cast<ULONGLONG>(phaseOffset * kSpinPeriodMs)) % kSpinPeriodMs;
            const float dotT         = static_cast<float>(dotPhase) / static_cast<float>(kSpinPeriodMs);
            const float bounce       = std::abs(std::sin(dotT * 3.14159265f));

            const float dotX      = centerX + (static_cast<float>(i) - 1.0f) * dotSpacing;
            const float dotY      = centerY - bounce * DipsToPixels(8.0f, _dpi);
            const float dotRadius = DipsToPixels(3.0f, _dpi);

            D2D1_COLOR_F dotColor = _graphFillBaseColor;
            if (rainbowMode)
            {
                // Rainbow: each dot has its own hue offset
                constexpr ULONGLONG kDotRainbowPeriodMs = 2000ull;
                const ULONGLONG dotRainbowPhase         = tick + static_cast<ULONGLONG>(i * 667); // 120 degree offset per dot
                dotColor                                = computeRainbowColor(dotRainbowPhase, kDotRainbowPeriodMs, 0.85f, 0.9f, 0.6f + 0.4f * bounce);
            }
            else
            {
                dotColor.a = 0.6f + 0.4f * bounce;
            }

            _graphDynamicBrush->SetColor(dotColor);
            _target->FillEllipse(D2D1::Ellipse(D2D1::Point2F(dotX, dotY), dotRadius, dotRadius), _graphDynamicBrush.get());
        }
    }

    if (_borderBrush)
    {
        _target->DrawRectangle(rect, _borderBrush.get(), 1.0f);
    }

    float maxSpeed = 0.0f;
    for (size_t i = 0; i < history.count; ++i)
    {
        const size_t index = (history.writeIndex + RateHistory::kMaxSamples - history.count + i) % RateHistory::kMaxSamples;
        maxSpeed           = std::max(maxSpeed, history.samples[index]);
    }

    if (limitBytesPerSecond > 0)
    {
        maxSpeed = std::max(maxSpeed, static_cast<float>(limitBytesPerSecond));
    }

    if (maxSpeed <= 0.0f)
    {
        maxSpeed = 1.0f;
    }

    const float axisMax = std::max(1.0f, maxSpeed * 1.10f);

    const bool canDrawSamples = _graphLineBrush && history.count >= 2;

    std::array<D2D1_POINT_2F, RateHistory::kMaxSamples> points{};
    std::array<float, RateHistory::kMaxSamples> sampleHues{};
    size_t count  = 0;
    size_t oldest = 0;
    if (canDrawSamples)
    {
        count  = history.count;
        oldest = (history.writeIndex + RateHistory::kMaxSamples - count) % RateHistory::kMaxSamples;

        for (size_t i = 0; i < count; ++i)
        {
            const size_t index = (oldest + i) % RateHistory::kMaxSamples;
            const float speed  = history.samples[index];
            sampleHues[i]      = history.hues[index];

            const float xFrac = static_cast<float>(i) / static_cast<float>(count - 1u);
            const float yFrac = Clamp01(speed / axisMax);

            const float x = rect.left + w * xFrac;
            const float y = rect.bottom - h * yFrac;
            points[i]     = D2D1::Point2F(x, y);
        }

        if (_graphFillBrush && _d2dFactory)
        {
            if (rainbowMode && _graphDynamicBrush && count >= 2)
            {
                // Rainbow: draw per-segment trapezoids with individual hue colors
                const float fillAlpha = _graphFillBaseColor.a;
                for (size_t i = 1; i < count; ++i)
                {
                    // Use the hue from the right-side sample (newer) for each segment
                    const float hue                = sampleHues[i];
                    const D2D1_COLOR_F segmentFill = sampleColorFromHue(hue, fillAlpha);
                    _graphDynamicBrush->SetColor(segmentFill);

                    // Build a trapezoid: from points[i-1] to points[i] to bottom-right to bottom-left
                    wil::com_ptr<ID2D1PathGeometry> trapezoid;
                    const HRESULT hrGeo = _d2dFactory->CreatePathGeometry(trapezoid.put());
                    if (SUCCEEDED(hrGeo) && trapezoid)
                    {
                        wil::com_ptr<ID2D1GeometrySink> sink;
                        const HRESULT hrSink = trapezoid->Open(sink.put());
                        if (SUCCEEDED(hrSink) && sink)
                        {
                            sink->SetFillMode(D2D1_FILL_MODE_WINDING);
                            sink->BeginFigure(points[i - 1u], D2D1_FIGURE_BEGIN_FILLED);
                            sink->AddLine(points[i]);
                            sink->AddLine(D2D1::Point2F(points[i].x, rect.bottom));
                            sink->AddLine(D2D1::Point2F(points[i - 1u].x, rect.bottom));
                            sink->EndFigure(D2D1_FIGURE_END_CLOSED);
                            sink->Close();
                            _target->FillGeometry(trapezoid.get(), _graphDynamicBrush.get());
                        }
                    }
                }
            }
            else
            {
                // Non-rainbow: draw single fill geometry
                wil::com_ptr<ID2D1PathGeometry> geometry;
                const HRESULT hrGeo = _d2dFactory->CreatePathGeometry(geometry.put());
                if (SUCCEEDED(hrGeo) && geometry)
                {
                    wil::com_ptr<ID2D1GeometrySink> sink;
                    const HRESULT hrSink = geometry->Open(sink.put());
                    if (SUCCEEDED(hrSink) && sink)
                    {
                        sink->SetFillMode(D2D1_FILL_MODE_WINDING);
                        sink->BeginFigure(points[0], D2D1_FIGURE_BEGIN_FILLED);
                        sink->AddLines(points.data() + 1, static_cast<UINT32>(count - 1u));

                        sink->AddLine(D2D1::Point2F(points[count - 1u].x, rect.bottom));
                        sink->AddLine(D2D1::Point2F(points[0].x, rect.bottom));

                        sink->EndFigure(D2D1_FIGURE_END_CLOSED);
                        sink->Close();

                        _target->FillGeometry(geometry.get(), _graphFillBrush.get());
                    }
                }
            }
        }
    }

    if (_graphGridBrush)
    {
        for (int i = 1; i <= 3; ++i)
        {
            const float frac = static_cast<float>(i) / 4.0f;
            const float y    = rect.bottom - h * frac;
            _target->DrawLine(D2D1::Point2F(rect.left, y), D2D1::Point2F(rect.right, y), _graphGridBrush.get(), 1.0f);
        }
    }

    if (limitBytesPerSecond > 0 && _graphLimitBrush)
    {
        const float limitFrac = Clamp01(static_cast<float>(static_cast<double>(limitBytesPerSecond) / static_cast<double>(axisMax)));
        const float y         = rect.bottom - h * limitFrac;
        _target->DrawLine(D2D1::Point2F(rect.left, y), D2D1::Point2F(rect.right, y), _graphLimitBrush.get(), 1.0f);
    }

    if (canDrawSamples && rainbowMode)
    {
        // Rainbow: draw each line segment with its own hue from the stored per-sample hue
        if (_graphDynamicBrush)
        {
            for (size_t i = 1; i < count; ++i)
            {
                const float hue                = sampleHues[i];
                const D2D1_COLOR_F segmentLine = sampleColorFromHue(hue, 1.0f);
                _graphDynamicBrush->SetColor(segmentLine);
                _target->DrawLine(points[i - 1u], points[i], _graphDynamicBrush.get(), 1.5f);
            }
        }
        else
        {
            for (size_t i = 1; i < count; ++i)
            {
                _target->DrawLine(points[i - 1u], points[i], _graphLineBrush.get(), 1.5f);
            }
        }
    }
    else if (canDrawSamples)
    {
        for (size_t i = 1; i < count; ++i)
        {
            _target->DrawLine(points[i - 1u], points[i], _graphLineBrush.get(), 1.5f);
        }
    }

    if (! overlayText.empty() && _graphOverlayFormat && _textBrush)
    {
        // Draw shadow behind text for better visibility
        if (_graphTextShadowBrush)
        {
            const float shadowOffset = DipsToPixels(1.0f, _dpi);
            const D2D1_RECT_F shadowRect =
                D2D1::RectF(rect.left + shadowOffset, rect.top + shadowOffset, rect.right + shadowOffset, rect.bottom + shadowOffset);
            _target->DrawTextW(overlayText.data(), static_cast<UINT32>(overlayText.size()), _graphOverlayFormat.get(), shadowRect, _graphTextShadowBrush.get());
        }

        // Draw main text
        _target->DrawTextW(overlayText.data(), static_cast<UINT32>(overlayText.size()), _graphOverlayFormat.get(), rect, _textBrush.get());
    }
}

void FileOperationsPopupInternal::FileOperationsPopupState::Render(HWND hwnd) noexcept
{
    if (! hwnd)
    {
        return;
    }

    PAINTSTRUCT ps{};
    wil::unique_hdc_paint hdc = wil::BeginPaint(hwnd, &ps);
    static_cast<void>(hdc.get());
    static_cast<void>(ps);

    if (! hostLifetime.lock())
    {
        return;
    }

    const FolderWindow* folderWindowPtr = folderWindow;
    if (! folderWindowPtr)
    {
        return;
    }
    const AppTheme& appTheme = folderWindowPtr->GetTheme();

    EnsureTarget(hwnd);
    EnsureTextFormats();
    EnsureBrushes();

    if (! _target || ! _bgBrush || ! _textBrush || ! _borderBrush)
    {
        return;
    }

    const bool capturePerf                   = Debug::Perf::IsCaptureEnabled();
    const uint64_t renderStartedUs           = capturePerf ? PerfNowUs() : 0u;
    const uint64_t snapshotStartedUs         = capturePerf ? PerfNowUs() : 0u;
    const std::vector<TaskSnapshot> snapshot = BuildSnapshot();
    const uint64_t snapshotUs                = capturePerf ? PerfElapsedUs(snapshotStartedUs) : 0u;
    CleanupCollapsedTasks(snapshot);
    UpdateCaptionStatus(hwnd, snapshot);

    constexpr ULONGLONG kCompletedInFlightGraceMs = 300ull;
    const ULONGLONG renderTick                    = GetTickCount64();

    float width  = 0.0f;
    float height = 0.0f;

    const float padding = DipsToPixels(10.0f, _dpi);
    const float cardGap = DipsToPixels(10.0f, _dpi);

    const float expandedCardH  = DipsToPixels(280.0f, _dpi);
    const float collapsedCardH = DipsToPixels(44.0f, _dpi);
    const float baseLineH      = DipsToPixels(18.0f, _dpi);
    const float fromToGapY     = DipsToPixels(4.0f, _dpi);

    const uint64_t cardLayoutStartedUs = capturePerf ? PerfNowUs() : 0u;
    std::vector<float> cardHeights;
    cardHeights.reserve(snapshot.size());
    for (const TaskSnapshot& task : snapshot)
    {
        if (task.kind == TaskSnapshot::Kind::Informational)
        {
            const FolderWindow::InformationalTaskUpdate& info = task.informational;
            const float expandedBase                          = task.finished ? DipsToPixels(180.0f, _dpi) : DipsToPixels(210.0f, _dpi);
            float h                                           = IsTaskCollapsed(task.taskId) ? collapsedCardH : expandedBase;

            if (! IsTaskCollapsed(task.taskId) && ! task.finished && info.kind == FolderWindow::InformationalTaskUpdate::Kind::CompareDirectories &&
                info.contentActive && info.contentInFlightCount > 1u)
            {
                size_t activeInFlightCount = 0;
                for (size_t i = 0; i < info.contentInFlightCount; ++i)
                {
                    const auto& entry          = info.contentInFlight[i];
                    const bool active          = entry.totalBytes == 0 || entry.completedBytes < entry.totalBytes;
                    const bool recentCompleted = ! active && entry.totalBytes > 0 && entry.completedBytes >= entry.totalBytes && entry.lastUpdateTick != 0 &&
                                                 renderTick >= entry.lastUpdateTick && (renderTick - entry.lastUpdateTick) <= kCompletedInFlightGraceMs;
                    if (active || recentCompleted)
                    {
                        ++activeInFlightCount;
                    }
                }

                const size_t lineCount = std::max<size_t>(1u, activeInFlightCount);
                if (lineCount > 1u)
                {
                    h += static_cast<float>(lineCount - 1u) * baseLineH;
                }
            }

            cardHeights.push_back(h);
            continue;
        }

        float h = IsTaskCollapsed(task.taskId) ? collapsedCardH : expandedCardH;
        if (! IsTaskCollapsed(task.taskId) && task.finished)
        {
            h = DipsToPixels(178.0f, _dpi);
        }
        if (! IsTaskCollapsed(task.taskId) && ! task.finished && (task.operation == FILESYSTEM_COPY || task.operation == FILESYSTEM_MOVE))
        {
            size_t activeInFlightCount = 0;
            for (size_t i = 0; i < task.inFlightFileCount; ++i)
            {
                const auto& entry          = task.inFlightFiles[i];
                const bool active          = entry.totalBytes == 0 || entry.completedBytes < entry.totalBytes;
                const bool recentCompleted = ! active && entry.totalBytes > 0 && entry.completedBytes >= entry.totalBytes && entry.lastUpdateTick != 0 &&
                                             renderTick >= entry.lastUpdateTick && (renderTick - entry.lastUpdateTick) <= kCompletedInFlightGraceMs;
                if (active || recentCompleted)
                {
                    ++activeInFlightCount;
                }
            }

            const size_t lineCount = std::max<size_t>(1u, activeInFlightCount);
            if (lineCount > 1u)
            {
                h += static_cast<float>(lineCount - 1u) * baseLineH;
            }
            h += fromToGapY;
        }
        if (! IsTaskCollapsed(task.taskId) && ! task.finished && task.conflict.active)
        {
            // Extra room for inline conflict prompt + action buttons.
            h += baseLineH * 3.0f;
        }
        cardHeights.push_back(h);
    }
    const uint64_t cardLayoutUs = capturePerf ? PerfElapsedUs(cardLayoutStartedUs) : 0u;

    const size_t taskCount = snapshot.size();
    if (taskCount == 0)
    {
        _contentHeight = padding * 2.0f;
    }
    else
    {
        float sumHeights = 0.0f;
        for (const float h : cardHeights)
        {
            sumHeights += h;
        }
        _contentHeight = padding * 2.0f + sumHeights + static_cast<float>(taskCount - 1u) * cardGap;
    }

    // Auto-resize window to fit content (limited to screen height)
    const uint64_t autoResizeStartedUs = capturePerf ? PerfNowUs() : 0u;
    AutoResizeWindow(hwnd, _contentHeight, taskCount);
    const uint64_t autoResizeUs = capturePerf ? PerfElapsedUs(autoResizeStartedUs) : 0u;

    const uint64_t scrollLayoutStartedUs = capturePerf ? PerfNowUs() : 0u;
    bool scrollReady                     = false;
    for (int pass = 0; pass < 2; ++pass)
    {
        RECT clientRc{};
        GetClientRect(hwnd, &clientRc);
        const UINT clientW = static_cast<UINT>(std::max(0L, clientRc.right - clientRc.left));
        const UINT clientH = static_cast<UINT>(std::max(0L, clientRc.bottom - clientRc.top));

        if (_target && (_clientSize.cx != static_cast<LONG>(clientW) || _clientSize.cy != static_cast<LONG>(clientH)))
        {
            _clientSize.cx = static_cast<LONG>(clientW);
            _clientSize.cy = static_cast<LONG>(clientH);
            _target->Resize(D2D1::SizeU(clientW, clientH));
        }

        width  = static_cast<float>(clientW);
        height = static_cast<float>(clientH);

        LayoutChrome(width, height);

        const float viewH              = std::max(0.0f, _listViewportRect.bottom - _listViewportRect.top);
        const bool shouldShowScrollBar = _contentHeight > viewH;
        if (shouldShowScrollBar != _scrollBarVisible)
        {
            _scrollBarVisible = shouldShowScrollBar;
            if (! shouldShowScrollBar)
            {
                _scrollPos = 0;
                _scrollY   = 0.0f;
            }

            ShowScrollBar(hwnd, SB_VERT, shouldShowScrollBar ? TRUE : FALSE);

            _hotHit     = {};
            _pressedHit = {};

            SetWindowPos(hwnd, nullptr, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);
            continue;
        }

        const float maxScroll = std::max(0.0f, _contentHeight - viewH);
        _scrollPos            = std::clamp(_scrollPos, 0, static_cast<int>(std::ceil(maxScroll)));
        UpdateScrollBar(hwnd, viewH, _contentHeight);
        _scrollY    = static_cast<float>(_scrollPos);
        scrollReady = true;
        break;
    }

    if (! scrollReady)
    {
        const float viewH     = std::max(0.0f, _listViewportRect.bottom - _listViewportRect.top);
        const float maxScroll = std::max(0.0f, _contentHeight - viewH);
        _scrollPos            = std::clamp(_scrollPos, 0, static_cast<int>(std::ceil(maxScroll)));
        UpdateScrollBar(hwnd, viewH, _contentHeight);
        _scrollY = static_cast<float>(_scrollPos);
    }
    const uint64_t scrollLayoutUs = capturePerf ? PerfElapsedUs(scrollLayoutStartedUs) : 0u;

    _buttons.clear();

    HRESULT hrEndDraw            = S_OK;
    const uint64_t drawStartedUs = capturePerf ? PerfNowUs() : 0u;
    {
        _target->BeginDraw();
        auto endDraw = wil::scope_exit([&] { hrEndDraw = _target->EndDraw(); });

        _target->SetTransform(D2D1::Matrix3x2F::Identity());

        const D2D1_RECT_F clientRect = D2D1::RectF(0.0f, 0.0f, width, height);
        _target->FillRectangle(clientRect, _bgBrush.get());

        const float footerH          = DipsToPixels(44.0f, _dpi);
        const float footerTop        = std::max(0.0f, height - footerH);
        const D2D1_RECT_F footerRect = D2D1::RectF(0.0f, footerTop, width, height);
        _target->DrawRectangle(footerRect, _borderBrush.get(), 1.0f);

        PopupButton cancelAllBtn{};
        cancelAllBtn.bounds   = _footerCancelAllRect;
        cancelAllBtn.hit.kind = PopupHitTest::Kind::FooterCancelAll;
        _buttons.push_back(cancelAllBtn);

        PopupButton queueBtn{};
        queueBtn.bounds   = _footerQueueModeRect;
        queueBtn.hit.kind = PopupHitTest::Kind::FooterQueueMode;
        _buttons.push_back(queueBtn);

        const float footerBtnW = DipsToPixels(120.0f, _dpi);
        const float footerGap  = DipsToPixels(10.0f, _dpi);
        PopupButton autoDismissBtn{};
        autoDismissBtn.bounds   = D2D1::RectF(_footerQueueModeRect.right + footerGap,
                                              _footerQueueModeRect.top,
                                              _footerQueueModeRect.right + footerGap + (footerBtnW * 1.7f),
                                              _footerQueueModeRect.bottom);
        autoDismissBtn.hit.kind = PopupHitTest::Kind::FooterAutoDismissSuccess;
        _buttons.push_back(autoDismissBtn);

        const bool hasActiveOperations = fileOps ? fileOps->HasActiveOperations() : false;
        const UINT footerActionId = hasActiveOperations ? static_cast<UINT>(IDS_FILEOPS_BTN_CANCEL_ALL) : static_cast<UINT>(IDS_FILEOPS_BTN_CLEAR_COMPLETED);
        const std::wstring cancelAllText = LoadStringResource(nullptr, footerActionId);
        DrawButton(cancelAllBtn, _buttonFormat.get(), cancelAllText);

        const bool queueMode        = fileOps ? fileOps->GetQueueNewTasks() : true;
        const UINT modeId           = queueMode ? static_cast<UINT>(IDS_FILEOPS_BTN_MODE_QUEUE) : static_cast<UINT>(IDS_FILEOPS_BTN_MODE_PARALLEL);
        const std::wstring modeText = LoadStringResource(nullptr, modeId);
        DrawButton(queueBtn, _buttonFormat.get(), modeText);

        const bool autoDismissSuccess = fileOps ? fileOps->GetAutoDismissSuccess() : false;
        if (_smallFormat && _textBrush)
        {
            const float insetX    = DipsToPixels(10.0f, _dpi);
            const float checkSize = DipsToPixels(12.0f, _dpi);
            const float checkTop  = autoDismissBtn.bounds.top + (autoDismissBtn.bounds.bottom - autoDismissBtn.bounds.top - checkSize) * 0.5f;
            const D2D1_RECT_F checkRc =
                D2D1::RectF(autoDismissBtn.bounds.left + insetX, checkTop, autoDismissBtn.bounds.left + insetX + checkSize, checkTop + checkSize);
            DrawCheckboxBox(checkRc, autoDismissSuccess);

            const float gapX          = DipsToPixels(8.0f, _dpi);
            const float labelLeft     = checkRc.right + gapX;
            const float rightInset    = insetX;
            const float labelRight    = std::max(labelLeft, autoDismissBtn.bounds.right - rightInset);
            const D2D1_RECT_F labelRc = D2D1::RectF(labelLeft, autoDismissBtn.bounds.top, labelRight, autoDismissBtn.bounds.bottom);

            const UINT labelId =
                autoDismissSuccess ? static_cast<UINT>(IDS_FILEOPS_CHECK_AUTODISMISS_ON) : static_cast<UINT>(IDS_FILEOPS_CHECK_AUTODISMISS_OFF);
            const std::wstring label = LoadStringResource(nullptr, labelId);
            _target->DrawTextW(label.data(), static_cast<UINT32>(label.size()), _smallFormat.get(), labelRc, _textBrush.get());
        }

        float y = _listViewportRect.top + padding - _scrollY;

        const float cardW = std::max(0.0f, width - padding * 2.0f);

        _target->PushAxisAlignedClip(_listViewportRect, D2D1_ANTIALIAS_MODE_PER_PRIMITIVE);

        for (size_t taskIndex = 0; taskIndex < taskCount; ++taskIndex)
        {
            const TaskSnapshot& task   = snapshot[taskIndex];
            const float taskCardH      = cardHeights[taskIndex];
            const D2D1_RECT_F cardRect = D2D1::RectF(padding, y, padding + cardW, y + taskCardH);

            const bool visible = cardRect.bottom >= _listViewportRect.top && cardRect.top <= _listViewportRect.bottom;
            if (visible)
            {
                _target->DrawRoundedRectangle(D2D1::RoundedRect(cardRect, DipsToPixels(2.0f, _dpi), DipsToPixels(2.0f, _dpi)), _borderBrush.get(), 1.0f);

                _target->PushAxisAlignedClip(cardRect, D2D1_ANTIALIAS_MODE_PER_PRIMITIVE);
                auto popCardClip = wil::scope_exit([&] { _target->PopAxisAlignedClip(); });

                const float padX         = DipsToPixels(10.0f, _dpi);
                const float textX        = cardRect.left + padX;
                const float contentRight = cardRect.right - padX;
                const float lineH        = DipsToPixels(18.0f, _dpi);
                float textY              = cardRect.top + DipsToPixels(8.0f, _dpi);
                const float textMaxW     = std::max(0.0f, contentRight - textX);

                const bool isCollapsedTask = IsTaskCollapsed(task.taskId);

                if (task.kind == TaskSnapshot::Kind::Informational)
                {
                    const FolderWindow::InformationalTaskUpdate& info = task.informational;
                    const ULONGLONG nowTick                           = renderTick;

                    std::wstring headerText = info.title;
                    if (info.finished && ! info.title.empty())
                    {
                        const HRESULT cancelledHr = HRESULT_FROM_WIN32(ERROR_CANCELLED);
                        std::wstring statusText;
                        if (SUCCEEDED(info.resultHr))
                        {
                            statusText = LoadStringResource(nullptr, IDS_FILEOPS_STATUS_COMPLETED);
                        }
                        else if (info.resultHr == cancelledHr || info.resultHr == E_ABORT)
                        {
                            statusText = LoadStringResource(nullptr, IDS_FILEOPS_STATUS_CANCELED);
                        }
                        else
                        {
                            statusText = FormatStringResource(nullptr, IDS_FMT_FILEOPS_STATUS_FAILED, static_cast<unsigned long>(info.resultHr));
                        }
                        headerText = FormatStringResource(nullptr, IDS_FMT_FILEOPS_OP_STATUS, info.title, statusText);
                    }

                    const float collapseBtnSize = DipsToPixels(18.0f, _dpi);
                    const float collapseBtnGap  = DipsToPixels(6.0f, _dpi);

                    const float headerTop    = isCollapsedTask ? cardRect.top + (taskCardH - lineH) * 0.5f : textY;
                    const float headerBottom = headerTop + lineH;
                    const float collapseTop  = headerTop + (lineH - collapseBtnSize) * 0.5f;
                    const float collapseLeft = std::max(textX, contentRight - collapseBtnSize);

                    PopupButton collapseBtn{};
                    collapseBtn.bounds     = D2D1::RectF(collapseLeft, collapseTop, contentRight, collapseTop + collapseBtnSize);
                    collapseBtn.hit.kind   = PopupHitTest::Kind::TaskToggleCollapse;
                    collapseBtn.hit.taskId = task.taskId;
                    _buttons.push_back(collapseBtn);
                    DrawButton(collapseBtn, nullptr, {});
                    DrawCollapseChevron(collapseBtn.bounds, isCollapsedTask);

                    const float headerRight = std::max(textX, collapseBtn.bounds.left - collapseBtnGap);
                    float headerLeft        = textX;

                    CaptionStatus statusIcon = CaptionStatus::None;
                    {
                        const HRESULT cancelledHr = HRESULT_FROM_WIN32(ERROR_CANCELLED);
                        if (info.finished && FAILED(info.resultHr) && info.resultHr != cancelledHr && info.resultHr != E_ABORT)
                        {
                            statusIcon = CaptionStatus::Error;
                        }
                        else if (info.finished && SUCCEEDED(info.resultHr))
                        {
                            statusIcon = CaptionStatus::Ok;
                        }
                    }

                    if (statusIcon != CaptionStatus::None && _target)
                    {
                        const float iconSize = DipsToPixels(16.0f, _dpi);
                        const float iconGap  = DipsToPixels(6.0f, _dpi);

                        D2D1_RECT_F iconRc = D2D1::RectF(textX, headerTop, textX + iconSize, headerBottom);
                        iconRc.right       = std::min(iconRc.right, headerRight);

                        wchar_t fluentGlyph = 0;
                        wchar_t fallback    = 0;
                        ID2D1Brush* brush   = _textBrush.get();
                        switch (statusIcon)
                        {
                            case CaptionStatus::Ok:
                                fluentGlyph = FluentIcons::kCheckMark;
                                fallback    = FluentIcons::kFallbackCheckMark;
                                brush       = _statusOkBrush ? _statusOkBrush.get() : (_textBrush ? _textBrush.get() : nullptr);
                                break;
                            case CaptionStatus::Error:
                                fluentGlyph = FluentIcons::kError;
                                fallback    = FluentIcons::kFallbackError;
                                brush       = _statusErrorBrush ? _statusErrorBrush.get() : (_textBrush ? _textBrush.get() : nullptr);
                                break;
                            case CaptionStatus::Warning:
                            case CaptionStatus::None:
                            default: break;
                        }

                        const bool useFluentFormat =
                            _statusIconFormat != nullptr && DirectWriteFormatHasGlyph(_dwriteFactory.get(), _statusIconFormat.get(), fluentGlyph);
                        const wchar_t glyph       = useFluentFormat ? fluentGlyph : fallback;
                        IDWriteTextFormat* format = useFluentFormat ? _statusIconFormat.get() : _statusIconFallbackFormat.get();

                        if (format && brush && glyph != 0 && iconRc.right > iconRc.left)
                        {
                            const wchar_t text[2]{glyph, 0};
                            _target->DrawText(text, 1u, format, iconRc, brush, D2D1_DRAW_TEXT_OPTIONS_CLIP);
                            headerLeft = std::min(headerRight, iconRc.right + iconGap);
                        }
                    }

                    if (_headerFormat)
                    {
                        const D2D1_RECT_F headerRc = D2D1::RectF(headerLeft, headerTop, headerRight, headerBottom);
                        _target->DrawTextW(headerText.data(),
                                           static_cast<UINT32>(headerText.size()),
                                           _headerFormat.get(),
                                           headerRc,
                                           _textBrush.get(),
                                           D2D1_DRAW_TEXT_OPTIONS_CLIP);
                    }

                    if (isCollapsedTask)
                    {
                        const float gapAfter = (taskIndex + 1u < taskCount) ? cardGap : 0.0f;
                        y += taskCardH + gapAfter;
                        continue;
                    }

                    textY = headerBottom + DipsToPixels(6.0f, _dpi);

                    const auto drawLabeledPathLine = [&](UINT labelId, const std::filesystem::path& path) noexcept
                    {
                        if (! _dwriteFactory || ! _smallFormat || ! _bodyFormat || ! _textBrush || ! _subTextBrush)
                        {
                            return;
                        }

                        if (textY + lineH > cardRect.bottom)
                        {
                            return;
                        }

                        const std::wstring label = LoadStringResource(nullptr, labelId);
                        if (label.empty())
                        {
                            return;
                        }

                        const float labelW   = MeasureTextWidth(_dwriteFactory.get(), _smallFormat.get(), label, textMaxW, lineH);
                        const float labelGap = DipsToPixels(6.0f, _dpi);
                        const float pathLeft = textX + labelW + labelGap;
                        const float pathW    = std::max(0.0f, contentRight - pathLeft);

                        const D2D1_RECT_F labelRc = D2D1::RectF(textX, textY, textX + labelW, textY + lineH);
                        _target->DrawTextW(
                            label.data(), static_cast<UINT32>(label.size()), _smallFormat.get(), labelRc, _subTextBrush.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);

                        const std::wstring pathText = TruncatePathMiddleToWidth(_dwriteFactory.get(), _bodyFormat.get(), path.native(), pathW, lineH);
                        const D2D1_RECT_F pathRc    = D2D1::RectF(pathLeft, textY, contentRight, textY + lineH);
                        _target->DrawTextW(
                            pathText.data(), static_cast<UINT32>(pathText.size()), _bodyFormat.get(), pathRc, _textBrush.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);

                        textY += lineH;
                    };

                    if (info.kind == FolderWindow::InformationalTaskUpdate::Kind::CompareDirectories)
                    {
                        drawLabeledPathLine(IDS_PREFS_PANES_HEADER_LEFT, info.leftRoot);
                        drawLabeledPathLine(IDS_PREFS_PANES_HEADER_RIGHT, info.rightRoot);

                        if (_smallFormat && _textBrush && (info.scanActive || info.scanFolderCount > 0 || info.scanEntryCount > 0))
                        {
                            const std::wstring scanPath = info.scanCurrentRelative.empty() ? std::wstring(L".") : info.scanCurrentRelative.native();
                            const std::wstring scanText =
                                FormatStringResource(nullptr, IDS_FMT_COMPARE_SCAN_STATUS, scanPath, info.scanFolderCount, info.scanEntryCount);
                            const D2D1_RECT_F scanRc = D2D1::RectF(textX, textY, contentRight, textY + lineH);
                            _target->DrawTextW(scanText.data(),
                                               static_cast<UINT32>(scanText.size()),
                                               _smallFormat.get(),
                                               scanRc,
                                               _textBrush.get(),
                                               D2D1_DRAW_TEXT_OPTIONS_CLIP);
                            textY += lineH;
                        }

                        if (_smallFormat && _subTextBrush && ! info.finished && info.scanElapsedSeconds.has_value())
                        {
                            if (textY + lineH <= cardRect.bottom)
                            {
                                const std::wstring duration = FormatDurationHms(info.scanElapsedSeconds.value());
                                if (! duration.empty())
                                {
                                    const std::wstring elapsedText = FormatStringResource(nullptr, IDS_FMT_COMPARE_ELAPSED, duration);
                                    const D2D1_RECT_F elapsedRc    = D2D1::RectF(textX, textY, contentRight, textY + lineH);
                                    _target->DrawTextW(elapsedText.data(),
                                                       static_cast<UINT32>(elapsedText.size()),
                                                       _smallFormat.get(),
                                                       elapsedRc,
                                                       _subTextBrush.get(),
                                                       D2D1_DRAW_TEXT_OPTIONS_CLIP);
                                    textY += lineH;
                                }
                            }
                        }

                        if (_smallFormat && _subTextBrush && (info.scanCandidateFileCount > 0 || info.scanCandidateTotalBytes > 0))
                        {
                            if (textY + lineH <= cardRect.bottom)
                            {
                                const std::wstring totalBytes = FormatBytesCompact(info.scanCandidateTotalBytes);
                                const std::wstring candidateText =
                                    FormatStringResource(nullptr, IDS_FMT_COMPARE_SCAN_CANDIDATES_STATUS, info.scanCandidateFileCount, totalBytes);
                                const D2D1_RECT_F candidatesRc = D2D1::RectF(textX, textY, contentRight, textY + lineH);
                                _target->DrawTextW(candidateText.data(),
                                                   static_cast<UINT32>(candidateText.size()),
                                                   _smallFormat.get(),
                                                   candidatesRc,
                                                   _subTextBrush.get(),
                                                   D2D1_DRAW_TEXT_OPTIONS_CLIP);
                                textY += lineH;
                            }
                        }
                    }
                    else if (info.kind == FolderWindow::InformationalTaskUpdate::Kind::ChangeCase)
                    {
                        if (! info.changeCaseCurrentPath.empty())
                        {
                            drawLabeledPathLine(IDS_FILEOPS_LABEL_FROM, info.changeCaseCurrentPath);
                        }

                        if (_smallFormat && _textBrush &&
                            (info.changeCaseEnumerating || info.changeCaseScannedFolders > 0 || info.changeCaseScannedEntries > 0))
                        {
                            const std::wstring scanPath = info.changeCaseCurrentPath.empty() ? std::wstring(L".") : info.changeCaseCurrentPath.native();
                            const std::wstring scanText = FormatStringResource(
                                nullptr, IDS_FMT_COMPARE_SCAN_STATUS, scanPath, info.changeCaseScannedFolders, info.changeCaseScannedEntries);
                            const D2D1_RECT_F scanRc = D2D1::RectF(textX, textY, contentRight, textY + lineH);
                            _target->DrawTextW(scanText.data(),
                                               static_cast<UINT32>(scanText.size()),
                                               _smallFormat.get(),
                                               scanRc,
                                               _textBrush.get(),
                                               D2D1_DRAW_TEXT_OPTIONS_CLIP);
                            textY += lineH;
                        }

                        if (_smallFormat && _subTextBrush &&
                            (info.changeCaseRenaming || info.changeCasePlannedRenames > 0 || info.changeCaseCompletedRenames > 0))
                        {
                            if (textY + lineH <= cardRect.bottom)
                            {
                                const std::wstring countsText =
                                    info.changeCasePlannedRenames > 0
                                        ? FormatStringResource(
                                              nullptr, IDS_FMT_FILEOPS_OP_COUNTS, info.title, info.changeCaseCompletedRenames, info.changeCasePlannedRenames)
                                        : FormatStringResource(nullptr, IDS_FMT_FILEOPS_OP_COUNTS_UNKNOWN_TOTAL, info.title, info.changeCaseCompletedRenames);
                                const D2D1_RECT_F countsRc = D2D1::RectF(textX, textY, contentRight, textY + lineH);
                                _target->DrawTextW(countsText.data(),
                                                   static_cast<UINT32>(countsText.size()),
                                                   _smallFormat.get(),
                                                   countsRc,
                                                   _subTextBrush.get(),
                                                   D2D1_DRAW_TEXT_OPTIONS_CLIP);
                                textY += lineH;
                            }
                        }
                    }
                    else if (info.kind == FolderWindow::InformationalTaskUpdate::Kind::ChangeAttributes)
                    {
                        if (! info.changeAttributesCurrentPath.empty())
                        {
                            drawLabeledPathLine(IDS_FILEOPS_LABEL_FROM, info.changeAttributesCurrentPath);
                        }

                        if (_smallFormat && _textBrush &&
                            (info.changeAttributesEnumerating || info.changeAttributesScannedFolders > 0 ||
                             info.changeAttributesScannedEntries > 0))
                        {
                            const std::wstring scanPath =
                                info.changeAttributesCurrentPath.empty() ? std::wstring(L".") : info.changeAttributesCurrentPath.native();
                            const std::wstring scanText = FormatStringResource(
                                nullptr, IDS_FMT_COMPARE_SCAN_STATUS, scanPath, info.changeAttributesScannedFolders, info.changeAttributesScannedEntries);
                            const D2D1_RECT_F scanRc = D2D1::RectF(textX, textY, contentRight, textY + lineH);
                            _target->DrawTextW(scanText.data(),
                                               static_cast<UINT32>(scanText.size()),
                                               _smallFormat.get(),
                                               scanRc,
                                               _textBrush.get(),
                                               D2D1_DRAW_TEXT_OPTIONS_CLIP);
                            textY += lineH;
                        }

                        if (_smallFormat && _subTextBrush &&
                            (info.changeAttributesApplying || info.changeAttributesPlannedItems > 0 ||
                             info.changeAttributesCompletedItems > 0))
                        {
                            if (textY + lineH <= cardRect.bottom)
                            {
                                const std::wstring countsText =
                                    info.changeAttributesPlannedItems > 0
                                        ? FormatStringResource(nullptr,
                                                               IDS_FMT_FILEOPS_OP_COUNTS,
                                                               info.title,
                                                               info.changeAttributesCompletedItems,
                                                               info.changeAttributesPlannedItems)
                                        : FormatStringResource(
                                              nullptr, IDS_FMT_FILEOPS_OP_COUNTS_UNKNOWN_TOTAL, info.title, info.changeAttributesCompletedItems);
                                const D2D1_RECT_F countsRc = D2D1::RectF(textX, textY, contentRight, textY + lineH);
                                _target->DrawTextW(countsText.data(),
                                                   static_cast<UINT32>(countsText.size()),
                                                   _smallFormat.get(),
                                                   countsRc,
                                                   _subTextBrush.get(),
                                                   D2D1_DRAW_TEXT_OPTIONS_CLIP);
                                textY += lineH;
                            }
                        }
                    }

                    if (_smallFormat && _textBrush && info.contentActive)
                    {
                        std::array<size_t, FolderWindow::InformationalTaskUpdate::kMaxContentInFlightFiles> activeInFlightIndices{};
                        size_t activeInFlightCount = 0;
                        for (size_t i = 0; i < info.contentInFlightCount && activeInFlightCount < activeInFlightIndices.size(); ++i)
                        {
                            const auto& entry          = info.contentInFlight[i];
                            const bool active          = entry.totalBytes == 0 || entry.completedBytes < entry.totalBytes;
                            const bool recentCompleted = ! active && entry.totalBytes > 0 && entry.completedBytes >= entry.totalBytes &&
                                                         entry.lastUpdateTick != 0 && nowTick >= entry.lastUpdateTick &&
                                                         (nowTick - entry.lastUpdateTick) <= kCompletedInFlightGraceMs;
                            if (active || recentCompleted)
                            {
                                activeInFlightIndices[activeInFlightCount] = i;
                                ++activeInFlightCount;
                            }
                        }

                        if (activeInFlightCount == 0u)
                        {
                            const std::wstring contentPath = info.contentCurrentRelative.empty() ? std::wstring{} : info.contentCurrentRelative.native();
                            const std::wstring bytesRead   = FormatBytesCompact(info.contentCurrentCompletedBytes);
                            std::wstring contentText;
                            if (info.contentCurrentTotalBytes > 0)
                            {
                                const std::wstring bytesTotal = FormatBytesCompact(info.contentCurrentTotalBytes);
                                contentText = FormatStringResource(nullptr, IDS_FMT_COMPARE_CONTENT_STATUS, contentPath, bytesRead, bytesTotal);
                            }
                            else
                            {
                                contentText = FormatStringResource(nullptr, IDS_FMT_COMPARE_CONTENT_STATUS_UNKNOWN, contentPath, bytesRead);
                            }

                            const D2D1_RECT_F contentRc = D2D1::RectF(textX, textY, contentRight, textY + lineH);
                            _target->DrawTextW(contentText.data(),
                                               static_cast<UINT32>(contentText.size()),
                                               _smallFormat.get(),
                                               contentRc,
                                               _textBrush.get(),
                                               D2D1_DRAW_TEXT_OPTIONS_CLIP);
                            textY += lineH;
                        }
                        else
                        {
                            const float rightEdge       = textX + textMaxW;
                            const float miniBarGap      = DipsToPixels(8.0f, _dpi);
                            const float miniBarWDesired = DipsToPixels(92.0f, _dpi);
                            const float miniBarH        = DipsToPixels(6.0f, _dpi);

                            for (size_t i = 0; i < activeInFlightCount; ++i)
                            {
                                if (textY + lineH > cardRect.bottom)
                                {
                                    break;
                                }

                                const auto& entry = info.contentInFlight[activeInFlightIndices[i]];

                                const std::wstring_view sourcePathText = entry.relativePath.native();
                                const uint64_t fileTotalBytes          = entry.totalBytes;
                                const uint64_t fileCompletedBytes      = entry.completedBytes;

                                const float availableW     = std::max(0.0f, rightEdge - textX);
                                const float miniBarWMin    = DipsToPixels(40.0f, _dpi);
                                const float minTextW       = DipsToPixels(48.0f, _dpi);
                                float miniBarW             = std::min(miniBarWDesired, availableW);
                                const float maxBarWithText = std::max(0.0f, availableW - miniBarGap - minTextW);
                                if (maxBarWithText > 0.0f)
                                {
                                    miniBarW = std::clamp(miniBarW, std::min(miniBarWMin, maxBarWithText), maxBarWithText);
                                }

                                if (fileTotalBytes > 0u && fileCompletedBytes >= fileTotalBytes)
                                {
                                    miniBarW = 0.0f;
                                }

                                const float barRight  = rightEdge;
                                const float barLeft   = barRight - miniBarW;
                                const float pathRight = (miniBarW > 0.0f) ? std::max(textX, barLeft - miniBarGap) : rightEdge;
                                const float pathW     = std::max(0.0f, pathRight - textX);

                                const std::wstring fromPath = TruncatePathMiddleToWidth(_dwriteFactory.get(), _bodyFormat.get(), sourcePathText, pathW, lineH);
                                const D2D1_RECT_F pathRc    = D2D1::RectF(textX, textY, textX + pathW, textY + lineH);
                                _target->DrawTextW(fromPath.data(),
                                                   static_cast<UINT32>(fromPath.size()),
                                                   _bodyFormat.get(),
                                                   pathRc,
                                                   _textBrush.get(),
                                                   D2D1_DRAW_TEXT_OPTIONS_CLIP);

                                if (miniBarW > 0.0f && _progressBgBrush && _progressItemBrush)
                                {
                                    const float barTop          = textY + (lineH - miniBarH) * 0.5f;
                                    const D2D1_RECT_F miniBarRc = D2D1::RectF(barLeft, barTop, barRight, barTop + miniBarH);

                                    const float radiusTrack = ClampCornerRadius(miniBarRc, DipsToPixels(2.0f, _dpi));
                                    _target->FillRoundedRectangle(D2D1::RoundedRect(miniBarRc, radiusTrack, radiusTrack), _progressBgBrush.get());

                                    const bool hasTotal = fileTotalBytes > 0u;
                                    const float frac =
                                        hasTotal && fileCompletedBytes <= fileTotalBytes
                                            ? Clamp01(static_cast<float>(static_cast<double>(fileCompletedBytes) / static_cast<double>(fileTotalBytes)))
                                            : 0.0f;

                                    if (appTheme.menu.rainbowMode)
                                    {
                                        const D2D1::ColorF rainbow = RainbowProgressColor(appTheme, sourcePathText);
                                        _progressItemBrush->SetColor(rainbow);
                                    }
                                    else
                                    {
                                        _progressItemBrush->SetColor(_progressItemBaseColor);
                                    }

                                    const D2D1_RECT_F fill =
                                        hasTotal
                                            ? D2D1::RectF(
                                                  miniBarRc.left, miniBarRc.top, miniBarRc.left + (miniBarRc.right - miniBarRc.left) * frac, miniBarRc.bottom)
                                            : ComputeIndeterminateBarFill(miniBarRc, nowTick);
                                    const float radiusFill = ClampCornerRadius(fill, DipsToPixels(2.0f, _dpi));
                                    _target->FillRoundedRectangle(D2D1::RoundedRect(fill, radiusFill, radiusFill), _progressItemBrush.get());
                                }

                                textY += lineH;
                            }
                        }
                    }

                    if (_smallFormat && _subTextBrush && (info.contentPendingCount > 0 || info.contentCompletedCount > 0))
                    {
                        if (textY + lineH <= cardRect.bottom)
                        {
                            const std::wstring countsText =
                                FormatStringResource(nullptr, IDS_FMT_COMPARE_CONTENT_COUNTS_STATUS, info.contentPendingCount, info.contentCompletedCount);
                            const D2D1_RECT_F countsRc = D2D1::RectF(textX, textY, contentRight, textY + lineH);
                            _target->DrawTextW(countsText.data(),
                                               static_cast<UINT32>(countsText.size()),
                                               _smallFormat.get(),
                                               countsRc,
                                               _subTextBrush.get(),
                                               D2D1_DRAW_TEXT_OPTIONS_CLIP);
                            textY += lineH;
                        }
                    }

                    if (_smallFormat && _subTextBrush && info.contentTotalBytes > 0)
                    {
                        if (textY + lineH <= cardRect.bottom)
                        {
                            const std::wstring completedBytes = FormatBytesCompact(info.contentCompletedBytes);
                            const std::wstring totalBytes     = FormatBytesCompact(info.contentTotalBytes);
                            const std::wstring bytesText =
                                FormatStringResource(nullptr, IDS_FMT_COMPARE_CONTENT_TOTAL_BYTES_STATUS, completedBytes, totalBytes);
                            const D2D1_RECT_F bytesRc = D2D1::RectF(textX, textY, contentRight, textY + lineH);
                            _target->DrawTextW(bytesText.data(),
                                               static_cast<UINT32>(bytesText.size()),
                                               _smallFormat.get(),
                                               bytesRc,
                                               _subTextBrush.get(),
                                               D2D1_DRAW_TEXT_OPTIONS_CLIP);
                            textY += lineH;
                        }
                    }

                    if (_smallFormat && _subTextBrush && ! info.finished && info.contentEtaSeconds.has_value())
                    {
                        if (textY + lineH <= cardRect.bottom)
                        {
                            const std::wstring duration = FormatDurationHms(info.contentEtaSeconds.value());
                            if (! duration.empty())
                            {
                                const std::wstring etaText = FormatStringResource(nullptr, IDS_FMT_COMPARE_ETA, duration);
                                const D2D1_RECT_F etaRc    = D2D1::RectF(textX, textY, contentRight, textY + lineH);
                                _target->DrawTextW(etaText.data(),
                                                   static_cast<UINT32>(etaText.size()),
                                                   _smallFormat.get(),
                                                   etaRc,
                                                   _subTextBrush.get(),
                                                   D2D1_DRAW_TEXT_OPTIONS_CLIP);
                                textY += lineH;
                            }
                        }
                    }

                    if (_smallFormat && _subTextBrush && info.finished && ! info.doneSummary.empty())
                    {
                        const D2D1_RECT_F doneRc = D2D1::RectF(textX, textY, contentRight, textY + lineH);
                        _target->DrawTextW(info.doneSummary.data(),
                                           static_cast<UINT32>(info.doneSummary.size()),
                                           _smallFormat.get(),
                                           doneRc,
                                           _subTextBrush.get(),
                                           D2D1_DRAW_TEXT_OPTIONS_CLIP);
                        textY += lineH;
                    }

                    const bool showProgressBar =
                        ! info.finished &&
                        ((info.kind == FolderWindow::InformationalTaskUpdate::Kind::CompareDirectories && (info.scanActive || info.contentActive)) ||
                         (info.kind == FolderWindow::InformationalTaskUpdate::Kind::ChangeCase && (info.changeCaseEnumerating || info.changeCaseRenaming)) ||
                         (info.kind == FolderWindow::InformationalTaskUpdate::Kind::ChangeAttributes &&
                          (info.changeAttributesEnumerating || info.changeAttributesApplying)));
                    if (showProgressBar)
                    {
                        const float barH        = DipsToPixels(8.0f, _dpi);
                        const float bottomPad   = DipsToPixels(10.0f, _dpi);
                        const float barBottom   = cardRect.bottom - bottomPad;
                        const float barTop      = barBottom - barH;
                        const D2D1_RECT_F barRc = D2D1::RectF(textX, barTop, contentRight, barBottom);

                        if (_progressBgBrush)
                        {
                            const float radius = ClampCornerRadius(barRc, DipsToPixels(2.0f, _dpi));
                            _target->FillRoundedRectangle(D2D1::RoundedRect(barRc, radius, radius), _progressBgBrush.get());
                        }

                        if (_progressGlobalBrush)
                        {
                            bool hasTotal = false;
                            float frac    = 0.0f;
                            if (info.kind == FolderWindow::InformationalTaskUpdate::Kind::CompareDirectories)
                            {
                                hasTotal = info.contentTotalBytes > 0 && info.contentCompletedBytes <= info.contentTotalBytes;
                                frac     = hasTotal ? Clamp01(static_cast<float>(static_cast<double>(info.contentCompletedBytes) /
                                                                                 static_cast<double>(info.contentTotalBytes)))
                                                    : 0.0f;
                            }
                            else if (info.kind == FolderWindow::InformationalTaskUpdate::Kind::ChangeCase)
                            {
                                hasTotal = info.changeCasePlannedRenames > 0 && info.changeCaseCompletedRenames <= info.changeCasePlannedRenames;
                                frac     = hasTotal ? Clamp01(static_cast<float>(static_cast<double>(info.changeCaseCompletedRenames) /
                                                                                 static_cast<double>(info.changeCasePlannedRenames)))
                                                    : 0.0f;
                            }
                            else if (info.kind == FolderWindow::InformationalTaskUpdate::Kind::ChangeAttributes)
                            {
                                hasTotal = info.changeAttributesPlannedItems > 0 &&
                                           info.changeAttributesCompletedItems <= info.changeAttributesPlannedItems;
                                frac     = hasTotal ? Clamp01(static_cast<float>(static_cast<double>(info.changeAttributesCompletedItems) /
                                                                                 static_cast<double>(info.changeAttributesPlannedItems)))
                                                    : 0.0f;
                            }

                            const D2D1_RECT_F fill = hasTotal ? D2D1::RectF(barRc.left, barRc.top, barRc.left + (barRc.right - barRc.left) * frac, barRc.bottom)
                                                              : ComputeIndeterminateBarFill(barRc, nowTick);
                            const float radius     = ClampCornerRadius(fill, DipsToPixels(2.0f, _dpi));
                            _target->FillRoundedRectangle(D2D1::RoundedRect(fill, radius, radius), _progressGlobalBrush.get());
                        }
                    }

                    if (info.finished)
                    {
                        const float dismissButtonH         = DipsToPixels(24.0f, _dpi);
                        const float dismissButtonBottomPad = DipsToPixels(8.0f, _dpi);
                        const float dismissButtonTop       = cardRect.bottom - dismissButtonBottomPad - dismissButtonH;

                        PopupButton dismissBtn{};
                        dismissBtn.bounds     = D2D1::RectF(textX, dismissButtonTop, contentRight, dismissButtonTop + dismissButtonH);
                        dismissBtn.hit.kind   = PopupHitTest::Kind::TaskDismiss;
                        dismissBtn.hit.taskId = task.taskId;
                        _buttons.push_back(dismissBtn);
                        DrawButton(dismissBtn, _buttonSmallFormat.get(), LoadStringResource(nullptr, IDS_FILEOP_BTN_DISMISS));
                    }

                    const float gapAfter = (taskIndex + 1u < taskCount) ? cardGap : 0.0f;
                    y += taskCardH + gapAfter;
                    continue;
                }

                const UINT pauseId           = task.paused ? static_cast<UINT>(IDS_FILEOP_BTN_RESUME) : static_cast<UINT>(IDS_FILEOP_BTN_PAUSE);
                const std::wstring pauseText = LoadStringResource(nullptr, pauseId);

                const std::wstring cancelText = LoadStringResource(nullptr, IDS_FILEOP_BTN_CANCEL);

                const bool showCopyMoveControls = task.operation == FILESYSTEM_COPY || task.operation == FILESYSTEM_MOVE;
                std::wstring speedLimitText;
                if (showCopyMoveControls)
                {
                    if (task.desiredSpeedLimitBytesPerSecond == 0)
                    {
                        speedLimitText = LoadStringResource(nullptr, IDS_FILEOP_SPEED_LIMIT_BUTTON_UNLIMITED);
                    }
                    else
                    {
                        speedLimitText =
                            FormatStringResource(nullptr, IDS_FMT_FILEOP_SPEED_LIMIT_BUTTON_BYTES, FormatBytesCompact(task.desiredSpeedLimitBytesPerSecond));
                    }
                }

                const UINT opTextId = [&]() -> UINT
                {
                    switch (task.operation)
                    {
                        case FILESYSTEM_COPY: return static_cast<UINT>(IDS_FILEOP_OPERATION_COPY);
                        case FILESYSTEM_MOVE: return static_cast<UINT>(IDS_FILEOP_OPERATION_MOVE);
                        case FILESYSTEM_DELETE: return static_cast<UINT>(IDS_FILEOP_OPERATION_DELETE);
                        case FILESYSTEM_RENAME: return static_cast<UINT>(IDS_FILEOP_OPERATION_RENAME);
                    }
                    return static_cast<UINT>(IDS_FILEOP_OPERATION_COPY);
                }();

                const std::wstring opText = LoadStringResource(nullptr, opTextId);
                const ULONGLONG nowTick   = renderTick;

                // Build header text - show calculating status or operation progress
                std::wstring headerText;
                if (task.finished)
                {
                    const HRESULT partialHr   = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
                    const HRESULT cancelledHr = HRESULT_FROM_WIN32(ERROR_CANCELLED);
                    std::wstring statusText;
                    if (SUCCEEDED(task.resultHr))
                    {
                        statusText = LoadStringResource(nullptr, IDS_FILEOPS_STATUS_COMPLETED);
                    }
                    else if (task.resultHr == cancelledHr || task.resultHr == E_ABORT)
                    {
                        statusText = LoadStringResource(nullptr, IDS_FILEOPS_STATUS_CANCELED);
                    }
                    else if (task.resultHr == partialHr)
                    {
                        statusText = LoadStringResource(nullptr, IDS_FILEOPS_STATUS_PARTIAL);
                    }
                    else
                    {
                        statusText = FormatStringResource(nullptr, IDS_FMT_FILEOPS_STATUS_FAILED, static_cast<unsigned long>(task.resultHr));
                    }
                    headerText = FormatStringResource(nullptr, IDS_FMT_FILEOPS_OP_STATUS, opText, statusText);
                }
                else
                {
                    const bool isWaiting = task.queuePaused || task.waitingInQueue;
                    if (isWaiting)
                    {
                        headerText = FormatStringResource(nullptr, IDS_FMT_FILEOPS_OP_STATUS, opText, LoadStringResource(nullptr, IDS_FILEOPS_GRAPH_WAITING));
                    }
                    else if (task.paused)
                    {
                        headerText = FormatStringResource(nullptr, IDS_FMT_FILEOPS_OP_STATUS, opText, LoadStringResource(nullptr, IDS_FILEOPS_GRAPH_PAUSED));
                    }
                    else if (task.preCalcInProgress)
                    {
                        // Show "Calculating... (Xs elapsed)"
                        const uint64_t elapsedSec   = task.preCalcElapsedMs / 1000;
                        const std::wstring calcText = elapsedSec > 0
                                                          ? FormatStringResource(nullptr, IDS_FMT_FILEOPS_CALCULATING_TIME, FormatDurationHms(elapsedSec))
                                                          : LoadStringResource(nullptr, IDS_FILEOPS_CALCULATING);
                        headerText                  = FormatStringResource(nullptr, IDS_FMT_FILEOPS_OP_STATUS, opText, calcText);
                    }
                    else
                    {
                        const bool hasProgressNumbers = task.completedItems > 0 || task.completedBytes > 0 || task.totalItems > 0 || task.totalBytes > 0;
                        const bool showPreparing =
                            ! task.started || ! task.hasProgressCallbacks || (task.operation == FILESYSTEM_DELETE && ! hasProgressNumbers);
                        if (showPreparing)
                        {
                            const ULONGLONG opStartTick = task.operationStartTick;
                            const uint64_t elapsedSec =
                                (opStartTick > 0 && nowTick >= opStartTick) ? static_cast<uint64_t>((nowTick - opStartTick) / 1000ull) : 0ull;
                            const std::wstring prepText = elapsedSec > 0
                                                              ? FormatStringResource(nullptr, IDS_FMT_FILEOPS_PREPARING_TIME, FormatDurationHms(elapsedSec))
                                                              : LoadStringResource(nullptr, IDS_FILEOPS_PREPARING);
                            headerText                  = FormatStringResource(nullptr, IDS_FMT_FILEOPS_OP_STATUS, opText, prepText);
                        }
                        else if (task.totalItems > 0)
                        {
                            headerText = FormatStringResource(nullptr, IDS_FMT_FILEOPS_OP_COUNTS, opText, task.completedItems, task.totalItems);
                        }
                        else
                        {
                            headerText = FormatStringResource(nullptr, IDS_FMT_FILEOPS_OP_COUNTS_UNKNOWN_TOTAL, opText, task.completedItems);
                        }
                    }
                }

                const float collapseBtnSize = DipsToPixels(18.0f, _dpi);
                const float collapseBtnGap  = DipsToPixels(6.0f, _dpi);

                const float headerTop    = isCollapsedTask ? cardRect.top + (taskCardH - lineH) * 0.5f : textY;
                const float headerBottom = headerTop + lineH;
                const float collapseTop  = headerTop + (lineH - collapseBtnSize) * 0.5f;
                const float collapseLeft = std::max(textX, contentRight - collapseBtnSize);

                PopupButton collapseBtn{};
                collapseBtn.bounds     = D2D1::RectF(collapseLeft, collapseTop, contentRight, collapseTop + collapseBtnSize);
                collapseBtn.hit.kind   = PopupHitTest::Kind::TaskToggleCollapse;
                collapseBtn.hit.taskId = task.taskId;
                _buttons.push_back(collapseBtn);
                DrawButton(collapseBtn, nullptr, {});
                DrawCollapseChevron(collapseBtn.bounds, isCollapsedTask);

                const float headerRight = std::max(textX, collapseBtn.bounds.left - collapseBtnGap);
                float headerLeft        = textX;

                CaptionStatus statusIcon = CaptionStatus::None;
                {
                    const HRESULT partialHr   = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
                    const HRESULT cancelledHr = HRESULT_FROM_WIN32(ERROR_CANCELLED);
                    if (task.errorCount > 0 ||
                        (task.finished && FAILED(task.resultHr) && task.resultHr != partialHr && task.resultHr != cancelledHr && task.resultHr != E_ABORT))
                    {
                        statusIcon = CaptionStatus::Error;
                    }
                    else if (task.warningCount > 0 || (task.finished && task.resultHr == partialHr))
                    {
                        statusIcon = CaptionStatus::Warning;
                    }
                    else if (task.finished && SUCCEEDED(task.resultHr))
                    {
                        statusIcon = CaptionStatus::Ok;
                    }
                }

                if (statusIcon != CaptionStatus::None && _target)
                {
                    const float iconSize = DipsToPixels(16.0f, _dpi);
                    const float iconGap  = DipsToPixels(6.0f, _dpi);

                    D2D1_RECT_F iconRc = D2D1::RectF(textX, headerTop, textX + iconSize, headerBottom);
                    iconRc.right       = std::min(iconRc.right, headerRight);

                    wchar_t fluentGlyph = 0;
                    wchar_t fallback    = 0;
                    ID2D1Brush* brush   = _textBrush.get();
                    switch (statusIcon)
                    {
                        case CaptionStatus::Ok:
                            fluentGlyph = FluentIcons::kCheckMark;
                            fallback    = FluentIcons::kFallbackCheckMark;
                            brush       = _statusOkBrush ? _statusOkBrush.get() : (_textBrush ? _textBrush.get() : nullptr);
                            break;
                        case CaptionStatus::Warning:
                            fluentGlyph = FluentIcons::kWarning;
                            fallback    = FluentIcons::kFallbackWarning;
                            brush       = _statusWarningBrush ? _statusWarningBrush.get() : (_textBrush ? _textBrush.get() : nullptr);
                            break;
                        case CaptionStatus::Error:
                            fluentGlyph = FluentIcons::kError;
                            fallback    = FluentIcons::kFallbackError;
                            brush       = _statusErrorBrush ? _statusErrorBrush.get() : (_textBrush ? _textBrush.get() : nullptr);
                            break;
                        case CaptionStatus::None:
                        default: break;
                    }

                    const bool useFluentFormat = _statusIconFormat != nullptr && fluentGlyph != 0;
                    const wchar_t glyph        = useFluentFormat ? fluentGlyph : fallback;
                    IDWriteTextFormat* format  = useFluentFormat ? _statusIconFormat.get() : _statusIconFallbackFormat.get();

                    if (format && brush && glyph != 0 && iconRc.right > iconRc.left)
                    {
                        const wchar_t text[2]{glyph, 0};
                        _target->DrawTextW(text, 1u, format, iconRc, brush, D2D1_DRAW_TEXT_OPTIONS_CLIP);
                        headerLeft = std::min(headerRight, iconRc.right + iconGap);
                    }
                }

                if (_headerFormat)
                {
                    const D2D1_RECT_F headerRc = D2D1::RectF(headerLeft, headerTop, headerRight, headerBottom);
                    _target->DrawTextW(headerText.data(),
                                       static_cast<UINT32>(headerText.size()),
                                       _headerFormat.get(),
                                       headerRc,
                                       _textBrush.get(),
                                       D2D1_DRAW_TEXT_OPTIONS_CLIP);
                }

                if (isCollapsedTask)
                {
                    const float gapAfter = (taskIndex + 1u < taskCount) ? cardGap : 0.0f;
                    y += taskCardH + gapAfter;
                    continue;
                }

                textY = headerBottom;

                if (task.finished)
                {
                    const HRESULT partialHr = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
                    const bool showHrLine   = FAILED(task.resultHr) && task.resultHr != partialHr;

                    const std::wstring diagCounts = FormatStringResource(nullptr, IDS_FMT_FILEOPS_WARNINGS_ERRORS, task.warningCount, task.errorCount);
                    const D2D1_RECT_F countsRc    = D2D1::RectF(textX, textY, textX + textMaxW, textY + lineH);
                    _target->DrawTextW(diagCounts.data(), static_cast<UINT32>(diagCounts.size()), _bodyFormat.get(), countsRc, _subTextBrush.get());
                    textY += lineH;

                    if (showHrLine)
                    {
                        const std::wstring hrText = FormatStringResource(nullptr, IDS_FMT_FILEOPS_RESULT_HRESULT, static_cast<unsigned long>(task.resultHr));
                        const D2D1_RECT_F hrRc    = D2D1::RectF(textX, textY, textX + textMaxW, textY + lineH);
                        _target->DrawTextW(hrText.data(), static_cast<UINT32>(hrText.size()), _bodyFormat.get(), hrRc, _subTextBrush.get());
                        textY += lineH;
                    }

                    const float labelWDesired   = DipsToPixels(56.0f, _dpi);
                    const float labelGapDesired = DipsToPixels(6.0f, _dpi);
                    const float labelW          = std::min(labelWDesired, textMaxW);
                    const float labelGap        = (labelW < textMaxW) ? std::min(labelGapDesired, textMaxW - labelW) : 0.0f;
                    const float pathW           = std::max(0.0f, textMaxW - labelW - labelGap);

                    if (task.operation == FILESYSTEM_DELETE)
                    {
                        const std::wstring label  = LoadStringResource(nullptr, IDS_FILEOPS_LABEL_DELETING);
                        const D2D1_RECT_F labelRc = D2D1::RectF(textX, textY, textX + labelW, textY + lineH);
                        _target->DrawTextW(label.data(), static_cast<UINT32>(label.size()), _smallFormat.get(), labelRc, _subTextBrush.get());

                        const std::wstring path  = TruncatePathMiddleToWidth(_dwriteFactory.get(), _bodyFormat.get(), task.currentSourcePath, pathW, lineH);
                        const D2D1_RECT_F pathRc = D2D1::RectF(textX + labelW + labelGap, textY, textX + labelW + labelGap + pathW, textY + lineH);
                        _target->DrawTextW(
                            path.data(), static_cast<UINT32>(path.size()), _bodyFormat.get(), pathRc, _textBrush.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);
                        textY += lineH;
                    }
                    else
                    {
                        const std::wstring fromLabel  = LoadStringResource(nullptr, IDS_FILEOPS_LABEL_FROM);
                        const D2D1_RECT_F fromLabelRc = D2D1::RectF(textX, textY, textX + labelW, textY + lineH);
                        _target->DrawTextW(fromLabel.data(), static_cast<UINT32>(fromLabel.size()), _smallFormat.get(), fromLabelRc, _subTextBrush.get());

                        const std::wstring fromPath  = TruncatePathMiddleToWidth(_dwriteFactory.get(), _bodyFormat.get(), task.currentSourcePath, pathW, lineH);
                        const D2D1_RECT_F fromPathRc = D2D1::RectF(textX + labelW + labelGap, textY, textX + labelW + labelGap + pathW, textY + lineH);
                        _target->DrawTextW(fromPath.data(),
                                           static_cast<UINT32>(fromPath.size()),
                                           _bodyFormat.get(),
                                           fromPathRc,
                                           _textBrush.get(),
                                           D2D1_DRAW_TEXT_OPTIONS_CLIP);
                        textY += lineH;

                        const std::wstring toLabel  = LoadStringResource(nullptr, IDS_FILEOPS_LABEL_TO);
                        const D2D1_RECT_F toLabelRc = D2D1::RectF(textX, textY, textX + labelW, textY + lineH);
                        _target->DrawTextW(toLabel.data(), static_cast<UINT32>(toLabel.size()), _smallFormat.get(), toLabelRc, _subTextBrush.get());

                        const std::wstring toPath =
                            TruncatePathMiddleToWidth(_dwriteFactory.get(), _bodyFormat.get(), task.currentDestinationPath, pathW, lineH);
                        const D2D1_RECT_F toPathRc = D2D1::RectF(textX + labelW + labelGap, textY, textX + labelW + labelGap + pathW, textY + lineH);
                        _target->DrawTextW(
                            toPath.data(), static_cast<UINT32>(toPath.size()), _bodyFormat.get(), toPathRc, _textBrush.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);
                        textY += lineH;
                    }

                    if (! task.lastDiagnosticMessage.empty())
                    {
                        const std::wstring diagText = FormatStringResource(nullptr, IDS_FMT_FILEOPS_LAST_NOTE, task.lastDiagnosticMessage);
                        const D2D1_RECT_F diagRc    = D2D1::RectF(textX, textY, textX + textMaxW, textY + lineH);
                        _target->DrawTextW(diagText.data(), static_cast<UINT32>(diagText.size()), _smallFormat.get(), diagRc, _subTextBrush.get());
                        textY += lineH;
                    }

                    const float dismissButtonH         = DipsToPixels(24.0f, _dpi);
                    const float dismissButtonBottomPad = DipsToPixels(8.0f, _dpi);
                    const float dismissButtonTop       = std::max(textY + DipsToPixels(4.0f, _dpi), cardRect.bottom - dismissButtonBottomPad - dismissButtonH);

                    const float progressBarH         = DipsToPixels(8.0f, _dpi);
                    const float progressBarBottomPad = DipsToPixels(6.0f, _dpi);
                    const float progressBarTop       = std::max(textY + DipsToPixels(2.0f, _dpi), dismissButtonTop - progressBarBottomPad - progressBarH);
                    const D2D1_RECT_F progressRc     = D2D1::RectF(textX, progressBarTop, contentRight, progressBarTop + progressBarH);

                    if (_progressBgBrush)
                    {
                        const float radius = ClampCornerRadius(progressRc, DipsToPixels(2.0f, _dpi));
                        _target->FillRoundedRectangle(D2D1::RoundedRect(progressRc, radius, radius), _progressBgBrush.get());
                    }

                    const float completeFraction = ComputeFileOperationsTaskCompleteFractionForDisplay(task);

                    if (_progressGlobalBrush)
                    {
                        const D2D1_RECT_F fillRc = D2D1::RectF(
                            progressRc.left, progressRc.top, progressRc.left + (progressRc.right - progressRc.left) * completeFraction, progressRc.bottom);
                        const float radius = ClampCornerRadius(fillRc, DipsToPixels(2.0f, _dpi));
                        _target->FillRoundedRectangle(D2D1::RoundedRect(fillRc, radius, radius), _progressGlobalBrush.get());
                    }

                    const bool hasDiagnosticsActions = task.warningCount > 0 || task.errorCount > 0;
                    if (hasDiagnosticsActions)
                    {
                        const float btnGap = DipsToPixels(6.0f, _dpi);
                        const float totalW = std::max(0.0f, contentRight - textX);
                        const float btnW   = std::max(0.0f, (totalW - btnGap * 2.0f) / 3.0f);

                        PopupButton showLogBtn{};
                        showLogBtn.bounds     = D2D1::RectF(textX, dismissButtonTop, textX + btnW, dismissButtonTop + dismissButtonH);
                        showLogBtn.hit.kind   = PopupHitTest::Kind::TaskShowLog;
                        showLogBtn.hit.taskId = task.taskId;
                        _buttons.push_back(showLogBtn);
                        DrawButton(showLogBtn, _buttonSmallFormat.get(), LoadStringResource(nullptr, IDS_FILEOP_BTN_SHOW_LOG));

                        PopupButton exportIssuesBtn{};
                        exportIssuesBtn.bounds =
                            D2D1::RectF(textX + btnW + btnGap, dismissButtonTop, textX + btnW * 2.0f + btnGap, dismissButtonTop + dismissButtonH);
                        exportIssuesBtn.hit.kind   = PopupHitTest::Kind::TaskExportIssues;
                        exportIssuesBtn.hit.taskId = task.taskId;
                        _buttons.push_back(exportIssuesBtn);
                        DrawButton(exportIssuesBtn, _buttonSmallFormat.get(), LoadStringResource(nullptr, IDS_FILEOP_BTN_EXPORT_ISSUES));

                        PopupButton dismissBtn{};
                        dismissBtn.bounds = D2D1::RectF(textX + btnW * 2.0f + btnGap * 2.0f, dismissButtonTop, contentRight, dismissButtonTop + dismissButtonH);
                        dismissBtn.hit.kind   = PopupHitTest::Kind::TaskDismiss;
                        dismissBtn.hit.taskId = task.taskId;
                        _buttons.push_back(dismissBtn);
                        DrawButton(dismissBtn, _buttonSmallFormat.get(), LoadStringResource(nullptr, IDS_FILEOP_BTN_DISMISS));
                    }
                    else
                    {
                        PopupButton dismissBtn{};
                        dismissBtn.bounds     = D2D1::RectF(textX, dismissButtonTop, contentRight, dismissButtonTop + dismissButtonH);
                        dismissBtn.hit.kind   = PopupHitTest::Kind::TaskDismiss;
                        dismissBtn.hit.taskId = task.taskId;
                        _buttons.push_back(dismissBtn);
                        DrawButton(dismissBtn, _buttonSmallFormat.get(), LoadStringResource(nullptr, IDS_FILEOP_BTN_DISMISS));
                    }

                    const float gapAfter = (taskIndex + 1u < taskCount) ? cardGap : 0.0f;
                    y += taskCardH + gapAfter;
                    continue;
                }

                const AppTheme& theme = folderWindow->GetTheme();

                const RateHistory* history = nullptr;
                const auto historyIt       = _rates.find(task.taskId);
                if (historyIt != _rates.end())
                {
                    history = &historyIt->second;
                }

                // During pre-calculation, show calculating info instead of speed
                if (task.preCalcInProgress)
                {
                    const std::wstring sizeText = FormatBytesCompact(task.preCalcTotalBytes);
                    const uint64_t totalItems   = static_cast<uint64_t>(task.preCalcFileCount) + static_cast<uint64_t>(task.preCalcDirectoryCount);
                    const std::wstring countsText =
                        FormatStringResource(nullptr, IDS_FMT_FILEOPS_FILES_FOLDERS, totalItems, task.preCalcFileCount, task.preCalcDirectoryCount);
                    const D2D1_RECT_F countsRc = D2D1::RectF(textX, textY, textX + textMaxW, textY + lineH);
                    _target->DrawTextW(countsText.data(), static_cast<UINT32>(countsText.size()), _bodyFormat.get(), countsRc, _subTextBrush.get());
                    textY += lineH;

                    const D2D1_RECT_F sizeRc = D2D1::RectF(textX, textY, textX + textMaxW, textY + lineH);
                    _target->DrawTextW(sizeText.data(), static_cast<UINT32>(sizeText.size()), _bodyFormat.get(), sizeRc, _subTextBrush.get());
                    textY += lineH;
                }
                else if (task.operation == FILESYSTEM_DELETE)
                {
                    const bool hasProgressNumbers = task.completedItems > 0 || task.completedBytes > 0 || task.totalItems > 0 || task.totalBytes > 0;
                    const bool showPreparing      = ! hasProgressNumbers;

                    if (showPreparing)
                    {
                        const ULONGLONG opStartTick = task.operationStartTick;
                        const uint64_t elapsedSec =
                            (opStartTick > 0 && nowTick >= opStartTick) ? static_cast<uint64_t>((nowTick - opStartTick) / 1000ull) : 0ull;
                        const std::wstring prepText = elapsedSec > 0
                                                          ? FormatStringResource(nullptr, IDS_FMT_FILEOPS_PREPARING_TIME, FormatDurationHms(elapsedSec))
                                                          : LoadStringResource(nullptr, IDS_FILEOPS_PREPARING);
                        const D2D1_RECT_F prepRc    = D2D1::RectF(textX, textY, textX + textMaxW, textY + lineH);
                        _target->DrawTextW(prepText.data(), static_cast<UINT32>(prepText.size()), _bodyFormat.get(), prepRc, _subTextBrush.get());
                        textY += lineH;
                    }
                    else
                    {
                        const double itemsPerSec     = history ? history->displayedItemsPerSec : 0.0;
                        const std::wstring speedText = FormatStringResource(nullptr, IDS_FMT_FILEOP_SPEED_ITEMS, itemsPerSec);
                        const D2D1_RECT_F speedRc    = D2D1::RectF(textX, textY, textX + textMaxW, textY + lineH);
                        _target->DrawTextW(speedText.data(), static_cast<UINT32>(speedText.size()), _bodyFormat.get(), speedRc, _subTextBrush.get());
                        textY += lineH;

                        const bool showSizeProgress = task.preCalcCompleted && task.preCalcTotalBytes > 0 && task.completedBytes > 0;
                        if (showSizeProgress)
                        {
                            const std::wstring sizeProgressText = FormatStringResource(
                                nullptr, IDS_FMT_FILEOPS_SIZE_PROGRESS, FormatBytesCompact(task.completedBytes), FormatBytesCompact(task.preCalcTotalBytes));
                            const D2D1_RECT_F sizeRc = D2D1::RectF(textX, textY, textX + textMaxW, textY + lineH);
                            _target->DrawTextW(
                                sizeProgressText.data(), static_cast<UINT32>(sizeProgressText.size()), _bodyFormat.get(), sizeRc, _subTextBrush.get());
                            textY += lineH;
                        }
                        else if (task.totalItems > 0)
                        {
                            const std::wstring itemsProgressText = FormatStringResource(nullptr, IDS_FMT_FILEOP_ITEMS, task.completedItems, task.totalItems);
                            const D2D1_RECT_F itemsRc            = D2D1::RectF(textX, textY, textX + textMaxW, textY + lineH);
                            _target->DrawTextW(
                                itemsProgressText.data(), static_cast<UINT32>(itemsProgressText.size()), _bodyFormat.get(), itemsRc, _subTextBrush.get());
                            textY += lineH;
                        }
                        else
                        {
                            const std::wstring itemsProgressText = FormatStringResource(nullptr, IDS_FMT_FILEOP_ITEMS_UNKNOWN_TOTAL, task.completedItems);
                            const D2D1_RECT_F itemsRc            = D2D1::RectF(textX, textY, textX + textMaxW, textY + lineH);
                            _target->DrawTextW(
                                itemsProgressText.data(), static_cast<UINT32>(itemsProgressText.size()), _bodyFormat.get(), itemsRc, _subTextBrush.get());
                            textY += lineH;
                        }
                    }
                }
                else
                {
                    if (task.preCalcSkipped && (task.operation == FILESYSTEM_COPY || task.operation == FILESYSTEM_MOVE))
                    {
                        const uint64_t completedTotal = static_cast<uint64_t>(task.completedFiles) + static_cast<uint64_t>(task.completedFolders);
                        const bool haveBreakdown      = completedTotal == static_cast<uint64_t>(task.completedItems);
                        if (haveBreakdown && task.completedItems > 0)
                        {
                            const std::wstring countsText =
                                FormatStringResource(nullptr, IDS_FMT_FILEOPS_FILES_FOLDERS, completedTotal, task.completedFiles, task.completedFolders);
                            const D2D1_RECT_F countsRc = D2D1::RectF(textX, textY, textX + textMaxW, textY + lineH);
                            _target->DrawTextW(countsText.data(), static_cast<UINT32>(countsText.size()), _bodyFormat.get(), countsRc, _subTextBrush.get());
                            textY += lineH;
                        }
                    }

                    const double bytesPerSec          = history ? history->displayedBytesPerSec : 0.0;
                    const uint64_t bytesPerSecRounded = bytesPerSec > 0.0 ? static_cast<uint64_t>(bytesPerSec + 0.5) : 0ull;
                    const std::wstring bytesText      = FormatBytesCompact(bytesPerSecRounded);
                    const std::wstring speedText      = FormatStringResource(nullptr, IDS_FMT_FILEOP_SPEED_BYTES, bytesText);
                    const D2D1_RECT_F speedRc         = D2D1::RectF(textX, textY, textX + textMaxW, textY + lineH);
                    _target->DrawTextW(speedText.data(), static_cast<UINT32>(speedText.size()), _bodyFormat.get(), speedRc, _subTextBrush.get());
                    textY += lineH;

                    // Show size progress (transferred / total) if we have data
                    if (task.totalBytes > 0)
                    {
                        const std::wstring sizeProgressText = FormatStringResource(
                            nullptr, IDS_FMT_FILEOPS_SIZE_PROGRESS, FormatBytesCompact(task.completedBytes), FormatBytesCompact(task.totalBytes));
                        const D2D1_RECT_F sizeRc = D2D1::RectF(textX, textY, textX + textMaxW, textY + lineH);
                        _target->DrawTextW(
                            sizeProgressText.data(), static_cast<UINT32>(sizeProgressText.size()), _bodyFormat.get(), sizeRc, _subTextBrush.get());
                        textY += lineH;
                    }

                    if (task.totalBytes > 0 && history && history->hasSmoothedEta && history->smoothedEtaSeconds > 0.0 &&
                        task.completedBytes <= task.totalBytes)
                    {
                        const uint64_t seconds     = static_cast<uint64_t>(std::ceil(history->smoothedEtaSeconds));
                        const std::wstring etaText = FormatStringResource(nullptr, IDS_FMT_FILEOPS_ETA, FormatDurationHms(seconds));
                        const D2D1_RECT_F etaRc    = D2D1::RectF(textX, textY, textX + textMaxW, textY + lineH);
                        _target->DrawTextW(etaText.data(), static_cast<UINT32>(etaText.size()), _bodyFormat.get(), etaRc, _subTextBrush.get());
                        textY += lineH;
                    }
                }

                if (task.autoConcurrencyUsed && task.autoTunedConcurrency > 0u)
                {
                    const std::wstring autoConcurrencyText =
                        FormatStringResource(nullptr, IDS_FMT_FILEOPS_AUTO_CONCURRENCY, task.autoTunedConcurrency, task.effectiveConcurrencyBudget);
                    const D2D1_RECT_F autoConcurrencyRc = D2D1::RectF(textX, textY, textX + textMaxW, textY + lineH);
                    _target->DrawTextW(
                        autoConcurrencyText.data(), static_cast<UINT32>(autoConcurrencyText.size()), _bodyFormat.get(), autoConcurrencyRc, _subTextBrush.get());
                    textY += lineH;
                }

                const float labelWDesired   = DipsToPixels(56.0f, _dpi);
                const float labelGapDesired = DipsToPixels(6.0f, _dpi);
                const float labelW          = std::min(labelWDesired, textMaxW);
                const float labelGap        = (labelW < textMaxW) ? std::min(labelGapDesired, textMaxW - labelW) : 0.0f;

                if (task.operation == FILESYSTEM_DELETE)
                {
                    const std::wstring label  = LoadStringResource(nullptr, IDS_FILEOPS_LABEL_DELETING);
                    const D2D1_RECT_F labelRc = D2D1::RectF(textX, textY, textX + labelW, textY + lineH);
                    _target->DrawTextW(label.data(), static_cast<UINT32>(label.size()), _smallFormat.get(), labelRc, _subTextBrush.get());

                    const float pathW        = std::max(0.0f, textMaxW - labelW - labelGap);
                    const std::wstring path  = TruncatePathMiddleToWidth(_dwriteFactory.get(), _bodyFormat.get(), task.currentSourcePath, pathW, lineH);
                    const D2D1_RECT_F pathRc = D2D1::RectF(textX + labelW + labelGap, textY, textX + labelW + labelGap + pathW, textY + lineH);
                    _target->DrawTextW(path.data(), static_cast<UINT32>(path.size()), _bodyFormat.get(), pathRc, _textBrush.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);
                    textY += lineH;
                }
                else
                {
                    const std::wstring fromLabel = LoadStringResource(nullptr, IDS_FILEOPS_LABEL_FROM);
                    const float miniBarGap       = DipsToPixels(8.0f, _dpi);
                    const float miniBarWDesired  = DipsToPixels(92.0f, _dpi);
                    const float miniBarH         = DipsToPixels(6.0f, _dpi);

                    const float pathLeft  = textX + labelW + labelGap;
                    const float rightEdge = textX + textMaxW;

                    const bool showInFlightFiles = task.operation == FILESYSTEM_COPY || task.operation == FILESYSTEM_MOVE;

                    std::array<size_t, TaskSnapshot::kMaxInFlightFiles> activeInFlightIndices{};
                    size_t activeInFlightCount = 0;
                    if (showInFlightFiles)
                    {
                        for (size_t j = 0; j < task.inFlightFileCount && activeInFlightCount < activeInFlightIndices.size(); ++j)
                        {
                            const auto& entry          = task.inFlightFiles[j];
                            const bool active          = entry.totalBytes == 0 || entry.completedBytes < entry.totalBytes;
                            const bool recentCompleted = ! active && entry.totalBytes > 0 && entry.completedBytes >= entry.totalBytes &&
                                                         entry.lastUpdateTick != 0 && nowTick >= entry.lastUpdateTick &&
                                                         (nowTick - entry.lastUpdateTick) <= kCompletedInFlightGraceMs;
                            if (active || recentCompleted)
                            {
                                activeInFlightIndices[activeInFlightCount] = j;
                                ++activeInFlightCount;
                            }
                        }
                    }

                    const size_t inFlightCount = showInFlightFiles ? std::max<size_t>(1u, activeInFlightCount) : 1u;

                    for (size_t i = 0; i < inFlightCount; ++i)
                    {
                        if (i == 0u)
                        {
                            const D2D1_RECT_F fromRc = D2D1::RectF(textX, textY, textX + labelW, textY + lineH);
                            _target->DrawTextW(fromLabel.data(), static_cast<UINT32>(fromLabel.size()), _smallFormat.get(), fromRc, _subTextBrush.get());
                        }

                        std::wstring_view sourcePathText;
                        uint64_t fileTotalBytes     = 0;
                        uint64_t fileCompletedBytes = 0;

                        const bool hasActiveInFlight = showInFlightFiles && activeInFlightCount > 0;
                        const bool useInFlightEntry  = hasActiveInFlight && i < activeInFlightCount;

                        if (useInFlightEntry)
                        {
                            const auto& entry  = task.inFlightFiles[activeInFlightIndices[i]];
                            sourcePathText     = entry.sourcePath;
                            fileTotalBytes     = entry.totalBytes;
                            fileCompletedBytes = entry.completedBytes;
                        }
                        else
                        {
                            sourcePathText     = task.currentSourcePath;
                            fileTotalBytes     = task.itemTotalBytes;
                            fileCompletedBytes = task.itemCompletedBytes;
                        }

                        const float availableW     = std::max(0.0f, rightEdge - pathLeft);
                        const float miniBarWMin    = DipsToPixels(40.0f, _dpi);
                        const float minTextW       = DipsToPixels(48.0f, _dpi);
                        float miniBarW             = std::min(miniBarWDesired, availableW);
                        const float maxBarWithText = std::max(0.0f, availableW - miniBarGap - minTextW);
                        if (maxBarWithText > 0.0f)
                        {
                            miniBarW = std::clamp(miniBarW, std::min(miniBarWMin, maxBarWithText), maxBarWithText);
                        }

                        // If nothing is actively copying (e.g., end-of-file or finalization), avoid showing a "stuck at 100%" mini bar.
                        if (! useInFlightEntry && fileTotalBytes > 0 && fileCompletedBytes >= fileTotalBytes)
                        {
                            miniBarW = 0.0f;
                        }

                        const float barRight = rightEdge;
                        const float barLeft  = barRight - miniBarW;

                        const float pathRight = (miniBarW > 0.0f) ? std::max(pathLeft, barLeft - miniBarGap) : rightEdge;
                        const float pathW     = std::max(0.0f, pathRight - pathLeft);

                        const std::wstring fromPath  = TruncatePathMiddleToWidth(_dwriteFactory.get(), _bodyFormat.get(), sourcePathText, pathW, lineH);
                        const D2D1_RECT_F fromPathRc = D2D1::RectF(pathLeft, textY, pathLeft + pathW, textY + lineH);
                        _target->DrawTextW(fromPath.data(),
                                           static_cast<UINT32>(fromPath.size()),
                                           _bodyFormat.get(),
                                           fromPathRc,
                                           _textBrush.get(),
                                           D2D1_DRAW_TEXT_OPTIONS_CLIP);

                        if (miniBarW > 0.0f && _progressBgBrush && _progressItemBrush)
                        {
                            const float barTop          = textY + (lineH - miniBarH) * 0.5f;
                            const D2D1_RECT_F miniBarRc = D2D1::RectF(barLeft, barTop, barRight, barTop + miniBarH);

                            const float radiusTrack = ClampCornerRadius(miniBarRc, DipsToPixels(2.0f, _dpi));
                            _target->FillRoundedRectangle(D2D1::RoundedRect(miniBarRc, radiusTrack, radiusTrack), _progressBgBrush.get());

                            const bool hasTotal = fileTotalBytes > 0;
                            const float frac = hasTotal && fileCompletedBytes <= fileTotalBytes
                                                   ? Clamp01(static_cast<float>(static_cast<double>(fileCompletedBytes) / static_cast<double>(fileTotalBytes)))
                                                   : 0.0f;

                            if (theme.menu.rainbowMode)
                            {
                                const D2D1::ColorF rainbow = RainbowProgressColor(theme, sourcePathText);
                                _progressItemBrush->SetColor(rainbow);
                            }
                            else
                            {
                                _progressItemBrush->SetColor(_progressItemBaseColor);
                            }

                            const D2D1_RECT_F fill =
                                hasTotal
                                    ? D2D1::RectF(miniBarRc.left, miniBarRc.top, miniBarRc.left + (miniBarRc.right - miniBarRc.left) * frac, miniBarRc.bottom)
                                    : ComputeIndeterminateBarFill(miniBarRc, nowTick);
                            const float radiusFill = ClampCornerRadius(fill, DipsToPixels(2.0f, _dpi));
                            _target->FillRoundedRectangle(D2D1::RoundedRect(fill, radiusFill, radiusFill), _progressItemBrush.get());
                        }

                        textY += lineH;
                    }

                    textY += fromToGapY;

                    const std::wstring toLabel = LoadStringResource(nullptr, IDS_FILEOPS_LABEL_TO);
                    const D2D1_RECT_F toRc     = D2D1::RectF(textX, textY, textX + labelW, textY + lineH);
                    _target->DrawTextW(toLabel.data(), static_cast<UINT32>(toLabel.size()), _smallFormat.get(), toRc, _subTextBrush.get());

                    const std::wstring destText = task.destinationFolder.wstring();

                    const float toPathLeft = textX + labelW + labelGap;
                    const float toRight    = textX + textMaxW;

                    const bool canSelectDestination =
                        (task.operation == FILESYSTEM_COPY || task.operation == FILESYSTEM_MOVE) && ! task.started && task.destinationPane.has_value();

                    float destMenuW         = canSelectDestination ? DipsToPixels(28.0f, _dpi) : 0.0f;
                    const float destMenuGap = (destMenuW > 0.0f) ? DipsToPixels(6.0f, _dpi) : 0.0f;

                    const float minPathW = DipsToPixels(80.0f, _dpi);
                    if (destMenuW > 0.0f && (toRight - toPathLeft) < (minPathW + destMenuGap + destMenuW))
                    {
                        destMenuW = 0.0f;
                    }

                    const float toPathRight    = (destMenuW > 0.0f) ? std::max(toPathLeft, toRight - destMenuW - destMenuGap) : toRight;
                    const float toPathW        = std::max(0.0f, toPathRight - toPathLeft);
                    const std::wstring toPath  = TruncatePathMiddleToWidth(_dwriteFactory.get(), _bodyFormat.get(), destText, toPathW, lineH);
                    const D2D1_RECT_F toPathRc = D2D1::RectF(toPathLeft, textY, toPathLeft + toPathW, textY + lineH);
                    _target->DrawTextW(
                        toPath.data(), static_cast<UINT32>(toPath.size()), _bodyFormat.get(), toPathRc, _textBrush.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);

                    if (destMenuW > 0.0f)
                    {
                        PopupButton destBtn{};
                        destBtn.bounds     = D2D1::RectF(toRight - destMenuW, textY, toRight, textY + lineH);
                        destBtn.hit.kind   = PopupHitTest::Kind::TaskDestination;
                        destBtn.hit.taskId = task.taskId;
                        _buttons.push_back(destBtn);
                        DrawMenuButton(destBtn, nullptr, {});
                    }
                    textY += lineH;
                }

                const float barInsetX = DipsToPixels(10.0f, _dpi);
                const float barW      = std::max(0.0f, cardRect.right - cardRect.left - barInsetX * 2.0f);
                const float barX      = cardRect.left + barInsetX;

                const float barHItem  = DipsToPixels(10.0f, _dpi);
                const float barHTotal = DipsToPixels(6.0f, _dpi);
                const float barGapY   = DipsToPixels(4.0f, _dpi);

                const bool hasConflictPrompt = task.conflict.active;

                const float barsHeight    = task.operation == FILESYSTEM_DELETE ? barHItem : (barHItem + barGapY + barHTotal);
                const float bottomPadding = DipsToPixels(10.0f, _dpi);
                const float buttonGapY    = DipsToPixels(8.0f, _dpi);
                const float buttonH       = DipsToPixels(24.0f, _dpi);

                const float conflictRowGapY = DipsToPixels(6.0f, _dpi);
                const int conflictRows      = hasConflictPrompt ? ((task.conflict.actionCount > 3u) ? 2 : 1) : 1;
                const float conflictButtonsHeight =
                    buttonH * static_cast<float>(conflictRows) + conflictRowGapY * static_cast<float>(std::max(0, conflictRows - 1));
                const float conflictApplyLineHeight = hasConflictPrompt ? (lineH + conflictRowGapY) : 0.0f;
                const float buttonsHeight           = conflictButtonsHeight + conflictApplyLineHeight;

                const float buttonRowBottom = cardRect.bottom - bottomPadding;
                const float buttonRowTop    = buttonRowBottom - buttonsHeight;

                const float barsBottom = buttonRowTop - buttonGapY;
                const float barsTop    = barsBottom - barsHeight;

                const auto conflictBucketToMessageId = [&](uint8_t bucket) noexcept -> UINT
                {
                    using Bucket = FolderWindow::FileOperationState::Task::ConflictBucket;
                    switch (static_cast<Bucket>(bucket))
                    {
                        case Bucket::Exists: return static_cast<UINT>(IDS_FILEOPS_CONFLICT_EXISTS);
                        case Bucket::ReadOnly: return static_cast<UINT>(IDS_FILEOPS_CONFLICT_READONLY);
                        case Bucket::AccessDenied: return static_cast<UINT>(IDS_FILEOPS_CONFLICT_ACCESS_DENIED);
                        case Bucket::SharingViolation: return static_cast<UINT>(IDS_FILEOPS_CONFLICT_SHARING);
                        case Bucket::DiskFull: return static_cast<UINT>(IDS_FILEOPS_CONFLICT_DISK_FULL);
                        case Bucket::PathTooLong: return static_cast<UINT>(IDS_FILEOPS_CONFLICT_PATH_TOO_LONG);
                        case Bucket::RecycleBinFailed: return static_cast<UINT>(IDS_FILEOPS_CONFLICT_RECYCLE_BIN);
                        case Bucket::NetworkOffline: return static_cast<UINT>(IDS_FILEOPS_CONFLICT_NETWORK);
                        case Bucket::UnsupportedReparse: return static_cast<UINT>(IDS_FILEOPS_CONFLICT_UNSUPPORTED_REPARSE);
                        case Bucket::Unknown: return static_cast<UINT>(IDS_FILEOPS_CONFLICT_UNKNOWN);
                        case Bucket::Count: return static_cast<UINT>(IDS_FILEOPS_CONFLICT_UNKNOWN);
                        default: return static_cast<UINT>(IDS_FILEOPS_CONFLICT_UNKNOWN);
                    }
                };

                const auto drawConflictPromptInfo = [&](const D2D1_RECT_F& rc) noexcept
                {
                    if (! _bodyFormat || ! _smallFormat)
                    {
                        return;
                    }

                    float yPrompt           = rc.top;
                    const float maxW        = std::max(0.0f, rc.right - rc.left);
                    const float maxDetailsY = rc.bottom;

                    std::wstring message = LoadStringResource(nullptr, conflictBucketToMessageId(task.conflict.bucket));
                    if (task.conflict.retryFailed)
                    {
                        const std::wstring retryFailed = LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_RETRY_FAILED);
                        message                        = std::format(L"{} {}", retryFailed, message);
                    }

                    if (task.conflict.bucket == static_cast<uint8_t>(FolderWindow::FileOperationState::Task::ConflictBucket::Unknown))
                    {
                        message = std::format(L"{} (0x{:08X})", message, static_cast<unsigned long>(task.conflict.status));
                    }

                    const D2D1_RECT_F msgRc = D2D1::RectF(rc.left, yPrompt, rc.left + maxW, yPrompt + lineH);
                    _target->DrawTextW(
                        message.data(), static_cast<UINT32>(message.size()), _bodyFormat.get(), msgRc, _textBrush.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);
                    yPrompt += lineH;

                    const auto drawConflictPathLine = [&](std::wstring_view label, std::wstring_view path) noexcept
                    {
                        if (label.empty() || path.empty() || ! _dwriteFactory)
                        {
                            return;
                        }

                        if (yPrompt + lineH > maxDetailsY)
                        {
                            return;
                        }

                        const float labelW   = MeasureTextWidth(_dwriteFactory.get(), _smallFormat.get(), label, maxW, lineH);
                        const float labelGap = DipsToPixels(6.0f, _dpi);
                        const float pathLeft = rc.left + labelW + labelGap;
                        const float pathW    = std::max(0.0f, rc.right - pathLeft);

                        const D2D1_RECT_F labelRc = D2D1::RectF(rc.left, yPrompt, rc.left + labelW, yPrompt + lineH);
                        _target->DrawTextW(
                            label.data(), static_cast<UINT32>(label.size()), _smallFormat.get(), labelRc, _subTextBrush.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);

                        const std::wstring truncated = TruncatePathMiddleToWidth(_dwriteFactory.get(), _smallFormat.get(), path, pathW, lineH);
                        const D2D1_RECT_F pathRc     = D2D1::RectF(pathLeft, yPrompt, rc.right, yPrompt + lineH);
                        _target->DrawTextW(
                            truncated.data(), static_cast<UINT32>(truncated.size()), _smallFormat.get(), pathRc, _textBrush.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);

                        yPrompt += lineH;
                    };

                    if (task.operation == FILESYSTEM_DELETE)
                    {
                        drawConflictPathLine(LoadStringResource(nullptr, IDS_FILEOPS_LABEL_DELETING), task.conflict.sourcePath);
                    }
                    else
                    {
                        drawConflictPathLine(LoadStringResource(nullptr, IDS_FILEOPS_LABEL_FROM), task.conflict.sourcePath);
                        drawConflictPathLine(LoadStringResource(nullptr, IDS_FILEOPS_LABEL_TO), task.conflict.destinationPath);
                    }
                };

                if (task.operation != FILESYSTEM_DELETE)
                {
                    const float graphTop    = textY + DipsToPixels(4.0f, _dpi);
                    const float graphBottom = hasConflictPrompt ? barsBottom : (barsTop - DipsToPixels(6.0f, _dpi));
                    const float graphMinH   = DipsToPixels(32.0f, _dpi);

                    if ((graphBottom - graphTop) >= graphMinH)
                    {
                        const D2D1_RECT_F graphRc = D2D1::RectF(barX, graphTop, barX + barW, graphBottom);
                        if (hasConflictPrompt)
                        {
                            drawConflictPromptInfo(graphRc);
                        }
                        else
                        {
                            uint64_t limit = 0;
                            if (task.operation != FILESYSTEM_DELETE)
                            {
                                limit =
                                    task.effectiveSpeedLimitBytesPerSecond != 0 ? task.effectiveSpeedLimitBytesPerSecond : task.desiredSpeedLimitBytesPerSecond;
                            }
                            const RateHistory empty{};
                            const RateHistory& graphHistory = history ? *history : empty;
                            std::wstring overlayText;
                            bool showAnimation = false;
                            if (task.paused)
                            {
                                overlayText = LoadStringResource(nullptr, IDS_FILEOPS_GRAPH_PAUSED);
                            }
                            else if (task.queuePaused || task.waitingInQueue)
                            {
                                overlayText = LoadStringResource(nullptr, IDS_FILEOPS_GRAPH_WAITING);
                            }
                            else if (task.preCalcInProgress)
                            {
                                overlayText   = LoadStringResource(nullptr, IDS_FILEOPS_GRAPH_CALCULATING);
                                showAnimation = true;
                            }
                            else if (! task.started || ! task.hasProgressCallbacks)
                            {
                                overlayText   = LoadStringResource(nullptr, IDS_FILEOPS_PREPARING);
                                showAnimation = true;
                            }
                            const bool rainbowMode = folderWindow && folderWindow->GetTheme().menu.rainbowMode;
                            DrawBandwidthGraph(graphRc, graphHistory, limit, overlayText, showAnimation, rainbowMode, nowTick);
                        }
                    }
                }
                else
                {
                    const float graphTop    = textY + DipsToPixels(4.0f, _dpi);
                    const float graphBottom = hasConflictPrompt ? barsBottom : (barsTop - DipsToPixels(6.0f, _dpi));
                    const float graphMinH   = DipsToPixels(32.0f, _dpi);

                    if ((graphBottom - graphTop) >= graphMinH)
                    {
                        const D2D1_RECT_F graphRc = D2D1::RectF(barX, graphTop, barX + barW, graphBottom);
                        if (hasConflictPrompt)
                        {
                            drawConflictPromptInfo(graphRc);
                        }
                        else
                        {
                            const RateHistory empty{};
                            const RateHistory& graphHistory = history ? *history : empty;
                            std::wstring overlayText;
                            bool showAnimation = false;
                            if (task.paused)
                            {
                                overlayText = LoadStringResource(nullptr, IDS_FILEOPS_GRAPH_PAUSED);
                            }
                            else if (task.queuePaused || task.waitingInQueue)
                            {
                                overlayText = LoadStringResource(nullptr, IDS_FILEOPS_GRAPH_WAITING);
                            }
                            else if (task.preCalcInProgress)
                            {
                                overlayText   = LoadStringResource(nullptr, IDS_FILEOPS_GRAPH_CALCULATING);
                                showAnimation = true;
                            }
                            else
                            {
                                const bool hasProgressNumbers =
                                    task.completedItems > 0 || task.completedBytes > 0 || task.totalItems > 0 || task.totalBytes > 0;
                                const bool showPreparing =
                                    ! task.started || ! task.hasProgressCallbacks || (task.operation == FILESYSTEM_DELETE && ! hasProgressNumbers);
                                if (showPreparing)
                                {
                                    overlayText   = LoadStringResource(nullptr, IDS_FILEOPS_PREPARING);
                                    showAnimation = true;
                                }
                            }
                            const bool rainbowMode = folderWindow && folderWindow->GetTheme().menu.rainbowMode;
                            DrawBandwidthGraph(graphRc, graphHistory, 0, overlayText, showAnimation, rainbowMode, nowTick);
                        }
                    }
                }

                // During pre-calculation, show marquee progress bar
                if (task.preCalcInProgress)
                {
                    const D2D1_RECT_F barRc = D2D1::RectF(barX, barsTop, barX + barW, barsTop + barHItem);

                    if (_progressBgBrush)
                    {
                        const float radius = ClampCornerRadius(barRc, DipsToPixels(2.0f, _dpi));
                        _target->FillRoundedRectangle(D2D1::RoundedRect(barRc, radius, radius), _progressBgBrush.get());
                    }

                    if (_progressItemBrush)
                    {
                        const D2D1_RECT_F fill = ComputeIndeterminateBarFill(barRc, nowTick);
                        const float radius     = ClampCornerRadius(fill, DipsToPixels(2.0f, _dpi));
                        _target->FillRoundedRectangle(D2D1::RoundedRect(fill, radius, radius), _progressItemBrush.get());
                    }
                }
                else if (hasConflictPrompt)
                {
                    // Conflict prompt uses the progress bar area so actions and the apply-to-all toggle sit close together.
                }
                else if (task.operation == FILESYSTEM_DELETE)
                {
                    const D2D1_RECT_F totalBarRc = D2D1::RectF(barX, barsTop, barX + barW, barsTop + barHItem);

                    if (_progressBgBrush)
                    {
                        const float radius = ClampCornerRadius(totalBarRc, DipsToPixels(2.0f, _dpi));
                        _target->FillRoundedRectangle(D2D1::RoundedRect(totalBarRc, radius, radius), _progressBgBrush.get());
                    }

                    if (_progressGlobalBrush)
                    {
                        const bool hasTotalBytes  = task.totalBytes > 0 && task.completedBytes <= task.totalBytes;
                        const bool hasUsefulItems = task.totalItems > 1;

                        const bool useBytes = hasTotalBytes && task.completedBytes > 0;
                        const bool useItems = ! useBytes && hasUsefulItems && task.completedItems > 0;

                        float totalFrac = 0.0f;
                        if (useBytes)
                        {
                            totalFrac = Clamp01(static_cast<float>(static_cast<double>(task.completedBytes) / static_cast<double>(task.totalBytes)));
                        }
                        else if (useItems)
                        {
                            const double denom = static_cast<double>(task.totalItems);
                            const double numer = static_cast<double>(std::min(task.completedItems, task.totalItems));
                            totalFrac          = Clamp01(static_cast<float>(numer / denom));
                        }

                        const D2D1_RECT_F fill =
                            (useBytes || useItems)
                                ? D2D1::RectF(
                                      totalBarRc.left, totalBarRc.top, totalBarRc.left + (totalBarRc.right - totalBarRc.left) * totalFrac, totalBarRc.bottom)
                                : ComputeIndeterminateBarFill(totalBarRc, nowTick);
                        const float radius = ClampCornerRadius(fill, DipsToPixels(2.0f, _dpi));
                        _target->FillRoundedRectangle(D2D1::RoundedRect(fill, radius, radius), _progressGlobalBrush.get());
                    }
                }
                else
                {
                    const D2D1_RECT_F itemBarRc  = D2D1::RectF(barX, barsTop, barX + barW, barsTop + barHItem);
                    const D2D1_RECT_F totalBarRc = D2D1::RectF(barX, itemBarRc.bottom + barGapY, barX + barW, itemBarRc.bottom + barGapY + barHTotal);

                    if (_progressBgBrush)
                    {
                        const float radiusItem  = ClampCornerRadius(itemBarRc, DipsToPixels(2.0f, _dpi));
                        const float radiusTotal = ClampCornerRadius(totalBarRc, DipsToPixels(2.0f, _dpi));
                        _target->FillRoundedRectangle(D2D1::RoundedRect(itemBarRc, radiusItem, radiusItem), _progressBgBrush.get());
                        _target->FillRoundedRectangle(D2D1::RoundedRect(totalBarRc, radiusTotal, radiusTotal), _progressBgBrush.get());
                    }

                    const bool hasItemBytes = task.itemTotalBytes > 0;
                    const float itemFrac =
                        hasItemBytes ? Clamp01(static_cast<float>(static_cast<double>(task.itemCompletedBytes) / static_cast<double>(task.itemTotalBytes)))
                                     : 0.0f;

                    if (_progressItemBrush)
                    {
                        if (theme.menu.rainbowMode)
                        {
                            const D2D1::ColorF rainbow = RainbowProgressColor(theme, task.currentSourcePath);
                            _progressItemBrush->SetColor(rainbow);
                        }
                        else
                        {
                            _progressItemBrush->SetColor(_progressItemBaseColor);
                        }

                        const D2D1_RECT_F fill =
                            hasItemBytes
                                ? D2D1::RectF(itemBarRc.left, itemBarRc.top, itemBarRc.left + (itemBarRc.right - itemBarRc.left) * itemFrac, itemBarRc.bottom)
                                : ComputeIndeterminateBarFill(itemBarRc, nowTick);
                        const float radius = ClampCornerRadius(fill, DipsToPixels(2.0f, _dpi));
                        _target->FillRoundedRectangle(D2D1::RoundedRect(fill, radius, radius), _progressItemBrush.get());
                    }

                    float totalFrac = 0.0f;
                    if (task.totalBytes > 0 && task.completedBytes <= task.totalBytes)
                    {
                        totalFrac = Clamp01(static_cast<float>(static_cast<double>(task.completedBytes) / static_cast<double>(task.totalBytes)));
                    }
                    else if (task.totalItems > 0)
                    {
                        const double denom = static_cast<double>(task.totalItems);
                        const double numer = static_cast<double>(std::min(task.completedItems, task.totalItems)) + static_cast<double>(itemFrac);
                        totalFrac          = Clamp01(static_cast<float>(numer / denom));
                    }

                    if (_progressGlobalBrush)
                    {
                        const D2D1_RECT_F fill =
                            D2D1::RectF(totalBarRc.left, totalBarRc.top, totalBarRc.left + (totalBarRc.right - totalBarRc.left) * totalFrac, totalBarRc.bottom);
                        const float radius = ClampCornerRadius(fill, DipsToPixels(2.0f, _dpi));
                        _target->FillRoundedRectangle(D2D1::RoundedRect(fill, radius, radius), _progressGlobalBrush.get());
                    }
                }

                {
                    const float btnGapX = DipsToPixels(8.0f, _dpi);
                    const float rowW    = std::max(0.0f, contentRight - textX);
                    if (rowW > 1.0f)
                    {
                        const float rowTop    = buttonRowTop;
                        const float rowBottom = buttonRowBottom;

                        if (hasConflictPrompt)
                        {
                            // "Apply to all" is placed directly above the conflict action buttons so it's easy to notice and use.
                            const float applyTop    = rowTop;
                            const float applyBottom = applyTop + lineH;
                            const float buttonsTop  = applyBottom + conflictRowGapY;

                            const float checkSize     = DipsToPixels(16.0f, _dpi);
                            const float checkTop      = applyTop + (lineH - checkSize) * 0.5f;
                            const D2D1_RECT_F checkRc = D2D1::RectF(textX, checkTop, textX + checkSize, checkTop + checkSize);
                            DrawCheckboxBox(checkRc, task.conflict.applyToAllChecked);

                            const std::wstring applyText = LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_APPLY_TO_ALL);
                            const float labelLeft        = textX + checkSize + DipsToPixels(8.0f, _dpi);
                            const D2D1_RECT_F labelRc    = D2D1::RectF(labelLeft, applyTop, contentRight, applyBottom);

                            IDWriteTextFormat* applyFormat = _bodyFormat.get();
                            ID2D1Brush* applyBrush         = _textBrush ? _textBrush.get() : (_subTextBrush ? _subTextBrush.get() : nullptr);
                            if (applyFormat && applyBrush && ! applyText.empty())
                            {
                                _target->DrawTextW(
                                    applyText.data(), static_cast<UINT32>(applyText.size()), applyFormat, labelRc, applyBrush, D2D1_DRAW_TEXT_OPTIONS_CLIP);
                            }

                            PopupButton applyBtn{};
                            applyBtn.bounds     = D2D1::RectF(textX, applyTop, contentRight, applyBottom);
                            applyBtn.hit.kind   = PopupHitTest::Kind::TaskConflictToggleApplyToAll;
                            applyBtn.hit.taskId = task.taskId;
                            _buttons.push_back(applyBtn);

                            const auto conflictActionText = [&](FolderWindow::FileOperationState::Task::ConflictAction action) noexcept -> std::wstring
                            {
                                switch (action)
                                {
                                    case FolderWindow::FileOperationState::Task::ConflictAction::Overwrite:
                                        return LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_BTN_OVERWRITE);
                                    case FolderWindow::FileOperationState::Task::ConflictAction::ReplaceReadOnly:
                                        return LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_BTN_REPLACE_READONLY);
                                    case FolderWindow::FileOperationState::Task::ConflictAction::PermanentDelete:
                                        return LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_BTN_PERMANENT_DELETE);
                                    case FolderWindow::FileOperationState::Task::ConflictAction::Retry:
                                        return LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_BTN_RETRY);
                                    case FolderWindow::FileOperationState::Task::ConflictAction::Skip:
                                        return LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_BTN_SKIP);
                                    case FolderWindow::FileOperationState::Task::ConflictAction::SkipAll:
                                        return LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_BTN_SKIP_ALL);
                                    case FolderWindow::FileOperationState::Task::ConflictAction::Cancel:
                                        return LoadStringResource(nullptr, IDS_FILEOP_BTN_CANCEL);
                                    case FolderWindow::FileOperationState::Task::ConflictAction::None:
                                    default: break;
                                }
                                return LoadStringResource(nullptr, IDS_FILEOP_BTN_CANCEL);
                            };

                            constexpr size_t kMaxPerRow = 3u;
                            size_t actionIndex          = 0;
                            const size_t totalActions   = task.conflict.actionCount;

                            for (int row = 0; row < conflictRows; ++row)
                            {
                                if (actionIndex >= totalActions)
                                {
                                    break;
                                }

                                const float rowY       = buttonsTop + static_cast<float>(row) * (buttonH + conflictRowGapY);
                                const float rowYBottom = rowY + buttonH;

                                const size_t remaining   = totalActions - actionIndex;
                                const size_t buttonCount = std::min(kMaxPerRow, remaining);
                                if (buttonCount == 0)
                                {
                                    break;
                                }

                                const float totalGapX = btnGapX * static_cast<float>(buttonCount - 1u);
                                const float btnW      = std::max(0.0f, (rowW - totalGapX) / static_cast<float>(buttonCount));

                                float xBtn = textX;
                                for (size_t i = 0; i < buttonCount; ++i)
                                {
                                    const uint8_t rawAction  = task.conflict.actions[actionIndex];
                                    const auto action        = static_cast<FolderWindow::FileOperationState::Task::ConflictAction>(rawAction);
                                    const std::wstring label = conflictActionText(action);

                                    PopupButton btn{};
                                    btn.bounds     = D2D1::RectF(xBtn, rowY, xBtn + btnW, rowYBottom);
                                    btn.hit.kind   = PopupHitTest::Kind::TaskConflictAction;
                                    btn.hit.taskId = task.taskId;
                                    btn.hit.data   = static_cast<uint32_t>(rawAction);
                                    _buttons.push_back(btn);
                                    DrawButton(btn, _buttonSmallFormat.get(), label);

                                    xBtn += btnW + btnGapX;
                                    ++actionIndex;
                                }
                            }
                        }
                        // During pre-calculation, show Skip and Cancel buttons
                        else if (task.preCalcInProgress)
                        {
                            const std::wstring skipText = LoadStringResource(nullptr, IDS_FILEOPS_BTN_SKIP);
                            const float skipW           = std::max(0.0f, (rowW - btnGapX) * 0.5f);
                            const float calcCancelW     = std::max(0.0f, rowW - btnGapX - skipW);

                            PopupButton skipBtn{};
                            skipBtn.bounds     = D2D1::RectF(textX, rowTop, textX + skipW, rowBottom);
                            skipBtn.hit.kind   = PopupHitTest::Kind::TaskSkip;
                            skipBtn.hit.taskId = task.taskId;
                            _buttons.push_back(skipBtn);
                            DrawButton(skipBtn, _buttonSmallFormat.get(), skipText);

                            PopupButton calcCancelBtn{};
                            calcCancelBtn.bounds     = D2D1::RectF(textX + skipW + btnGapX, rowTop, textX + skipW + btnGapX + calcCancelW, rowBottom);
                            calcCancelBtn.hit.kind   = PopupHitTest::Kind::TaskCancel;
                            calcCancelBtn.hit.taskId = task.taskId;
                            _buttons.push_back(calcCancelBtn);
                            DrawButton(calcCancelBtn, _buttonSmallFormat.get(), cancelText);
                        }
                        else if (showCopyMoveControls && ! speedLimitText.empty())
                        {
                            const float available = std::max(0.0f, rowW - btnGapX * 2.0f);
                            const float minEach   = DipsToPixels(68.0f, _dpi);

                            float pauseW  = DipsToPixels(84.0f, _dpi);
                            float cancelW = DipsToPixels(84.0f, _dpi);
                            float limitW  = std::max(0.0f, available - pauseW - cancelW);

                            if (available < minEach * 3.0f)
                            {
                                const float eachW = available / 3.0f;
                                pauseW            = eachW;
                                cancelW           = eachW;
                                limitW            = eachW;
                            }
                            else
                            {
                                const float minLimitW = DipsToPixels(140.0f, _dpi);
                                if (limitW < minLimitW)
                                {
                                    const float minSideW          = DipsToPixels(72.0f, _dpi);
                                    const float remainingForSides = std::max(0.0f, available - minLimitW);
                                    const float sideW             = std::max(minSideW, remainingForSides / 2.0f);
                                    pauseW                        = std::min(pauseW, sideW);
                                    cancelW                       = std::min(cancelW, sideW);
                                    limitW                        = std::max(0.0f, available - pauseW - cancelW);
                                }
                            }

                            float xBtn = textX;

                            PopupButton pauseBtn{};
                            pauseBtn.bounds     = D2D1::RectF(xBtn, rowTop, xBtn + pauseW, rowBottom);
                            pauseBtn.hit.kind   = PopupHitTest::Kind::TaskPause;
                            pauseBtn.hit.taskId = task.taskId;
                            _buttons.push_back(pauseBtn);
                            DrawButton(pauseBtn, _buttonSmallFormat.get(), pauseText);
                            xBtn += pauseW + btnGapX;

                            PopupButton limitBtn{};
                            limitBtn.bounds     = D2D1::RectF(xBtn, rowTop, xBtn + limitW, rowBottom);
                            limitBtn.hit.kind   = PopupHitTest::Kind::TaskSpeedLimit;
                            limitBtn.hit.taskId = task.taskId;
                            _buttons.push_back(limitBtn);
                            DrawMenuButton(limitBtn, _buttonSmallFormat.get(), speedLimitText);
                            xBtn += limitW + btnGapX;

                            PopupButton cancelBtn{};
                            cancelBtn.bounds     = D2D1::RectF(xBtn, rowTop, xBtn + cancelW, rowBottom);
                            cancelBtn.hit.kind   = PopupHitTest::Kind::TaskCancel;
                            cancelBtn.hit.taskId = task.taskId;
                            _buttons.push_back(cancelBtn);
                            DrawButton(cancelBtn, _buttonSmallFormat.get(), cancelText);
                        }
                        else
                        {
                            const float pauseW  = std::max(0.0f, (rowW - btnGapX) * 0.5f);
                            const float cancelW = std::max(0.0f, rowW - btnGapX - pauseW);

                            PopupButton pauseBtn{};
                            pauseBtn.bounds     = D2D1::RectF(textX, rowTop, textX + pauseW, rowBottom);
                            pauseBtn.hit.kind   = PopupHitTest::Kind::TaskPause;
                            pauseBtn.hit.taskId = task.taskId;
                            _buttons.push_back(pauseBtn);
                            DrawButton(pauseBtn, _buttonSmallFormat.get(), pauseText);

                            PopupButton cancelBtn{};
                            cancelBtn.bounds     = D2D1::RectF(textX + pauseW + btnGapX, rowTop, textX + pauseW + btnGapX + cancelW, rowBottom);
                            cancelBtn.hit.kind   = PopupHitTest::Kind::TaskCancel;
                            cancelBtn.hit.taskId = task.taskId;
                            _buttons.push_back(cancelBtn);
                            DrawButton(cancelBtn, _buttonSmallFormat.get(), cancelText);
                        }
                    }
                }
            }

            const float gapAfter = (taskIndex + 1u < taskCount) ? cardGap : 0.0f;
            y += taskCardH + gapAfter;
        }

        _target->PopAxisAlignedClip();
    }
    const uint64_t drawUs = capturePerf ? PerfElapsedUs(drawStartedUs) : 0u;

    if (capturePerf)
    {
        uint64_t informationalTaskCount = 0u;
        for (const TaskSnapshot& task : snapshot)
        {
            if (task.kind == TaskSnapshot::Kind::Informational)
            {
                ++informationalTaskCount;
            }
        }

        Debug::Perf::Emit(L"FileOps.Popup.Render.BuildSnapshotUs", L"", snapshotUs, informationalTaskCount, static_cast<uint64_t>(taskCount), S_OK);
        Debug::Perf::Emit(L"FileOps.Popup.Render.CardLayoutUs", L"", cardLayoutUs, informationalTaskCount, static_cast<uint64_t>(taskCount), S_OK);
        Debug::Perf::Emit(L"FileOps.Popup.Render.AutoResizeUs", L"", autoResizeUs, informationalTaskCount, static_cast<uint64_t>(taskCount), S_OK);
        Debug::Perf::Emit(L"FileOps.Popup.Render.ScrollLayoutUs", L"", scrollLayoutUs, informationalTaskCount, static_cast<uint64_t>(taskCount), S_OK);
        Debug::Perf::Emit(L"FileOps.Popup.Render.DrawUs", L"", drawUs, informationalTaskCount, static_cast<uint64_t>(taskCount), hrEndDraw);
        Debug::Perf::Emit(
            L"FileOps.Popup.Render.TotalUs", L"", PerfElapsedUs(renderStartedUs), informationalTaskCount, static_cast<uint64_t>(taskCount), hrEndDraw);
    }

    if (hrEndDraw == D2DERR_RECREATE_TARGET)
    {
        DiscardDeviceResources();
    }
}

void FileOperationsPopupInternal::FileOperationsPopupState::UpdateLastPopupRect(HWND hwnd) noexcept
{
    if (! hwnd || ! fileOps)
    {
        return;
    }

    if (! hostLifetime.lock())
    {
        return;
    }

    if (! IsWindowVisible(hwnd) || IsIconic(hwnd))
    {
        return;
    }

    RECT rc{};
    if (! GetWindowRect(hwnd, &rc))
    {
        return;
    }

    fileOps->UpdateLastPopupRect(rc);
    fileOps->SavePopupPlacement(hwnd);
}

void FileOperationsPopupInternal::FileOperationsPopupState::UpdateCaptionStatus(HWND hwnd, const std::vector<TaskSnapshot>& snapshot) noexcept
{
    const HRESULT partialHr   = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
    const HRESULT cancelledHr = HRESULT_FROM_WIN32(ERROR_CANCELLED);

    CaptionStatus computed = snapshot.empty() ? CaptionStatus::None : CaptionStatus::Ok;

    bool sawWarning = false;
    for (const TaskSnapshot& task : snapshot)
    {
        if (task.errorCount > 0)
        {
            computed = CaptionStatus::Error;
            break;
        }

        if (task.finished && FAILED(task.resultHr) && task.resultHr != partialHr && task.resultHr != cancelledHr && task.resultHr != E_ABORT)
        {
            computed = CaptionStatus::Error;
            break;
        }

        if (task.warningCount > 0 || (task.finished && task.resultHr == partialHr))
        {
            sawWarning = true;
        }
    }

    if (computed != CaptionStatus::Error && sawWarning)
    {
        computed = CaptionStatus::Warning;
    }

    if (_captionStatus == computed)
    {
        return;
    }

    _captionStatus = computed;

    if (hwnd)
    {
        RedrawWindow(hwnd, nullptr, nullptr, RDW_FRAME | RDW_NOERASE | RDW_NOCHILDREN);
    }
}

void FileOperationsPopupInternal::FileOperationsPopupState::PaintCaptionStatusGlyph(HWND hwnd) noexcept
{
    if (! hwnd)
    {
        return;
    }

    if (! hostLifetime.lock())
    {
        return;
    }

    if (! folderWindow)
    {
        return;
    }

    if (_captionStatus == CaptionStatus::None)
    {
        return;
    }

    const AppTheme& theme = folderWindow->GetTheme();
    if (theme.highContrast)
    {
        return;
    }

    RECT windowScreen{};
    if (! GetWindowRect(hwnd, &windowScreen))
    {
        return;
    }

    RECT client{};
    if (! GetClientRect(hwnd, &client))
    {
        return;
    }

    POINT clientTopLeftScreen{0, 0};
    if (! ClientToScreen(hwnd, &clientTopLeftScreen))
    {
        return;
    }

    const int windowW         = std::max(0, static_cast<int>(windowScreen.right - windowScreen.left));
    const int clientW         = std::max(0, static_cast<int>(client.right - client.left));
    const int nonClientTopH   = std::max(0, static_cast<int>(clientTopLeftScreen.y - windowScreen.top));
    const int nonClientRightW = std::max(0, static_cast<int>(windowScreen.right - (clientTopLeftScreen.x + clientW)));

    if (windowW <= 0 || nonClientTopH <= 0)
    {
        return;
    }

    const UINT dpi    = GetDpiForWindow(hwnd);
    const DWORD style = static_cast<DWORD>(GetWindowLongPtrW(hwnd, GWL_STYLE));
    const bool hasSys = (style & WS_SYSMENU) != 0;
    const bool hasMin = (style & WS_MINIMIZEBOX) != 0;
    const bool hasMax = (style & WS_MAXIMIZEBOX) != 0;
    const int buttonW = GetSystemMetricsForDpi(SM_CXSIZE, dpi);

    int buttonCount = 0;
    if (hasSys)
    {
        buttonCount += 1; // Close
    }
    if (hasMax)
    {
        buttonCount += 1;
    }
    if (hasMin)
    {
        buttonCount += 1;
    }

    if (buttonCount <= 0 || buttonW <= 0)
    {
        return;
    }

    const int iconSize = DipsToPixels(20, dpi);
    const int gap      = DipsToPixels(8, dpi);

    const int buttonsLeft = windowW - nonClientRightW - buttonW * buttonCount;
    const int iconRight   = buttonsLeft - gap;
    const int iconLeft    = iconRight - iconSize;
    const int iconTop     = std::max(0, (nonClientTopH - iconSize) / 2);

    if (iconRight <= iconLeft || iconTop + iconSize <= 0)
    {
        return;
    }

    RECT iconRc{iconLeft, iconTop, iconRight, iconTop + iconSize};
    RECT targetRc{0, 0, windowW, nonClientTopH};

    wchar_t fluentGlyph = 0;
    wchar_t fallback    = 0;
    D2D1::ColorF color  = ColorFromCOLORREF(theme.menu.text);

    switch (_captionStatus)
    {
        case CaptionStatus::Ok:
            fluentGlyph = FluentIcons::kCheckMark;
            fallback    = FluentIcons::kFallbackCheckMark;
            color       = theme.accent;
            break;
        case CaptionStatus::Warning:
            fluentGlyph = FluentIcons::kWarning;
            fallback    = FluentIcons::kFallbackWarning;
            color       = theme.folderView.warningText;
            break;
        case CaptionStatus::Error:
            fluentGlyph = FluentIcons::kError;
            fallback    = FluentIcons::kFallbackError;
            color       = theme.folderView.errorText;
            break;
        case CaptionStatus::None:
        default: return;
    }

    if (! EnsureCaptionGlyphTarget(dpi))
    {
        return;
    }

    const bool useFluentGlyph = _captionGlyphFormat != nullptr && DirectWriteFormatHasGlyph(_dwriteFactory.get(), _captionGlyphFormat.get(), fluentGlyph);
    IDWriteTextFormat* format = useFluentGlyph ? _captionGlyphFormat.get() : _captionGlyphFallbackFormat.get();
    const wchar_t glyph       = useFluentGlyph ? fluentGlyph : fallback;
    if (! format || glyph == 0)
    {
        return;
    }

    wil::unique_hdc_window hdc(GetWindowDC(hwnd));
    if (! hdc)
    {
        return;
    }

    const HRESULT bindHr = _captionGlyphTarget->BindDC(hdc.get(), &targetRc);
    if (FAILED(bindHr))
    {
        _captionGlyphBrush.reset();
        _captionGlyphTarget.reset();
        return;
    }

    _captionGlyphBrush->SetColor(color);
    const D2D1_RECT_F glyphRc = RectPixelsToDips(iconRc, dpi);
    const wchar_t glyphText[2]{glyph, 0};

    _captionGlyphTarget->BeginDraw();
    _captionGlyphTarget->DrawText(glyphText, 1u, format, glyphRc, _captionGlyphBrush.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP, DWRITE_MEASURING_MODE_NATURAL);
    const HRESULT endHr = _captionGlyphTarget->EndDraw();
    if (endHr == D2DERR_RECREATE_TARGET)
    {
        _captionGlyphBrush.reset();
        _captionGlyphTarget.reset();
    }
}

PopupHitTest FileOperationsPopupInternal::FileOperationsPopupState::HitTest(float x, float y) const noexcept
{
    for (auto it = _buttons.rbegin(); it != _buttons.rend(); ++it)
    {
        if (PointInRectF(it->bounds, x, y))
        {
            return it->hit;
        }
    }
    return {};
}

void FileOperationsPopupInternal::FileOperationsPopupState::Invalidate(HWND hwnd) const noexcept
{
    if (hwnd)
    {
        InvalidateRect(hwnd, nullptr, FALSE);
    }
}

bool FileOperationsPopupInternal::FileOperationsPopupState::ConfirmCancelAll(HWND hwnd) noexcept
{
    if (! hwnd)
    {
        return false;
    }

    if (! hostLifetime.lock())
    {
        return true;
    }

    if (! fileOps || ! fileOps->HasActiveOperations())
    {
        return true;
    }

    const std::wstring title   = LoadStringResource(nullptr, IDS_CAPTION_FILEOPS_CANCEL_ALL);
    const std::wstring message = LoadStringResource(nullptr, IDS_MSG_FILEOPS_CANCEL_ALL_POPUP);

    HostPromptRequest prompt{};
    prompt.version       = 1;
    prompt.sizeBytes     = sizeof(prompt);
    prompt.scope         = HOST_ALERT_SCOPE_WINDOW;
    prompt.severity      = HOST_ALERT_INFO;
    prompt.buttons       = HOST_PROMPT_BUTTONS_OK_CANCEL;
    prompt.targetWindow  = hwnd;
    prompt.title         = title.c_str();
    prompt.message       = message.c_str();
    prompt.defaultResult = HOST_PROMPT_RESULT_OK;

    HostPromptResult promptResult = HOST_PROMPT_RESULT_NONE;
    const HRESULT hrPrompt        = HostShowPrompt(prompt, nullptr, &promptResult);
    if (FAILED(hrPrompt) || promptResult != HOST_PROMPT_RESULT_OK)
    {
        return false;
    }

    if (fileOps)
    {
        fileOps->CancelAll();
    }

    return true;
}

void FileOperationsPopupInternal::FileOperationsPopupState::ShowSpeedLimitMenu(HWND hwnd, uint64_t taskId) noexcept
{
    if (! hwnd || ! fileOps || ! folderWindow)
    {
        return;
    }

    if (! hostLifetime.lock())
    {
        return;
    }

    FolderWindow::FileOperationState::Task* task = fileOps->FindTask(taskId);
    if (! task)
    {
        return;
    }

    const FileSystemOperation operation = task->GetOperation();
    if (operation != FILESYSTEM_COPY && operation != FILESYSTEM_MOVE)
    {
        return;
    }

    const uint64_t currentLimit = task->_desiredSpeedLimitBytesPerSecond.load(std::memory_order_acquire);

    constexpr UINT kCmdUnlimited  = 1u;
    constexpr UINT kCmdCustom     = 2u;
    constexpr UINT kCmdPresetBase = 10u;

    static constexpr std::array<uint64_t, 6> kPresets = {{
        1ull * 1024ull * 1024ull,
        5ull * 1024ull * 1024ull,
        10ull * 1024ull * 1024ull,
        50ull * 1024ull * 1024ull,
        100ull * 1024ull * 1024ull,
        1ull * 1024ull * 1024ull * 1024ull,
    }};

    std::vector<RedSalamander::DxUi::MenuFlyoutItem> items;
    items.reserve(kPresets.size() + 4u);

    auto appendRadioItem = [&](UINT commandId, std::wstring text, bool checked) noexcept
    {
        RedSalamander::DxUi::MenuFlyoutItem item{};
        item.kind      = RedSalamander::DxUi::MenuItemKind::Radio;
        item.text      = std::move(text);
        item.commandId = static_cast<int>(commandId);
        item.checked   = checked;
        items.push_back(std::move(item));
    };

    appendRadioItem(kCmdUnlimited, LoadStringResource(nullptr, IDS_FILEOP_SPEED_LIMIT_MENU_UNLIMITED), currentLimit == 0);

    RedSalamander::DxUi::MenuFlyoutItem separator{};
    separator.kind = RedSalamander::DxUi::MenuItemKind::Separator;
    items.push_back(separator);

    for (size_t i = 0; i < kPresets.size(); ++i)
    {
        const uint64_t bytesPerSecond = kPresets[i];
        appendRadioItem(kCmdPresetBase + static_cast<UINT>(i),
                        FormatStringResource(nullptr, IDS_FMT_FILEOP_SPEED_LIMIT_MENU_BYTES, FormatBytesCompact(bytesPerSecond)),
                        currentLimit == bytesPerSecond);
    }

    items.push_back(separator);

    RedSalamander::DxUi::MenuFlyoutItem customItem{};
    customItem.text      = LoadStringResource(nullptr, IDS_FILEOP_SPEED_LIMIT_MENU_CUSTOM);
    customItem.commandId = static_cast<int>(kCmdCustom);
    items.push_back(std::move(customItem));

    POINT pt{};
    GetCursorPos(&pt);

    const auto chosenOpt = RedSalamander::DxUi::ContextMenu::Show(hwnd, pt, items, MakeAppThemeDxPalette(folderWindow->GetTheme()));
    if (! chosenOpt.has_value())
    {
        return;
    }
    const UINT chosen = static_cast<UINT>(chosenOpt.value());

    uint64_t newLimit = currentLimit;
    if (chosen == kCmdUnlimited)
    {
        newLimit = 0;
    }
    else if (chosen >= kCmdPresetBase && chosen < (kCmdPresetBase + static_cast<UINT>(kPresets.size())))
    {
        const size_t index = static_cast<size_t>(chosen - kCmdPresetBase);
        newLimit           = kPresets[index];
    }
    else if (chosen == kCmdCustom)
    {
        const auto promptResult = ShowCustomSpeedLimitPrompt(hwnd, folderWindow->GetTheme(), currentLimit);
        if (! promptResult.has_value())
        {
            return;
        }

        newLimit = promptResult.value();
    }

    task->SetDesiredSpeedLimit(newLimit);
}

bool FileOperationsPopupInternal::FileOperationsPopupState::ShowCustomSpeedLimitPromptForTask(HWND hwnd, uint64_t requestedTaskId) noexcept
{
    if (! fileOps || ! folderWindow)
    {
        return false;
    }

    FolderWindow::FileOperationState::Task* task = requestedTaskId != 0 ? fileOps->FindTask(requestedTaskId) : nullptr;
    if (! task)
    {
        const std::vector<TaskSnapshot> snapshot = BuildSnapshot();
        const auto isActiveFileOperation         = [](const TaskSnapshot& candidate) noexcept
        { return candidate.kind == TaskSnapshot::Kind::FileOperation && ! candidate.finished && candidate.taskId != 0; };

        const auto activeIt = std::find_if(snapshot.begin(), snapshot.end(), isActiveFileOperation);
        if (activeIt != snapshot.end())
        {
            task = fileOps->FindTask(activeIt->taskId);
        }
    }

    if (! task)
    {
        return false;
    }

    const uint64_t currentLimit = task->_desiredSpeedLimitBytesPerSecond.load(std::memory_order_acquire);
    const auto promptResult     = ShowCustomSpeedLimitPrompt(hwnd, folderWindow->GetTheme(), currentLimit);
    if (promptResult.has_value())
    {
        task->SetDesiredSpeedLimit(promptResult.value());
    }
    Invalidate(hwnd);
    return true;
}

void FileOperationsPopupInternal::FileOperationsPopupState::ShowDestinationMenu(HWND hwnd, uint64_t taskId) noexcept
{
    if (! hwnd || ! fileOps || ! folderWindow)
    {
        return;
    }

    if (! hostLifetime.lock())
    {
        return;
    }

    FolderWindow::FileOperationState::Task* task = fileOps->FindTask(taskId);
    if (! task)
    {
        return;
    }

    const FileSystemOperation operation = task->GetOperation();
    if (operation != FILESYSTEM_COPY && operation != FILESYSTEM_MOVE)
    {
        return;
    }

    if (task->HasStarted())
    {
        return;
    }

    const std::optional<FolderWindow::Pane> destinationPaneOpt = task->GetDestinationPane();
    if (! destinationPaneOpt.has_value())
    {
        return;
    }

    const FolderWindow::Pane destinationPane                  = destinationPaneOpt.value();
    const std::optional<std::filesystem::path> otherPanelPath = folderWindow->GetCurrentPluginPath(destinationPane);
    const std::vector<std::filesystem::path> history          = folderWindow->GetFolderHistory(destinationPane);

    constexpr UINT kCmdOtherPanel  = 1u;
    constexpr UINT kCmdHistoryBase = 10u;

    const std::filesystem::path currentDestination = task->GetDestinationFolder();

    NavigationLocation::Location destinationLocation;
    const std::optional<std::filesystem::path> displayDestination = folderWindow->GetCurrentPath(destinationPane);
    if (displayDestination.has_value())
    {
        static_cast<void>(NavigationLocation::TryParseLocation(displayDestination.value().wstring(), destinationLocation));
    }

    struct DestinationEntry
    {
        std::filesystem::path folder;
        std::wstring label;
    };

    std::vector<DestinationEntry> entries;
    entries.reserve(history.size());

    for (const auto& h : history)
    {
        if (h.empty())
        {
            continue;
        }

        NavigationLocation::Location parsed;
        if (! NavigationLocation::TryParseLocation(h.wstring(), parsed))
        {
            continue;
        }

        const bool destIsFile  = NavigationLocation::IsFilePluginShortId(destinationLocation.pluginShortId);
        const bool entryIsFile = NavigationLocation::IsFilePluginShortId(parsed.pluginShortId);
        if (destIsFile != entryIsFile)
        {
            continue;
        }

        if (! destIsFile)
        {
            if (! NavigationLocation::EqualsNoCase(parsed.pluginShortId, destinationLocation.pluginShortId))
            {
                continue;
            }

            if (! NavigationLocation::EqualsNoCase(parsed.instanceContext, destinationLocation.instanceContext))
            {
                continue;
            }
        }

        if (parsed.pluginPath.empty())
        {
            continue;
        }

        DestinationEntry entry{};
        entry.folder = parsed.pluginPath;
        entry.label  = h.wstring();
        entries.push_back(std::move(entry));
    }

    std::vector<RedSalamander::DxUi::MenuFlyoutItem> items;
    items.reserve(entries.size() + 2u);

    RedSalamander::DxUi::MenuFlyoutItem otherPanelItem{};
    otherPanelItem.kind      = RedSalamander::DxUi::MenuItemKind::Radio;
    otherPanelItem.text      = LoadStringResource(nullptr, IDS_FILEOP_DEST_OTHER_PANEL);
    otherPanelItem.commandId = static_cast<int>(kCmdOtherPanel);
    otherPanelItem.checked   = otherPanelPath.has_value() && otherPanelPath.value() == currentDestination;
    items.push_back(std::move(otherPanelItem));

    if (! entries.empty())
    {
        RedSalamander::DxUi::MenuFlyoutItem separator{};
        separator.kind = RedSalamander::DxUi::MenuItemKind::Separator;
        items.push_back(separator);
    }

    for (size_t i = 0; i < entries.size(); ++i)
    {
        RedSalamander::DxUi::MenuFlyoutItem item{};
        item.kind      = RedSalamander::DxUi::MenuItemKind::Radio;
        item.text      = entries[i].label;
        item.commandId = static_cast<int>(kCmdHistoryBase + static_cast<UINT>(i));
        item.checked   = entries[i].folder == currentDestination;
        items.push_back(std::move(item));
    }

    POINT pt{};
    GetCursorPos(&pt);

    const auto chosenOpt = RedSalamander::DxUi::ContextMenu::Show(hwnd, pt, items, MakeAppThemeDxPalette(folderWindow->GetTheme()));
    if (! chosenOpt.has_value())
    {
        return;
    }
    const UINT chosen = static_cast<UINT>(chosenOpt.value());

    if (chosen == kCmdOtherPanel)
    {
        if (otherPanelPath.has_value())
        {
            task->SetDestinationFolder(otherPanelPath.value());
        }
        return;
    }

    if (chosen >= kCmdHistoryBase && chosen < (kCmdHistoryBase + static_cast<UINT>(entries.size())))
    {
        const size_t index = static_cast<size_t>(chosen - kCmdHistoryBase);
        task->SetDestinationFolder(entries[index].folder);
    }
}

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnCreate(HWND hwnd) noexcept
{
    _dpi = GetDpiForWindow(hwnd);

    if (folderWindow && hostLifetime.lock())
    {
        ApplyWindowChromeTheme(hwnd, folderWindow->GetTheme(), WindowBackdropTarget::Tool, GetActiveWindow() == hwnd);
    }
    ApplyScrollBarTheme(hwnd);
    ShowScrollBar(hwnd, SB_VERT, FALSE);
    _scrollBarVisible = false;

    UpdateLastPopupRect(hwnd);

    SetTimer(hwnd, kFileOperationsPopupTimerId, kFileOperationsPopupTimerIntervalMs, nullptr);
    return 0;
}

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnThemeChanged(HWND hwnd) noexcept
{
    if (_inThemeChange)
    {
        return 0;
    }

    _inThemeChange      = true;
    auto clearThemeFlag = wil::scope_exit([&] { _inThemeChange = false; });

    DiscardDeviceResources();

    if (folderWindow && hostLifetime.lock())
    {
        ApplyWindowChromeTheme(hwnd, folderWindow->GetTheme(), WindowBackdropTarget::Tool, GetActiveWindow() == hwnd);
    }
    ApplyScrollBarTheme(hwnd);

    RedrawWindow(hwnd, nullptr, nullptr, RDW_FRAME | RDW_NOERASE | RDW_NOCHILDREN);
    Invalidate(hwnd);
    return 0;
}

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnNcDestroy(HWND hwnd) noexcept
{
    KillTimer(hwnd, kFileOperationsPopupTimerId);

    if (fileOps && hostLifetime.lock())
    {
        fileOps->OnPopupDestroyed(hwnd);
    }

    DiscardDeviceResources();

    _headerFormat.reset();
    _bodyFormat.reset();
    _smallFormat.reset();
    _buttonFormat.reset();
    _buttonSmallFormat.reset();
    _graphOverlayFormat.reset();
    _statusIconFormat.reset();
    _statusIconFallbackFormat.reset();
    _captionGlyphFormat.reset();
    _captionGlyphFallbackFormat.reset();
    _dwriteFactory.reset();
    _d2dFactory.reset();

    SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
    _deletePending = true;
    if (_dispatchDepth == 0u)
    {
        delete this;
    }
    return 0;
}

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnSize(HWND hwnd, UINT width, UINT height) noexcept
{
    _clientSize.cx = static_cast<LONG>(width);
    _clientSize.cy = static_cast<LONG>(height);

    if (_target)
    {
        _target->Resize(D2D1::SizeU(width, height));
    }

    UpdateLastPopupRect(hwnd);
    Invalidate(hwnd);
    return 0;
}

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnDpiChanged(HWND hwnd, UINT newDpi, const RECT& suggested) noexcept
{
    _dpi = newDpi;

    _headerFormat.reset();
    _bodyFormat.reset();
    _smallFormat.reset();
    _buttonFormat.reset();
    _buttonSmallFormat.reset();
    _graphOverlayFormat.reset();
    _statusIconFormat.reset();
    _statusIconFallbackFormat.reset();
    _captionGlyphFormat.reset();
    _captionGlyphFallbackFormat.reset();
    _captionGlyphDpi = 0;

    if (_target)
    {
        _target->SetDpi(96.0f, 96.0f);
    }

    SetWindowPos(hwnd,
                 nullptr,
                 suggested.left,
                 suggested.top,
                 std::max(0L, suggested.right - suggested.left),
                 std::max(0L, suggested.bottom - suggested.top),
                 SWP_NOZORDER | SWP_NOACTIVATE);

    _maxAutoSizedWindowHeight = std::max(0L, suggested.bottom - suggested.top);

    UpdateLastPopupRect(hwnd);
    Invalidate(hwnd);
    return 0;
}

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnGetMinMaxInfo(HWND hwnd, MINMAXINFO* info) noexcept
{
    if (! hwnd || ! info)
    {
        return 0;
    }

    const UINT dpiForWindow = GetDpiForWindow(hwnd);

    constexpr int kMinClientWidthDip  = 480;
    constexpr int kMinClientHeightDip = 320;

    const int minClientW = DipsToPixels(kMinClientWidthDip, dpiForWindow);
    const int minClientH = DipsToPixels(kMinClientHeightDip, dpiForWindow);

    const DWORD style   = static_cast<DWORD>(GetWindowLongPtrW(hwnd, GWL_STYLE));
    const DWORD exStyle = static_cast<DWORD>(GetWindowLongPtrW(hwnd, GWL_EXSTYLE));

    RECT rc{0, 0, minClientW, minClientH};
    AdjustWindowRectExForDpi(&rc, style, FALSE, exStyle, dpiForWindow);

    const long minTrackW = std::max(0L, rc.right - rc.left);
    const long minTrackH = std::max(0L, rc.bottom - rc.top);

    info->ptMinTrackSize.x = std::max(static_cast<LONG>(info->ptMinTrackSize.x), minTrackW);
    info->ptMinTrackSize.y = std::max(static_cast<LONG>(info->ptMinTrackSize.y), minTrackH);

    static_cast<void>(WindowMaximizeBehavior::ApplyVerticalMaximize(hwnd, *info));
    return 0;
}

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnMove(HWND hwnd) noexcept
{
    UpdateLastPopupRect(hwnd);
    return 0;
}

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnTimer(HWND hwnd, UINT_PTR timerId) noexcept
{
    if (timerId == kFileOperationsPopupTimerId)
    {
        if (! hostLifetime.lock())
        {
            DestroyWindow(hwnd);
            return 0;
        }

        if (! IsWindowVisible(hwnd) || IsIconic(hwnd))
        {
            return 0;
        }

        UpdateRates();
        Invalidate(hwnd);
    }
    return 0;
}

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnEnterSizeMove(HWND hwnd) noexcept
{
    static_cast<void>(hwnd);
    _inSizeMove = true;
    return 0;
}

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnExitSizeMove(HWND hwnd) noexcept
{
    if (hwnd)
    {
        RECT rc{};
        GetWindowRect(hwnd, &rc);
        _maxAutoSizedWindowHeight = std::max(0L, rc.bottom - rc.top);
    }
    _inSizeMove = false;
    return 0;
}

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnVScroll(HWND hwnd, UINT request) noexcept
{
    if (! hwnd)
    {
        return 0;
    }

    SCROLLINFO si{};
    si.cbSize = sizeof(si);
    si.fMask  = SIF_ALL;
    if (! GetScrollInfo(hwnd, SB_VERT, &si))
    {
        return 0;
    }

    const int page     = std::max(1, static_cast<int>(si.nPage));
    const int maxPos   = std::max(0, si.nMax - page + 1);
    const int lineStep = std::max(1, DipsToPixels(36, _dpi));
    const int pageStep = page;

    int newPos = _scrollPos;
    switch (request)
    {
        case SB_TOP: newPos = 0; break;
        case SB_BOTTOM: newPos = maxPos; break;
        case SB_LINEUP: newPos -= lineStep; break;
        case SB_LINEDOWN: newPos += lineStep; break;
        case SB_PAGEUP: newPos -= pageStep; break;
        case SB_PAGEDOWN: newPos += pageStep; break;
        case SB_THUMBTRACK:
        case SB_THUMBPOSITION: newPos = si.nTrackPos; break;
        default: return 0;
    }

    newPos = std::clamp(newPos, 0, maxPos);
    if (newPos == _scrollPos)
    {
        return 0;
    }

    _scrollPos = newPos;

    SCROLLINFO set{};
    set.cbSize = sizeof(set);
    set.fMask  = SIF_POS;
    set.nPos   = _scrollPos;
    SetScrollInfo(hwnd, SB_VERT, &set, TRUE);

    Invalidate(hwnd);
    return 0;
}

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnMouseMove(HWND hwnd, POINT pt) noexcept
{
    if (! _trackingMouse)
    {
        TRACKMOUSEEVENT tme{};
        tme.cbSize    = sizeof(tme);
        tme.dwFlags   = TME_LEAVE;
        tme.hwndTrack = hwnd;
        TrackMouseEvent(&tme);
        _trackingMouse = true;
    }

    const PopupHitTest hit = HitTest(static_cast<float>(pt.x), static_cast<float>(pt.y));
    if (hit.kind != _hotHit.kind || hit.taskId != _hotHit.taskId || hit.data != _hotHit.data)
    {
        _hotHit = hit;
        Invalidate(hwnd);
    }
    return 0;
}

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnMouseLeave(HWND hwnd) noexcept
{
    _trackingMouse = false;
    if (_hotHit.kind != PopupHitTest::Kind::None)
    {
        _hotHit = {};
        Invalidate(hwnd);
    }
    return 0;
}

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnLButtonDown(HWND hwnd, POINT pt) noexcept
{
    SetCapture(hwnd);
    _pressedHit = HitTest(static_cast<float>(pt.x), static_cast<float>(pt.y));
    _hotHit     = _pressedHit;
    Invalidate(hwnd);
    return 0;
}

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnLButtonUp(HWND hwnd, POINT pt) noexcept
{
    ReleaseCapture();

    const PopupHitTest released = HitTest(static_cast<float>(pt.x), static_cast<float>(pt.y));
    const bool activated        = _pressedHit.kind != PopupHitTest::Kind::None && _pressedHit.kind == released.kind && _pressedHit.taskId == released.taskId &&
                                  _pressedHit.data == released.data;
    const PopupHitTest hit      = _pressedHit;
    _pressedHit                 = {};

    if (! activated)
    {
        return 0;
    }

    return OnActivatedHit(hwnd, hit);
}

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnActivatedHit(HWND hwnd, const PopupHitTest& hit) noexcept
{
    if (! hostLifetime.lock())
    {
        DestroyWindow(hwnd);
        return 0;
    }

    if (hit.kind == PopupHitTest::Kind::FooterCancelAll)
    {
        if (fileOps && ! fileOps->HasActiveOperations())
        {
            std::vector<FolderWindow::FileOperationState::CompletedTaskSummary> completed;
            fileOps->CollectCompletedTasks(completed);
            for (const auto& summary : completed)
            {
                fileOps->DismissCompletedTask(summary.taskId);
            }
            Invalidate(hwnd);
            return 0;
        }

        static_cast<void>(ConfirmCancelAll(hwnd));
        return 0;
    }

    if (hit.kind == PopupHitTest::Kind::FooterQueueMode)
    {
        if (fileOps)
        {
            const bool queueMode    = fileOps->GetQueueNewTasks();
            const bool newQueueMode = ! queueMode;
            fileOps->ApplyQueueMode(newQueueMode);
        }
        Invalidate(hwnd);
        return 0;
    }

    if (hit.kind == PopupHitTest::Kind::TaskToggleCollapse)
    {
        ToggleTaskCollapsed(hit.taskId);
        Invalidate(hwnd);
        return 0;
    }

    if (hit.kind == PopupHitTest::Kind::TaskPause)
    {
        if (fileOps)
        {
            FolderWindow::FileOperationState::Task* task = fileOps->FindTask(hit.taskId);
            if (task)
            {
                task->TogglePause();
            }
        }
        Invalidate(hwnd);
        return 0;
    }

    if (hit.kind == PopupHitTest::Kind::TaskCancel)
    {
        if (fileOps)
        {
            FolderWindow::FileOperationState::Task* task = fileOps->FindTask(hit.taskId);
            if (task)
            {
                task->RequestCancel();
            }
        }
        Invalidate(hwnd);
        return 0;
    }

    if (hit.kind == PopupHitTest::Kind::FooterAutoDismissSuccess)
    {
        if (fileOps)
        {
            const bool enabled = fileOps->GetAutoDismissSuccess();
            fileOps->SetAutoDismissSuccess(! enabled);
        }
        Invalidate(hwnd);
        return 0;
    }

    if (hit.kind == PopupHitTest::Kind::TaskDismiss)
    {
        if (fileOps)
        {
            fileOps->DismissCompletedTask(hit.taskId);
            fileOps->DismissInformationalTask(hit.taskId);
        }
        Invalidate(hwnd);
        return 0;
    }

    if (hit.kind == PopupHitTest::Kind::TaskShowLog)
    {
        if (fileOps)
        {
            static_cast<void>(fileOps->OpenDiagnosticsLogForTask(hit.taskId));
        }
        Invalidate(hwnd);
        return 0;
    }

    if (hit.kind == PopupHitTest::Kind::TaskExportIssues)
    {
        if (fileOps)
        {
            static_cast<void>(fileOps->ExportTaskIssuesReport(hit.taskId));
        }
        Invalidate(hwnd);
        return 0;
    }

    if (hit.kind == PopupHitTest::Kind::TaskSkip)
    {
        if (fileOps)
        {
            FolderWindow::FileOperationState::Task* task = fileOps->FindTask(hit.taskId);
            if (task)
            {
                task->SkipPreCalculation();
            }
        }
        Invalidate(hwnd);
        return 0;
    }

    if (hit.kind == PopupHitTest::Kind::TaskSpeedLimit)
    {
        ShowSpeedLimitMenu(hwnd, hit.taskId);
        Invalidate(hwnd);
        return 0;
    }

    if (hit.kind == PopupHitTest::Kind::TaskDestination)
    {
        ShowDestinationMenu(hwnd, hit.taskId);
        Invalidate(hwnd);
        return 0;
    }

    if (hit.kind == PopupHitTest::Kind::TaskConflictToggleApplyToAll)
    {
        if (fileOps)
        {
            FolderWindow::FileOperationState::Task* task = fileOps->FindTask(hit.taskId);
            if (task)
            {
                task->ToggleConflictApplyToAllChecked();
            }
        }
        Invalidate(hwnd);
        return 0;
    }

    if (hit.kind == PopupHitTest::Kind::TaskConflictAction)
    {
        if (fileOps)
        {
            FolderWindow::FileOperationState::Task* task = fileOps->FindTask(hit.taskId);
            if (task)
            {
                bool applyToAll = false;
                {
                    std::scoped_lock lock(task->_conflictMutex);
                    applyToAll = task->_conflictPrompt.applyToAllChecked;
                }

                const auto action = static_cast<FolderWindow::FileOperationState::Task::ConflictAction>(hit.data);
                task->SubmitConflictDecision(action, applyToAll);
            }
        }
        Invalidate(hwnd);
        return 0;
    }

    return 0;
}

#ifdef ENABLE_TESTS
LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnSelfTestInvoke(HWND hwnd, const PopupSelfTestInvoke* payload) noexcept
{
    if (! payload)
    {
        return 0;
    }

    if (payload->kind == PopupHitTest::Kind::TaskSpeedLimit && payload->data == 1u)
    {
        return ShowCustomSpeedLimitPromptForTask(hwnd, payload->taskId) ? 1 : 0;
    }

    return OnActivatedHit(hwnd, PopupHitTest{payload->kind, payload->taskId, payload->data});
}

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnTaskSnapshotRequest(const PopupTaskSnapshotRequest* request) const noexcept
{
    if (! request)
    {
        return 0;
    }

    PopupTaskSnapshotRequest& mutableRequest = *const_cast<PopupTaskSnapshotRequest*>(request);
    mutableRequest.found                     = false;

    const std::vector<TaskSnapshot> snapshot = BuildSnapshot();
    auto it = mutableRequest.taskId == 0
                  ? std::find_if(snapshot.begin(),
                                 snapshot.end(),
                                 [](const TaskSnapshot& task) noexcept { return task.kind == TaskSnapshot::Kind::FileOperation && ! task.finished; })
                  : std::find_if(snapshot.begin(), snapshot.end(), [&mutableRequest](const TaskSnapshot& task) noexcept {
        return task.taskId == mutableRequest.taskId;
    });
    if (mutableRequest.taskId == 0 && it == snapshot.end())
    {
        it = std::find_if(snapshot.begin(), snapshot.end(), [](const TaskSnapshot& task) noexcept { return task.kind == TaskSnapshot::Kind::FileOperation; });
    }
    if (it == snapshot.end())
    {
        return 0;
    }

    mutableRequest.snapshot = *it;
    mutableRequest.found    = true;
    return 1;
}

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnCaptionGlyphSnapshotRequest(CaptionGlyphDebugSnapshot* snapshot) const noexcept
{
    if (! snapshot)
    {
        return 0;
    }

    const bool highContrast                 = folderWindow && folderWindow->GetTheme().highContrast;
    snapshot->statusVisible                 = _captionStatus != CaptionStatus::None && ! highContrast;
    snapshot->highContrastSuppressed        = _captionStatus != CaptionStatus::None && highContrast;
    snapshot->usesDirectWriteGlyphRendering = true;
    snapshot->usesGdiTextFallback           = false;
    return 1;
}
#endif

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnMouseWheel(HWND hwnd, int delta) noexcept
{
    const int step = std::max(1, DipsToPixels(36, _dpi));
    _mouseWheelRemainder += delta;

    const int steps      = _mouseWheelRemainder / WHEEL_DELTA;
    _mouseWheelRemainder = _mouseWheelRemainder % WHEEL_DELTA;

    SCROLLINFO si{};
    si.cbSize = sizeof(si);
    si.fMask  = SIF_RANGE | SIF_PAGE | SIF_POS;
    if (GetScrollInfo(hwnd, SB_VERT, &si))
    {
        const int page   = std::max(1, static_cast<int>(si.nPage));
        const int maxPos = std::max(0, si.nMax - page + 1);
        _scrollPos       = std::clamp(_scrollPos - steps * step, 0, maxPos);

        SCROLLINFO set{};
        set.cbSize = sizeof(set);
        set.fMask  = SIF_POS;
        set.nPos   = _scrollPos;
        SetScrollInfo(hwnd, SB_VERT, &set, TRUE);
    }

    Invalidate(hwnd);
    return 0;
}

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnClose(HWND hwnd) noexcept
{
    if (ConfirmCancelAll(hwnd))
    {
        DestroyWindow(hwnd);
    }
    return 0;
}

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnNcPaint(HWND hwnd, WPARAM wParam, LPARAM lParam) noexcept
{
    const bool capturePerf   = Debug::Perf::IsCaptureEnabled();
    const uint64_t startedUs = capturePerf ? PerfNowUs() : 0u;
    const LRESULT result     = DefWindowProcW(hwnd, WM_NCPAINT, wParam, lParam);
    PaintCaptionStatusGlyph(hwnd);
    if (capturePerf)
    {
        Debug::Perf::Emit(L"FileOps.Popup.WmNcPaintUs", L"", PerfElapsedUs(startedUs), 0u, 0u, result >= 0 ? S_OK : E_FAIL);
    }
    return result;
}

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnNcActivate(HWND hwnd, WPARAM wParam, LPARAM lParam) noexcept
{
    const bool capturePerf   = Debug::Perf::IsCaptureEnabled();
    const uint64_t startedUs = capturePerf ? PerfNowUs() : 0u;
    if (folderWindow && hostLifetime.lock())
    {
        ApplyTitleBarTheme(hwnd, folderWindow->GetTheme(), wParam != FALSE);
    }

    const LRESULT result = DefWindowProcW(hwnd, WM_NCACTIVATE, wParam, lParam);
    PaintCaptionStatusGlyph(hwnd);
    if (capturePerf)
    {
        Debug::Perf::Emit(L"FileOps.Popup.WmNcActivateUs", L"", PerfElapsedUs(startedUs), wParam != FALSE ? 1u : 0u, 0u, result >= 0 ? S_OK : E_FAIL);
    }
    return result;
}

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    switch (msg)
    {
        case WM_CREATE: return OnCreate(hwnd);
        case WM_NCDESTROY: return OnNcDestroy(hwnd);
        case WM_NCACTIVATE: return OnNcActivate(hwnd, wp, lp);
        case WM_NCPAINT: return OnNcPaint(hwnd, wp, lp);
        case WM_ERASEBKGND: return 1;
        case WM_PAINT:
        {
            const bool capturePerf   = Debug::Perf::IsCaptureEnabled();
            const uint64_t startedUs = capturePerf ? PerfNowUs() : 0u;
            Render(hwnd);
            if (capturePerf)
            {
                Debug::Perf::Emit(L"FileOps.Popup.WmPaintUs", L"", PerfElapsedUs(startedUs), 0u, 0u, S_OK);
            }
            return 0;
        }
        case WM_SIZE: return OnSize(hwnd, LOWORD(lp), HIWORD(lp));
        case WM_MOVE: return OnMove(hwnd);
        case WM_GETMINMAXINFO: return OnGetMinMaxInfo(hwnd, reinterpret_cast<MINMAXINFO*>(lp));
        case WM_ENTERSIZEMOVE: return OnEnterSizeMove(hwnd);
        case WM_EXITSIZEMOVE: return OnExitSizeMove(hwnd);
        case WM_TIMER: return OnTimer(hwnd, static_cast<UINT_PTR>(wp));
        case WM_VSCROLL: return OnVScroll(hwnd, static_cast<UINT>(LOWORD(wp)));
        case WM_MOUSEMOVE: return OnMouseMove(hwnd, {GET_X_LPARAM(lp), GET_Y_LPARAM(lp)});
        case WM_MOUSELEAVE: return OnMouseLeave(hwnd);
        case WM_LBUTTONDOWN: return OnLButtonDown(hwnd, {GET_X_LPARAM(lp), GET_Y_LPARAM(lp)});
        case WM_LBUTTONUP: return OnLButtonUp(hwnd, {GET_X_LPARAM(lp), GET_Y_LPARAM(lp)});
        case WM_MOUSEWHEEL: return OnMouseWheel(hwnd, GET_WHEEL_DELTA_WPARAM(wp));
        case kFileOperationsPopupDeferredSpeedLimitPromptMessage:
        {
            return ShowCustomSpeedLimitPromptForTask(hwnd, static_cast<uint64_t>(lp)) ? 1 : 0;
        }
        case WM_DPICHANGED:
        {
            const auto* suggested = reinterpret_cast<const RECT*>(lp);
            return suggested ? OnDpiChanged(hwnd, LOWORD(wp), *suggested) : 0;
        }
        case WM_THEMECHANGED:
        case WM_SYSCOLORCHANGE: return OnThemeChanged(hwnd);
        case WM_CLOSE: return OnClose(hwnd);
#ifdef ENABLE_TESTS
        case WndMsg::kFileOpsPopupSelfTestInvoke: return OnSelfTestInvoke(hwnd, reinterpret_cast<const PopupSelfTestInvoke*>(lp));
        case WndMsg::kFileOpsPopupSelfTestSnapshot: return OnTaskSnapshotRequest(reinterpret_cast<const PopupTaskSnapshotRequest*>(lp));
        case WndMsg::kFileOpsPopupCaptionGlyphSnapshot: return OnCaptionGlyphSnapshotRequest(reinterpret_cast<CaptionGlyphDebugSnapshot*>(lp));
#endif
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

namespace
{
} // namespace

LRESULT CALLBACK FileOperationsPopupInternal::FileOperationsPopupState::WndProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    FileOperationsPopupState* state = nullptr;

    if (msg == WM_NCCREATE)
    {
        auto* cs = reinterpret_cast<CREATESTRUCTW*>(lp);
        state    = cs ? reinterpret_cast<FileOperationsPopupState*>(cs->lpCreateParams) : nullptr;
        SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(state));
    }
    else
    {
        state = reinterpret_cast<FileOperationsPopupState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    }

    if (! state)
    {
        return DefWindowProcW(hwnd, msg, wp, lp);
    }

    ++state->_dispatchDepth;
    const auto finishDispatch = wil::scope_exit([state]() noexcept
    {
        if (state->_dispatchDepth > 0u)
        {
            --state->_dispatchDepth;
        }
        if (state->_dispatchDepth == 0u && state->_deletePending)
        {
            delete state;
        }
    });

    return state->WndProc(hwnd, msg, wp, lp);
}

HWND FileOperationsPopup::Create(FolderWindow::FileOperationState* fileOps,
                                 FolderWindow* folderWindow,
                                 HWND ownerWindow,
                                 std::weak_ptr<void> hostLifetime) noexcept
{
    if (! fileOps || ! folderWindow)
    {
        return nullptr;
    }

    if (hostLifetime.expired())
    {
        return nullptr;
    }

    if (! RegisterFileOperationsPopupWndClass(GetModuleHandleW(nullptr)))
    {
        return nullptr;
    }

    auto statePtr          = std::make_unique<FileOperationsPopupInternal::FileOperationsPopupState>();
    statePtr->fileOps      = fileOps;
    statePtr->folderWindow = folderWindow;
    statePtr->hostLifetime = std::move(hostLifetime);

    const UINT ownerDpi           = ownerWindow ? GetDpiForWindow(ownerWindow) : USER_DEFAULT_SCREEN_DPI;
    const int desiredClientWidth  = DipsToPixels(480, ownerDpi);
    const int desiredClientHeight = DipsToPixels(460, ownerDpi);

    const DWORD style   = WS_POPUP | WS_CAPTION | WS_SYSMENU | WS_THICKFRAME | WS_MINIMIZEBOX | WS_MAXIMIZEBOX | WS_VSCROLL;
    const DWORD exStyle = WS_EX_APPWINDOW;

    int width  = 0;
    int height = 0;
    int x      = 0;
    int y      = 0;

    bool useSavedPlacement = false;
    RECT savedRect{};
    [[maybe_unused]] bool startMaximized = false;
    if (fileOps)
    {
        if (fileOps->TryGetPopupPlacement(savedRect, startMaximized, ownerDpi))
        {
            useSavedPlacement = true;
        }
        else
        {
            const std::optional<RECT> lastRectOpt = fileOps->GetLastPopupRect();
            if (lastRectOpt.has_value() && IsRectFullyVisible(lastRectOpt.value()))
            {
                savedRect         = lastRectOpt.value();
                useSavedPlacement = true;
            }
        }
    }

    if (useSavedPlacement)
    {
        width  = std::max(0L, savedRect.right - savedRect.left);
        height = std::max(0L, savedRect.bottom - savedRect.top);
        x      = static_cast<int>(savedRect.left);
        y      = static_cast<int>(savedRect.top);
    }
    else
    {
        const HWND monitorOwner = ownerWindow ? ownerWindow : folderWindow->GetHwnd();
        HMONITOR monitor        = MonitorFromWindow(monitorOwner, MONITOR_DEFAULTTONEAREST);
        MONITORINFO mi{};
        mi.cbSize = sizeof(mi);
        if (! GetMonitorInfoW(monitor, &mi))
        {
            return nullptr;
        }

        const RECT work = mi.rcWork;

        RECT desiredWindowRect{0, 0, desiredClientWidth, desiredClientHeight};
        AdjustWindowRectExForDpi(&desiredWindowRect, style, FALSE, exStyle, ownerDpi);
        width  = std::max(0L, desiredWindowRect.right - desiredWindowRect.left);
        height = std::max(0L, desiredWindowRect.bottom - desiredWindowRect.top);

        RECT ownerRect{};
        bool useOwnerCenter = false;
        if (ownerWindow && ! IsIconic(ownerWindow) && GetWindowRect(ownerWindow, &ownerRect))
        {
            useOwnerCenter = true;
        }

        int centerX = work.left + (work.right - work.left - width) / 2;
        int centerY = work.top + (work.bottom - work.top - height) / 2;

        if (useOwnerCenter)
        {
            const int ownerW = std::max(0L, ownerRect.right - ownerRect.left);
            const int ownerH = std::max(0L, ownerRect.bottom - ownerRect.top);
            centerX          = ownerRect.left + (ownerW - width) / 2;
            centerY          = ownerRect.top + (ownerH - height) / 2;
        }

        const int maxX = work.right - width;
        if (maxX >= work.left)
        {
            x = std::clamp(centerX, static_cast<int>(work.left), maxX);
        }
        else
        {
            x = work.left;
        }

        const int maxY = work.bottom - height;
        if (maxY >= work.top)
        {
            y = std::clamp(centerY, static_cast<int>(work.top), maxY);
        }
        else
        {
            y = work.top;
        }
    }

    const std::wstring title = LoadStringResource(nullptr, IDS_FILEOPS_POPUP_TITLE);

    // Transfer ownership to window - it will delete itself in WM_DESTROY
    auto* state = statePtr.release();
    HWND popup =
        CreateWindowExW(exStyle, kFileOperationsPopupClassName, title.c_str(), style, x, y, width, height, nullptr, nullptr, GetModuleHandleW(nullptr), state);

    if (! popup)
    {
        // Reclaim ownership via unique_ptr destructor
        std::unique_ptr<FileOperationsPopupInternal::FileOperationsPopupState> reclaimed(state);
        return nullptr;
    }

    return popup;
}

#ifdef ENABLE_TESTS
HWND FindFileOperationsDebugWindowForCurrentProcess(const wchar_t* className) noexcept
{
    if (! className || *className == L'\0')
    {
        return nullptr;
    }

    struct SearchState
    {
        DWORD processId          = 0;
        const wchar_t* className = nullptr;
        HWND hwnd                = nullptr;
    } state{GetCurrentProcessId(), className, nullptr};

    EnumWindows(
        [](HWND hwnd, LPARAM lParam) noexcept -> BOOL
    {
        auto* state = reinterpret_cast<SearchState*>(lParam);
        if (! state || ! hwnd || IsWindow(hwnd) == FALSE || IsWindowVisible(hwnd) == FALSE)
        {
            return TRUE;
        }

        DWORD processId = 0;
        GetWindowThreadProcessId(hwnd, &processId);
        if (processId != state->processId)
        {
            return TRUE;
        }

        wchar_t windowClass[128]{};
        if (GetClassNameW(hwnd, windowClass, static_cast<int>(std::size(windowClass))) == 0)
        {
            return TRUE;
        }

        if (wcscmp(windowClass, state->className) != 0)
        {
            return TRUE;
        }

        state->hwnd = hwnd;
        return FALSE;
    },
        reinterpret_cast<LPARAM>(&state));

    return state.hwnd;
}

bool DebugInvokeFileOperationsPopup(HWND popup, const FileOperationsPopupInternal::PopupSelfTestInvoke& invoke) noexcept
{
    // Keep the selftest invoke synchronous: the stack payload must remain live while
    // the popup opens any modal test surface.
    return popup && IsWindow(popup) != FALSE && SendMessageW(popup, WndMsg::kFileOpsPopupSelfTestInvoke, 0, reinterpret_cast<LPARAM>(&invoke)) != 0;
}

bool DebugGetFileOperationsPopupTaskSnapshot(HWND popup, uint64_t taskId, FileOperationsPopupInternal::TaskSnapshot& out) noexcept
{
    if (! popup || IsWindow(popup) == FALSE)
    {
        return false;
    }

    FileOperationsPopupInternal::PopupTaskSnapshotRequest request{};
    request.taskId = taskId;
    const bool ok  = SendMessageW(popup, WndMsg::kFileOpsPopupSelfTestSnapshot, 0, reinterpret_cast<LPARAM>(&request)) != FALSE;
    if (ok && request.found)
    {
        out = std::move(request.snapshot);
    }
    return ok && request.found;
}

bool DebugGetFileOperationsPopupCaptionGlyphSnapshot(HWND popup, FileOperationsPopupInternal::CaptionGlyphDebugSnapshot& out) noexcept
{
    if (! popup || IsWindow(popup) == FALSE)
    {
        return false;
    }

    FileOperationsPopupInternal::CaptionGlyphDebugSnapshot snapshot{};
    const bool ok = SendMessageW(popup, WndMsg::kFileOpsPopupCaptionGlyphSnapshot, 0, reinterpret_cast<LPARAM>(&snapshot)) != FALSE;
    if (ok)
    {
        out = snapshot;
    }
    return ok;
}

float DebugComputeFileOperationsTaskCompleteFraction(const FileOperationsPopupInternal::TaskSnapshot& task) noexcept
{
    return ComputeFileOperationsTaskCompleteFractionForDisplay(task);
}

double DebugSmoothRateForDisplay(double previousRate, double sampleRate, ULONGLONG elapsedMs) noexcept
{
    return SmoothRateForDisplay(previousRate, sampleRate, elapsedMs);
}

double DebugDecayRateForCallbackSilence(double smoothedRate, ULONGLONG silenceMs) noexcept
{
    return DecayRateForCallbackSilence(smoothedRate, silenceMs);
}

double DebugSmoothEtaSecondsForDisplay(double previousEtaSeconds, double sampleEtaSeconds, ULONGLONG elapsedMs) noexcept
{
    return SmoothEtaSecondsForDisplay(previousEtaSeconds, sampleEtaSeconds, elapsedMs);
}

HWND GetFileOperationsSpeedLimitPromptHandle() noexcept
{
    const HWND hwnd = FindFileOperationsDebugWindowForCurrentProcess(kFileOperationsSpeedLimitPromptClassName);
    return hwnd && IsWindow(hwnd) != FALSE ? hwnd : nullptr;
}

bool DebugGetFileOperationsSpeedLimitPromptSnapshot(FileOperationsSpeedLimitPromptDebugSnapshot& out) noexcept
{
    const HWND hwnd = GetFileOperationsSpeedLimitPromptHandle();
    if (! hwnd)
    {
        return false;
    }

    auto snapshot = std::make_unique<FileOperationsSpeedLimitPromptDebugSnapshot>();
    const bool ok = SendMessageW(hwnd,
                                 kFileOperationsSpeedLimitPromptDebugMessage,
                                 static_cast<WPARAM>(FileOperationsSpeedLimitPromptWindow::DebugCommand::GetSnapshot),
                                 reinterpret_cast<LPARAM>(snapshot.get())) != FALSE;
    if (ok)
    {
        out = std::move(*snapshot);
    }
    return ok;
}

bool DebugSetFileOperationsSpeedLimitPromptText(std::wstring_view text) noexcept
{
    const HWND hwnd = GetFileOperationsSpeedLimitPromptHandle();
    if (! hwnd)
    {
        return false;
    }

    auto payload = std::make_unique<std::wstring>(text);
    return SendMessageW(hwnd,
                        kFileOperationsSpeedLimitPromptDebugMessage,
                        static_cast<WPARAM>(FileOperationsSpeedLimitPromptWindow::DebugCommand::SetText),
                        reinterpret_cast<LPARAM>(payload.get())) != FALSE;
}

bool DebugConfirmFileOperationsSpeedLimitPrompt() noexcept
{
    const HWND hwnd = GetFileOperationsSpeedLimitPromptHandle();
    return hwnd &&
           SendMessageW(
               hwnd, kFileOperationsSpeedLimitPromptDebugMessage, static_cast<WPARAM>(FileOperationsSpeedLimitPromptWindow::DebugCommand::Confirm), 0) != FALSE;
}

bool DebugCancelFileOperationsSpeedLimitPrompt() noexcept
{
    const HWND hwnd = GetFileOperationsSpeedLimitPromptHandle();
    return hwnd &&
           SendMessageW(
               hwnd, kFileOperationsSpeedLimitPromptDebugMessage, static_cast<WPARAM>(FileOperationsSpeedLimitPromptWindow::DebugCommand::Cancel), 0) != FALSE;
}
#endif
