#include "FolderWindow.FileOperations.Popup.h"

#include "FluentIcons.h"
#include "FolderWindow.FileOperationsInternal.h"
#include "HostServices.h"
#include "NavigationLocation.h"
#include "WindowMaximizeBehavior.h"

#include <array>
#include <cmath>
#include <d2d1.h>
#include <dwrite.h>
#include <limits>
#include <unordered_map>
#include <windowsx.h>

namespace
{
constexpr wchar_t kFileOperationsPopupClassName[] = L"RedSalamander.FileOperationsPopup";

constexpr UINT_PTR kFileOperationsPopupTimerId     = 1;
constexpr UINT kFileOperationsPopupTimerIntervalMs = 100;

constexpr std::wstring_view kEllipsisText = L"\u2026";

struct SpeedLimitDialogState
{
    SpeedLimitDialogState() = default;

    SpeedLimitDialogState& operator=(SpeedLimitDialogState&) = delete;
    SpeedLimitDialogState(SpeedLimitDialogState&)            = delete;

    unsigned __int64 initialLimitBytesPerSecond = 0;
    unsigned __int64 resultLimitBytesPerSecond  = 0;
    AppTheme theme{};
    wil::unique_any<HBRUSH, decltype(&::DeleteObject), ::DeleteObject> backgroundBrush;
    std::wstring hintText;
    bool showingValidationError = false;
};
INT_PTR CALLBACK SpeedLimitDialogProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp);

COLORREF ColorRefFromColorF(const D2D1::ColorF& color) noexcept
{
    const auto toByte = [](float v) noexcept
    {
        const float clamped = std::clamp(v, 0.0f, 1.0f);
        const float scaled  = (clamped * 255.0f) + 0.5f;
        const int asInt     = static_cast<int>(scaled);
        const int bounded   = std::clamp(asInt, 0, 255);
        return static_cast<BYTE>(bounded);
    };

    return RGB(toByte(color.r), toByte(color.g), toByte(color.b));
}

float DipsToPixels(float dip, UINT dpi) noexcept
{
    return dip * (static_cast<float>(dpi) / static_cast<float>(USER_DEFAULT_SCREEN_DPI));
}

int DipsToPixels(int dip, UINT dpi) noexcept
{
    return MulDiv(dip, static_cast<int>(dpi), USER_DEFAULT_SCREEN_DPI);
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

bool TryParseThroughputText(std::wstring_view text, unsigned __int64& outBytesPerSecond) noexcept
{
    constexpr unsigned __int64 kKiB = 1024ull;
    constexpr unsigned __int64 kMiB = 1024ull * 1024ull;
    constexpr unsigned __int64 kGiB = 1024ull * 1024ull * 1024ull;
    constexpr unsigned __int64 kTiB = 1024ull * 1024ull * 1024ull * 1024ull;
    constexpr unsigned __int64 kPiB = 1024ull * 1024ull * 1024ull * 1024ull * 1024ull;

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

        if (ch == L'.' && ! sawDecimal)
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

    unsigned __int64 multiplier = 0;
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

    constexpr double maxValue = static_cast<double>(std::numeric_limits<unsigned __int64>::max());
    if (result >= maxValue)
    {
        outBytesPerSecond = std::numeric_limits<unsigned __int64>::max();
        return true;
    }

    outBytesPerSecond = static_cast<unsigned __int64>(result + 0.5);
    return true;
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
    if (! hwnd || ! folderWindow)
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

    const wchar_t* fontName = L"Segoe UI";

    if (! _headerFormat)
    {
        _dwriteFactory->CreateTextFormat(fontName,
                                         nullptr,
                                         DWRITE_FONT_WEIGHT_SEMI_BOLD,
                                         DWRITE_FONT_STYLE_NORMAL,
                                         DWRITE_FONT_STRETCH_NORMAL,
                                         DipsToPixels(12.0f, _dpi),
                                         L"",
                                         _headerFormat.put());
    }

    if (! _bodyFormat)
    {
        _dwriteFactory->CreateTextFormat(fontName,
                                         nullptr,
                                         DWRITE_FONT_WEIGHT_NORMAL,
                                         DWRITE_FONT_STYLE_NORMAL,
                                         DWRITE_FONT_STRETCH_NORMAL,
                                         DipsToPixels(12.0f, _dpi),
                                         L"",
                                         _bodyFormat.put());
    }

    if (! _smallFormat)
    {
        _dwriteFactory->CreateTextFormat(fontName,
                                         nullptr,
                                         DWRITE_FONT_WEIGHT_NORMAL,
                                         DWRITE_FONT_STYLE_NORMAL,
                                         DWRITE_FONT_STRETCH_NORMAL,
                                         DipsToPixels(11.0f, _dpi),
                                         L"",
                                         _smallFormat.put());
    }

    if (! _buttonFormat)
    {
        _dwriteFactory->CreateTextFormat(fontName,
                                         nullptr,
                                         DWRITE_FONT_WEIGHT_NORMAL,
                                         DWRITE_FONT_STYLE_NORMAL,
                                         DWRITE_FONT_STRETCH_NORMAL,
                                         DipsToPixels(12.0f, _dpi),
                                         L"",
                                         _buttonFormat.put());
    }

    if (! _buttonSmallFormat)
    {
        _dwriteFactory->CreateTextFormat(fontName,
                                         nullptr,
                                         DWRITE_FONT_WEIGHT_NORMAL,
                                         DWRITE_FONT_STYLE_NORMAL,
                                         DWRITE_FONT_STRETCH_NORMAL,
                                         DipsToPixels(11.0f, _dpi),
                                         L"",
                                         _buttonSmallFormat.put());
    }

    if (! _graphOverlayFormat)
    {
        _dwriteFactory->CreateTextFormat(fontName,
                                         nullptr,
                                         DWRITE_FONT_WEIGHT_SEMI_BOLD,
                                         DWRITE_FONT_STYLE_NORMAL,
                                         DWRITE_FONT_STRETCH_NORMAL,
                                         DipsToPixels(14.0f, _dpi),
                                         L"",
                                         _graphOverlayFormat.put());
    }

    if (! _statusIconFallbackFormat)
    {
        _dwriteFactory->CreateTextFormat(fontName,
                                         nullptr,
                                         DWRITE_FONT_WEIGHT_SEMI_BOLD,
                                         DWRITE_FONT_STYLE_NORMAL,
                                         DWRITE_FONT_STRETCH_NORMAL,
                                         DipsToPixels(14.0f, _dpi),
                                         L"",
                                         _statusIconFallbackFormat.put());
    }

    if (! _statusIconFormat)
    {
        // Optional: Segoe Fluent Icons. If missing, fallback format draws standard Unicode glyphs.
        const HRESULT hr = _dwriteFactory->CreateTextFormat(FluentIcons::kFontFamily.data(),
                                                            nullptr,
                                                            DWRITE_FONT_WEIGHT_NORMAL,
                                                            DWRITE_FONT_STYLE_NORMAL,
                                                            DWRITE_FONT_STRETCH_NORMAL,
                                                            DipsToPixels(14.0f, _dpi),
                                                            L"",
                                                            _statusIconFormat.put());
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

void FileOperationsPopupInternal::FileOperationsPopupState::EnsureBrushes() noexcept
{
    if (! _target || ! folderWindow)
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
    std::vector<FolderWindow::FileOperationState::CompletedTaskSummary> completedTasks;
    if (fileOps)
    {
        fileOps->CollectTasks(tasks);
        fileOps->CollectCompletedTasks(completedTasks);
    }

    std::vector<TaskSnapshot> result;
    result.reserve(tasks.size() + completedTasks.size());
    std::unordered_map<uint64_t, bool> activeTaskIds;
    activeTaskIds.reserve(tasks.size());

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

        {
            std::scoped_lock lock(task->_progressMutex);
            snap.totalItems             = task->_progressTotalItems;
            snap.completedItems         = task->_progressCompletedItems;
            snap.totalBytes             = task->_progressTotalBytes;
            snap.completedBytes         = task->_progressCompletedBytes;
            snap.itemTotalBytes         = task->_progressItemTotalBytes;
            snap.itemCompletedBytes     = task->_progressItemCompletedBytes;
            snap.currentSourcePath      = task->_progressSourcePath;
            snap.currentDestinationPath = task->_progressDestinationPath;
            snap.hasProgressCallbacks   = ! task->_lastProgressCallbackSourcePath.empty() || ! task->_lastProgressCallbackDestinationPath.empty();

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
                    constexpr unsigned __int64 kClampThresholdBytes = 64ull * 1024ull;
                    const unsigned __int64 delta                    = snap.inFlightFiles[i].completedBytes - snap.inFlightFiles[i].totalBytes;
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

        snap.desiredSpeedLimitBytesPerSecond   = task->_desiredSpeedLimitBytesPerSecond.load(std::memory_order_acquire);
        snap.effectiveSpeedLimitBytesPerSecond = task->_effectiveSpeedLimitBytesPerSecond.load(std::memory_order_acquire);

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

        result.push_back(std::move(snap));
    }

    for (const auto& completed : completedTasks)
    {
        if (activeTaskIds.find(completed.taskId) != activeTaskIds.end())
        {
            continue;
        }

        TaskSnapshot snap{};
        snap.taskId                 = completed.taskId;
        snap.operation              = completed.operation;
        snap.totalItems             = completed.totalItems;
        snap.completedItems         = completed.completedItems;
        snap.totalBytes             = completed.totalBytes;
        snap.completedBytes         = completed.completedBytes;
        snap.currentSourcePath      = completed.sourcePath;
        snap.currentDestinationPath = completed.destinationPath;
        snap.destinationFolder      = completed.destinationFolder;
        snap.destinationPane        = completed.destinationPane;
        snap.started                = true;
        snap.finished               = true;
        snap.resultHr               = completed.resultHr;
        snap.warningCount           = completed.warningCount;
        snap.errorCount             = completed.errorCount;
        snap.lastDiagnosticMessage  = completed.lastDiagnosticMessage;

        if (snap.totalItems > 0)
        {
            snap.completedItems = std::min(snap.completedItems, snap.totalItems);
        }
        if (snap.totalBytes > 0)
        {
            snap.completedBytes = std::min(snap.completedBytes, snap.totalBytes);
        }

        result.push_back(std::move(snap));
    }

    return result;
}

std::vector<RateSnapshot> FileOperationsPopupInternal::FileOperationsPopupState::BuildRateSnapshot() const
{
    std::vector<FolderWindow::FileOperationState::Task*> tasks;
    if (fileOps)
    {
        fileOps->CollectTasks(tasks);
    }

    std::vector<RateSnapshot> result;
    result.reserve(tasks.size());

    for (auto* task : tasks)
    {
        if (! task)
        {
            continue;
        }

        RateSnapshot snap{};
        snap.taskId    = task->GetId();
        snap.operation = task->GetOperation();

        {
            std::scoped_lock lock(task->_progressMutex);
            snap.completedItems    = task->_progressCompletedItems;
            snap.completedBytes    = task->_progressCompletedBytes;
            snap.currentSourcePath = task->_progressSourcePath;
        }

        snap.started          = task->HasStarted();
        snap.paused           = task->IsPaused();
        snap.waitingForOthers = task->IsWaitingForOthers();
        snap.waitingInQueue   = task->IsWaitingInQueue();
        snap.queuePaused      = task->IsQueuePaused();

        result.push_back(snap);
    }

    return result;
}

void FileOperationsPopupInternal::FileOperationsPopupState::UpdateRates() noexcept
{
    const ULONGLONG nowTick                  = GetTickCount64();
    const std::vector<RateSnapshot> snapshot = BuildRateSnapshot();

    std::unordered_map<uint64_t, bool> seen;
    seen.reserve(snapshot.size());

    for (const RateSnapshot& task : snapshot)
    {
        seen[task.taskId] = true;

        RateHistory& history     = _rates[task.taskId];
        const ULONGLONG lastTick = history.lastTick;

        if (task.paused || task.queuePaused || task.waitingInQueue)
        {
            history.lastBytes = task.completedBytes;
            history.lastItems = task.completedItems;
            history.lastTick  = nowTick;
            continue;
        }

        if (lastTick != 0 && nowTick > lastTick)
        {
            const double dtSec = static_cast<double>(nowTick - lastTick) / 1000.0;
            if (dtSec > 0.0)
            {
                if (task.operation == FILESYSTEM_DELETE)
                {
                    unsigned long prevItems = history.lastItems;
                    if (task.completedItems < prevItems)
                    {
                        prevItems = task.completedItems;
                    }

                    const unsigned long deltaItems = task.completedItems - prevItems;
                    const double instItemsPerSec   = static_cast<double>(deltaItems) / dtSec;
                    const float instF              = instItemsPerSec > 0.0 ? static_cast<float>(instItemsPerSec) : 0.0f;

                    history.samples[history.writeIndex] = instF;
                    // Compute hue from the current source path to match progress bar color
                    if (task.currentSourcePath.empty())
                    {
                        history.hues[history.writeIndex] = -1.0f;
                    }
                    else
                    {
                        const uint32_t pathHash          = StableHash32(task.currentSourcePath);
                        history.hues[history.writeIndex] = static_cast<float>(pathHash % 360u);
                    }
                    history.writeIndex = (history.writeIndex + 1u) % RateHistory::kMaxSamples;
                    history.count      = std::min(RateHistory::kMaxSamples, history.count + 1u);

                    if (history.smoothedItemsPerSec <= 0.0f)
                    {
                        history.smoothedItemsPerSec = instF;
                    }
                    else
                    {
                        history.smoothedItemsPerSec = history.smoothedItemsPerSec * 0.85f + instF * 0.15f;
                    }

                    history.lastItems = task.completedItems;
                }
                else
                {
                    unsigned __int64 prevBytes = history.lastBytes;
                    if (task.completedBytes < prevBytes)
                    {
                        prevBytes = task.completedBytes;
                    }

                    const unsigned __int64 deltaBytes = task.completedBytes - prevBytes;
                    const double instBytesPerSec      = static_cast<double>(deltaBytes) / dtSec;
                    const float instF                 = instBytesPerSec > 0.0 ? static_cast<float>(instBytesPerSec) : 0.0f;

                    history.samples[history.writeIndex] = instF;
                    // Compute hue from the current source path to match progress bar color
                    if (task.currentSourcePath.empty())
                    {
                        history.hues[history.writeIndex] = -1.0f;
                    }
                    else
                    {
                        const uint32_t pathHash          = StableHash32(task.currentSourcePath);
                        history.hues[history.writeIndex] = static_cast<float>(pathHash % 360u);
                    }
                    history.writeIndex = (history.writeIndex + 1u) % RateHistory::kMaxSamples;
                    history.count      = std::min(RateHistory::kMaxSamples, history.count + 1u);

                    if (history.smoothedBytesPerSec <= 0.0f)
                    {
                        history.smoothedBytesPerSec = instF;
                    }
                    else
                    {
                        history.smoothedBytesPerSec = history.smoothedBytesPerSec * 0.85f + instF * 0.15f;
                    }

                    history.lastBytes = task.completedBytes;
                }
            }
        }
        else
        {
            history.lastBytes = task.completedBytes;
            history.lastItems = task.completedItems;
        }

        history.lastTick = nowTick;
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
                                                                               unsigned __int64 limitBytesPerSecond,
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

    EnsureTarget(hwnd);
    EnsureTextFormats();
    EnsureBrushes();

    if (! _target || ! _bgBrush || ! _textBrush || ! _borderBrush)
    {
        return;
    }

    const std::vector<TaskSnapshot> snapshot = BuildSnapshot();
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

    std::vector<float> cardHeights;
    cardHeights.reserve(snapshot.size());
    for (const TaskSnapshot& task : snapshot)
    {
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
                const auto& entry = task.inFlightFiles[i];
                const bool active = entry.totalBytes == 0 || entry.completedBytes < entry.totalBytes;
                const bool recentCompleted =
                    ! active && entry.totalBytes > 0 && entry.completedBytes >= entry.totalBytes && entry.lastUpdateTick != 0 &&
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
    AutoResizeWindow(hwnd, _contentHeight, taskCount);

    bool scrollReady = false;
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

    _buttons.clear();

    HRESULT hrEndDraw = S_OK;
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

                    float completeFraction = 0.0f;
                    if (task.operation == FILESYSTEM_DELETE)
                    {
                        if (task.totalBytes > 0 && task.completedBytes > 0)
                        {
                            completeFraction = Clamp01(static_cast<float>(static_cast<double>(task.completedBytes) / static_cast<double>(task.totalBytes)));
                        }
                        else if (task.totalItems > 0)
                        {
                            completeFraction = Clamp01(static_cast<float>(static_cast<double>(task.completedItems) / static_cast<double>(task.totalItems)));
                        }
                        else
                        {
                            completeFraction = SUCCEEDED(task.resultHr) ? 1.0f : 0.0f;
                        }
                    }
                    else
                    {
                        if (task.totalBytes > 0)
                        {
                            completeFraction = Clamp01(static_cast<float>(static_cast<double>(task.completedBytes) / static_cast<double>(task.totalBytes)));
                        }
                        else if (task.totalItems > 0)
                        {
                            completeFraction = Clamp01(static_cast<float>(static_cast<double>(task.completedItems) / static_cast<double>(task.totalItems)));
                        }
                        else
                        {
                            completeFraction = SUCCEEDED(task.resultHr) ? 1.0f : 0.0f;
                        }
                    }

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
                    const std::wstring sizeText   = FormatBytesCompact(task.preCalcTotalBytes);
                    const unsigned __int64 totalItems = static_cast<unsigned __int64>(task.preCalcFileCount) +
                                                        static_cast<unsigned __int64>(task.preCalcDirectoryCount);
                    const std::wstring countsText = FormatStringResource(nullptr,
                                                                         IDS_FMT_FILEOPS_FILES_FOLDERS,
                                                                         totalItems,
                                                                         task.preCalcFileCount,
                                                                         task.preCalcDirectoryCount);
                    const D2D1_RECT_F countsRc    = D2D1::RectF(textX, textY, textX + textMaxW, textY + lineH);
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
                        const double itemsPerSec     = history ? static_cast<double>(history->smoothedItemsPerSec) : 0.0;
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
                    const double bytesPerSec                  = history ? static_cast<double>(history->smoothedBytesPerSec) : 0.0;
                    const unsigned __int64 bytesPerSecRounded = bytesPerSec > 0.0 ? static_cast<unsigned __int64>(bytesPerSec + 0.5) : 0ull;
                    const std::wstring bytesText              = FormatBytesCompact(bytesPerSecRounded);
                    const std::wstring speedText              = FormatStringResource(nullptr, IDS_FMT_FILEOP_SPEED_BYTES, bytesText);
                    const D2D1_RECT_F speedRc                 = D2D1::RectF(textX, textY, textX + textMaxW, textY + lineH);
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

                    if (task.totalBytes > 0 && bytesPerSec > 0.0 && task.completedBytes <= task.totalBytes)
                    {
                        const unsigned __int64 remainingBytes = task.totalBytes - task.completedBytes;
                        const double secondsD                 = static_cast<double>(remainingBytes) / bytesPerSec;
                        const uint64_t seconds                = secondsD > 0.0 ? static_cast<uint64_t>(std::ceil(secondsD)) : 0u;
                        const std::wstring etaText            = FormatStringResource(nullptr, IDS_FMT_FILEOPS_ETA, FormatDurationHms(seconds));
                        const D2D1_RECT_F etaRc               = D2D1::RectF(textX, textY, textX + textMaxW, textY + lineH);
                        _target->DrawTextW(etaText.data(), static_cast<UINT32>(etaText.size()), _bodyFormat.get(), etaRc, _subTextBrush.get());
                        textY += lineH;
                    }
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
                            const auto& entry = task.inFlightFiles[j];
                            const bool active = entry.totalBytes == 0 || entry.completedBytes < entry.totalBytes;
                            const bool recentCompleted =
                                ! active && entry.totalBytes > 0 && entry.completedBytes >= entry.totalBytes && entry.lastUpdateTick != 0 &&
                                nowTick >= entry.lastUpdateTick && (nowTick - entry.lastUpdateTick) <= kCompletedInFlightGraceMs;
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
                        unsigned __int64 fileTotalBytes     = 0;
                        unsigned __int64 fileCompletedBytes = 0;

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

                        const float availableW   = std::max(0.0f, rightEdge - pathLeft);
                        const float miniBarWMin  = DipsToPixels(40.0f, _dpi);
                        const float minTextW     = DipsToPixels(48.0f, _dpi);
                        float miniBarW           = std::min(miniBarWDesired, availableW);
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
                            const float frac    = hasTotal && fileCompletedBytes <= fileTotalBytes
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

                    float yPrompt                    = rc.top;
                    const float maxW                 = std::max(0.0f, rc.right - rc.left);
                    const float maxDetailsY          = rc.bottom;

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
                            unsigned __int64 limit = 0;
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
                                _target->DrawTextW(applyText.data(),
                                                   static_cast<UINT32>(applyText.size()),
                                                   applyFormat,
                                                   labelRc,
                                                   applyBrush,
                                                   D2D1_DRAW_TEXT_OPTIONS_CLIP);
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

void FileOperationsPopupInternal::FileOperationsPopupState::PaintCaptionStatusGlyph(HWND hwnd) const noexcept
{
    if (! hwnd || ! folderWindow)
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

    wil::unique_hdc_window hdc(GetWindowDC(hwnd));
    if (! hdc)
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

    wchar_t fluentGlyph = 0;
    wchar_t fallback    = 0;
    COLORREF color      = theme.menu.text;

    switch (_captionStatus)
    {
        case CaptionStatus::Ok:
            fluentGlyph = FluentIcons::kCheckMark;
            fallback    = FluentIcons::kFallbackCheckMark;
            color       = ColorToCOLORREF(theme.accent);
            break;
        case CaptionStatus::Warning:
            fluentGlyph = FluentIcons::kWarning;
            fallback    = FluentIcons::kFallbackWarning;
            color       = ColorToCOLORREF(theme.folderView.warningText);
            break;
        case CaptionStatus::Error:
            fluentGlyph = FluentIcons::kError;
            fallback    = FluentIcons::kFallbackError;
            color       = ColorToCOLORREF(theme.folderView.errorText);
            break;
        case CaptionStatus::None:
        default: return;
    }

    const int sizeDip         = 20;
    auto iconFont             = FluentIcons::CreateFontForDpi(dpi, sizeDip);
    const bool useFluentGlyph = iconFont && FluentIcons::FontHasGlyph(hdc.get(), iconFont.get(), fluentGlyph);

    const wchar_t glyph = useFluentGlyph ? fluentGlyph : fallback;
    wchar_t glyphText[2]{glyph, 0};

    HFONT fontToUse = useFluentGlyph ? iconFont.get() : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));

    SetBkMode(hdc.get(), TRANSPARENT);
    SetTextColor(hdc.get(), color);

    auto oldFont = wil::SelectObject(hdc.get(), fontToUse);
    DrawTextW(hdc.get(), glyphText, 1, &iconRc, DT_CENTER | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX);
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
    if (! hwnd || ! fileOps)
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

    const unsigned __int64 currentLimit = task->_desiredSpeedLimitBytesPerSecond.load(std::memory_order_acquire);

    HMENU menu       = CreatePopupMenu();
    auto menuCleanup = wil::scope_exit(
        [&]
        {
            if (menu)
                DestroyMenu(menu);
        });
    if (! menu)
    {
        return;
    }

    constexpr UINT kCmdUnlimited  = 1u;
    constexpr UINT kCmdCustom     = 2u;
    constexpr UINT kCmdPresetBase = 10u;

    static constexpr std::array<unsigned __int64, 6> kPresets = {{
        1ull * 1024ull * 1024ull,
        5ull * 1024ull * 1024ull,
        10ull * 1024ull * 1024ull,
        50ull * 1024ull * 1024ull,
        100ull * 1024ull * 1024ull,
        1ull * 1024ull * 1024ull * 1024ull,
    }};

    const std::wstring unlimitedText = LoadStringResource(nullptr, IDS_FILEOP_SPEED_LIMIT_MENU_UNLIMITED);
    const UINT unlimitedFlags        = static_cast<UINT>(MF_STRING) | (currentLimit == 0 ? static_cast<UINT>(MF_CHECKED) : 0u);
    AppendMenuW(menu, unlimitedFlags, static_cast<UINT_PTR>(kCmdUnlimited), unlimitedText.c_str());
    AppendMenuW(menu, static_cast<UINT>(MF_SEPARATOR), 0u, nullptr);

    for (size_t i = 0; i < kPresets.size(); ++i)
    {
        const unsigned __int64 bytesPerSecond = kPresets[i];
        const std::wstring label              = FormatStringResource(nullptr, IDS_FMT_FILEOP_SPEED_LIMIT_MENU_BYTES, FormatBytesCompact(bytesPerSecond));
        const UINT cmd                        = kCmdPresetBase + static_cast<UINT>(i);
        const UINT flags                      = static_cast<UINT>(MF_STRING) | (currentLimit == bytesPerSecond ? static_cast<UINT>(MF_CHECKED) : 0u);
        AppendMenuW(menu, flags, static_cast<UINT_PTR>(cmd), label.c_str());
    }

    AppendMenuW(menu, static_cast<UINT>(MF_SEPARATOR), 0u, nullptr);
    const std::wstring customText = LoadStringResource(nullptr, IDS_FILEOP_SPEED_LIMIT_MENU_CUSTOM);
    AppendMenuW(menu, static_cast<UINT>(MF_STRING), static_cast<UINT_PTR>(kCmdCustom), customText.c_str());

    POINT pt{};
    GetCursorPos(&pt);

    const UINT chosen = static_cast<UINT>(
        TrackPopupMenuEx(menu, static_cast<UINT>(TPM_RETURNCMD | TPM_RIGHTBUTTON), static_cast<int>(pt.x), static_cast<int>(pt.y), hwnd, nullptr));
    if (chosen == 0)
    {
        return;
    }

    unsigned __int64 newLimit = currentLimit;
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
        SpeedLimitDialogState dlgState{};
        dlgState.initialLimitBytesPerSecond = currentLimit;
        dlgState.resultLimitBytesPerSecond  = currentLimit;
        dlgState.theme                      = folderWindow->GetTheme();

#pragma warning(push)
#pragma warning(disable : 5039) // C5039: pointer or reference to potentially throwing function passed to 'extern "C"' function
        const INT_PTR result = DialogBoxParamW(
            GetModuleHandleW(nullptr), MAKEINTRESOURCEW(IDD_FILEOP_SPEED_LIMIT_CUSTOM), hwnd, SpeedLimitDialogProc, reinterpret_cast<LPARAM>(&dlgState));
#pragma warning(pop)
        if (result != IDOK)
        {
            return;
        }

        newLimit = dlgState.resultLimitBytesPerSecond;
    }

    task->SetDesiredSpeedLimit(newLimit);
}

void FileOperationsPopupInternal::FileOperationsPopupState::ShowDestinationMenu(HWND hwnd, uint64_t taskId) noexcept
{
    if (! hwnd || ! fileOps || ! folderWindow)
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

    HMENU menu       = CreatePopupMenu();
    auto menuCleanup = wil::scope_exit(
        [&]
        {
            if (menu)
                DestroyMenu(menu);
        });
    if (! menu)
    {
        return;
    }

    constexpr UINT kCmdOtherPanel  = 1u;
    constexpr UINT kCmdHistoryBase = 10u;

    const std::filesystem::path currentDestination = task->GetDestinationFolder();
    const std::wstring otherPanelText              = LoadStringResource(nullptr, IDS_FILEOP_DEST_OTHER_PANEL);
    const bool otherPanelSelected                  = otherPanelPath.has_value() && otherPanelPath.value() == currentDestination;
    const UINT otherFlags                          = static_cast<UINT>(MF_STRING) | (otherPanelSelected ? static_cast<UINT>(MF_CHECKED) : 0u);
    AppendMenuW(menu, otherFlags, static_cast<UINT_PTR>(kCmdOtherPanel), otherPanelText.c_str());
    AppendMenuW(menu, static_cast<UINT>(MF_SEPARATOR), 0u, nullptr);

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

    for (size_t i = 0; i < entries.size(); ++i)
    {
        const UINT cmd           = kCmdHistoryBase + static_cast<UINT>(i);
        const std::wstring label = entries[i].label;
        const UINT flags         = static_cast<UINT>(MF_STRING) | (entries[i].folder == currentDestination ? static_cast<UINT>(MF_CHECKED) : 0u);
        AppendMenuW(menu, flags, static_cast<UINT_PTR>(cmd), label.c_str());
    }

    POINT pt{};
    GetCursorPos(&pt);

    const UINT chosen = static_cast<UINT>(
        TrackPopupMenuEx(menu, static_cast<UINT>(TPM_RETURNCMD | TPM_RIGHTBUTTON), static_cast<int>(pt.x), static_cast<int>(pt.y), hwnd, nullptr));
    if (chosen == 0)
    {
        return;
    }

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

    if (folderWindow)
    {
        ApplyTitleBarTheme(hwnd, folderWindow->GetTheme(), GetActiveWindow() == hwnd);
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

    if (folderWindow)
    {
        ApplyTitleBarTheme(hwnd, folderWindow->GetTheme(), GetActiveWindow() == hwnd);
    }
    ApplyScrollBarTheme(hwnd);

    RedrawWindow(hwnd, nullptr, nullptr, RDW_FRAME | RDW_NOERASE | RDW_NOCHILDREN);
    Invalidate(hwnd);
    return 0;
}

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnNcDestroy(HWND hwnd) noexcept
{
    KillTimer(hwnd, kFileOperationsPopupTimerId);

    if (fileOps)
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
    _dwriteFactory.reset();
    _d2dFactory.reset();

    SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
    delete this;
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
    const PopupHitTest hit = _pressedHit;
    _pressedHit            = {};

    if (! activated)
    {
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

#ifdef _DEBUG
LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnSelfTestInvoke(HWND hwnd, const PopupSelfTestInvoke* payload) noexcept
{
    if (! payload)
    {
        return 0;
    }

    const PopupHitTest hit{payload->kind, payload->taskId, payload->data};

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
    const LRESULT result = DefWindowProcW(hwnd, WM_NCPAINT, wParam, lParam);
    PaintCaptionStatusGlyph(hwnd);
    return result;
}

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnNcActivate(HWND hwnd, WPARAM wParam, LPARAM lParam) noexcept
{
    if (folderWindow)
    {
        ApplyTitleBarTheme(hwnd, folderWindow->GetTheme(), wParam != FALSE);
    }

    const LRESULT result = DefWindowProcW(hwnd, WM_NCACTIVATE, wParam, lParam);
    PaintCaptionStatusGlyph(hwnd);
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
        case WM_PAINT: Render(hwnd); return 0;
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
        case WM_DPICHANGED:
        {
            const auto* suggested = reinterpret_cast<const RECT*>(lp);
            return suggested ? OnDpiChanged(hwnd, LOWORD(wp), *suggested) : 0;
        }
        case WM_THEMECHANGED:
        case WM_SYSCOLORCHANGE: return OnThemeChanged(hwnd);
        case WM_CLOSE: return OnClose(hwnd);
#ifdef _DEBUG
        case WndMsg::kFileOpsPopupSelfTestInvoke: return OnSelfTestInvoke(hwnd, reinterpret_cast<const PopupSelfTestInvoke*>(lp));
#endif
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

namespace
{

std::wstring ReadDialogItemText(HWND dlg, int controlId)
{
    const HWND control = GetDlgItem(dlg, controlId);
    if (! control)
    {
        return {};
    }

    const int length = GetWindowTextLengthW(control);
    if (length <= 0)
    {
        return {};
    }

    std::wstring text;
    text.resize(static_cast<size_t>(length) + 1u);
    GetWindowTextW(control, text.data(), length + 1);
    text.resize(static_cast<size_t>(length));
    return text;
}

void RestoreSpeedLimitDialogHint(HWND hwnd, SpeedLimitDialogState* state) noexcept
{
    if (! hwnd || ! state)
    {
        return;
    }

    const HWND message = GetDlgItem(hwnd, IDC_FILEOP_SPEED_LIMIT_CUSTOM_VALIDATION);
    if (! message)
    {
        return;
    }

    SetWindowTextW(message, state->hintText.c_str());
    state->showingValidationError = false;
    InvalidateRect(message, nullptr, TRUE);
}

void ShowSpeedLimitDialogValidationError(HWND hwnd, SpeedLimitDialogState* state, UINT messageId) noexcept
{
    if (! hwnd || ! state)
    {
        return;
    }

    const HWND message = GetDlgItem(hwnd, IDC_FILEOP_SPEED_LIMIT_CUSTOM_VALIDATION);
    if (! message)
    {
        return;
    }

    const std::wstring text = LoadStringResource(nullptr, messageId);
    SetWindowTextW(message, text.c_str());
    state->showingValidationError = true;
    InvalidateRect(message, nullptr, TRUE);
}

void FocusSpeedLimitDialogEdit(HWND hwnd) noexcept
{
    const HWND edit = hwnd ? GetDlgItem(hwnd, IDC_FILEOP_SPEED_LIMIT_CUSTOM_EDIT) : nullptr;
    if (! edit)
    {
        return;
    }

    SetFocus(edit);
    SendMessageW(edit, EM_SETSEL, 0, -1);
}

INT_PTR OnSpeedLimitDialogInit(HWND hwnd, SpeedLimitDialogState* state)
{
    SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(state));

    if (state)
    {
        ApplyTitleBarTheme(hwnd, state->theme, GetActiveWindow() == hwnd);
        state->backgroundBrush.reset(CreateSolidBrush(state->theme.windowBackground));
        state->hintText = ReadDialogItemText(hwnd, IDC_FILEOP_SPEED_LIMIT_CUSTOM_VALIDATION);
        RestoreSpeedLimitDialogHint(hwnd, state);

        std::wstring text = L"0";
        if (state->initialLimitBytesPerSecond != 0)
        {
            text = FormatBytesCompact(state->initialLimitBytesPerSecond);
        }
        SetDlgItemTextW(hwnd, IDC_FILEOP_SPEED_LIMIT_CUSTOM_EDIT, text.c_str());
    }

    return TRUE;
}

INT_PTR OnSpeedLimitDialogCtlColorDialog(SpeedLimitDialogState* state) noexcept
{
    if (! state || ! state->backgroundBrush)
    {
        return FALSE;
    }

    return reinterpret_cast<INT_PTR>(state->backgroundBrush.get());
}

INT_PTR OnSpeedLimitDialogCtlColorStatic(SpeedLimitDialogState* state, HDC hdc, HWND control) noexcept
{
    if (! state || ! state->backgroundBrush)
    {
        return FALSE;
    }

    COLORREF textColor = state->theme.menu.text;
    if (control)
    {
        const int controlId = GetDlgCtrlID(control);
        if (controlId == IDC_FILEOP_SPEED_LIMIT_CUSTOM_VALIDATION)
        {
            textColor = state->showingValidationError ? ColorRefFromColorF(state->theme.folderView.errorText) : state->theme.menu.disabledText;
        }
    }

    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, textColor);
    return reinterpret_cast<INT_PTR>(state->backgroundBrush.get());
}

INT_PTR OnSpeedLimitDialogCtlColorEdit(SpeedLimitDialogState* state, HDC hdc) noexcept
{
    if (! state || ! state->backgroundBrush)
    {
        return FALSE;
    }

    SetBkColor(hdc, state->theme.windowBackground);
    SetTextColor(hdc, state->theme.menu.text);
    return reinterpret_cast<INT_PTR>(state->backgroundBrush.get());
}

INT_PTR OnSpeedLimitDialogCommand(HWND hwnd, SpeedLimitDialogState* state, UINT commandId, UINT notifyCode) noexcept
{
    if (commandId == IDC_FILEOP_SPEED_LIMIT_CUSTOM_EDIT && notifyCode == EN_CHANGE)
    {
        if (state && state->showingValidationError)
        {
            RestoreSpeedLimitDialogHint(hwnd, state);
        }
        return TRUE;
    }

    if (commandId == IDOK)
    {
        if (! state)
        {
            EndDialog(hwnd, IDCANCEL);
            return TRUE;
        }

        RestoreSpeedLimitDialogHint(hwnd, state);

        const std::wstring text = ReadDialogItemText(hwnd, IDC_FILEOP_SPEED_LIMIT_CUSTOM_EDIT);

        unsigned __int64 parsed = 0;
        if (! TryParseThroughputText(std::wstring_view(text), parsed))
        {
            MessageBeep(MB_ICONERROR);
            ShowSpeedLimitDialogValidationError(hwnd, state, IDS_MSG_FILEOP_SPEED_LIMIT_INVALID);
            FocusSpeedLimitDialogEdit(hwnd);
            return TRUE;
        }

        state->resultLimitBytesPerSecond = parsed;
        EndDialog(hwnd, IDOK);
        return TRUE;
    }

    if (commandId == IDCANCEL)
    {
        EndDialog(hwnd, IDCANCEL);
        return TRUE;
    }

    return FALSE;
}

INT_PTR CALLBACK SpeedLimitDialogProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp)
{
    auto* state = reinterpret_cast<SpeedLimitDialogState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));

    switch (msg)
    {
        case WM_INITDIALOG: return OnSpeedLimitDialogInit(hwnd, reinterpret_cast<SpeedLimitDialogState*>(lp));
        case WM_CTLCOLORDLG: return OnSpeedLimitDialogCtlColorDialog(state);
        case WM_CTLCOLORSTATIC: return OnSpeedLimitDialogCtlColorStatic(state, reinterpret_cast<HDC>(wp), reinterpret_cast<HWND>(lp));
        case WM_CTLCOLOREDIT: return OnSpeedLimitDialogCtlColorEdit(state, reinterpret_cast<HDC>(wp));
        case WM_NCACTIVATE:
            if (state)
            {
                ApplyTitleBarTheme(hwnd, state->theme, wp != FALSE);
            }
            return FALSE;
        case WM_COMMAND: return OnSpeedLimitDialogCommand(hwnd, state, LOWORD(wp), HIWORD(wp));
    }

    return FALSE;
}

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

    if (state)
    {
        return state->WndProc(hwnd, msg, wp, lp);
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

HWND FileOperationsPopup::Create(FolderWindow::FileOperationState* fileOps, FolderWindow* folderWindow, HWND ownerWindow) noexcept
{
    if (! fileOps || ! folderWindow)
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
