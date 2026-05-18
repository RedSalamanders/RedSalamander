#include "DxUi.Internal.h"
#include "Helpers.h"

#include <algorithm>
#include <array>
#include <atomic>
#include <chrono>
#include <cmath>
#include <format>
#include <limits>
#include <mutex>
#include <new>
#include <optional>
#include <string>

#include <UIAutomation.h>
#include <oleauto.h>

#pragma comment(lib, "uiautomationcore.lib")

namespace RedSalamander::DxUi
{
namespace
{
constexpr PCWSTR kWindowHostPropName                 = L"RedSalamander.DxUi.WindowHost";
constexpr uint32_t kAccessibilityMaxDepth            = 16u;
constexpr UINT kWindowHostAccessibilityActionMessage = WM_APP + 0x6A;

class AccessibilityProvider;
class AccessibilityTextRangeProvider;

enum class AccessibilityUiActionKind : uint8_t
{
    SetFocus,
    Invoke,
    Toggle,
    SetStringValue,
    SetRangeValue,
    Select,
    AddToSelection,
    RemoveFromSelection,
    Expand,
    Collapse,
};

struct AccessibilityUiActionRequest
{
    AccessibilityProvider* provider = nullptr;
    AccessibilityTextRangeProvider* textRangeProvider = nullptr;
    AccessibilityUiActionKind kind  = AccessibilityUiActionKind::Invoke;
    std::wstring stringValue{};
    double numberValue = 0.0;
    HRESULT result     = static_cast<HRESULT>(UIA_E_NOTSUPPORTED);
};

struct ControlPath
{
    uint32_t depth = 0u;
    std::array<uint16_t, kAccessibilityMaxDepth> indices{};
};

[[nodiscard]] bool AreControlPathsEqual(const ControlPath& left, const ControlPath& right) noexcept
{
    if (left.depth != right.depth)
    {
        return false;
    }

    for (uint32_t index = 0u; index < left.depth; ++index)
    {
        if (left.indices[index] != right.indices[index])
        {
            return false;
        }
    }
    return true;
}

enum class AccessibilityFragmentKind : uint8_t
{
    Root,
    Control,
    TreeItem,
    GridHeader,
    GridRow,
    GridCell,
    TextFieldPasswordRevealButton,
};

struct WindowHostAccessibilityTarget final
{
    explicit WindowHostAccessibilityTarget(HWND hwnd, WindowHost* host) noexcept : hwnd(hwnd), host(host)
    {
    }

    WindowHostAccessibilityTarget(const WindowHostAccessibilityTarget&)            = delete;
    WindowHostAccessibilityTarget& operator=(const WindowHostAccessibilityTarget&) = delete;
    WindowHostAccessibilityTarget(WindowHostAccessibilityTarget&&)                 = delete;
    WindowHostAccessibilityTarget& operator=(WindowHostAccessibilityTarget&&)      = delete;

    [[nodiscard]] ULONG AddRef() noexcept
    {
        return _referenceCount.fetch_add(1u, std::memory_order_relaxed) + 1u;
    }

    [[nodiscard]] ULONG Release() noexcept
    {
        const ULONG remaining = _referenceCount.fetch_sub(1u, std::memory_order_acq_rel) - 1u;
        if (remaining == 0u)
        {
            delete this;
        }
        return remaining;
    }

    [[nodiscard]] WindowHost* ResolveHost() const noexcept
    {
        if (! hwnd || IsWindow(hwnd) == FALSE)
        {
            return nullptr;
        }

        return host.load(std::memory_order_acquire);
    }

    std::atomic<ULONG> _referenceCount{1u};
    HWND hwnd = nullptr;
    std::atomic<WindowHost*> host{nullptr};
};

[[nodiscard]] std::recursive_mutex& GetAccessibilityTargetMutex() noexcept
{
    static std::recursive_mutex mutex;
    return mutex;
}

[[nodiscard]] WindowHostAccessibilityTarget* AcquireWindowHostAccessibilityTarget(HWND hwnd) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return nullptr;
    }

    const std::scoped_lock lock(GetAccessibilityTargetMutex());
    auto* target = static_cast<WindowHostAccessibilityTarget*>(GetPropW(hwnd, kWindowHostPropName));
    if (! target)
    {
        return nullptr;
    }

    static_cast<void>(target->AddRef());
    return target;
}

[[nodiscard]] bool TryAppendPathIndex(const ControlPath& source, size_t childIndex, ControlPath& outPath) noexcept
{
    if (source.depth >= source.indices.size() || childIndex > (std::numeric_limits<uint16_t>::max)())
    {
        return false;
    }

    outPath                        = source;
    outPath.indices[outPath.depth] = static_cast<uint16_t>(childIndex);
    ++outPath.depth;
    return true;
}

template <typename TControl> [[nodiscard]] TControl* ResolveControlAtPath(TControl* root, const ControlPath& path) noexcept
{
    TControl* current = root;
    for (uint32_t depth = 0u; depth < path.depth; ++depth)
    {
        auto* panel = dynamic_cast<Panel*>(current);
        if (! panel)
        {
            return nullptr;
        }

        const auto children = panel->GetChildren();
        const size_t index  = path.indices[depth];
        if (index >= children.size() || ! children[index])
        {
            return nullptr;
        }

        current = children[index].get();
    }

    return current;
}

[[nodiscard]] bool IsControlPathVisible(const Control* root, const ControlPath& path) noexcept
{
    const Control* current = root;
    if (! current || ! current->IsVisible())
    {
        return false;
    }

    for (uint32_t depth = 0u; depth < path.depth; ++depth)
    {
        const auto* panel = dynamic_cast<const Panel*>(current);
        if (! panel)
        {
            return false;
        }

        const auto children = panel->GetChildren();
        const size_t index  = path.indices[depth];
        if (index >= children.size() || ! children[index])
        {
            return false;
        }

        current = children[index].get();
        if (! current->IsVisible())
        {
            return false;
        }
    }

    return true;
}

[[nodiscard]] bool IsSemanticAccessibilityControl(const Control* control) noexcept
{
    return dynamic_cast<const Label*>(control) != nullptr || dynamic_cast<const Button*>(control) != nullptr ||
           dynamic_cast<const TextField*>(control) != nullptr || dynamic_cast<const ComboBox*>(control) != nullptr ||
           dynamic_cast<const Tree*>(control) != nullptr || dynamic_cast<const Grid*>(control) != nullptr || dynamic_cast<const Slider*>(control) != nullptr ||
           dynamic_cast<const ColorSwatch*>(control) != nullptr;
}

[[nodiscard]] bool FindFirstSemanticControl(const Control* current, const ControlPath& basePath, ControlPath& outPath) noexcept
{
    if (! current || ! current->IsVisible())
    {
        return false;
    }

    if (IsSemanticAccessibilityControl(current))
    {
        outPath = basePath;
        return true;
    }

    auto* panel = dynamic_cast<const Panel*>(current);
    if (! panel)
    {
        return false;
    }

    const auto children = panel->GetChildren();
    for (size_t index = 0u; index < children.size(); ++index)
    {
        if (! children[index])
        {
            continue;
        }

        ControlPath childPath{};
        if (! TryAppendPathIndex(basePath, index, childPath))
        {
            continue;
        }

        if (FindFirstSemanticControl(children[index].get(), childPath, outPath))
        {
            return true;
        }
    }

    return false;
}

[[nodiscard]] bool FindLastSemanticControl(const Control* current, const ControlPath& basePath, ControlPath& outPath) noexcept
{
    if (! current || ! current->IsVisible())
    {
        return false;
    }

    if (const auto* panel = dynamic_cast<const Panel*>(current))
    {
        const auto children = panel->GetChildren();
        for (size_t index = children.size(); index-- > 0u;)
        {
            if (! children[index])
            {
                continue;
            }

            ControlPath childPath{};
            if (! TryAppendPathIndex(basePath, index, childPath))
            {
                continue;
            }

            if (FindLastSemanticControl(children[index].get(), childPath, outPath))
            {
                return true;
            }
        }
    }

    if (IsSemanticAccessibilityControl(current))
    {
        outPath = basePath;
        return true;
    }

    return false;
}

[[nodiscard]] bool FindNextSemanticControl(const Control* root, const ControlPath& currentPath, ControlPath& outPath) noexcept
{
    if (! root)
    {
        return false;
    }

    ControlPath cursor = currentPath;
    while (cursor.depth != 0u)
    {
        const uint32_t parentDepth = cursor.depth - 1u;
        ControlPath parentPath     = cursor;
        parentPath.depth           = parentDepth;
        const Control* parent      = ResolveControlAtPath(const_cast<Control*>(root), parentPath);
        const auto* panel          = dynamic_cast<const Panel*>(parent);
        if (! panel)
        {
            return false;
        }

        const auto children       = panel->GetChildren();
        const size_t currentIndex = cursor.indices[parentDepth];
        for (size_t siblingIndex = currentIndex + 1u; siblingIndex < children.size(); ++siblingIndex)
        {
            if (! children[siblingIndex])
            {
                continue;
            }

            ControlPath siblingPath{};
            if (! TryAppendPathIndex(parentPath, siblingIndex, siblingPath))
            {
                continue;
            }

            if (FindFirstSemanticControl(children[siblingIndex].get(), siblingPath, outPath))
            {
                return true;
            }
        }

        cursor.depth = parentDepth;
    }

    return false;
}

[[nodiscard]] bool FindPreviousSemanticControl(const Control* root, const ControlPath& currentPath, ControlPath& outPath) noexcept
{
    if (! root)
    {
        return false;
    }

    ControlPath cursor = currentPath;
    while (cursor.depth != 0u)
    {
        const uint32_t parentDepth = cursor.depth - 1u;
        ControlPath parentPath     = cursor;
        parentPath.depth           = parentDepth;
        const Control* parent      = ResolveControlAtPath(const_cast<Control*>(root), parentPath);
        const auto* panel          = dynamic_cast<const Panel*>(parent);
        if (! panel)
        {
            return false;
        }

        const auto children       = panel->GetChildren();
        const size_t currentIndex = cursor.indices[parentDepth];
        for (size_t siblingIndex = currentIndex; siblingIndex-- > 0u;)
        {
            if (! children[siblingIndex])
            {
                continue;
            }

            ControlPath siblingPath{};
            if (! TryAppendPathIndex(parentPath, siblingIndex, siblingPath))
            {
                continue;
            }

            if (FindLastSemanticControl(children[siblingIndex].get(), siblingPath, outPath))
            {
                return true;
            }
        }

        cursor.depth = parentDepth;
    }

    return false;
}

[[nodiscard]] bool ShouldExposeSingleSemanticRootControl(const Control* control) noexcept
{
    return dynamic_cast<const Checkbox*>(control) != nullptr || dynamic_cast<const Toggle*>(control) != nullptr ||
           dynamic_cast<const Button*>(control) != nullptr || dynamic_cast<const TextField*>(control) != nullptr ||
           dynamic_cast<const ComboBox*>(control) != nullptr || dynamic_cast<const Tree*>(control) != nullptr ||
           dynamic_cast<const Slider*>(control) != nullptr || dynamic_cast<const Grid*>(control) != nullptr ||
           dynamic_cast<const ColorSwatch*>(control) != nullptr;
}

[[nodiscard]] bool TryResolveSingleSemanticRootControlPath(const Control* root, ControlPath& outPath) noexcept
{
    ControlPath firstPath{};
    if (! FindFirstSemanticControl(root, ControlPath{}, firstPath))
    {
        return false;
    }

    const Control* control = ResolveControlAtPath(const_cast<Control*>(root), firstPath);
    if (! ShouldExposeSingleSemanticRootControl(control))
    {
        return false;
    }

    ControlPath nextPath{};
    if (FindNextSemanticControl(root, firstPath, nextPath))
    {
        return false;
    }

    outPath = firstPath;
    return true;
}

[[nodiscard]] bool FindSemanticControlAtPoint(const Control* current, const ControlPath& basePath, D2D1_POINT_2F pointDip, ControlPath& outPath) noexcept
{
    if (! current || ! current->IsVisible() || ! PointInRect(current->GetHitBounds(), pointDip))
    {
        return false;
    }

    if (const auto* panel = dynamic_cast<const Panel*>(current))
    {
        const auto children = panel->GetChildren();
        for (size_t index = children.size(); index-- > 0u;)
        {
            if (! children[index])
            {
                continue;
            }

            ControlPath childPath{};
            if (! TryAppendPathIndex(basePath, index, childPath))
            {
                continue;
            }

            if (FindSemanticControlAtPoint(children[index].get(), childPath, pointDip, outPath))
            {
                return true;
            }
        }
    }

    if (IsSemanticAccessibilityControl(current))
    {
        outPath = basePath;
        return true;
    }

    return false;
}

[[nodiscard]] std::wstring_view FindAssociatedLabelText(const Control* root, const Control* target, uint32_t depth = 0u) noexcept
{
    if (! root || ! target || depth >= kAccessibilityMaxDepth)
    {
        return {};
    }

    if (const auto* label = dynamic_cast<const Label*>(root))
    {
        if (label->GetMnemonicTarget() == target)
        {
            return label->GetText();
        }
    }

    if (const auto* panel = dynamic_cast<const Panel*>(root))
    {
        const auto children = panel->GetChildren();
        for (const auto& child : children)
        {
            if (! child)
            {
                continue;
            }

            if (const std::wstring_view text = FindAssociatedLabelText(child.get(), target, depth + 1u); ! text.empty())
            {
                return text;
            }
        }
    }

    return {};
}

[[nodiscard]] std::wstring_view GetControlAccessibleName(const Control* root, const Control* control) noexcept
{
    if (! control)
    {
        return {};
    }

    if (const std::wstring_view explicitName = control->GetAccessibleName(); ! explicitName.empty())
    {
        return explicitName;
    }

    if (const auto* toggle = dynamic_cast<const Toggle*>(control))
    {
        if (const std::wstring_view labelText = FindAssociatedLabelText(root, toggle); ! labelText.empty())
        {
            return labelText;
        }
        return toggle->GetDisplayedText();
    }
    if (const auto* button = dynamic_cast<const Button*>(control))
    {
        return button->GetText();
    }
    if (const auto* label = dynamic_cast<const Label*>(control))
    {
        return label->GetText();
    }
    if (const auto* textField = dynamic_cast<const TextField*>(control))
    {
        if (const std::wstring_view labelText = FindAssociatedLabelText(root, textField); ! labelText.empty())
        {
            return labelText;
        }
        if (textField->IsMasked())
        {
            return {};
        }
        return textField->GetText();
    }
    if (const auto* comboBox = dynamic_cast<const ComboBox*>(control))
    {
        if (const std::wstring_view labelText = FindAssociatedLabelText(root, comboBox); ! labelText.empty())
        {
            return labelText;
        }
        return comboBox->GetDisplayedText();
    }
    if (const auto* slider = dynamic_cast<const Slider*>(control))
    {
        if (const std::wstring_view labelText = FindAssociatedLabelText(root, slider); ! labelText.empty())
        {
            return labelText;
        }
    }

    return FindAssociatedLabelText(root, control);
}

[[nodiscard]] std::wstring_view GetControlAccessibleValue(const Control* control) noexcept
{
    if (const auto* textField = dynamic_cast<const TextField*>(control))
    {
        if (textField->IsMasked())
        {
            return {};
        }
        return textField->GetText();
    }
    if (const auto* comboBox = dynamic_cast<const ComboBox*>(control))
    {
        return comboBox->GetDisplayedText();
    }
    return {};
}

[[nodiscard]] bool IsTextFieldPasswordRevealButtonVisible(const Control* control) noexcept
{
    const auto* textField = dynamic_cast<const TextField*>(control);
    return textField && textField->IsPasswordRevealButtonVisibleForAccessibility();
}

[[nodiscard]] bool SupportsTextPattern(const Control* control) noexcept
{
    if (dynamic_cast<const TextField*>(control) != nullptr)
    {
        return true;
    }
    if (const auto* comboBox = dynamic_cast<const ComboBox*>(control))
    {
        return comboBox->IsEditable();
    }
    return false;
}

[[nodiscard]] std::wstring GetControlAccessibleTextRangeText(const Control* control)
{
    if (const auto* textField = dynamic_cast<const TextField*>(control))
    {
        if (textField->IsMasked())
        {
            return {};
        }
        return std::wstring(textField->GetText());
    }
    if (const auto* comboBox = dynamic_cast<const ComboBox*>(control))
    {
        if (! comboBox->IsEditable())
        {
            return {};
        }
        return std::wstring(comboBox->GetText());
    }
    return {};
}

[[nodiscard]] D2D1_RECT_F ResolveTextPatternViewportRect(const Control* control) noexcept
{
    if (! control)
    {
        return D2D1::RectF();
    }
    if (const std::optional<D2D1_RECT_F> viewport = control->TryGetTextInputViewportRect(); viewport.has_value())
    {
        return viewport.value();
    }
    return control->GetHitBounds();
}

struct TextRangeSpan
{
    size_t start = 0u;
    size_t end   = 0u;
};

[[nodiscard]] bool IsSimpleTextRangeCaretGeometrySupported(const Control& control, std::wstring_view text) noexcept
{
    const DWRITE_READING_DIRECTION readingDirection = ResolveReadingDirection(control.GetFlowDirection());
    return ! ShouldEmitSingleLineBiDiTextMetric(text, readingDirection);
}

[[nodiscard]] bool TextRangeContainsLineBreak(std::wstring_view text, const TextRangeSpan& range) noexcept
{
    const size_t start = (std::min)(range.start, text.size());
    const size_t end   = (std::min)(range.end, text.size());
    return text.substr(start, end - start).find_first_of(L"\r\n") != std::wstring_view::npos;
}

[[nodiscard]] size_t FindNextLineBreak(std::wstring_view text, size_t start, size_t end) noexcept
{
    const size_t clampedStart = (std::min)(start, text.size());
    const size_t clampedEnd   = (std::min)(end, text.size());
    const size_t breakIndex   = text.substr(clampedStart, clampedEnd - clampedStart).find_first_of(L"\r\n");
    return breakIndex == std::wstring_view::npos ? clampedEnd : clampedStart + breakIndex;
}

[[nodiscard]] size_t AdvancePastLineBreak(std::wstring_view text, size_t index, size_t end) noexcept
{
    const size_t clampedIndex = (std::min)(index, text.size());
    const size_t clampedEnd   = (std::min)(end, text.size());
    if (clampedIndex >= clampedEnd)
    {
        return clampedIndex;
    }
    if (text[clampedIndex] == L'\r' && clampedIndex + 1u < clampedEnd && text[clampedIndex + 1u] == L'\n')
    {
        return clampedIndex + 2u;
    }
    return clampedIndex + 1u;
}

[[nodiscard]] bool CaretRectsShareVisualLine(const D2D1_RECT_F& first, const D2D1_RECT_F& second) noexcept
{
    constexpr float kLineToleranceDip = 1.0f;
    return std::fabs(first.top - second.top) <= kLineToleranceDip && std::fabs(first.bottom - second.bottom) <= kLineToleranceDip;
}

[[nodiscard]] D2D1_RECT_F ClipRectToBounds(D2D1_RECT_F rect, const D2D1_RECT_F& bounds) noexcept
{
    rect.left   = std::clamp(rect.left, bounds.left, bounds.right);
    rect.right  = std::clamp(rect.right, rect.left, bounds.right);
    rect.top    = std::clamp(rect.top, bounds.top, bounds.bottom);
    rect.bottom = std::clamp(rect.bottom, rect.top, bounds.bottom);
    if (rect.right <= rect.left)
    {
        rect.right = (std::min)(bounds.right, rect.left + 1.0f);
    }
    if (rect.bottom <= rect.top)
    {
        rect.bottom = (std::min)(bounds.bottom, rect.top + 1.0f);
    }
    return rect;
}

[[nodiscard]] std::optional<D2D1_RECT_F> TryResolveSimpleTextRangeCaretRect(
    const WindowHost& host, const Control& control, std::wstring_view text, const TextRangeSpan& range) noexcept
{
    if (range.start == range.end || ! IsSimpleTextRangeCaretGeometrySupported(control, text) || TextRangeContainsLineBreak(text, range))
    {
        return std::nullopt;
    }

    const std::optional<D2D1_RECT_F> startRect = control.TryGetTextInputCaretRect(host, range.start);
    const std::optional<D2D1_RECT_F> endRect   = control.TryGetTextInputCaretRect(host, range.end);
    if (! startRect.has_value() || ! endRect.has_value())
    {
        return std::nullopt;
    }
    if (! CaretRectsShareVisualLine(startRect.value(), endRect.value()))
    {
        return std::nullopt;
    }

    const D2D1_RECT_F viewport = ResolveTextPatternViewportRect(&control);
    const D2D1_RECT_F rect     = D2D1::RectF((std::min)(startRect->left, endRect->left),
                                             (std::min)(startRect->top, endRect->top),
                                             (std::max)(startRect->right, endRect->right),
                                             (std::max)(startRect->bottom, endRect->bottom));
    return ClipRectToBounds(rect, viewport);
}

[[nodiscard]] std::optional<std::vector<D2D1_RECT_F>> TryResolveTextRangeCaretRects(
    const WindowHost& host, const Control& control, std::wstring_view text, const TextRangeSpan& range)
{
    if (range.start == range.end)
    {
        return std::nullopt;
    }

    if (! IsSimpleTextRangeCaretGeometrySupported(control, text))
    {
        if (std::optional<std::vector<D2D1_RECT_F>> rects = control.TryGetTextInputRangeRects(host, range.start, range.end);
            rects.has_value() && ! rects->empty())
        {
            return rects;
        }
        return std::nullopt;
    }

    if (! TextRangeContainsLineBreak(text, range))
    {
        if (std::optional<D2D1_RECT_F> rect = TryResolveSimpleTextRangeCaretRect(host, control, text, range); rect.has_value())
        {
            return std::vector<D2D1_RECT_F>{rect.value()};
        }
        if (std::optional<std::vector<D2D1_RECT_F>> rects = control.TryGetTextInputRangeRects(host, range.start, range.end);
            rects.has_value() && ! rects->empty())
        {
            return rects;
        }
        return std::nullopt;
    }

    const D2D1_RECT_F viewport = ResolveTextPatternViewportRect(&control);
    std::vector<D2D1_RECT_F> rects;
    size_t segmentStart = (std::min)(range.start, text.size());
    const size_t rangeEnd = (std::min)(range.end, text.size());
    while (segmentStart < rangeEnd)
    {
        const size_t segmentEnd = FindNextLineBreak(text, segmentStart, rangeEnd);
        if (segmentStart < segmentEnd)
        {
            const std::optional<D2D1_RECT_F> startRect = control.TryGetTextInputCaretRect(host, segmentStart);
            const std::optional<D2D1_RECT_F> endRect   = control.TryGetTextInputCaretRect(host, segmentEnd);
            if (! startRect.has_value() || ! endRect.has_value())
            {
                return std::nullopt;
            }

            if (CaretRectsShareVisualLine(startRect.value(), endRect.value()))
            {
                const D2D1_RECT_F rect = D2D1::RectF((std::min)(startRect->left, endRect->left),
                                                    (std::min)(startRect->top, endRect->top),
                                                    (std::max)(startRect->right, endRect->right),
                                                    (std::max)(startRect->bottom, endRect->bottom));
                rects.push_back(ClipRectToBounds(rect, viewport));
            }
            else if (std::optional<std::vector<D2D1_RECT_F>> segmentRects = control.TryGetTextInputRangeRects(host, segmentStart, segmentEnd);
                     segmentRects.has_value() && ! segmentRects->empty())
            {
                rects.insert(rects.end(), segmentRects->begin(), segmentRects->end());
            }
            else
            {
                return std::nullopt;
            }
        }

        if (segmentEnd >= rangeEnd)
        {
            break;
        }
        segmentStart = AdvancePastLineBreak(text, segmentEnd, rangeEnd);
    }

    if (rects.empty())
    {
        return std::nullopt;
    }
    return rects;
}

[[nodiscard]] size_t GetTextRangeEndpointPosition(const TextRangeSpan& range, TextPatternRangeEndpoint endpoint) noexcept
{
    return endpoint == TextPatternRangeEndpoint_Start ? range.start : range.end;
}

struct TextRangeUnitMoveResult
{
    size_t position = 0u;
    int moved       = 0;
};

[[nodiscard]] TextRangeUnitMoveResult MoveTextRangePositionByUnit(
    std::wstring_view text, size_t position, TextUnit unit, int count) noexcept
{
    const size_t clampedPosition = std::min(position, text.size());
    if (count == 0)
    {
        return TextRangeUnitMoveResult{clampedPosition, 0};
    }

    if (unit == TextUnit_Character)
    {
        size_t cursor = SnapCaretIndexToTextElementBoundary(text, clampedPosition);
        int moved     = 0;
        if (count > 0)
        {
            while (moved < count)
            {
                const size_t next = StepToNextTextElement(text, cursor);
                if (next == cursor)
                {
                    break;
                }

                cursor = next;
                ++moved;
            }
        }
        else
        {
            while (moved > count)
            {
                const size_t previous = StepToPreviousTextElement(text, cursor);
                if (previous == cursor)
                {
                    break;
                }

                cursor = previous;
                --moved;
            }
        }

        return TextRangeUnitMoveResult{cursor, moved};
    }

    if (unit != TextUnit_Word)
    {
        return TextRangeUnitMoveResult{clampedPosition, 0};
    }

    size_t cursor = clampedPosition;
    int moved     = 0;
    if (count > 0)
    {
        while (moved < count)
        {
            const size_t next = FindNextWordBoundary(text, cursor);
            if (next == cursor)
            {
                break;
            }

            cursor = next;
            ++moved;
        }
    }
    else
    {
        while (moved > count)
        {
            const size_t previous = FindPreviousWordBoundary(text, cursor);
            if (previous == cursor)
            {
                break;
            }

            cursor = previous;
            --moved;
        }
    }

    return TextRangeUnitMoveResult{cursor, moved};
}

[[nodiscard]] TextRangeSpan ClampTextRangeSpan(size_t start, size_t end, size_t textLength) noexcept
{
    start = (std::min)(start, textLength);
    end   = (std::min)(end, textLength);
    if (start > end)
    {
        std::swap(start, end);
    }
    return TextRangeSpan{start, end};
}

[[nodiscard]] bool IsAccessibilityTextRangeWhitespace(wchar_t value) noexcept
{
    WORD charType = 0u;
    return GetStringTypeW(CT_CTYPE1, &value, 1, &charType) && ((charType & C1_SPACE) != 0u);
}

[[nodiscard]] size_t TrimTrailingTextRangeWhitespace(std::wstring_view text, size_t start, size_t end) noexcept
{
    start = (std::min)(start, text.size());
    end   = (std::min)(end, text.size());
    while (end > start)
    {
        const size_t previous = StepToPreviousCodePoint(text, end);
        if (! IsAccessibilityTextRangeWhitespace(text[previous]))
        {
            break;
        }

        end = previous;
    }

    return end;
}

[[nodiscard]] TextRangeSpan GetTextRangeWordSpanAtPosition(std::wstring_view text, size_t position) noexcept
{
    const size_t start = (std::min)(position, text.size());
    if (start >= text.size())
    {
        return TextRangeSpan{text.size(), text.size()};
    }

    const TextRangeUnitMoveResult endBoundary = MoveTextRangePositionByUnit(text, start, TextUnit_Word, 1);
    const size_t end = TrimTrailingTextRangeWhitespace(text, start, endBoundary.position);
    return TextRangeSpan{start, end};
}

struct TextRangeSpanMoveResult
{
    TextRangeSpan range{};
    int moved = 0;
};

[[nodiscard]] TextRangeSpanMoveResult MoveTextRangeSpanByWord(std::wstring_view text, const TextRangeSpan& range, int count) noexcept
{
    const TextRangeSpan clampedRange = ClampTextRangeSpan(range.start, range.end, text.size());
    if (count == 0)
    {
        return TextRangeSpanMoveResult{clampedRange, 0};
    }

    size_t cursor = clampedRange.start;
    int moved     = 0;
    if (count > 0)
    {
        while (moved < count)
        {
            const TextRangeUnitMoveResult moveResult = MoveTextRangePositionByUnit(text, cursor, TextUnit_Word, 1);
            if (moveResult.position == cursor || moveResult.moved == 0)
            {
                break;
            }

            cursor = moveResult.position;
            ++moved;
        }
    }
    else
    {
        while (moved > count)
        {
            const TextRangeUnitMoveResult moveResult = MoveTextRangePositionByUnit(text, cursor, TextUnit_Word, -1);
            if (moveResult.position == cursor || moveResult.moved == 0)
            {
                break;
            }

            cursor = moveResult.position;
            --moved;
        }
    }

    if (moved == 0)
    {
        return TextRangeSpanMoveResult{clampedRange, 0};
    }

    return TextRangeSpanMoveResult{GetTextRangeWordSpanAtPosition(text, cursor), moved};
}

[[nodiscard]] bool IsAccessibilityTextRangeLineBreak(wchar_t value) noexcept
{
    return value == L'\r' || value == L'\n';
}

[[nodiscard]] size_t GetTextRangeLineStart(std::wstring_view text, size_t position) noexcept
{
    size_t cursor = (std::min)(position, text.size());
    while (cursor > 0u)
    {
        const size_t previous = StepToPreviousCodePoint(text, cursor);
        if (IsAccessibilityTextRangeLineBreak(text[previous]))
        {
            return cursor;
        }

        cursor = previous;
    }

    return 0u;
}

[[nodiscard]] size_t GetTextRangeLineEnd(std::wstring_view text, size_t lineStart) noexcept
{
    size_t cursor = (std::min)(lineStart, text.size());
    while (cursor < text.size() && ! IsAccessibilityTextRangeLineBreak(text[cursor]))
    {
        cursor = StepToNextCodePoint(text, cursor);
    }

    return cursor;
}

[[nodiscard]] size_t AdvancePastTextRangeLineBreak(std::wstring_view text, size_t position) noexcept
{
    position = (std::min)(position, text.size());
    if (position >= text.size())
    {
        return text.size();
    }

    if (text[position] == L'\r')
    {
        const size_t afterCarriageReturn = StepToNextCodePoint(text, position);
        if (afterCarriageReturn < text.size() && text[afterCarriageReturn] == L'\n')
        {
            return StepToNextCodePoint(text, afterCarriageReturn);
        }

        return afterCarriageReturn;
    }

    if (text[position] == L'\n')
    {
        return StepToNextCodePoint(text, position);
    }

    return position;
}

[[nodiscard]] TextRangeSpan GetTextRangeLineSpanAtPosition(std::wstring_view text, size_t position) noexcept
{
    const size_t lineStart = GetTextRangeLineStart(text, position);
    return TextRangeSpan{lineStart, GetTextRangeLineEnd(text, lineStart)};
}

[[nodiscard]] size_t GetPreviousTextRangeLineStart(std::wstring_view text, size_t lineStart) noexcept
{
    lineStart = (std::min)(lineStart, text.size());
    if (lineStart == 0u)
    {
        return 0u;
    }

    size_t previousLineEnd = StepToPreviousCodePoint(text, lineStart);
    if (text[previousLineEnd] == L'\n' && previousLineEnd > 0u)
    {
        const size_t beforeLineFeed = StepToPreviousCodePoint(text, previousLineEnd);
        if (text[beforeLineFeed] == L'\r')
        {
            previousLineEnd = beforeLineFeed;
        }
    }

    return GetTextRangeLineStart(text, previousLineEnd);
}

[[nodiscard]] TextRangeSpanMoveResult MoveTextRangeSpanByLine(std::wstring_view text, const TextRangeSpan& range, int count) noexcept
{
    const TextRangeSpan clampedRange = ClampTextRangeSpan(range.start, range.end, text.size());
    if (count == 0)
    {
        return TextRangeSpanMoveResult{clampedRange, 0};
    }

    size_t cursor = GetTextRangeLineStart(text, clampedRange.start);
    int moved     = 0;
    if (count > 0)
    {
        while (moved < count)
        {
            const size_t lineEnd = GetTextRangeLineEnd(text, cursor);
            if (lineEnd >= text.size())
            {
                break;
            }

            const size_t nextLineStart = AdvancePastTextRangeLineBreak(text, lineEnd);
            if (nextLineStart == cursor)
            {
                break;
            }

            cursor = nextLineStart;
            ++moved;
        }
    }
    else
    {
        while (moved > count)
        {
            const size_t previousLineStart = GetPreviousTextRangeLineStart(text, cursor);
            if (previousLineStart == cursor)
            {
                break;
            }

            cursor = previousLineStart;
            --moved;
        }
    }

    if (moved == 0)
    {
        return TextRangeSpanMoveResult{clampedRange, 0};
    }

    return TextRangeSpanMoveResult{GetTextRangeLineSpanAtPosition(text, cursor), moved};
}

[[nodiscard]] std::optional<std::vector<TextRangeSpan>> TryResolveTextRangeVisualLineSpans(
    const WindowHost& host, const Control& control, std::wstring_view text) noexcept
{
    if (text.empty() || ! IsSimpleTextRangeCaretGeometrySupported(control, text))
    {
        return std::nullopt;
    }

    std::vector<TextRangeSpan> spans;
    size_t logicalLineCount = 0u;
    size_t logicalLineStart = 0u;
    while (logicalLineStart <= text.size())
    {
        ++logicalLineCount;
        const size_t logicalLineEnd = GetTextRangeLineEnd(text, logicalLineStart);
        if (logicalLineStart == logicalLineEnd)
        {
            spans.push_back(TextRangeSpan{logicalLineStart, logicalLineEnd});
        }
        else
        {
            size_t visualLineStart = logicalLineStart;
            std::optional<D2D1_RECT_F> visualLineRect = control.TryGetTextInputCaretRect(host, visualLineStart);
            if (! visualLineRect.has_value())
            {
                return std::nullopt;
            }

            size_t cursor = logicalLineStart;
            while (cursor < logicalLineEnd)
            {
                const size_t next = StepToNextCodePoint(text, cursor);
                const std::optional<D2D1_RECT_F> nextRect = control.TryGetTextInputCaretRect(host, next);
                if (! nextRect.has_value())
                {
                    return std::nullopt;
                }

                if (! CaretRectsShareVisualLine(visualLineRect.value(), nextRect.value()))
                {
                    if (visualLineStart < next)
                    {
                        spans.push_back(TextRangeSpan{visualLineStart, next});
                    }
                    visualLineStart = next;
                    visualLineRect  = nextRect;
                }

                cursor = next;
            }

            if (visualLineStart < logicalLineEnd)
            {
                spans.push_back(TextRangeSpan{visualLineStart, logicalLineEnd});
            }
        }

        if (logicalLineEnd >= text.size())
        {
            break;
        }

        const size_t nextLogicalLineStart = AdvancePastTextRangeLineBreak(text, logicalLineEnd);
        if (nextLogicalLineStart <= logicalLineStart)
        {
            break;
        }
        logicalLineStart = nextLogicalLineStart;
    }

    if (spans.size() <= logicalLineCount)
    {
        return std::nullopt;
    }
    return spans;
}

[[nodiscard]] size_t FindTextRangeVisualLineIndex(std::span<const TextRangeSpan> spans, size_t position) noexcept
{
    if (spans.empty())
    {
        return 0u;
    }

    for (size_t index = 0u; index < spans.size(); ++index)
    {
        const TextRangeSpan& span = spans[index];
        if (position == span.start || (position > span.start && position < span.end))
        {
            return index;
        }
        if (position < span.start)
        {
            return index == 0u ? 0u : index - 1u;
        }
    }

    return spans.size() - 1u;
}

[[nodiscard]] std::optional<TextRangeSpanMoveResult> TryMoveTextRangeSpanByVisualLine(
    const WindowHost& host, const Control& control, std::wstring_view text, const TextRangeSpan& range, int count) noexcept
{
    const TextRangeSpan clampedRange = ClampTextRangeSpan(range.start, range.end, text.size());
    if (count == 0)
    {
        return TextRangeSpanMoveResult{clampedRange, 0};
    }

    const std::optional<std::vector<TextRangeSpan>> spans = TryResolveTextRangeVisualLineSpans(host, control, text);
    if (! spans.has_value() || spans->empty())
    {
        return std::nullopt;
    }

    const size_t currentIndex = FindTextRangeVisualLineIndex(*spans, clampedRange.start);
    const int requestedIndex  = static_cast<int>(currentIndex) + count;
    const int targetIndex     = std::clamp(requestedIndex, 0, static_cast<int>(spans->size() - 1u));
    const int moved           = targetIndex - static_cast<int>(currentIndex);
    if (moved == 0)
    {
        return TextRangeSpanMoveResult{clampedRange, 0};
    }

    return TextRangeSpanMoveResult{spans->at(static_cast<size_t>(targetIndex)), moved};
}

[[nodiscard]] TextRangeUnitMoveResult MoveTextRangePositionByLine(std::wstring_view text, size_t position, int count) noexcept
{
    const size_t clampedPosition = (std::min)(position, text.size());
    if (count == 0)
    {
        return TextRangeUnitMoveResult{clampedPosition, 0};
    }

    const TextRangeSpan lineRange = GetTextRangeLineSpanAtPosition(text, clampedPosition);
    const TextRangeSpanMoveResult moveResult = MoveTextRangeSpanByLine(text, lineRange, count);
    if (moveResult.moved == 0)
    {
        return TextRangeUnitMoveResult{clampedPosition, 0};
    }

    return TextRangeUnitMoveResult{moveResult.range.start, moveResult.moved};
}

[[nodiscard]] std::optional<TextRangeUnitMoveResult> TryMoveTextRangePositionByVisualLine(
    const WindowHost& host, const Control& control, std::wstring_view text, size_t position, int count) noexcept
{
    const size_t clampedPosition = (std::min)(position, text.size());
    if (count == 0)
    {
        return TextRangeUnitMoveResult{clampedPosition, 0};
    }

    const std::optional<TextRangeSpanMoveResult> moveResult =
        TryMoveTextRangeSpanByVisualLine(host, control, text, TextRangeSpan{clampedPosition, clampedPosition}, count);
    if (! moveResult.has_value())
    {
        return std::nullopt;
    }

    return TextRangeUnitMoveResult{moveResult->range.start, moveResult->moved};
}

[[nodiscard]] TextRangeSpan GetTextRangeSpanFromState(size_t caretIndex, std::optional<size_t> selectionAnchorIndex, size_t textLength) noexcept
{
    const size_t clampedCaret = (std::min)(caretIndex, textLength);
    if (! selectionAnchorIndex)
    {
        return TextRangeSpan{clampedCaret, clampedCaret};
    }

    return ClampTextRangeSpan(selectionAnchorIndex.value(), clampedCaret, textLength);
}

[[nodiscard]] TextRangeSpan GetControlAccessibleSelectionRange(const WindowHost* host, const Control* control, size_t textLength) noexcept
{
    if (host && control)
    {
        NativeTextInputState nativeState{};
        if (host->TryReadNativeTextInputState(control, nativeState))
        {
            return GetTextRangeSpanFromState(nativeState.caretIndex, nativeState.selectionAnchorIndex, textLength);
        }

        TextInputState textInputState{};
        if (host->TryReadTextInputState(control, textInputState))
        {
            return GetTextRangeSpanFromState(textInputState.caretIndex, textInputState.selectionAnchorIndex, textLength);
        }
    }

    if (const auto* textField = dynamic_cast<const TextField*>(control))
    {
        if (const std::optional<std::pair<size_t, size_t>> selectionRange = textField->GetSelectionRange())
        {
            return ClampTextRangeSpan(selectionRange->first, selectionRange->second, textLength);
        }
    }
    if (const auto* comboBox = dynamic_cast<const ComboBox*>(control); comboBox && comboBox->IsEditable())
    {
        if (const std::optional<std::pair<size_t, size_t>> selectionRange = comboBox->GetEditableSelectionRange())
        {
            return ClampTextRangeSpan(selectionRange->first, selectionRange->second, textLength);
        }
    }

    return TextRangeSpan{textLength, textLength};
}

[[nodiscard]] std::wstring BuildGridCellAccessibleText(const GridCellData& cellData)
{
    std::wstring text;
    if (cellData.kind == GridCellKind::Checkbox)
    {
        text.assign(cellData.checked ? L"[x]" : L"[ ]");
        if (! cellData.text.empty())
        {
            text.push_back(L' ');
        }
    }

    if (! cellData.text.empty())
    {
        text.append(cellData.text);
    }
    else if (cellData.kind == GridCellKind::ColorSwatch && cellData.hasSwatchValue)
    {
        text.append(std::format(L"#{:08X}", cellData.swatchArgb));
    }
    else if (cellData.kind == GridCellKind::IconText && ! cellData.iconText.empty())
    {
        text.append(cellData.iconText);
    }

    if (! cellData.badgeText.empty())
    {
        if (! text.empty())
        {
            text.push_back(L' ');
        }
        text.push_back(L'[');
        text.append(cellData.badgeText);
        text.push_back(L']');
    }

    return text;
}

[[nodiscard]] std::wstring BuildGridRowAccessibleName(const IDxGridModel& model, size_t rowIndex)
{
    std::wstring text;
    for (size_t columnIndex = 0u; columnIndex < model.GetColumnCount(); ++columnIndex)
    {
        GridCellData cellData{};
        model.GetCellData(rowIndex, columnIndex, cellData);
        std::wstring cellText = BuildGridCellAccessibleText(cellData);
        if (cellText.empty())
        {
            continue;
        }

        if (! text.empty())
        {
            text.append(L" | ");
        }
        text.append(cellText);
    }

    return text;
}

[[nodiscard]] std::wstring_view GetGridHeaderAccessibleName(const GridColumnDesc& columnDesc) noexcept
{
    if (! columnDesc.title.empty())
    {
        return columnDesc.title;
    }

    return columnDesc.id;
}

[[nodiscard]] CONTROLTYPEID GetGridCellControlTypeId(const GridCellData& cellData) noexcept
{
    switch (cellData.kind)
    {
        case GridCellKind::Checkbox: return UIA_CheckBoxControlTypeId;
        case GridCellKind::ColorSwatch: return UIA_ImageControlTypeId;
        case GridCellKind::Spinner:
        case GridCellKind::Marquee: return UIA_ProgressBarControlTypeId;
        case GridCellKind::IconText: return (! cellData.iconText.empty() && cellData.text.empty()) ? UIA_ImageControlTypeId : UIA_TextControlTypeId;
        case GridCellKind::Text:
        default: return UIA_TextControlTypeId;
    }
}

[[nodiscard]] bool GridCellSupportsTogglePattern(const GridCellData& cellData) noexcept
{
    return cellData.kind == GridCellKind::Checkbox && cellData.enabled;
}

[[nodiscard]] bool GridCellSupportsValuePattern(const GridCellData& cellData) noexcept
{
    switch (cellData.kind)
    {
        case GridCellKind::Text:
        case GridCellKind::ColorSwatch: return true;
        case GridCellKind::IconText: return ! cellData.text.empty();
        case GridCellKind::Checkbox:
        case GridCellKind::Spinner:
        case GridCellKind::Marquee:
        default: return false;
    }
}

[[nodiscard]] bool GridCellSupportsRangeValuePattern(const GridCellData& cellData) noexcept
{
    return cellData.kind == GridCellKind::Marquee && std::isfinite(cellData.progress) && cellData.progress > 0.0f;
}

[[nodiscard]] double GetGridCellRangeValue(const GridCellData& cellData) noexcept
{
    return std::clamp(static_cast<double>(cellData.progress), 0.0, 1.0);
}

[[nodiscard]] CONTROLTYPEID GetControlTypeId(const Control* control) noexcept
{
    if (dynamic_cast<const Checkbox*>(control) != nullptr)
    {
        return UIA_CheckBoxControlTypeId;
    }
    if (dynamic_cast<const Toggle*>(control) != nullptr || dynamic_cast<const Button*>(control) != nullptr)
    {
        return UIA_ButtonControlTypeId;
    }
    if (dynamic_cast<const TextField*>(control) != nullptr)
    {
        return UIA_EditControlTypeId;
    }
    if (dynamic_cast<const ComboBox*>(control) != nullptr)
    {
        return UIA_ComboBoxControlTypeId;
    }
    if (dynamic_cast<const Slider*>(control) != nullptr)
    {
        return UIA_SliderControlTypeId;
    }
    if (dynamic_cast<const Tree*>(control) != nullptr)
    {
        return UIA_TreeControlTypeId;
    }
    if (dynamic_cast<const Grid*>(control) != nullptr)
    {
        return UIA_DataGridControlTypeId;
    }
    if (dynamic_cast<const ColorSwatch*>(control) != nullptr)
    {
        return UIA_ImageControlTypeId;
    }
    return UIA_TextControlTypeId;
}

[[nodiscard]] bool SupportsInvokePattern(const Control* control) noexcept
{
    return dynamic_cast<const Button*>(control) != nullptr && dynamic_cast<const Toggle*>(control) == nullptr;
}

[[nodiscard]] bool SupportsTogglePattern(const Control* control) noexcept
{
    return dynamic_cast<const Toggle*>(control) != nullptr;
}

[[nodiscard]] bool SupportsValuePattern(const Control* control) noexcept
{
    if (const auto* textField = dynamic_cast<const TextField*>(control))
    {
        return ! textField->IsMultiline();
    }
    return dynamic_cast<const ComboBox*>(control) != nullptr;
}

[[nodiscard]] bool SupportsSelectionProviderPattern(const Control* control) noexcept
{
    return dynamic_cast<const Tree*>(control) != nullptr || dynamic_cast<const Grid*>(control) != nullptr;
}

[[nodiscard]] bool SupportsRangeValuePattern(const Control* control) noexcept
{
    return dynamic_cast<const Slider*>(control) != nullptr;
}

[[nodiscard]] bool IsValueReadOnly(const Control* control) noexcept
{
    if (const auto* textField = dynamic_cast<const TextField*>(control))
    {
        return textField->IsReadOnly();
    }
    if (const auto* comboBox = dynamic_cast<const ComboBox*>(control))
    {
        return ! comboBox->IsEditable();
    }
    if (dynamic_cast<const Slider*>(control) != nullptr)
    {
        return false;
    }
    return true;
}

[[nodiscard]] VARIANT VariantFromBool(bool value) noexcept
{
    VARIANT variant{};
    variant.vt      = VT_BOOL;
    variant.boolVal = value ? VARIANT_TRUE : VARIANT_FALSE;
    return variant;
}

[[nodiscard]] VARIANT VariantFromInt(int value) noexcept
{
    VARIANT variant{};
    variant.vt   = VT_I4;
    variant.lVal = value;
    return variant;
}

[[nodiscard]] VARIANT VariantFromDouble(double value) noexcept
{
    VARIANT variant{};
    variant.vt     = VT_R8;
    variant.dblVal = value;
    return variant;
}

HRESULT SetVariantFromString(VARIANT* outVariant, std::wstring_view value) noexcept
{
    if (! outVariant)
    {
        return E_POINTER;
    }

    VariantInit(outVariant);
    BSTR text = SysAllocStringLen(value.data(), static_cast<UINT>(value.size()));
    if (! text && ! value.empty())
    {
        return E_OUTOFMEMORY;
    }

    outVariant->vt      = VT_BSTR;
    outVariant->bstrVal = text;
    return S_OK;
}

HRESULT SetRuntimeId(SAFEARRAY** outArray, HWND hwnd, const ControlPath* path) noexcept
{
    if (! outArray)
    {
        return E_POINTER;
    }

    *outArray          = nullptr;
    const ULONG length = path ? (path->depth + 4u) : 3u;
    unique_safearray array(SafeArrayCreateVector(VT_I4, 0, length));
    if (! array)
    {
        return E_OUTOFMEMORY;
    }

    LONG index = 0;
    LONG value = UiaAppendRuntimeId;
    HRESULT hr = SafeArrayPutElement(array.get(), &index, &value);
    if (FAILED(hr))
    {
        return hr;
    }

    ++index;
    value = HandleToLong(hwnd);
    hr    = SafeArrayPutElement(array.get(), &index, &value);
    if (FAILED(hr))
    {
        return hr;
    }

    ++index;
    value = static_cast<LONG>((static_cast<ULONG_PTR>(reinterpret_cast<ULONG_PTR>(hwnd)) >> 32) & 0x7FFFFFFF);
    hr    = SafeArrayPutElement(array.get(), &index, &value);
    if (FAILED(hr))
    {
        return hr;
    }

    if (path)
    {
        for (uint32_t depth = 0u; depth < path->depth; ++depth)
        {
            ++index;
            value = static_cast<LONG>(path->indices[depth]);
            hr    = SafeArrayPutElement(array.get(), &index, &value);
            if (FAILED(hr))
            {
                return hr;
            }
        }

        ++index;
        value = static_cast<LONG>(path->depth);
        hr    = SafeArrayPutElement(array.get(), &index, &value);
        if (FAILED(hr))
        {
            return hr;
        }
    }

    *outArray = array.release();
    return S_OK;
}

HRESULT SetTextFieldPasswordRevealButtonRuntimeId(SAFEARRAY** outArray, const ControlPath& path) noexcept
{
    if (! outArray)
    {
        return E_POINTER;
    }

    *outArray          = nullptr;
    const ULONG length = path.depth + 4u;
    unique_safearray array(SafeArrayCreateVector(VT_I4, 0, length));
    if (! array)
    {
        return E_OUTOFMEMORY;
    }

    LONG index = 0;
    LONG value = UiaAppendRuntimeId;
    HRESULT hr = SafeArrayPutElement(array.get(), &index, &value);
    if (FAILED(hr))
    {
        return hr;
    }

    for (uint32_t depth = 0u; depth < path.depth; ++depth)
    {
        ++index;
        value = static_cast<LONG>(path.indices[depth]);
        hr    = SafeArrayPutElement(array.get(), &index, &value);
        if (FAILED(hr))
        {
            return hr;
        }
    }

    ++index;
    value = static_cast<LONG>(path.depth);
    hr    = SafeArrayPutElement(array.get(), &index, &value);
    if (FAILED(hr))
    {
        return hr;
    }

    ++index;
    value = 1'005;
    hr    = SafeArrayPutElement(array.get(), &index, &value);
    if (FAILED(hr))
    {
        return hr;
    }

    *outArray = array.release();
    return S_OK;
}

HRESULT SetTreeItemRuntimeId(SAFEARRAY** outArray, const ControlPath& path, size_t visibleIndex) noexcept
{
    if (! outArray)
    {
        return E_POINTER;
    }

    *outArray = nullptr;
    if (visibleIndex > static_cast<size_t>((std::numeric_limits<LONG>::max)()))
    {
        return E_INVALIDARG;
    }

    const ULONG length = path.depth + 4u;
    unique_safearray array(SafeArrayCreateVector(VT_I4, 0, length));
    if (! array)
    {
        return E_OUTOFMEMORY;
    }

    LONG index = 0;
    LONG value = UiaAppendRuntimeId;
    HRESULT hr = SafeArrayPutElement(array.get(), &index, &value);
    if (FAILED(hr))
    {
        return hr;
    }

    for (uint32_t depth = 0u; depth < path.depth; ++depth)
    {
        ++index;
        value = static_cast<LONG>(path.indices[depth]);
        hr    = SafeArrayPutElement(array.get(), &index, &value);
        if (FAILED(hr))
        {
            return hr;
        }
    }

    ++index;
    value = static_cast<LONG>(path.depth);
    hr    = SafeArrayPutElement(array.get(), &index, &value);
    if (FAILED(hr))
    {
        return hr;
    }

    ++index;
    value = 1'001;
    hr    = SafeArrayPutElement(array.get(), &index, &value);
    if (FAILED(hr))
    {
        return hr;
    }

    ++index;
    value = static_cast<LONG>(visibleIndex);
    hr    = SafeArrayPutElement(array.get(), &index, &value);
    if (FAILED(hr))
    {
        return hr;
    }

    *outArray = array.release();
    return S_OK;
}

HRESULT SetGridRowRuntimeId(SAFEARRAY** outArray, const ControlPath& path, uint64_t rowId) noexcept
{
    if (! outArray)
    {
        return E_POINTER;
    }

    *outArray          = nullptr;
    const ULONG length = path.depth + 5u;
    unique_safearray array(SafeArrayCreateVector(VT_I4, 0, length));
    if (! array)
    {
        return E_OUTOFMEMORY;
    }

    LONG index = 0;
    LONG value = UiaAppendRuntimeId;
    HRESULT hr = SafeArrayPutElement(array.get(), &index, &value);
    if (FAILED(hr))
    {
        return hr;
    }

    for (uint32_t depth = 0u; depth < path.depth; ++depth)
    {
        ++index;
        value = static_cast<LONG>(path.indices[depth]);
        hr    = SafeArrayPutElement(array.get(), &index, &value);
        if (FAILED(hr))
        {
            return hr;
        }
    }

    ++index;
    value = static_cast<LONG>(path.depth);
    hr    = SafeArrayPutElement(array.get(), &index, &value);
    if (FAILED(hr))
    {
        return hr;
    }

    ++index;
    value = 1'002;
    hr    = SafeArrayPutElement(array.get(), &index, &value);
    if (FAILED(hr))
    {
        return hr;
    }

    ++index;
    value = static_cast<LONG>(rowId & 0xFFFFFFFFu);
    hr    = SafeArrayPutElement(array.get(), &index, &value);
    if (FAILED(hr))
    {
        return hr;
    }

    ++index;
    value = static_cast<LONG>((rowId >> 32u) & 0xFFFFFFFFu);
    hr    = SafeArrayPutElement(array.get(), &index, &value);
    if (FAILED(hr))
    {
        return hr;
    }

    *outArray = array.release();
    return S_OK;
}

HRESULT SetGridHeaderRuntimeId(SAFEARRAY** outArray, const ControlPath& path, size_t columnIndex) noexcept
{
    if (! outArray)
    {
        return E_POINTER;
    }

    *outArray = nullptr;
    if (columnIndex > static_cast<size_t>((std::numeric_limits<LONG>::max)()))
    {
        return E_INVALIDARG;
    }

    const ULONG length = path.depth + 4u;
    unique_safearray array(SafeArrayCreateVector(VT_I4, 0, length));
    if (! array)
    {
        return E_OUTOFMEMORY;
    }

    LONG index = 0;
    LONG value = UiaAppendRuntimeId;
    HRESULT hr = SafeArrayPutElement(array.get(), &index, &value);
    if (FAILED(hr))
    {
        return hr;
    }

    for (uint32_t depth = 0u; depth < path.depth; ++depth)
    {
        ++index;
        value = static_cast<LONG>(path.indices[depth]);
        hr    = SafeArrayPutElement(array.get(), &index, &value);
        if (FAILED(hr))
        {
            return hr;
        }
    }

    ++index;
    value = static_cast<LONG>(path.depth);
    hr    = SafeArrayPutElement(array.get(), &index, &value);
    if (FAILED(hr))
    {
        return hr;
    }

    ++index;
    value = 1'004;
    hr    = SafeArrayPutElement(array.get(), &index, &value);
    if (FAILED(hr))
    {
        return hr;
    }

    ++index;
    value = static_cast<LONG>(columnIndex);
    hr    = SafeArrayPutElement(array.get(), &index, &value);
    if (FAILED(hr))
    {
        return hr;
    }

    *outArray = array.release();
    return S_OK;
}

HRESULT SetGridCellRuntimeId(SAFEARRAY** outArray, const ControlPath& path, uint64_t rowId, size_t columnIndex) noexcept
{
    if (! outArray)
    {
        return E_POINTER;
    }

    *outArray = nullptr;
    if (columnIndex > static_cast<size_t>((std::numeric_limits<LONG>::max)()))
    {
        return E_INVALIDARG;
    }

    const ULONG length = path.depth + 6u;
    unique_safearray array(SafeArrayCreateVector(VT_I4, 0, length));
    if (! array)
    {
        return E_OUTOFMEMORY;
    }

    LONG index = 0;
    LONG value = UiaAppendRuntimeId;
    HRESULT hr = SafeArrayPutElement(array.get(), &index, &value);
    if (FAILED(hr))
    {
        return hr;
    }

    for (uint32_t depth = 0u; depth < path.depth; ++depth)
    {
        ++index;
        value = static_cast<LONG>(path.indices[depth]);
        hr    = SafeArrayPutElement(array.get(), &index, &value);
        if (FAILED(hr))
        {
            return hr;
        }
    }

    ++index;
    value = static_cast<LONG>(path.depth);
    hr    = SafeArrayPutElement(array.get(), &index, &value);
    if (FAILED(hr))
    {
        return hr;
    }

    ++index;
    value = 1'003;
    hr    = SafeArrayPutElement(array.get(), &index, &value);
    if (FAILED(hr))
    {
        return hr;
    }

    ++index;
    value = static_cast<LONG>(rowId & 0xFFFFFFFFu);
    hr    = SafeArrayPutElement(array.get(), &index, &value);
    if (FAILED(hr))
    {
        return hr;
    }

    ++index;
    value = static_cast<LONG>((rowId >> 32u) & 0xFFFFFFFFu);
    hr    = SafeArrayPutElement(array.get(), &index, &value);
    if (FAILED(hr))
    {
        return hr;
    }

    ++index;
    value = static_cast<LONG>(columnIndex);
    hr    = SafeArrayPutElement(array.get(), &index, &value);
    if (FAILED(hr))
    {
        return hr;
    }

    *outArray = array.release();
    return S_OK;
}

HRESULT SetProviderArray(SAFEARRAY** outArray, std::span<IRawElementProviderSimple* const> providers) noexcept
{
    if (! outArray)
    {
        return E_POINTER;
    }

    *outArray = nullptr;
    unique_safearray array(SafeArrayCreateVector(VT_UNKNOWN, 0, static_cast<ULONG>(providers.size())));
    if (! array)
    {
        return E_OUTOFMEMORY;
    }

    for (LONG index = 0; index < static_cast<LONG>(providers.size()); ++index)
    {
        IRawElementProviderSimple* provider = providers[static_cast<size_t>(index)];
        const HRESULT hr                    = SafeArrayPutElement(array.get(), &index, provider);
        if (FAILED(hr))
        {
            return hr;
        }
    }

    *outArray = array.release();
    return S_OK;
}

HRESULT SetTextRangeProviderArray(SAFEARRAY** outArray, std::span<ITextRangeProvider* const> providers) noexcept
{
    if (! outArray)
    {
        return E_POINTER;
    }

    *outArray = nullptr;
    unique_safearray array(SafeArrayCreateVector(VT_UNKNOWN, 0, static_cast<ULONG>(providers.size())));
    if (! array)
    {
        return E_OUTOFMEMORY;
    }

    for (LONG index = 0; index < static_cast<LONG>(providers.size()); ++index)
    {
        ITextRangeProvider* provider = providers[static_cast<size_t>(index)];
        const HRESULT hr            = SafeArrayPutElement(array.get(), &index, provider);
        if (FAILED(hr))
        {
            return hr;
        }
    }

    *outArray = array.release();
    return S_OK;
}

HRESULT SetDoubleArray(SAFEARRAY** outArray, std::span<const double> values) noexcept
{
    if (! outArray)
    {
        return E_POINTER;
    }

    *outArray = nullptr;
    unique_safearray array(SafeArrayCreateVector(VT_R8, 0, static_cast<ULONG>(values.size())));
    if (! array)
    {
        return E_OUTOFMEMORY;
    }

    for (LONG index = 0; index < static_cast<LONG>(values.size()); ++index)
    {
        double value     = values[static_cast<size_t>(index)];
        const HRESULT hr = SafeArrayPutElement(array.get(), &index, &value);
        if (FAILED(hr))
        {
            return hr;
        }
    }

    *outArray = array.release();
    return S_OK;
}

class AccessibilityTextRangeProvider final : public ITextRangeProvider
{
public:
    AccessibilityTextRangeProvider(WindowHostAccessibilityTarget* target, HWND hwnd, const ControlPath& path, size_t start, size_t end) noexcept
        : _target(target),
          _hwnd(hwnd),
          _path(path),
          _rangeStart(start),
          _rangeEnd(end)
    {
    }

    AccessibilityTextRangeProvider(const AccessibilityTextRangeProvider&)            = delete;
    AccessibilityTextRangeProvider& operator=(const AccessibilityTextRangeProvider&) = delete;
    AccessibilityTextRangeProvider(AccessibilityTextRangeProvider&&)                 = delete;
    AccessibilityTextRangeProvider& operator=(AccessibilityTextRangeProvider&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override;
    ULONG STDMETHODCALLTYPE AddRef() noexcept override;
    ULONG STDMETHODCALLTYPE Release() noexcept override;

    HRESULT STDMETHODCALLTYPE Clone(ITextRangeProvider** outClone) noexcept override;
    HRESULT STDMETHODCALLTYPE Compare(ITextRangeProvider* range, BOOL* outSame) noexcept override;
    HRESULT STDMETHODCALLTYPE CompareEndpoints(TextPatternRangeEndpoint endpoint,
                                               ITextRangeProvider* targetRange,
                                               TextPatternRangeEndpoint targetEndpoint,
                                               int* outResult) noexcept override;
    HRESULT STDMETHODCALLTYPE ExpandToEnclosingUnit(TextUnit unit) noexcept override;
    HRESULT STDMETHODCALLTYPE FindAttribute(TEXTATTRIBUTEID attributeId,
                                            VARIANT value,
                                            BOOL backward,
                                            ITextRangeProvider** outRange) noexcept override;
    HRESULT STDMETHODCALLTYPE FindText(BSTR text, BOOL backward, BOOL ignoreCase, ITextRangeProvider** outRange) noexcept override;
    HRESULT STDMETHODCALLTYPE GetAttributeValue(TEXTATTRIBUTEID attributeId, VARIANT* outValue) noexcept override;
    HRESULT STDMETHODCALLTYPE GetBoundingRectangles(SAFEARRAY** outRectangles) noexcept override;
    HRESULT STDMETHODCALLTYPE GetEnclosingElement(IRawElementProviderSimple** outElement) noexcept override;
    HRESULT STDMETHODCALLTYPE GetText(int maxLength, BSTR* outText) noexcept override;
    HRESULT STDMETHODCALLTYPE Move(TextUnit unit, int count, int* outMoved) noexcept override;
    HRESULT STDMETHODCALLTYPE MoveEndpointByUnit(TextPatternRangeEndpoint endpoint, TextUnit unit, int count, int* outMoved) noexcept override;
    HRESULT STDMETHODCALLTYPE MoveEndpointByRange(TextPatternRangeEndpoint endpoint,
                                                  ITextRangeProvider* targetRange,
                                                  TextPatternRangeEndpoint targetEndpoint) noexcept override;
    HRESULT STDMETHODCALLTYPE Select() noexcept override;
    HRESULT STDMETHODCALLTYPE AddToSelection() noexcept override;
    HRESULT STDMETHODCALLTYPE RemoveFromSelection() noexcept override;
    HRESULT STDMETHODCALLTYPE ScrollIntoView(BOOL alignToTop) noexcept override;
    HRESULT STDMETHODCALLTYPE GetChildren(SAFEARRAY** outChildren) noexcept override;
    HRESULT ExecuteSelectOnWindowThread() noexcept;

private:
    [[nodiscard]] bool IsCurrentThreadWindowThread() const noexcept;
    HRESULT DispatchActionToWindowThread(AccessibilityUiActionRequest& request) const noexcept;
    [[nodiscard]] WindowHost* ResolveHost() const noexcept;
    [[nodiscard]] const Control* ResolveControl() const noexcept;
    [[nodiscard]] Control* ResolveMutableControl() const noexcept;
    [[nodiscard]] std::wstring ResolveText() const;
    [[nodiscard]] TextRangeSpan ClampCurrentRange(size_t textLength) const noexcept;
    [[nodiscard]] ITextRangeProvider* CreateRange(size_t start, size_t end) const noexcept;

    std::atomic<ULONG> _referenceCount{1u};
    WindowHostAccessibilityTarget* _target = nullptr;
    HWND _hwnd                             = nullptr;
    ControlPath _path{};
    size_t _rangeStart = 0u;
    size_t _rangeEnd   = 0u;
};

class AccessibilityProvider final : public IRawElementProviderSimple,
                                    public IRawElementProviderFragment,
                                    public IRawElementProviderFragmentRoot,
                                    public IInvokeProvider,
                                    public ITextEditProvider,
                                    public ITableProvider,
                                    public IToggleProvider,
                                    public IValueProvider,
                                    public IRangeValueProvider,
                                    public ISelectionProvider,
                                    public ISelectionItemProvider,
                                    public IExpandCollapseProvider,
                                    public IGridItemProvider,
                                    public ITableItemProvider
{
public:
    struct GridHeaderTag
    {
    };

    AccessibilityProvider(WindowHostAccessibilityTarget* target, HWND hwnd) noexcept : _target(target), _hwnd(hwnd)
    {
    }

    AccessibilityProvider(WindowHostAccessibilityTarget* target, HWND hwnd, const ControlPath& path) noexcept
        : _target(target),
          _hwnd(hwnd),
          _path(path),
          _kind(AccessibilityFragmentKind::Control)
    {
    }

    AccessibilityProvider(WindowHostAccessibilityTarget* target, HWND hwnd, const ControlPath& path, AccessibilityFragmentKind kind) noexcept
        : _target(target),
          _hwnd(hwnd),
          _path(path),
          _kind(kind)
    {
    }

    AccessibilityProvider(WindowHostAccessibilityTarget* target, HWND hwnd, const ControlPath& path, size_t treeVisibleIndex) noexcept
        : _target(target),
          _hwnd(hwnd),
          _path(path),
          _kind(AccessibilityFragmentKind::TreeItem),
          _treeVisibleIndex(treeVisibleIndex)
    {
    }

    AccessibilityProvider(WindowHostAccessibilityTarget* target, HWND hwnd, const ControlPath& path, size_t gridColumnIndex, GridHeaderTag) noexcept
        : _target(target),
          _hwnd(hwnd),
          _path(path),
          _kind(AccessibilityFragmentKind::GridHeader),
          _gridColumnIndex(gridColumnIndex)
    {
    }

    AccessibilityProvider(
        WindowHostAccessibilityTarget* target, HWND hwnd, const ControlPath& path, uint64_t gridRowId, AccessibilityFragmentKind kind) noexcept
        : _target(target),
          _hwnd(hwnd),
          _path(path),
          _kind(kind),
          _gridRowId(gridRowId)
    {
    }

    AccessibilityProvider(WindowHostAccessibilityTarget* target, HWND hwnd, const ControlPath& path, uint64_t gridRowId, size_t gridColumnIndex) noexcept
        : _target(target),
          _hwnd(hwnd),
          _path(path),
          _kind(AccessibilityFragmentKind::GridCell),
          _gridRowId(gridRowId),
          _gridColumnIndex(gridColumnIndex)
    {
    }

    AccessibilityProvider(const AccessibilityProvider&)            = delete;
    AccessibilityProvider& operator=(const AccessibilityProvider&) = delete;
    AccessibilityProvider(AccessibilityProvider&&)                 = delete;
    AccessibilityProvider& operator=(AccessibilityProvider&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override;
    ULONG STDMETHODCALLTYPE AddRef() noexcept override;
    ULONG STDMETHODCALLTYPE Release() noexcept override;

    HRESULT STDMETHODCALLTYPE get_ProviderOptions(ProviderOptions* outOptions) noexcept override;
    HRESULT STDMETHODCALLTYPE GetPatternProvider(PATTERNID patternId, IUnknown** outProvider) noexcept override;
    HRESULT STDMETHODCALLTYPE GetPropertyValue(PROPERTYID propertyId, VARIANT* outValue) noexcept override;
    HRESULT STDMETHODCALLTYPE get_HostRawElementProvider(IRawElementProviderSimple** outProvider) noexcept override;

    HRESULT STDMETHODCALLTYPE Navigate(NavigateDirection direction, IRawElementProviderFragment** outProvider) noexcept override;
    HRESULT STDMETHODCALLTYPE GetRuntimeId(SAFEARRAY** outRuntimeId) noexcept override;
    HRESULT STDMETHODCALLTYPE get_BoundingRectangle(UiaRect* outRect) noexcept override;
    HRESULT STDMETHODCALLTYPE GetEmbeddedFragmentRoots(SAFEARRAY** outRoots) noexcept override;
    HRESULT STDMETHODCALLTYPE SetFocus() noexcept override;
    HRESULT STDMETHODCALLTYPE get_FragmentRoot(IRawElementProviderFragmentRoot** outRoot) noexcept override;

    HRESULT STDMETHODCALLTYPE ElementProviderFromPoint(double x, double y, IRawElementProviderFragment** outProvider) noexcept override;
    HRESULT STDMETHODCALLTYPE GetFocus(IRawElementProviderFragment** outProvider) noexcept override;

    HRESULT STDMETHODCALLTYPE Invoke() noexcept override;
    HRESULT STDMETHODCALLTYPE Toggle() noexcept override;
    HRESULT STDMETHODCALLTYPE get_ToggleState(ToggleState* outState) noexcept override;
    HRESULT STDMETHODCALLTYPE GetVisibleRanges(SAFEARRAY** outRanges) noexcept override;
    HRESULT STDMETHODCALLTYPE RangeFromChild(IRawElementProviderSimple* childElement, ITextRangeProvider** outRange) noexcept override;
    HRESULT STDMETHODCALLTYPE RangeFromPoint(UiaPoint point, ITextRangeProvider** outRange) noexcept override;
    HRESULT STDMETHODCALLTYPE get_DocumentRange(ITextRangeProvider** outRange) noexcept override;
    HRESULT STDMETHODCALLTYPE get_SupportedTextSelection(SupportedTextSelection* outSupportedSelection) noexcept override;
    HRESULT STDMETHODCALLTYPE GetActiveComposition(ITextRangeProvider** outRange) noexcept override;
    HRESULT STDMETHODCALLTYPE GetConversionTarget(ITextRangeProvider** outRange) noexcept override;
    HRESULT STDMETHODCALLTYPE SetValue(LPCWSTR value) noexcept override;
    HRESULT STDMETHODCALLTYPE SetValue(double value) noexcept override;
    HRESULT STDMETHODCALLTYPE get_Value(BSTR* outValue) noexcept override;
    HRESULT STDMETHODCALLTYPE get_Value(double* outValue) noexcept override;
    HRESULT STDMETHODCALLTYPE get_IsReadOnly(BOOL* outReadOnly) noexcept override;
    HRESULT STDMETHODCALLTYPE get_Maximum(double* outMaximum) noexcept override;
    HRESULT STDMETHODCALLTYPE get_Minimum(double* outMinimum) noexcept override;
    HRESULT STDMETHODCALLTYPE get_LargeChange(double* outLargeChange) noexcept override;
    HRESULT STDMETHODCALLTYPE get_SmallChange(double* outSmallChange) noexcept override;
    HRESULT STDMETHODCALLTYPE GetSelection(SAFEARRAY** outSelection) noexcept override;
    HRESULT STDMETHODCALLTYPE get_CanSelectMultiple(BOOL* outCanSelectMultiple) noexcept override;
    HRESULT STDMETHODCALLTYPE get_IsSelectionRequired(BOOL* outIsSelectionRequired) noexcept override;
    HRESULT STDMETHODCALLTYPE GetRowHeaders(SAFEARRAY** outRowHeaders) noexcept override;
    HRESULT STDMETHODCALLTYPE GetColumnHeaders(SAFEARRAY** outColumnHeaders) noexcept override;
    HRESULT STDMETHODCALLTYPE get_RowOrColumnMajor(RowOrColumnMajor* outRowOrColumnMajor) noexcept override;
    HRESULT STDMETHODCALLTYPE Select() noexcept override;
    HRESULT STDMETHODCALLTYPE AddToSelection() noexcept override;
    HRESULT STDMETHODCALLTYPE RemoveFromSelection() noexcept override;
    HRESULT STDMETHODCALLTYPE get_IsSelected(BOOL* outSelected) noexcept override;
    HRESULT STDMETHODCALLTYPE get_SelectionContainer(IRawElementProviderSimple** outContainer) noexcept override;
    HRESULT STDMETHODCALLTYPE Expand() noexcept override;
    HRESULT STDMETHODCALLTYPE Collapse() noexcept override;
    HRESULT STDMETHODCALLTYPE get_ExpandCollapseState(ExpandCollapseState* outState) noexcept override;
    HRESULT STDMETHODCALLTYPE get_Row(int* outRow) noexcept override;
    HRESULT STDMETHODCALLTYPE get_Column(int* outColumn) noexcept override;
    HRESULT STDMETHODCALLTYPE get_RowSpan(int* outRowSpan) noexcept override;
    HRESULT STDMETHODCALLTYPE get_ColumnSpan(int* outColumnSpan) noexcept override;
    HRESULT STDMETHODCALLTYPE get_ContainingGrid(IRawElementProviderSimple** outContainingGrid) noexcept override;
    HRESULT STDMETHODCALLTYPE GetRowHeaderItems(SAFEARRAY** outRowHeaderItems) noexcept override;
    HRESULT STDMETHODCALLTYPE GetColumnHeaderItems(SAFEARRAY** outColumnHeaderItems) noexcept override;
    HRESULT ExecuteUiThreadAction(AccessibilityUiActionKind kind, LPCWSTR stringValue, double numberValue) noexcept;

private:
    [[nodiscard]] WindowHost* ResolveHost() const noexcept;
    [[nodiscard]] const Control* ResolveRootControl() const noexcept;
    [[nodiscard]] bool ResolveControlPath(ControlPath& outPath) const noexcept;
    [[nodiscard]] const Control* ResolveControl() const noexcept;
    [[nodiscard]] Control* ResolveMutableControl() const noexcept;
    [[nodiscard]] const Tree* ResolveTreeControl() const noexcept;
    [[nodiscard]] Tree* ResolveMutableTreeControl() const noexcept;
    [[nodiscard]] const Grid* ResolveGridControl() const noexcept;
    [[nodiscard]] Grid* ResolveMutableGridControl() const noexcept;
    [[nodiscard]] bool ResolveTreeItemData(TreeItemData& outItem) const noexcept;
    [[nodiscard]] bool ResolveGridHeaderColumn(size_t& outColumnIndex, GridColumnDesc& outColumnDesc) const noexcept;
    [[nodiscard]] bool ResolveGridRowIndex(size_t& outRowIndex) const noexcept;
    [[nodiscard]] bool ResolveGridCellData(size_t& outRowIndex, size_t& outColumnIndex, GridCellData& outCellData) const noexcept;
    [[nodiscard]] bool SupportsTreeItemSelectionPattern() const noexcept;
    [[nodiscard]] bool SupportsTreeItemExpandCollapsePattern() const noexcept;
    [[nodiscard]] bool SupportsGridTablePattern() const noexcept;
    [[nodiscard]] bool SupportsGridRowSelectionPattern() const noexcept;
    [[nodiscard]] bool SupportsGridCellPattern() const noexcept;
    [[nodiscard]] bool SupportsGridCellTableItemPattern() const noexcept;
    [[nodiscard]] bool SupportsGridCellTogglePattern() const noexcept;
    [[nodiscard]] bool SupportsGridCellValuePattern() const noexcept;
    [[nodiscard]] bool SupportsGridCellRangeValuePattern() const noexcept;
    [[nodiscard]] std::optional<size_t> FindTreeItemAtPoint(WindowHost& host, const Tree& tree, D2D1_POINT_2F pointDip) const noexcept;
    [[nodiscard]] IRawElementProviderFragmentRoot* CreateRootProvider() noexcept;
    [[nodiscard]] IRawElementProviderFragment* CreateChildProvider(const ControlPath& path) noexcept;
    [[nodiscard]] IRawElementProviderFragment* CreateTextFieldPasswordRevealButtonProvider(const ControlPath& path) noexcept;
    [[nodiscard]] IRawElementProviderFragment* CreateTreeItemProvider(const ControlPath& path, size_t visibleIndex) noexcept;
    [[nodiscard]] IRawElementProviderFragment* CreateGridHeaderProvider(const ControlPath& path, size_t columnIndex) noexcept;
    [[nodiscard]] IRawElementProviderFragment* CreateGridRowProvider(const ControlPath& path, uint64_t rowId) noexcept;
    [[nodiscard]] IRawElementProviderFragment* CreateGridCellProvider(const ControlPath& path, uint64_t rowId, size_t columnIndex) noexcept;
    [[nodiscard]] ITextRangeProvider* CreateTextRangeProvider(size_t start, size_t end) noexcept;
    [[nodiscard]] ITextRangeProvider* CreateTextDocumentRangeProvider(const Control* control) noexcept;
    static bool FindPathForTarget(const Control* current, const ControlPath& basePath, const Control* target, ControlPath& outPath) noexcept;
    [[nodiscard]] WindowHostAccessibilityTarget* AddRefTarget() const noexcept;
    [[nodiscard]] bool IsCurrentThreadWindowThread() const noexcept;
    HRESULT DispatchActionToWindowThread(AccessibilityUiActionRequest& request) noexcept;
    HRESULT ExecuteSetFocusOnWindowThread() noexcept;
    HRESULT ExecuteInvokeOnWindowThread() noexcept;
    HRESULT ExecuteToggleOnWindowThread() noexcept;
    HRESULT ExecuteSetStringValueOnWindowThread(LPCWSTR value) noexcept;
    HRESULT ExecuteSetRangeValueOnWindowThread(double value) noexcept;
    HRESULT ExecuteSelectOnWindowThread() noexcept;
    HRESULT ExecuteAddToSelectionOnWindowThread() noexcept;
    HRESULT ExecuteRemoveFromSelectionOnWindowThread() noexcept;
    HRESULT ExecuteExpandOnWindowThread(bool expanded) noexcept;

    std::atomic<ULONG> _referenceCount{1u};
    WindowHostAccessibilityTarget* _target = nullptr;
    HWND _hwnd                             = nullptr;
    ControlPath _path{};
    AccessibilityFragmentKind _kind = AccessibilityFragmentKind::Root;
    size_t _treeVisibleIndex        = 0u;
    uint64_t _gridRowId             = 0u;
    size_t _gridColumnIndex         = 0u;
};

HRESULT AccessibilityTextRangeProvider::QueryInterface(REFIID riid, void** ppvObject) noexcept
{
    if (! ppvObject)
    {
        return E_POINTER;
    }

    *ppvObject = nullptr;
    if (riid == __uuidof(IUnknown) || riid == __uuidof(ITextRangeProvider))
    {
        *ppvObject = static_cast<ITextRangeProvider*>(this);
        AddRef();
        return S_OK;
    }

    return E_NOINTERFACE;
}

ULONG AccessibilityTextRangeProvider::AddRef() noexcept
{
    return _referenceCount.fetch_add(1u, std::memory_order_relaxed) + 1u;
}

ULONG AccessibilityTextRangeProvider::Release() noexcept
{
    const ULONG remaining = _referenceCount.fetch_sub(1u, std::memory_order_acq_rel) - 1u;
    if (remaining == 0u)
    {
        if (_target)
        {
            static_cast<void>(_target->Release());
            _target = nullptr;
        }
        delete this;
    }
    return remaining;
}

HRESULT AccessibilityTextRangeProvider::Clone(ITextRangeProvider** outClone) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outClone)
    {
        return E_POINTER;
    }

    *outClone = CreateRange(_rangeStart, _rangeEnd);
    return *outClone ? S_OK : E_OUTOFMEMORY;
}

HRESULT AccessibilityTextRangeProvider::Compare(ITextRangeProvider* range, BOOL* outSame) noexcept
{
    if (! outSame)
    {
        return E_POINTER;
    }

    const auto* other = dynamic_cast<AccessibilityTextRangeProvider*>(range);
    *outSame = (other && _target == other->_target && _hwnd == other->_hwnd && AreControlPathsEqual(_path, other->_path) &&
                _rangeStart == other->_rangeStart && _rangeEnd == other->_rangeEnd)
                   ? TRUE
                   : FALSE;
    return S_OK;
}

HRESULT AccessibilityTextRangeProvider::CompareEndpoints(TextPatternRangeEndpoint endpoint,
                                                         ITextRangeProvider* targetRange,
                                                         TextPatternRangeEndpoint targetEndpoint,
                                                         int* outResult) noexcept
{
    if (! outResult)
    {
        return E_POINTER;
    }
    if (! targetRange)
    {
        return E_INVALIDARG;
    }

    const auto* other = dynamic_cast<AccessibilityTextRangeProvider*>(targetRange);
    if (! other || _target != other->_target || _hwnd != other->_hwnd || ! AreControlPathsEqual(_path, other->_path))
    {
        return E_INVALIDARG;
    }

    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    const std::wstring text = ResolveText();
    const TextRangeSpan range = ClampCurrentRange(text.size());
    const TextRangeSpan otherRange = other->ClampCurrentRange(text.size());
    const size_t position = GetTextRangeEndpointPosition(range, endpoint);
    const size_t targetPosition = GetTextRangeEndpointPosition(otherRange, targetEndpoint);

    *outResult = position < targetPosition ? -1 : (position > targetPosition ? 1 : 0);
    return S_OK;
}

HRESULT AccessibilityTextRangeProvider::ExpandToEnclosingUnit(TextUnit unit) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    const std::wstring text = ResolveText();
    if (unit == TextUnit_Document)
    {
        _rangeStart = 0u;
        _rangeEnd   = text.size();
    }
    else
    {
        const TextRangeSpan range = ClampCurrentRange(text.size());
        _rangeStart              = range.start;
        _rangeEnd                = range.end;
    }
    return S_OK;
}

HRESULT AccessibilityTextRangeProvider::FindAttribute(TEXTATTRIBUTEID /*attributeId*/,
                                                      VARIANT /*value*/,
                                                      BOOL /*backward*/,
                                                      ITextRangeProvider** outRange) noexcept
{
    if (! outRange)
    {
        return E_POINTER;
    }

    *outRange = nullptr;
    return S_OK;
}

HRESULT AccessibilityTextRangeProvider::FindText(BSTR /*text*/, BOOL /*backward*/, BOOL /*ignoreCase*/, ITextRangeProvider** outRange) noexcept
{
    if (! outRange)
    {
        return E_POINTER;
    }

    *outRange = nullptr;
    return S_OK;
}

HRESULT AccessibilityTextRangeProvider::GetAttributeValue(TEXTATTRIBUTEID /*attributeId*/, VARIANT* outValue) noexcept
{
    if (! outValue)
    {
        return E_POINTER;
    }

    VariantInit(outValue);
    return S_OK;
}

HRESULT AccessibilityTextRangeProvider::GetBoundingRectangles(SAFEARRAY** outRectangles) noexcept
{
    if (! outRectangles)
    {
        return E_POINTER;
    }

    const auto startedAt = std::chrono::steady_clock::now();
    *outRectangles       = nullptr;

    WindowHost* const host    = ResolveHost();
    const Control* const control = ResolveControl();
    const std::wstring text   = ResolveText();
    const TextRangeSpan range = ClampCurrentRange(text.size());
    if (! host || ! control || range.start == range.end)
    {
        const HRESULT hr = SetDoubleArray(outRectangles, {});
        Debug::Perf::Emit(L"dxui.uia.text_range_us", L"bounding-rectangles-empty", Debug::Perf::ElapsedUs(startedAt), 0u, 0u, hr);
        return hr;
    }

    std::vector<D2D1_RECT_F> boundsDip =
        TryResolveTextRangeCaretRects(*host, *control, text, range).value_or(std::vector<D2D1_RECT_F>{ResolveTextPatternViewportRect(control)});
    std::vector<double> values;
    values.reserve(boundsDip.size() * 4u);
    for (const D2D1_RECT_F& rectDip : boundsDip)
    {
        POINT topLeft{static_cast<LONG>(std::lround(host->DipsToPixels(rectDip.left))), static_cast<LONG>(std::lround(host->DipsToPixels(rectDip.top)))};
        POINT bottomRight{static_cast<LONG>(std::lround(host->DipsToPixels(rectDip.right))),
                          static_cast<LONG>(std::lround(host->DipsToPixels(rectDip.bottom)))};
        ClientToScreen(_hwnd, &topLeft);
        ClientToScreen(_hwnd, &bottomRight);
        values.push_back(static_cast<double>(topLeft.x));
        values.push_back(static_cast<double>(topLeft.y));
        values.push_back(static_cast<double>(bottomRight.x - topLeft.x));
        values.push_back(static_cast<double>(bottomRight.y - topLeft.y));
    }
    const HRESULT hr = SetDoubleArray(outRectangles, values);
    Debug::Perf::Emit(
        L"dxui.uia.text_range_us", L"bounding-rectangles", Debug::Perf::ElapsedUs(startedAt), boundsDip.size(), range.end - range.start, hr);
    return hr;
}

HRESULT AccessibilityTextRangeProvider::GetEnclosingElement(IRawElementProviderSimple** outElement) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outElement)
    {
        return E_POINTER;
    }

    *outElement = nullptr;
    if (! ResolveControl())
    {
        return UIA_E_ELEMENTNOTAVAILABLE;
    }

    WindowHostAccessibilityTarget* target = _target;
    if (target)
    {
        static_cast<void>(target->AddRef());
    }

    auto* provider = new (std::nothrow) AccessibilityProvider(target, _hwnd, _path);
    if (! provider)
    {
        if (target)
        {
            static_cast<void>(target->Release());
        }
        return E_OUTOFMEMORY;
    }

    *outElement = static_cast<IRawElementProviderSimple*>(provider);
    return S_OK;
}

HRESULT AccessibilityTextRangeProvider::GetText(int maxLength, BSTR* outText) noexcept
{
    const auto startedAt = std::chrono::steady_clock::now();
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outText)
    {
        return E_POINTER;
    }

    *outText              = nullptr;
    const std::wstring text = ResolveText();
    const TextRangeSpan range = ClampCurrentRange(text.size());
    std::wstring_view value(text.data() + range.start, range.end - range.start);
    if (maxLength >= 0)
    {
        value = value.substr(0u, (std::min)(value.size(), static_cast<size_t>(maxLength)));
    }

    *outText = SysAllocStringLen(value.data(), static_cast<UINT>(value.size()));
    const HRESULT hr = (*outText || value.empty()) ? S_OK : E_OUTOFMEMORY;
    Debug::Perf::Emit(L"dxui.uia.text_range_us", L"gettext", Debug::Perf::ElapsedUs(startedAt), value.size(), text.size(), hr);
    return hr;
}

HRESULT AccessibilityTextRangeProvider::Move(TextUnit unit, int count, int* outMoved) noexcept
{
    const auto startedAt = std::chrono::steady_clock::now();
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outMoved)
    {
        return E_POINTER;
    }

    const std::wstring text = ResolveText();
    const TextRangeSpan range = ClampCurrentRange(text.size());
    *outMoved = 0;
    if ((unit != TextUnit_Character && unit != TextUnit_Word && unit != TextUnit_Line) || count == 0)
    {
        Debug::Perf::Emit(L"dxui.uia.text_range_us", L"move-unsupported", Debug::Perf::ElapsedUs(startedAt), 0u, text.size(), S_OK);
        return S_OK;
    }

    if (unit == TextUnit_Line)
    {
        WindowHost* const host          = ResolveHost();
        const Control* const control    = ResolveControl();
        const TextRangeSpanMoveResult moveResult =
            (host && control)
                ? TryMoveTextRangeSpanByVisualLine(*host, *control, text, range, count).value_or(MoveTextRangeSpanByLine(text, range, count))
                : MoveTextRangeSpanByLine(text, range, count);
        _rangeStart = moveResult.range.start;
        _rangeEnd   = moveResult.range.end;
        *outMoved   = moveResult.moved;
        Debug::Perf::Emit(
            L"dxui.uia.text_range_us", L"move-line", Debug::Perf::ElapsedUs(startedAt), _rangeEnd - _rangeStart, text.size(), S_OK);
        return S_OK;
    }

    if (unit == TextUnit_Word)
    {
        if (range.start != range.end)
        {
            const TextRangeSpanMoveResult moveResult = MoveTextRangeSpanByWord(text, range, count);
            _rangeStart                              = moveResult.range.start;
            _rangeEnd                                = moveResult.range.end;
            *outMoved                                = moveResult.moved;
            Debug::Perf::Emit(
                L"dxui.uia.text_range_us", L"move-word", Debug::Perf::ElapsedUs(startedAt), _rangeEnd - _rangeStart, text.size(), S_OK);
            return S_OK;
        }

        const TextRangeUnitMoveResult moveResult = MoveTextRangePositionByUnit(text, range.start, unit, count);
        _rangeStart = moveResult.position;
        _rangeEnd   = moveResult.position;
        *outMoved   = moveResult.moved;
        Debug::Perf::Emit(L"dxui.uia.text_range_us", L"move-word", Debug::Perf::ElapsedUs(startedAt), 0u, text.size(), S_OK);
        return S_OK;
    }

    const int64_t rangeStart = static_cast<int64_t>(range.start);
    const int64_t rangeEnd   = static_cast<int64_t>(range.end);
    const int64_t textLength = static_cast<int64_t>(text.size());
    const int64_t minimum    = -rangeStart;
    const int64_t maximum    = textLength - rangeEnd;
    const int64_t requested  = static_cast<int64_t>(count);
    const int64_t moved      = (std::min)((std::max)(requested, minimum), maximum);
    _rangeStart              = static_cast<size_t>(rangeStart + moved);
    _rangeEnd                = static_cast<size_t>(rangeEnd + moved);
    *outMoved                = static_cast<int>(moved);
    Debug::Perf::Emit(L"dxui.uia.text_range_us", L"move-character", Debug::Perf::ElapsedUs(startedAt), _rangeEnd - _rangeStart, text.size(), S_OK);
    return S_OK;
}

HRESULT AccessibilityTextRangeProvider::MoveEndpointByUnit(TextPatternRangeEndpoint endpoint,
                                                           TextUnit unit,
                                                           int count,
                                                           int* outMoved) noexcept
{
    const auto startedAt = std::chrono::steady_clock::now();
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outMoved)
    {
        return E_POINTER;
    }

    const std::wstring text = ResolveText();
    const TextRangeSpan range = ClampCurrentRange(text.size());
    *outMoved = 0;
    if ((unit != TextUnit_Character && unit != TextUnit_Word && unit != TextUnit_Line) || count == 0)
    {
        Debug::Perf::Emit(L"dxui.uia.text_range_us", L"move-endpoint-unsupported", Debug::Perf::ElapsedUs(startedAt), 0u, text.size(), S_OK);
        return S_OK;
    }

    const size_t endpointPosition = GetTextRangeEndpointPosition(range, endpoint);
    TextRangeUnitMoveResult moveResult{};
    if (unit == TextUnit_Line)
    {
        WindowHost* const host       = ResolveHost();
        const Control* const control = ResolveControl();
        moveResult = (host && control) ? TryMoveTextRangePositionByVisualLine(*host, *control, text, endpointPosition, count)
                                             .value_or(MoveTextRangePositionByLine(text, endpointPosition, count))
                                      : MoveTextRangePositionByLine(text, endpointPosition, count);
    }
    else
    {
        moveResult = MoveTextRangePositionByUnit(text, endpointPosition, unit, count);
    }
    const size_t movedTo = moveResult.position;
    *outMoved            = moveResult.moved;
    if (endpoint == TextPatternRangeEndpoint_Start)
    {
        _rangeStart = movedTo;
        _rangeEnd   = range.end < _rangeStart ? _rangeStart : range.end;
    }
    else
    {
        _rangeEnd   = movedTo;
        _rangeStart = _rangeEnd < range.start ? _rangeEnd : range.start;
    }
    const wchar_t* detail = unit == TextUnit_Line ? L"move-endpoint-line" :
                            unit == TextUnit_Word ? L"move-endpoint-word" :
                                                    L"move-endpoint-character";
    Debug::Perf::Emit(
        L"dxui.uia.text_range_us", detail, Debug::Perf::ElapsedUs(startedAt), _rangeEnd - _rangeStart, text.size(), S_OK);
    return S_OK;
}

HRESULT AccessibilityTextRangeProvider::MoveEndpointByRange(TextPatternRangeEndpoint /*endpoint*/,
                                                            ITextRangeProvider* /*targetRange*/,
                                                            TextPatternRangeEndpoint /*targetEndpoint*/) noexcept
{
    return S_OK;
}

HRESULT AccessibilityTextRangeProvider::Select() noexcept
{
    if (! IsCurrentThreadWindowThread())
    {
        AccessibilityUiActionRequest request{};
        request.textRangeProvider = this;
        request.kind              = AccessibilityUiActionKind::Select;
        return DispatchActionToWindowThread(request);
    }

    return ExecuteSelectOnWindowThread();
}

HRESULT AccessibilityTextRangeProvider::ExecuteSelectOnWindowThread() noexcept
{
    const auto startedAt = std::chrono::steady_clock::now();
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());

    WindowHost* host = ResolveHost();
    Control* control = ResolveMutableControl();
    if (! host || ! control || ! SupportsTextPattern(control))
    {
        return UIA_E_ELEMENTNOTAVAILABLE;
    }

    const std::wstring text = GetControlAccessibleTextRangeText(control);
    const TextRangeSpan range = ClampCurrentRange(text.size());
    if (auto* textField = dynamic_cast<TextField*>(control))
    {
        textField->SetSelectionRange(range.start, range.end);
        host->SyncTextInput(textField);
        host->Invalidate();
        Debug::Perf::Emit(L"dxui.uia.text_range_us", L"select", Debug::Perf::ElapsedUs(startedAt), range.end - range.start, text.size(), S_OK);
        return S_OK;
    }
    if (auto* comboBox = dynamic_cast<ComboBox*>(control); comboBox && comboBox->IsEditable())
    {
        comboBox->SetEditableSelectionRange(range.start, range.end);
        host->SyncTextInput(comboBox);
        host->Invalidate();
        Debug::Perf::Emit(L"dxui.uia.text_range_us", L"select", Debug::Perf::ElapsedUs(startedAt), range.end - range.start, text.size(), S_OK);
        return S_OK;
    }

    return UIA_E_NOTSUPPORTED;
}

HRESULT AccessibilityTextRangeProvider::AddToSelection() noexcept
{
    return S_OK;
}

HRESULT AccessibilityTextRangeProvider::RemoveFromSelection() noexcept
{
    return S_OK;
}

HRESULT AccessibilityTextRangeProvider::ScrollIntoView(BOOL /*alignToTop*/) noexcept
{
    return S_OK;
}

HRESULT AccessibilityTextRangeProvider::GetChildren(SAFEARRAY** outChildren) noexcept
{
    if (! outChildren)
    {
        return E_POINTER;
    }

    return SetProviderArray(outChildren, {});
}

bool AccessibilityTextRangeProvider::IsCurrentThreadWindowThread() const noexcept
{
    return _hwnd && GetWindowThreadProcessId(_hwnd, nullptr) == GetCurrentThreadId();
}

HRESULT AccessibilityTextRangeProvider::DispatchActionToWindowThread(AccessibilityUiActionRequest& request) const noexcept
{
    if (! _hwnd || IsWindow(_hwnd) == FALSE)
    {
        return UIA_E_ELEMENTNOTAVAILABLE;
    }

    DWORD_PTR messageResult = 0;
    if (SendMessageTimeoutW(
            _hwnd, kWindowHostAccessibilityActionMessage, 0, reinterpret_cast<LPARAM>(&request), SMTO_ABORTIFHUNG | SMTO_BLOCK, 5000u, &messageResult) == 0)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    return request.result;
}

WindowHost* AccessibilityTextRangeProvider::ResolveHost() const noexcept
{
    return (_target && _target->hwnd == _hwnd) ? _target->ResolveHost() : nullptr;
}

const Control* AccessibilityTextRangeProvider::ResolveControl() const noexcept
{
    WindowHost* host = ResolveHost();
    if (! host || ! IsControlPathVisible(host->GetRoot(), _path))
    {
        return nullptr;
    }

    return ResolveControlAtPath(host->GetRoot(), _path);
}

Control* AccessibilityTextRangeProvider::ResolveMutableControl() const noexcept
{
    WindowHost* host = ResolveHost();
    if (! host || ! IsControlPathVisible(host->GetRoot(), _path))
    {
        return nullptr;
    }

    return ResolveControlAtPath(host->GetRoot(), _path);
}

std::wstring AccessibilityTextRangeProvider::ResolveText() const
{
    const Control* control = ResolveControl();
    if (! SupportsTextPattern(control))
    {
        return {};
    }

    return GetControlAccessibleTextRangeText(control);
}

TextRangeSpan AccessibilityTextRangeProvider::ClampCurrentRange(size_t textLength) const noexcept
{
    return ClampTextRangeSpan(_rangeStart, _rangeEnd, textLength);
}

ITextRangeProvider* AccessibilityTextRangeProvider::CreateRange(size_t start, size_t end) const noexcept
{
    WindowHostAccessibilityTarget* target = _target;
    if (target)
    {
        static_cast<void>(target->AddRef());
    }

    auto* provider = new (std::nothrow) AccessibilityTextRangeProvider(target, _hwnd, _path, start, end);
    if (! provider && target)
    {
        static_cast<void>(target->Release());
    }
    return provider ? static_cast<ITextRangeProvider*>(provider) : nullptr;
}

HRESULT AccessibilityProvider::QueryInterface(REFIID riid, void** ppvObject) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! ppvObject)
    {
        return E_POINTER;
    }

    *ppvObject             = nullptr;
    const bool pathVisible = _kind == AccessibilityFragmentKind::Root || IsControlPathVisible(ResolveRootControl(), _path);
    const Control* control = pathVisible ? ResolveControl() : nullptr;
    if (riid == __uuidof(IUnknown) || riid == __uuidof(IRawElementProviderSimple))
    {
        *ppvObject = static_cast<IRawElementProviderSimple*>(this);
    }
    else if (riid == __uuidof(IRawElementProviderFragment))
    {
        *ppvObject = static_cast<IRawElementProviderFragment*>(this);
    }
    else if (_kind == AccessibilityFragmentKind::Root && riid == __uuidof(IRawElementProviderFragmentRoot))
    {
        *ppvObject = static_cast<IRawElementProviderFragmentRoot*>(this);
    }
    else if ((_kind == AccessibilityFragmentKind::Control || _kind == AccessibilityFragmentKind::Root) && riid == __uuidof(IInvokeProvider) &&
             SupportsInvokePattern(control))
    {
        *ppvObject = static_cast<IInvokeProvider*>(this);
    }
    else if (_kind == AccessibilityFragmentKind::TextFieldPasswordRevealButton && pathVisible && riid == __uuidof(IInvokeProvider) &&
             IsTextFieldPasswordRevealButtonVisible(control))
    {
        *ppvObject = static_cast<IInvokeProvider*>(this);
    }
    else if ((_kind == AccessibilityFragmentKind::Control || _kind == AccessibilityFragmentKind::Root) && riid == __uuidof(IToggleProvider) &&
             SupportsTogglePattern(control))
    {
        *ppvObject = static_cast<IToggleProvider*>(this);
    }
    else if ((_kind == AccessibilityFragmentKind::Control || _kind == AccessibilityFragmentKind::Root) && riid == __uuidof(IValueProvider) &&
             SupportsValuePattern(control))
    {
        *ppvObject = static_cast<IValueProvider*>(this);
    }
    else if ((_kind == AccessibilityFragmentKind::Control || _kind == AccessibilityFragmentKind::Root) && pathVisible &&
             riid == __uuidof(ITextProvider) && SupportsTextPattern(control))
    {
        *ppvObject = static_cast<ITextProvider*>(static_cast<ITextEditProvider*>(this));
    }
    else if ((_kind == AccessibilityFragmentKind::Control || _kind == AccessibilityFragmentKind::Root) && pathVisible &&
             riid == __uuidof(ITextEditProvider) && SupportsTextPattern(control))
    {
        *ppvObject = static_cast<ITextEditProvider*>(this);
    }
    else if ((_kind == AccessibilityFragmentKind::Control || _kind == AccessibilityFragmentKind::Root) && riid == __uuidof(IRangeValueProvider) &&
             SupportsRangeValuePattern(control))
    {
        *ppvObject = static_cast<IRangeValueProvider*>(this);
    }
    else if (pathVisible && _kind == AccessibilityFragmentKind::GridCell && riid == __uuidof(IToggleProvider) && SupportsGridCellTogglePattern())
    {
        *ppvObject = static_cast<IToggleProvider*>(this);
    }
    else if (pathVisible && _kind == AccessibilityFragmentKind::GridCell && riid == __uuidof(IValueProvider) && SupportsGridCellValuePattern())
    {
        *ppvObject = static_cast<IValueProvider*>(this);
    }
    else if (pathVisible && _kind == AccessibilityFragmentKind::GridCell && riid == __uuidof(IRangeValueProvider) && SupportsGridCellRangeValuePattern())
    {
        *ppvObject = static_cast<IRangeValueProvider*>(this);
    }
    else if ((_kind == AccessibilityFragmentKind::Control || _kind == AccessibilityFragmentKind::Root) && riid == __uuidof(ISelectionProvider) &&
             SupportsSelectionProviderPattern(control))
    {
        *ppvObject = static_cast<ISelectionProvider*>(this);
    }
    else if ((_kind == AccessibilityFragmentKind::Control || _kind == AccessibilityFragmentKind::Root) && riid == __uuidof(ITableProvider) && pathVisible &&
             SupportsGridTablePattern())
    {
        *ppvObject = static_cast<ITableProvider*>(this);
    }
    else if (pathVisible && _kind == AccessibilityFragmentKind::TreeItem && riid == __uuidof(ISelectionItemProvider) && SupportsTreeItemSelectionPattern())
    {
        *ppvObject = static_cast<ISelectionItemProvider*>(this);
    }
    else if (pathVisible && _kind == AccessibilityFragmentKind::GridRow && riid == __uuidof(ISelectionItemProvider) && SupportsGridRowSelectionPattern())
    {
        *ppvObject = static_cast<ISelectionItemProvider*>(this);
    }
    else if (pathVisible && _kind == AccessibilityFragmentKind::GridCell && riid == __uuidof(IGridItemProvider) && SupportsGridCellPattern())
    {
        *ppvObject = static_cast<IGridItemProvider*>(this);
    }
    else if (pathVisible && _kind == AccessibilityFragmentKind::GridCell && riid == __uuidof(ITableItemProvider) && SupportsGridCellTableItemPattern())
    {
        *ppvObject = static_cast<ITableItemProvider*>(this);
    }
    else if (pathVisible && _kind == AccessibilityFragmentKind::TreeItem && riid == __uuidof(IExpandCollapseProvider) &&
             SupportsTreeItemExpandCollapsePattern())
    {
        *ppvObject = static_cast<IExpandCollapseProvider*>(this);
    }
    else
    {
        return E_NOINTERFACE;
    }

    AddRef();
    return S_OK;
}

ULONG AccessibilityProvider::AddRef() noexcept
{
    return _referenceCount.fetch_add(1u, std::memory_order_relaxed) + 1u;
}

ULONG AccessibilityProvider::Release() noexcept
{
    const ULONG remaining = _referenceCount.fetch_sub(1u, std::memory_order_acq_rel) - 1u;
    if (remaining == 0u)
    {
        if (_target)
        {
            static_cast<void>(_target->Release());
            _target = nullptr;
        }
        delete this;
    }
    return remaining;
}

HRESULT AccessibilityProvider::get_ProviderOptions(ProviderOptions* outOptions) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outOptions)
    {
        return E_POINTER;
    }

    *outOptions = ProviderOptions_ServerSideProvider;
    return S_OK;
}

HRESULT AccessibilityProvider::GetPatternProvider(PATTERNID patternId, IUnknown** outProvider) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outProvider)
    {
        return E_POINTER;
    }

    *outProvider = nullptr;
    if (_kind != AccessibilityFragmentKind::Root && ! IsControlPathVisible(ResolveRootControl(), _path))
    {
        return S_OK;
    }

    if (_kind == AccessibilityFragmentKind::TreeItem)
    {
        if (patternId == UIA_SelectionItemPatternId && SupportsTreeItemSelectionPattern())
        {
            *outProvider = static_cast<ISelectionItemProvider*>(this);
        }
        else if (patternId == UIA_ExpandCollapsePatternId && SupportsTreeItemExpandCollapsePattern())
        {
            *outProvider = static_cast<IExpandCollapseProvider*>(this);
        }
        else
        {
            return S_OK;
        }

        AddRef();
        return S_OK;
    }

    if (_kind == AccessibilityFragmentKind::GridRow)
    {
        if (patternId == UIA_SelectionItemPatternId && SupportsGridRowSelectionPattern())
        {
            *outProvider = static_cast<ISelectionItemProvider*>(this);
            AddRef();
        }
        return S_OK;
    }

    if (_kind == AccessibilityFragmentKind::GridCell)
    {
        if (patternId == UIA_TogglePatternId && SupportsGridCellTogglePattern())
        {
            *outProvider = static_cast<IToggleProvider*>(this);
            AddRef();
        }
        else if (patternId == UIA_ValuePatternId && SupportsGridCellValuePattern())
        {
            *outProvider = static_cast<IValueProvider*>(this);
            AddRef();
        }
        else if (patternId == UIA_RangeValuePatternId && SupportsGridCellRangeValuePattern())
        {
            *outProvider = static_cast<IRangeValueProvider*>(this);
            AddRef();
        }
        else if (patternId == UIA_GridItemPatternId && SupportsGridCellPattern())
        {
            *outProvider = static_cast<IGridItemProvider*>(this);
            AddRef();
        }
        else if (patternId == UIA_TableItemPatternId && SupportsGridCellTableItemPattern())
        {
            *outProvider = static_cast<ITableItemProvider*>(this);
            AddRef();
        }
        return S_OK;
    }

    const Control* control = ResolveControl();
    if (! control)
    {
        return S_OK;
    }

    if (_kind == AccessibilityFragmentKind::TextFieldPasswordRevealButton)
    {
        if (patternId == UIA_InvokePatternId && IsTextFieldPasswordRevealButtonVisible(control))
        {
            *outProvider = static_cast<IInvokeProvider*>(this);
            AddRef();
        }
        return S_OK;
    }

    if (patternId == UIA_InvokePatternId && SupportsInvokePattern(control))
    {
        *outProvider = static_cast<IInvokeProvider*>(this);
    }
    else if (patternId == UIA_TogglePatternId && SupportsTogglePattern(control))
    {
        *outProvider = static_cast<IToggleProvider*>(this);
    }
    else if (patternId == UIA_ValuePatternId && SupportsValuePattern(control))
    {
        *outProvider = static_cast<IValueProvider*>(this);
    }
    else if (patternId == UIA_TextPatternId && SupportsTextPattern(control))
    {
        *outProvider = static_cast<ITextProvider*>(static_cast<ITextEditProvider*>(this));
    }
    else if (patternId == UIA_TextEditPatternId && SupportsTextPattern(control))
    {
        *outProvider = static_cast<ITextEditProvider*>(this);
    }
    else if (patternId == UIA_RangeValuePatternId && SupportsRangeValuePattern(control))
    {
        *outProvider = static_cast<IRangeValueProvider*>(this);
    }
    else if (patternId == UIA_SelectionPatternId && SupportsSelectionProviderPattern(control))
    {
        *outProvider = static_cast<ISelectionProvider*>(this);
    }
    else if (patternId == UIA_TablePatternId && SupportsGridTablePattern())
    {
        *outProvider = static_cast<ITableProvider*>(this);
    }
    else
    {
        return S_OK;
    }

    AddRef();
    return S_OK;
}

HRESULT AccessibilityProvider::GetPropertyValue(PROPERTYID propertyId, VARIANT* outValue) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outValue)
    {
        return E_POINTER;
    }

    VariantInit(outValue);
    const Control* root    = ResolveRootControl();
    const bool pathVisible = _kind == AccessibilityFragmentKind::Root || IsControlPathVisible(root, _path);
    const Control* control = ResolveControl();
    if (_kind == AccessibilityFragmentKind::Root)
    {
        if (! control)
        {
            switch (propertyId)
            {
                case UIA_ControlTypePropertyId: *outValue = VariantFromInt(UIA_PaneControlTypeId); return S_OK;
                case UIA_IsControlElementPropertyId:
                case UIA_IsContentElementPropertyId:
                case UIA_IsEnabledPropertyId:
                case UIA_IsKeyboardFocusablePropertyId: *outValue = VariantFromBool(true); return S_OK;
                case UIA_HasKeyboardFocusPropertyId: *outValue = VariantFromBool(false); return S_OK;
                case UIA_NamePropertyId:
                {
                    wchar_t buffer[128]{};
                    const int length = GetWindowTextW(_hwnd, buffer, static_cast<int>(std::size(buffer)));
                    return SetVariantFromString(outValue, std::wstring_view(buffer, static_cast<size_t>((std::max)(0, length))));
                }
                default: return S_OK;
            }
        }
    }

    if (_kind == AccessibilityFragmentKind::TreeItem)
    {
        TreeItemData item{};
        const Tree* tree = ResolveTreeControl();
        if (! tree || ! ResolveTreeItemData(item))
        {
            return S_OK;
        }

        switch (propertyId)
        {
            case UIA_ControlTypePropertyId: *outValue = VariantFromInt(UIA_TreeItemControlTypeId); return S_OK;
            case UIA_NamePropertyId: return SetVariantFromString(outValue, item.text);
            case UIA_IsControlElementPropertyId:
            case UIA_IsContentElementPropertyId:
            case UIA_IsKeyboardFocusablePropertyId: *outValue = VariantFromBool(pathVisible); return S_OK;
            case UIA_IsEnabledPropertyId: *outValue = VariantFromBool(pathVisible && tree->IsEnabled()); return S_OK;
            case UIA_HasKeyboardFocusPropertyId:
                *outValue = VariantFromBool(pathVisible && tree->HasFocus() && tree->GetSelectedItemId() && tree->GetSelectedItemId().value() == item.id);
                return S_OK;
            case UIA_IsOffscreenPropertyId: *outValue = VariantFromBool(! pathVisible || ! tree->IsVisible()); return S_OK;
            case UIA_SelectionItemIsSelectedPropertyId:
                *outValue = VariantFromBool(pathVisible && tree->GetSelectedItemId() && tree->GetSelectedItemId().value() == item.id);
                return S_OK;
            case UIA_LevelPropertyId: *outValue = VariantFromInt(static_cast<LONG>(item.depth + 1u)); return S_OK;
            case UIA_ExpandCollapseExpandCollapseStatePropertyId:
                if (item.hasChildren)
                {
                    *outValue = VariantFromInt(item.expanded ? ExpandCollapseState_Expanded : ExpandCollapseState_Collapsed);
                }
                return S_OK;
            default: return S_OK;
        }
    }

    if (_kind == AccessibilityFragmentKind::GridHeader)
    {
        size_t columnIndex = 0u;
        GridColumnDesc columnDesc{};
        const Grid* grid = ResolveGridControl();
        if (! grid || ! ResolveGridHeaderColumn(columnIndex, columnDesc))
        {
            return S_OK;
        }

        switch (propertyId)
        {
            case UIA_ControlTypePropertyId: *outValue = VariantFromInt(UIA_HeaderItemControlTypeId); return S_OK;
            case UIA_NamePropertyId: return SetVariantFromString(outValue, GetGridHeaderAccessibleName(columnDesc));
            case UIA_IsControlElementPropertyId:
            case UIA_IsContentElementPropertyId: *outValue = VariantFromBool(pathVisible); return S_OK;
            case UIA_IsEnabledPropertyId: *outValue = VariantFromBool(pathVisible && grid->IsEnabled()); return S_OK;
            case UIA_IsKeyboardFocusablePropertyId:
            case UIA_HasKeyboardFocusPropertyId: *outValue = VariantFromBool(false); return S_OK;
            case UIA_IsOffscreenPropertyId:
                *outValue = VariantFromBool(! pathVisible || ! grid->IsVisible() || ! grid->GetVisibleColumnHeaderRect(columnIndex).has_value());
                return S_OK;
            default: return S_OK;
        }
    }

    if (_kind == AccessibilityFragmentKind::GridRow)
    {
        size_t rowIndex   = 0u;
        const Grid* grid  = ResolveGridControl();
        const auto* model = grid ? grid->GetModel() : nullptr;
        if (! grid || ! model || ! ResolveGridRowIndex(rowIndex))
        {
            return S_OK;
        }

        switch (propertyId)
        {
            case UIA_ControlTypePropertyId: *outValue = VariantFromInt(UIA_DataItemControlTypeId); return S_OK;
            case UIA_NamePropertyId: return SetVariantFromString(outValue, BuildGridRowAccessibleName(*model, rowIndex));
            case UIA_IsControlElementPropertyId:
            case UIA_IsContentElementPropertyId:
            case UIA_IsKeyboardFocusablePropertyId: *outValue = VariantFromBool(pathVisible); return S_OK;
            case UIA_IsEnabledPropertyId: *outValue = VariantFromBool(pathVisible && grid->IsEnabled()); return S_OK;
            case UIA_HasKeyboardFocusPropertyId: *outValue = VariantFromBool(pathVisible && grid->HasFocus() && grid->IsRowSelected(rowIndex)); return S_OK;
            case UIA_IsOffscreenPropertyId:
                *outValue = VariantFromBool(! pathVisible || ! grid->IsVisible() || ! grid->GetVisibleRowRect(rowIndex).has_value());
                return S_OK;
            case UIA_SelectionItemIsSelectedPropertyId: *outValue = VariantFromBool(pathVisible && grid->IsRowSelected(rowIndex)); return S_OK;
            default: return S_OK;
        }
    }

    if (_kind == AccessibilityFragmentKind::GridCell)
    {
        size_t rowIndex    = 0u;
        size_t columnIndex = 0u;
        GridCellData cellData{};
        const Grid* grid = ResolveGridControl();
        if (! grid || ! ResolveGridCellData(rowIndex, columnIndex, cellData))
        {
            return S_OK;
        }

        switch (propertyId)
        {
            case UIA_ControlTypePropertyId: *outValue = VariantFromInt(GetGridCellControlTypeId(cellData)); return S_OK;
            case UIA_NamePropertyId: return SetVariantFromString(outValue, BuildGridCellAccessibleText(cellData));
            case UIA_HelpTextPropertyId:
            {
                const std::wstring accessibleText = BuildGridCellAccessibleText(cellData);
                if (! cellData.tooltipText.empty() && cellData.tooltipText != cellData.text && cellData.tooltipText != accessibleText)
                {
                    return SetVariantFromString(outValue, cellData.tooltipText);
                }
                return S_OK;
            }
            case UIA_IsControlElementPropertyId:
            case UIA_IsContentElementPropertyId: *outValue = VariantFromBool(pathVisible); return S_OK;
            case UIA_IsEnabledPropertyId: *outValue = VariantFromBool(pathVisible && grid->IsEnabled() && cellData.enabled); return S_OK;
            case UIA_IsKeyboardFocusablePropertyId:
            case UIA_HasKeyboardFocusPropertyId: *outValue = VariantFromBool(false); return S_OK;
            case UIA_IsOffscreenPropertyId:
                *outValue = VariantFromBool(! pathVisible || ! grid->IsVisible() || ! grid->GetVisibleCellRect(rowIndex, columnIndex).has_value());
                return S_OK;
            case UIA_ToggleToggleStatePropertyId:
                if (pathVisible && SupportsGridCellTogglePattern())
                {
                    *outValue = VariantFromInt(cellData.checked ? ToggleState_On : ToggleState_Off);
                }
                return S_OK;
            case UIA_ValueValuePropertyId:
                if (pathVisible && SupportsGridCellValuePattern())
                {
                    return SetVariantFromString(outValue, BuildGridCellAccessibleText(cellData));
                }
                return S_OK;
            case UIA_ValueIsReadOnlyPropertyId:
                if (pathVisible && (SupportsGridCellValuePattern() || SupportsGridCellRangeValuePattern()))
                {
                    *outValue = VariantFromBool(true);
                }
                return S_OK;
            case UIA_RangeValueValuePropertyId:
                if (pathVisible && SupportsGridCellRangeValuePattern())
                {
                    *outValue = VariantFromDouble(GetGridCellRangeValue(cellData));
                }
                return S_OK;
            case UIA_RangeValueMinimumPropertyId:
                if (pathVisible && SupportsGridCellRangeValuePattern())
                {
                    *outValue = VariantFromDouble(0.0);
                }
                return S_OK;
            case UIA_RangeValueMaximumPropertyId:
                if (pathVisible && SupportsGridCellRangeValuePattern())
                {
                    *outValue = VariantFromDouble(1.0);
                }
                return S_OK;
            case UIA_RangeValueLargeChangePropertyId:
            case UIA_RangeValueSmallChangePropertyId:
                if (pathVisible && SupportsGridCellRangeValuePattern())
                {
                    *outValue = VariantFromDouble(0.0);
                }
                return S_OK;
            default: return S_OK;
        }
    }

    if (_kind == AccessibilityFragmentKind::TextFieldPasswordRevealButton)
    {
        const auto* textField = dynamic_cast<const TextField*>(control);
        const bool buttonVisible = pathVisible && textField && textField->IsPasswordRevealButtonVisibleForAccessibility();
        switch (propertyId)
        {
            case UIA_ControlTypePropertyId: *outValue = VariantFromInt(UIA_ButtonControlTypeId); return S_OK;
            case UIA_NamePropertyId: return SetVariantFromString(outValue, textField ? textField->GetPasswordRevealAccessibleName() : std::wstring_view{});
            case UIA_IsControlElementPropertyId:
            case UIA_IsContentElementPropertyId: *outValue = VariantFromBool(buttonVisible); return S_OK;
            case UIA_IsEnabledPropertyId: *outValue = VariantFromBool(buttonVisible && textField->IsEnabled()); return S_OK;
            case UIA_IsKeyboardFocusablePropertyId: *outValue = VariantFromBool(buttonVisible); return S_OK;
            case UIA_HasKeyboardFocusPropertyId: *outValue = VariantFromBool(false); return S_OK;
            case UIA_IsOffscreenPropertyId: *outValue = VariantFromBool(! buttonVisible); return S_OK;
            default: return S_OK;
        }
    }

    if (! control)
    {
        return S_OK;
    }

    switch (propertyId)
    {
        case UIA_ControlTypePropertyId: *outValue = VariantFromInt(GetControlTypeId(control)); return S_OK;
        case UIA_NamePropertyId: return SetVariantFromString(outValue, GetControlAccessibleName(root, control));
        case UIA_HelpTextPropertyId: return SetVariantFromString(outValue, control->GetAccessibleHelpText());
        case UIA_IsControlElementPropertyId:
        case UIA_IsContentElementPropertyId: *outValue = VariantFromBool(pathVisible); return S_OK;
        case UIA_IsEnabledPropertyId: *outValue = VariantFromBool(pathVisible && control->IsEnabled()); return S_OK;
        case UIA_IsKeyboardFocusablePropertyId: *outValue = VariantFromBool(pathVisible && control->IsFocusable()); return S_OK;
        case UIA_HasKeyboardFocusPropertyId: *outValue = VariantFromBool(pathVisible && control->HasFocus()); return S_OK;
        case UIA_IsOffscreenPropertyId: *outValue = VariantFromBool(! pathVisible || ! control->IsVisible()); return S_OK;
        case UIA_IsPasswordPropertyId:
            *outValue = VariantFromBool(pathVisible && dynamic_cast<const TextField*>(control) != nullptr &&
                                        static_cast<const TextField*>(control)->IsMasked());
            return S_OK;
        case UIA_ValueValuePropertyId:
            if (pathVisible && SupportsValuePattern(control))
            {
                return SetVariantFromString(outValue, GetControlAccessibleValue(control));
            }
            return S_OK;
        case UIA_ValueIsReadOnlyPropertyId:
            if (pathVisible && (SupportsValuePattern(control) || SupportsRangeValuePattern(control)))
            {
                *outValue = VariantFromBool(IsValueReadOnly(control));
            }
            return S_OK;
        case UIA_RangeValueValuePropertyId:
            if (pathVisible)
            {
                if (const auto* slider = dynamic_cast<const Slider*>(control))
                {
                    *outValue = VariantFromDouble(slider->GetValue());
                }
            }
            return S_OK;
        case UIA_RangeValueMinimumPropertyId:
            if (pathVisible)
            {
                if (const auto* slider = dynamic_cast<const Slider*>(control))
                {
                    *outValue = VariantFromDouble(slider->GetMinimum());
                }
            }
            return S_OK;
        case UIA_RangeValueMaximumPropertyId:
            if (pathVisible)
            {
                if (const auto* slider = dynamic_cast<const Slider*>(control))
                {
                    *outValue = VariantFromDouble(slider->GetMaximum());
                }
            }
            return S_OK;
        case UIA_RangeValueSmallChangePropertyId:
            if (pathVisible)
            {
                if (const auto* slider = dynamic_cast<const Slider*>(control))
                {
                    *outValue = VariantFromDouble(slider->GetStep());
                }
            }
            return S_OK;
        case UIA_RangeValueLargeChangePropertyId:
            if (pathVisible)
            {
                if (const auto* slider = dynamic_cast<const Slider*>(control))
                {
                    *outValue = VariantFromDouble(slider->GetLargeStep());
                }
            }
            return S_OK;
        case UIA_GridRowCountPropertyId:
            if (const auto* grid = dynamic_cast<const Grid*>(control))
            {
                if (const auto* model = grid->GetModel())
                {
                    const size_t rowCount = model->GetRowCount();
                    *outValue             = VariantFromInt(static_cast<LONG>((std::min)(rowCount, static_cast<size_t>((std::numeric_limits<LONG>::max)()))));
                }
            }
            return S_OK;
        case UIA_GridColumnCountPropertyId:
            if (const auto* grid = dynamic_cast<const Grid*>(control))
            {
                if (const auto* model = grid->GetModel())
                {
                    const size_t columnCount = model->GetColumnCount();
                    *outValue = VariantFromInt(static_cast<LONG>((std::min)(columnCount, static_cast<size_t>((std::numeric_limits<LONG>::max)()))));
                }
            }
            return S_OK;
        default: return S_OK;
    }
}

HRESULT AccessibilityProvider::get_HostRawElementProvider(IRawElementProviderSimple** outProvider) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outProvider)
    {
        return E_POINTER;
    }

    *outProvider = nullptr;
    if (_kind != AccessibilityFragmentKind::Root)
    {
        return S_OK;
    }
    return UiaHostProviderFromHwnd(_hwnd, outProvider);
}

HRESULT AccessibilityProvider::Navigate(NavigateDirection direction, IRawElementProviderFragment** outProvider) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outProvider)
    {
        return E_POINTER;
    }

    *outProvider        = nullptr;
    const Control* root = ResolveRootControl();
    if (! root)
    {
        return S_OK;
    }

    ControlPath resolvedPath{};
    switch (direction)
    {
        case NavigateDirection_FirstChild:
            if (_kind == AccessibilityFragmentKind::Root && FindFirstSemanticControl(root, ControlPath{}, resolvedPath))
            {
                *outProvider = CreateChildProvider(resolvedPath);
            }
            else if (_kind == AccessibilityFragmentKind::Control)
            {
                if (const auto* tree = dynamic_cast<const Tree*>(ResolveControl()))
                {
                    if (const auto* model = tree->GetModel(); model && model->GetVisibleItemCount() > 0u)
                    {
                        *outProvider = CreateTreeItemProvider(_path, 0u);
                    }
                }
                else if (const auto* grid = dynamic_cast<const Grid*>(ResolveControl()))
                {
                    if (const std::optional<size_t> columnIndex = grid->GetVisibleColumnAt(0u))
                    {
                        *outProvider = CreateGridHeaderProvider(_path, columnIndex.value());
                    }
                    else if (const std::optional<size_t> rowIndex = grid->GetVisibleRowAt(0u))
                    {
                        if (const auto* model = grid->GetModel())
                        {
                            *outProvider = CreateGridRowProvider(_path, model->GetStableRowId(rowIndex.value()));
                        }
                    }
                }
                else if (IsTextFieldPasswordRevealButtonVisible(ResolveControl()))
                {
                    *outProvider = CreateTextFieldPasswordRevealButtonProvider(_path);
                }
            }
            else if (_kind == AccessibilityFragmentKind::GridRow)
            {
                size_t rowIndex  = 0u;
                const Grid* grid = ResolveGridControl();
                if (grid && ResolveGridRowIndex(rowIndex))
                {
                    if (const std::optional<size_t> columnIndex = grid->GetVisibleColumnAt(0u))
                    {
                        *outProvider = CreateGridCellProvider(_path, _gridRowId, columnIndex.value());
                    }
                }
            }
            return S_OK;
        case NavigateDirection_LastChild:
            if (_kind == AccessibilityFragmentKind::Root && FindLastSemanticControl(root, ControlPath{}, resolvedPath))
            {
                *outProvider = CreateChildProvider(resolvedPath);
            }
            else if (_kind == AccessibilityFragmentKind::Control)
            {
                if (const auto* tree = dynamic_cast<const Tree*>(ResolveControl()))
                {
                    if (const auto* model = tree->GetModel(); model && model->GetVisibleItemCount() > 0u)
                    {
                        *outProvider = CreateTreeItemProvider(_path, model->GetVisibleItemCount() - 1u);
                    }
                }
                else if (const auto* grid = dynamic_cast<const Grid*>(ResolveControl()))
                {
                    const size_t visibleRowCount = grid->GetVisibleRowCount();
                    if (visibleRowCount > 0u)
                    {
                        if (const std::optional<size_t> rowIndex = grid->GetVisibleRowAt(visibleRowCount - 1u))
                        {
                            if (const auto* model = grid->GetModel())
                            {
                                *outProvider = CreateGridRowProvider(_path, model->GetStableRowId(rowIndex.value()));
                            }
                        }
                    }
                    else
                    {
                        const size_t visibleColumnCount = grid->GetVisibleColumnCount();
                        if (visibleColumnCount > 0u)
                        {
                            if (const std::optional<size_t> columnIndex = grid->GetVisibleColumnAt(visibleColumnCount - 1u))
                            {
                                *outProvider = CreateGridHeaderProvider(_path, columnIndex.value());
                            }
                        }
                    }
                }
                else if (IsTextFieldPasswordRevealButtonVisible(ResolveControl()))
                {
                    *outProvider = CreateTextFieldPasswordRevealButtonProvider(_path);
                }
            }
            else if (_kind == AccessibilityFragmentKind::GridHeader)
            {
                return S_OK;
            }
            else if (_kind == AccessibilityFragmentKind::GridRow)
            {
                size_t rowIndex  = 0u;
                const Grid* grid = ResolveGridControl();
                if (grid && ResolveGridRowIndex(rowIndex))
                {
                    const size_t visibleColumnCount = grid->GetVisibleColumnCount();
                    if (visibleColumnCount > 0u)
                    {
                        if (const std::optional<size_t> columnIndex = grid->GetVisibleColumnAt(visibleColumnCount - 1u))
                        {
                            *outProvider = CreateGridCellProvider(_path, _gridRowId, columnIndex.value());
                        }
                    }
                }
            }
            return S_OK;
        case NavigateDirection_Parent:
            if (_kind == AccessibilityFragmentKind::Control)
            {
                WindowHostAccessibilityTarget* target = AddRefTarget();
                auto* rootProvider                    = new (std::nothrow) AccessibilityProvider(target, _hwnd);
                if (! rootProvider && target)
                {
                    static_cast<void>(target->Release());
                }
                *outProvider = rootProvider ? static_cast<IRawElementProviderFragment*>(rootProvider) : nullptr;
            }
            else if (_kind == AccessibilityFragmentKind::TextFieldPasswordRevealButton)
            {
                *outProvider = CreateChildProvider(_path);
            }
            else if (_kind == AccessibilityFragmentKind::TreeItem)
            {
                *outProvider = CreateChildProvider(_path);
            }
            else if (_kind == AccessibilityFragmentKind::GridHeader)
            {
                *outProvider = CreateChildProvider(_path);
            }
            else if (_kind == AccessibilityFragmentKind::GridRow)
            {
                *outProvider = CreateChildProvider(_path);
            }
            else if (_kind == AccessibilityFragmentKind::GridCell)
            {
                *outProvider = CreateGridRowProvider(_path, _gridRowId);
            }
            return S_OK;
        case NavigateDirection_NextSibling:
            if (_kind == AccessibilityFragmentKind::Control && FindNextSemanticControl(root, _path, resolvedPath))
            {
                *outProvider = CreateChildProvider(resolvedPath);
            }
            else if (_kind == AccessibilityFragmentKind::TreeItem)
            {
                if (const Tree* tree = ResolveTreeControl())
                {
                    if (const auto* model = tree->GetModel(); model && (_treeVisibleIndex + 1u) < model->GetVisibleItemCount())
                    {
                        *outProvider = CreateTreeItemProvider(_path, _treeVisibleIndex + 1u);
                    }
                }
            }
            else if (_kind == AccessibilityFragmentKind::GridHeader)
            {
                const Grid* grid = ResolveGridControl();
                if (grid)
                {
                    if (const std::optional<size_t> visibleOrdinal = grid->FindVisibleColumnOrdinal(_gridColumnIndex))
                    {
                        if (const std::optional<size_t> nextColumnIndex = grid->GetVisibleColumnAt(visibleOrdinal.value() + 1u))
                        {
                            *outProvider = CreateGridHeaderProvider(_path, nextColumnIndex.value());
                        }
                        else if (const std::optional<size_t> firstRowIndex = grid->GetVisibleRowAt(0u))
                        {
                            if (const auto* model = grid->GetModel())
                            {
                                *outProvider = CreateGridRowProvider(_path, model->GetStableRowId(firstRowIndex.value()));
                            }
                        }
                    }
                }
            }
            else if (_kind == AccessibilityFragmentKind::GridRow)
            {
                size_t rowIndex  = 0u;
                const Grid* grid = ResolveGridControl();
                if (grid && ResolveGridRowIndex(rowIndex))
                {
                    if (const std::optional<size_t> visibleOrdinal = grid->FindVisibleRowOrdinal(rowIndex))
                    {
                        if (const std::optional<size_t> nextRowIndex = grid->GetVisibleRowAt(visibleOrdinal.value() + 1u))
                        {
                            if (const auto* model = grid->GetModel())
                            {
                                *outProvider = CreateGridRowProvider(_path, model->GetStableRowId(nextRowIndex.value()));
                            }
                        }
                    }
                }
            }
            else if (_kind == AccessibilityFragmentKind::GridCell)
            {
                size_t rowIndex    = 0u;
                size_t columnIndex = 0u;
                GridCellData cellData{};
                const Grid* grid = ResolveGridControl();
                if (grid && ResolveGridCellData(rowIndex, columnIndex, cellData))
                {
                    if (const std::optional<size_t> visibleOrdinal = grid->FindVisibleColumnOrdinal(columnIndex))
                    {
                        if (const std::optional<size_t> nextColumnIndex = grid->GetVisibleColumnAt(visibleOrdinal.value() + 1u))
                        {
                            *outProvider = CreateGridCellProvider(_path, _gridRowId, nextColumnIndex.value());
                        }
                    }
                }
            }
            return S_OK;
        case NavigateDirection_PreviousSibling:
            if (_kind == AccessibilityFragmentKind::Control && FindPreviousSemanticControl(root, _path, resolvedPath))
            {
                *outProvider = CreateChildProvider(resolvedPath);
            }
            else if (_kind == AccessibilityFragmentKind::TreeItem)
            {
                if (_treeVisibleIndex > 0u)
                {
                    *outProvider = CreateTreeItemProvider(_path, _treeVisibleIndex - 1u);
                }
            }
            else if (_kind == AccessibilityFragmentKind::GridRow)
            {
                size_t rowIndex  = 0u;
                const Grid* grid = ResolveGridControl();
                if (grid && ResolveGridRowIndex(rowIndex))
                {
                    if (const std::optional<size_t> visibleOrdinal = grid->FindVisibleRowOrdinal(rowIndex); visibleOrdinal && visibleOrdinal.value() > 0u)
                    {
                        if (const std::optional<size_t> previousRowIndex = grid->GetVisibleRowAt(visibleOrdinal.value() - 1u))
                        {
                            if (const auto* model = grid->GetModel())
                            {
                                *outProvider = CreateGridRowProvider(_path, model->GetStableRowId(previousRowIndex.value()));
                            }
                        }
                    }
                    else
                    {
                        const size_t visibleColumnCount = grid->GetVisibleColumnCount();
                        if (visibleColumnCount > 0u)
                        {
                            if (const std::optional<size_t> previousHeaderColumnIndex = grid->GetVisibleColumnAt(visibleColumnCount - 1u))
                            {
                                *outProvider = CreateGridHeaderProvider(_path, previousHeaderColumnIndex.value());
                            }
                        }
                    }
                }
            }
            else if (_kind == AccessibilityFragmentKind::GridHeader)
            {
                const Grid* grid = ResolveGridControl();
                if (grid)
                {
                    if (const std::optional<size_t> visibleOrdinal = grid->FindVisibleColumnOrdinal(_gridColumnIndex);
                        visibleOrdinal && visibleOrdinal.value() > 0u)
                    {
                        if (const std::optional<size_t> previousColumnIndex = grid->GetVisibleColumnAt(visibleOrdinal.value() - 1u))
                        {
                            *outProvider = CreateGridHeaderProvider(_path, previousColumnIndex.value());
                        }
                    }
                }
            }
            else if (_kind == AccessibilityFragmentKind::GridCell)
            {
                size_t rowIndex    = 0u;
                size_t columnIndex = 0u;
                GridCellData cellData{};
                const Grid* grid = ResolveGridControl();
                if (grid && ResolveGridCellData(rowIndex, columnIndex, cellData))
                {
                    if (const std::optional<size_t> visibleOrdinal = grid->FindVisibleColumnOrdinal(columnIndex); visibleOrdinal && visibleOrdinal.value() > 0u)
                    {
                        if (const std::optional<size_t> previousColumnIndex = grid->GetVisibleColumnAt(visibleOrdinal.value() - 1u))
                        {
                            *outProvider = CreateGridCellProvider(_path, _gridRowId, previousColumnIndex.value());
                        }
                    }
                }
            }
            return S_OK;
        default: return S_OK;
    }
}

HRESULT AccessibilityProvider::GetRuntimeId(SAFEARRAY** outRuntimeId) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (_kind == AccessibilityFragmentKind::Root)
    {
        return SetRuntimeId(outRuntimeId, _hwnd, nullptr);
    }
    if (_kind == AccessibilityFragmentKind::TreeItem)
    {
        return SetTreeItemRuntimeId(outRuntimeId, _path, _treeVisibleIndex);
    }
    if (_kind == AccessibilityFragmentKind::GridHeader)
    {
        return SetGridHeaderRuntimeId(outRuntimeId, _path, _gridColumnIndex);
    }
    if (_kind == AccessibilityFragmentKind::GridRow)
    {
        return SetGridRowRuntimeId(outRuntimeId, _path, _gridRowId);
    }
    if (_kind == AccessibilityFragmentKind::GridCell)
    {
        return SetGridCellRuntimeId(outRuntimeId, _path, _gridRowId, _gridColumnIndex);
    }
    if (_kind == AccessibilityFragmentKind::TextFieldPasswordRevealButton)
    {
        return SetTextFieldPasswordRevealButtonRuntimeId(outRuntimeId, _path);
    }
    return SetRuntimeId(outRuntimeId, _hwnd, &_path);
}

HRESULT AccessibilityProvider::get_BoundingRectangle(UiaRect* outRect) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outRect)
    {
        return E_POINTER;
    }

    *outRect = UiaRect{};
    if (_kind == AccessibilityFragmentKind::Root)
    {
        RECT windowRect{};
        if (GetWindowRect(_hwnd, &windowRect) == FALSE)
        {
            return S_OK;
        }

        outRect->left   = static_cast<double>(windowRect.left);
        outRect->top    = static_cast<double>(windowRect.top);
        outRect->width  = static_cast<double>(windowRect.right - windowRect.left);
        outRect->height = static_cast<double>(windowRect.bottom - windowRect.top);
        return S_OK;
    }

    WindowHost* host = ResolveHost();
    if (! host)
    {
        return S_OK;
    }
    if (! IsControlPathVisible(host->GetRoot(), _path))
    {
        return S_OK;
    }

    D2D1_RECT_F boundsDip = D2D1::RectF();
    if (_kind == AccessibilityFragmentKind::TextFieldPasswordRevealButton)
    {
        const auto* textField = dynamic_cast<const TextField*>(ResolveControl());
        if (! textField || ! textField->IsPasswordRevealButtonVisibleForAccessibility())
        {
            return S_OK;
        }
        boundsDip = textField->GetPasswordRevealButtonAccessibilityRect();
    }
    else if (_kind == AccessibilityFragmentKind::TreeItem)
    {
        const Tree* tree  = ResolveTreeControl();
        const auto* model = tree ? tree->GetModel() : nullptr;
        if (! tree || ! model || _treeVisibleIndex >= model->GetVisibleItemCount())
        {
            return S_OK;
        }

        boundsDip = tree->GetItemLayoutMetrics(*host, _treeVisibleIndex).rowRect;
    }
    else if (_kind == AccessibilityFragmentKind::GridHeader)
    {
        size_t columnIndex = 0u;
        GridColumnDesc columnDesc{};
        const Grid* grid = ResolveGridControl();
        if (! grid || ! ResolveGridHeaderColumn(columnIndex, columnDesc))
        {
            return S_OK;
        }

        const std::optional<D2D1_RECT_F> headerRect = grid->GetVisibleColumnHeaderRect(columnIndex);
        if (! headerRect)
        {
            return S_OK;
        }

        boundsDip = headerRect.value();
    }
    else if (_kind == AccessibilityFragmentKind::GridRow)
    {
        size_t rowIndex  = 0u;
        const Grid* grid = ResolveGridControl();
        if (! grid || ! ResolveGridRowIndex(rowIndex))
        {
            return S_OK;
        }

        const std::optional<D2D1_RECT_F> rowRect = grid->GetVisibleRowRect(rowIndex);
        if (! rowRect)
        {
            return S_OK;
        }

        boundsDip = rowRect.value();
    }
    else if (_kind == AccessibilityFragmentKind::GridCell)
    {
        size_t rowIndex    = 0u;
        size_t columnIndex = 0u;
        GridCellData cellData{};
        const Grid* grid = ResolveGridControl();
        if (! grid || ! ResolveGridCellData(rowIndex, columnIndex, cellData))
        {
            return S_OK;
        }

        const std::optional<D2D1_RECT_F> cellRect = grid->GetVisibleCellRect(rowIndex, columnIndex);
        if (! cellRect)
        {
            return S_OK;
        }

        boundsDip = cellRect.value();
    }
    else
    {
        const Control* control = ResolveControl();
        if (! control)
        {
            return S_OK;
        }
        boundsDip = control->GetHitBounds();
    }

    POINT topLeft{static_cast<LONG>(std::lround(host->DipsToPixels(boundsDip.left))), static_cast<LONG>(std::lround(host->DipsToPixels(boundsDip.top)))};
    POINT bottomRight{static_cast<LONG>(std::lround(host->DipsToPixels(boundsDip.right))),
                      static_cast<LONG>(std::lround(host->DipsToPixels(boundsDip.bottom)))};
    ClientToScreen(_hwnd, &topLeft);
    ClientToScreen(_hwnd, &bottomRight);
    outRect->left   = static_cast<double>(topLeft.x);
    outRect->top    = static_cast<double>(topLeft.y);
    outRect->width  = static_cast<double>(bottomRight.x - topLeft.x);
    outRect->height = static_cast<double>(bottomRight.y - topLeft.y);
    return S_OK;
}

HRESULT AccessibilityProvider::GetEmbeddedFragmentRoots(SAFEARRAY** outRoots) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outRoots)
    {
        return E_POINTER;
    }

    *outRoots = nullptr;
    return S_OK;
}

HRESULT AccessibilityProvider::SetFocus() noexcept
{
    if (! IsCurrentThreadWindowThread())
    {
        AccessibilityUiActionRequest request{};
        request.provider = this;
        request.kind     = AccessibilityUiActionKind::SetFocus;
        return DispatchActionToWindowThread(request);
    }

    return ExecuteSetFocusOnWindowThread();
}

HRESULT AccessibilityProvider::get_FragmentRoot(IRawElementProviderFragmentRoot** outRoot) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outRoot)
    {
        return E_POINTER;
    }

    *outRoot = CreateRootProvider();
    return S_OK;
}

HRESULT AccessibilityProvider::ElementProviderFromPoint(double x, double y, IRawElementProviderFragment** outProvider) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outProvider)
    {
        return E_POINTER;
    }

    *outProvider        = nullptr;
    WindowHost* host    = ResolveHost();
    const Control* root = ResolveRootControl();
    if (! host || ! root)
    {
        return S_OK;
    }

    POINT pointPx{static_cast<LONG>(std::lround(x)), static_cast<LONG>(std::lround(y))};
    ScreenToClient(_hwnd, &pointPx);
    const D2D1_POINT_2F pointDip = D2D1::Point2F(host->PixelsToDip(static_cast<float>(pointPx.x)), host->PixelsToDip(static_cast<float>(pointPx.y)));
    ControlPath hitPath{};
    if (FindSemanticControlAtPoint(root, ControlPath{}, pointDip, hitPath))
    {
        const Control* control = ResolveControlAtPath(const_cast<Control*>(root), hitPath);
        if (const auto* textField = dynamic_cast<const TextField*>(control);
            textField && textField->IsPasswordRevealButtonVisibleForAccessibility() &&
            PointInRect(textField->GetPasswordRevealButtonAccessibilityRect(), pointDip))
        {
            *outProvider = CreateTextFieldPasswordRevealButtonProvider(hitPath);
            return S_OK;
        }
        if (const auto* tree = dynamic_cast<const Tree*>(control))
        {
            if (const std::optional<size_t> visibleIndex = FindTreeItemAtPoint(*host, *tree, pointDip))
            {
                *outProvider = CreateTreeItemProvider(hitPath, visibleIndex.value());
                return S_OK;
            }
        }
        else if (const auto* grid = dynamic_cast<const Grid*>(control))
        {
            if (const std::optional<size_t> headerColumnIndex = grid->FindHeaderColumnAtPoint(MakePointDip(pointDip)))
            {
                *outProvider = CreateGridHeaderProvider(hitPath, headerColumnIndex.value());
                return S_OK;
            }
            if (const std::optional<std::pair<size_t, size_t>> cellIndex = grid->FindCellAtPoint(MakePointDip(pointDip)))
            {
                if (const auto* model = grid->GetModel())
                {
                    *outProvider = CreateGridCellProvider(hitPath, model->GetStableRowId(cellIndex->first), cellIndex->second);
                }
                return S_OK;
            }
            if (const std::optional<size_t> rowIndex = grid->FindRowAtPoint(MakePointDip(pointDip)))
            {
                if (const auto* model = grid->GetModel())
                {
                    *outProvider = CreateGridRowProvider(hitPath, model->GetStableRowId(rowIndex.value()));
                }
                return S_OK;
            }
        }

        *outProvider = CreateChildProvider(hitPath);
    }
    return S_OK;
}

HRESULT AccessibilityProvider::GetFocus(IRawElementProviderFragment** outProvider) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outProvider)
    {
        return E_POINTER;
    }

    *outProvider        = nullptr;
    WindowHost* host    = ResolveHost();
    const Control* root = ResolveRootControl();
    if (! host || ! root)
    {
        return S_OK;
    }

    Control* focused = host->GetFocusControl();
    if (! focused || ! IsSemanticAccessibilityControl(focused))
    {
        return S_OK;
    }

    if (const auto* tree = dynamic_cast<const Tree*>(focused))
    {
        if (const auto* model = tree->GetModel(); model && tree->GetSelectedItemId())
        {
            if (const std::optional<size_t> visibleIndex = model->FindVisibleItemById(tree->GetSelectedItemId().value()))
            {
                ControlPath treePath{};
                if (FindPathForTarget(root, ControlPath{}, focused, treePath))
                {
                    *outProvider = CreateTreeItemProvider(treePath, visibleIndex.value());
                    return S_OK;
                }
            }
        }
    }
    else if (const auto* grid = dynamic_cast<const Grid*>(focused))
    {
        if (const std::optional<size_t> selectedRow = grid->GetPrimarySelectedRow())
        {
            ControlPath gridPath{};
            if (FindPathForTarget(root, ControlPath{}, focused, gridPath))
            {
                if (const auto* model = grid->GetModel())
                {
                    *outProvider = CreateGridRowProvider(gridPath, model->GetStableRowId(selectedRow.value()));
                }
                return S_OK;
            }
        }
    }

    ControlPath focusedPath{};
    if (FindPathForTarget(root, ControlPath{}, focused, focusedPath))
    {
        *outProvider = CreateChildProvider(focusedPath);
    }
    return S_OK;
}

HRESULT AccessibilityProvider::Invoke() noexcept
{
    if (! IsCurrentThreadWindowThread())
    {
        AccessibilityUiActionRequest request{};
        request.provider = this;
        request.kind     = AccessibilityUiActionKind::Invoke;
        return DispatchActionToWindowThread(request);
    }

    return ExecuteInvokeOnWindowThread();
}

HRESULT AccessibilityProvider::Toggle() noexcept
{
    if (! IsCurrentThreadWindowThread())
    {
        AccessibilityUiActionRequest request{};
        request.provider = this;
        request.kind     = AccessibilityUiActionKind::Toggle;
        return DispatchActionToWindowThread(request);
    }

    return ExecuteToggleOnWindowThread();
}

HRESULT AccessibilityProvider::get_ToggleState(ToggleState* outState) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outState)
    {
        return E_POINTER;
    }

    if (_kind == AccessibilityFragmentKind::GridCell)
    {
        size_t rowIndex    = 0u;
        size_t columnIndex = 0u;
        GridCellData cellData{};
        if (! ResolveGridCellData(rowIndex, columnIndex, cellData) || ! GridCellSupportsTogglePattern(cellData))
        {
            return UIA_E_NOTSUPPORTED;
        }

        *outState = cellData.checked ? ToggleState_On : ToggleState_Off;
        return S_OK;
    }

    auto* toggle = dynamic_cast<RedSalamander::DxUi::Toggle*>(ResolveMutableControl());
    if (! toggle)
    {
        return UIA_E_NOTSUPPORTED;
    }

    *outState = toggle->IsChecked() ? ToggleState_On : ToggleState_Off;
    return S_OK;
}

HRESULT AccessibilityProvider::GetVisibleRanges(SAFEARRAY** outRanges) noexcept
{
    const auto startedAt = std::chrono::steady_clock::now();
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outRanges)
    {
        return E_POINTER;
    }

    *outRanges             = nullptr;
    const Control* control = ResolveControl();
    if (! control || ! SupportsTextPattern(control))
    {
        return UIA_E_NOTSUPPORTED;
    }

    wil::com_ptr_nothrow<ITextRangeProvider> range;
    range.attach(CreateTextDocumentRangeProvider(control));
    if (! range)
    {
        return E_OUTOFMEMORY;
    }

    ITextRangeProvider* rawRanges[] = {range.get()};
    const HRESULT hr                = SetTextRangeProviderArray(outRanges, rawRanges);
    Debug::Perf::Emit(L"dxui.uia.text_range_us", L"visible-ranges", Debug::Perf::ElapsedUs(startedAt), 1u, GetControlAccessibleTextRangeText(control).size(), hr);
    return hr;
}

HRESULT AccessibilityProvider::RangeFromChild(IRawElementProviderSimple* /*childElement*/, ITextRangeProvider** outRange) noexcept
{
    return get_DocumentRange(outRange);
}

HRESULT AccessibilityProvider::RangeFromPoint(UiaPoint point, ITextRangeProvider** outRange) noexcept
{
    const auto startedAt = std::chrono::steady_clock::now();
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outRange)
    {
        return E_POINTER;
    }

    *outRange              = nullptr;
    WindowHost* host       = ResolveHost();
    const Control* control = ResolveControl();
    if (! host || ! control || ! SupportsTextPattern(control))
    {
        return UIA_E_NOTSUPPORTED;
    }

    const std::wstring text = GetControlAccessibleTextRangeText(control);
    const POINT pointScreen{static_cast<LONG>(std::lround(point.x)), static_cast<LONG>(std::lround(point.y))};
    const std::optional<PointDip> pointDip = host->ScreenPointToDipPoint(pointScreen);
    if (! pointDip)
    {
        return E_INVALIDARG;
    }

    const D2D1_POINT_2F hitPoint = D2D1::Point2F(pointDip->x, pointDip->y);
    size_t caretIndex = 0u;
    if (const std::optional<size_t> hitIndex = control->TryHitTestTextInputPoint(*host, hitPoint); hitIndex.has_value())
    {
        caretIndex = (std::min)(hitIndex.value(), text.size());
    }
    else if (const auto* comboBox = dynamic_cast<const ComboBox*>(control); comboBox && comboBox->IsEditable())
    {
        caretIndex = comboBox->HitTestEditableCaretIndex(*host, hitPoint);
    }
    else
    {
        caretIndex = HitTestCaretIndexDip(
            host, text, FontRole::Body, ResolveTextPatternViewportRect(control), 0.0f, hitPoint, ResolveReadingDirection(control->GetFlowDirection()));
    }
    *outRange = CreateTextRangeProvider(caretIndex, caretIndex);
    const HRESULT hr = *outRange ? S_OK : E_OUTOFMEMORY;
    Debug::Perf::Emit(L"dxui.uia.text_range_us", L"range-from-point", Debug::Perf::ElapsedUs(startedAt), caretIndex, text.size(), hr);
    return hr;
}

HRESULT AccessibilityProvider::get_DocumentRange(ITextRangeProvider** outRange) noexcept
{
    const auto startedAt = std::chrono::steady_clock::now();
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outRange)
    {
        return E_POINTER;
    }

    *outRange              = nullptr;
    const Control* control = ResolveControl();
    if (! control || ! SupportsTextPattern(control))
    {
        return UIA_E_NOTSUPPORTED;
    }

    *outRange        = CreateTextDocumentRangeProvider(control);
    const HRESULT hr = *outRange ? S_OK : E_OUTOFMEMORY;
    Debug::Perf::Emit(
        L"dxui.uia.text_range_us", L"document-range", Debug::Perf::ElapsedUs(startedAt), GetControlAccessibleTextRangeText(control).size(), 0u, hr);
    return hr;
}

HRESULT AccessibilityProvider::get_SupportedTextSelection(SupportedTextSelection* outSupportedSelection) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outSupportedSelection)
    {
        return E_POINTER;
    }

    const Control* control = ResolveControl();
    if (! control || ! SupportsTextPattern(control))
    {
        return UIA_E_NOTSUPPORTED;
    }

    *outSupportedSelection = SupportedTextSelection_Single;
    return S_OK;
}

HRESULT AccessibilityProvider::GetActiveComposition(ITextRangeProvider** outRange) noexcept
{
    const auto startedAt = std::chrono::steady_clock::now();
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outRange)
    {
        return E_POINTER;
    }

    *outRange              = nullptr;
    const Control* control = ResolveControl();
    if (! control || ! SupportsTextPattern(control))
    {
        return UIA_E_NOTSUPPORTED;
    }

    WindowHost* host = ResolveHost();
    NativeTextInputState state{};
    if (! host || ! host->TryReadNativeTextInputState(control, state) || ! state.compositionStartIndex || ! state.compositionEndIndex)
    {
        Debug::Perf::Emit(L"dxui.uia.text_range_us", L"active-composition", Debug::Perf::ElapsedUs(startedAt), 0u, 0u, S_OK);
        return S_OK;
    }

    const TextRangeSpan range = ClampTextRangeSpan(state.compositionStartIndex.value(), state.compositionEndIndex.value(), state.text.size());
    *outRange                 = CreateTextRangeProvider(range.start, range.end);
    const HRESULT hr          = *outRange ? S_OK : E_OUTOFMEMORY;
    Debug::Perf::Emit(L"dxui.uia.text_range_us", L"active-composition", Debug::Perf::ElapsedUs(startedAt), range.end - range.start, state.text.size(), hr);
    return hr;
}

HRESULT AccessibilityProvider::GetConversionTarget(ITextRangeProvider** outRange) noexcept
{
    const auto startedAt = std::chrono::steady_clock::now();
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outRange)
    {
        return E_POINTER;
    }

    *outRange              = nullptr;
    const Control* control = ResolveControl();
    if (! control || ! SupportsTextPattern(control))
    {
        return UIA_E_NOTSUPPORTED;
    }

    WindowHost* host = ResolveHost();
    NativeTextInputState state{};
    if (! host || ! host->TryReadNativeTextInputState(control, state) || ! state.conversionTargetStartIndex || ! state.conversionTargetEndIndex)
    {
        Debug::Perf::Emit(L"dxui.uia.text_range_us", L"conversion-target", Debug::Perf::ElapsedUs(startedAt), 0u, 0u, S_OK);
        return S_OK;
    }

    const TextRangeSpan range = ClampTextRangeSpan(state.conversionTargetStartIndex.value(), state.conversionTargetEndIndex.value(), state.text.size());
    *outRange                 = CreateTextRangeProvider(range.start, range.end);
    const HRESULT hr          = *outRange ? S_OK : E_OUTOFMEMORY;
    Debug::Perf::Emit(L"dxui.uia.text_range_us", L"conversion-target", Debug::Perf::ElapsedUs(startedAt), range.end - range.start, state.text.size(), hr);
    return hr;
}

HRESULT AccessibilityProvider::SetValue(LPCWSTR value) noexcept
{
    if (! IsCurrentThreadWindowThread())
    {
        AccessibilityUiActionRequest request{};
        request.provider    = this;
        request.kind        = AccessibilityUiActionKind::SetStringValue;
        request.stringValue = value ? value : L"";
        return DispatchActionToWindowThread(request);
    }

    return ExecuteSetStringValueOnWindowThread(value);
}

HRESULT AccessibilityProvider::SetValue(double value) noexcept
{
    if (! IsCurrentThreadWindowThread())
    {
        AccessibilityUiActionRequest request{};
        request.provider    = this;
        request.kind        = AccessibilityUiActionKind::SetRangeValue;
        request.numberValue = value;
        return DispatchActionToWindowThread(request);
    }

    return ExecuteSetRangeValueOnWindowThread(value);
}

HRESULT AccessibilityProvider::get_Value(BSTR* outValue) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outValue)
    {
        return E_POINTER;
    }

    *outValue = nullptr;
    if (_kind == AccessibilityFragmentKind::GridCell)
    {
        size_t rowIndex    = 0u;
        size_t columnIndex = 0u;
        GridCellData cellData{};
        if (! ResolveGridCellData(rowIndex, columnIndex, cellData) || ! GridCellSupportsValuePattern(cellData))
        {
            return UIA_E_NOTSUPPORTED;
        }

        const std::wstring value = BuildGridCellAccessibleText(cellData);
        *outValue                = SysAllocStringLen(value.data(), static_cast<UINT>(value.size()));
        return (*outValue || value.empty()) ? S_OK : E_OUTOFMEMORY;
    }

    const Control* control = ResolveControl();
    if (! control || ! SupportsValuePattern(control))
    {
        return UIA_E_NOTSUPPORTED;
    }

    const std::wstring_view value = GetControlAccessibleValue(control);
    *outValue                     = SysAllocStringLen(value.data(), static_cast<UINT>(value.size()));
    return (*outValue || value.empty()) ? S_OK : E_OUTOFMEMORY;
}

HRESULT AccessibilityProvider::get_Value(double* outValue) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outValue)
    {
        return E_POINTER;
    }

    size_t rowIndex    = 0u;
    size_t columnIndex = 0u;
    GridCellData cellData{};
    if (! ResolveGridCellData(rowIndex, columnIndex, cellData) || ! GridCellSupportsRangeValuePattern(cellData))
    {
        const Control* control = ResolveControl();
        if (const auto* slider = dynamic_cast<const Slider*>(control))
        {
            *outValue = slider->GetValue();
            return S_OK;
        }
        return UIA_E_NOTSUPPORTED;
    }

    *outValue = GetGridCellRangeValue(cellData);
    return S_OK;
}

HRESULT AccessibilityProvider::get_IsReadOnly(BOOL* outReadOnly) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outReadOnly)
    {
        return E_POINTER;
    }

    if (_kind == AccessibilityFragmentKind::GridCell)
    {
        size_t rowIndex    = 0u;
        size_t columnIndex = 0u;
        GridCellData cellData{};
        if (! ResolveGridCellData(rowIndex, columnIndex, cellData) ||
            (! GridCellSupportsValuePattern(cellData) && ! GridCellSupportsRangeValuePattern(cellData)))
        {
            return UIA_E_NOTSUPPORTED;
        }

        *outReadOnly = TRUE;
        return S_OK;
    }

    const Control* control = ResolveControl();
    if (! control || (! SupportsValuePattern(control) && ! SupportsRangeValuePattern(control)))
    {
        return UIA_E_NOTSUPPORTED;
    }

    *outReadOnly = IsValueReadOnly(control) ? TRUE : FALSE;
    return S_OK;
}

HRESULT AccessibilityProvider::get_Maximum(double* outMaximum) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outMaximum)
    {
        return E_POINTER;
    }

    if (SupportsGridCellRangeValuePattern())
    {
        *outMaximum = 1.0;
        return S_OK;
    }

    if (const auto* slider = dynamic_cast<const Slider*>(ResolveControl()))
    {
        *outMaximum = slider->GetMaximum();
        return S_OK;
    }

    return UIA_E_NOTSUPPORTED;
}

HRESULT AccessibilityProvider::get_Minimum(double* outMinimum) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outMinimum)
    {
        return E_POINTER;
    }

    if (SupportsGridCellRangeValuePattern())
    {
        *outMinimum = 0.0;
        return S_OK;
    }

    if (const auto* slider = dynamic_cast<const Slider*>(ResolveControl()))
    {
        *outMinimum = slider->GetMinimum();
        return S_OK;
    }

    return UIA_E_NOTSUPPORTED;
}

HRESULT AccessibilityProvider::get_LargeChange(double* outLargeChange) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outLargeChange)
    {
        return E_POINTER;
    }

    if (SupportsGridCellRangeValuePattern())
    {
        *outLargeChange = 0.0;
        return S_OK;
    }

    if (const auto* slider = dynamic_cast<const Slider*>(ResolveControl()))
    {
        *outLargeChange = slider->GetLargeStep();
        return S_OK;
    }

    return UIA_E_NOTSUPPORTED;
}

HRESULT AccessibilityProvider::get_SmallChange(double* outSmallChange) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outSmallChange)
    {
        return E_POINTER;
    }

    if (SupportsGridCellRangeValuePattern())
    {
        *outSmallChange = 0.0;
        return S_OK;
    }

    if (const auto* slider = dynamic_cast<const Slider*>(ResolveControl()))
    {
        *outSmallChange = slider->GetStep();
        return S_OK;
    }

    return UIA_E_NOTSUPPORTED;
}

HRESULT AccessibilityProvider::GetSelection(SAFEARRAY** outSelection) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outSelection)
    {
        return E_POINTER;
    }

    *outSelection = nullptr;
    ControlPath controlPath{};
    if (! ResolveControlPath(controlPath))
    {
        return UIA_E_NOTSUPPORTED;
    }
    const Control* control = ResolveControl();
    if (control && SupportsTextPattern(control))
    {
        const std::wstring text = GetControlAccessibleTextRangeText(control);
        const TextRangeSpan rangeSpan = GetControlAccessibleSelectionRange(ResolveHost(), control, text.size());
        wil::com_ptr_nothrow<ITextRangeProvider> range;
        range.attach(CreateTextRangeProvider(rangeSpan.start, rangeSpan.end));
        if (! range)
        {
            return E_OUTOFMEMORY;
        }

        ITextRangeProvider* rawRanges[] = {range.get()};
        return SetTextRangeProviderArray(outSelection, rawRanges);
    }

    if (! control || ! SupportsSelectionProviderPattern(control))
    {
        return UIA_E_NOTSUPPORTED;
    }

    std::vector<wil::com_ptr_nothrow<IRawElementProviderSimple>> selectionProviders;
    if (const auto* tree = dynamic_cast<const Tree*>(control))
    {
        if (const auto* model = tree->GetModel(); model && tree->GetSelectedItemId())
        {
            if (const std::optional<size_t> visibleIndex = model->FindVisibleItemById(tree->GetSelectedItemId().value()))
            {
                wil::com_ptr_nothrow<IRawElementProviderFragment> fragment;
                fragment.attach(CreateTreeItemProvider(controlPath, visibleIndex.value()));
                wil::com_ptr_nothrow<IRawElementProviderSimple> simple;
                if (! fragment || FAILED(fragment.query_to(simple.put())))
                {
                    return E_OUTOFMEMORY;
                }
                selectionProviders.push_back(std::move(simple));
            }
        }
    }
    else if (const auto* grid = dynamic_cast<const Grid*>(control))
    {
        const auto selection = grid->GetSelectionModel().GetOrderedSelection();
        selectionProviders.reserve(selection.size());
        for (const uint64_t rowId : selection)
        {
            if (const auto* model = grid->GetModel())
            {
                const std::optional<size_t> rowIndex = model->FindRowByStableId(rowId);
                if (! rowIndex)
                {
                    continue;
                }

                wil::com_ptr_nothrow<IRawElementProviderFragment> fragment;
                fragment.attach(CreateGridRowProvider(controlPath, rowId));
                wil::com_ptr_nothrow<IRawElementProviderSimple> simple;
                if (! fragment || FAILED(fragment.query_to(simple.put())))
                {
                    return E_OUTOFMEMORY;
                }
                selectionProviders.push_back(std::move(simple));
            }
        }
    }

    std::vector<IRawElementProviderSimple*> rawProviders;
    rawProviders.reserve(selectionProviders.size());
    for (const auto& provider : selectionProviders)
    {
        rawProviders.push_back(provider.get());
    }

    return SetProviderArray(outSelection, rawProviders);
}

HRESULT AccessibilityProvider::get_CanSelectMultiple(BOOL* outCanSelectMultiple) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outCanSelectMultiple)
    {
        return E_POINTER;
    }

    const Control* control = ResolveControl();
    if (! control || ! SupportsSelectionProviderPattern(control))
    {
        return UIA_E_NOTSUPPORTED;
    }

    if (const auto* tree = dynamic_cast<const Tree*>(control))
    {
        *outCanSelectMultiple = FALSE;
        return S_OK;
    }
    if (const auto* grid = dynamic_cast<const Grid*>(control))
    {
        *outCanSelectMultiple = grid->GetSelectionMode() == GridSelectionMode::Single ? FALSE : TRUE;
        return S_OK;
    }

    return UIA_E_NOTSUPPORTED;
}

HRESULT AccessibilityProvider::get_IsSelectionRequired(BOOL* outIsSelectionRequired) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outIsSelectionRequired)
    {
        return E_POINTER;
    }

    const Control* control = ResolveControl();
    if (! control || ! SupportsSelectionProviderPattern(control))
    {
        return UIA_E_NOTSUPPORTED;
    }

    *outIsSelectionRequired = FALSE;
    return S_OK;
}

HRESULT AccessibilityProvider::GetRowHeaders(SAFEARRAY** outRowHeaders) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outRowHeaders)
    {
        return E_POINTER;
    }

    *outRowHeaders = nullptr;
    ControlPath controlPath{};
    if (! ResolveControlPath(controlPath) || ! SupportsGridTablePattern())
    {
        return UIA_E_NOTSUPPORTED;
    }

    return SetProviderArray(outRowHeaders, {});
}

HRESULT AccessibilityProvider::GetColumnHeaders(SAFEARRAY** outColumnHeaders) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outColumnHeaders)
    {
        return E_POINTER;
    }

    *outColumnHeaders = nullptr;
    ControlPath controlPath{};
    if (! ResolveControlPath(controlPath) || ! SupportsGridTablePattern())
    {
        return UIA_E_NOTSUPPORTED;
    }

    const Grid* grid = ResolveGridControl();
    if (! grid)
    {
        return UIA_E_NOTSUPPORTED;
    }

    std::vector<wil::com_ptr_nothrow<IRawElementProviderSimple>> headerProviders;
    const size_t visibleColumnCount = grid->GetVisibleColumnCount();
    headerProviders.reserve(visibleColumnCount);
    for (size_t visibleOrdinal = 0u; visibleOrdinal < visibleColumnCount; ++visibleOrdinal)
    {
        const std::optional<size_t> columnIndex = grid->GetVisibleColumnAt(visibleOrdinal);
        if (! columnIndex)
        {
            continue;
        }

        wil::com_ptr_nothrow<IRawElementProviderFragment> fragment;
        fragment.attach(CreateGridHeaderProvider(controlPath, columnIndex.value()));
        wil::com_ptr_nothrow<IRawElementProviderSimple> simple;
        if (! fragment || FAILED(fragment.query_to(simple.put())))
        {
            return E_OUTOFMEMORY;
        }

        headerProviders.push_back(std::move(simple));
    }

    std::vector<IRawElementProviderSimple*> rawProviders;
    rawProviders.reserve(headerProviders.size());
    for (const auto& provider : headerProviders)
    {
        rawProviders.push_back(provider.get());
    }

    return SetProviderArray(outColumnHeaders, rawProviders);
}

HRESULT AccessibilityProvider::get_RowOrColumnMajor(RowOrColumnMajor* outRowOrColumnMajor) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outRowOrColumnMajor)
    {
        return E_POINTER;
    }

    ControlPath controlPath{};
    if (! ResolveControlPath(controlPath) || ! SupportsGridTablePattern())
    {
        return UIA_E_NOTSUPPORTED;
    }

    *outRowOrColumnMajor = RowOrColumnMajor_Indeterminate;
    return S_OK;
}

HRESULT AccessibilityProvider::Select() noexcept
{
    if (! IsCurrentThreadWindowThread())
    {
        AccessibilityUiActionRequest request{};
        request.provider = this;
        request.kind     = AccessibilityUiActionKind::Select;
        return DispatchActionToWindowThread(request);
    }

    return ExecuteSelectOnWindowThread();
}

HRESULT AccessibilityProvider::AddToSelection() noexcept
{
    if (! IsCurrentThreadWindowThread())
    {
        AccessibilityUiActionRequest request{};
        request.provider = this;
        request.kind     = AccessibilityUiActionKind::AddToSelection;
        return DispatchActionToWindowThread(request);
    }

    return ExecuteAddToSelectionOnWindowThread();
}

HRESULT AccessibilityProvider::RemoveFromSelection() noexcept
{
    if (! IsCurrentThreadWindowThread())
    {
        AccessibilityUiActionRequest request{};
        request.provider = this;
        request.kind     = AccessibilityUiActionKind::RemoveFromSelection;
        return DispatchActionToWindowThread(request);
    }

    return ExecuteRemoveFromSelectionOnWindowThread();
}

HRESULT AccessibilityProvider::get_IsSelected(BOOL* outSelected) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outSelected)
    {
        return E_POINTER;
    }

    if (_kind == AccessibilityFragmentKind::TreeItem)
    {
        TreeItemData item;
        const Tree* tree = ResolveTreeControl();
        if (! tree || ! ResolveTreeItemData(item))
        {
            return UIA_E_NOTSUPPORTED;
        }

        *outSelected = (tree->GetSelectedItemId() && tree->GetSelectedItemId().value() == item.id) ? TRUE : FALSE;
        return S_OK;
    }

    if (_kind == AccessibilityFragmentKind::GridRow)
    {
        size_t rowIndex  = 0u;
        const Grid* grid = ResolveGridControl();
        if (! grid || ! ResolveGridRowIndex(rowIndex))
        {
            return UIA_E_NOTSUPPORTED;
        }

        *outSelected = grid->IsRowSelected(rowIndex) ? TRUE : FALSE;
        return S_OK;
    }

    return UIA_E_NOTSUPPORTED;
}

HRESULT AccessibilityProvider::get_SelectionContainer(IRawElementProviderSimple** outContainer) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outContainer)
    {
        return E_POINTER;
    }

    *outContainer = nullptr;
    if ((_kind == AccessibilityFragmentKind::TreeItem && ! SupportsTreeItemSelectionPattern()) ||
        (_kind == AccessibilityFragmentKind::GridRow && ! SupportsGridRowSelectionPattern()) ||
        (_kind != AccessibilityFragmentKind::TreeItem && _kind != AccessibilityFragmentKind::GridRow))
    {
        return UIA_E_NOTSUPPORTED;
    }

    WindowHostAccessibilityTarget* target = AddRefTarget();
    auto* provider                        = new (std::nothrow) AccessibilityProvider(target, _hwnd, _path);
    if (! provider)
    {
        if (target)
        {
            static_cast<void>(target->Release());
        }
        return E_OUTOFMEMORY;
    }

    *outContainer = static_cast<IRawElementProviderSimple*>(provider);
    return S_OK;
}

HRESULT AccessibilityProvider::Expand() noexcept
{
    if (! IsCurrentThreadWindowThread())
    {
        AccessibilityUiActionRequest request{};
        request.provider = this;
        request.kind     = AccessibilityUiActionKind::Expand;
        return DispatchActionToWindowThread(request);
    }

    return ExecuteExpandOnWindowThread(true);
}

HRESULT AccessibilityProvider::Collapse() noexcept
{
    if (! IsCurrentThreadWindowThread())
    {
        AccessibilityUiActionRequest request{};
        request.provider = this;
        request.kind     = AccessibilityUiActionKind::Collapse;
        return DispatchActionToWindowThread(request);
    }

    return ExecuteExpandOnWindowThread(false);
}

HRESULT AccessibilityProvider::get_ExpandCollapseState(ExpandCollapseState* outState) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outState)
    {
        return E_POINTER;
    }

    TreeItemData item;
    if (! ResolveTreeItemData(item) || ! item.hasChildren)
    {
        return UIA_E_NOTSUPPORTED;
    }

    *outState = item.expanded ? ExpandCollapseState_Expanded : ExpandCollapseState_Collapsed;
    return S_OK;
}

HRESULT AccessibilityProvider::get_Row(int* outRow) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outRow)
    {
        return E_POINTER;
    }

    size_t rowIndex    = 0u;
    size_t columnIndex = 0u;
    GridCellData cellData{};
    if (! ResolveGridCellData(rowIndex, columnIndex, cellData) || rowIndex > static_cast<size_t>((std::numeric_limits<int>::max)()))
    {
        return UIA_E_NOTSUPPORTED;
    }

    *outRow = static_cast<int>(rowIndex);
    return S_OK;
}

HRESULT AccessibilityProvider::get_Column(int* outColumn) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outColumn)
    {
        return E_POINTER;
    }

    size_t rowIndex    = 0u;
    size_t columnIndex = 0u;
    GridCellData cellData{};
    if (! ResolveGridCellData(rowIndex, columnIndex, cellData) || columnIndex > static_cast<size_t>((std::numeric_limits<int>::max)()))
    {
        return UIA_E_NOTSUPPORTED;
    }

    *outColumn = static_cast<int>(columnIndex);
    return S_OK;
}

HRESULT AccessibilityProvider::get_RowSpan(int* outRowSpan) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outRowSpan)
    {
        return E_POINTER;
    }

    if (! SupportsGridCellPattern())
    {
        return UIA_E_NOTSUPPORTED;
    }

    *outRowSpan = 1;
    return S_OK;
}

HRESULT AccessibilityProvider::get_ColumnSpan(int* outColumnSpan) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outColumnSpan)
    {
        return E_POINTER;
    }

    if (! SupportsGridCellPattern())
    {
        return UIA_E_NOTSUPPORTED;
    }

    *outColumnSpan = 1;
    return S_OK;
}

HRESULT AccessibilityProvider::get_ContainingGrid(IRawElementProviderSimple** outContainingGrid) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outContainingGrid)
    {
        return E_POINTER;
    }

    *outContainingGrid = nullptr;
    if (! SupportsGridCellPattern())
    {
        return UIA_E_NOTSUPPORTED;
    }

    WindowHostAccessibilityTarget* target = AddRefTarget();
    auto* provider                        = new (std::nothrow) AccessibilityProvider(target, _hwnd, _path);
    if (! provider)
    {
        if (target)
        {
            static_cast<void>(target->Release());
        }
        return E_OUTOFMEMORY;
    }

    *outContainingGrid = static_cast<IRawElementProviderSimple*>(provider);
    return S_OK;
}

HRESULT AccessibilityProvider::GetRowHeaderItems(SAFEARRAY** outRowHeaderItems) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outRowHeaderItems)
    {
        return E_POINTER;
    }

    *outRowHeaderItems = nullptr;
    if (! SupportsGridCellTableItemPattern())
    {
        return UIA_E_NOTSUPPORTED;
    }

    return SetProviderArray(outRowHeaderItems, {});
}

HRESULT AccessibilityProvider::GetColumnHeaderItems(SAFEARRAY** outColumnHeaderItems) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outColumnHeaderItems)
    {
        return E_POINTER;
    }

    *outColumnHeaderItems = nullptr;
    if (! SupportsGridCellTableItemPattern())
    {
        return UIA_E_NOTSUPPORTED;
    }

    size_t rowIndex    = 0u;
    size_t columnIndex = 0u;
    GridCellData cellData{};
    if (! ResolveGridCellData(rowIndex, columnIndex, cellData))
    {
        return UIA_E_NOTSUPPORTED;
    }

    wil::com_ptr_nothrow<IRawElementProviderFragment> fragment;
    fragment.attach(CreateGridHeaderProvider(_path, columnIndex));
    wil::com_ptr_nothrow<IRawElementProviderSimple> simple;
    if (! fragment || FAILED(fragment.query_to(simple.put())))
    {
        return E_OUTOFMEMORY;
    }

    IRawElementProviderSimple* provider = simple.get();
    return SetProviderArray(outColumnHeaderItems, std::span<IRawElementProviderSimple* const>(&provider, 1u));
}

WindowHost* AccessibilityProvider::ResolveHost() const noexcept
{
    return _target ? _target->ResolveHost() : nullptr;
}

const Control* AccessibilityProvider::ResolveRootControl() const noexcept
{
    WindowHost* host = ResolveHost();
    return host ? host->GetRoot() : nullptr;
}

bool AccessibilityProvider::ResolveControlPath(ControlPath& outPath) const noexcept
{
    if (_kind == AccessibilityFragmentKind::Control || _kind == AccessibilityFragmentKind::TextFieldPasswordRevealButton)
    {
        outPath = _path;
        return true;
    }

    if (_kind == AccessibilityFragmentKind::Root)
    {
        return TryResolveSingleSemanticRootControlPath(ResolveRootControl(), outPath);
    }

    return false;
}

const Control* AccessibilityProvider::ResolveControl() const noexcept
{
    ControlPath path{};
    if (! ResolveControlPath(path))
    {
        return nullptr;
    }

    WindowHost* host = ResolveHost();
    return host ? ResolveControlAtPath(host->GetRoot(), path) : nullptr;
}

Control* AccessibilityProvider::ResolveMutableControl() const noexcept
{
    ControlPath path{};
    if (! ResolveControlPath(path))
    {
        return nullptr;
    }

    WindowHost* host = ResolveHost();
    return host ? ResolveControlAtPath(host->GetRoot(), path) : nullptr;
}

const Tree* AccessibilityProvider::ResolveTreeControl() const noexcept
{
    ControlPath path{};
    if (_kind == AccessibilityFragmentKind::TreeItem)
    {
        path = _path;
    }
    else if (! ResolveControlPath(path))
    {
        return nullptr;
    }

    WindowHost* host = ResolveHost();
    return host ? dynamic_cast<const Tree*>(ResolveControlAtPath(host->GetRoot(), path)) : nullptr;
}

Tree* AccessibilityProvider::ResolveMutableTreeControl() const noexcept
{
    ControlPath path{};
    if (_kind == AccessibilityFragmentKind::TreeItem)
    {
        path = _path;
    }
    else if (! ResolveControlPath(path))
    {
        return nullptr;
    }

    WindowHost* host = ResolveHost();
    return host ? dynamic_cast<Tree*>(ResolveControlAtPath(host->GetRoot(), path)) : nullptr;
}

const Grid* AccessibilityProvider::ResolveGridControl() const noexcept
{
    ControlPath path{};
    if (_kind == AccessibilityFragmentKind::GridHeader || _kind == AccessibilityFragmentKind::GridRow || _kind == AccessibilityFragmentKind::GridCell)
    {
        path = _path;
    }
    else if (! ResolveControlPath(path))
    {
        return nullptr;
    }

    WindowHost* host = ResolveHost();
    return host ? dynamic_cast<const Grid*>(ResolveControlAtPath(host->GetRoot(), path)) : nullptr;
}

Grid* AccessibilityProvider::ResolveMutableGridControl() const noexcept
{
    ControlPath path{};
    if (_kind == AccessibilityFragmentKind::GridHeader || _kind == AccessibilityFragmentKind::GridRow || _kind == AccessibilityFragmentKind::GridCell)
    {
        path = _path;
    }
    else if (! ResolveControlPath(path))
    {
        return nullptr;
    }

    WindowHost* host = ResolveHost();
    return host ? dynamic_cast<Grid*>(ResolveControlAtPath(host->GetRoot(), path)) : nullptr;
}

bool AccessibilityProvider::ResolveTreeItemData(TreeItemData& outItem) const noexcept
{
    const Tree* tree = ResolveTreeControl();
    if (! tree)
    {
        return false;
    }

    const auto* model = tree->GetModel();
    if (! model || _treeVisibleIndex >= model->GetVisibleItemCount())
    {
        return false;
    }

    model->GetVisibleItem(_treeVisibleIndex, outItem);
    return true;
}

bool AccessibilityProvider::ResolveGridHeaderColumn(size_t& outColumnIndex, GridColumnDesc& outColumnDesc) const noexcept
{
    const Grid* grid  = ResolveGridControl();
    const auto* model = grid ? grid->GetModel() : nullptr;
    if (_kind != AccessibilityFragmentKind::GridHeader || ! grid || ! model || _gridColumnIndex >= model->GetColumnCount() ||
        ! grid->FindVisibleColumnOrdinal(_gridColumnIndex) || ! grid->GetVisibleColumnHeaderRect(_gridColumnIndex))
    {
        return false;
    }

    outColumnIndex = _gridColumnIndex;
    outColumnDesc  = model->GetColumn(_gridColumnIndex);
    return true;
}

bool AccessibilityProvider::ResolveGridRowIndex(size_t& outRowIndex) const noexcept
{
    const Grid* grid  = ResolveGridControl();
    const auto* model = grid ? grid->GetModel() : nullptr;
    if (! grid || ! model)
    {
        return false;
    }

    const std::optional<size_t> rowIndex = model->FindRowByStableId(_gridRowId);
    if (! rowIndex)
    {
        return false;
    }

    outRowIndex = rowIndex.value();
    return true;
}

bool AccessibilityProvider::ResolveGridCellData(size_t& outRowIndex, size_t& outColumnIndex, GridCellData& outCellData) const noexcept
{
    const Grid* grid  = ResolveGridControl();
    const auto* model = grid ? grid->GetModel() : nullptr;
    if (! grid || ! model || ! ResolveGridRowIndex(outRowIndex) || ! grid->FindVisibleRowOrdinal(outRowIndex) ||
        ! grid->FindVisibleColumnOrdinal(_gridColumnIndex))
    {
        return false;
    }

    outColumnIndex = _gridColumnIndex;
    outCellData    = {};
    model->GetCellData(outRowIndex, outColumnIndex, outCellData);
    return true;
}

bool AccessibilityProvider::SupportsTreeItemSelectionPattern() const noexcept
{
    TreeItemData item;
    return ResolveTreeItemData(item);
}

bool AccessibilityProvider::SupportsTreeItemExpandCollapsePattern() const noexcept
{
    TreeItemData item;
    return ResolveTreeItemData(item) && item.hasChildren;
}

bool AccessibilityProvider::SupportsGridTablePattern() const noexcept
{
    const Grid* grid  = ResolveGridControl();
    const auto* model = grid ? grid->GetModel() : nullptr;
    return grid && model && model->GetColumnCount() > 0u;
}

bool AccessibilityProvider::SupportsGridRowSelectionPattern() const noexcept
{
    size_t rowIndex = 0u;
    return ResolveGridRowIndex(rowIndex);
}

bool AccessibilityProvider::SupportsGridCellPattern() const noexcept
{
    size_t rowIndex    = 0u;
    size_t columnIndex = 0u;
    GridCellData cellData{};
    return ResolveGridCellData(rowIndex, columnIndex, cellData);
}

bool AccessibilityProvider::SupportsGridCellTableItemPattern() const noexcept
{
    const Grid* grid  = ResolveGridControl();
    const auto* model = grid ? grid->GetModel() : nullptr;
    return SupportsGridCellPattern() && grid && model && model->GetColumnCount() > 0u;
}

bool AccessibilityProvider::SupportsGridCellTogglePattern() const noexcept
{
    size_t rowIndex    = 0u;
    size_t columnIndex = 0u;
    GridCellData cellData{};
    return ResolveGridCellData(rowIndex, columnIndex, cellData) && GridCellSupportsTogglePattern(cellData);
}

bool AccessibilityProvider::SupportsGridCellValuePattern() const noexcept
{
    size_t rowIndex    = 0u;
    size_t columnIndex = 0u;
    GridCellData cellData{};
    return ResolveGridCellData(rowIndex, columnIndex, cellData) && GridCellSupportsValuePattern(cellData);
}

bool AccessibilityProvider::SupportsGridCellRangeValuePattern() const noexcept
{
    size_t rowIndex    = 0u;
    size_t columnIndex = 0u;
    GridCellData cellData{};
    return ResolveGridCellData(rowIndex, columnIndex, cellData) && GridCellSupportsRangeValuePattern(cellData);
}

std::optional<size_t> AccessibilityProvider::FindTreeItemAtPoint(WindowHost& host, const Tree& tree, D2D1_POINT_2F pointDip) const noexcept
{
    const auto* model = tree.GetModel();
    if (! model)
    {
        return std::nullopt;
    }

    for (size_t visibleIndex = 0u; visibleIndex < model->GetVisibleItemCount(); ++visibleIndex)
    {
        if (PointInRect(tree.GetItemLayoutMetrics(host, visibleIndex).rowRect, pointDip))
        {
            return visibleIndex;
        }
    }

    return std::nullopt;
}

WindowHostAccessibilityTarget* AccessibilityProvider::AddRefTarget() const noexcept
{
    if (_target)
    {
        static_cast<void>(_target->AddRef());
    }

    return _target;
}

bool AccessibilityProvider::IsCurrentThreadWindowThread() const noexcept
{
    return _hwnd && GetWindowThreadProcessId(_hwnd, nullptr) == GetCurrentThreadId();
}

HRESULT AccessibilityProvider::DispatchActionToWindowThread(AccessibilityUiActionRequest& request) noexcept
{
    if (! _hwnd || IsWindow(_hwnd) == FALSE)
    {
        return UIA_E_ELEMENTNOTAVAILABLE;
    }

    DWORD_PTR messageResult = 0;
    if (SendMessageTimeoutW(
            _hwnd, kWindowHostAccessibilityActionMessage, 0, reinterpret_cast<LPARAM>(&request), SMTO_ABORTIFHUNG | SMTO_BLOCK, 5000u, &messageResult) == 0)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    return request.result;
}

HRESULT AccessibilityProvider::ExecuteUiThreadAction(AccessibilityUiActionKind kind, LPCWSTR stringValue, double numberValue) noexcept
{
    switch (kind)
    {
        case AccessibilityUiActionKind::SetFocus: return ExecuteSetFocusOnWindowThread();
        case AccessibilityUiActionKind::Invoke: return ExecuteInvokeOnWindowThread();
        case AccessibilityUiActionKind::Toggle: return ExecuteToggleOnWindowThread();
        case AccessibilityUiActionKind::SetStringValue: return ExecuteSetStringValueOnWindowThread(stringValue);
        case AccessibilityUiActionKind::SetRangeValue: return ExecuteSetRangeValueOnWindowThread(numberValue);
        case AccessibilityUiActionKind::Select: return ExecuteSelectOnWindowThread();
        case AccessibilityUiActionKind::AddToSelection: return ExecuteAddToSelectionOnWindowThread();
        case AccessibilityUiActionKind::RemoveFromSelection: return ExecuteRemoveFromSelectionOnWindowThread();
        case AccessibilityUiActionKind::Expand: return ExecuteExpandOnWindowThread(true);
        case AccessibilityUiActionKind::Collapse: return ExecuteExpandOnWindowThread(false);
    }

    return UIA_E_NOTSUPPORTED;
}

HRESULT AccessibilityProvider::ExecuteSetFocusOnWindowThread() noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    WindowHost* host = ResolveHost();
    if (! host)
    {
        return S_OK;
    }

    if (_kind == AccessibilityFragmentKind::Root)
    {
        ::SetFocus(_hwnd);
        return S_OK;
    }

    if (_kind == AccessibilityFragmentKind::TreeItem)
    {
        TreeItemData item{};
        Tree* tree = ResolveMutableTreeControl();
        if (tree && ResolveTreeItemData(item))
        {
            ::SetFocus(_hwnd);
            tree->SetSelectedItemId(item.id);
            host->SetFocusControl(tree);
            host->Invalidate();
        }
        return S_OK;
    }

    if (_kind == AccessibilityFragmentKind::GridHeader)
    {
        Grid* grid = ResolveMutableGridControl();
        if (grid)
        {
            ::SetFocus(_hwnd);
            host->SetFocusControl(grid);
            host->Invalidate();
        }
        return S_OK;
    }

    if (_kind == AccessibilityFragmentKind::GridRow)
    {
        size_t rowIndex = 0u;
        Grid* grid      = ResolveMutableGridControl();
        if (grid && ResolveGridRowIndex(rowIndex) && grid->RequestSelectRow(rowIndex, 0u))
        {
            ::SetFocus(_hwnd);
            host->SetFocusControl(grid);
            host->Invalidate();
        }
        return S_OK;
    }

    if (_kind == AccessibilityFragmentKind::GridCell)
    {
        size_t rowIndex    = 0u;
        size_t columnIndex = 0u;
        GridCellData cellData{};
        Grid* grid = ResolveMutableGridControl();
        if (grid && ResolveGridCellData(rowIndex, columnIndex, cellData) && grid->RequestSelectRow(rowIndex, 0u))
        {
            ::SetFocus(_hwnd);
            host->SetFocusControl(grid);
            host->Invalidate();
        }
        return S_OK;
    }

    if (_kind == AccessibilityFragmentKind::TextFieldPasswordRevealButton)
    {
        auto* textField = dynamic_cast<TextField*>(ResolveMutableControl());
        if (textField && textField->IsPasswordRevealButtonVisibleForAccessibility())
        {
            ::SetFocus(_hwnd);
            host->SetFocusControl(textField);
            host->Invalidate();
        }
        return S_OK;
    }

    Control* control = ResolveMutableControl();
    if (control && control->IsFocusable())
    {
        ::SetFocus(_hwnd);
        host->SetFocusControl(control);
    }
    return S_OK;
}

HRESULT AccessibilityProvider::ExecuteInvokeOnWindowThread() noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    WindowHost* host = ResolveHost();
    if (_kind == AccessibilityFragmentKind::TextFieldPasswordRevealButton)
    {
        auto* textField = dynamic_cast<TextField*>(ResolveMutableControl());
        if (! host || ! textField || ! textField->InvokePasswordRevealButton(*host))
        {
            return UIA_E_NOTSUPPORTED;
        }
        return S_OK;
    }

    auto* button     = dynamic_cast<Button*>(ResolveMutableControl());
    if (! host || ! button || ! SupportsInvokePattern(button))
    {
        return UIA_E_NOTSUPPORTED;
    }

    return button->Invoke(*host, true) ? S_OK : UIA_E_NOTSUPPORTED;
}

HRESULT AccessibilityProvider::ExecuteToggleOnWindowThread() noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    WindowHost* host = ResolveHost();
    if (! host)
    {
        return UIA_E_NOTSUPPORTED;
    }

    if (_kind == AccessibilityFragmentKind::GridCell)
    {
        size_t rowIndex    = 0u;
        size_t columnIndex = 0u;
        GridCellData cellData{};
        Grid* grid = ResolveMutableGridControl();
        if (! grid || ! ResolveGridCellData(rowIndex, columnIndex, cellData) || ! GridCellSupportsTogglePattern(cellData))
        {
            return UIA_E_NOTSUPPORTED;
        }

        return grid->RequestToggleCheckboxCell(*host, rowIndex, columnIndex) ? S_OK : UIA_E_NOTSUPPORTED;
    }

    auto* toggle = dynamic_cast<RedSalamander::DxUi::Toggle*>(ResolveMutableControl());
    if (! toggle)
    {
        return UIA_E_NOTSUPPORTED;
    }

    static_cast<void>(toggle->OnMnemonic(*host));
    return S_OK;
}

HRESULT AccessibilityProvider::ExecuteSetStringValueOnWindowThread(LPCWSTR value) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    WindowHost* host = ResolveHost();
    Control* control = ResolveMutableControl();
    if (! host || ! control || ! SupportsValuePattern(control) || IsValueReadOnly(control))
    {
        return UIA_E_NOTSUPPORTED;
    }

    if (auto* textField = dynamic_cast<TextField*>(control))
    {
        textField->SetTextAndNotify(value ? value : L"");
        host->SyncTextInput(textField);
        host->Invalidate();
        return S_OK;
    }
    if (auto* comboBox = dynamic_cast<ComboBox*>(control))
    {
        comboBox->SetTextAndNotify(value ? value : L"");
        host->SyncTextInput(comboBox);
        host->Invalidate();
        return S_OK;
    }

    return UIA_E_NOTSUPPORTED;
}

HRESULT AccessibilityProvider::ExecuteSetRangeValueOnWindowThread(double value) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    WindowHost* host = ResolveHost();
    Control* control = ResolveMutableControl();
    if (! host || ! control || ! SupportsRangeValuePattern(control) || IsValueReadOnly(control))
    {
        return UIA_E_NOTSUPPORTED;
    }

    if (auto* slider = dynamic_cast<Slider*>(control))
    {
        slider->SetValue(value);
        host->Invalidate();
        return S_OK;
    }

    return UIA_E_NOTSUPPORTED;
}

HRESULT AccessibilityProvider::ExecuteSelectOnWindowThread() noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    WindowHost* host = ResolveHost();
    if (! host)
    {
        return UIA_E_NOTSUPPORTED;
    }

    if (_kind == AccessibilityFragmentKind::TreeItem)
    {
        Tree* tree = ResolveMutableTreeControl();
        if (! tree || ! SupportsTreeItemSelectionPattern() || ! tree->RequestSelectVisibleItem(_treeVisibleIndex))
        {
            return UIA_E_NOTSUPPORTED;
        }

        host->SetFocusControl(tree);
        host->Invalidate();
        return S_OK;
    }

    if (_kind == AccessibilityFragmentKind::GridRow)
    {
        size_t rowIndex = 0u;
        Grid* grid      = ResolveMutableGridControl();
        if (! grid || ! ResolveGridRowIndex(rowIndex) || ! grid->RequestSelectRow(rowIndex, 0u))
        {
            return UIA_E_NOTSUPPORTED;
        }

        host->SetFocusControl(grid);
        host->Invalidate();
        return S_OK;
    }

    return UIA_E_NOTSUPPORTED;
}

HRESULT AccessibilityProvider::ExecuteAddToSelectionOnWindowThread() noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (_kind == AccessibilityFragmentKind::TreeItem)
    {
        return ExecuteSelectOnWindowThread();
    }

    WindowHost* host = ResolveHost();
    size_t rowIndex  = 0u;
    Grid* grid       = ResolveMutableGridControl();
    if (! host || ! grid || _kind != AccessibilityFragmentKind::GridRow || ! ResolveGridRowIndex(rowIndex) || ! grid->RequestSelectRow(rowIndex, MK_CONTROL))
    {
        return UIA_E_NOTSUPPORTED;
    }

    host->SetFocusControl(grid);
    host->Invalidate();
    return S_OK;
}

HRESULT AccessibilityProvider::ExecuteRemoveFromSelectionOnWindowThread() noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    WindowHost* host = ResolveHost();
    if (! host)
    {
        return UIA_E_NOTSUPPORTED;
    }

    if (_kind == AccessibilityFragmentKind::TreeItem)
    {
        Tree* tree = ResolveMutableTreeControl();
        TreeItemData item;
        if (! tree || ! ResolveTreeItemData(item))
        {
            return UIA_E_NOTSUPPORTED;
        }

        if (tree->GetSelectedItemId() && tree->GetSelectedItemId().value() == item.id)
        {
            tree->SetSelectedItemId(std::nullopt);
            host->Invalidate();
        }

        return S_OK;
    }

    if (_kind == AccessibilityFragmentKind::GridRow)
    {
        size_t rowIndex = 0u;
        Grid* grid      = ResolveMutableGridControl();
        if (! grid || ! ResolveGridRowIndex(rowIndex) || ! grid->RequestRemoveRowSelection(rowIndex))
        {
            return UIA_E_NOTSUPPORTED;
        }

        host->Invalidate();
        return S_OK;
    }

    return UIA_E_NOTSUPPORTED;
}

HRESULT AccessibilityProvider::ExecuteExpandOnWindowThread(bool expanded) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    WindowHost* host = ResolveHost();
    Tree* tree       = ResolveMutableTreeControl();
    if (! host || ! tree || ! SupportsTreeItemExpandCollapsePattern())
    {
        return UIA_E_NOTSUPPORTED;
    }

    if (! tree->RequestExpandedState(_treeVisibleIndex, expanded))
    {
        return UIA_E_NOTSUPPORTED;
    }

    host->Invalidate();
    return S_OK;
}

IRawElementProviderFragmentRoot* AccessibilityProvider::CreateRootProvider() noexcept
{
    WindowHostAccessibilityTarget* target = AddRefTarget();
    auto* provider                        = new (std::nothrow) AccessibilityProvider(target, _hwnd);
    if (! provider && target)
    {
        static_cast<void>(target->Release());
    }
    return provider ? static_cast<IRawElementProviderFragmentRoot*>(provider) : nullptr;
}

IRawElementProviderFragment* AccessibilityProvider::CreateChildProvider(const ControlPath& path) noexcept
{
    WindowHostAccessibilityTarget* target = AddRefTarget();
    auto* provider                        = new (std::nothrow) AccessibilityProvider(target, _hwnd, path);
    if (! provider && target)
    {
        static_cast<void>(target->Release());
    }
    return provider ? static_cast<IRawElementProviderFragment*>(provider) : nullptr;
}

IRawElementProviderFragment* AccessibilityProvider::CreateTextFieldPasswordRevealButtonProvider(const ControlPath& path) noexcept
{
    WindowHostAccessibilityTarget* target = AddRefTarget();
    auto* provider = new (std::nothrow) AccessibilityProvider(target, _hwnd, path, AccessibilityFragmentKind::TextFieldPasswordRevealButton);
    if (! provider && target)
    {
        static_cast<void>(target->Release());
    }
    return provider ? static_cast<IRawElementProviderFragment*>(provider) : nullptr;
}

IRawElementProviderFragment* AccessibilityProvider::CreateTreeItemProvider(const ControlPath& path, size_t visibleIndex) noexcept
{
    WindowHostAccessibilityTarget* target = AddRefTarget();
    auto* provider                        = new (std::nothrow) AccessibilityProvider(target, _hwnd, path, visibleIndex);
    if (! provider && target)
    {
        static_cast<void>(target->Release());
    }
    return provider ? static_cast<IRawElementProviderFragment*>(provider) : nullptr;
}

IRawElementProviderFragment* AccessibilityProvider::CreateGridHeaderProvider(const ControlPath& path, size_t columnIndex) noexcept
{
    WindowHostAccessibilityTarget* target = AddRefTarget();
    auto* provider                        = new (std::nothrow) AccessibilityProvider(target, _hwnd, path, columnIndex, AccessibilityProvider::GridHeaderTag{});
    if (! provider && target)
    {
        static_cast<void>(target->Release());
    }
    return provider ? static_cast<IRawElementProviderFragment*>(provider) : nullptr;
}

IRawElementProviderFragment* AccessibilityProvider::CreateGridRowProvider(const ControlPath& path, uint64_t rowId) noexcept
{
    WindowHostAccessibilityTarget* target = AddRefTarget();
    auto* provider                        = new (std::nothrow) AccessibilityProvider(target, _hwnd, path, rowId, AccessibilityFragmentKind::GridRow);
    if (! provider && target)
    {
        static_cast<void>(target->Release());
    }
    return provider ? static_cast<IRawElementProviderFragment*>(provider) : nullptr;
}

IRawElementProviderFragment* AccessibilityProvider::CreateGridCellProvider(const ControlPath& path, uint64_t rowId, size_t columnIndex) noexcept
{
    WindowHostAccessibilityTarget* target = AddRefTarget();
    auto* provider                        = new (std::nothrow) AccessibilityProvider(target, _hwnd, path, rowId, columnIndex);
    if (! provider && target)
    {
        static_cast<void>(target->Release());
    }
    return provider ? static_cast<IRawElementProviderFragment*>(provider) : nullptr;
}

ITextRangeProvider* AccessibilityProvider::CreateTextRangeProvider(size_t start, size_t end) noexcept
{
    ControlPath controlPath{};
    if (! ResolveControlPath(controlPath))
    {
        return nullptr;
    }

    WindowHostAccessibilityTarget* target = AddRefTarget();
    auto* provider                        = new (std::nothrow) AccessibilityTextRangeProvider(target, _hwnd, controlPath, start, end);
    if (! provider && target)
    {
        static_cast<void>(target->Release());
    }
    return provider ? static_cast<ITextRangeProvider*>(provider) : nullptr;
}

ITextRangeProvider* AccessibilityProvider::CreateTextDocumentRangeProvider(const Control* control) noexcept
{
    if (! control || ! SupportsTextPattern(control))
    {
        return nullptr;
    }

    const std::wstring text = GetControlAccessibleTextRangeText(control);
    return CreateTextRangeProvider(0u, text.size());
}

[[nodiscard]] bool FindAccessibilityPathForTarget(const Control* current, const ControlPath& basePath, const Control* target, ControlPath& outPath) noexcept
{
    if (! current || ! current->IsVisible())
    {
        return false;
    }

    if (current == target && IsSemanticAccessibilityControl(current))
    {
        outPath = basePath;
        return true;
    }

    const auto* panel = dynamic_cast<const Panel*>(current);
    if (! panel)
    {
        return false;
    }

    const auto children = panel->GetChildren();
    for (size_t index = 0u; index < children.size(); ++index)
    {
        if (! children[index])
        {
            continue;
        }

        ControlPath childPath{};
        if (! TryAppendPathIndex(basePath, index, childPath))
        {
            continue;
        }

        if (FindAccessibilityPathForTarget(children[index].get(), childPath, target, outPath))
        {
            return true;
        }
    }

    return false;
}

bool AccessibilityProvider::FindPathForTarget(const Control* current, const ControlPath& basePath, const Control* target, ControlPath& outPath) noexcept
{
    return FindAccessibilityPathForTarget(current, basePath, target, outPath);
}

} // namespace

bool RaiseWindowHostTextInputAutomationEvent(HWND hwnd, const Control* control, TextInputAutomationEventKind kind) noexcept
{
    if (! hwnd || ! control)
    {
        return false;
    }

    WindowHostAccessibilityTarget* target = AcquireWindowHostAccessibilityTarget(hwnd);
    if (! target)
    {
        return false;
    }

    WindowHost* const host = target->ResolveHost();
    ControlPath controlPath{};
    if (! host || ! FindAccessibilityPathForTarget(host->GetRoot(), ControlPath{}, control, controlPath))
    {
        static_cast<void>(target->Release());
        return false;
    }

    auto* provider = new (std::nothrow) AccessibilityProvider(target, hwnd, controlPath);
    if (! provider)
    {
        static_cast<void>(target->Release());
        return false;
    }
    const auto releaseProvider = wil::scope_exit([&] { static_cast<void>(provider->Release()); });

    switch (kind)
    {
        case TextInputAutomationEventKind::TextChanged:
            static_cast<void>(UiaRaiseAutomationEvent(static_cast<IRawElementProviderSimple*>(provider), UIA_Text_TextChangedEventId));
            return true;
        case TextInputAutomationEventKind::TextSelectionChanged:
            static_cast<void>(UiaRaiseAutomationEvent(static_cast<IRawElementProviderSimple*>(provider), UIA_Text_TextSelectionChangedEventId));
            return true;
        case TextInputAutomationEventKind::ActiveTextPositionChanged:
        {
            wil::com_ptr_nothrow<ITextRangeProvider> activeRange;
            NativeTextInputState state{};
            if (host->TryReadNativeTextInputState(control, state))
            {
                const size_t caretIndex = std::min(state.caretIndex, state.text.size());
                static_cast<void>(target->AddRef());
                auto* rangeProvider = new (std::nothrow) AccessibilityTextRangeProvider(target, hwnd, controlPath, caretIndex, caretIndex);
                if (rangeProvider)
                {
                    activeRange.attach(static_cast<ITextRangeProvider*>(rangeProvider));
                }
                else
                {
                    static_cast<void>(target->Release());
                }
            }
            static_cast<void>(
                UiaRaiseActiveTextPositionChangedEvent(static_cast<IRawElementProviderSimple*>(provider), activeRange.get()));
            return true;
        }
        case TextInputAutomationEventKind::TextEditCompositionChanged:
        {
            unique_safearray changedData(SafeArrayCreateVector(VT_BSTR, 0u, 0u));
            static_cast<void>(UiaRaiseTextEditTextChangedEvent(
                static_cast<IRawElementProviderSimple*>(provider), TextEditChangeType_Composition, changedData.get()));
            return true;
        }
        case TextInputAutomationEventKind::TextEditConversionTargetChanged:
            static_cast<void>(
                UiaRaiseAutomationEvent(static_cast<IRawElementProviderSimple*>(provider), UIA_TextEdit_ConversionTargetChangedEventId));
            return true;
        default: return false;
    }
}

void RegisterWindowHostAccessibilityTarget(HWND hwnd, WindowHost* host) noexcept
{
    if (hwnd && host)
    {
        const std::scoped_lock lock(GetAccessibilityTargetMutex());
        if (auto* target = static_cast<WindowHostAccessibilityTarget*>(GetPropW(hwnd, kWindowHostPropName)))
        {
            target->host.store(host, std::memory_order_release);
            return;
        }

        auto* target = new (std::nothrow) WindowHostAccessibilityTarget(hwnd, host);
        if (! target)
        {
            return;
        }

        if (SetPropW(hwnd, kWindowHostPropName, target) == 0)
        {
            static_cast<void>(target->Release());
        }
    }
}

void UnregisterWindowHostAccessibilityTarget(HWND hwnd, WindowHost* host) noexcept
{
    if (! hwnd)
    {
        return;
    }

    const std::scoped_lock lock(GetAccessibilityTargetMutex());
    auto* target = static_cast<WindowHostAccessibilityTarget*>(GetPropW(hwnd, kWindowHostPropName));
    if (! target)
    {
        return;
    }

    const WindowHost* current = target->host.load(std::memory_order_acquire);
    if (host && current != host)
    {
        return;
    }

    target->host.store(nullptr, std::memory_order_release);
    if (RemovePropW(hwnd, kWindowHostPropName) == target)
    {
        static_cast<void>(target->Release());
    }
}

LRESULT ReturnWindowHostAccessibilityProvider(HWND hwnd, WPARAM wp, LPARAM lp) noexcept
{
    WindowHostAccessibilityTarget* target = AcquireWindowHostAccessibilityTarget(hwnd);
    if (! target || target->ResolveHost() == nullptr)
    {
        if (target)
        {
            static_cast<void>(target->Release());
        }
        return 0;
    }

    auto* provider = new (std::nothrow) AccessibilityProvider(target, hwnd);
    if (! provider)
    {
        static_cast<void>(target->Release());
        return 0;
    }

    const LRESULT result = UiaReturnRawElementProvider(hwnd, wp, lp, static_cast<IRawElementProviderSimple*>(provider));
    static_cast<void>(provider->Release());
    return result;
}

bool TryHandleWindowHostAccessibilityMessage(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, LRESULT& outResult) noexcept
{
    outResult = 0;
    if (msg == WM_GETOBJECT)
    {
        if (lp != static_cast<LPARAM>(UiaRootObjectId))
        {
            return false;
        }

        const LRESULT providerResult = ReturnWindowHostAccessibilityProvider(hwnd, wp, lp);
        if (providerResult == 0)
        {
            return false;
        }

        outResult = providerResult;
        return true;
    }

    if (msg != kWindowHostAccessibilityActionMessage)
    {
        return false;
    }

    auto* request = reinterpret_cast<AccessibilityUiActionRequest*>(lp);
    if (! request)
    {
        return true;
    }

    if (request->provider)
    {
        request->result = request->provider->ExecuteUiThreadAction(request->kind, request->stringValue.c_str(), request->numberValue);
        return true;
    }
    if (request->textRangeProvider && request->kind == AccessibilityUiActionKind::Select)
    {
        request->result = request->textRangeProvider->ExecuteSelectOnWindowThread();
        return true;
    }

    request->result = UIA_E_NOTSUPPORTED;
    return true;
}

#if defined(ENABLE_TESTS)
IRawElementProviderFragmentRoot* CreateWindowHostAccessibilityProvider(HWND hwnd) noexcept
{
    WindowHostAccessibilityTarget* target = AcquireWindowHostAccessibilityTarget(hwnd);
    if (! target || target->ResolveHost() == nullptr)
    {
        if (target)
        {
            static_cast<void>(target->Release());
        }
        return nullptr;
    }

    auto* provider = new (std::nothrow) AccessibilityProvider(target, hwnd);
    if (! provider)
    {
        static_cast<void>(target->Release());
        return nullptr;
    }
    return provider ? static_cast<IRawElementProviderFragmentRoot*>(provider) : nullptr;
}
#endif
} // namespace RedSalamander::DxUi
