#include "DxUi.Internal.h"

#include <algorithm>
#include <cmath>
#include <cwctype>
#include <limits>

#include <imm.h>
#include <richedit.h>
#include <windowsx.h>

#include "DxUi.Typography.h"
#include "Helpers.h"
#include "WindowMessages.h"

#pragma comment(lib, "imm32.lib")

namespace RedSalamander::DxUi
{
namespace
{
constexpr UINT kModifierAlt                                        = 0x0100u;
constexpr int kTextInputBridgeMinWidthPx                           = 64;
constexpr int kTextInputBridgeMinHeightPx                          = 32;
constexpr int kTextInputBridgeControlId                            = 0x7A41;
constexpr wchar_t kTextInputBridgeWindowClass[]                    = L"DxUiTextInputBridgeWindow";
constexpr wchar_t kDxUiTextInputBridgeImeComposingProp[]           = L"DxUiTextInputBridgeImeComposing";
constexpr wchar_t kDxUiTextInputBridgeShiftProp[]                  = L"DxUiTextInputBridgeShift";
constexpr wchar_t kDxUiTextInputBridgeCtrlProp[]                   = L"DxUiTextInputBridgeCtrl";
constexpr wchar_t kDxUiTextInputBridgeAltProp[]                    = L"DxUiTextInputBridgeAlt";
constexpr float kTextInputBridgeFontSizeDip                        = 14.0f;
constexpr std::wstring_view kNonVisibleTextServiceBridgeFontFamily = Typography::kSegoeUiVariableTextFamily;
constexpr float kMultilineLayoutHeightDip                          = 32768.0f;
constexpr uint64_t kCaretBlinkPeriodMs                             = 530u;
constexpr size_t kTextInputBridgeMaxHistoryEntries                 = 256u;

[[nodiscard]] std::wstring NormalizeBridgeTextFromControlText(std::wstring_view text, bool multiline);
[[nodiscard]] std::wstring NormalizeControlTextFromBridgeText(std::wstring_view text, bool multiline);
[[nodiscard]] size_t MapBridgeIndexToControlIndex(std::wstring_view bridgeText, size_t bridgeIndex, bool multiline, bool logicalSelectionNewlines) noexcept;
[[nodiscard]] bool UseLogicalNewlinesForCollapsedBridgeCaret(const TextInputBridgeState& state) noexcept;

[[nodiscard]] bool ModifiersContainCtrl(UINT modifiers) noexcept
{
    return (modifiers & MK_CONTROL) != 0u;
}

[[nodiscard]] bool ModifiersContainShift(UINT modifiers) noexcept
{
    return (modifiers & MK_SHIFT) != 0u;
}

[[nodiscard]] bool ModifiersContainAlt(UINT modifiers) noexcept
{
    return (modifiers & kModifierAlt) != 0u;
}

[[nodiscard]] LONG DipsToNegativePixels(float sizeDip, UINT dpi) noexcept
{
    const UINT effectiveDpi = (std::max<UINT>)(dpi, USER_DEFAULT_SCREEN_DPI);
    return -(std::max<LONG>)(1L, static_cast<LONG>(std::lround((sizeDip * static_cast<float>(effectiveDpi)) / static_cast<float>(USER_DEFAULT_SCREEN_DPI))));
}

[[nodiscard]] wil::unique_hfont CreateNonVisibleTextServiceBridgeFont(UINT dpi) noexcept
{
    static_assert(kNonVisibleTextServiceBridgeFontFamily.size() < LF_FACESIZE);

    LOGFONTW lf{};
    lf.lfHeight        = DipsToNegativePixels(kTextInputBridgeFontSizeDip, dpi);
    lf.lfWeight        = FW_NORMAL;
    lf.lfQuality       = CLEARTYPE_QUALITY;
    lf.lfCharSet       = DEFAULT_CHARSET;
    lf.lfOutPrecision  = OUT_TT_PRECIS;
    lf.lfClipPrecision = CLIP_DEFAULT_PRECIS;
    std::copy(kNonVisibleTextServiceBridgeFontFamily.begin(), kNonVisibleTextServiceBridgeFontFamily.end(), lf.lfFaceName);
    lf.lfFaceName[kNonVisibleTextServiceBridgeFontFamily.size()] = L'\0';
    return wil::unique_hfont(CreateFontIndirectW(&lf));
}

[[nodiscard]] bool EnsureTextInputBridgeWindowClassRegistered() noexcept
{
    static bool registered = false;
    if (registered)
    {
        return true;
    }

    WNDCLASSEXW windowClass{};
    windowClass.cbSize        = sizeof(windowClass);
    windowClass.style         = CS_DBLCLKS;
    windowClass.lpfnWndProc   = TextInputBridgeWndProc;
    windowClass.hInstance     = GetModuleHandleW(nullptr);
    windowClass.lpszClassName = kTextInputBridgeWindowClass;
    if (RegisterClassExW(&windowClass) == 0)
    {
        const DWORD lastError = GetLastError();
        if (lastError != ERROR_CLASS_ALREADY_EXISTS)
        {
            Debug::Error(L"DxUi::WindowHost: RegisterClassExW failed for text-input bridge window ({:08X})", HRESULT_FROM_WIN32(lastError));
            return false;
        }
    }

    registered = true;
    return true;
}

[[nodiscard]] UINT ComputeModifierMask() noexcept
{
    UINT mask = 0;
    if ((GetKeyState(VK_SHIFT) & 0x8000) != 0)
    {
        mask |= MK_SHIFT;
    }
    if ((GetKeyState(VK_CONTROL) & 0x8000) != 0)
    {
        mask |= MK_CONTROL;
    }
    if ((GetKeyState(VK_MENU) & 0x8000) != 0)
    {
        mask |= kModifierAlt;
    }
    if ((GetKeyState(VK_LBUTTON) & 0x8000) != 0)
    {
        mask |= MK_LBUTTON;
    }
    if ((GetKeyState(VK_RBUTTON) & 0x8000) != 0)
    {
        mask |= MK_RBUTTON;
    }
    return mask;
}

void SetTextBridgeModifierState(HWND hwnd, UINT virtualKey, bool keyDown) noexcept
{
    if (! hwnd)
    {
        return;
    }

    const auto setPropFlag = [hwnd, keyDown](const wchar_t* propName) noexcept
    {
        if (keyDown)
        {
            static_cast<void>(SetPropW(hwnd, propName, reinterpret_cast<HANDLE>(1)));
        }
        else
        {
            RemovePropW(hwnd, propName);
        }
    };

    switch (virtualKey)
    {
        case VK_SHIFT: setPropFlag(kDxUiTextInputBridgeShiftProp); break;
        case VK_CONTROL: setPropFlag(kDxUiTextInputBridgeCtrlProp); break;
        case VK_MENU: setPropFlag(kDxUiTextInputBridgeAltProp); break;
        default: break;
    }
}

[[nodiscard]] UINT ComputeTextBridgeModifierMask(HWND hwnd) noexcept
{
    UINT mask = ComputeModifierMask();
    if (hwnd && GetPropW(hwnd, kDxUiTextInputBridgeShiftProp))
    {
        mask |= MK_SHIFT;
    }
    if (hwnd && GetPropW(hwnd, kDxUiTextInputBridgeCtrlProp))
    {
        mask |= MK_CONTROL;
    }
    if (hwnd && GetPropW(hwnd, kDxUiTextInputBridgeAltProp))
    {
        mask |= kModifierAlt;
    }
    return mask;
}

[[nodiscard]] bool IsTextBridgeSpecialKey(UINT msg, UINT virtualKey, UINT modifiers, bool multiline) noexcept
{
    if (virtualKey == VK_TAB || virtualKey == VK_ESCAPE || virtualKey == VK_APPS)
    {
        return true;
    }

    if (virtualKey == VK_RETURN)
    {
        return ! multiline;
    }

    if (virtualKey == VK_F10 && ModifiersContainShift(modifiers))
    {
        return true;
    }

    if (! ModifiersContainAlt(modifiers))
    {
        switch (virtualKey)
        {
            case VK_LEFT:
            case VK_RIGHT:
            case VK_HOME:
            case VK_END:
            case VK_BACK:
            case VK_DELETE: return true;
            case VK_UP:
            case VK_DOWN:
            case VK_PRIOR:
            case VK_NEXT: return multiline;
            case VK_INSERT: return ModifiersContainCtrl(modifiers) || ModifiersContainShift(modifiers);
            default: break;
        }

        if (ModifiersContainCtrl(modifiers))
        {
            switch (virtualKey)
            {
                case 'A':
                case 'C':
                case 'V':
                case 'X':
                case 'Y':
                case 'Z': return true;
                default: break;
            }
        }
    }

    return msg == WM_SYSKEYDOWN && ModifiersContainAlt(modifiers) && (virtualKey == VK_DOWN || virtualKey == VK_UP);
}

void SetTextBridgeImeComposing(HWND hwnd, bool composing) noexcept
{
    if (! hwnd)
    {
        return;
    }

    if (composing)
    {
        static_cast<void>(SetPropW(hwnd, kDxUiTextInputBridgeImeComposingProp, reinterpret_cast<HANDLE>(1)));
    }
    else
    {
        RemovePropW(hwnd, kDxUiTextInputBridgeImeComposingProp);
    }
}

[[nodiscard]] bool TextBridgeImeCompositionHasActiveComposition(LPARAM compositionFlags) noexcept
{
    constexpr LPARAM kActiveCompositionFlags =
        GCS_COMPSTR | GCS_COMPATTR | GCS_COMPCLAUSE | GCS_COMPREADSTR | GCS_COMPREADATTR | GCS_COMPREADCLAUSE | GCS_CURSORPOS | GCS_DELTASTART;
    return (compositionFlags & kActiveCompositionFlags) != 0;
}

[[nodiscard]] bool TextBridgeImeCompositionHasCommittedResult(LPARAM compositionFlags) noexcept
{
    constexpr LPARAM kCommittedResultFlags = GCS_RESULTSTR | GCS_RESULTCLAUSE | GCS_RESULTREADSTR | GCS_RESULTREADCLAUSE;
    return (compositionFlags & kCommittedResultFlags) != 0;
}

[[nodiscard]] bool TextBridgeImeCompositionKeepsCompositionActive(LPARAM compositionFlags) noexcept
{
    return TextBridgeImeCompositionHasActiveComposition(compositionFlags) || ! TextBridgeImeCompositionHasCommittedResult(compositionFlags);
}

[[nodiscard]] std::wstring NormalizeBridgeTextFromControlText(std::wstring_view text, bool multiline)
{
    if (! multiline || text.empty())
    {
        return std::wstring(text);
    }

    std::wstring normalized;
    normalized.reserve(text.size() + static_cast<size_t>(std::count(text.begin(), text.end(), L'\n')));
    for (wchar_t ch : text)
    {
        if (ch == L'\r')
        {
            continue;
        }

        if (ch == L'\n')
        {
            normalized.push_back(L'\r');
        }

        normalized.push_back(ch);
    }

    return normalized;
}

[[nodiscard]] std::wstring NormalizeControlTextFromBridgeText(std::wstring_view text, bool multiline)
{
    if (! multiline || text.empty())
    {
        return std::wstring(text);
    }

    std::wstring normalized;
    normalized.reserve(text.size());
    for (size_t index = 0u; index < text.size(); ++index)
    {
        const wchar_t ch = text[index];
        if (ch == L'\r')
        {
            if (index + 1u < text.size() && text[index + 1u] == L'\n')
            {
                ++index;
            }

            normalized.push_back(L'\n');
            continue;
        }

        normalized.push_back(ch);
    }

    return normalized;
}

[[nodiscard]] std::wstring NormalizePastedControlText(std::wstring_view text, bool multiline)
{
    if (text.empty())
    {
        return std::wstring(text);
    }

    if (multiline)
    {
        return NormalizeControlTextFromBridgeText(text, true);
    }

    std::wstring normalized;
    normalized.reserve(text.size());
    for (wchar_t ch : text)
    {
        if (std::iswcntrl(static_cast<wint_t>(ch)) != 0)
        {
            continue;
        }

        normalized.push_back(ch);
    }

    return normalized;
}

[[nodiscard]] size_t MapControlIndexToBridgeIndex(std::wstring_view controlText, size_t controlIndex, bool multiline, bool logicalSelectionNewlines) noexcept
{
    const size_t clampedIndex = std::min(controlIndex, controlText.size());
    if (! multiline)
    {
        return clampedIndex;
    }

    size_t bridgeIndex = 0u;
    for (size_t index = 0u; index < clampedIndex; ++index)
    {
        if (controlText[index] == L'\n')
        {
            bridgeIndex += logicalSelectionNewlines ? 1u : 2u;
            continue;
        }

        bridgeIndex += 1u;
    }
    return bridgeIndex;
}

[[nodiscard]] size_t MapBridgeIndexToControlIndex(std::wstring_view bridgeText, size_t bridgeIndex, bool multiline, bool logicalSelectionNewlines) noexcept
{
    const size_t clampedIndex = std::min(bridgeIndex, bridgeText.size());
    if (! multiline)
    {
        return clampedIndex;
    }

    size_t controlIndex = 0u;
    size_t index        = 0u;
    size_t logicalIndex = 0u;
    while (index < bridgeText.size() && logicalIndex < clampedIndex)
    {
        if (bridgeText[index] == L'\r' && index + 1u < bridgeText.size() && bridgeText[index + 1u] == L'\n')
        {
            if (logicalSelectionNewlines)
            {
                index += 2u;
                ++logicalIndex;
                ++controlIndex;
                continue;
            }

            if (index + 2u <= clampedIndex)
            {
                index += 2u;
                logicalIndex += 2u;
                ++controlIndex;
                continue;
            }

            break;
        }

        ++index;
        ++logicalIndex;
        ++controlIndex;
    }

    return controlIndex;
}

struct TextBridgeImeAnchor
{
    POINT compositionPoint{};
    RECT exclusionRect{};
};

[[nodiscard]] bool TryGetTextBridgeSelectionEnd(HWND edit, size_t& outSelectionEnd) noexcept
{
    if (! edit)
    {
        return false;
    }

    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    SendMessageW(edit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd));
    outSelectionEnd = static_cast<size_t>(selectionEnd);
    return true;
}

[[nodiscard]] bool IsRichEditBridge(HWND edit) noexcept
{
    if (! edit)
    {
        return false;
    }

    std::array<wchar_t, 16> className{};
    const int classLength = GetClassNameW(edit, className.data(), static_cast<int>(className.size()));
    return classLength > 0 && _wcsicmp(className.data(), MSFTEDIT_CLASS) == 0;
}

[[nodiscard]] bool TryGetTextBridgeCaretClientRect(HWND edit, RECT& outCaretRect) noexcept
{
    if (! edit)
    {
        return false;
    }

    size_t selectionEnd = 0u;
    static_cast<void>(TryGetTextBridgeSelectionEnd(edit, selectionEnd));

    std::optional<POINT> caretPoint;
    if (IsRichEditBridge(edit))
    {
        POINTL richEditPoint{};
        if (SendMessageW(edit, EM_POSFROMCHAR, reinterpret_cast<WPARAM>(&richEditPoint), static_cast<LPARAM>(selectionEnd)) != -1)
        {
            caretPoint = POINT{richEditPoint.x, richEditPoint.y};
        }
    }
    else
    {
        const LRESULT position = SendMessageW(edit, EM_POSFROMCHAR, static_cast<WPARAM>(selectionEnd), 0);
        caretPoint             = POINT{GET_X_LPARAM(position), GET_Y_LPARAM(position)};
    }

    if (! caretPoint.has_value())
    {
        GUITHREADINFO guiThreadInfo{};
        guiThreadInfo.cbSize = sizeof(guiThreadInfo);
        if (GetGUIThreadInfo(0, &guiThreadInfo) != FALSE && guiThreadInfo.hwndCaret)
        {
            RECT caretRect = guiThreadInfo.rcCaret;
            if (guiThreadInfo.hwndCaret != edit)
            {
                MapWindowPoints(guiThreadInfo.hwndCaret, edit, reinterpret_cast<POINT*>(&caretRect), 2);
            }

            if (caretRect.right <= caretRect.left)
            {
                caretRect.right = caretRect.left + 1;
            }
            if (caretRect.bottom <= caretRect.top)
            {
                caretRect.bottom = caretRect.top + std::max(12, GetSystemMetrics(SM_CYCURSOR));
            }

            outCaretRect = caretRect;
            return true;
        }

        return false;
    }

    RECT textRect{};
    if (SendMessageW(edit, EM_GETRECT, 0, reinterpret_cast<LPARAM>(&textRect)) == 0)
    {
        GetClientRect(edit, &textRect);
    }

    const LONG caretHeight = std::max<LONG>(12, std::max(1L, textRect.bottom - textRect.top));
    outCaretRect           = RECT{caretPoint->x, caretPoint->y, caretPoint->x + 1, caretPoint->y + caretHeight};
    return true;
}

void UpdateTextBridgeImeWindows(HWND edit) noexcept
{
    if (! edit)
    {
        return;
    }

    HIMC inputContext = ImmGetContext(edit);
    if (! inputContext)
    {
        return;
    }
    const auto releaseContext = wil::scope_exit([&] { ImmReleaseContext(edit, inputContext); });

    RECT caretRect{};
    if (! TryGetTextBridgeCaretClientRect(edit, caretRect))
    {
        return;
    }

    const TextBridgeImeAnchor anchor{
        .compositionPoint = POINT{caretRect.left, caretRect.top},
        .exclusionRect    = caretRect,
    };

    COMPOSITIONFORM compositionForm{};
    compositionForm.dwStyle      = CFS_FORCE_POSITION;
    compositionForm.ptCurrentPos = anchor.compositionPoint;
    static_cast<void>(ImmSetCompositionWindow(inputContext, &compositionForm));

    CANDIDATEFORM candidateForm{};
    candidateForm.dwIndex      = 0u;
    candidateForm.dwStyle      = CFS_EXCLUDE;
    candidateForm.ptCurrentPos = POINT{caretRect.left, caretRect.bottom};
    candidateForm.rcArea       = anchor.exclusionRect;
    static_cast<void>(ImmSetCandidateWindow(inputContext, &candidateForm));
}

[[nodiscard]] size_t FindLineStart(std::wstring_view text, size_t caretIndex) noexcept
{
    size_t index = std::min(caretIndex, text.size());
    while (index > 0u && text[index - 1u] != L'\n')
    {
        --index;
    }
    return index;
}

[[nodiscard]] size_t FindLineEnd(std::wstring_view text, size_t caretIndex) noexcept
{
    size_t index = std::min(caretIndex, text.size());
    while (index < text.size() && text[index] != L'\n')
    {
        ++index;
    }
    return index;
}

[[nodiscard]] wil::com_ptr<IDWriteTextLayout> CreateMultilineTextLayout(
    const WindowHost* host, std::wstring_view text, FontRole role, float widthDip, float heightDip) noexcept;
[[nodiscard]] std::vector<DWRITE_LINE_METRICS> GetMultilineLineMetrics(IDWriteTextLayout* layout) noexcept;
[[nodiscard]] std::optional<size_t> TryGetMultilineCaretLineIndex(IDWriteTextLayout* layout,
                                                                  const std::vector<DWRITE_LINE_METRICS>& metrics,
                                                                  size_t caretIndex,
                                                                  size_t textLength) noexcept;
struct WrappedLineTextRange
{
    size_t start = 0u;
    size_t end   = 0u;
};
[[nodiscard]] std::vector<WrappedLineTextRange> BuildWrappedLineTextRanges(const std::vector<DWRITE_LINE_METRICS>& metrics, size_t textLength) noexcept;
[[nodiscard]] std::optional<size_t> TryMoveCaretToWrappedLineBoundary(
    const WindowHost* host, std::wstring_view text, FontRole role, const D2D1_RECT_F& textRect, size_t caretIndex, bool moveToLineEnd) noexcept;

[[nodiscard]] size_t MoveCaretVerticallyByLogicalLine(std::wstring_view text,
                                                      size_t caretIndex,
                                                      bool moveDown,
                                                      std::optional<float>& preferredXOffsetDip) noexcept
{
    const size_t caret            = std::min(caretIndex, text.size());
    const size_t currentLineStart = FindLineStart(text, caret);
    const size_t currentLineEnd   = FindLineEnd(text, caret);
    const size_t currentColumn    = std::min(caret - currentLineStart, currentLineEnd - currentLineStart);
    const float targetXOffsetDip  = preferredXOffsetDip.value_or(static_cast<float>(currentColumn) * 7.0f);
    preferredXOffsetDip           = targetXOffsetDip;
    const size_t targetColumn     = static_cast<size_t>(std::max(0.0, std::floor(static_cast<double>(targetXOffsetDip) / 7.0)));

    if (moveDown)
    {
        if (currentLineEnd >= text.size())
        {
            return caret;
        }

        const size_t nextLineStart = currentLineEnd + 1u;
        const size_t nextLineEnd   = FindLineEnd(text, nextLineStart);
        return nextLineStart + std::min(targetColumn, nextLineEnd - nextLineStart);
    }

    if (currentLineStart == 0u)
    {
        return caret;
    }

    const size_t previousLineEnd   = currentLineStart - 1u;
    const size_t previousLineStart = FindLineStart(text, previousLineEnd);
    return previousLineStart + std::min(targetColumn, previousLineEnd - previousLineStart);
}

struct MultilineViewportMetrics
{
    size_t totalLineCount   = 1u;
    size_t visibleLineCount = 1u;
};

[[nodiscard]] size_t CountLogicalTextLines(std::wstring_view text) noexcept
{
    return 1u + static_cast<size_t>(std::count(text.begin(), text.end(), L'\n'));
}

[[nodiscard]] size_t ComputeMultilinePageLineCount(float viewportHeightDip, float lineHeightDip) noexcept
{
    const float effectiveLineHeightDip = std::max(1.0f, lineHeightDip);
    return std::max<size_t>(1u, static_cast<size_t>(std::floor(std::max(1.0f, viewportHeightDip) / effectiveLineHeightDip)));
}

[[nodiscard]] size_t ComputeMultilineWheelLineCount(size_t pageLineCount) noexcept
{
    UINT systemLines = 0u;
    if (SystemParametersInfoW(SPI_GETWHEELSCROLLLINES, 0, &systemLines, 0) == FALSE)
    {
        systemLines = 3u;
    }

    if (systemLines == WHEEL_PAGESCROLL)
    {
        return std::max<size_t>(1u, pageLineCount);
    }

    if (systemLines == 0u)
    {
        return 0u;
    }

    return static_cast<size_t>(systemLines);
}

[[nodiscard]] size_t MoveCaretByPage(
    std::wstring_view text, size_t caretIndex, bool moveDown, std::optional<float>& preferredXOffsetDip, size_t lineCount) noexcept
{
    size_t nextCaretIndex = caretIndex;
    for (size_t step = 0u; step < lineCount; ++step)
    {
        const size_t movedIndex = MoveCaretVerticallyByLogicalLine(text, nextCaretIndex, moveDown, preferredXOffsetDip);
        if (movedIndex == nextCaretIndex)
        {
            break;
        }

        nextCaretIndex = movedIndex;
    }

    return nextCaretIndex;
}

[[nodiscard]] std::optional<size_t> TryMoveCaretVerticallyByWrappedLines(const WindowHost* host,
                                                                         std::wstring_view text,
                                                                         FontRole role,
                                                                         const D2D1_RECT_F& textRect,
                                                                         size_t caretIndex,
                                                                         int visualLineDelta,
                                                                         std::optional<float>& preferredXOffsetDip) noexcept
{
    if (! host || text.empty() || visualLineDelta == 0 || (textRect.right - textRect.left) <= 1.0f || (textRect.bottom - textRect.top) <= 1.0f)
    {
        return std::nullopt;
    }

    wil::com_ptr<IDWriteTextLayout> layout =
        CreateMultilineTextLayout(host, text, role, std::max(1.0f, textRect.right - textRect.left), std::max(1.0f, textRect.bottom - textRect.top));
    if (! layout)
    {
        return std::nullopt;
    }

    const std::vector<DWRITE_LINE_METRICS> metrics = GetMultilineLineMetrics(layout.get());
    if (metrics.empty())
    {
        return std::nullopt;
    }

    float currentX = 0.0f;
    float currentY = 0.0f;
    DWRITE_HIT_TEST_METRICS currentHitMetrics{};
    if (FAILED(layout->HitTestTextPosition(static_cast<UINT32>(std::min(caretIndex, text.size())), FALSE, &currentX, &currentY, &currentHitMetrics)))
    {
        return std::nullopt;
    }

    const std::optional<size_t> currentLineIndex = TryGetMultilineCaretLineIndex(layout.get(), metrics, caretIndex, text.size());
    if (! currentLineIndex)
    {
        return std::nullopt;
    }

    const int maxLineIndex    = static_cast<int>(metrics.size()) - 1;
    const int targetLineIndex = std::clamp(static_cast<int>(currentLineIndex.value()) + visualLineDelta, 0, maxLineIndex);
    if (targetLineIndex == static_cast<int>(currentLineIndex.value()))
    {
        return caretIndex;
    }

    const float targetX = preferredXOffsetDip.value_or(currentX);
    preferredXOffsetDip = targetX;

    float targetTopDip = 0.0f;
    for (int index = 0; index < targetLineIndex; ++index)
    {
        targetTopDip += std::max(1.0f, metrics[static_cast<size_t>(index)].height);
    }

    const float targetHeightDip = std::max(1.0f, metrics[static_cast<size_t>(targetLineIndex)].height);
    const float targetY         = targetTopDip + std::min(targetHeightDip - 1.0f, targetHeightDip * 0.5f);

    BOOL isTrailingHit = FALSE;
    BOOL isInside      = FALSE;
    DWRITE_HIT_TEST_METRICS targetHitMetrics{};
    if (FAILED(layout->HitTestPoint(targetX, targetY, &isTrailingHit, &isInside, &targetHitMetrics)))
    {
        return std::nullopt;
    }

    const size_t textPosition    = static_cast<size_t>(targetHitMetrics.textPosition);
    const size_t trailingAdvance = isTrailingHit ? static_cast<size_t>(targetHitMetrics.length) : 0u;
    return std::min(text.size(), textPosition + trailingAdvance);
}

[[nodiscard]] size_t MoveMultilineCaretVertically(const WindowHost* host,
                                                  std::wstring_view text,
                                                  FontRole role,
                                                  const D2D1_RECT_F& textRect,
                                                  size_t caretIndex,
                                                  bool moveDown,
                                                  std::optional<float>& preferredXOffsetDip) noexcept
{
    if (const std::optional<size_t> wrappedCaretIndex =
            TryMoveCaretVerticallyByWrappedLines(host, text, role, textRect, caretIndex, moveDown ? 1 : -1, preferredXOffsetDip))
    {
        return wrappedCaretIndex.value();
    }

    return MoveCaretVerticallyByLogicalLine(text, caretIndex, moveDown, preferredXOffsetDip);
}

[[nodiscard]] size_t MoveMultilineCaretByPage(const WindowHost* host,
                                              std::wstring_view text,
                                              FontRole role,
                                              const D2D1_RECT_F& textRect,
                                              size_t caretIndex,
                                              bool moveDown,
                                              std::optional<float>& preferredXOffsetDip,
                                              size_t lineCount) noexcept
{
    if (lineCount == 0u)
    {
        return caretIndex;
    }

    if (const std::optional<size_t> wrappedCaretIndex = TryMoveCaretVerticallyByWrappedLines(
            host, text, role, textRect, caretIndex, moveDown ? static_cast<int>(lineCount) : -static_cast<int>(lineCount), preferredXOffsetDip))
    {
        return wrappedCaretIndex.value();
    }

    return MoveCaretByPage(text, caretIndex, moveDown, preferredXOffsetDip, lineCount);
}

[[nodiscard]] std::vector<WrappedLineTextRange> BuildWrappedLineTextRanges(const std::vector<DWRITE_LINE_METRICS>& metrics, size_t textLength) noexcept
{
    std::vector<WrappedLineTextRange> ranges;
    ranges.reserve(metrics.size());

    size_t lineStart = 0u;
    for (const DWRITE_LINE_METRICS& metric : metrics)
    {
        const size_t rawLineEnd = std::min(textLength, lineStart + static_cast<size_t>(metric.length));
        size_t visibleLineEnd   = rawLineEnd;
        if (metric.newlineLength > 0u)
        {
            const size_t newlineLength = static_cast<size_t>(metric.newlineLength);
            visibleLineEnd             = (visibleLineEnd >= newlineLength) ? visibleLineEnd - newlineLength : lineStart;
        }

        ranges.push_back({lineStart, std::max(lineStart, visibleLineEnd)});
        lineStart = rawLineEnd;
    }

    return ranges;
}

[[nodiscard]] std::optional<size_t> TryMoveCaretToWrappedLineBoundary(
    const WindowHost* host, std::wstring_view text, FontRole role, const D2D1_RECT_F& textRect, size_t caretIndex, bool moveToLineEnd) noexcept
{
    if (! host || text.empty() || (textRect.right - textRect.left) <= 1.0f || (textRect.bottom - textRect.top) <= 1.0f)
    {
        return std::nullopt;
    }

    const wil::com_ptr<IDWriteTextLayout> layout =
        CreateMultilineTextLayout(host, text, role, std::max(1.0f, textRect.right - textRect.left), std::max(1.0f, textRect.bottom - textRect.top));
    if (! layout)
    {
        return std::nullopt;
    }

    const std::vector<DWRITE_LINE_METRICS> metrics = GetMultilineLineMetrics(layout.get());
    if (metrics.empty() || metrics.size() <= CountLogicalTextLines(text))
    {
        return std::nullopt;
    }

    const std::optional<size_t> currentLineIndex = TryGetMultilineCaretLineIndex(layout.get(), metrics, caretIndex, text.size());
    if (! currentLineIndex)
    {
        return std::nullopt;
    }

    const std::vector<WrappedLineTextRange> ranges = BuildWrappedLineTextRanges(metrics, text.size());
    if (currentLineIndex.value() >= ranges.size())
    {
        return std::nullopt;
    }

    const WrappedLineTextRange& range = ranges[currentLineIndex.value()];
    return moveToLineEnd ? range.end : range.start;
}

[[nodiscard]] bool UseLogicalNewlinesForCollapsedBridgeCaret(const TextInputBridgeState& state) noexcept
{
    return state.multiline && ! state.selectionAnchorIndex.has_value();
}

[[nodiscard]] wil::com_ptr<IDWriteTextLayout> CreateMultilineTextLayout(
    const WindowHost* host, std::wstring_view text, FontRole role, float widthDip, float heightDip) noexcept
{
    if (! host)
    {
        return {};
    }

    auto* factory = host->GetWriteFactory();
    auto* format  = host->GetTextFormat(role, DWRITE_TEXT_ALIGNMENT_LEADING, DWRITE_PARAGRAPH_ALIGNMENT_NEAR, true);
    if (! factory || ! format)
    {
        return {};
    }

    wil::com_ptr<IDWriteTextLayout> layout;
    if (FAILED(factory->CreateTextLayout(text.data(),
                                         static_cast<UINT32>(text.size()),
                                         format,
                                         std::max(1.0f, widthDip),
                                         std::max(kMultilineLayoutHeightDip, heightDip),
                                         layout.addressof())) ||
        ! layout)
    {
        return {};
    }
    return layout;
}

[[nodiscard]] std::vector<DWRITE_LINE_METRICS> GetMultilineLineMetrics(IDWriteTextLayout* layout) noexcept
{
    if (! layout)
    {
        return {};
    }

    UINT32 actualLineCount = 0u;
    HRESULT hr             = layout->GetLineMetrics(nullptr, 0u, &actualLineCount);
    if (FAILED(hr) && actualLineCount == 0u)
    {
        return {};
    }

    std::vector<DWRITE_LINE_METRICS> metrics(actualLineCount);
    if (actualLineCount == 0u)
    {
        return metrics;
    }

    if (FAILED(layout->GetLineMetrics(metrics.data(), actualLineCount, &actualLineCount)))
    {
        return {};
    }

    metrics.resize(actualLineCount);
    return metrics;
}

[[nodiscard]] float EstimateMultilineFallbackLineHeightDip(const WindowHost* host, FontRole role) noexcept
{
    if (! host)
    {
        return 20.0f;
    }

    auto* factory = host->GetWriteFactory();
    auto* format  = host->GetTextFormat(role, DWRITE_TEXT_ALIGNMENT_LEADING, DWRITE_PARAGRAPH_ALIGNMENT_NEAR, true);
    if (! factory || ! format)
    {
        return 20.0f;
    }

    static constexpr wchar_t kSampleText[] = L"Ag";
    wil::com_ptr<IDWriteTextLayout> layout;
    if (FAILED(factory->CreateTextLayout(
            kSampleText, static_cast<UINT32>(std::size(kSampleText) - 1u), format, 1024.0f, kMultilineLayoutHeightDip, layout.addressof())) ||
        ! layout)
    {
        return 20.0f;
    }

    const std::vector<DWRITE_LINE_METRICS> metrics = GetMultilineLineMetrics(layout.get());
    if (! metrics.empty())
    {
        return std::max(1.0f, metrics.front().height);
    }

    DWRITE_TEXT_METRICS textMetrics{};
    if (SUCCEEDED(layout->GetMetrics(&textMetrics)) && textMetrics.height > 0.0f)
    {
        return textMetrics.height;
    }

    return 20.0f;
}

[[nodiscard]] float MeasureWrappedLineOffsetDip(const std::vector<DWRITE_LINE_METRICS>& metrics, size_t firstVisibleLine) noexcept
{
    if (metrics.empty() || firstVisibleLine == 0u)
    {
        return 0.0f;
    }

    const size_t clampedFirstVisibleLine = std::min(firstVisibleLine, metrics.size());
    float offsetDip                      = 0.0f;
    for (size_t index = 0u; index < clampedFirstVisibleLine; ++index)
    {
        offsetDip += metrics[index].height;
    }
    return offsetDip;
}

[[nodiscard]] size_t CountVisibleWrappedLines(const std::vector<DWRITE_LINE_METRICS>& metrics, float viewportHeightDip, float fallbackLineHeightDip) noexcept
{
    if (metrics.empty())
    {
        return ComputeMultilinePageLineCount(viewportHeightDip, fallbackLineHeightDip);
    }

    const float clampedViewportHeightDip = std::max(1.0f, viewportHeightDip);
    float consumedHeightDip              = 0.0f;
    size_t visibleLineCount              = 0u;
    for (const DWRITE_LINE_METRICS& metric : metrics)
    {
        ++visibleLineCount;
        consumedHeightDip += std::max(1.0f, metric.height);
        if (consumedHeightDip >= clampedViewportHeightDip)
        {
            break;
        }
    }

    return std::max<size_t>(1u, visibleLineCount);
}

[[nodiscard]] MultilineViewportMetrics BuildMultilineViewportMetrics(std::wstring_view text,
                                                                     float viewportHeightDip,
                                                                     const std::vector<DWRITE_LINE_METRICS>& metrics,
                                                                     float fallbackLineHeightDip) noexcept
{
    MultilineViewportMetrics viewportMetrics;
    if (metrics.empty())
    {
        viewportMetrics.totalLineCount   = CountLogicalTextLines(text);
        viewportMetrics.visibleLineCount = ComputeMultilinePageLineCount(viewportHeightDip, fallbackLineHeightDip);
        return viewportMetrics;
    }

    viewportMetrics.totalLineCount   = std::max<size_t>(1u, metrics.size());
    viewportMetrics.visibleLineCount = CountVisibleWrappedLines(metrics, viewportHeightDip, fallbackLineHeightDip);
    return viewportMetrics;
}

[[nodiscard]] size_t ComputeMultilineMaxFirstVisibleLine(const MultilineViewportMetrics& viewportMetrics) noexcept
{
    // Preserve imported/view-driven top-line state up to the last content line.
    // The DxUI surface can legitimately show trailing blank space below the last line.
    return viewportMetrics.totalLineCount > 0u ? viewportMetrics.totalLineCount - 1u : 0u;
}

[[nodiscard]] size_t ClampMultilineFirstVisibleLine(size_t firstVisibleLine, const MultilineViewportMetrics& viewportMetrics) noexcept
{
    return (std::min)(firstVisibleLine, ComputeMultilineMaxFirstVisibleLine(viewportMetrics));
}

[[nodiscard]] std::optional<size_t> TryGetMultilineCaretLineIndex(IDWriteTextLayout* layout,
                                                                  const std::vector<DWRITE_LINE_METRICS>& metrics,
                                                                  size_t caretIndex,
                                                                  size_t textLength) noexcept
{
    if (! layout || metrics.empty())
    {
        return std::nullopt;
    }

    float x = 0.0f;
    float y = 0.0f;
    DWRITE_HIT_TEST_METRICS hitMetrics{};
    if (FAILED(layout->HitTestTextPosition(static_cast<UINT32>(std::min(caretIndex, textLength)), FALSE, &x, &y, &hitMetrics)))
    {
        return std::nullopt;
    }

    float lineTopDip = 0.0f;
    for (size_t index = 0u; index < metrics.size(); ++index)
    {
        lineTopDip += std::max(1.0f, metrics[index].height);
        if (y < lineTopDip || index + 1u == metrics.size())
        {
            return index;
        }
    }

    return std::nullopt;
}

[[nodiscard]] MultilineViewportMetrics ComputeMultilineViewportMetrics(const WindowHost* host,
                                                                       std::wstring_view text,
                                                                       FontRole role,
                                                                       const D2D1_RECT_F& textRect) noexcept
{
    const float viewportHeightDip     = std::max(1.0f, textRect.bottom - textRect.top);
    const auto multilineLayout        = CreateMultilineTextLayout(host, text, role, std::max(1.0f, textRect.right - textRect.left), viewportHeightDip);
    const float fallbackLineHeightDip = EstimateMultilineFallbackLineHeightDip(host, role);
    return BuildMultilineViewportMetrics(text, viewportHeightDip, GetMultilineLineMetrics(multilineLayout.get()), fallbackLineHeightDip);
}

[[nodiscard]] size_t HitTestMultilineCaretIndexDip(
    const WindowHost* host, std::wstring_view text, FontRole role, const D2D1_RECT_F& textRect, float scrollDip, D2D1_POINT_2F point) noexcept
{
    if (text.empty())
    {
        return 0u;
    }

    const float localX = std::max(0.0f, point.x - textRect.left);
    const float localY = std::max(0.0f, point.y - textRect.top + scrollDip);
    if (wil::com_ptr<IDWriteTextLayout> layout =
            CreateMultilineTextLayout(host, text, role, std::max(1.0f, textRect.right - textRect.left), std::max(1.0f, textRect.bottom - textRect.top)))
    {
        BOOL isTrailingHit = FALSE;
        BOOL isInside      = FALSE;
        DWRITE_HIT_TEST_METRICS metrics{};
        if (SUCCEEDED(layout->HitTestPoint(localX, localY, &isTrailingHit, &isInside, &metrics)))
        {
            const size_t textPosition    = static_cast<size_t>(metrics.textPosition);
            const size_t trailingAdvance = isTrailingHit ? static_cast<size_t>(metrics.length) : 0u;
            return std::min(text.size(), textPosition + trailingAdvance);
        }
    }

    size_t lineStart = 0u;
    size_t lineIndex = static_cast<size_t>(std::max(0.0, std::floor(static_cast<double>(localY) / 18.0)));
    while (lineIndex > 0u && lineStart < text.size())
    {
        const size_t nextBreak = text.find(L'\n', lineStart);
        if (nextBreak == std::wstring_view::npos)
        {
            return text.size();
        }
        lineStart = nextBreak + 1u;
        --lineIndex;
    }

    const size_t lineEnd        = text.find(L'\n', lineStart);
    const size_t clampedLineEnd = lineEnd == std::wstring_view::npos ? text.size() : lineEnd;
    const size_t column = static_cast<size_t>(std::clamp(std::floor(static_cast<double>(localX) / 7.0), 0.0, static_cast<double>(clampedLineEnd - lineStart)));
    return std::min(text.size(), lineStart + column);
}

[[nodiscard]] D2D1_RECT_F MeasureMultilineCaretRectDip(
    const WindowHost* host, std::wstring_view text, FontRole role, const D2D1_RECT_F& textRect, float scrollDip, size_t caretIndex) noexcept
{
    const D2D1_RECT_F fallbackRect = D2D1::RectF(textRect.left, textRect.top + 2.0f, textRect.left + 1.0f, textRect.top + 18.0f);
    if (wil::com_ptr<IDWriteTextLayout> layout =
            CreateMultilineTextLayout(host, text, role, std::max(1.0f, textRect.right - textRect.left), std::max(1.0f, textRect.bottom - textRect.top)))
    {
        float x = 0.0f;
        float y = 0.0f;
        DWRITE_HIT_TEST_METRICS metrics{};
        if (SUCCEEDED(layout->HitTestTextPosition(static_cast<UINT32>(std::min(caretIndex, text.size())), FALSE, &x, &y, &metrics)))
        {
            const float caretTop    = textRect.top + y - scrollDip;
            const float caretHeight = std::max(12.0f, metrics.height);
            return D2D1::RectF(textRect.left + x, caretTop, textRect.left + x + 1.0f, caretTop + caretHeight);
        }
    }

    return fallbackRect;
}

void DrawMultilineSelection(WindowHost& host,
                            std::wstring_view text,
                            const D2D1_RECT_F& rect,
                            FontRole role,
                            const D2D1_COLOR_F& textColor,
                            const D2D1_COLOR_F& selectionFill,
                            const D2D1_COLOR_F& selectionText,
                            float scrollDip,
                            std::optional<std::pair<size_t, size_t>> selectionRange) noexcept
{
    auto* dc                = host.GetDeviceContext();
    auto* fillBrush         = host.GetSolidBrush(selectionFill);
    auto* textBrush         = host.GetSolidBrush(textColor);
    auto* selectedTextBrush = host.GetSolidBrush(selectionText);
    if (! dc || ! textBrush)
    {
        return;
    }

    const D2D1_RECT_F snappedRect          = SnapRectToPixel(host, rect);
    wil::com_ptr<IDWriteTextLayout> layout = CreateMultilineTextLayout(
        &host, text, role, std::max(1.0f, snappedRect.right - snappedRect.left), std::max(1.0f, snappedRect.bottom - snappedRect.top));
    if (! layout)
    {
        DrawCenteredText(host, text, snappedRect, role, textColor, DWRITE_TEXT_ALIGNMENT_LEADING, DWRITE_PARAGRAPH_ALIGNMENT_NEAR, true);
        return;
    }

    dc->PushAxisAlignedClip(snappedRect, D2D1_ANTIALIAS_MODE_PER_PRIMITIVE);
    dc->DrawTextLayout(D2D1::Point2F(snappedRect.left, snappedRect.top - scrollDip), layout.get(), textBrush, kTextDrawOptions);

    if (selectionRange && fillBrush && selectedTextBrush)
    {
        const auto [selectionStart, selectionEnd] = selectionRange.value();
        if (selectionStart < selectionEnd)
        {
            UINT32 actualCount = 0u;
            std::vector<DWRITE_HIT_TEST_METRICS> metrics(static_cast<size_t>(selectionEnd - selectionStart) + 4u);
            HRESULT hr = layout->HitTestTextRange(static_cast<UINT32>(selectionStart),
                                                  static_cast<UINT32>(selectionEnd - selectionStart),
                                                  snappedRect.left,
                                                  snappedRect.top - scrollDip,
                                                  metrics.data(),
                                                  static_cast<UINT32>(metrics.size()),
                                                  &actualCount);
            if (hr == HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER))
            {
                metrics.resize(actualCount);
                hr = layout->HitTestTextRange(static_cast<UINT32>(selectionStart),
                                              static_cast<UINT32>(selectionEnd - selectionStart),
                                              snappedRect.left,
                                              snappedRect.top - scrollDip,
                                              metrics.data(),
                                              static_cast<UINT32>(metrics.size()),
                                              &actualCount);
            }

            if (SUCCEEDED(hr))
            {
                for (UINT32 index = 0u; index < actualCount; ++index)
                {
                    const auto& hit           = metrics[index];
                    const D2D1_RECT_F hitRect = SnapRectToPixel(host, D2D1::RectF(hit.left, hit.top + 1.0f, hit.left + hit.width, hit.top + hit.height - 1.0f));
                    if (hitRect.right > hitRect.left)
                    {
                        dc->FillRectangle(hitRect, fillBrush);
                    }
                }

                for (UINT32 index = 0u; index < actualCount; ++index)
                {
                    const auto& hit           = metrics[index];
                    const D2D1_RECT_F hitRect = SnapRectToPixel(host, D2D1::RectF(hit.left, hit.top, hit.left + hit.width, hit.top + hit.height));
                    if (hitRect.right <= hitRect.left)
                    {
                        continue;
                    }

                    dc->PushAxisAlignedClip(hitRect, D2D1_ANTIALIAS_MODE_PER_PRIMITIVE);
                    dc->DrawTextLayout(D2D1::Point2F(snappedRect.left, snappedRect.top - scrollDip), layout.get(), selectedTextBrush, kTextDrawOptions);
                    dc->PopAxisAlignedClip();
                }
            }
        }
    }

    dc->PopAxisAlignedClip();
}

[[nodiscard]] std::wstring GetBridgeTextFromState(const TextInputBridgeState& state)
{
    return NormalizeBridgeTextFromControlText(state.text, state.multiline);
}

[[nodiscard]] std::pair<size_t, size_t> GetTextInputBridgeSelectionRange(const TextInputBridgeState& state) noexcept
{
    const size_t anchorIndex = state.selectionAnchorIndex.value_or(state.caretIndex);
    return {(std::min)(anchorIndex, state.caretIndex), (std::max)(anchorIndex, state.caretIndex)};
}

void SetTextInputBridgeSelectionRange(TextInputBridgeState& state, size_t start, size_t end) noexcept
{
    start = (std::min)(start, state.text.size());
    end   = (std::min)(end, state.text.size());
    if (start >= end)
    {
        state.selectionAnchorIndex.reset();
        state.caretIndex = end;
        return;
    }

    state.selectionAnchorIndex = start;
    state.caretIndex           = end;
}

[[nodiscard]] bool ReplaceTextInputBridgeSelection(TextInputBridgeState& state, std::wstring_view replacement)
{
    if (state.readOnly)
    {
        return false;
    }

    auto [selectionStart, selectionEnd] = GetTextInputBridgeSelectionRange(state);
    selectionStart                      = (std::min)(selectionStart, state.text.size());
    selectionEnd                        = (std::min)(selectionEnd, state.text.size());
    state.text.replace(selectionStart, selectionEnd - selectionStart, replacement);
    state.selectionAnchorIndex.reset();
    state.caretIndex = selectionStart + replacement.size();
    return true;
}

[[nodiscard]] std::optional<std::wstring> GetSelectedBridgeTextFromState(const TextInputBridgeState& state)
{
    const auto [selectionStart, selectionEnd] = GetTextInputBridgeSelectionRange(state);
    if (selectionEnd <= selectionStart || selectionEnd > state.text.size())
    {
        return std::nullopt;
    }

    return NormalizeBridgeTextFromControlText(state.text.substr(selectionStart, selectionEnd - selectionStart), state.multiline);
}

} // namespace

LRESULT CALLBACK TextInputBridgeWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    if (msg == WM_NCCREATE)
    {
        const auto* createStruct = reinterpret_cast<const CREATESTRUCTW*>(lp);
        static_cast<void>(SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(createStruct ? createStruct->lpCreateParams : nullptr)));
        return TRUE;
    }

    auto* host = reinterpret_cast<WindowHost*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (! host)
    {
        return DefWindowProcW(hwnd, msg, wp, lp);
    }

    if (msg == WM_NCDESTROY)
    {
        const LRESULT result = host->HandleTextInputBridgeWindowMessage(hwnd, msg, wp, lp);
        static_cast<void>(SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0));
        return result;
    }

    return host->HandleTextInputBridgeWindowMessage(hwnd, msg, wp, lp);
}

LRESULT WindowHost::HandleTextInputBridgeWindowMessage(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    if (hwnd != _textInputBridgeEdit.get())
    {
        return DefWindowProcW(hwnd, msg, wp, lp);
    }

    const auto readState        = [this]() noexcept -> std::optional<TextInputBridgeState> { return ReadTextInputBridgeState(); };
    const auto pushHistoryState = [](std::vector<TextInputBridgeState>& history, const TextInputBridgeState& state) noexcept
    {
        history.push_back(state);
        if (history.size() > kTextInputBridgeMaxHistoryEntries)
        {
            history.erase(history.begin());
        }
    };
    const auto pushUndoState = [this, &pushHistoryState](const TextInputBridgeState& before) noexcept
    {
        pushHistoryState(_textInputBridgeUndoHistory, before);
        _textInputBridgeRedoHistory.clear();
    };
    const auto commitState =
        [this, hwnd, &pushUndoState](const TextInputBridgeState& before, const TextInputBridgeState& after, bool notifyChange, bool recordUndo) noexcept
    {
        if (recordUndo && before.text != after.text)
        {
            pushUndoState(before);
        }

        ApplyTextInputBridgeState(after);
        SyncFocusedControlFromTextInputBridge(notifyChange);
        UpdateTextBridgeImeWindows(hwnd);
        return 0;
    };
    const auto tryGetViewportRect = [this]() noexcept -> std::optional<RECT>
    {
        if (! _textInputBridgeControl)
        {
            return std::nullopt;
        }

        RECT clientRect{};
        if (GetClientRect(_textInputBridgeEdit.get(), &clientRect) == FALSE)
        {
            return std::nullopt;
        }

        RECT viewportRect = clientRect;
        if (const std::optional<D2D1_RECT_F> viewportDip = _textInputBridgeControl->GetTextInputBridgeViewportRect(); viewportDip.has_value())
        {
            const D2D1_RECT_F bounds = _textInputBridgeControl->GetBounds();
            const int leftPx         = static_cast<int>(std::lround(DipsToPixels(viewportDip->left - bounds.left)));
            const int topPx          = static_cast<int>(std::lround(DipsToPixels(viewportDip->top - bounds.top)));
            const int rightPx        = static_cast<int>(std::lround(DipsToPixels(bounds.right - viewportDip->right)));
            const int bottomPx       = static_cast<int>(std::lround(DipsToPixels(bounds.bottom - viewportDip->bottom)));

            viewportRect.left   = static_cast<LONG>((std::clamp)(leftPx, 0, static_cast<int>(clientRect.right)));
            viewportRect.top    = static_cast<LONG>((std::clamp)(topPx, 0, static_cast<int>(clientRect.bottom)));
            viewportRect.right  = static_cast<LONG>((std::clamp)(static_cast<int>(clientRect.right) - (std::max)(0, rightPx),
                                                                 static_cast<int>(viewportRect.left) + 1,
                                                                 static_cast<int>(clientRect.right)));
            viewportRect.bottom = static_cast<LONG>((std::clamp)(static_cast<int>(clientRect.bottom) - (std::max)(0, bottomPx),
                                                                 static_cast<int>(viewportRect.top) + 1,
                                                                 static_cast<int>(clientRect.bottom)));
        }

        return viewportRect;
    };
    const auto tryGetCaretRect = [this](const TextInputBridgeState& state) noexcept -> std::optional<RECT>
    {
        if (! _textInputBridgeControl)
        {
            return std::nullopt;
        }

        if (const std::optional<D2D1_RECT_F> caretRectDip = _textInputBridgeControl->GetTextInputBridgeCaretRect(*this, state.caretIndex);
            caretRectDip.has_value())
        {
            const D2D1_RECT_F bounds = _textInputBridgeControl->GetBounds();
            RECT caretRect{static_cast<LONG>(std::lround(DipsToPixels(caretRectDip->left - bounds.left))),
                           static_cast<LONG>(std::lround(DipsToPixels(caretRectDip->top - bounds.top))),
                           static_cast<LONG>(std::lround(DipsToPixels(caretRectDip->right - bounds.left))),
                           static_cast<LONG>(std::lround(DipsToPixels(caretRectDip->bottom - bounds.top)))};
            if (caretRect.right <= caretRect.left)
            {
                caretRect.right = caretRect.left + 1;
            }
            if (caretRect.bottom <= caretRect.top)
            {
                caretRect.bottom = caretRect.top + std::max<LONG>(12, GetSystemMetrics(SM_CYCURSOR));
            }
            return caretRect;
        }

        return std::nullopt;
    };

    switch (msg)
    {
        case WM_NCHITTEST: return _textInputBridgeImeComposing ? HTCLIENT : HTTRANSPARENT;
        case WM_GETDLGCODE: return DLGC_WANTCHARS | DLGC_WANTARROWS | DLGC_WANTTAB;
        case WM_KILLFOCUS: return _hwnd ? SendMessageW(_hwnd, WndMsg::kDxUiTextInputBridgeBlur, wp, lp) : 0;
        case WM_LBUTTONDOWN:
        case WM_RBUTTONDOWN:
        case WM_LBUTTONDBLCLK:
        case WM_RBUTTONDBLCLK:
        {
            if (! _textInputBridgeControl)
            {
                return 0;
            }

            POINT pointPx{GET_X_LPARAM(lp), GET_Y_LPARAM(lp)};
            if (_hwnd)
            {
                static_cast<void>(MapWindowPoints(hwnd, _hwnd, &pointPx, 1));
            }

            const D2D1_POINT_2F point = D2D1::Point2F(PixelsToDip(static_cast<float>(pointPx.x)), PixelsToDip(static_cast<float>(pointPx.y)));
            const bool rightButton    = msg == WM_RBUTTONDOWN || msg == WM_RBUTTONDBLCLK;
            const bool doubleClick    = msg == WM_LBUTTONDBLCLK || msg == WM_RBUTTONDBLCLK;
            const bool handled        = doubleClick ? _textInputBridgeControl->OnMouseDoubleClick(*this, point, rightButton, static_cast<UINT>(wp))
                                                    : _textInputBridgeControl->OnMouseDown(*this, point, rightButton, static_cast<UINT>(wp));
            if (handled)
            {
                CaptureMouse(_textInputBridgeControl);
            }
            return 0;
        }
        case WM_LBUTTONUP:
        case WM_RBUTTONUP:
        {
            if (_textInputBridgeControl)
            {
                POINT pointPx{GET_X_LPARAM(lp), GET_Y_LPARAM(lp)};
                if (_hwnd)
                {
                    static_cast<void>(MapWindowPoints(hwnd, _hwnd, &pointPx, 1));
                }

                const D2D1_POINT_2F point = D2D1::Point2F(PixelsToDip(static_cast<float>(pointPx.x)), PixelsToDip(static_cast<float>(pointPx.y)));
                const bool rightButton    = msg == WM_RBUTTONUP;
                _textInputBridgeControl->OnMouseUp(*this, point, rightButton, static_cast<UINT>(wp));
            }

            ReleaseMouseCapture();
            return 0;
        }
        case WM_KEYDOWN:
        case WM_SYSKEYDOWN:
        {
            const UINT virtualKey = static_cast<UINT>(wp);
            SetTextBridgeModifierState(hwnd, virtualKey, true);

            const TextInputBridgeState state = readState().value_or(TextInputBridgeState{});
            UINT modifiers                   = ComputeTextBridgeModifierMask(hwnd);
            const bool systemKey             = msg == WM_SYSKEYDOWN;
            if (systemKey && virtualKey != VK_F10 && virtualKey != VK_MENU)
            {
                modifiers |= kModifierAlt;
            }

            if (IsTextBridgeSpecialKey(msg, virtualKey, modifiers, state.multiline))
            {
                if (_textInputBridgeImeComposing)
                {
                    return 0;
                }

                return _hwnd ? SendMessageW(_hwnd,
                                            WndMsg::kDxUiTextInputBridgeSpecialKey,
                                            wp,
                                            MAKELPARAM(static_cast<WORD>(modifiers), static_cast<WORD>(systemKey ? 1u : 0u)))
                             : 0;
            }
            break;
        }
        case WM_KEYUP:
        case WM_SYSKEYUP: SetTextBridgeModifierState(hwnd, static_cast<UINT>(wp), false); return 0;
        case WM_CHAR:
        case WM_SYSCHAR:
        {
            const std::optional<TextInputBridgeState> beforeState = readState();
            if (! beforeState.has_value())
            {
                return 0;
            }

            TextInputBridgeState afterState = beforeState.value();
            if (! afterState.multiline && wp == L'\t')
            {
                return 0;
            }
            if (afterState.readOnly)
            {
                return 0;
            }
            if (wp == L'\r')
            {
                if (! afterState.multiline)
                {
                    return 0;
                }

                _textInputBridgeSelectionLogicalNewlines = UseLogicalNewlinesForCollapsedBridgeCaret(afterState);
                return commitState(beforeState.value(), afterState, true, ReplaceTextInputBridgeSelection(afterState, L"\n"));
            }
            if (std::iswcntrl(static_cast<wint_t>(wp)) != 0)
            {
                return 0;
            }

            const wchar_t insertedChar[]             = {static_cast<wchar_t>(wp), L'\0'};
            _textInputBridgeSelectionLogicalNewlines = UseLogicalNewlinesForCollapsedBridgeCaret(afterState);
            return commitState(beforeState.value(), afterState, true, ReplaceTextInputBridgeSelection(afterState, insertedChar));
        }
        case WM_IME_STARTCOMPOSITION:
            _textInputBridgeImeComposing      = true;
            _textInputBridgeImeHasVisibleText = true;
            _textInputBridgeImeBaseState      = readState();
            SetTextBridgeImeComposing(hwnd, true);
            UpdateTextBridgeImeWindows(hwnd);
            return 0;
        case WM_IME_COMPOSITION:
            if (TextBridgeImeCompositionKeepsCompositionActive(lp))
            {
                _textInputBridgeImeComposing = true;
                if (TextBridgeImeCompositionHasActiveComposition(lp))
                {
                    _textInputBridgeImeHasVisibleText = true;
                }
                SetTextBridgeImeComposing(hwnd, true);
            }
            else if (TextBridgeImeCompositionHasCommittedResult(lp))
            {
                _textInputBridgeImeComposing = false;
                SetTextBridgeImeComposing(hwnd, false);
            }
            UpdateTextBridgeImeWindows(hwnd);
            return 0;
        case WM_IME_ENDCOMPOSITION:
            _textInputBridgeImeComposing      = false;
            _textInputBridgeImeHasVisibleText = false;
            _textInputBridgeImeBaseState.reset();
            SetTextBridgeImeComposing(hwnd, false);
            UpdateTextBridgeImeWindows(hwnd);
            return 0;
        case WM_UNDO:
        {
            const std::optional<TextInputBridgeState> currentState = readState();
            if (! currentState.has_value() || _textInputBridgeUndoHistory.empty())
            {
                return FALSE;
            }

            pushHistoryState(_textInputBridgeRedoHistory, currentState.value());
            const TextInputBridgeState restoredState = std::move(_textInputBridgeUndoHistory.back());
            _textInputBridgeUndoHistory.pop_back();
            _textInputBridgeSelectionLogicalNewlines = UseLogicalNewlinesForCollapsedBridgeCaret(restoredState);
            ApplyTextInputBridgeState(restoredState);
            SyncFocusedControlFromTextInputBridge(true);
            UpdateTextBridgeImeWindows(hwnd);
            return TRUE;
        }
        case EM_REDO:
        {
            const std::optional<TextInputBridgeState> currentState = readState();
            if (! currentState.has_value() || _textInputBridgeRedoHistory.empty())
            {
                return FALSE;
            }

            pushHistoryState(_textInputBridgeUndoHistory, currentState.value());
            const TextInputBridgeState restoredState = std::move(_textInputBridgeRedoHistory.back());
            _textInputBridgeRedoHistory.pop_back();
            _textInputBridgeSelectionLogicalNewlines = UseLogicalNewlinesForCollapsedBridgeCaret(restoredState);
            ApplyTextInputBridgeState(restoredState);
            SyncFocusedControlFromTextInputBridge(true);
            UpdateTextBridgeImeWindows(hwnd);
            return TRUE;
        }
        case EM_GETSEL:
        {
            const std::optional<TextInputBridgeState> state = readState();
            if (! state.has_value())
            {
                return 0;
            }

            const auto [selectionStart, selectionEnd] = GetTextInputBridgeSelectionRange(state.value());
            const bool logicalSelectionNewlines       = state->multiline && _textInputBridgeSelectionLogicalNewlines;
            const size_t bridgeSelectionStart         = MapControlIndexToBridgeIndex(state->text, selectionStart, state->multiline, logicalSelectionNewlines);
            const size_t bridgeSelectionEnd           = MapControlIndexToBridgeIndex(state->text, selectionEnd, state->multiline, logicalSelectionNewlines);
            const DWORD selectionStartDword = static_cast<DWORD>((std::min)(bridgeSelectionStart, static_cast<size_t>(std::numeric_limits<DWORD>::max())));
            const DWORD selectionEndDword   = static_cast<DWORD>((std::min)(bridgeSelectionEnd, static_cast<size_t>(std::numeric_limits<DWORD>::max())));

            if (wp != 0)
            {
                *reinterpret_cast<DWORD*>(wp) = selectionStartDword;
            }
            if (lp != 0)
            {
                *reinterpret_cast<DWORD*>(lp) = selectionEndDword;
            }

            return MAKELRESULT(static_cast<WORD>(selectionStartDword & 0xFFFFu), static_cast<WORD>(selectionEndDword & 0xFFFFu));
        }
        case EM_SETSEL:
        {
            const std::optional<TextInputBridgeState> beforeState = readState();
            if (! beforeState.has_value())
            {
                return 0;
            }

            TextInputBridgeState afterState     = beforeState.value();
            const bool logicalSelectionNewlines = afterState.multiline && _textInputBridgeSelectionLogicalNewlines;
            const std::wstring bridgeText       = GetBridgeTextFromState(afterState);
            const auto mapSelectionIndex        = [&](LPARAM value) noexcept -> size_t
            {
                if (value < 0)
                {
                    return afterState.text.size();
                }

                return MapBridgeIndexToControlIndex(bridgeText, static_cast<size_t>(value), afterState.multiline, logicalSelectionNewlines);
            };

            const size_t selectionStart = mapSelectionIndex(static_cast<LPARAM>(wp));
            const size_t selectionEnd   = mapSelectionIndex(lp);
            SetTextInputBridgeSelectionRange(afterState, selectionStart, selectionEnd);
            _textInputBridgeSelectionLogicalNewlines = UseLogicalNewlinesForCollapsedBridgeCaret(afterState);
            return commitState(beforeState.value(), afterState, false, false);
        }
        case EM_REPLACESEL:
        {
            const std::optional<TextInputBridgeState> beforeState = readState();
            if (! beforeState.has_value())
            {
                return FALSE;
            }

            TextInputBridgeState afterState          = beforeState.value();
            const wchar_t* const replacementText     = lp != 0 ? reinterpret_cast<const wchar_t*>(lp) : L"";
            const std::wstring normalizedReplacement = NormalizeControlTextFromBridgeText(replacementText, afterState.multiline);
            const bool replaced                      = ReplaceTextInputBridgeSelection(afterState, normalizedReplacement);
            _textInputBridgeSelectionLogicalNewlines = UseLogicalNewlinesForCollapsedBridgeCaret(afterState);
            return commitState(beforeState.value(), afterState, true, replaced && wp != FALSE);
        }
        case EM_POSFROMCHAR:
        {
            const std::optional<TextInputBridgeState> state = readState();
            if (! state.has_value())
            {
                return 0;
            }

            const bool logicalSelectionNewlines = state->multiline && _textInputBridgeSelectionLogicalNewlines;
            const std::wstring bridgeText       = GetBridgeTextFromState(state.value());
            TextInputBridgeState caretState     = state.value();
            caretState.selectionAnchorIndex.reset();
            caretState.caretIndex = MapBridgeIndexToControlIndex(bridgeText, static_cast<size_t>(wp), state->multiline, logicalSelectionNewlines);

            if (const std::optional<RECT> caretRect = tryGetCaretRect(caretState); caretRect.has_value())
            {
                return MAKELRESULT(static_cast<SHORT>(caretRect->left), static_cast<SHORT>(caretRect->top));
            }
            return 0;
        }
        case EM_GETRECT:
        {
            RECT viewportRect{};
            if (const std::optional<RECT> candidate = tryGetViewportRect(); candidate.has_value())
            {
                viewportRect = candidate.value();
            }
            else
            {
                GetClientRect(hwnd, &viewportRect);
            }

            if (lp != 0)
            {
                *reinterpret_cast<RECT*>(lp) = viewportRect;
            }
            return 1;
        }
        case EM_GETFIRSTVISIBLELINE:
        {
            const std::optional<TextInputBridgeState> state = readState();
            if (! state.has_value() || ! state->multiline)
            {
                return 0;
            }

            return static_cast<LRESULT>((std::min)(state->firstVisibleLine, static_cast<size_t>(std::numeric_limits<LRESULT>::max())));
        }
        case WM_COPY:
        {
            const std::optional<TextInputBridgeState> state = readState();
            if (! state.has_value() || state->masked)
            {
                return 0;
            }

            const std::optional<std::wstring> selectedText = GetSelectedBridgeTextFromState(state.value());
            if (! selectedText.has_value())
            {
                return 0;
            }

            return CopyTextToClipboard(selectedText.value()) ? 1 : 0;
        }
        case WM_CUT:
        {
            const std::optional<TextInputBridgeState> beforeState = readState();
            if (! beforeState.has_value() || beforeState->masked || beforeState->readOnly)
            {
                return 0;
            }

            const std::optional<std::wstring> selectedText = GetSelectedBridgeTextFromState(beforeState.value());
            if (! selectedText.has_value() || ! CopyTextToClipboard(selectedText.value()))
            {
                return 0;
            }

            TextInputBridgeState afterState = beforeState.value();
            if (! ReplaceTextInputBridgeSelection(afterState, L""))
            {
                return 0;
            }

            _textInputBridgeSelectionLogicalNewlines = UseLogicalNewlinesForCollapsedBridgeCaret(afterState);
            return commitState(beforeState.value(), afterState, true, true);
        }
        case WM_CLEAR:
        {
            const std::optional<TextInputBridgeState> beforeState = readState();
            if (! beforeState.has_value() || beforeState->readOnly)
            {
                return 0;
            }

            TextInputBridgeState afterState           = beforeState.value();
            const auto [selectionStart, selectionEnd] = GetTextInputBridgeSelectionRange(afterState);
            if (selectionEnd <= selectionStart || ! ReplaceTextInputBridgeSelection(afterState, L""))
            {
                return 0;
            }

            _textInputBridgeSelectionLogicalNewlines = UseLogicalNewlinesForCollapsedBridgeCaret(afterState);
            return commitState(beforeState.value(), afterState, true, true);
        }
        case WM_PASTE:
        {
            const std::optional<TextInputBridgeState> beforeState = readState();
            if (! beforeState.has_value() || beforeState->readOnly)
            {
                return 0;
            }

            const std::optional<std::wstring> clipboardText = ReadTextFromClipboard();
            if (! clipboardText.has_value())
            {
                return 0;
            }

            const std::wstring normalizedClipboardText = NormalizePastedControlText(clipboardText.value(), beforeState->multiline);
            TextInputBridgeState afterState            = beforeState.value();
            const auto [selectionStart, selectionEnd]  = GetTextInputBridgeSelectionRange(afterState);
            if (selectionEnd <= selectionStart && normalizedClipboardText.empty())
            {
                return 0;
            }

            if (! ReplaceTextInputBridgeSelection(afterState, normalizedClipboardText))
            {
                return 0;
            }

            _textInputBridgeSelectionLogicalNewlines = UseLogicalNewlinesForCollapsedBridgeCaret(afterState);
            return commitState(beforeState.value(), afterState, true, true);
        }
        case WM_SETTEXT:
        {
            TextInputBridgeState state = readState().value_or(TextInputBridgeState{});
            const wchar_t* const text  = lp != 0 ? reinterpret_cast<const wchar_t*>(lp) : L"";
            state.text                 = NormalizeControlTextFromBridgeText(text, state.multiline);
            state.selectionAnchorIndex.reset();
            state.caretIndex                         = state.text.size();
            state.firstVisibleLine                   = 0u;
            _textInputBridgeSelectionLogicalNewlines = UseLogicalNewlinesForCollapsedBridgeCaret(state);
            return commitState(state, state, true, false);
        }
        case WM_GETTEXTLENGTH:
        {
            const std::optional<TextInputBridgeState> state = readState();
            return state ? static_cast<LRESULT>(GetBridgeTextFromState(state.value()).size()) : 0;
        }
        case WM_GETTEXT:
        {
            if (wp == 0 || lp == 0)
            {
                return 0;
            }

            const std::optional<TextInputBridgeState> state = readState();
            const std::wstring bridgeText                   = state ? GetBridgeTextFromState(state.value()) : std::wstring{};
            const size_t bufferChars                        = static_cast<size_t>(wp);
            const size_t copyChars                          = bufferChars > 0u ? (std::min)(bridgeText.size(), bufferChars - 1u) : 0u;
            auto* buffer                                    = reinterpret_cast<wchar_t*>(lp);
            if (copyChars > 0u)
            {
                std::copy_n(bridgeText.data(), copyChars, buffer);
            }
            buffer[copyChars] = L'\0';
            return static_cast<LRESULT>(copyChars);
        }
        default: break;
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

bool WindowHost::EnsureTextInputBridge(bool multiline) noexcept
{
    if (_textInputBridgeEdit && _textInputBridgeMultiline == multiline)
    {
        return true;
    }

    if (_textInputBridgeEdit && _textInputBridgeMultiline != multiline)
    {
        DestroyTextInputBridge();
    }

    if (! _hwnd || ! EnsureTextInputBridgeWindowClassRegistered())
    {
        return false;
    }

    wil::unique_hwnd bridgeEdit(CreateWindowExW(WS_EX_NOPARENTNOTIFY,
                                                kTextInputBridgeWindowClass,
                                                L"",
                                                WS_CHILD | WS_VISIBLE,
                                                0,
                                                0,
                                                kTextInputBridgeMinWidthPx,
                                                kTextInputBridgeMinHeightPx,
                                                _hwnd,
                                                reinterpret_cast<HMENU>(static_cast<INT_PTR>(kTextInputBridgeControlId)),
                                                GetModuleHandleW(nullptr),
                                                this));
    if (! bridgeEdit)
    {
        static_cast<void>(Debug::ErrorWithLastError(L"DxUi::WindowHost: CreateWindowExW failed for text-input bridge service"));
        return false;
    }

    // Keep the native text helper anchored over the real DX control so IME candidate/composition UI
    // can follow the visible caret, but suppress all helper chrome so only the DX surface is visible.
    wil::unique_hrgn emptyRegion(CreateRectRgn(0, 0, 0, 0));
    if (emptyRegion && SetWindowRgn(bridgeEdit.get(), emptyRegion.get(), FALSE) != 0)
    {
        static_cast<void>(emptyRegion.release());
    }
    else if (emptyRegion)
    {
        static_cast<void>(Debug::Warning(L"DxUi::WindowHost: SetWindowRgn failed for hidden text bridge"));
    }

    _nonVisibleTextServiceBridgeFontDpi      = 0u;
    _textInputBridgeEdit                     = std::move(bridgeEdit);
    _textInputBridgeMultiline                = multiline;
    _textInputBridgeStateCache               = TextInputBridgeState{.multiline = multiline};
    _textInputBridgeStateCacheValid          = false;
    _textInputBridgeSelectionLogicalNewlines = true;
    return true;
}

void WindowHost::DestroyTextInputBridge() noexcept
{
    _textInputBridgeControl = nullptr;
    _textInputBridgeSyncing = false;
    _textInputBridgeEdit.reset();
    _nonVisibleTextServiceBridgeFont.reset();
    _nonVisibleTextServiceBridgeFontDpi = 0u;
    _textInputBridgeModule.reset();
    _textInputBridgeMultiline                = false;
    _textInputBridgeStateCache               = {};
    _textInputBridgeStateCacheValid          = false;
    _textInputBridgeSelectionLogicalNewlines = true;
    _textInputBridgeImeComposing             = false;
    _textInputBridgeImeHasVisibleText        = false;
    _textInputBridgeImeBaseState.reset();
    _textInputBridgeUndoHistory.clear();
    _textInputBridgeRedoHistory.clear();
}

void WindowHost::UpdateTextInputBridgeBounds(Control* control, bool multiline) noexcept
{
    if (! _textInputBridgeEdit || ! control)
    {
        return;
    }

    const D2D1_RECT_F bounds = control->GetBounds();
    const int xPx            = static_cast<int>(std::lround(DipsToPixels(bounds.left)));
    const int yPx            = static_cast<int>(std::lround(DipsToPixels(bounds.top)));
    const float widthDip     = std::max(1.0f, bounds.right - bounds.left);
    const float heightDip    = std::max(1.0f, bounds.bottom - bounds.top);
    const int widthPx        = std::max(kTextInputBridgeMinWidthPx, static_cast<int>(std::lround(DipsToPixels(widthDip))));
    const int heightPx =
        std::max(multiline ? kTextInputBridgeMinHeightPx : kTextInputBridgeMinHeightPx, static_cast<int>(std::lround(DipsToPixels(heightDip))));

    const UINT fontDpi = (std::max<UINT>)(_dpi, USER_DEFAULT_SCREEN_DPI);
    if (! _nonVisibleTextServiceBridgeFont || _nonVisibleTextServiceBridgeFontDpi != fontDpi)
    {
        wil::unique_hfont font = CreateNonVisibleTextServiceBridgeFont(fontDpi);
        if (font)
        {
            _nonVisibleTextServiceBridgeFont    = std::move(font);
            _nonVisibleTextServiceBridgeFontDpi = fontDpi;
        }
        else
        {
            Debug::Warning(L"DxUi::WindowHost: CreateFontIndirectW failed for non-visible text-service bridge font");
        }
    }
    if (_nonVisibleTextServiceBridgeFont)
    {
        static_cast<void>(SendMessageW(_textInputBridgeEdit.get(), WM_SETFONT, reinterpret_cast<WPARAM>(_nonVisibleTextServiceBridgeFont.get()), FALSE));
    }

    static_cast<void>(SetWindowPos(_textInputBridgeEdit.get(), nullptr, xPx, yPx, widthPx, heightPx, SWP_NOACTIVATE | SWP_NOOWNERZORDER | SWP_NOZORDER));
    _textInputBridgeMultiline = multiline;
    UpdateTextBridgeImeWindows(_textInputBridgeEdit.get());
}

void WindowHost::ActivateTextInputBridge(Control* control) noexcept
{
    if (! control || ! control->SupportsTextInputBridge())
    {
        return;
    }

    TextInputBridgeState state;
    if (! control->ExportTextInputBridgeState(state))
    {
        return;
    }

    if (! EnsureTextInputBridge(state.multiline))
    {
        return;
    }

    const bool controlChanged                = _textInputBridgeControl != control;
    _textInputBridgeControl                  = control;
    _textInputBridgeSelectionLogicalNewlines = true;
    ApplyTextInputBridgeState(state);
    if (controlChanged)
    {
        _textInputBridgeUndoHistory.clear();
        _textInputBridgeRedoHistory.clear();
    }
    if (GetFocus() != _textInputBridgeEdit.get())
    {
        SetFocus(_textInputBridgeEdit.get());
    }
    UpdateTextBridgeImeWindows(_textInputBridgeEdit.get());
}

void WindowHost::DeactivateTextInputBridge(bool restoreHostFocus) noexcept
{
    if (! _textInputBridgeEdit)
    {
        _textInputBridgeControl = nullptr;
        return;
    }

    if (_textInputBridgeControl && ! _textInputBridgeSyncing)
    {
        SyncFocusedControlFromTextInputBridge(true);
    }
    _textInputBridgeControl           = nullptr;
    _textInputBridgeImeComposing      = false;
    _textInputBridgeImeHasVisibleText = false;
    _textInputBridgeImeBaseState.reset();

    if (restoreHostFocus && _hwnd && GetFocus() == _textInputBridgeEdit.get())
    {
        SetFocus(_hwnd);
    }
}

void WindowHost::ApplyTextInputBridgeState(const TextInputBridgeState& state) noexcept
{
    _textInputBridgeStateCache      = state;
    _textInputBridgeStateCacheValid = true;
    if (_textInputBridgeEdit)
    {
        LONG_PTR style = GetWindowLongPtrW(_textInputBridgeEdit.get(), GWL_STYLE);
        if (state.multiline)
        {
            style |= ES_MULTILINE;
        }
        else
        {
            style &= ~static_cast<LONG_PTR>(ES_MULTILINE);
        }

        if (state.readOnly)
        {
            style |= ES_READONLY;
        }
        else
        {
            style &= ~static_cast<LONG_PTR>(ES_READONLY);
        }

        static_cast<void>(SetWindowLongPtrW(_textInputBridgeEdit.get(), GWL_STYLE, style));
    }
    if (_textInputBridgeControl)
    {
        UpdateTextInputBridgeBounds(_textInputBridgeControl, state.multiline);
    }
}

std::optional<TextInputBridgeState> WindowHost::ReadTextInputBridgeState() const noexcept
{
    if (! _textInputBridgeEdit || ! _textInputBridgeStateCacheValid)
    {
        return std::nullopt;
    }
    return _textInputBridgeStateCache;
}

void WindowHost::SyncFocusedControlFromTextInputBridge(bool notifyChange) noexcept
{
    if (_textInputBridgeSyncing || ! _textInputBridgeControl || _textInputBridgeControl != _focusedControl)
    {
        return;
    }

    const std::optional<TextInputBridgeState> state = ReadTextInputBridgeState();
    if (! state)
    {
        return;
    }

    _textInputBridgeSyncing = true;
    const auto resetSyncing = wil::scope_exit([&] { _textInputBridgeSyncing = false; });

    Control* const bridgeControl = _textInputBridgeControl;
    static_cast<void>(bridgeControl->ImportTextInputBridgeState(*this, state.value(), notifyChange));

    // If the OnTextChanged callback modified the control's text (e.g. input
    // normalization/filtering), sync the corrected state back to the bridge
    // EDIT so both sides stay in agreement.
    TextInputBridgeState postState;
    if (bridgeControl == _textInputBridgeControl && bridgeControl->ExportTextInputBridgeState(postState))
    {
        ApplyTextInputBridgeState(postState);
    }
}

void WindowHost::HandleTextInputBridgeCommand(UINT /*notifyCode*/) noexcept
{
}

void WindowHost::CommitFocusedTextInputBridge(bool notifyChange) noexcept
{
    SyncFocusedControlFromTextInputBridge(notifyChange);
}

bool WindowHost::TryReadTextInputBridgeState(const Control* control, TextInputBridgeState& outState) const noexcept
{
    if (! control || control != _textInputBridgeControl)
    {
        return false;
    }

    const std::optional<TextInputBridgeState> state = ReadTextInputBridgeState();
    if (! state.has_value())
    {
        return false;
    }

    outState = state.value();
    return true;
}

void WindowHost::SyncTextInputBridge(Control* control) noexcept
{
    TextInputBridgeState state;
    if (! control || control != _textInputBridgeControl || ! control->ExportTextInputBridgeState(state))
    {
        return;
    }

    if (! EnsureTextInputBridge(state.multiline))
    {
        return;
    }

    _textInputBridgeSelectionLogicalNewlines = true;
    ApplyTextInputBridgeState(state);
    UpdateTextBridgeImeWindows(_textInputBridgeEdit.get());
}

bool WindowHost::HasActiveTextInputBridge() const noexcept
{
    return _textInputBridgeEdit && _textInputBridgeControl != nullptr;
}

#if defined(ENABLE_TESTS)
bool WindowHost::DebugGetNonVisibleTextServiceBridgeFont(LOGFONTW& outLogFont) const noexcept
{
    outLogFont = LOGFONTW{};
    return _nonVisibleTextServiceBridgeFont && GetObjectW(_nonVisibleTextServiceBridgeFont.get(), sizeof(outLogFont), &outLogFont) == sizeof(outLogFont);
}
#endif

HWND WindowHost::GetTextInputBridgeHwnd() const noexcept
{
    return _textInputBridgeEdit.get();
}

TextField::TextField(std::wstring text) : _text(std::move(text))
{
    _caretIndex = _text.size();
    SetFocusable(true);
}

void TextField::SetText(std::wstring text)
{
    _text       = std::move(text);
    _caretIndex = std::min(_caretIndex, _text.size());
    _selectionAnchorIndex.reset();
    ResetSingleLineSelectionClickSequence(_selectionClickSequence);
    _preferredMultilineXOffsetDip.reset();
    _caretBlinkAnchorTickMs       = 0u;
    _caretVisible                 = true;
    _horizontalScrollDip          = 0.0f;
    _multilineFirstVisibleLine    = 0u;
    _multilineWheelDeltaRemainder = 0.0f;
    _dragSelecting                = false;
    _undoHistory.clear();
    _redoHistory.clear();
    InvalidateMultilineLayoutCache();
    RequestInvalidate();
}

std::wstring_view TextField::GetText() const noexcept
{
    return _text;
}

void TextField::SetSelectionRange(const size_t selectionStart, const size_t selectionEnd) noexcept
{
    const size_t clampedStart = std::min(selectionStart, _text.size());
    const size_t clampedEnd   = std::min(selectionEnd, _text.size());
    _preferredMultilineXOffsetDip.reset();
    _dragSelecting = false;
    ResetSingleLineSelectionClickSequence(_selectionClickSequence);
    if (clampedStart == clampedEnd)
    {
        _selectionAnchorIndex.reset();
        _caretIndex = clampedEnd;
    }
    else
    {
        _selectionAnchorIndex = clampedStart;
        _caretIndex           = clampedEnd;
    }
    _caretBlinkAnchorTickMs = 0u;
    _caretVisible           = true;
    RequestInvalidate();
}

void TextField::SetTextAndNotify(std::wstring text)
{
    SetText(std::move(text));
    NotifyChanged();
}

void TextField::SetMasked(bool masked) noexcept
{
    _masked = masked;
}

bool TextField::IsMasked() const noexcept
{
    return _masked;
}

void TextField::SetPlaceholder(std::wstring text)
{
    _placeholder = std::move(text);
}

void TextField::SetMultiline(bool multiline) noexcept
{
    _multiline = multiline;
    _preferredMultilineXOffsetDip.reset();
    _multilineFirstVisibleLine    = 0u;
    _multilineWheelDeltaRemainder = 0.0f;
}

void TextField::SetClearButtonEnabled(bool enabled) noexcept
{
    _clearButtonEnabled = enabled;
    if (! enabled)
    {
        _clearButtonHovered = false;
    }
}

bool TextField::IsClearButtonEnabled() const noexcept
{
    return _clearButtonEnabled;
}

void TextField::SetCaretColor(std::optional<D2D1_COLOR_F> caretColor) noexcept
{
    _caretColorOverride = caretColor;
}

void TextField::SetHorizontalTextPadding(float leftDip, float rightDip) noexcept
{
    _textPaddingLeftDip  = std::max(0.0f, leftDip);
    _textPaddingRightDip = std::max(0.0f, rightDip);
}

void TextField::SetVerticalTextPadding(float topDip, float bottomDip) noexcept
{
    _textPaddingTopDip    = std::max(0.0f, topDip);
    _textPaddingBottomDip = std::max(0.0f, bottomDip);
}

void TextField::SetReadOnly(bool readOnly) noexcept
{
    _readOnly = readOnly;
}

bool TextField::IsReadOnly() const noexcept
{
    return _readOnly;
}

void TextField::SetOnTextChanged(std::function<void(std::wstring_view)> onTextChanged)
{
    _onTextChanged = std::move(onTextChanged);
}

void TextField::SetOnSubmitted(std::function<void()> onSubmitted)
{
    _onSubmitted = std::move(onSubmitted);
}

void TextField::SetOnBlur(std::function<void()> onBlur)
{
    _onBlur = std::move(onBlur);
}

bool TextField::DebugGetMultilineState(const WindowHost& host, TextFieldDebugMultilineState& out) const noexcept
{
    out = {};
    if (! _multiline)
    {
        return false;
    }

    const D2D1_RECT_F textRect     = GetTextRect();
    const float viewportHeightDip  = std::max(1.0f, textRect.bottom - textRect.top);
    const std::wstring displayText = GetDisplayText();
    const auto multilineLayout =
        CreateMultilineTextLayout(&host, displayText, FontRole::Body, std::max(1.0f, textRect.right - textRect.left), viewportHeightDip);
    const std::vector<DWRITE_LINE_METRICS> lineMetrics = GetMultilineLineMetrics(multilineLayout.get());
    const float fallbackLineHeightDip                  = EstimateMultilineFallbackLineHeightDip(&host, FontRole::Body);
    const MultilineViewportMetrics viewportMetrics     = BuildMultilineViewportMetrics(displayText, viewportHeightDip, lineMetrics, fallbackLineHeightDip);
    const size_t clampedFirstVisibleLine               = ClampMultilineFirstVisibleLine(_multilineFirstVisibleLine, viewportMetrics);

    out.firstVisibleLine      = clampedFirstVisibleLine;
    out.visibleLineCount      = viewportMetrics.visibleLineCount;
    out.totalLineCount        = viewportMetrics.totalLineCount;
    out.canScrollVertically   = viewportMetrics.totalLineCount > viewportMetrics.visibleLineCount;
    out.cachedLayoutPresent   = static_cast<bool>(_cachedMultilineLayout);
    out.layoutDirty           = _multilineLayoutDirty;
    out.cachedLayoutWidthDip  = _cachedLayoutSize.width;
    out.cachedLayoutHeightDip = _cachedLayoutSize.height;
    return true;
}

bool TextField::DebugGetSingleLinePaintState(const WindowHost& host, TextFieldDebugSingleLinePaintState& out) const noexcept
{
    out = {};
    if (_multiline)
    {
        return false;
    }

    const D2D1_RECT_F textRect = GetTextRect();
    EnsureCaretVisible(&host, std::max(1.0f, textRect.right - textRect.left));

    out.textRect            = SnapRectToPixel(host, textRect);
    out.horizontalScrollDip = _horizontalScrollDip;
    if (const std::optional<D2D1_RECT_F> selectionPaintRect =
            ComputeSingleLineSelectionPaintRect(host, GetDisplayText(), textRect, FontRole::Body, _horizontalScrollDip, GetSelectionRange());
        selectionPaintRect.has_value())
    {
        out.selectionPaintRect    = selectionPaintRect.value();
        out.hasSelectionPaintRect = true;
    }
    return true;
}

void TextField::OnBoundsChanged() noexcept
{
    InvalidateMultilineLayoutCache();
    if (! _multiline)
    {
        _horizontalScrollDip = 0.0f;
        return;
    }

    WindowHost* const host = GetHost();
    if (! host)
    {
        _multilineFirstVisibleLine    = 0u;
        _multilineWheelDeltaRemainder = 0.0f;
        return;
    }

    const D2D1_RECT_F textRect     = GetTextRect();
    const std::wstring displayText = GetDisplayText();
    if (displayText.empty())
    {
        _multilineFirstVisibleLine    = 0u;
        _multilineWheelDeltaRemainder = 0.0f;
        return;
    }

    const float viewportHeightDip = std::max(1.0f, textRect.bottom - textRect.top);
    const auto multilineLayout =
        CreateMultilineTextLayout(host, displayText, FontRole::Body, std::max(1.0f, textRect.right - textRect.left), viewportHeightDip);
    const std::vector<DWRITE_LINE_METRICS> lineMetrics = GetMultilineLineMetrics(multilineLayout.get());
    const float fallbackLineHeightDip                  = EstimateMultilineFallbackLineHeightDip(host, FontRole::Body);
    const MultilineViewportMetrics viewportMetrics     = BuildMultilineViewportMetrics(displayText, viewportHeightDip, lineMetrics, fallbackLineHeightDip);
    _multilineFirstVisibleLine                         = ClampMultilineFirstVisibleLine(_multilineFirstVisibleLine, viewportMetrics);
    if (HasFocus() && ! _readOnly)
    {
        EnsureMultilineCaretVisible(host);
        host->SyncTextInputBridge(this);
    }
}

void TextField::Paint(WindowHost& host) const
{
    const TextFieldVisualStyle style =
        ResolveTextFieldVisualStyle(host.GetTheme(), IsEnabled(), IsHovered(), HasFocus(), HasFocus() && host.IsKeyboardFocusVisible(), _caretColorOverride);
    constexpr float kTextFieldCornerRadiusDip = 4.0f;
    DrawRoundedRect(host, GetBounds(), style.fill, style.border, kTextFieldCornerRadiusDip);
    if (style.showFocus)
    {
        // WinUI accent bottom border: 2px accent line at bottom of control when focused
        // Inset horizontally by corner radius to avoid clipping into rounded corners
        if (auto* dc = host.GetDeviceContext())
        {
            const D2D1_RECT_F bounds = GetBounds();
            const D2D1_RECT_F accentBar =
                D2D1::RectF(bounds.left + kTextFieldCornerRadiusDip, bounds.bottom - 2.0f, bounds.right - kTextFieldCornerRadiusDip, bounds.bottom);
            dc->FillRectangle(&accentBar, host.GetSolidBrush(host.GetTheme().accent));
        }
    }

    const D2D1_RECT_F textRect     = GetTextRect();
    const std::wstring displayText = GetDisplayText();
    const bool usePlaceholder      = _text.empty() && ! _placeholder.empty();
    if (_multiline)
    {
        const auto multilineLayout =
            GetOrCreateMultilineLayout(&host, displayText, std::max(1.0f, textRect.right - textRect.left), std::max(1.0f, textRect.bottom - textRect.top));
        const std::vector<DWRITE_LINE_METRICS> lineMetrics = GetMultilineLineMetrics(multilineLayout.get());
        const float multilineScrollDip                     = MeasureWrappedLineOffsetDip(lineMetrics, _multilineFirstVisibleLine);
        DrawMultilineSelection(
            host, displayText, textRect, FontRole::Body, style.text, style.selectionFill, style.selectionText, multilineScrollDip, GetSelectionRange());
    }
    else if (usePlaceholder)
    {
        DrawSingleLineTextClipped(host, std::wstring_view(_placeholder), textRect, FontRole::Body, style.placeholderText, 0.0f);
    }
    else
    {
        EnsureCaretVisible(&host, std::max(1.0f, textRect.right - textRect.left));
        DrawSingleLineSelection(
            host, displayText, textRect, FontRole::Body, style.text, style.selectionFill, style.selectionText, _horizontalScrollDip, GetSelectionRange());
    }

    if (HasFocus() && _caretVisible)
    {
        D2D1_RECT_F caretRect = D2D1::RectF();
        if (_multiline)
        {
            const auto multilineLayout =
                GetOrCreateMultilineLayout(&host, displayText, std::max(1.0f, textRect.right - textRect.left), std::max(1.0f, textRect.bottom - textRect.top));
            const std::vector<DWRITE_LINE_METRICS> lineMetrics = GetMultilineLineMetrics(multilineLayout.get());
            const float multilineScrollDip                     = MeasureWrappedLineOffsetDip(lineMetrics, _multilineFirstVisibleLine);
            caretRect = MeasureMultilineCaretRectDip(&host, displayText, FontRole::Body, textRect, multilineScrollDip, _caretIndex);
        }
        else
        {
            EnsureCaretVisible(&host, std::max(1.0f, textRect.right - textRect.left));
            const float caretOffset = MeasureCaretOffsetDip(&host, displayText, FontRole::Body, _caretIndex, std::max(1.0f, textRect.bottom - textRect.top));
            const float caretX      = std::clamp(textRect.left + caretOffset - _horizontalScrollDip, textRect.left, textRect.right - 1.0f);
            caretRect               = D2D1::RectF(caretX, textRect.top + 2.0f, caretX + 1.0f, textRect.bottom - 2.0f);
        }
        if (auto* dc = host.GetDeviceContext())
        {
            if (auto* brush = host.GetSolidBrush(style.caret))
            {
                const D2D1_RECT_F snappedCaretRect = SnapRectToPixel(host, caretRect);
                dc->DrawLine(
                    D2D1::Point2F(snappedCaretRect.left, snappedCaretRect.top), D2D1::Point2F(snappedCaretRect.left, snappedCaretRect.bottom), brush, 1.0f);
            }
        }
    }

    // Clear button: X icon when editable, single-line, has text, and focused
    if (IsClearButtonVisible())
    {
        const D2D1_RECT_F clearRect = GetClearButtonRect();
        if (_clearButtonHovered)
        {
            const D2D1_ROUNDED_RECT hoverBg =
                D2D1::RoundedRect(D2D1::RectF(clearRect.left + 4.0f, clearRect.top + 4.0f, clearRect.right - 4.0f, clearRect.bottom - 4.0f), 4.0f, 4.0f);
            if (auto* dc = host.GetDeviceContext())
            {
                dc->FillRoundedRectangle(&hoverBg, host.GetSolidBrush(host.GetTheme().hoverFill));
            }
        }
        DrawCenteredText(host, L"\xE711", clearRect, FontRole::Icon, style.text);
    }
}

bool TextField::Tick(WindowHost& /*host*/, uint64_t nowTickMs)
{
    if (! HasFocus())
    {
        _caretVisible           = true;
        _caretBlinkAnchorTickMs = 0u;
        return false;
    }

    if (_caretBlinkAnchorTickMs == 0u)
    {
        _caretBlinkAnchorTickMs = nowTickMs;
        _caretVisible           = true;
    }
    else
    {
        _caretVisible = (((nowTickMs - _caretBlinkAnchorTickMs) / kCaretBlinkPeriodMs) % 2u) == 0u;
    }

    return true;
}

void TextField::OnFocusChanged(WindowHost& host, bool focused)
{
    Control::OnFocusChanged(host, focused);
    if (focused)
    {
        if (_multiline)
        {
            if (! _readOnly)
            {
                EnsureMultilineCaretVisible(&host);
            }
        }
        ResetCaretBlink(host);
    }
    else
    {
        _caretBlinkAnchorTickMs = 0u;
        _caretVisible           = true;
        _dragSelecting          = false;
        ResetSingleLineSelectionClickSequence(_selectionClickSequence);
        _preferredMultilineXOffsetDip.reset();
        _multilineWheelDeltaRemainder = 0.0f;
        if (_onBlur)
        {
            _onBlur();
        }
        Invalidate(host);
    }
}

bool TextField::OnMouseDown(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT modifiers)
{
    if (rightButton)
    {
        ResetSingleLineSelectionClickSequence(_selectionClickSequence);
        return OnContextMenu(host, false, point);
    }

    // Clear button click — clear text and keep focus
    if (IsClearButtonVisible() && PointInRect(GetClearButtonRect(), point))
    {
        ResetSingleLineSelectionClickSequence(_selectionClickSequence);
        SetTextAndNotify({});
        ResetCaretBlink(host);
        Invalidate(host);
        return true;
    }

    host.SetFocusControl(this);
    if (! _multiline && ! ModifiersContainShift(modifiers) && ShouldPromoteSingleLineClickToSelectAll(host, _selectionClickSequence, point))
    {
        ResetSingleLineSelectionClickSequence(_selectionClickSequence);
        SelectAllText();
        _dragSelecting             = false;
        const D2D1_RECT_F textRect = GetTextRect();
        EnsureCaretVisible(&host, std::max(1.0f, textRect.right - textRect.left));
        ResetCaretBlink(host);
        host.SyncTextInputBridge(this);
        Invalidate(host);
        return true;
    }

    ResetSingleLineSelectionClickSequence(_selectionClickSequence);
    if (_multiline)
    {
        const D2D1_RECT_F textRect = GetTextRect();
        const auto multilineLayout = CreateMultilineTextLayout(
            &host, _text, FontRole::Body, std::max(1.0f, textRect.right - textRect.left), std::max(1.0f, textRect.bottom - textRect.top));
        const size_t hitIndex =
            HitTestMultilineCaretIndexDip(&host,
                                          _text,
                                          FontRole::Body,
                                          textRect,
                                          MeasureWrappedLineOffsetDip(GetMultilineLineMetrics(multilineLayout.get()), _multilineFirstVisibleLine),
                                          point);
        _preferredMultilineXOffsetDip.reset();
        if (ModifiersContainShift(modifiers))
        {
            SetCaretIndex(hitIndex, true);
        }
        else
        {
            _selectionAnchorIndex = hitIndex;
            _caretIndex           = hitIndex;
        }
        _dragSelecting = true;
        EnsureMultilineCaretVisible(&host);
    }
    else
    {
        const D2D1_RECT_F textRect     = GetTextRect();
        const std::wstring displayText = GetDisplayText();
        const size_t hitIndex          = HitTestCaretIndexDip(&host, displayText, FontRole::Body, textRect, _horizontalScrollDip, point);
        _preferredMultilineXOffsetDip.reset();
        if (ModifiersContainShift(modifiers))
        {
            SetCaretIndex(hitIndex, true);
        }
        else
        {
            _selectionAnchorIndex = hitIndex;
            _caretIndex           = hitIndex;
        }
        _dragSelecting = true;
        EnsureCaretVisible(&host, std::max(1.0f, textRect.right - textRect.left));
    }
    ResetCaretBlink(host);
    host.SyncTextInputBridge(this);
    Invalidate(host);
    return true;
}

bool TextField::OnMouseDoubleClick(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT modifiers)
{
    if (rightButton)
    {
        return OnMouseDown(host, point, rightButton, modifiers);
    }

    host.SetFocusControl(this);
    const D2D1_RECT_F textRect     = GetTextRect();
    const std::wstring displayText = GetDisplayText();
    size_t hitIndex                = 0u;
    if (_multiline)
    {
        const auto multilineLayout = CreateMultilineTextLayout(
            &host, displayText, FontRole::Body, std::max(1.0f, textRect.right - textRect.left), std::max(1.0f, textRect.bottom - textRect.top));
        hitIndex = HitTestMultilineCaretIndexDip(&host,
                                                 displayText,
                                                 FontRole::Body,
                                                 textRect,
                                                 MeasureWrappedLineOffsetDip(GetMultilineLineMetrics(multilineLayout.get()), _multilineFirstVisibleLine),
                                                 point);
    }
    else
    {
        hitIndex = HitTestCaretIndexDip(&host, displayText, FontRole::Body, textRect, _horizontalScrollDip, point);
    }
    _preferredMultilineXOffsetDip.reset();
    SelectWordAt(hitIndex);
    ResetSingleLineSelectionClickSequence(_selectionClickSequence);
    if (! _multiline)
    {
        ArmSingleLineSelectionClickSequence(_selectionClickSequence, point);
    }
    _dragSelecting = false;
    if (_multiline)
    {
        EnsureMultilineCaretVisible(&host);
    }
    else
    {
        EnsureCaretVisible(&host, std::max(1.0f, textRect.right - textRect.left));
    }
    ResetCaretBlink(host);
    host.SyncTextInputBridge(this);
    Invalidate(host);
    return true;
}

bool TextField::OnMouseMove(WindowHost& host, D2D1_POINT_2F point, UINT /*modifiers*/)
{
    // Track clear button hover state
    const bool wasClearHovered = _clearButtonHovered;
    _clearButtonHovered        = IsClearButtonVisible() && PointInRect(GetClearButtonRect(), point);
    if (_clearButtonHovered != wasClearHovered)
    {
        Invalidate(host);
    }

    if (! _dragSelecting)
    {
        return _clearButtonHovered;
    }

    const D2D1_RECT_F textRect     = GetTextRect();
    const std::wstring displayText = GetDisplayText();
    size_t hitIndex                = 0u;
    if (_multiline)
    {
        const auto multilineLayout = CreateMultilineTextLayout(
            &host, displayText, FontRole::Body, std::max(1.0f, textRect.right - textRect.left), std::max(1.0f, textRect.bottom - textRect.top));
        hitIndex = HitTestMultilineCaretIndexDip(&host,
                                                 displayText,
                                                 FontRole::Body,
                                                 textRect,
                                                 MeasureWrappedLineOffsetDip(GetMultilineLineMetrics(multilineLayout.get()), _multilineFirstVisibleLine),
                                                 point);
    }
    else
    {
        hitIndex = HitTestCaretIndexDip(&host, displayText, FontRole::Body, textRect, _horizontalScrollDip, point);
    }
    _preferredMultilineXOffsetDip.reset();
    SetCaretIndex(hitIndex, true);
    if (_multiline)
    {
        EnsureMultilineCaretVisible(&host);
    }
    else
    {
        EnsureCaretVisible(&host, std::max(1.0f, textRect.right - textRect.left));
    }
    ResetCaretBlink(host);
    host.SyncTextInputBridge(this);
    Invalidate(host);
    return true;
}

bool TextField::OnMouseUp(WindowHost& host, D2D1_POINT_2F /*point*/, bool rightButton, UINT /*modifiers*/)
{
    if (rightButton)
    {
        return false;
    }

    const bool wasDragging = _dragSelecting;
    _dragSelecting         = false;
    if (wasDragging)
    {
        Invalidate(host);
    }
    return wasDragging;
}

bool TextField::OnMouseWheel(WindowHost& host, D2D1_POINT_2F /*point*/, float wheelDelta, UINT /*modifiers*/)
{
    if (! _multiline)
    {
        return false;
    }

    const D2D1_RECT_F textRect     = GetTextRect();
    const float viewportHeightDip  = std::max(1.0f, textRect.bottom - textRect.top);
    const std::wstring displayText = GetDisplayText();
    const auto multilineLayout =
        CreateMultilineTextLayout(&host, displayText, FontRole::Body, std::max(1.0f, textRect.right - textRect.left), viewportHeightDip);
    const std::vector<DWRITE_LINE_METRICS> lineMetrics = GetMultilineLineMetrics(multilineLayout.get());
    const float fallbackLineHeightDip                  = EstimateMultilineFallbackLineHeightDip(&host, FontRole::Body);
    const MultilineViewportMetrics viewportMetrics     = BuildMultilineViewportMetrics(displayText, viewportHeightDip, lineMetrics, fallbackLineHeightDip);
    const size_t wheelLineCount                        = ComputeMultilineWheelLineCount(viewportMetrics.visibleLineCount);
    if (wheelLineCount == 0u)
    {
        return true;
    }

    if (viewportMetrics.totalLineCount <= viewportMetrics.visibleLineCount)
    {
        _multilineFirstVisibleLine    = 0u;
        _multilineWheelDeltaRemainder = 0.0f;
        if (host.GetFocusControl() == this)
        {
            host.SyncTextInputBridge(this);
        }
        return true;
    }

    _multilineWheelDeltaRemainder += wheelDelta;
    const int wheelStepCount = static_cast<int>(_multilineWheelDeltaRemainder / static_cast<float>(WHEEL_DELTA));
    if (wheelStepCount == 0)
    {
        return true;
    }
    _multilineWheelDeltaRemainder -= static_cast<float>(wheelStepCount * WHEEL_DELTA);

    const int direction              = wheelStepCount > 0 ? -1 : 1;
    const size_t deltaLines          = static_cast<size_t>(std::abs(wheelStepCount)) * (std::max)(static_cast<size_t>(1u), wheelLineCount);
    const size_t maxFirstVisibleLine = ComputeMultilineMaxFirstVisibleLine(viewportMetrics);
    if (_multilineFirstVisibleLine > maxFirstVisibleLine)
    {
        _multilineFirstVisibleLine = maxFirstVisibleLine;
    }
    const int64_t nextFirstVisibleLine =
        std::clamp<int64_t>(static_cast<int64_t>(_multilineFirstVisibleLine) + (static_cast<int64_t>(direction) * static_cast<int64_t>(deltaLines)),
                            0ll,
                            static_cast<int64_t>(maxFirstVisibleLine));

    if (_multilineFirstVisibleLine == static_cast<size_t>(nextFirstVisibleLine))
    {
        return true;
    }

    _multilineFirstVisibleLine = static_cast<size_t>(nextFirstVisibleLine);
    if (host.GetFocusControl() == this)
    {
        host.SyncTextInputBridge(this);
    }
    Invalidate(host);
    return true;
}

bool TextField::OnKeyDown(WindowHost& host, UINT virtualKey, UINT modifiers)
{
    ResetSingleLineSelectionClickSequence(_selectionClickSequence);
    const auto textRect     = GetTextRect();
    const auto refreshCaret = [this, &host, &textRect]() noexcept
    {
        ResetCaretBlink(host);
        if (_multiline)
        {
            EnsureMultilineCaretVisible(&host);
        }
        else
        {
            EnsureCaretVisible(&host, std::max(1.0f, textRect.right - textRect.left));
        }
        Invalidate(host);
    };

    if (ModifiersContainCtrl(modifiers))
    {
        if (virtualKey == 'Z' && ! ModifiersContainShift(modifiers) && ! _readOnly)
        {
            if (TryUndoDirectEdit())
            {
                NotifyChanged();
                refreshCaret();
                host.SyncTextInputBridge(this);
                return true;
            }
            return false;
        }
        if (virtualKey == 'Y' && ! ModifiersContainShift(modifiers) && ! _readOnly)
        {
            if (TryRedoDirectEdit())
            {
                NotifyChanged();
                refreshCaret();
                host.SyncTextInputBridge(this);
                return true;
            }
            return false;
        }
        if (virtualKey == 'A')
        {
            SelectAllText();
            refreshCaret();
            return true;
        }
        if (virtualKey == 'C')
        {
            return OnCopy(host);
        }
        if (virtualKey == 'X' && ! _readOnly)
        {
            if (OnCopy(host))
            {
                RecordUndoStateForDirectEdit();
                _preferredMultilineXOffsetDip.reset();
                if (DeleteSelection())
                {
                    NotifyChanged();
                    refreshCaret();
                }
                return true;
            }
            return false;
        }
        if (virtualKey == VK_INSERT)
        {
            return OnCopy(host);
        }
        if (virtualKey == 'V' && ! _readOnly)
        {
            const auto clipboardText = host.ReadTextFromClipboard();
            if (clipboardText)
            {
                const std::wstring normalizedClipboardText = NormalizePastedControlText(clipboardText.value(), _multiline);
                const bool willMutate                      = HasSelection() || ! normalizedClipboardText.empty();
                if (willMutate)
                {
                    RecordUndoStateForDirectEdit();
                    static_cast<void>(DeleteSelection());
                    _text.insert(_caretIndex, normalizedClipboardText);
                    _caretIndex += normalizedClipboardText.size();
                    _selectionAnchorIndex.reset();
                    NotifyChanged();
                    refreshCaret();
                }
                return true;
            }
        }
        if (virtualKey == VK_BACK && ! _readOnly)
        {
            const size_t eraseFrom = FindPreviousWordBoundary(_text, _caretIndex);
            if (HasSelection() || eraseFrom < _caretIndex)
            {
                RecordUndoStateForDirectEdit();
                _preferredMultilineXOffsetDip.reset();
            }
            if (DeleteSelection())
            {
                NotifyChanged();
                refreshCaret();
                return true;
            }
            if (eraseFrom < _caretIndex)
            {
                _text.erase(eraseFrom, _caretIndex - eraseFrom);
                _caretIndex = eraseFrom;
                _selectionAnchorIndex.reset();
                NotifyChanged();
                refreshCaret();
                return true;
            }
            return true;
        }
        if (virtualKey == VK_DELETE && ! _readOnly)
        {
            const size_t eraseTo = FindNextWordBoundary(_text, _caretIndex);
            if (HasSelection() || eraseTo > _caretIndex)
            {
                RecordUndoStateForDirectEdit();
                _preferredMultilineXOffsetDip.reset();
            }
            if (DeleteSelection())
            {
                NotifyChanged();
                refreshCaret();
                return true;
            }
            if (eraseTo > _caretIndex)
            {
                _text.erase(_caretIndex, eraseTo - _caretIndex);
                _selectionAnchorIndex.reset();
                NotifyChanged();
                refreshCaret();
                return true;
            }
            return true;
        }
    }

    if (virtualKey == VK_INSERT && ModifiersContainShift(modifiers) && ! _readOnly)
    {
        const auto clipboardText = host.ReadTextFromClipboard();
        if (clipboardText)
        {
            const std::wstring normalizedClipboardText = NormalizePastedControlText(clipboardText.value(), _multiline);
            const bool willMutate                      = HasSelection() || ! normalizedClipboardText.empty();
            if (willMutate)
            {
                RecordUndoStateForDirectEdit();
                static_cast<void>(DeleteSelection());
                _text.insert(_caretIndex, normalizedClipboardText);
                _caretIndex += normalizedClipboardText.size();
                _selectionAnchorIndex.reset();
                NotifyChanged();
                refreshCaret();
            }
            return true;
        }
    }
    if (virtualKey == VK_DELETE && ModifiersContainShift(modifiers) && ! _readOnly)
    {
        if (OnCopy(host))
        {
            RecordUndoStateForDirectEdit();
            _preferredMultilineXOffsetDip.reset();
            if (DeleteSelection())
            {
                NotifyChanged();
                refreshCaret();
            }
            return true;
        }
        return false;
    }

    if (virtualKey == VK_LEFT)
    {
        _preferredMultilineXOffsetDip.reset();
        const bool extendSelection = ModifiersContainShift(modifiers);
        if (ModifiersContainCtrl(modifiers))
        {
            SetCaretIndex(FindPreviousWordBoundary(_text, _caretIndex), extendSelection);
            refreshCaret();
            return true;
        }
        if (! extendSelection && HasSelection())
        {
            _caretIndex = GetSelectionRange().value().first;
            _selectionAnchorIndex.reset();
            refreshCaret();
            return true;
        }
        if (_caretIndex > 0u)
        {
            SetCaretIndex(StepToPreviousCodePoint(_text, _caretIndex), extendSelection);
            refreshCaret();
        }
        return true;
    }
    if (virtualKey == VK_RIGHT)
    {
        _preferredMultilineXOffsetDip.reset();
        const bool extendSelection = ModifiersContainShift(modifiers);
        if (ModifiersContainCtrl(modifiers))
        {
            SetCaretIndex(FindNextWordBoundary(_text, _caretIndex), extendSelection);
            refreshCaret();
            return true;
        }
        if (! extendSelection && HasSelection())
        {
            _caretIndex = GetSelectionRange().value().second;
            _selectionAnchorIndex.reset();
            refreshCaret();
            return true;
        }
        SetCaretIndex(StepToNextCodePoint(_text, _caretIndex), extendSelection);
        refreshCaret();
        return true;
    }
    if (_multiline && virtualKey == VK_UP)
    {
        const bool extendSelection = ModifiersContainShift(modifiers);
        const size_t nextCaretIndex =
            MoveMultilineCaretVertically(&host, _text, FontRole::Body, GetTextRect(), _caretIndex, false, _preferredMultilineXOffsetDip);
        SetSingleLineCaretIndex(_caretIndex, _selectionAnchorIndex, nextCaretIndex, extendSelection);
        refreshCaret();
        return true;
    }
    if (_multiline && virtualKey == VK_DOWN)
    {
        const bool extendSelection = ModifiersContainShift(modifiers);
        const size_t nextCaretIndex =
            MoveMultilineCaretVertically(&host, _text, FontRole::Body, GetTextRect(), _caretIndex, true, _preferredMultilineXOffsetDip);
        SetSingleLineCaretIndex(_caretIndex, _selectionAnchorIndex, nextCaretIndex, extendSelection);
        refreshCaret();
        return true;
    }
    if (_multiline && virtualKey == VK_PRIOR)
    {
        const bool extendSelection                     = ModifiersContainShift(modifiers);
        const D2D1_RECT_F pageTextRect                 = GetTextRect();
        const MultilineViewportMetrics viewportMetrics = ComputeMultilineViewportMetrics(&host, _text, FontRole::Body, pageTextRect);
        const size_t nextCaretIndex                    = MoveMultilineCaretByPage(
            &host, _text, FontRole::Body, pageTextRect, _caretIndex, false, _preferredMultilineXOffsetDip, viewportMetrics.visibleLineCount);
        SetSingleLineCaretIndex(_caretIndex, _selectionAnchorIndex, nextCaretIndex, extendSelection);
        refreshCaret();
        return true;
    }
    if (_multiline && virtualKey == VK_NEXT)
    {
        const bool extendSelection                     = ModifiersContainShift(modifiers);
        const D2D1_RECT_F pageTextRect                 = GetTextRect();
        const MultilineViewportMetrics viewportMetrics = ComputeMultilineViewportMetrics(&host, _text, FontRole::Body, pageTextRect);
        const size_t nextCaretIndex                    = MoveMultilineCaretByPage(
            &host, _text, FontRole::Body, pageTextRect, _caretIndex, true, _preferredMultilineXOffsetDip, viewportMetrics.visibleLineCount);
        SetSingleLineCaretIndex(_caretIndex, _selectionAnchorIndex, nextCaretIndex, extendSelection);
        refreshCaret();
        return true;
    }
    if (virtualKey == VK_HOME)
    {
        _preferredMultilineXOffsetDip.reset();
        if (_multiline && ! ModifiersContainCtrl(modifiers))
        {
            const size_t nextCaretIndex =
                TryMoveCaretToWrappedLineBoundary(&host, _text, FontRole::Body, GetTextRect(), _caretIndex, false).value_or(FindLineStart(_text, _caretIndex));
            SetCaretIndex(nextCaretIndex, ModifiersContainShift(modifiers));
        }
        else
        {
            SetCaretIndex(0u, ModifiersContainShift(modifiers));
        }
        refreshCaret();
        return true;
    }
    if (virtualKey == VK_END)
    {
        _preferredMultilineXOffsetDip.reset();
        if (_multiline && ! ModifiersContainCtrl(modifiers))
        {
            const size_t nextCaretIndex =
                TryMoveCaretToWrappedLineBoundary(&host, _text, FontRole::Body, GetTextRect(), _caretIndex, true).value_or(FindLineEnd(_text, _caretIndex));
            SetCaretIndex(nextCaretIndex, ModifiersContainShift(modifiers));
        }
        else
        {
            SetCaretIndex(_text.size(), ModifiersContainShift(modifiers));
        }
        refreshCaret();
        return true;
    }
    if (virtualKey == VK_RETURN && _multiline)
    {
        return true;
    }
    if (virtualKey == VK_RETURN && ! _multiline)
    {
        if (_onSubmitted)
        {
            _onSubmitted();
            return true;
        }
        return false;
    }
    if (virtualKey == VK_BACK && ! _readOnly)
    {
        if (HasSelection() || (_caretIndex > 0 && ! _text.empty()))
        {
            RecordUndoStateForDirectEdit();
            _preferredMultilineXOffsetDip.reset();
        }
        if (DeleteSelection())
        {
            NotifyChanged();
            refreshCaret();
            return true;
        }
        if (_caretIndex > 0 && ! _text.empty())
        {
            const size_t eraseFrom = StepToPreviousCodePoint(_text, _caretIndex);
            _text.erase(eraseFrom, _caretIndex - eraseFrom);
            _caretIndex = eraseFrom;
            _selectionAnchorIndex.reset();
            NotifyChanged();
            refreshCaret();
        }
        return true;
    }
    if (virtualKey == VK_DELETE && ! _readOnly)
    {
        if (HasSelection() || _caretIndex < _text.size())
        {
            RecordUndoStateForDirectEdit();
            _preferredMultilineXOffsetDip.reset();
        }
        if (DeleteSelection())
        {
            NotifyChanged();
            refreshCaret();
            return true;
        }
        if (_caretIndex < _text.size())
        {
            const size_t eraseTo = StepToNextCodePoint(_text, _caretIndex);
            _text.erase(_caretIndex, eraseTo - _caretIndex);
            _selectionAnchorIndex.reset();
            NotifyChanged();
            refreshCaret();
        }
        return true;
    }
    return false;
}

bool TextField::OnChar(WindowHost& host, wchar_t ch, UINT /*modifiers*/)
{
    ResetSingleLineSelectionClickSequence(_selectionClickSequence);
    if (_readOnly)
    {
        return false;
    }

    if (ch == L'\r')
    {
        if (! _multiline)
        {
            return true;
        }
        RecordUndoStateForDirectEdit();
        _preferredMultilineXOffsetDip.reset();
        static_cast<void>(DeleteSelection());
        _text.insert(_caretIndex, 1u, L'\n');
        _caretIndex += 1u;
        _selectionAnchorIndex.reset();
    }
    else if (std::iswcntrl(static_cast<wint_t>(ch)) == 0 && (_multiline || ch != L'\t'))
    {
        RecordUndoStateForDirectEdit();
        _preferredMultilineXOffsetDip.reset();
        static_cast<void>(DeleteSelection());
        _text.insert(_caretIndex, 1u, ch);
        _caretIndex += 1u;
        _selectionAnchorIndex.reset();
    }
    else
    {
        return false;
    }

    NotifyChanged();
    ResetCaretBlink(host);
    if (_multiline)
    {
        EnsureMultilineCaretVisible(&host);
    }
    else
    {
        EnsureCaretVisible(&host, std::max(1.0f, GetTextRect().right - GetTextRect().left));
    }
    Invalidate(host);
    return true;
}

bool TextField::OnContextMenu(WindowHost& host, bool keyboardInvocation, D2D1_POINT_2F pointDip)
{
    ResetSingleLineSelectionClickSequence(_selectionClickSequence);
    host.SetFocusControl(this);
    host.SyncTextInputBridge(this);
    ResetCaretBlink(host);
    Invalidate(host);
    return Control::OnContextMenu(host, keyboardInvocation, pointDip);
}

bool TextField::OnCopy(WindowHost& host)
{
    if (_masked)
    {
        return false;
    }

    if (const std::optional<std::pair<size_t, size_t>> selectionRange = GetSelectionRange())
    {
        const auto [selectionStart, selectionEnd] = selectionRange.value();
        return host.CopyTextToClipboard(_text.substr(selectionStart, selectionEnd - selectionStart));
    }

    return false;
}

bool TextField::OnSelectAll(WindowHost& host)
{
    ResetSingleLineSelectionClickSequence(_selectionClickSequence);
    _preferredMultilineXOffsetDip.reset();
    SelectAllText();
    ResetCaretBlink(host);
    if (_multiline)
    {
        EnsureMultilineCaretVisible(&host);
    }
    else
    {
        EnsureCaretVisible(&host, std::max(1.0f, GetTextRect().right - GetTextRect().left));
    }
    host.SyncTextInputBridge(this);
    Invalidate(host);
    return true;
}

bool TextField::SupportsTextInputBridge() const noexcept
{
    return true;
}

std::optional<D2D1_RECT_F> TextField::GetTextInputBridgeViewportRect() const noexcept
{
    return GetTextRect();
}

std::optional<D2D1_RECT_F> TextField::GetTextInputBridgeCaretRect(const WindowHost& host, size_t controlTextIndex) const noexcept
{
    const D2D1_RECT_F textRect     = GetTextRect();
    const std::wstring displayText = GetDisplayText();
    const size_t clampedCaretIndex = (std::min)(controlTextIndex, displayText.size());

    if (_multiline)
    {
        const auto multilineLayout =
            GetOrCreateMultilineLayout(&host, displayText, std::max(1.0f, textRect.right - textRect.left), std::max(1.0f, textRect.bottom - textRect.top));
        const std::vector<DWRITE_LINE_METRICS> lineMetrics = GetMultilineLineMetrics(multilineLayout.get());
        const float multilineScrollDip                     = MeasureWrappedLineOffsetDip(lineMetrics, _multilineFirstVisibleLine);
        return MeasureMultilineCaretRectDip(&host, displayText, FontRole::Body, textRect, multilineScrollDip, clampedCaretIndex);
    }

    const float caretOffset = MeasureCaretOffsetDip(&host, displayText, FontRole::Body, clampedCaretIndex, std::max(1.0f, textRect.bottom - textRect.top));
    const float caretX      = std::clamp(textRect.left + caretOffset - _horizontalScrollDip, textRect.left, textRect.right - 1.0f);
    return D2D1::RectF(caretX, textRect.top + 2.0f, caretX + 1.0f, textRect.bottom - 2.0f);
}

bool TextField::ExportTextInputBridgeState(TextInputBridgeState& outState) const
{
    outState.text                 = _text;
    outState.selectionAnchorIndex = _selectionAnchorIndex;
    outState.caretIndex           = _caretIndex;
    outState.firstVisibleLine     = _multilineFirstVisibleLine;
    outState.readOnly             = _readOnly;
    outState.masked               = _masked;
    outState.multiline            = _multiline;
    return true;
}

bool TextField::ImportTextInputBridgeState(WindowHost& host, const TextInputBridgeState& state, bool notifyChange)
{
    const std::wstring previousText       = _text;
    const size_t previousFirstVisibleLine = _multilineFirstVisibleLine;
    _undoHistory.clear();
    _redoHistory.clear();
    ResetSingleLineSelectionClickSequence(_selectionClickSequence);
    _text       = state.text;
    _caretIndex = std::min(state.caretIndex, _text.size());
    if (state.selectionAnchorIndex)
    {
        _selectionAnchorIndex = std::min(state.selectionAnchorIndex.value(), _text.size());
        if (_selectionAnchorIndex.value() == _caretIndex)
        {
            _selectionAnchorIndex.reset();
        }
    }
    else
    {
        _selectionAnchorIndex.reset();
    }

    const bool preserveReadOnlyMultilineViewport = _multiline && _readOnly && state.multiline;
    _multilineFirstVisibleLine = state.multiline ? (preserveReadOnlyMultilineViewport ? previousFirstVisibleLine : state.firstVisibleLine) : 0u;
    _preferredMultilineXOffsetDip.reset();
    _multilineWheelDeltaRemainder = 0.0f;
    _dragSelecting                = false;
    ResetCaretBlink(host);
    if (_multiline)
    {
        if (preserveReadOnlyMultilineViewport)
        {
            if (previousText != _text)
            {
                const D2D1_RECT_F textRect     = GetTextRect();
                const std::wstring displayText = GetDisplayText();
                if (displayText.empty())
                {
                    _multilineFirstVisibleLine = 0u;
                }
                else
                {
                    const float viewportHeightDip = std::max(1.0f, textRect.bottom - textRect.top);
                    const auto multilineLayout =
                        CreateMultilineTextLayout(&host, displayText, FontRole::Body, std::max(1.0f, textRect.right - textRect.left), viewportHeightDip);
                    const std::vector<DWRITE_LINE_METRICS> lineMetrics = GetMultilineLineMetrics(multilineLayout.get());
                    const float fallbackLineHeightDip                  = EstimateMultilineFallbackLineHeightDip(&host, FontRole::Body);
                    const MultilineViewportMetrics viewportMetrics =
                        BuildMultilineViewportMetrics(displayText, viewportHeightDip, lineMetrics, fallbackLineHeightDip);
                    _multilineFirstVisibleLine = ClampMultilineFirstVisibleLine(_multilineFirstVisibleLine, viewportMetrics);
                }
            }
        }
        else
        {
            EnsureMultilineCaretVisible(&host);
        }
    }
    else
    {
        EnsureCaretVisible(&host, std::max(1.0f, GetTextRect().right - GetTextRect().left));
    }
    Invalidate(host);
    if (notifyChange && previousText != _text)
    {
        NotifyChanged();
        return true;
    }
    return true;
}

TextField::EditHistoryState TextField::CaptureEditHistoryState() const
{
    return EditHistoryState{
        .text = _text, .caretIndex = _caretIndex, .selectionAnchorIndex = _selectionAnchorIndex, .firstVisibleLine = _multilineFirstVisibleLine};
}

void TextField::RestoreEditHistoryState(const EditHistoryState& state) noexcept
{
    _text       = state.text;
    _caretIndex = std::min(state.caretIndex, _text.size());
    if (state.selectionAnchorIndex)
    {
        _selectionAnchorIndex = std::min(state.selectionAnchorIndex.value(), _text.size());
        if (_selectionAnchorIndex.value() == _caretIndex)
        {
            _selectionAnchorIndex.reset();
        }
    }
    else
    {
        _selectionAnchorIndex.reset();
    }
    _multilineFirstVisibleLine = _multiline ? state.firstVisibleLine : 0u;
    _preferredMultilineXOffsetDip.reset();
    _multilineWheelDeltaRemainder = 0.0f;
    _dragSelecting                = false;
    _horizontalScrollDip          = 0.0f;
    InvalidateMultilineLayoutCache();
}

void TextField::RecordUndoStateForDirectEdit()
{
    _undoHistory.push_back(CaptureEditHistoryState());
    if (_undoHistory.size() > kMaxEditHistoryEntries)
    {
        _undoHistory.erase(_undoHistory.begin());
    }
    _redoHistory.clear();
}

bool TextField::TryUndoDirectEdit() noexcept
{
    if (_undoHistory.empty())
    {
        return false;
    }

    _redoHistory.push_back(CaptureEditHistoryState());
    if (_redoHistory.size() > kMaxEditHistoryEntries)
    {
        _redoHistory.erase(_redoHistory.begin());
    }

    const EditHistoryState state = std::move(_undoHistory.back());
    _undoHistory.pop_back();
    RestoreEditHistoryState(state);
    return true;
}

bool TextField::TryRedoDirectEdit() noexcept
{
    if (_redoHistory.empty())
    {
        return false;
    }

    _undoHistory.push_back(CaptureEditHistoryState());
    if (_undoHistory.size() > kMaxEditHistoryEntries)
    {
        _undoHistory.erase(_undoHistory.begin());
    }

    const EditHistoryState state = std::move(_redoHistory.back());
    _redoHistory.pop_back();
    RestoreEditHistoryState(state);
    return true;
}

void TextField::SetCaretIndex(size_t caretIndex, bool extendSelection) noexcept
{
    _preferredMultilineXOffsetDip.reset();
    SetSingleLineCaretIndex(_caretIndex, _selectionAnchorIndex, std::min(caretIndex, _text.size()), extendSelection);
}

bool TextField::HasSelection() const noexcept
{
    return GetSelectionRange().has_value();
}

std::optional<std::pair<size_t, size_t>> TextField::GetSelectionRange() const noexcept
{
    return GetSingleLineSelectionRange(_selectionAnchorIndex, _caretIndex);
}

bool TextField::DeleteSelection() noexcept
{
    return DeleteSingleLineSelection(_text, _caretIndex, _selectionAnchorIndex);
}

void TextField::SelectAllText() noexcept
{
    SelectAllSingleLineText(_text.size(), _caretIndex, _selectionAnchorIndex);
}

void TextField::SelectWordAt(size_t hitIndex) noexcept
{
    SelectSingleLineWordAt(_text, hitIndex, _caretIndex, _selectionAnchorIndex);
}

D2D1_RECT_F TextField::GetTextRect() const noexcept
{
    const float rightInset = IsClearButtonVisible() ? std::max(38.0f, _textPaddingRightDip) : _textPaddingRightDip; // 30 DIP button + padding
    return D2D1::RectF(GetBounds().left + _textPaddingLeftDip,
                       GetBounds().top + _textPaddingTopDip,
                       GetBounds().right - rightInset,
                       GetBounds().bottom - _textPaddingBottomDip);
}

bool TextField::IsClearButtonVisible() const noexcept
{
    return _clearButtonEnabled && ! _readOnly && ! _multiline && ! _text.empty() && HasFocus();
}

D2D1_RECT_F TextField::GetClearButtonRect() const noexcept
{
    constexpr float kClearButtonWidthDip = 30.0f;
    return D2D1::RectF(GetBounds().right - kClearButtonWidthDip, GetBounds().top, GetBounds().right, GetBounds().bottom);
}

std::wstring TextField::GetDisplayText() const
{
    if (! _masked || _text.empty())
    {
        return _text;
    }

    return std::wstring(_text.size(), L'\u2022');
}

wil::com_ptr<IDWriteTextLayout> TextField::GetOrCreateMultilineLayout(const WindowHost* host,
                                                                      std::wstring_view text,
                                                                      float widthDip,
                                                                      float heightDip) const noexcept
{
    if (! host || ! _multiline)
    {
        return CreateMultilineTextLayout(host, text, FontRole::Body, widthDip, heightDip);
    }

    const D2D1_SIZE_F desiredSize = D2D1::SizeF(widthDip, heightDip);
    const bool sizeChanged = std::abs(_cachedLayoutSize.width - desiredSize.width) > 0.5f || std::abs(_cachedLayoutSize.height - desiredSize.height) > 0.5f;
    const bool textChanged = text != _cachedLayoutText;

    if (_multilineLayoutDirty || sizeChanged || textChanged || ! _cachedMultilineLayout)
    {
        _cachedMultilineLayout = CreateMultilineTextLayout(host, text, FontRole::Body, widthDip, heightDip);
        _cachedLayoutText      = std::wstring(text);
        _cachedLayoutSize      = desiredSize;
        _multilineLayoutDirty  = false;
    }

    return _cachedMultilineLayout;
}

void TextField::InvalidateMultilineLayoutCache() const noexcept
{
    _multilineLayoutDirty = true;
    _cachedMultilineLayout.reset();
    _cachedLayoutText.clear();
}

void TextField::EnsureCaretVisible(const WindowHost* host, float availableWidthDip) const noexcept
{
    if (_multiline || _text.empty())
    {
        _horizontalScrollDip = 0.0f;
        return;
    }

    const std::wstring displayText = GetDisplayText();
    const float caretOffset = MeasureCaretOffsetDip(host, displayText, FontRole::Body, _caretIndex, std::max(1.0f, GetTextRect().bottom - GetTextRect().top));
    const float padding     = 6.0f;
    if (caretOffset < _horizontalScrollDip + padding)
    {
        _horizontalScrollDip = std::max(0.0f, caretOffset - padding);
    }
    else if (caretOffset > _horizontalScrollDip + availableWidthDip - padding)
    {
        _horizontalScrollDip = std::max(0.0f, caretOffset - availableWidthDip + padding);
    }
}

void TextField::EnsureMultilineCaretVisible(const WindowHost* host) noexcept
{
    if (! _multiline)
    {
        _multilineFirstVisibleLine = 0u;
        return;
    }

    const D2D1_RECT_F textRect     = GetTextRect();
    const std::wstring displayText = GetDisplayText();
    if (displayText.empty())
    {
        _multilineFirstVisibleLine = 0u;
        return;
    }

    const float viewportHeightDip = std::max(1.0f, textRect.bottom - textRect.top);
    const auto multilineLayout =
        CreateMultilineTextLayout(host, displayText, FontRole::Body, std::max(1.0f, textRect.right - textRect.left), viewportHeightDip);
    const std::vector<DWRITE_LINE_METRICS> lineMetrics = GetMultilineLineMetrics(multilineLayout.get());
    const float fallbackLineHeightDip                  = EstimateMultilineFallbackLineHeightDip(host, FontRole::Body);
    const MultilineViewportMetrics viewportMetrics     = BuildMultilineViewportMetrics(displayText, viewportHeightDip, lineMetrics, fallbackLineHeightDip);
    const size_t clampedCaretIndex                     = (std::min)(_caretIndex, displayText.size());
    const size_t caretLine =
        TryGetMultilineCaretLineIndex(multilineLayout.get(), lineMetrics, clampedCaretIndex, displayText.size())
            .value_or(static_cast<size_t>(
                std::count(displayText.begin(), std::next(displayText.begin(), static_cast<std::wstring::difference_type>(clampedCaretIndex)), L'\n')));
    const size_t maxFirstVisibleLine = ComputeMultilineMaxFirstVisibleLine(viewportMetrics);
    size_t firstVisibleLine          = ClampMultilineFirstVisibleLine(_multilineFirstVisibleLine, viewportMetrics);
    if (caretLine < firstVisibleLine)
    {
        firstVisibleLine = caretLine;
    }
    else if (caretLine >= firstVisibleLine + viewportMetrics.visibleLineCount)
    {
        firstVisibleLine = caretLine - viewportMetrics.visibleLineCount + 1u;
    }

    _multilineFirstVisibleLine = (std::min)(firstVisibleLine, maxFirstVisibleLine);
}

void TextField::OnHostDpiChanged(WindowHost& host) noexcept
{
    Control::OnHostDpiChanged(host);
    InvalidateMultilineLayoutCache();
    if (SupportsTextInputBridge() && host.GetFocusControl() == this)
    {
        host.SyncTextInputBridge(this);
    }
}

void TextField::ResetCaretBlink(WindowHost& host) noexcept
{
    _caretBlinkAnchorTickMs = ::GetTickCount64();
    _caretVisible           = true;
    host.RequestAnimation();
}

void TextField::NotifyChanged() const
{
    if (_onTextChanged)
    {
        _onTextChanged(_text);
    }
}
} // namespace RedSalamander::DxUi
