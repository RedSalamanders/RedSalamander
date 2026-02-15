#pragma once

#include <algorithm>
#include <cmath>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#pragma warning(push)
// Windows headers: C4710 (not inlined), C4711 (auto inline), C4514 (unreferenced inline)
#pragma warning(disable : 4710 4711 4514)
#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <windows.h>

#include <d2d1.h>
#include <d2d1helper.h>
#include <dwrite.h>
#pragma warning(pop)

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027)
#include <wil/com.h>
#include <wil/resource.h>
#pragma warning(pop)

namespace RedSalamander::Ui
{
enum class AlertSeverity : uint8_t
{
    Error,
    Warning,
    Info,
    Busy,
};

struct AlertButton
{
    uint32_t id = 0;
    std::wstring label;
    bool primary = false;
};

struct AlertModel
{
    AlertSeverity severity = AlertSeverity::Error;
    std::wstring title;
    std::wstring message;
    bool closable = true;
    std::vector<AlertButton> buttons;
};

struct AlertTheme
{
    // Base surface colors.
    D2D1::ColorF background = D2D1::ColorF(D2D1::ColorF::White);
    D2D1::ColorF text       = D2D1::ColorF(D2D1::ColorF::Black);
    D2D1::ColorF accent     = D2D1::ColorF(0.0f, 0.478f, 1.0f, 1.0f);

    // Selection colors (used for hover / button emphasis).
    D2D1::ColorF selectionBackground = D2D1::ColorF(0.0f, 0.478f, 1.0f, 1.0f);
    D2D1::ColorF selectionText       = D2D1::ColorF(D2D1::ColorF::White);

    // Alert palettes (per severity).
    D2D1::ColorF errorBackground = D2D1::ColorF(1.0f, 0.95f, 0.95f);
    D2D1::ColorF errorText       = D2D1::ColorF(0.8f, 0.0f, 0.0f);

    D2D1::ColorF warningBackground = D2D1::ColorF(1.0f, 0.98f, 0.90f);
    D2D1::ColorF warningText       = D2D1::ColorF(0.65f, 0.38f, 0.0f);

    D2D1::ColorF infoBackground = D2D1::ColorF(0.90f, 0.95f, 1.0f);
    D2D1::ColorF infoText       = D2D1::ColorF(0.0f, 0.47f, 0.84f);

    bool darkBase     = false;
    bool highContrast = false;
};

struct AlertHitTest
{
    enum class Part : uint8_t
    {
        None,
        Close,
        Button,
    };

    Part part         = Part::None;
    uint32_t buttonId = 0;
};

inline D2D1::ColorF ColorFFromArgb(uint32_t argb) noexcept
{
    const uint8_t a = static_cast<uint8_t>((argb >> 24) & 0xFFu);
    const uint8_t r = static_cast<uint8_t>((argb >> 16) & 0xFFu);
    const uint8_t g = static_cast<uint8_t>((argb >> 8) & 0xFFu);
    const uint8_t b = static_cast<uint8_t>(argb & 0xFFu);

    const float af = static_cast<float>(a) / 255.0f;
    const float rf = static_cast<float>(r) / 255.0f;
    const float gf = static_cast<float>(g) / 255.0f;
    const float bf = static_cast<float>(b) / 255.0f;
    return D2D1::ColorF(rf, gf, bf, af);
}

class AlertOverlay final
{
private:
    enum class IconGlyphSet : uint8_t
    {
        None,
        Fluent,
        Unicode,
    };

public:
    void SetModel(AlertModel model)
    {
        _model = std::move(model);
        InvalidateTextLayouts();
        InvalidateButtonLayouts();
        ResetFocus();
    }

    const AlertModel& GetModel() const noexcept
    {
        return _model;
    }

    void SetTheme(const AlertTheme& theme) noexcept
    {
        _theme = theme;
    }

    const AlertTheme& GetTheme() const noexcept
    {
        return _theme;
    }

    void ResetDeviceResources() noexcept
    {
        _scrimBrush.reset();
        _backgroundBrush.reset();
        _textBrush.reset();
        _targetIdentity = nullptr;
    }

    void ResetTextResources() noexcept
    {
        _titleFormat.reset();
        _bodyFormat.reset();
        _buttonFormat.reset();
        _iconFormat.reset();
        _iconGlyphSet   = IconGlyphSet::None;
        _dwriteIdentity = nullptr;
        InvalidateTextLayouts();
        InvalidateButtonLayouts();
    }

    void SetStartTick(uint64_t tickMs) noexcept
    {
        _startTickMs = tickMs;
    }

    void ClearHotState() noexcept
    {
        _hot = {};
    }

    void ClearFocusedButton() noexcept
    {
        _focusedButtonId.reset();
    }

    [[nodiscard]] std::optional<uint32_t> GetFocusedButtonId() const noexcept
    {
        return _focusedButtonId;
    }

    [[nodiscard]] bool FocusNextButton(bool reverse) noexcept
    {
        if (_model.buttons.empty())
        {
            const bool changed = _focusedButtonId.has_value();
            _focusedButtonId.reset();
            return changed;
        }

        size_t nextIndex = 0;
        if (_focusedButtonId.has_value())
        {
            std::optional<size_t> currentIndex;
            for (size_t i = 0; i < _model.buttons.size(); ++i)
            {
                if (_model.buttons[i].id == _focusedButtonId.value())
                {
                    currentIndex = i;
                    break;
                }
            }

            if (currentIndex.has_value())
            {
                if (reverse)
                {
                    nextIndex = currentIndex.value() == 0 ? (_model.buttons.size() - 1u) : (currentIndex.value() - 1u);
                }
                else
                {
                    nextIndex = (currentIndex.value() + 1u) % _model.buttons.size();
                }
            }
            else
            {
                nextIndex = reverse ? (_model.buttons.size() - 1u) : 0u;
            }
        }
        else
        {
            nextIndex = reverse ? (_model.buttons.size() - 1u) : 0u;
        }

        const uint32_t nextId = _model.buttons[nextIndex].id;
        if (_focusedButtonId.has_value() && _focusedButtonId.value() == nextId)
        {
            return false;
        }

        _focusedButtonId = nextId;
        return true;
    }

    [[nodiscard]] bool HasLayout() const noexcept
    {
        return _hasLayout;
    }

    [[nodiscard]] D2D1_RECT_F GetPanelRect() const noexcept
    {
        return _panelRect;
    }

    [[nodiscard]] bool IsPointInPanel(D2D1_POINT_2F ptDip) const noexcept
    {
        if (! _hasLayout)
        {
            return false;
        }

        return ptDip.x >= _panelRect.left && ptDip.x <= _panelRect.right && ptDip.y >= _panelRect.top && ptDip.y <= _panelRect.bottom;
    }

    AlertHitTest HitTest(D2D1_POINT_2F ptDip) const noexcept
    {
        if (! _hasLayout)
        {
            return {};
        }

        if (_model.closable && PointInRect(ptDip, _closeRect))
        {
            AlertHitTest hit{};
            hit.part = AlertHitTest::Part::Close;
            return hit;
        }

        for (const auto& button : _buttonRects)
        {
            if (PointInRect(ptDip, button.rect))
            {
                AlertHitTest hit{};
                hit.part     = AlertHitTest::Part::Button;
                hit.buttonId = button.id;
                return hit;
            }
        }

        return {};
    }

    bool UpdateHotState(D2D1_POINT_2F ptDip) noexcept
    {
        const AlertHitTest hit = HitTest(ptDip);
        if (hit.part == _hot.part && hit.buttonId == _hot.buttonId)
        {
            return false;
        }

        _hot = hit;
        return true;
    }

    void Draw(ID2D1RenderTarget* target, IDWriteFactory* dwriteFactory, float clientWidthDip, float clientHeightDip, uint64_t nowTickMs) noexcept
    {
        if (! target || ! dwriteFactory)
        {
            return;
        }

        if (clientWidthDip <= 0.0f || clientHeightDip <= 0.0f)
        {
            return;
        }

        EnsureDeviceResources(target);
        EnsureTextResources(dwriteFactory);
        if (! _scrimBrush || ! _backgroundBrush || ! _textBrush || ! _titleFormat || ! _bodyFormat || ! _buttonFormat)
        {
            return;
        }

        constexpr float kOuterMarginDip       = 24.0f;
        constexpr float kInnerPaddingDip      = 20.0f;
        constexpr float kMaxWidthDip          = 780.0f;
        constexpr float kMaxHeightDip         = 420.0f;
        constexpr float kCornerRadiusDip      = 12.0f;
        constexpr float kIconSizeDip          = 80.0f;
        constexpr float kIconTextGapDip       = 18.0f;
        constexpr float kTitleBodyGapDip      = 6.0f;
        constexpr float kCardOpacity          = 0.96f;
        constexpr float kBorderOpacity        = 0.90f;
        constexpr float kCloseSizeDip         = 22.0f;
        constexpr float kCloseInsetDip        = 8.0f;
        constexpr float kButtonsGapDip        = 14.0f;
        constexpr float kButtonsRowGapDip     = 14.0f;
        constexpr float kButtonHeightDip      = 32.0f;
        constexpr float kButtonMinWidthDip    = 84.0f;
        constexpr float kButtonHorzPaddingDip = 14.0f;
        constexpr float kButtonCornerDip      = 6.0f;
        constexpr uint64_t kShowAnimationMs   = 220;

        const float minDimDip      = std::min(clientWidthDip, clientHeightDip);
        const float outerMarginDip = std::min(kOuterMarginDip, minDimDip * 0.06f);

        const float availableWidth  = std::max(0.0f, clientWidthDip - outerMarginDip * 2.0f);
        const float availableHeight = std::max(0.0f, clientHeightDip - outerMarginDip * 2.0f);

        const float panelWidth     = std::max(0.0f, std::min(kMaxWidthDip, availableWidth));
        const float maxPanelHeight = std::max(0.0f, std::min(kMaxHeightDip, availableHeight));
        if (panelWidth <= 0.0f || maxPanelHeight <= 0.0f)
        {
            return;
        }

        float innerPaddingDip     = kInnerPaddingDip;
        const float maxPaddingDip = std::max(0.0f, std::min(panelWidth, maxPanelHeight) * 0.05f);
        if (maxPaddingDip > 0.0f)
        {
            innerPaddingDip = std::min(innerPaddingDip, maxPaddingDip);
        }
        else
        {
            innerPaddingDip = 0.0f;
        }

        const float closeReserveDip = _model.closable ? (kCloseSizeDip + kCloseInsetDip) : 0.0f;
        const float maxTextWidthDip = std::max(1.0f, panelWidth - innerPaddingDip * 2.0f - closeReserveDip);

        const bool allowIcon = maxPanelHeight >= (innerPaddingDip * 2.0f + kIconSizeDip);
        bool showIcon        = allowIcon;

        float textWidthDip = maxTextWidthDip;
        if (showIcon)
        {
            textWidthDip = std::max(1.0f, maxTextWidthDip - (kIconSizeDip + kIconTextGapDip));
        }

        constexpr float kMinTextWidthForIconDip = 120.0f;
        if (showIcon && textWidthDip < kMinTextWidthForIconDip)
        {
            showIcon     = false;
            textWidthDip = maxTextWidthDip;
        }

        EnsureTextLayouts(dwriteFactory, textWidthDip);
        EnsureButtonLayouts(dwriteFactory);

        float titleHeightDip = _titleLayoutHeightDip;
        float bodyHeightDip  = _bodyLayoutHeightDip;

        float textHeightDip = 0.0f;
        if (titleHeightDip > 0.0f)
        {
            textHeightDip += titleHeightDip;
        }
        if (bodyHeightDip > 0.0f)
        {
            if (textHeightDip > 0.0f)
            {
                textHeightDip += kTitleBodyGapDip;
            }
            textHeightDip += bodyHeightDip;
        }

        const float buttonRowHeightDip = _model.buttons.empty() ? 0.0f : kButtonHeightDip;
        const float contentHeightDip   = std::max(showIcon ? kIconSizeDip : 0.0f, textHeightDip);

        float desiredPanelHeight = innerPaddingDip * 2.0f + contentHeightDip;
        if (buttonRowHeightDip > 0.0f)
        {
            desiredPanelHeight += kButtonsRowGapDip + buttonRowHeightDip;
        }

        const float clampedPanelHeight = std::clamp(desiredPanelHeight, std::min(desiredPanelHeight, maxPanelHeight), maxPanelHeight);
        const float panelLeft          = (clientWidthDip - panelWidth) * 0.5f;
        const float panelTop           = (clientHeightDip - clampedPanelHeight) * 0.5f;

        _panelRect = D2D1::RectF(panelLeft, panelTop, panelLeft + panelWidth, panelTop + clampedPanelHeight);
        _hasLayout = true;

        const uint64_t elapsedMs      = (nowTickMs >= _startTickMs) ? (nowTickMs - _startTickMs) : 0u;
        const float showT             = static_cast<float>(std::min<uint64_t>(elapsedMs, kShowAnimationMs)) / static_cast<float>(kShowAnimationMs);
        const float ease              = EaseOutCubic(showT);
        const float overlayOpacity    = ease;
        const float overlayScale      = std::lerp(0.975f, 1.0f, ease);
        const float overlayTranslateY = std::lerp(10.0f, 0.0f, ease);

        const float scrimOpacity = _theme.darkBase ? 0.65f : 0.50f;
        _scrimBrush->SetOpacity(scrimOpacity * overlayOpacity);
        target->FillRectangle(D2D1::RectF(0.0f, 0.0f, clientWidthDip, clientHeightDip), _scrimBrush.get());

        D2D1_MATRIX_3X2_F baseTransform{};
        target->GetTransform(&baseTransform);
        const D2D1_POINT_2F panelCenter = D2D1::Point2F((_panelRect.left + _panelRect.right) * 0.5f, (_panelRect.top + _panelRect.bottom) * 0.5f);
        target->SetTransform(D2D1::Matrix3x2F::Scale(overlayScale, overlayScale, panelCenter) * D2D1::Matrix3x2F::Translation(0.0f, overlayTranslateY));

        D2D1::ColorF panelColor  = _theme.background;
        D2D1::ColorF accentColor = _theme.accent;
        D2D1::ColorF textColor   = _theme.text;
        ResolvePalette(_model.severity, panelColor, accentColor, textColor);

        _backgroundBrush->SetColor(panelColor);
        _backgroundBrush->SetOpacity(kCardOpacity * overlayOpacity);

        const float cornerRadiusDip          = std::min(kCornerRadiusDip, std::max(0.0f, std::min(panelWidth, clampedPanelHeight) * 0.25f));
        const D2D1_ROUNDED_RECT roundedPanel = D2D1::RoundedRect(_panelRect, cornerRadiusDip, cornerRadiusDip);
        target->FillRoundedRectangle(roundedPanel, _backgroundBrush.get());

        _textBrush->SetColor(accentColor);
        _textBrush->SetOpacity(kBorderOpacity * overlayOpacity);
        target->DrawRoundedRectangle(roundedPanel, _textBrush.get(), 1.0f);

        D2D1_RECT_F contentRect = D2D1::RectF(
            _panelRect.left + innerPaddingDip, _panelRect.top + innerPaddingDip, _panelRect.right - innerPaddingDip, _panelRect.bottom - innerPaddingDip);

        if (_model.closable)
        {
            const float closeRight = _panelRect.right - kCloseInsetDip;
            const float closeTop   = _panelRect.top + kCloseInsetDip;
            _closeRect             = D2D1::RectF(closeRight - kCloseSizeDip, closeTop, closeRight, closeTop + kCloseSizeDip);
            DrawCloseButton(target, _closeRect, accentColor, overlayOpacity);

            contentRect.right = std::min(contentRect.right, _closeRect.left - kCloseInsetDip);
        }
        else
        {
            _closeRect = {};
        }

        D2D1_RECT_F bodyTextRect = contentRect;
        if (buttonRowHeightDip > 0.0f)
        {
            bodyTextRect.bottom = bodyTextRect.bottom - buttonRowHeightDip - kButtonsRowGapDip;
        }

        D2D1_RECT_F iconRect{};
        D2D1_RECT_F textRect = bodyTextRect;
        if (showIcon)
        {
            const float iconTop = bodyTextRect.top + (contentHeightDip - kIconSizeDip) * 0.5f;
            iconRect            = D2D1::RectF(bodyTextRect.left, iconTop, bodyTextRect.left + kIconSizeDip, iconTop + kIconSizeDip);

            textRect = D2D1::RectF(iconRect.right + kIconTextGapDip, bodyTextRect.top, bodyTextRect.right, bodyTextRect.bottom);

            const float dividerX = iconRect.right + kIconTextGapDip * 0.5f;
            _textBrush->SetColor(accentColor);
            _textBrush->SetOpacity(0.15f * overlayOpacity);
            target->DrawLine(D2D1::Point2F(dividerX, bodyTextRect.top), D2D1::Point2F(dividerX, bodyTextRect.bottom), _textBrush.get(), 1.0f);

            _textBrush->SetColor(accentColor);
            _textBrush->SetOpacity(overlayOpacity);
            DrawSeverityIcon(target, _textBrush.get(), _model.severity, iconRect, overlayOpacity, elapsedMs);
        }

        _textBrush->SetColor(textColor);
        _textBrush->SetOpacity(overlayOpacity);

        float textY = textRect.top;
        target->PushAxisAlignedClip(textRect, D2D1_ANTIALIAS_MODE_PER_PRIMITIVE);

        if (_titleLayout)
        {
            target->DrawTextLayout(D2D1::Point2F(textRect.left, textY), _titleLayout.get(), _textBrush.get());
            textY += titleHeightDip;
        }
        if (_bodyLayout)
        {
            if (textY > textRect.top)
            {
                textY += kTitleBodyGapDip;
            }
            target->DrawTextLayout(D2D1::Point2F(textRect.left, textY), _bodyLayout.get(), _textBrush.get());
        }

        target->PopAxisAlignedClip();

        if (buttonRowHeightDip > 0.0f)
        {
            const D2D1_RECT_F buttonsRect =
                D2D1::RectF(contentRect.left, contentRect.bottom - buttonRowHeightDip, _panelRect.right - innerPaddingDip, contentRect.bottom);
            LayoutButtons(buttonsRect, kButtonsGapDip, kButtonHeightDip, kButtonMinWidthDip, kButtonHorzPaddingDip);
            DrawButtons(target, overlayOpacity, kButtonCornerDip);
        }
        else
        {
            _buttonRects.clear();
        }

        target->SetTransform(baseTransform);
    }

private:
    struct ButtonRect
    {
        uint32_t id = 0;
        D2D1_RECT_F rect{};
        bool primary = false;
        std::wstring_view label;
        float labelWidthDip = 0.0f;
    };

    static bool PointInRect(D2D1_POINT_2F pt, const D2D1_RECT_F& rc) noexcept
    {
        return pt.x >= rc.left && pt.x <= rc.right && pt.y >= rc.top && pt.y <= rc.bottom;
    }

    static float EaseOutCubic(float t) noexcept
    {
        t               = std::clamp(t, 0.0f, 1.0f);
        const float inv = 1.0f - t;
        return 1.0f - inv * inv * inv;
    }

    void EnsureDeviceResources(ID2D1RenderTarget* target) noexcept
    {
        if (! target)
        {
            return;
        }

        if (_targetIdentity == target && _scrimBrush && _backgroundBrush && _textBrush)
        {
            return;
        }

        ResetDeviceResources();
        _targetIdentity = target;

        static_cast<void>(target->CreateSolidColorBrush(D2D1::ColorF(0.0f, 0.0f, 0.0f, 1.0f), _scrimBrush.put()));
        static_cast<void>(target->CreateSolidColorBrush(_theme.background, _backgroundBrush.put()));
        static_cast<void>(target->CreateSolidColorBrush(_theme.text, _textBrush.put()));
    }

    void EnsureTextResources(IDWriteFactory* dwriteFactory) noexcept
    {
        if (! dwriteFactory)
        {
            return;
        }

        if (_dwriteIdentity == dwriteFactory && _titleFormat && _bodyFormat && _buttonFormat)
        {
            return;
        }

        ResetTextResources();
        _dwriteIdentity = dwriteFactory;

        constexpr float kTitleSizeDip  = 18.0f;
        constexpr float kBodySizeDip   = 14.0f;
        constexpr float kButtonSizeDip = 13.0f;
        constexpr float kIconSizeDip   = 56.0f;

        static_cast<void>(dwriteFactory->CreateTextFormat(
            L"Segoe UI", nullptr, DWRITE_FONT_WEIGHT_SEMI_BOLD, DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, kTitleSizeDip, L"", _titleFormat.put()));
        static_cast<void>(dwriteFactory->CreateTextFormat(
            L"Segoe UI", nullptr, DWRITE_FONT_WEIGHT_NORMAL, DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, kBodySizeDip, L"", _bodyFormat.put()));
        static_cast<void>(dwriteFactory->CreateTextFormat(L"Segoe UI",
                                                          nullptr,
                                                          DWRITE_FONT_WEIGHT_SEMI_BOLD,
                                                          DWRITE_FONT_STYLE_NORMAL,
                                                          DWRITE_FONT_STRETCH_NORMAL,
                                                          kButtonSizeDip,
                                                          L"",
                                                          _buttonFormat.put()));

        if (FAILED(dwriteFactory->CreateTextFormat(L"Segoe Fluent Icons",
                                                   nullptr,
                                                   DWRITE_FONT_WEIGHT_NORMAL,
                                                   DWRITE_FONT_STYLE_NORMAL,
                                                   DWRITE_FONT_STRETCH_NORMAL,
                                                   kIconSizeDip,
                                                   L"",
                                                   _iconFormat.put())))
        {
            _iconFormat.reset();
            static_cast<void>(dwriteFactory->CreateTextFormat(L"Segoe MDL2 Assets",
                                                              nullptr,
                                                              DWRITE_FONT_WEIGHT_NORMAL,
                                                              DWRITE_FONT_STYLE_NORMAL,
                                                              DWRITE_FONT_STRETCH_NORMAL,
                                                              kIconSizeDip,
                                                              L"",
                                                              _iconFormat.put()));
        }

        if (_iconFormat)
        {
            _iconGlyphSet = IconGlyphSet::Fluent;
        }
        else
        {
            static_cast<void>(dwriteFactory->CreateTextFormat(L"Segoe UI Symbol",
                                                              nullptr,
                                                              DWRITE_FONT_WEIGHT_NORMAL,
                                                              DWRITE_FONT_STYLE_NORMAL,
                                                              DWRITE_FONT_STRETCH_NORMAL,
                                                              kIconSizeDip,
                                                              L"",
                                                              _iconFormat.put()));
            if (_iconFormat)
            {
                _iconGlyphSet = IconGlyphSet::Unicode;
            }
        }

        if (_titleFormat)
        {
            static_cast<void>(_titleFormat->SetWordWrapping(DWRITE_WORD_WRAPPING_WRAP));
        }
        if (_bodyFormat)
        {
            static_cast<void>(_bodyFormat->SetWordWrapping(DWRITE_WORD_WRAPPING_WRAP));
        }
        if (_buttonFormat)
        {
            static_cast<void>(_buttonFormat->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP));
            static_cast<void>(_buttonFormat->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_CENTER));
            static_cast<void>(_buttonFormat->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER));
        }
        if (_iconFormat)
        {
            static_cast<void>(_iconFormat->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP));
            static_cast<void>(_iconFormat->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_CENTER));
            static_cast<void>(_iconFormat->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER));
        }
    }

    void ResetFocus() noexcept
    {
        _focusedButtonId.reset();
        for (const auto& button : _model.buttons)
        {
            if (button.primary)
            {
                _focusedButtonId = button.id;
                return;
            }
        }

        if (! _model.buttons.empty())
        {
            _focusedButtonId = _model.buttons.front().id;
        }
    }

    void InvalidateTextLayouts() noexcept
    {
        _cachedTitle.clear();
        _cachedMessage.clear();
        _cachedTextWidthDip = 0.0f;
        _titleLayout.reset();
        _bodyLayout.reset();
        _titleLayoutHeightDip = 0.0f;
        _bodyLayoutHeightDip  = 0.0f;
    }

    void EnsureTextLayouts(IDWriteFactory* dwriteFactory, float textWidthDip) noexcept
    {
        if (! dwriteFactory || ! _titleFormat || ! _bodyFormat)
        {
            return;
        }

        if (_cachedTitle == _model.title && _cachedMessage == _model.message && std::abs(_cachedTextWidthDip - textWidthDip) <= 0.5f)
        {
            return;
        }

        _cachedTitle        = _model.title;
        _cachedMessage      = _model.message;
        _cachedTextWidthDip = textWidthDip;

        _titleLayout.reset();
        _bodyLayout.reset();
        _titleLayoutHeightDip = 0.0f;
        _bodyLayoutHeightDip  = 0.0f;

        if (! _model.title.empty())
        {
            const HRESULT hr = dwriteFactory->CreateTextLayout(
                _model.title.c_str(), static_cast<UINT32>(_model.title.size()), _titleFormat.get(), textWidthDip, 1000.0f, _titleLayout.put());
            if (SUCCEEDED(hr) && _titleLayout)
            {
                DWRITE_TEXT_METRICS metrics{};
                if (SUCCEEDED(_titleLayout->GetMetrics(&metrics)))
                {
                    _titleLayoutHeightDip = metrics.height;
                }
            }
        }

        if (! _model.message.empty())
        {
            const HRESULT hr = dwriteFactory->CreateTextLayout(
                _model.message.c_str(), static_cast<UINT32>(_model.message.size()), _bodyFormat.get(), textWidthDip, 1000.0f, _bodyLayout.put());
            if (SUCCEEDED(hr) && _bodyLayout)
            {
                DWRITE_TEXT_METRICS metrics{};
                if (SUCCEEDED(_bodyLayout->GetMetrics(&metrics)))
                {
                    _bodyLayoutHeightDip = metrics.height;
                }
            }
        }
    }

    void InvalidateButtonLayouts() noexcept
    {
        _cachedButtonLabels.clear();
        _buttonBaseRects.clear();
        _buttonRects.clear();
    }

    void EnsureButtonLayouts(IDWriteFactory* dwriteFactory) noexcept
    {
        if (! dwriteFactory || ! _buttonFormat)
        {
            return;
        }

        std::vector<std::wstring_view> labels;
        labels.reserve(_model.buttons.size());
        for (const auto& button : _model.buttons)
        {
            labels.emplace_back(button.label);
        }

        if (labels.size() == _cachedButtonLabels.size())
        {
            bool same = true;
            for (size_t i = 0; i < labels.size(); ++i)
            {
                if (labels[i] != _cachedButtonLabels[i])
                {
                    same = false;
                    break;
                }
            }
            if (same)
            {
                return;
            }
        }

        _cachedButtonLabels.clear();
        _cachedButtonLabels.reserve(labels.size());
        for (const auto& label : labels)
        {
            _cachedButtonLabels.emplace_back(label);
        }

        _buttonBaseRects.clear();
        _buttonBaseRects.reserve(_model.buttons.size());

        for (const auto& button : _model.buttons)
        {
            ButtonRect btn{};
            btn.id            = button.id;
            btn.primary       = button.primary;
            btn.label         = button.label;
            btn.labelWidthDip = 0.0f;

            if (! button.label.empty())
            {
                wil::com_ptr<IDWriteTextLayout> layout;
                const HRESULT hr = dwriteFactory->CreateTextLayout(
                    button.label.c_str(), static_cast<UINT32>(button.label.size()), _buttonFormat.get(), 1000.0f, 1000.0f, layout.put());
                if (SUCCEEDED(hr) && layout)
                {
                    DWRITE_TEXT_METRICS metrics{};
                    if (SUCCEEDED(layout->GetMetrics(&metrics)))
                    {
                        btn.labelWidthDip = metrics.widthIncludingTrailingWhitespace;
                    }
                }
            }

            _buttonBaseRects.emplace_back(std::move(btn));
        }
    }

    void ResolvePalette(AlertSeverity severity, D2D1::ColorF& panelBackground, D2D1::ColorF& accentColor, D2D1::ColorF& textColor) const noexcept
    {
        if (severity == AlertSeverity::Error)
        {
            panelBackground = _theme.errorBackground;
            accentColor     = _theme.errorText;
            textColor       = _theme.errorText;
            return;
        }

        if (severity == AlertSeverity::Warning)
        {
            panelBackground = _theme.warningBackground;
            accentColor     = _theme.warningText;
            textColor       = _theme.warningText;
            return;
        }

        if (severity == AlertSeverity::Info)
        {
            panelBackground = _theme.infoBackground;
            accentColor     = _theme.infoText;
            textColor       = _theme.infoText;
            return;
        }

        if (severity == AlertSeverity::Busy)
        {
            panelBackground = _theme.background;
            accentColor     = _theme.accent;
            textColor       = _theme.text;
            return;
        }
    }

    void DrawCloseButton(ID2D1RenderTarget* target, const D2D1_RECT_F& rect, const D2D1::ColorF& accentColor, float opacity) noexcept
    {
        if (! target || ! _backgroundBrush || ! _textBrush)
        {
            return;
        }

        const bool hot = (_hot.part == AlertHitTest::Part::Close);
        if (hot)
        {
            const D2D1::ColorF bg = D2D1::ColorF(accentColor.r, accentColor.g, accentColor.b, 0.14f);
            _backgroundBrush->SetColor(bg);
            _backgroundBrush->SetOpacity(opacity);
            const float r = std::min(rect.right - rect.left, rect.bottom - rect.top) * 0.35f;
            target->FillRoundedRectangle(D2D1::RoundedRect(rect, r, r), _backgroundBrush.get());
        }

        const float w      = std::max(0.0f, rect.right - rect.left);
        const float h      = std::max(0.0f, rect.bottom - rect.top);
        const float size   = std::min(w, h);
        const float stroke = std::clamp(size * 0.10f, 1.5f, 2.5f);
        const float pad    = size * 0.28f;

        const D2D1_POINT_2F a = D2D1::Point2F(rect.left + pad, rect.top + pad);
        const D2D1_POINT_2F b = D2D1::Point2F(rect.right - pad, rect.bottom - pad);
        const D2D1_POINT_2F c = D2D1::Point2F(rect.right - pad, rect.top + pad);
        const D2D1_POINT_2F d = D2D1::Point2F(rect.left + pad, rect.bottom - pad);

        _textBrush->SetColor(accentColor);
        _textBrush->SetOpacity(opacity);
        target->DrawLine(a, b, _textBrush.get(), stroke);
        target->DrawLine(c, d, _textBrush.get(), stroke);
    }

    void LayoutButtons(const D2D1_RECT_F& rowRect, float gapDip, float heightDip, float minWidthDip, float horzPaddingDip) noexcept
    {
        _buttonRects.clear();

        if (_buttonBaseRects.empty())
        {
            return;
        }

        float right        = rowRect.right;
        const float bottom = rowRect.bottom;
        const float top    = bottom - heightDip;

        for (size_t i = 0; i < _buttonBaseRects.size(); ++i)
        {
            const auto& src = _buttonBaseRects[_buttonBaseRects.size() - 1u - i];

            const float widthDip = std::max(minWidthDip, src.labelWidthDip + horzPaddingDip * 2.0f);

            const float left = right - widthDip;
            if (left < rowRect.left)
            {
                break;
            }

            ButtonRect btn = src;
            btn.rect       = D2D1::RectF(left, top, right, bottom);
            _buttonRects.emplace_back(std::move(btn));

            right = left - gapDip;
        }

        std::reverse(_buttonRects.begin(), _buttonRects.end());
    }

    void DrawButtons(ID2D1RenderTarget* target, float opacity, float cornerDip) noexcept
    {
        if (! target || ! _backgroundBrush || ! _textBrush)
        {
            return;
        }

        for (const auto& btn : _buttonRects)
        {
            const bool hot     = (_hot.part == AlertHitTest::Part::Button && _hot.buttonId == btn.id);
            const bool focused = (_focusedButtonId.has_value() && _focusedButtonId.value() == btn.id);

            const D2D1_ROUNDED_RECT rr = D2D1::RoundedRect(btn.rect, cornerDip, cornerDip);
            if (btn.primary)
            {
                const D2D1::ColorF bg = hot ? D2D1::ColorF(_theme.accent.r, _theme.accent.g, _theme.accent.b, 0.95f) : _theme.accent;
                _backgroundBrush->SetColor(bg);
                _backgroundBrush->SetOpacity(opacity);
                target->FillRoundedRectangle(rr, _backgroundBrush.get());

                _textBrush->SetColor(_theme.selectionText);
                _textBrush->SetOpacity(opacity);
                target->DrawRoundedRectangle(rr, _textBrush.get(), 1.0f);
            }
            else
            {
                const D2D1::ColorF border = hot ? _theme.accent : _theme.text;
                const D2D1::ColorF bg     = hot ? D2D1::ColorF(border.r, border.g, border.b, 0.10f) : D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);

                _backgroundBrush->SetColor(bg);
                _backgroundBrush->SetOpacity(opacity);
                target->FillRoundedRectangle(rr, _backgroundBrush.get());

                _textBrush->SetColor(border);
                _textBrush->SetOpacity(opacity);
                target->DrawRoundedRectangle(rr, _textBrush.get(), 1.0f);
            }

            if (! btn.label.empty())
            {
                _textBrush->SetColor(btn.primary ? _theme.selectionText : _theme.text);
                _textBrush->SetOpacity(opacity);
                target->DrawTextW(btn.label.data(), static_cast<UINT32>(btn.label.size()), _buttonFormat.get(), btn.rect, _textBrush.get());
            }

            if (focused)
            {
                constexpr float kFocusOutsetDip = 2.0f;

                D2D1_RECT_F focusRect = btn.rect;
                focusRect.left -= kFocusOutsetDip;
                focusRect.top -= kFocusOutsetDip;
                focusRect.right += kFocusOutsetDip;
                focusRect.bottom += kFocusOutsetDip;

                const float focusCorner         = cornerDip + kFocusOutsetDip;
                const D2D1_ROUNDED_RECT focusRr = D2D1::RoundedRect(focusRect, focusCorner, focusCorner);

                const D2D1::ColorF focusColor = btn.primary ? _theme.selectionText : _theme.accent;
                _textBrush->SetColor(focusColor);
                _textBrush->SetOpacity(opacity);
                target->DrawRoundedRectangle(focusRr, _textBrush.get(), 2.0f);
            }
        }
    }

    void DrawSeverityIcon(
        ID2D1RenderTarget* target, ID2D1SolidColorBrush* brush, AlertSeverity severity, const D2D1_RECT_F& rect, float opacity, uint64_t elapsedMs) noexcept
    {
        if (! target || ! brush)
        {
            return;
        }

        if (severity != AlertSeverity::Busy && _iconFormat && _iconGlyphSet != IconGlyphSet::None)
        {
            constexpr wchar_t kFallbackInfo = L'\u2139'; // ℹ

            wchar_t glyph = L'?';
            if (_iconGlyphSet == IconGlyphSet::Fluent)
            {
                switch (severity)
                {
                    case AlertSeverity::Error: glyph = L'\uEA39'; break;
                    case AlertSeverity::Warning: glyph = L'\uE7BA'; break;
                    case AlertSeverity::Info: glyph = L'\uE946'; break;
                    case AlertSeverity::Busy: break;
                    default: break;
                }
            }
            else
            {
                switch (severity)
                {
                    case AlertSeverity::Error: glyph = L'\u2716'; break;   // ✖
                    case AlertSeverity::Warning: glyph = L'\u26A0'; break; // ⚠
                    case AlertSeverity::Info: glyph = kFallbackInfo; break;
                    case AlertSeverity::Busy: break;
                    default: break;
                }
            }

            const wchar_t text[2]{glyph, 0};
            brush->SetOpacity(opacity);
            target->DrawTextW(text, 1, _iconFormat.get(), rect, brush, D2D1_DRAW_TEXT_OPTIONS_NO_SNAP, DWRITE_MEASURING_MODE_NATURAL);
            return;
        }

        constexpr float kPi = 3.14159265358979323846f;

        const float width  = std::max(0.0f, rect.right - rect.left);
        const float height = std::max(0.0f, rect.bottom - rect.top);
        const float size   = std::min(width, height);
        if (size <= 0.0f)
        {
            return;
        }

        const D2D1_POINT_2F center = D2D1::Point2F((rect.left + rect.right) * 0.5f, (rect.top + rect.bottom) * 0.5f);
        const float radius         = size * 0.46f;
        const float stroke         = std::clamp(size * 0.06f, 2.0f, 4.0f);

        brush->SetOpacity(opacity);

        if (severity == AlertSeverity::Error)
        {
            target->DrawEllipse(D2D1::Ellipse(center, radius, radius), brush, stroke);

            const float x = radius * 0.45f;
            const float y = radius * 0.45f;
            target->DrawLine(D2D1::Point2F(center.x - x, center.y - y), D2D1::Point2F(center.x + x, center.y + y), brush, stroke);
            target->DrawLine(D2D1::Point2F(center.x - x, center.y + y), D2D1::Point2F(center.x + x, center.y - y), brush, stroke);
            return;
        }

        if (severity == AlertSeverity::Warning)
        {
            const float a          = -kPi / 2.0f;
            const float b          = 5.0f * kPi / 6.0f;
            const float c          = kPi / 6.0f;
            const D2D1_POINT_2F p0 = D2D1::Point2F(center.x + radius * std::cos(a), center.y + radius * std::sin(a));
            const D2D1_POINT_2F p1 = D2D1::Point2F(center.x + radius * std::cos(b), center.y + radius * std::sin(b));
            const D2D1_POINT_2F p2 = D2D1::Point2F(center.x + radius * std::cos(c), center.y + radius * std::sin(c));

            target->DrawLine(p0, p1, brush, stroke);
            target->DrawLine(p1, p2, brush, stroke);
            target->DrawLine(p2, p0, brush, stroke);

            const float exHeight = radius * 0.55f;
            const float exTop    = center.y - exHeight * 0.45f;
            const float exBottom = center.y + exHeight * 0.15f;
            target->DrawLine(D2D1::Point2F(center.x, exTop), D2D1::Point2F(center.x, exBottom), brush, stroke);
            target->FillEllipse(D2D1::Ellipse(D2D1::Point2F(center.x, center.y + exHeight * 0.40f), stroke * 0.35f, stroke * 0.35f), brush);
            return;
        }

        if (severity == AlertSeverity::Info)
        {
            target->DrawEllipse(D2D1::Ellipse(center, radius, radius), brush, stroke);

            const float dotY = center.y - radius * 0.28f;
            target->FillEllipse(D2D1::Ellipse(D2D1::Point2F(center.x, dotY), stroke * 0.35f, stroke * 0.35f), brush);

            const float lineTop    = center.y - radius * 0.05f;
            const float lineBottom = center.y + radius * 0.38f;
            target->DrawLine(D2D1::Point2F(center.x, lineTop), D2D1::Point2F(center.x, lineBottom), brush, stroke);
            return;
        }

        if (severity == AlertSeverity::Busy)
        {
            constexpr int dotCount = 12;
            const float ringRadius = radius * 0.72f;
            const float dotRadius  = std::max(stroke * 0.45f, 2.0f);
            const float cycle      = static_cast<float>(elapsedMs % 900u) / 900.0f;

            for (int i = 0; i < dotCount; ++i)
            {
                const float angle = (static_cast<float>(i) / static_cast<float>(dotCount)) * (2.0f * kPi);
                const float local =
                    std::fmod((cycle * static_cast<float>(dotCount)) - static_cast<float>(i) + static_cast<float>(dotCount), static_cast<float>(dotCount));
                const float intensity  = 1.0f - (local / static_cast<float>(dotCount));
                const float dotOpacity = std::lerp(0.18f, 1.0f, intensity * intensity);

                brush->SetOpacity(opacity * dotOpacity);
                const D2D1_POINT_2F dotCenter = D2D1::Point2F(center.x + ringRadius * std::cos(angle), center.y + ringRadius * std::sin(angle));
                target->FillEllipse(D2D1::Ellipse(dotCenter, dotRadius, dotRadius), brush);
            }

            brush->SetOpacity(opacity);
            return;
        }
    }

private:
    AlertModel _model{};
    AlertTheme _theme{};

    ID2D1RenderTarget* _targetIdentity = nullptr;
    IDWriteFactory* _dwriteIdentity    = nullptr;

    wil::com_ptr<ID2D1SolidColorBrush> _scrimBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _backgroundBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _textBrush;

    wil::com_ptr<IDWriteTextFormat> _titleFormat;
    wil::com_ptr<IDWriteTextFormat> _bodyFormat;
    wil::com_ptr<IDWriteTextFormat> _buttonFormat;
    wil::com_ptr<IDWriteTextFormat> _iconFormat;
    IconGlyphSet _iconGlyphSet = IconGlyphSet::None;

    std::wstring _cachedTitle;
    std::wstring _cachedMessage;
    float _cachedTextWidthDip = 0.0f;
    wil::com_ptr<IDWriteTextLayout> _titleLayout;
    wil::com_ptr<IDWriteTextLayout> _bodyLayout;
    float _titleLayoutHeightDip = 0.0f;
    float _bodyLayoutHeightDip  = 0.0f;

    std::vector<std::wstring_view> _cachedButtonLabels;
    std::vector<ButtonRect> _buttonBaseRects;
    std::vector<ButtonRect> _buttonRects;

    D2D1_RECT_F _panelRect{};
    D2D1_RECT_F _closeRect{};
    bool _hasLayout = false;

    AlertHitTest _hot{};
    std::optional<uint32_t> _focusedButtonId;
    uint64_t _startTickMs = 0;
};
} // namespace RedSalamander::Ui
