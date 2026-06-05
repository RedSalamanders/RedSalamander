#pragma once

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <Windows.h>

#include "DxUi.h"
#include "RedConfigureApp.h"
#include "RedConfigureUiHelpers.h"
#include "Themes/ThemePreviewModel.h"
#include "resource.h"

#include <algorithm>
#include <array>
#include <cmath>
#include <cstdint>
#include <functional>
#include <string>
#include <string_view>
#include <vector>

namespace RedConfigure::Ui
{
using RedSalamander::DxUi::Control;
using RedSalamander::DxUi::FontRole;
using RedSalamander::DxUi::ThemePalette;
using RedSalamander::DxUi::WindowHost;

class ThemeExampleControl final : public Control
{
public:
    explicit ThemeExampleControl(HINSTANCE instance)
        : _navText(LoadAppString(instance, IDS_REDCONFIGURE_THEME_PREVIEW_NAV)),
          _menuText(LoadAppString(instance, IDS_REDCONFIGURE_THEME_PREVIEW_MENU)),
          _folderText(LoadAppString(instance, IDS_REDCONFIGURE_THEME_PREVIEW_FOLDER)),
          _hoverText(LoadAppString(instance, IDS_REDCONFIGURE_THEME_PREVIEW_HOVER)),
          _selectedText(LoadAppString(instance, IDS_REDCONFIGURE_THEME_PREVIEW_SELECTED)),
          _dialogText(LoadAppString(instance, IDS_REDCONFIGURE_THEME_PREVIEW_DIALOG)),
          _buttonText(LoadAppString(instance, IDS_REDCONFIGURE_THEME_PREVIEW_BUTTON)),
          _progressText(LoadAppString(instance, IDS_REDCONFIGURE_THEME_PREVIEW_PROGRESS)),
          _warningText(LoadAppString(instance, IDS_REDCONFIGURE_THEME_PREVIEW_WARNING)),
          _diffAddedText(LoadAppString(instance, IDS_REDCONFIGURE_THEME_PREVIEW_DIFF_ADDED)),
          _diffRemovedText(LoadAppString(instance, IDS_REDCONFIGURE_THEME_PREVIEW_DIFF_REMOVED))
    {
    }

    void SetModel(const RedConfigure::Themes::ThemePreviewModel* model) noexcept
    {
        _model = model;
        RequestInvalidate();
    }

    void Refresh() noexcept
    {
        RequestInvalidate();
    }

    void SetSelectedToken(std::wstring_view key)
    {
        _selectedTokenKey = std::wstring(key);
        RequestInvalidate();
    }

    void SetOnTokenSelected(std::function<void(std::wstring_view)> onTokenSelected)
    {
        _onTokenSelected = std::move(onTokenSelected);
    }

    bool OnMouseDown(WindowHost&, D2D1_POINT_2F point, bool rightButton, UINT) override
    {
        if (rightButton || ! _onTokenSelected)
        {
            return false;
        }

        const PreviewLayout layout                                           = BuildLayout(GetBounds());
        const std::vector<RedConfigure::ThemePreviewHitCandidate> candidates = BuildHitCandidates(layout);
        const bool continuingCycle                                           = IsSameClickPoint(point, _lastClickPoint);
        const std::wstring selectedKey =
            RedConfigure::SelectThemePreviewHitKey(candidates, point.x, point.y, continuingCycle ? std::wstring_view(_lastClickKey) : std::wstring_view{});
        if (! selectedKey.empty())
        {
            _lastClickPoint = point;
            _lastClickKey   = selectedKey;
            _onTokenSelected(selectedKey);
            return true;
        }

        return false;
    }

    void Paint(WindowHost& host) const override
    {
        auto* dc = host.GetDeviceContext();
        if (! dc || ! _model)
        {
            return;
        }

        const D2D1_RECT_F bounds    = GetBounds();
        const PreviewLayout layout  = BuildLayout(bounds);
        const ThemePalette& palette = host.GetTheme();

        const D2D1_COLOR_F windowColor        = ColorFromArgb(ColorOrDefault(*_model, L"window.background", 0xFFFFFFFFu));
        const D2D1_COLOR_F windowTextColor    = ColorFromArgb(ColorOrDefault(*_model, L"window.text", 0xFF111111u));
        const D2D1_COLOR_F navColor           = ColorFromArgb(ColorOrDefault(*_model, L"navigation.background", 0xFFEAF2FEu));
        const D2D1_COLOR_F navTextColor       = ColorFromArgb(ColorOrDefault(*_model, L"navigation.text", 0xFF1F2937u));
        const D2D1_COLOR_F navAccentColor     = ColorFromArgb(ColorOrDefault(*_model, L"app.accent", 0xFF0F6CBDu));
        const D2D1_COLOR_F menuColor          = ColorFromArgb(ColorOrDefault(*_model, L"menu.background", 0xFFFFFFFFu));
        const D2D1_COLOR_F menuTextColor      = ColorFromArgb(ColorOrDefault(*_model, L"menu.text", 0xFF111111u));
        const D2D1_COLOR_F menuSelectionColor = ColorFromArgb(ColorOrDefault(*_model, L"menu.selectionBackground", 0xFFE8F1FFu));
        const D2D1_COLOR_F menuBorderColor    = ColorFromArgb(ColorOrDefault(*_model, L"menu.border", 0xFFD8D8D8u));
        const D2D1_COLOR_F folderColor        = ColorFromArgb(ColorOrDefault(*_model, L"folderView.background", 0xFFFFFFFFu));
        const D2D1_COLOR_F folderTextColor    = ColorFromArgb(ColorOrDefault(*_model, L"folderView.itemForeground", 0xFF111111u));
        const D2D1_COLOR_F hoverColor         = ColorFromArgb(ColorOrDefault(*_model, L"folderView.itemBackgroundHovered", 0xFFF0F6FFu));
        const D2D1_COLOR_F selectedColor      = ColorFromArgb(ColorOrDefault(*_model, L"folderView.itemBackgroundSelected", 0xFFCFE8FFu));
        const D2D1_COLOR_F selectedTextColor  = ColorFromArgb(ColorOrDefault(*_model, L"folderView.itemForegroundSelected", 0xFF0F172Au));
        const D2D1_COLOR_F warningColor       = ColorFromArgb(ColorOrDefault(*_model, L"folderView.warningForeground", 0xFF8A4B00u));
        const D2D1_COLOR_F dialogColor        = ColorFromArgb(ColorOrDefault(*_model, L"dialog.background", 0xFFF7F7F7u));
        const D2D1_COLOR_F dialogTextColor    = ColorFromArgb(ColorOrDefault(*_model, L"dialog.text", 0xFF111111u));
        const D2D1_COLOR_F buttonColor        = ColorFromArgb(ColorOrDefault(*_model, L"dialog.buttonBackground", 0xFFFFFFFFu));
        const D2D1_COLOR_F buttonTextColor    = ColorFromArgb(ColorOrDefault(*_model, L"dialog.buttonText", 0xFF111111u));
        const D2D1_COLOR_F progressBgColor    = ColorFromArgb(ColorOrDefault(*_model, L"progress.background", 0xFFE5E7EBu));
        const D2D1_COLOR_F progressFillColor  = ColorFromArgb(ColorOrDefault(*_model, L"progress.fill", 0xFF2563EBu));
        const D2D1_COLOR_F diffAddedColor     = ColorFromArgb(ColorOrDefault(*_model, L"diff.addedBackground", 0xFFEAF8EFu));
        const D2D1_COLOR_F diffRemovedColor   = ColorFromArgb(ColorOrDefault(*_model, L"diff.removedBackground", 0xFFFFECEFu));

        if (auto* brush = host.GetSolidBrush(windowColor))
        {
            dc->FillRectangle(bounds, brush);
        }
        if (auto* brush = host.GetSolidBrush(palette.borderDefault))
        {
            dc->DrawRectangle(bounds, brush, 1.0f);
        }

        if (auto* brush = host.GetSolidBrush(navColor))
        {
            dc->FillRectangle(layout.navRect, brush);
        }
        if (auto* brush = host.GetSolidBrush(navAccentColor))
        {
            dc->FillRectangle(layout.navAccentRect, brush);
        }
        if (auto* brush = host.GetSolidBrush(menuColor))
        {
            dc->FillRectangle(layout.menuRect, brush);
        }
        if (auto* brush = host.GetSolidBrush(menuSelectionColor))
        {
            dc->FillRectangle(layout.menuSelectionRect, brush);
        }
        if (auto* brush = host.GetSolidBrush(folderColor))
        {
            dc->FillRectangle(layout.folderRect, brush);
        }
        if (auto* brush = host.GetSolidBrush(hoverColor))
        {
            dc->FillRectangle(layout.hoverRowRect, brush);
        }
        if (auto* brush = host.GetSolidBrush(selectedColor))
        {
            dc->FillRectangle(layout.selectedRowRect, brush);
        }
        if (auto* brush = host.GetSolidBrush(dialogColor))
        {
            dc->FillRectangle(layout.dialogRect, brush);
        }
        if (auto* brush = host.GetSolidBrush(buttonColor))
        {
            dc->FillRectangle(layout.buttonRect, brush);
        }
        if (auto* brush = host.GetSolidBrush(progressBgColor))
        {
            dc->FillRectangle(layout.progressBgRect, brush);
        }
        if (auto* brush = host.GetSolidBrush(progressFillColor))
        {
            dc->FillRectangle(layout.progressFillRect, brush);
        }
        if (auto* brush = host.GetSolidBrush(diffAddedColor))
        {
            dc->FillRectangle(layout.diffAddedRect, brush);
        }
        if (auto* brush = host.GetSolidBrush(diffRemovedColor))
        {
            dc->FillRectangle(layout.diffRemovedRect, brush);
        }
        if (auto* brush = host.GetSolidBrush(palette.borderDefault))
        {
            dc->DrawRectangle(layout.menuRect, brush, 1.0f);
            dc->DrawRectangle(layout.folderRect, brush, 1.0f);
            dc->DrawRectangle(layout.dialogRect, brush, 1.0f);
            dc->DrawRectangle(layout.buttonRect, brush, 1.0f);
        }
        if (auto* brush = host.GetSolidBrush(menuBorderColor))
        {
            dc->DrawLine(
                D2D1::Point2F(layout.menuRect.left, layout.menuRect.bottom), D2D1::Point2F(layout.menuRect.right, layout.menuRect.bottom), brush, 1.0f);
        }

        const auto drawText = [&](std::wstring_view text, const D2D1_RECT_F& rect, FontRole role, const D2D1_COLOR_F& color)
        {
            if (auto* brush = host.GetSolidBrush(color))
            {
                dc->DrawTextW(text.data(), static_cast<UINT32>(text.size()), host.GetTextFormat(role), rect, brush);
            }
        };

        drawText(_navText,
                 D2D1::RectF(layout.navRect.left + 14.0f, layout.navRect.top + 14.0f, layout.navRect.right - 14.0f, layout.navRect.top + 44.0f),
                 FontRole::BodyStrong,
                 navTextColor);
        drawText(_menuText,
                 D2D1::RectF(layout.menuRect.left + 12.0f, layout.menuRect.top + 8.0f, layout.menuRect.right - 12.0f, layout.menuRect.bottom),
                 FontRole::Body,
                 menuTextColor);
        drawText(_folderText,
                 D2D1::RectF(layout.folderRect.left + 12.0f, layout.folderRect.top + 10.0f, layout.folderRect.right - 12.0f, layout.folderRect.top + 34.0f),
                 FontRole::BodyStrong,
                 windowTextColor);
        drawText(_hoverText,
                 D2D1::RectF(layout.hoverRowRect.left + 10.0f, layout.hoverRowRect.top + 6.0f, layout.hoverRowRect.right, layout.hoverRowRect.bottom),
                 FontRole::Body,
                 folderTextColor);
        drawText(
            _selectedText,
            D2D1::RectF(layout.selectedRowRect.left + 10.0f, layout.selectedRowRect.top + 6.0f, layout.selectedRowRect.right, layout.selectedRowRect.bottom),
            FontRole::Body,
            selectedTextColor);
        drawText(_warningText,
                 D2D1::RectF(layout.warningRowRect.left + 10.0f, layout.warningRowRect.top + 6.0f, layout.warningRowRect.right, layout.warningRowRect.bottom),
                 FontRole::Body,
                 warningColor);
        drawText(_dialogText,
                 D2D1::RectF(layout.dialogRect.left + 12.0f, layout.dialogRect.top + 10.0f, layout.dialogRect.right - 12.0f, layout.dialogRect.top + 38.0f),
                 FontRole::BodyStrong,
                 dialogTextColor);
        drawText(_buttonText,
                 D2D1::RectF(layout.buttonRect.left + 18.0f, layout.buttonRect.top + 7.0f, layout.buttonRect.right - 18.0f, layout.buttonRect.bottom),
                 FontRole::Body,
                 buttonTextColor);
        drawText(_progressText,
                 D2D1::RectF(layout.progressBgRect.left, layout.progressBgRect.top - 24.0f, layout.progressBgRect.right, layout.progressBgRect.top - 2.0f),
                 FontRole::Small,
                 folderTextColor);
        drawText(_diffAddedText, layout.diffAddedRect, FontRole::Small, folderTextColor);
        drawText(_diffRemovedText, layout.diffRemovedRect, FontRole::Small, folderTextColor);
        DrawSelectedTokenHighlight(host, layout);
    }

private:
    struct PreviewLayout
    {
        D2D1_RECT_F navRect{};
        D2D1_RECT_F navAccentRect{};
        D2D1_RECT_F menuRect{};
        D2D1_RECT_F menuSelectionRect{};
        D2D1_RECT_F folderRect{};
        D2D1_RECT_F hoverRowRect{};
        D2D1_RECT_F selectedRowRect{};
        D2D1_RECT_F warningRowRect{};
        D2D1_RECT_F dialogRect{};
        D2D1_RECT_F buttonRect{};
        D2D1_RECT_F progressBgRect{};
        D2D1_RECT_F progressFillRect{};
        D2D1_RECT_F diffAddedRect{};
        D2D1_RECT_F diffRemovedRect{};
    };

    struct PreviewHitRegion
    {
        D2D1_RECT_F rect{};
        std::wstring_view key;
    };

    [[nodiscard]] static std::array<PreviewHitRegion, 14> BuildHitRegions(const PreviewLayout& layout) noexcept
    {
        return {{
            {layout.navRect, L"navigation.background"},
            {layout.navAccentRect, L"app.accent"},
            {layout.menuRect, L"menu.background"},
            {layout.menuSelectionRect, L"menu.selectionBackground"},
            {layout.folderRect, L"folderView.background"},
            {layout.hoverRowRect, L"folderView.itemBackgroundHovered"},
            {layout.selectedRowRect, L"folderView.itemBackgroundSelected"},
            {layout.warningRowRect, L"folderView.warningForeground"},
            {layout.dialogRect, L"dialog.background"},
            {layout.buttonRect, L"dialog.buttonBackground"},
            {layout.progressBgRect, L"progress.background"},
            {layout.progressFillRect, L"progress.fill"},
            {layout.diffAddedRect, L"diff.addedBackground"},
            {layout.diffRemovedRect, L"diff.removedBackground"},
        }};
    }

    [[nodiscard]] static std::vector<RedConfigure::ThemePreviewHitCandidate> BuildHitCandidates(const PreviewLayout& layout)
    {
        const std::array<PreviewHitRegion, 14> regions = BuildHitRegions(layout);
        std::vector<RedConfigure::ThemePreviewHitCandidate> candidates;
        candidates.reserve(regions.size());
        for (const PreviewHitRegion& region : regions)
        {
            candidates.push_back(RedConfigure::ThemePreviewHitCandidate{
                .key = std::wstring(region.key), .left = region.rect.left, .top = region.rect.top, .right = region.rect.right, .bottom = region.rect.bottom});
        }
        return candidates;
    }

    [[nodiscard]] static bool IsSameClickPoint(const D2D1_POINT_2F& lhs, const D2D1_POINT_2F& rhs) noexcept
    {
        constexpr float thresholdDip = 3.0f;
        return std::abs(lhs.x - rhs.x) <= thresholdDip && std::abs(lhs.y - rhs.y) <= thresholdDip;
    }

    [[nodiscard]] static D2D1_RECT_F InflateLocalRect(const D2D1_RECT_F& rect, float amount) noexcept
    {
        return D2D1::RectF(rect.left - amount, rect.top - amount, rect.right + amount, rect.bottom + amount);
    }

    void DrawSelectedTokenHighlight(WindowHost& host, const PreviewLayout& layout) const
    {
        if (_selectedTokenKey.empty())
        {
            return;
        }

        auto* dc = host.GetDeviceContext();
        if (! dc)
        {
            return;
        }

        const ThemePalette& palette                    = host.GetTheme();
        const std::array<PreviewHitRegion, 14> regions = BuildHitRegions(layout);
        for (const PreviewHitRegion& region : regions)
        {
            if (region.key != _selectedTokenKey)
            {
                continue;
            }

            const D2D1_RECT_F outer = InflateLocalRect(region.rect, 2.0f);
            if (auto* brush = host.GetSolidBrush(palette.focusStrokeOuter))
            {
                dc->DrawRectangle(outer, brush, 3.0f);
            }
            if (auto* brush = host.GetSolidBrush(palette.focusStroke))
            {
                dc->DrawRectangle(InflateLocalRect(region.rect, 0.5f), brush, 2.0f);
            }
        }
    }

    [[nodiscard]] static PreviewLayout BuildLayout(const D2D1_RECT_F& bounds) noexcept
    {
        const float inset = 16.0f;
        PreviewLayout layout{};
        layout.navRect       = D2D1::RectF(bounds.left + inset, bounds.top + inset, bounds.left + 172.0f, bounds.bottom - inset);
        layout.navAccentRect = D2D1::RectF(layout.navRect.left, layout.navRect.top, layout.navRect.left + 5.0f, layout.navRect.bottom);
        layout.menuRect      = D2D1::RectF(layout.navRect.right + 12.0f, bounds.top + inset, bounds.right - inset, bounds.top + inset + 36.0f);
        layout.menuSelectionRect =
            D2D1::RectF(layout.menuRect.left + 92.0f, layout.menuRect.top + 5.0f, layout.menuRect.left + 150.0f, layout.menuRect.bottom - 5.0f);
        layout.folderRect = D2D1::RectF(layout.navRect.right + 12.0f, layout.menuRect.bottom + 12.0f, bounds.right - inset, bounds.bottom - inset);

        const float rowLeft  = layout.folderRect.left + 12.0f;
        const float rowRight = std::max(rowLeft + 60.0f, layout.folderRect.right - 250.0f);
        float rowTop         = layout.folderRect.top + 48.0f;
        layout.hoverRowRect  = D2D1::RectF(rowLeft, rowTop, rowRight, rowTop + 30.0f);
        rowTop += 34.0f;
        layout.selectedRowRect = D2D1::RectF(rowLeft, rowTop, rowRight, rowTop + 30.0f);
        rowTop += 34.0f;
        layout.warningRowRect = D2D1::RectF(rowLeft, rowTop, rowRight, rowTop + 30.0f);

        layout.dialogRect = D2D1::RectF(std::max(rowRight + 16.0f, layout.folderRect.right - 230.0f),
                                        layout.folderRect.top + 48.0f,
                                        layout.folderRect.right - 12.0f,
                                        layout.folderRect.top + 154.0f);
        layout.buttonRect =
            D2D1::RectF(layout.dialogRect.left + 14.0f, layout.dialogRect.bottom - 42.0f, layout.dialogRect.left + 112.0f, layout.dialogRect.bottom - 12.0f);

        layout.progressBgRect   = D2D1::RectF(rowLeft,
                                              std::max(layout.warningRowRect.bottom + 44.0f, layout.folderRect.bottom - 104.0f),
                                              rowRight,
                                              std::max(layout.warningRowRect.bottom + 56.0f, layout.folderRect.bottom - 92.0f));
        layout.progressFillRect = D2D1::RectF(layout.progressBgRect.left,
                                              layout.progressBgRect.top,
                                              layout.progressBgRect.left + ((layout.progressBgRect.right - layout.progressBgRect.left) * 0.62f),
                                              layout.progressBgRect.bottom);
        layout.diffAddedRect    = D2D1::RectF(rowLeft, layout.progressBgRect.bottom + 22.0f, rowRight, layout.progressBgRect.bottom + 48.0f);
        layout.diffRemovedRect  = D2D1::RectF(rowLeft, layout.diffAddedRect.bottom + 4.0f, rowRight, layout.diffAddedRect.bottom + 30.0f);
        return layout;
    }

    const RedConfigure::Themes::ThemePreviewModel* _model = nullptr;
    std::function<void(std::wstring_view)> _onTokenSelected;
    std::wstring _selectedTokenKey;
    D2D1_POINT_2F _lastClickPoint = D2D1::Point2F(-10000.0f, -10000.0f);
    std::wstring _lastClickKey;
    std::wstring _navText;
    std::wstring _menuText;
    std::wstring _folderText;
    std::wstring _hoverText;
    std::wstring _selectedText;
    std::wstring _dialogText;
    std::wstring _buttonText;
    std::wstring _progressText;
    std::wstring _warningText;
    std::wstring _diffAddedText;
    std::wstring _diffRemovedText;
};
} // namespace RedConfigure::Ui
