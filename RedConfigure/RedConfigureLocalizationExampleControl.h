#pragma once

#include "DxUi.h"

#include <string>
#include <string_view>

namespace RedConfigure::Ui
{
class LocalizationExampleControl final : public RedSalamander::DxUi::Control
{
public:
    enum class Kind : uint8_t
    {
        String,
        FormattedString,
        CommandLabel,
        Menu,
        Dialog,
    };

    void SetExample(Kind kind, std::wstring_view source, std::wstring_view target)
    {
        _kind   = kind;
        _source = source;
        _target = target;
        RequestInvalidate();
    }

    void Paint(RedSalamander::DxUi::WindowHost& host) const override
    {
        auto* dc = host.GetDeviceContext();
        if (! dc)
        {
            return;
        }
        const D2D1_RECT_F bounds = GetBounds();
        const auto& palette      = host.GetTheme();
        if (auto* brush = host.GetSolidBrush(palette.cardBackground))
        {
            dc->FillRoundedRectangle(D2D1::RoundedRect(bounds, 6.0f, 6.0f), brush);
        }
        if (auto* brush = host.GetSolidBrush(palette.borderDefault))
        {
            dc->DrawRoundedRectangle(D2D1::RoundedRect(bounds, 6.0f, 6.0f), brush, 1.0f);
        }

        std::wstring sample = _target.empty() ? _source : _target;
        size_t placeholder  = sample.find(L"{0}");
        if (placeholder != std::wstring::npos)
        {
            sample.replace(placeholder, 3u, L"42");
        }
        const auto draw = [&](std::wstring_view text, const D2D1_RECT_F& rect, RedSalamander::DxUi::FontRole role, const D2D1_COLOR_F& color)
        {
            if (auto* brush = host.GetSolidBrush(color))
            {
                dc->DrawTextW(text.data(), static_cast<UINT32>(text.size()), host.GetTextFormat(role), rect, brush);
            }
        };
        draw(sample,
             D2D1::RectF(bounds.left + 12.0f, bounds.top + 18.0f, bounds.right - 12.0f, bounds.bottom - 10.0f),
             RedSalamander::DxUi::FontRole::BodyStrong,
             palette.text);
    }

private:
    Kind _kind = Kind::String;
    std::wstring _source;
    std::wstring _target;
};
} // namespace RedConfigure::Ui
