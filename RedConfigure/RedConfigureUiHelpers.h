#pragma once

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <Windows.h>

#include "DxUi.h"
#include "Localization/PlaceholderValidation.h"
#include "Localization/RcParser.h"
#include "Themes/ThemePreviewModel.h"

#include <cstdint>
#include <string>
#include <string_view>

namespace RedConfigure::Ui
{
[[nodiscard]] std::wstring LoadAppString(HINSTANCE instance, UINT resourceId);
[[nodiscard]] D2D1_COLOR_F ColorFromArgb(uint32_t argb) noexcept;
[[nodiscard]] uint32_t ColorOrDefault(const RedConfigure::Themes::ThemePreviewModel& model, std::wstring_view key, uint32_t fallback);
[[nodiscard]] std::wstring PlaceholderStatusText(HINSTANCE instance, RedConfigure::Localization::PlaceholderStatus status);
[[nodiscard]] std::wstring LocalizableKindText(HINSTANCE instance, RedConfigure::Localization::RcLocalizableKind kind);
}
