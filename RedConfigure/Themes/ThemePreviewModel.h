#pragma once

#include "SettingsStore.h"

#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

namespace RedConfigure::Themes
{
enum class ThemeSourceTransform : uint8_t
{
    Darken10,
    BlendAccent16,
};

class ThemePreviewModel final
{
public:
    void SetTheme(const Common::Settings::ThemeDefinition& theme);

    [[nodiscard]] const Common::Settings::ThemeDefinition& GetTheme() const noexcept;
    [[nodiscard]] std::optional<uint32_t> GetEffectiveColor(std::wstring_view key) const;
    [[nodiscard]] std::wstring GetAuthoredColorText(std::wstring_view key) const;
    [[nodiscard]] bool TryEditOverride(std::wstring_view key, std::wstring_view colorText);
    [[nodiscard]] bool CreatePaletteEntry(std::wstring_view name, std::wstring_view colorText, bool replaceMatchingDirectSources);
    [[nodiscard]] bool WrapSourceWithTransform(std::wstring_view key, ThemeSourceTransform transform);
    [[nodiscard]] bool ResetOverride(std::wstring_view key);
    [[nodiscard]] bool RenamePaletteEntry(std::wstring_view oldName, std::wstring_view newName);
    [[nodiscard]] std::vector<std::wstring> GetDependencies(std::wstring_view key) const;
    [[nodiscard]] std::vector<std::wstring> GetAffected(std::wstring_view key) const;
    [[nodiscard]] Common::Settings::ThemeColorEvaluationPhase GetEvaluationPhase(std::wstring_view key) const;
    [[nodiscard]] std::optional<Common::Settings::ThemeColorSourceKind> GetSourceKind(std::wstring_view key) const noexcept;
    [[nodiscard]] std::wstring_view GetLastError() const noexcept;
    void SetPreviewSeed(uint32_t seed) noexcept;
    [[nodiscard]] uint32_t GetPreviewSeed() const noexcept;

private:
    [[nodiscard]] bool Recompute();

    Common::Settings::ThemeDefinition _theme;
    Common::Settings::ResolvedThemeColors _resolved;
    std::wstring _lastError;
    uint32_t _previewSeed = 0x52ED5EEDu;
};
} // namespace RedConfigure::Themes
