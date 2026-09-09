#pragma once

#include <cstddef>

#include "SettingsStore.h"

namespace ShortcutDefaults
{
struct WindowsTerminalReviewCoverage final
{
    size_t p1Mapped            = 48u;
    size_t deferredPassThrough = 12u;
    size_t notApplicable       = 8u;

    [[nodiscard]] constexpr size_t Total() const noexcept
    {
        return p1Mapped + deferredPassThrough + notApplicable;
    }
};

inline constexpr WindowsTerminalReviewCoverage kWindowsTerminalReviewCoverage{};
static_assert(kWindowsTerminalReviewCoverage.Total() == 68u);

[[nodiscard]] Common::Settings::ShortcutsSettings CreateDefaultShortcuts();

[[nodiscard]] bool AreShortcutsDefault(const Common::Settings::ShortcutsSettings& shortcuts);

[[nodiscard]] bool IsDefaultFunctionBarBinding(const Common::Settings::ShortcutBinding& binding);

[[nodiscard]] bool IsDefaultFolderViewBinding(const Common::Settings::ShortcutBinding& binding);

[[nodiscard]] bool IsDefaultApplicationBinding(const Common::Settings::ShortcutBinding& binding);

[[nodiscard]] bool IsDefaultTerminalBinding(const Common::Settings::ShortcutBinding& binding);

void EnsureShortcutsInitialized(Common::Settings::Settings& settings);
} // namespace ShortcutDefaults
