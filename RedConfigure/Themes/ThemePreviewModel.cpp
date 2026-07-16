#include "ThemePreviewModel.h"

#include "ThemeDefinitionIo.h"
#include "Helpers.h"

#include <array>
#include <chrono>
#include <format>

namespace
{
struct DefaultColor
{
    std::wstring_view key;
    uint32_t darkValue  = 0u;
    uint32_t lightValue = 0u;
};

constexpr std::array<DefaultColor, 26> kDefaultColors = {{
    {L"app.accent", 0xFF0078D4u, 0xFF005FB8u},
    {L"window.background", 0xFF202020u, 0xFFFFFFFFu},
    {L"navigation.background", 0xFF252525u, 0xFFEAF2FEu},
    {L"navigation.text", 0xFFF3F3F3u, 0xFF1F2937u},
    {L"navigation.accent", 0xFF0078D4u, 0xFF005FB8u},
    {L"navigation.progressOk", 0xFF22C55Eu, 0xFF15803Du},
    {L"navigation.progressBackground", 0xFF3A3A3Au, 0xFFE5E7EBu},
    {L"menu.background", 0xFF2B2B2Bu, 0xFFFFFFFFu},
    {L"menu.text", 0xFFF4F4F4u, 0xFF111111u},
    {L"menu.disabledText", 0xFFB8B8B8u, 0xFF5B5B5Bu},
    {L"menu.selectionBg", 0xFF3B3B3Bu, 0xFFE8F1FFu},
    {L"menu.selectionText", 0xFFFFFFFFu, 0xFF0F172Au},
    {L"menu.border", 0xFF505050u, 0xFFD8D8D8u},
    {L"folderView.background", 0xFF1E1E1Eu, 0xFFFFFFFFu},
    {L"folderView.textNormal", 0xFFEDEDEDu, 0xFF111111u},
    {L"folderView.itemBackgroundHovered", 0xFF333333u, 0xFFF0F6FFu},
    {L"folderView.itemBackgroundSelected", 0xFF264F78u, 0xFFCFE8FFu},
    {L"folderView.textSelected", 0xFFFFFFFFu, 0xFF0F172Au},
    {L"folderView.warningBackground", 0xFF4A3512u, 0xFFFFF4CEu},
    {L"folderView.warningText", 0xFFFFD166u, 0xFF8A4B00u},
    {L"fileOps.progressBackground", 0xFF3A3A3Au, 0xFFE5E7EBu},
    {L"fileOps.progressTotal", 0xFF0078D4u, 0xFF2563EBu},
    {L"viewer.diff.addedBackground", 0xFF173B24u, 0xFFEAF8EFu},
    {L"viewer.diff.removedBackground", 0xFF4A1F25u, 0xFFFFECEFu},
    {L"monitor.textView.bg", 0xFF1E1E1Eu, 0xFFFFFFFFu},
    {L"monitor.textView.fg", 0xFFF3F3F3u, 0xFF111111u},
}};

[[nodiscard]] bool IsLightBase(std::wstring_view baseThemeId) noexcept
{
    return baseThemeId == L"builtin/light" || baseThemeId == L"builtin/system";
}

[[nodiscard]] std::optional<std::wstring_view> PaletteNameFromEditorKey(std::wstring_view key) noexcept
{
    constexpr std::wstring_view kPrefix = L"palette.";
    if (! key.starts_with(kPrefix) || key.size() <= kPrefix.size()) return std::nullopt;
    return key.substr(kPrefix.size());
}
} // namespace

namespace RedConfigure::Themes
{
void ThemePreviewModel::SetTheme(const Common::Settings::ThemeDefinition& theme)
{
    _theme = theme;
    static_cast<void>(Recompute());
}

const Common::Settings::ThemeDefinition& ThemePreviewModel::GetTheme() const noexcept
{
    return _theme;
}

std::optional<uint32_t> ThemePreviewModel::GetEffectiveColor(std::wstring_view key) const
{
    if (const std::optional<std::wstring_view> paletteName = PaletteNameFromEditorKey(key); paletteName.has_value())
    {
        const auto found = _resolved.paletteColors.find(std::wstring(paletteName.value()));
        return found == _resolved.paletteColors.end() ? std::nullopt : std::optional<uint32_t>(found->second);
    }
    if (const auto dynamic = _resolved.dynamicColors.find(std::wstring(key)); dynamic != _resolved.dynamicColors.end())
    {
        return Common::Settings::EvaluateDynamicThemeColor(
            dynamic->second, Common::Settings::ThemeRuntimeContext{.seedHash32 = _previewSeed, .highContrast = false});
    }
    const auto found = _resolved.colors.find(std::wstring(key));
    if (found != _resolved.colors.end())
    {
        return found->second;
    }
    const bool light = IsLightBase(_theme.baseThemeId);
    for (const DefaultColor& color : kDefaultColors)
    {
        if (color.key == key)
        {
            return light ? color.lightValue : color.darkValue;
        }
    }
    return std::nullopt;
}

std::wstring ThemePreviewModel::GetAuthoredColorText(std::wstring_view key) const
{
    if (const std::optional<std::wstring_view> paletteName = PaletteNameFromEditorKey(key); paletteName.has_value())
    {
        const auto found = _theme.palette.find(std::wstring(paletteName.value()));
        return found == _theme.palette.end() ? std::wstring{} : Common::Settings::FormatThemeColorSource(found->second);
    }
    const auto found = _theme.colors.find(std::wstring(key));
    return found == _theme.colors.end() ? std::wstring{} : Common::Settings::FormatThemeColorSource(found->second);
}

bool ThemePreviewModel::TryEditOverride(std::wstring_view key, std::wstring_view colorText)
{
    const std::optional<std::wstring_view> paletteName = PaletteNameFromEditorKey(key);
    if ((! paletteName.has_value() && ! Common::Settings::IsValidThemeColorKey(key)) ||
        (paletteName.has_value() && ! Common::Settings::IsValidThemePaletteName(paletteName.value())))
    {
        _lastError = L"The theme key or palette name is invalid.";
        return false;
    }
    Common::Settings::ThemeColorSource source;
    if (FAILED(Common::Settings::ParseThemeColorSource(colorText, source, &_lastError)))
    {
        return false;
    }
    const std::wstring keyText(key);
    auto& destination = paletteName.has_value() ? _theme.palette : _theme.colors;
    const std::wstring destinationKey = paletteName.has_value() ? std::wstring(paletteName.value()) : keyText;
    const auto previous = destination.find(destinationKey);
    const std::optional<Common::Settings::ThemeColorSource> previousSource =
        previous == destination.end() ? std::nullopt : std::optional<Common::Settings::ThemeColorSource>(previous->second);
    destination[destinationKey] = std::move(source);
    if (Recompute())
    {
        return true;
    }
    const std::wstring failedError = _lastError;
    if (previousSource.has_value())
    {
        destination[destinationKey] = previousSource.value();
    }
    else
    {
        destination.erase(destinationKey);
    }
    static_cast<void>(Recompute());
    _lastError = failedError;
    return false;
}

bool ThemePreviewModel::CreatePaletteEntry(std::wstring_view name, std::wstring_view colorText, bool replaceMatchingDirectSources)
{
    if (! Common::Settings::IsValidThemePaletteName(name) || _theme.palette.contains(std::wstring(name)))
    {
        _lastError = L"The palette name is invalid or already exists.";
        return false;
    }

    Common::Settings::ThemeColorSource source;
    if (FAILED(Common::Settings::ParseThemeColorSource(colorText, source, &_lastError)))
    {
        return false;
    }

    Common::Settings::ThemeDefinition candidate = _theme;
    candidate.palette.emplace(std::wstring(name), source);
    if (replaceMatchingDirectSources && source.kind == Common::Settings::ThemeColorSourceKind::Direct)
    {
        Common::Settings::ThemeColorSource reference;
        reference.kind = Common::Settings::ThemeColorSourceKind::Reference;
        reference.references.push_back(std::wstring(L"palette.") + std::wstring(name));
        for (auto& [paletteName, value] : candidate.palette)
        {
            if (paletteName != name && value == source) value = reference;
        }
        for (auto& [_, value] : candidate.colors)
        {
            if (value == source) value = reference;
        }
    }

    const Common::Settings::ThemeDefinition previous = _theme;
    _theme = std::move(candidate);
    if (Recompute()) return true;
    const std::wstring failedError = _lastError;
    _theme = previous;
    static_cast<void>(Recompute());
    _lastError = failedError;
    return false;
}

bool ThemePreviewModel::WrapSourceWithTransform(std::wstring_view key, ThemeSourceTransform transform)
{
    const std::optional<std::wstring_view> paletteName = PaletteNameFromEditorKey(key);
    if ((! paletteName.has_value() && ! Common::Settings::IsValidThemeColorKey(key)) ||
        (paletteName.has_value() && ! Common::Settings::IsValidThemePaletteName(paletteName.value())) ||
        (transform == ThemeSourceTransform::BlendAccent16 && key == L"app.accent"))
    {
        _lastError = L"The selected theme source cannot use this transform.";
        return false;
    }

    const auto& sources = paletteName.has_value() ? _theme.palette : _theme.colors;
    const std::wstring sourceKey = paletteName.has_value() ? std::wstring(paletteName.value()) : std::wstring(key);
    Common::Settings::ThemeColorSource original;
    if (const auto found = sources.find(sourceKey); found != sources.end())
    {
        original = found->second;
    }
    else if (const std::optional<uint32_t> effective = GetEffectiveColor(key); effective.has_value())
    {
        original = Common::Settings::ThemeColorSource(effective.value());
    }
    else
    {
        _lastError = L"The selected theme source has no effective color to transform.";
        return false;
    }

    std::wstring stem = L"source_";
    for (const wchar_t ch : sourceKey)
    {
        const bool alphaNumeric = (ch >= L'a' && ch <= L'z') || (ch >= L'A' && ch <= L'Z') || (ch >= L'0' && ch <= L'9');
        stem.push_back(alphaNumeric ? ch : L'_');
        if (stem.size() >= 52u) break;
    }
    std::wstring generatedName = stem;
    for (uint32_t suffix = 2u; _theme.palette.contains(generatedName); ++suffix)
    {
        generatedName = std::format(L"{}_{}", stem, suffix);
    }

    Common::Settings::ThemeDefinition candidate = _theme;
    candidate.palette.emplace(generatedName, std::move(original));
    const std::wstring expression = transform == ThemeSourceTransform::Darken10
                                        ? std::format(L"darken(palette.{},10%)", generatedName)
                                        : std::format(L"blend(palette.{0},app.accent,16%)", generatedName);
    Common::Settings::ThemeColorSource transformed;
    if (FAILED(Common::Settings::ParseThemeColorSource(expression, transformed, &_lastError))) return false;
    auto& candidateSources = paletteName.has_value() ? candidate.palette : candidate.colors;
    candidateSources[sourceKey] = std::move(transformed);

    const Common::Settings::ThemeDefinition previous = _theme;
    _theme = std::move(candidate);
    if (Recompute()) return true;
    const std::wstring failedError = _lastError;
    _theme = previous;
    static_cast<void>(Recompute());
    _lastError = failedError;
    return false;
}

bool ThemePreviewModel::ResetOverride(std::wstring_view key)
{
    if (const std::optional<std::wstring_view> paletteName = PaletteNameFromEditorKey(key); paletteName.has_value())
    {
        const std::wstring dependencyKey = std::wstring(L"palette.") + std::wstring(paletteName.value());
        if (const auto affected = _resolved.affected.find(dependencyKey); affected != _resolved.affected.end() && ! affected->second.empty())
        {
            _lastError = std::format(L"Palette entry '{}' is still referenced by {} theme source(s).", paletteName.value(), affected->second.size());
            return false;
        }
        _theme.palette.erase(std::wstring(paletteName.value()));
        return Recompute();
    }
    _theme.colors.erase(std::wstring(key));
    return Recompute();
}

bool ThemePreviewModel::RenamePaletteEntry(std::wstring_view oldName, std::wstring_view newName)
{
    if (! Common::Settings::IsValidThemePaletteName(oldName) || ! Common::Settings::IsValidThemePaletteName(newName) || oldName == newName ||
        ! _theme.palette.contains(std::wstring(oldName)) || _theme.palette.contains(std::wstring(newName)))
    {
        _lastError = L"The palette rename is invalid or the destination name already exists.";
        return false;
    }

    Common::Settings::ThemeDefinition candidate = _theme;
    auto source = candidate.palette.extract(std::wstring(oldName));
    source.key() = std::wstring(newName);
    candidate.palette.insert(std::move(source));
    const std::wstring oldReference = std::wstring(L"palette.") + std::wstring(oldName);
    const std::wstring newReference = std::wstring(L"palette.") + std::wstring(newName);
    const auto rewrite = [&](auto& sources)
    {
        for (auto& [_, value] : sources)
        {
            for (std::wstring& reference : value.references)
            {
                if (reference == oldReference) reference = newReference;
            }
        }
    };
    rewrite(candidate.palette);
    rewrite(candidate.colors);

    const Common::Settings::ThemeDefinition previous = _theme;
    _theme = std::move(candidate);
    if (Recompute()) return true;
    const std::wstring failedError = _lastError;
    _theme = previous;
    static_cast<void>(Recompute());
    _lastError = failedError;
    return false;
}

std::vector<std::wstring> ThemePreviewModel::GetDependencies(std::wstring_view key) const
{
    const auto found = _resolved.dependencies.find(std::wstring(key));
    return found == _resolved.dependencies.end() ? std::vector<std::wstring>{} : found->second;
}

std::vector<std::wstring> ThemePreviewModel::GetAffected(std::wstring_view key) const
{
    const auto found = _resolved.affected.find(std::wstring(key));
    return found == _resolved.affected.end() ? std::vector<std::wstring>{} : found->second;
}

Common::Settings::ThemeColorEvaluationPhase ThemePreviewModel::GetEvaluationPhase(std::wstring_view key) const
{
    const auto phaseFor = [&](auto&& self, std::wstring_view sourceName, size_t depth) -> Common::Settings::ThemeColorEvaluationPhase
    {
        if (depth >= 32u) return Common::Settings::ThemeColorEvaluationPhase::Load;
        const std::optional<std::wstring_view> paletteName = PaletteNameFromEditorKey(sourceName);
        const auto& sources = paletteName.has_value() ? _theme.palette : _theme.colors;
        const std::wstring sourceKey = paletteName.has_value() ? std::wstring(paletteName.value()) : std::wstring(sourceName);
        const auto found = sources.find(sourceKey);
        Common::Settings::ThemeColorEvaluationPhase phase =
            found == sources.end() ? Common::Settings::ThemeColorEvaluationPhase::Load : Common::Settings::GetThemeColorEvaluationPhase(found->second);
        if (phase == Common::Settings::ThemeColorEvaluationPhase::Paint) return phase;
        if (const auto dependencies = _resolved.dependencies.find(std::wstring(sourceName)); dependencies != _resolved.dependencies.end())
        {
            for (const std::wstring& dependency : dependencies->second)
            {
                const auto dependencyPhase = self(self, dependency, depth + 1u);
                if (dependencyPhase == Common::Settings::ThemeColorEvaluationPhase::Paint) return dependencyPhase;
                if (dependencyPhase == Common::Settings::ThemeColorEvaluationPhase::Event) phase = dependencyPhase;
            }
        }
        return phase;
    };
    return phaseFor(phaseFor, key, 0u);
}

std::optional<Common::Settings::ThemeColorSourceKind> ThemePreviewModel::GetSourceKind(std::wstring_view key) const noexcept
{
    const std::optional<std::wstring_view> paletteName = PaletteNameFromEditorKey(key);
    const auto& sources = paletteName.has_value() ? _theme.palette : _theme.colors;
    const std::wstring sourceKey = paletteName.has_value() ? std::wstring(paletteName.value()) : std::wstring(key);
    const auto found = sources.find(sourceKey);
    return found == sources.end() ? std::nullopt : std::optional<Common::Settings::ThemeColorSourceKind>(found->second.kind);
}

std::wstring_view ThemePreviewModel::GetLastError() const noexcept
{
    return _lastError;
}

void ThemePreviewModel::SetPreviewSeed(uint32_t seed) noexcept
{
    _previewSeed = seed;
}

uint32_t ThemePreviewModel::GetPreviewSeed() const noexcept
{
    return _previewSeed;
}

bool ThemePreviewModel::Recompute()
{
    const auto startedAt = std::chrono::steady_clock::now();
    Common::Settings::ThemeResolutionContext context;
    const bool light = IsLightBase(_theme.baseThemeId);
    context.effectiveDark = ! light;
    context.baseColor = [light](std::wstring_view key) -> std::optional<uint32_t>
    {
        for (const DefaultColor& color : kDefaultColors)
        {
            if (color.key == key)
            {
                return light ? color.lightValue : color.darkValue;
            }
        }
        return std::nullopt;
    };
    context.systemColors = {{0xFF0078D4u, 0xFF60A5FAu, 0xFF005FB8u, light ? 0xFFFFFFFFu : 0xFF202020u,
                             light ? 0xFF111111u : 0xFFF3F3F3u, 0xFF0078D4u, 0xFFFFFFFFu}};
    Common::Settings::ResolvedThemeColors candidate;
    std::wstring message;
    if (FAILED(Common::Settings::ResolveThemeDefinition(_theme, context, candidate, &message)))
    {
        Debug::Perf::EmitDurationUs(L"redconfigure.theme.preview_resolve_us",
                                    Debug::Perf::ElapsedUs(startedAt),
                                    static_cast<uint64_t>(_theme.palette.size()),
                                    static_cast<uint64_t>(_theme.colors.size()),
                                    HRESULT_FROM_WIN32(ERROR_INVALID_DATA));
        _lastError = std::move(message);
        return false;
    }
    _resolved = std::move(candidate);
    Debug::Perf::EmitDurationUs(L"redconfigure.theme.preview_resolve_us",
                                Debug::Perf::ElapsedUs(startedAt),
                                static_cast<uint64_t>(_theme.palette.size()),
                                static_cast<uint64_t>(_theme.colors.size()),
                                S_OK);
    _lastError.clear();
    return true;
}
} // namespace RedConfigure::Themes
