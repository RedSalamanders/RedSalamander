#pragma once

#include "Localization/PlaceholderValidation.h"
#include "Localization/RcParser.h"
#include "Themes/ThemeCatalog.h"
#include "Themes/ThemePreviewModel.h"
#include "Workspace/WorkspaceDiscovery.h"

#include <filesystem>
#include <span>
#include <string>
#include <string_view>
#include <vector>

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <Windows.h>

namespace RedConfigure
{
struct TranslationEntry
{
    std::wstring id;
    std::wstring sourceText;
    std::wstring targetText;
    Localization::PlaceholderValidationResult validation;
};

struct InventoryEntry
{
    Localization::RcLocalizableKind kind = Localization::RcLocalizableKind::StringTable;
    std::wstring ownerName;
    std::wstring resourceId;
    std::wstring itemId;
    std::wstring sourceText;
    size_t sourceLine = 0u;
};

enum class LocalizationViewColumn : uint8_t
{
    Id,
    Source,
    Target,
    Status,
};

enum class LocalizationSortDirection : uint8_t
{
    None,
    Ascending,
    Descending,
};

enum class LocalizationStatusFilter : uint8_t
{
    All,
    Ok,
    Problems,
};

struct LocalizationViewOptions
{
    std::wstring searchText;
    std::wstring idFilterText;
    LocalizationStatusFilter statusFilter = LocalizationStatusFilter::All;
    LocalizationViewColumn sortColumn     = LocalizationViewColumn::Id;
    LocalizationSortDirection sortDirection = LocalizationSortDirection::None;
};

[[nodiscard]] std::vector<size_t> BuildTranslationView(std::span<const TranslationEntry> translations, const LocalizationViewOptions& options);

class RedConfigureSession final
{
public:
    [[nodiscard]] HRESULT LoadWorkspace(const std::filesystem::path& root, std::wstring cultureName);

    [[nodiscard]] const Workspace::WorkspaceScanResult& GetWorkspace() const noexcept;
    [[nodiscard]] const Themes::ThemeCatalog& GetThemeCatalog() const noexcept;
    [[nodiscard]] const Themes::ThemePreviewModel& GetThemePreviewModel() const noexcept;
    [[nodiscard]] Themes::ThemePreviewModel& GetThemePreviewModel() noexcept;
    [[nodiscard]] std::span<const TranslationEntry> GetTranslations() const noexcept;
    [[nodiscard]] std::span<const InventoryEntry> GetInventoryEntries() const noexcept;
    [[nodiscard]] std::wstring_view GetCultureName() const noexcept;
    [[nodiscard]] std::wstring_view GetActiveResourceOwnerName() const noexcept;
    [[nodiscard]] size_t GetActiveResourceOwnerIndex() const noexcept;
    [[nodiscard]] size_t GetActiveThemeIndex() const noexcept;

    [[nodiscard]] HRESULT SetActiveResourceOwner(size_t ownerIndex);
    [[nodiscard]] bool SetActiveTheme(size_t themeIndex);
    [[nodiscard]] bool UpdateTranslation(size_t rowIndex, std::wstring_view targetText);
    [[nodiscard]] bool UpdateThemeColor(std::wstring_view colorKey, std::wstring_view colorText);

    [[nodiscard]] std::filesystem::path GetDefaultLocalizationExportPath() const;
    [[nodiscard]] std::filesystem::path GetDefaultThemeExportPath() const;
    [[nodiscard]] HRESULT BuildLocalizationExportText(std::wstring& outText) const;
    [[nodiscard]] HRESULT BuildThemeExportText(std::string& outJson) const;
    [[nodiscard]] HRESULT ExportLocalization(const std::filesystem::path& path) const;
    [[nodiscard]] HRESULT ExportTheme(const std::filesystem::path& path) const;

private:
    [[nodiscard]] HRESULT LoadLocalizationForActiveOwner();

    Workspace::WorkspaceScanResult _workspace;
    Themes::ThemeCatalog _themeCatalog;
    Themes::ThemePreviewModel _themePreview;
    std::vector<TranslationEntry> _translations;
    std::vector<InventoryEntry> _inventoryEntries;
    std::wstring _cultureName = L"en-US";
    std::wstring _activeResourceOwnerName;
    size_t _activeResourceOwnerIndex = 0u;
    size_t _activeThemeIndex         = 0u;
};
} // namespace RedConfigure
