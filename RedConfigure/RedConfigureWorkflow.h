#pragma once

#include "RedConfigureSession.h"

#include <cstdint>
#include <filesystem>
#include <optional>
#include <span>
#include <string>
#include <string_view>
#include <vector>

namespace RedConfigure::Workflow
{
enum class BatchApprovalResult : uint8_t
{
    NoChanges,
    Ready,
    Stale,
    Invalid,
    Applied,
};

enum class ValidationSeverity : uint8_t
{
    Warning,
    Error,
};

enum class ValidationCategory : uint8_t
{
    Workspace,
    Theme,
    Localization,
    Export,
    Accelerator,
};

enum class ValidationCode : uint8_t
{
    WorkspaceProcessingError,
    ThemeCatalogError,
    PlaceholderMismatch,
    MissingTranslation,
    InvalidThemeId,
    ThemeResolutionError,
    EmptyOutputPath,
    OutputConflict,
    LocalizationPreviewBuildFailed,
    DuplicateLocalizationOutputPath,
    LocalizationThemeOutputConflict,
    DuplicateAccelerator,
};

struct ValidationIssue
{
    ValidationSeverity severity = ValidationSeverity::Error;
    ValidationCategory category = ValidationCategory::Workspace;
    ValidationCode code = ValidationCode::WorkspaceProcessingError;
    std::vector<std::wstring> arguments;
    std::wstring ownerName;
    std::wstring resourceId;
    std::wstring cultureName;
};

struct ValidationSummary
{
    std::vector<ValidationIssue> issues;
    size_t errorCount   = 0u;
    size_t warningCount = 0u;

    [[nodiscard]] bool CanExport() const noexcept
    {
        return errorCount == 0u;
    }
};

struct LanguageColumn
{
    std::wstring cultureName;
    bool pinned = false;
};

class LanguageColumnModel final
{
public:
    void Set(std::span<const std::wstring> cultures);
    [[nodiscard]] bool Add(std::wstring_view cultureName);
    [[nodiscard]] bool Remove(std::wstring_view cultureName);
    [[nodiscard]] bool Move(std::wstring_view cultureName, size_t newIndex);
    [[nodiscard]] bool SetPinned(std::wstring_view cultureName, bool pinned);
    [[nodiscard]] std::span<const LanguageColumn> Get() const noexcept;
    [[nodiscard]] std::vector<std::wstring> GetOrderedCultures() const;

private:
    void NormalizePinnedOrder();
    std::vector<LanguageColumn> _columns;
};

enum class LocalizationBatchKind : uint8_t
{
    CopyEnglish,
    CopyCulture,
    Clear,
    FindReplace,
    NormalizePlaceholderWhitespace,
    PreserveAccelerators,
    MarkReviewed,
};

struct LocalizationBatchRequest
{
    LocalizationBatchKind kind = LocalizationBatchKind::CopyEnglish;
    std::wstring sourceCulture;
    std::wstring targetCulture;
    std::wstring findText;
    std::wstring replaceText;
    std::vector<size_t> rowIndices;

    bool operator==(const LocalizationBatchRequest&) const = default;
};

struct LocalizationBatchChange
{
    std::wstring ownerName;
    std::wstring resourceId;
    std::wstring cultureName;
    std::wstring before;
    std::wstring after;
    bool beforeReviewed = false;
    bool afterReviewed  = false;
};

struct LocalizationBatchPreview
{
    LocalizationBatchRequest request;
    std::vector<LocalizationBatchChange> changes;
    BatchApprovalResult result = BatchApprovalResult::NoChanges;
};

struct DuplicateAccelerator
{
    std::wstring ownerName;
    wchar_t accelerator = L'\0';
    std::vector<std::wstring> resourceIds;
};

struct ClipboardMatrix
{
    std::vector<std::vector<std::wstring>> rows;
};

enum class ThemeRecipe : uint8_t
{
    DarkVariant,
    LightVariant,
    AccentRecolor,
    SoftenedSelections,
    IncreasedContrast,
    SemanticStatusColors,
    SetAlpha,
    ReplaceReference,
    ConvertSolidsToReferences,
    RemoveOverrides,
};

struct ThemeMassRequest
{
    ThemeRecipe recipe = ThemeRecipe::DarkVariant;
    std::vector<std::wstring> keys;
    std::wstring argument;
    uint32_t alphaPercent = 80u;

    bool operator==(const ThemeMassRequest&) const = default;
};

struct ThemeMassChange
{
    std::wstring key;
    std::wstring before;
    std::wstring after;
    std::optional<Common::Settings::ThemeColorSource> beforeSource;
    std::optional<Common::Settings::ThemeColorSource> afterSource;
    std::wstring sourcePaletteName;
    std::wstring sourceValue;
    std::optional<Common::Settings::ThemeColorSource> sourceValueSource;
};

enum class ThemeMassDiagnostic : uint8_t
{
    None,
    InvalidAlpha,
    MalformedReferenceReplacement,
    MissingPaletteTarget,
    InvalidExistingSource,
    InvalidGeneratedSource,
    InvalidCandidate,
};

struct ThemeMassPreview
{
    ThemeMassRequest request;
    std::vector<ThemeMassChange> changes;
    Common::Settings::ThemeDefinition beforeTheme;
    BatchApprovalResult result = BatchApprovalResult::NoChanges;
    ThemeMassDiagnostic diagnostic = ThemeMassDiagnostic::None;
};

struct DuplicateThemeCandidate
{
    std::wstring id;
    std::wstring name;
};

enum class ThemeTokenSourceKind : uint8_t
{
    Inherited,
    Literal,
    Reference,
    Function,
};

struct ThemeTokenMetadata
{
    std::wstring key;
    std::wstring group;
    ThemeTokenSourceKind sourceKind = ThemeTokenSourceKind::Inherited;
    size_t usageCount = 0u;
    double contrastRatio = 0.0;
    bool contrastKnown = false;
    bool contrastPass  = false;
};

[[nodiscard]] size_t CountDirtyLocalizationCells(std::span<const LocalizationReviewRow> rows) noexcept;
[[nodiscard]] ValidationSummary ValidateWorkspace(const RedConfigureSession& session,
                                                  const std::filesystem::path& localizationOutputRoot,
                                                  const std::filesystem::path& themeOutputPath);
[[nodiscard]] LocalizationBatchPreview PreviewLocalizationBatch(std::span<const LocalizationReviewRow> rows,
                                                                const LocalizationBatchRequest& request);
[[nodiscard]] std::vector<DuplicateAccelerator> FindDuplicateSiblingAccelerators(std::span<const LocalizationReviewRow> rows,
                                                                                 std::wstring_view cultureName);
[[nodiscard]] ClipboardMatrix ParseClipboardMatrix(std::wstring_view text);
[[nodiscard]] std::wstring SerializeClipboardMatrix(const ClipboardMatrix& matrix);
[[nodiscard]] ThemeMassPreview PreviewThemeMassChange(const Themes::ThemePreviewModel& model, const ThemeMassRequest& request);
[[nodiscard]] BatchApprovalResult ApplyThemeMassChange(Themes::ThemePreviewModel& model, const ThemeMassPreview& preview);
[[nodiscard]] std::optional<DuplicateThemeCandidate> BuildDuplicateThemeCandidate(std::wstring_view sourceId,
                                                                                  std::wstring_view sourceName,
                                                                                  std::wstring_view localizedCopyLabel,
                                                                                  uint32_t sequence);
[[nodiscard]] ThemeTokenMetadata BuildThemeTokenMetadata(const Themes::ThemePreviewModel& model, std::wstring_view key);
} // namespace RedConfigure::Workflow
