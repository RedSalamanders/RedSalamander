#include "RedConfigureWorkflow.h"

#include "Helpers.h"
#include "SettingsStore.h"
#include "ThemeDefinitionIo.h"

#include <algorithm>
#include <cmath>
#include <cwctype>
#include <format>
#include <map>
#include <optional>
#include <set>
#include <unordered_map>

namespace
{
[[nodiscard]] bool EqualIgnoreCase(std::wstring_view lhs, std::wstring_view rhs) noexcept
{
    return lhs.size() == rhs.size() &&
           ::CompareStringOrdinal(lhs.data(), static_cast<int>(lhs.size()), rhs.data(), static_cast<int>(rhs.size()), TRUE) == CSTR_EQUAL;
}

[[nodiscard]] const RedConfigure::LocalizationTargetCell* FindCell(const RedConfigure::LocalizationReviewRow& row, std::wstring_view cultureName) noexcept
{
    for (const auto& cell : row.targets)
    {
        if (EqualIgnoreCase(cell.cultureName, cultureName))
        {
            return &cell;
        }
    }
    return nullptr;
}

[[nodiscard]] std::wstring NormalizePlaceholderWhitespace(std::wstring_view text)
{
    std::wstring result;
    result.reserve(text.size());
    bool insidePlaceholder = false;
    bool pendingSpace      = false;
    for (const wchar_t ch : text)
    {
        if (ch == L'{')
        {
            insidePlaceholder = true;
            pendingSpace      = false;
            result.push_back(ch);
        }
        else if (ch == L'}')
        {
            while (! result.empty() && result.back() == L' ')
            {
                result.pop_back();
            }
            result.push_back(ch);
            insidePlaceholder = false;
            pendingSpace      = false;
        }
        else if (insidePlaceholder && std::iswspace(ch) != 0)
        {
            pendingSpace = ! result.empty() && result.back() != L'{';
        }
        else
        {
            if (pendingSpace)
            {
                result.push_back(L' ');
                pendingSpace = false;
            }
            result.push_back(ch);
        }
    }
    return result;
}

[[nodiscard]] std::optional<wchar_t> Accelerator(std::wstring_view text) noexcept
{
    for (size_t index = 0u; index + 1u < text.size(); ++index)
    {
        if (text[index] != L'&')
        {
            continue;
        }
        if (text[index + 1u] == L'&')
        {
            ++index;
            continue;
        }
        return static_cast<wchar_t>(std::towlower(text[index + 1u]));
    }
    return std::nullopt;
}

[[nodiscard]] std::wstring PreserveSourceAccelerator(std::wstring_view source, std::wstring_view target)
{
    if (Accelerator(target).has_value())
    {
        return std::wstring(target);
    }
    const std::optional<wchar_t> sourceAccelerator = Accelerator(source);
    if (! sourceAccelerator.has_value())
    {
        return std::wstring(target);
    }
    std::wstring result(target);
    for (size_t index = 0u; index < result.size(); ++index)
    {
        if (std::towlower(result[index]) == sourceAccelerator.value())
        {
            result.insert(index, 1u, L'&');
            return result;
        }
    }
    return result;
}

[[nodiscard]] std::wstring Replaced(std::wstring_view text, std::wstring_view findText, std::wstring_view replacement)
{
    if (findText.empty())
    {
        return std::wstring(text);
    }
    std::wstring result(text);
    size_t offset = 0u;
    while ((offset = result.find(findText, offset)) != std::wstring::npos)
    {
        result.replace(offset, findText.size(), replacement);
        offset += replacement.size();
    }
    return result;
}

[[nodiscard]] std::wstring RecipeValue(
    RedConfigure::Workflow::ThemeRecipe recipe, std::wstring_view key, std::wstring_view before, std::wstring_view argument, uint32_t alphaPercent)
{
    using RedConfigure::Workflow::ThemeRecipe;
    const std::wstring_view operand = before.empty() ? key : before;
    switch (recipe)
    {
        case ThemeRecipe::DarkVariant: return std::format(L"darken({},18%)", operand);
        case ThemeRecipe::LightVariant: return std::format(L"lighten({},18%)", operand);
        case ThemeRecipe::AccentRecolor: return std::format(L"blend({},app.accent,32%)", operand);
        case ThemeRecipe::SoftenedSelections: return std::format(L"alpha({},72%)", operand);
        case ThemeRecipe::IncreasedContrast: return std::format(L"contrast({})", operand);
        case ThemeRecipe::SemanticStatusColors:
            if (key.find(L"warning") != std::wstring_view::npos)
                return L"ref(palette.warning)";
            if (key.find(L"error") != std::wstring_view::npos)
                return L"ref(palette.error)";
            if (key.find(L"success") != std::wstring_view::npos || key.find(L"progressOk") != std::wstring_view::npos)
                return L"ref(palette.success)";
            return std::wstring(before);
        case ThemeRecipe::SetAlpha: return std::format(L"alpha({},{}%)", operand, std::min(alphaPercent, 100u));
        case ThemeRecipe::ReplaceReference:
            return Replaced(before,
                            argument.substr(0u, argument.find(L'=')),
                            argument.find(L'=') == std::wstring_view::npos ? std::wstring_view{} : argument.substr(argument.find(L'=') + 1u));
        case ThemeRecipe::ConvertSolidsToReferences: return argument.empty() ? std::wstring(before) : std::format(L"ref(palette.{})", argument);
        case ThemeRecipe::RemoveOverrides: return {};
        default: return std::wstring(before);
    }
}

[[nodiscard]] bool RecipeNeedsCapturedSource(RedConfigure::Workflow::ThemeRecipe recipe) noexcept
{
    using RedConfigure::Workflow::ThemeRecipe;
    return recipe == ThemeRecipe::DarkVariant || recipe == ThemeRecipe::LightVariant || recipe == ThemeRecipe::AccentRecolor ||
           recipe == ThemeRecipe::SoftenedSelections || recipe == ThemeRecipe::IncreasedContrast || recipe == ThemeRecipe::SetAlpha;
}

[[nodiscard]] std::wstring RecipePaletteName(std::wstring_view key)
{
    std::wstring result = L"recipe_";
    for (const wchar_t ch : key)
    {
        const bool alphaNumeric = (ch >= L'a' && ch <= L'z') || (ch >= L'A' && ch <= L'Z') || (ch >= L'0' && ch <= L'9');
        result.push_back(alphaNumeric ? ch : L'_');
    }
    return result;
}

[[nodiscard]] bool ThemeDefinitionsEqual(const Common::Settings::ThemeDefinition& lhs, const Common::Settings::ThemeDefinition& rhs) noexcept
{
    return lhs.formatVersion == rhs.formatVersion && lhs.id == rhs.id && lhs.name == rhs.name && lhs.baseThemeId == rhs.baseThemeId &&
           lhs.palette == rhs.palette && lhs.colors == rhs.colors;
}

[[nodiscard]] bool BuildCandidateTheme(const RedConfigure::Workflow::ThemeMassPreview& preview,
                                       Common::Settings::ThemeDefinition& candidate,
                                       RedConfigure::Workflow::ThemeMassDiagnostic& diagnostic)
{
    candidate = preview.beforeTheme;
    for (const RedConfigure::Workflow::ThemeMassChange& change : preview.changes)
    {
        if (! change.sourcePaletteName.empty())
        {
            if (! change.sourceValueSource.has_value() || ! candidate.palette.emplace(change.sourcePaletteName, change.sourceValueSource.value()).second)
            {
                diagnostic = RedConfigure::Workflow::ThemeMassDiagnostic::InvalidGeneratedSource;
                return false;
            }
        }

        constexpr std::wstring_view palettePrefix = L"palette.";
        const bool isPalette                      = change.key.starts_with(palettePrefix);
        auto& destination                         = isPalette ? candidate.palette : candidate.colors;
        const std::wstring destinationKey(isPalette ? std::wstring_view(change.key).substr(palettePrefix.size()) : std::wstring_view(change.key));
        if (preview.request.recipe == RedConfigure::Workflow::ThemeRecipe::RemoveOverrides)
        {
            if (destination.erase(destinationKey) == 0u)
            {
                diagnostic = RedConfigure::Workflow::ThemeMassDiagnostic::InvalidCandidate;
                return false;
            }
        }
        else if (! change.afterSource.has_value())
        {
            diagnostic = RedConfigure::Workflow::ThemeMassDiagnostic::InvalidGeneratedSource;
            return false;
        }
        else
        {
            destination[destinationKey] = change.afterSource.value();
        }
    }
    return true;
}

using Common::Colors::ContrastRatioFromRelativeLuminance;
using Common::Colors::RelativeLuminanceFromArgb;

[[nodiscard]] std::optional<std::wstring_view> ContrastPeer(std::wstring_view key) noexcept
{
    static constexpr std::pair<std::wstring_view, std::wstring_view> pairs[] = {
        {L"navigation.text", L"navigation.background"},
        {L"menu.text", L"menu.background"},
        {L"menu.selectionText", L"menu.selectionBg"},
        {L"folderView.textNormal", L"folderView.background"},
        {L"folderView.textSelected", L"folderView.itemBackgroundSelected"},
        {L"folderView.warningText", L"folderView.warningBackground"},
        {L"monitor.textView.fg", L"monitor.textView.bg"},
    };
    for (const auto& [foreground, background] : pairs)
    {
        if (key == foreground)
            return background;
        if (key == background)
            return foreground;
    }
    return std::nullopt;
}
} // namespace

namespace RedConfigure::Workflow
{
void LanguageColumnModel::Set(std::span<const std::wstring> cultures)
{
    _columns.clear();
    for (const std::wstring& culture : cultures)
    {
        static_cast<void>(Add(culture));
    }
}

bool LanguageColumnModel::Add(std::wstring_view cultureName)
{
    if (cultureName.empty() || std::ranges::any_of(_columns, [cultureName](const auto& column) { return EqualIgnoreCase(column.cultureName, cultureName); }))
    {
        return false;
    }
    _columns.push_back(LanguageColumn{.cultureName = std::wstring(cultureName)});
    return true;
}

bool LanguageColumnModel::Remove(std::wstring_view cultureName)
{
    const auto it = std::ranges::find_if(_columns, [cultureName](const auto& column) { return EqualIgnoreCase(column.cultureName, cultureName); });
    if (it == _columns.end())
        return false;
    _columns.erase(it);
    return true;
}

bool LanguageColumnModel::Move(std::wstring_view cultureName, size_t newIndex)
{
    const auto it = std::ranges::find_if(_columns, [cultureName](const auto& column) { return EqualIgnoreCase(column.cultureName, cultureName); });
    if (it == _columns.end())
        return false;
    LanguageColumn column = std::move(*it);
    _columns.erase(it);
    _columns.insert(_columns.begin() + static_cast<std::ptrdiff_t>(std::min(newIndex, _columns.size())), std::move(column));
    NormalizePinnedOrder();
    return true;
}

bool LanguageColumnModel::SetPinned(std::wstring_view cultureName, bool pinned)
{
    const auto it = std::ranges::find_if(_columns, [cultureName](const auto& column) { return EqualIgnoreCase(column.cultureName, cultureName); });
    if (it == _columns.end() || it->pinned == pinned)
        return false;
    it->pinned = pinned;
    NormalizePinnedOrder();
    return true;
}

std::span<const LanguageColumn> LanguageColumnModel::Get() const noexcept
{
    return _columns;
}

std::vector<std::wstring> LanguageColumnModel::GetOrderedCultures() const
{
    std::vector<std::wstring> result;
    result.reserve(_columns.size());
    for (const LanguageColumn& column : _columns)
        result.push_back(column.cultureName);
    return result;
}

void LanguageColumnModel::NormalizePinnedOrder()
{
    std::stable_partition(_columns.begin(), _columns.end(), [](const LanguageColumn& column) noexcept { return column.pinned; });
}

size_t CountDirtyLocalizationCells(std::span<const LocalizationReviewRow> rows) noexcept
{
    size_t count = 0u;
    for (const auto& row : rows)
    {
        count += static_cast<size_t>(std::ranges::count_if(row.targets, [](const auto& cell) noexcept { return cell.dirty; }));
    }
    return count;
}

ValidationSummary ValidateWorkspace(const RedConfigureSession& session,
                                    const std::filesystem::path& localizationOutputRoot,
                                    const std::filesystem::path& themeOutputPath)
{
    ValidationSummary summary;
    const auto add = [&summary](ValidationIssue issue)
    {
        if (issue.severity == ValidationSeverity::Error)
            ++summary.errorCount;
        else
            ++summary.warningCount;
        summary.issues.push_back(std::move(issue));
    };

    for (const std::wstring& error : session.GetWorkspace().errors)
    {
        add({.severity  = ValidationSeverity::Error,
             .category  = ValidationCategory::Workspace,
             .code      = ValidationCode::WorkspaceProcessingError,
             .arguments = {error}});
    }
    for (const std::wstring& error : session.GetThemeCatalog().errors)
    {
        add({.severity = ValidationSeverity::Error, .category = ValidationCategory::Theme, .code = ValidationCode::ThemeCatalogError, .arguments = {error}});
    }
    for (const auto& row : session.GetLocalizationReviewRows())
    {
        for (const auto& cell : row.targets)
        {
            if (cell.validation.status != Localization::PlaceholderStatus::Ok)
            {
                add({.severity    = ValidationSeverity::Error,
                     .category    = ValidationCategory::Localization,
                     .code        = ValidationCode::PlaceholderMismatch,
                     .ownerName   = row.ownerName,
                     .resourceId  = row.id,
                     .cultureName = cell.cultureName});
            }
            else if (! cell.hasExistingTranslation && ! cell.dirty)
            {
                add({.severity    = ValidationSeverity::Warning,
                     .category    = ValidationCategory::Localization,
                     .code        = ValidationCode::MissingTranslation,
                     .ownerName   = row.ownerName,
                     .resourceId  = row.id,
                     .cultureName = cell.cultureName});
            }
        }
    }

    const auto& theme = session.GetThemePreviewModel().GetTheme();
    if (theme.id.empty() || theme.id.find_first_of(L" \\:\t\r\n") != std::wstring::npos)
    {
        add({.severity = ValidationSeverity::Error, .category = ValidationCategory::Theme, .code = ValidationCode::InvalidThemeId});
    }
    if (! session.GetThemePreviewModel().GetLastError().empty())
    {
        add({.severity  = ValidationSeverity::Error,
             .category  = ValidationCategory::Theme,
             .code      = ValidationCode::ThemeResolutionError,
             .arguments = {std::wstring(session.GetThemePreviewModel().GetLastError())}});
    }
    if (localizationOutputRoot.empty() || themeOutputPath.empty())
    {
        add({.severity = ValidationSeverity::Error, .category = ValidationCategory::Export, .code = ValidationCode::EmptyOutputPath});
    }
    else if (localizationOutputRoot == themeOutputPath)
    {
        add({.severity = ValidationSeverity::Error, .category = ValidationCategory::Export, .code = ValidationCode::OutputConflict});
    }

    std::vector<LocalizationExportPreview> previews;
    if (FAILED(session.BuildLocalizationReviewExportPreviews(previews)))
    {
        add({.severity = ValidationSeverity::Error, .category = ValidationCategory::Export, .code = ValidationCode::LocalizationPreviewBuildFailed});
    }
    else
    {
        std::set<std::wstring> outputPaths;
        for (const LocalizationExportPreview& preview : previews)
        {
            std::wstring normalized = (localizationOutputRoot / preview.path.filename()).lexically_normal().wstring();
            std::ranges::transform(normalized, normalized.begin(), [](wchar_t ch) noexcept { return static_cast<wchar_t>(std::towlower(ch)); });
            if (! outputPaths.insert(normalized).second)
            {
                add({.severity = ValidationSeverity::Error, .category = ValidationCategory::Export, .code = ValidationCode::DuplicateLocalizationOutputPath});
            }
            if (EqualIgnoreCase(normalized, themeOutputPath.lexically_normal().wstring()))
            {
                add({.severity = ValidationSeverity::Error, .category = ValidationCategory::Export, .code = ValidationCode::LocalizationThemeOutputConflict});
            }
        }
    }

    for (const std::wstring& culture : session.GetLocalizationReviewCultures())
    {
        for (const DuplicateAccelerator& duplicate : FindDuplicateSiblingAccelerators(session.GetLocalizationReviewRows(), culture))
        {
            add({.severity    = ValidationSeverity::Warning,
                 .category    = ValidationCategory::Accelerator,
                 .code        = ValidationCode::DuplicateAccelerator,
                 .arguments   = {std::wstring(1u, duplicate.accelerator)},
                 .ownerName   = duplicate.ownerName,
                 .cultureName = std::wstring(culture)});
        }
    }
    return summary;
}

LocalizationBatchPreview PreviewLocalizationBatch(std::span<const LocalizationReviewRow> rows, const LocalizationBatchRequest& request)
{
    LocalizationBatchPreview preview{.request = request};
    if (request.targetCulture.empty() || (request.kind == LocalizationBatchKind::CopyCulture && request.sourceCulture.empty()) ||
        (request.kind == LocalizationBatchKind::FindReplace && request.findText.empty()))
    {
        preview.result = BatchApprovalResult::Invalid;
        return preview;
    }
    std::vector<size_t> indices = request.rowIndices;
    if (indices.empty())
    {
        indices.resize(rows.size());
        for (size_t index = 0u; index < rows.size(); ++index)
            indices[index] = index;
    }
    for (const size_t rowIndex : indices)
    {
        if (rowIndex >= rows.size())
        {
            preview.changes.clear();
            preview.result = BatchApprovalResult::Invalid;
            return preview;
        }
        const LocalizationReviewRow& row     = rows[rowIndex];
        const LocalizationTargetCell* target = FindCell(row, request.targetCulture);
        if (! target)
            continue;
        std::wstring after = target->targetText;
        switch (request.kind)
        {
            case LocalizationBatchKind::CopyEnglish: after = row.sourceText; break;
            case LocalizationBatchKind::CopyCulture:
                if (const LocalizationTargetCell* source = FindCell(row, request.sourceCulture))
                    after = source->targetText;
                break;
            case LocalizationBatchKind::Clear: after.clear(); break;
            case LocalizationBatchKind::FindReplace: after = Replaced(after, request.findText, request.replaceText); break;
            case LocalizationBatchKind::NormalizePlaceholderWhitespace: after = NormalizePlaceholderWhitespace(after); break;
            case LocalizationBatchKind::PreserveAccelerators: after = PreserveSourceAccelerator(row.sourceText, after); break;
            case LocalizationBatchKind::MarkReviewed: break;
            default: break;
        }
        const bool afterReviewed = request.kind == LocalizationBatchKind::MarkReviewed || target->reviewed;
        if (after != target->targetText || afterReviewed != target->reviewed)
        {
            preview.changes.push_back({.ownerName      = row.ownerName,
                                       .resourceId     = row.id,
                                       .cultureName    = request.targetCulture,
                                       .before         = target->targetText,
                                       .after          = std::move(after),
                                       .beforeReviewed = target->reviewed,
                                       .afterReviewed  = afterReviewed});
        }
    }
    preview.result = preview.changes.empty() ? BatchApprovalResult::NoChanges : BatchApprovalResult::Ready;
    return preview;
}

std::vector<DuplicateAccelerator> FindDuplicateSiblingAccelerators(std::span<const LocalizationReviewRow> rows, std::wstring_view cultureName)
{
    std::map<std::pair<std::wstring, wchar_t>, std::vector<std::wstring>> grouped;
    for (const auto& row : rows)
    {
        if (row.id.find(L"MENU") == std::wstring::npos && row.id.find(L"CMD") == std::wstring::npos)
            continue;
        const LocalizationTargetCell* cell = FindCell(row, cultureName);
        const std::wstring_view text       = cell && ! cell->targetText.empty() ? std::wstring_view(cell->targetText) : std::wstring_view(row.sourceText);
        if (const std::optional<wchar_t> accelerator = Accelerator(text))
        {
            grouped[{row.ownerName, accelerator.value()}].push_back(row.id);
        }
    }
    std::vector<DuplicateAccelerator> result;
    for (auto& [key, ids] : grouped)
    {
        if (ids.size() > 1u)
        {
            result.push_back({.ownerName = key.first, .accelerator = key.second, .resourceIds = std::move(ids)});
        }
    }
    return result;
}

ClipboardMatrix ParseClipboardMatrix(std::wstring_view text)
{
    ClipboardMatrix matrix;
    if (text.empty())
    {
        return matrix;
    }
    size_t lineStart = 0u;
    while (lineStart < text.size())
    {
        size_t lineEnd = text.find_first_of(L"\r\n", lineStart);
        if (lineEnd == std::wstring_view::npos)
            lineEnd = text.size();
        std::vector<std::wstring> row;
        size_t cellStart = lineStart;
        while (cellStart <= lineEnd)
        {
            size_t cellEnd = text.find(L'\t', cellStart);
            if (cellEnd == std::wstring_view::npos || cellEnd > lineEnd)
                cellEnd = lineEnd;
            row.emplace_back(text.substr(cellStart, cellEnd - cellStart));
            if (cellEnd == lineEnd)
                break;
            cellStart = cellEnd + 1u;
        }
        matrix.rows.push_back(std::move(row));
        if (lineEnd == text.size())
            break;
        lineStart = lineEnd + 1u;
        if (lineStart < text.size() && text[lineEnd] == L'\r' && text[lineStart] == L'\n')
            ++lineStart;
    }
    return matrix;
}

std::wstring SerializeClipboardMatrix(const ClipboardMatrix& matrix)
{
    std::wstring result;
    for (size_t rowIndex = 0u; rowIndex < matrix.rows.size(); ++rowIndex)
    {
        if (rowIndex > 0u)
            result += L"\r\n";
        for (size_t columnIndex = 0u; columnIndex < matrix.rows[rowIndex].size(); ++columnIndex)
        {
            if (columnIndex > 0u)
                result.push_back(L'\t');
            result += matrix.rows[rowIndex][columnIndex];
        }
    }
    return result;
}

ThemeMassPreview PreviewThemeMassChange(const Themes::ThemePreviewModel& model, const ThemeMassRequest& request)
{
    ThemeMassPreview preview{.request = request, .beforeTheme = model.GetTheme()};
    if (request.recipe == ThemeRecipe::SetAlpha && request.alphaPercent > 100u)
    {
        preview.result     = BatchApprovalResult::Invalid;
        preview.diagnostic = ThemeMassDiagnostic::InvalidAlpha;
        return preview;
    }

    std::wstring replaceFrom;
    std::wstring replaceTo;
    if (request.recipe == ThemeRecipe::ReplaceReference)
    {
        const size_t separator = request.argument.find(L'=');
        if (separator == std::wstring::npos || separator == 0u || separator + 1u >= request.argument.size() ||
            request.argument.find(L'=', separator + 1u) != std::wstring::npos)
        {
            preview.result     = BatchApprovalResult::Invalid;
            preview.diagnostic = ThemeMassDiagnostic::MalformedReferenceReplacement;
            return preview;
        }
        replaceFrom = request.argument.substr(0u, separator);
        replaceTo   = request.argument.substr(separator + 1u);
    }
    if (request.recipe == ThemeRecipe::ConvertSolidsToReferences &&
        (! Common::Settings::IsValidThemePaletteName(request.argument) || ! model.GetTheme().palette.contains(request.argument)))
    {
        preview.result     = BatchApprovalResult::Invalid;
        preview.diagnostic = ThemeMassDiagnostic::MissingPaletteTarget;
        return preview;
    }

    for (const std::wstring& key : request.keys)
    {
        const std::wstring before = model.GetAuthoredColorText(key);
        std::optional<Common::Settings::ThemeColorSource> beforeSource;
        if (! before.empty())
        {
            Common::Settings::ThemeColorSource parsed;
            std::wstring parseMessage;
            if (FAILED(Common::Settings::ParseThemeColorSource(before, parsed, &parseMessage)))
            {
                preview.result     = BatchApprovalResult::Invalid;
                preview.diagnostic = ThemeMassDiagnostic::InvalidExistingSource;
                preview.changes.clear();
                return preview;
            }
            beforeSource = std::move(parsed);
        }

        std::wstring sourcePaletteName;
        std::wstring sourceValue;
        std::wstring operand = before;
        if (RecipeNeedsCapturedSource(request.recipe))
        {
            sourcePaletteName = RecipePaletteName(key);
            for (uint32_t suffix = 2u; model.GetTheme().palette.contains(sourcePaletteName); ++suffix)
            {
                sourcePaletteName = std::format(L"{}_{}", RecipePaletteName(key), suffix);
            }
            sourceValue = before;
            if (sourceValue.empty())
            {
                if (const std::optional<uint32_t> effective = model.GetEffectiveColor(key))
                {
                    sourceValue = Common::Settings::FormatColor(effective.value());
                }
            }
            operand = std::wstring(L"palette.") + sourcePaletteName;
        }
        std::wstring after;
        std::optional<Common::Settings::ThemeColorSource> afterSource;
        if (request.recipe == ThemeRecipe::ReplaceReference)
        {
            if (! beforeSource.has_value())
                continue;
            Common::Settings::ThemeColorSource rewritten = beforeSource.value();
            bool changed                                 = false;
            for (std::wstring& reference : rewritten.references)
            {
                if (reference == replaceFrom)
                {
                    reference = replaceTo;
                    changed   = true;
                }
            }
            if (! changed)
                continue;
            after       = Common::Settings::FormatThemeColorSource(rewritten);
            afterSource = std::move(rewritten);
        }
        else if (request.recipe == ThemeRecipe::ConvertSolidsToReferences)
        {
            if (! beforeSource.has_value() || beforeSource->kind != Common::Settings::ThemeColorSourceKind::Direct)
                continue;
            Common::Settings::ThemeColorSource reference;
            reference.kind = Common::Settings::ThemeColorSourceKind::Reference;
            reference.references.push_back(std::wstring(L"palette.") + request.argument);
            after       = Common::Settings::FormatThemeColorSource(reference);
            afterSource = std::move(reference);
        }
        else if (request.recipe == ThemeRecipe::RemoveOverrides)
        {
            if (! beforeSource.has_value())
                continue;
        }
        else
        {
            after = RecipeValue(request.recipe, key, operand, request.argument, request.alphaPercent);
            Common::Settings::ThemeColorSource parsed;
            std::wstring parseMessage;
            if (FAILED(Common::Settings::ParseThemeColorSource(after, parsed, &parseMessage)))
            {
                preview.result     = BatchApprovalResult::Invalid;
                preview.diagnostic = ThemeMassDiagnostic::InvalidGeneratedSource;
                preview.changes.clear();
                return preview;
            }
            afterSource = std::move(parsed);
        }
        if (before != after)
        {
            std::optional<Common::Settings::ThemeColorSource> sourceValueSource;
            if (! sourcePaletteName.empty())
            {
                Common::Settings::ThemeColorSource parsed;
                std::wstring parseMessage;
                if (sourceValue.empty() || FAILED(Common::Settings::ParseThemeColorSource(sourceValue, parsed, &parseMessage)))
                {
                    preview.result     = BatchApprovalResult::Invalid;
                    preview.diagnostic = ThemeMassDiagnostic::InvalidGeneratedSource;
                    preview.changes.clear();
                    return preview;
                }
                sourceValueSource = std::move(parsed);
            }
            preview.changes.push_back({.key               = key,
                                       .before            = before,
                                       .after             = after,
                                       .beforeSource      = std::move(beforeSource),
                                       .afterSource       = std::move(afterSource),
                                       .sourcePaletteName = std::move(sourcePaletteName),
                                       .sourceValue       = std::move(sourceValue),
                                       .sourceValueSource = std::move(sourceValueSource)});
        }
    }

    if (preview.changes.empty())
    {
        preview.result = BatchApprovalResult::NoChanges;
        return preview;
    }
    Common::Settings::ThemeDefinition candidate;
    if (! BuildCandidateTheme(preview, candidate, preview.diagnostic))
    {
        preview.result = BatchApprovalResult::Invalid;
        return preview;
    }
    Themes::ThemePreviewModel candidateModel;
    candidateModel.SetTheme(candidate);
    if (! candidateModel.GetLastError().empty())
    {
        preview.result     = BatchApprovalResult::Invalid;
        preview.diagnostic = ThemeMassDiagnostic::InvalidCandidate;
        return preview;
    }
    preview.result = BatchApprovalResult::Ready;
    return preview;
}

BatchApprovalResult ApplyThemeMassChange(Themes::ThemePreviewModel& model, const ThemeMassPreview& preview)
{
    if (preview.result != BatchApprovalResult::Ready)
    {
        return preview.result == BatchApprovalResult::NoChanges ? BatchApprovalResult::NoChanges : BatchApprovalResult::Invalid;
    }
    if (! ThemeDefinitionsEqual(model.GetTheme(), preview.beforeTheme))
    {
        return BatchApprovalResult::Stale;
    }
    Common::Settings::ThemeDefinition candidate;
    ThemeMassDiagnostic diagnostic = ThemeMassDiagnostic::None;
    if (! BuildCandidateTheme(preview, candidate, diagnostic))
    {
        return BatchApprovalResult::Invalid;
    }
    Themes::ThemePreviewModel candidateModel;
    candidateModel.SetTheme(candidate);
    if (! candidateModel.GetLastError().empty())
    {
        return BatchApprovalResult::Invalid;
    }
    model.SetTheme(candidate);
    return BatchApprovalResult::Applied;
}

std::optional<DuplicateThemeCandidate> BuildDuplicateThemeCandidate(std::wstring_view sourceId,
                                                                    std::wstring_view sourceName,
                                                                    std::wstring_view localizedCopyLabel,
                                                                    uint32_t sequence)
{
    constexpr size_t maximumLength = 64u;
    if (sequence == 0u || sourceName.empty() || localizedCopyLabel.empty())
    {
        return std::nullopt;
    }

    std::wstring idStem(sourceId);
    if (sourceId.starts_with(L"builtin/"))
    {
        idStem = std::wstring(L"user/") + std::wstring(sourceId.substr(8u));
    }
    const std::wstring idSuffix   = sequence == 1u ? L"-copy" : std::format(L"-copy-{}", sequence);
    const std::wstring nameSuffix = sequence == 1u ? std::format(L" {}", localizedCopyLabel) : std::format(L" {} {}", localizedCopyLabel, sequence);
    if (idSuffix.size() >= maximumLength || nameSuffix.size() >= maximumLength)
    {
        return std::nullopt;
    }
    idStem.resize(std::min(idStem.size(), maximumLength - idSuffix.size()));
    std::wstring nameStem(sourceName.substr(0u, std::min(sourceName.size(), maximumLength - nameSuffix.size())));
    DuplicateThemeCandidate candidate{.id = idStem + idSuffix, .name = nameStem + nameSuffix};
    if (! Common::Settings::IsValidUserThemeId(candidate.id) || candidate.name.empty() || candidate.name.size() > maximumLength)
    {
        return std::nullopt;
    }
    return candidate;
}

ThemeTokenMetadata BuildThemeTokenMetadata(const Themes::ThemePreviewModel& model, std::wstring_view key)
{
    ThemeTokenMetadata metadata;
    metadata.key        = key;
    const size_t dot    = key.find(L'.');
    metadata.group      = std::wstring(dot == std::wstring_view::npos ? key : key.substr(0u, dot));
    const auto kind     = model.GetSourceKind(key);
    metadata.sourceKind = ! kind.has_value()                                                  ? ThemeTokenSourceKind::Inherited
                          : kind.value() == Common::Settings::ThemeColorSourceKind::Direct    ? ThemeTokenSourceKind::Literal
                          : kind.value() == Common::Settings::ThemeColorSourceKind::Reference ? ThemeTokenSourceKind::Reference
                                                                                              : ThemeTokenSourceKind::Function;
    metadata.usageCount = model.GetAffected(key).size() + 1u;
    if (const std::optional<std::wstring_view> peer = ContrastPeer(key))
    {
        const std::optional<uint32_t> lhs = model.GetEffectiveColor(key);
        const std::optional<uint32_t> rhs = model.GetEffectiveColor(peer.value());
        if (lhs.has_value() && rhs.has_value())
        {
            const double lhsLuminance = RelativeLuminanceFromArgb(lhs.value());
            const double rhsLuminance = RelativeLuminanceFromArgb(rhs.value());
            metadata.contrastRatio    = ContrastRatioFromRelativeLuminance(lhsLuminance, rhsLuminance);
            metadata.contrastKnown    = true;
            metadata.contrastPass     = metadata.contrastRatio >= 4.5;
        }
    }
    return metadata;
}
} // namespace RedConfigure::Workflow
