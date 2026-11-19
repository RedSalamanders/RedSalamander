#include "RedConfigurePagePresenters.h"

#include "resource.h"

namespace RedConfigure::Ui
{
StartPageSummary StartPagePresenter::Build(const RedConfigureSession& session) noexcept
{
    return {.resourceOwnerCount = session.GetWorkspace().resourceOwners.size(),
            .themeFileCount     = session.GetThemeCatalog().themes.size(),
            .scanErrorCount     = session.GetWorkspace().errors.size()};
}

BatchInteraction LocalizationPagePresenter::Execute(RedConfigureSession& session, const Workflow::LocalizationBatchRequest& request)
{
    if (_pending.has_value() && _pending->request == request)
    {
        const Workflow::BatchApprovalResult result = session.ApplyLocalizationBatch(_pending.value());
        _pending.reset();
        return {.phase = BatchInteractionPhase::Apply, .result = result};
    }

    _pending = Workflow::PreviewLocalizationBatch(session.GetLocalizationReviewRows(), request);
    BatchInteraction interaction{.phase = BatchInteractionPhase::Preview, .result = _pending->result, .changeCount = _pending->changes.size()};
    if (! _pending->changes.empty())
    {
        const Workflow::LocalizationBatchChange& first = _pending->changes.front();
        interaction.firstChange                        = BatchChangeSummary{.identity = first.resourceId, .before = first.before, .after = first.after};
    }
    if (_pending->result != Workflow::BatchApprovalResult::Ready)
    {
        _pending.reset();
    }
    return interaction;
}

void LocalizationPagePresenter::Invalidate() noexcept
{
    _pending.reset();
}

BatchInteraction ThemesPagePresenter::Execute(RedConfigureSession& session, const Workflow::ThemeMassRequest& request)
{
    if (_pending.has_value() && _pending->request == request)
    {
        const Workflow::BatchApprovalResult result = session.ApplyThemeMassChange(_pending.value());
        _pending.reset();
        return {.phase = BatchInteractionPhase::Apply, .result = result};
    }

    _pending = Workflow::PreviewThemeMassChange(session.GetThemePreviewModel(), request);
    BatchInteraction interaction{.phase = BatchInteractionPhase::Preview, .result = _pending->result, .changeCount = _pending->changes.size()};
    if (! _pending->changes.empty())
    {
        const Workflow::ThemeMassChange& first = _pending->changes.front();
        interaction.firstChange                = BatchChangeSummary{.identity = first.key, .before = first.before, .after = first.after};
    }
    if (_pending->result != Workflow::BatchApprovalResult::Ready)
    {
        _pending.reset();
    }
    return interaction;
}

void ThemesPagePresenter::Invalidate() noexcept
{
    _pending.reset();
}

uint32_t ThemesPagePresenter::GetOriginResourceId(Themes::ThemeCatalogOrigin origin) noexcept
{
    switch (origin)
    {
        case Themes::ThemeCatalogOrigin::BuiltIn: return IDS_REDCONFIGURE_THEME_ORIGIN_BUILTIN;
        case Themes::ThemeCatalogOrigin::User: return IDS_REDCONFIGURE_THEME_ORIGIN_USER;
        case Themes::ThemeCatalogOrigin::File: return IDS_REDCONFIGURE_THEME_ORIGIN_FILE;
        default: return IDS_REDCONFIGURE_THEME_ORIGIN_FILE;
    }
}

uint32_t ReviewExportPagePresenter::GetCategoryResourceId(Workflow::ValidationCategory category) noexcept
{
    switch (category)
    {
        case Workflow::ValidationCategory::Workspace: return IDS_REDCONFIGURE_CATEGORY_WORKSPACE;
        case Workflow::ValidationCategory::Theme: return IDS_REDCONFIGURE_CATEGORY_THEME;
        case Workflow::ValidationCategory::Localization: return IDS_REDCONFIGURE_CATEGORY_LOCALIZATION;
        case Workflow::ValidationCategory::Export: return IDS_REDCONFIGURE_CATEGORY_EXPORT;
        case Workflow::ValidationCategory::Accelerator: return IDS_REDCONFIGURE_CATEGORY_ACCELERATOR;
        default: return IDS_REDCONFIGURE_CATEGORY_WORKSPACE;
    }
}

uint32_t ReviewExportPagePresenter::GetMessageResourceId(Workflow::ValidationCode code) noexcept
{
    switch (code)
    {
        case Workflow::ValidationCode::WorkspaceProcessingError: return IDS_REDCONFIGURE_VALIDATION_WORKSPACE_ERROR;
        case Workflow::ValidationCode::ThemeCatalogError: return IDS_REDCONFIGURE_VALIDATION_THEME_CATALOG_ERROR;
        case Workflow::ValidationCode::PlaceholderMismatch: return IDS_REDCONFIGURE_VALIDATION_PLACEHOLDER_MISMATCH;
        case Workflow::ValidationCode::MissingTranslation: return IDS_REDCONFIGURE_VALIDATION_MISSING_TRANSLATION;
        case Workflow::ValidationCode::InvalidThemeId: return IDS_REDCONFIGURE_VALIDATION_INVALID_THEME_ID;
        case Workflow::ValidationCode::ThemeResolutionError: return IDS_REDCONFIGURE_VALIDATION_THEME_RESOLUTION_ERROR;
        case Workflow::ValidationCode::EmptyOutputPath: return IDS_REDCONFIGURE_VALIDATION_EMPTY_OUTPUT_PATH;
        case Workflow::ValidationCode::OutputConflict: return IDS_REDCONFIGURE_VALIDATION_OUTPUT_CONFLICT;
        case Workflow::ValidationCode::LocalizationPreviewBuildFailed: return IDS_REDCONFIGURE_VALIDATION_PREVIEW_FAILED;
        case Workflow::ValidationCode::DuplicateLocalizationOutputPath: return IDS_REDCONFIGURE_VALIDATION_DUPLICATE_OUTPUT;
        case Workflow::ValidationCode::LocalizationThemeOutputConflict: return IDS_REDCONFIGURE_VALIDATION_LOCALIZATION_THEME_CONFLICT;
        case Workflow::ValidationCode::DuplicateAccelerator: return IDS_REDCONFIGURE_FMT_VALIDATION_DUPLICATE_ACCELERATOR;
        default: return IDS_REDCONFIGURE_VALIDATION_WORKSPACE_ERROR;
    }
}
} // namespace RedConfigure::Ui
