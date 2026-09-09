#pragma once

#include "RedConfigureSession.h"
#include "RedConfigureWorkflow.h"

#include <cstdint>
#include <optional>
#include <string>

namespace RedConfigure::Ui
{
struct StartPageSummary
{
    size_t resourceOwnerCount = 0u;
    size_t themeFileCount     = 0u;
    size_t scanErrorCount     = 0u;
};

class StartPagePresenter final
{
public:
    [[nodiscard]] static StartPageSummary Build(const RedConfigureSession& session) noexcept;
};

enum class BatchInteractionPhase : uint8_t
{
    Preview,
    Apply,
};

struct BatchChangeSummary
{
    std::wstring identity;
    std::wstring before;
    std::wstring after;
};

struct BatchInteraction
{
    BatchInteractionPhase phase          = BatchInteractionPhase::Preview;
    Workflow::BatchApprovalResult result = Workflow::BatchApprovalResult::NoChanges;
    size_t changeCount                   = 0u;
    std::optional<BatchChangeSummary> firstChange;
};

class LocalizationPagePresenter final
{
public:
    [[nodiscard]] BatchInteraction Execute(RedConfigureSession& session, const Workflow::LocalizationBatchRequest& request);
    void Invalidate() noexcept;

private:
    std::optional<Workflow::LocalizationBatchPreview> _pending;
};

class ThemesPagePresenter final
{
public:
    [[nodiscard]] BatchInteraction Execute(RedConfigureSession& session, const Workflow::ThemeMassRequest& request);
    void Invalidate() noexcept;

    [[nodiscard]] static uint32_t GetOriginResourceId(Themes::ThemeCatalogOrigin origin) noexcept;

private:
    std::optional<Workflow::ThemeMassPreview> _pending;
};

class ReviewExportPagePresenter final
{
public:
    [[nodiscard]] static uint32_t GetCategoryResourceId(Workflow::ValidationCategory category) noexcept;
    [[nodiscard]] static uint32_t GetMessageResourceId(Workflow::ValidationCode code) noexcept;
};
} // namespace RedConfigure::Ui
