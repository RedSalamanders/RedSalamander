#pragma once

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <Windows.h>

#include "DxUi.h"

#include <filesystem>
#include <memory>

namespace RedConfigure
{
class RedConfigureSession;
}

namespace RedConfigure::Ui
{
class RedConfigureRootController
{
public:
    RedConfigureRootController() = default;
    RedConfigureRootController(const RedConfigureRootController&) = delete;
    RedConfigureRootController& operator=(const RedConfigureRootController&) = delete;
    RedConfigureRootController(RedConfigureRootController&&) = delete;
    RedConfigureRootController& operator=(RedConfigureRootController&&) = delete;
    virtual ~RedConfigureRootController() = default;

    virtual void ReloadWorkspaceFromFields() = 0;
};

struct RedConfigureRootCreateResult
{
    RedConfigureRootCreateResult() = default;
    RedConfigureRootCreateResult(const RedConfigureRootCreateResult&) = delete;
    RedConfigureRootCreateResult& operator=(const RedConfigureRootCreateResult&) = delete;
    RedConfigureRootCreateResult(RedConfigureRootCreateResult&&) noexcept = default;
    RedConfigureRootCreateResult& operator=(RedConfigureRootCreateResult&&) noexcept = default;
    ~RedConfigureRootCreateResult() = default;

    std::unique_ptr<RedSalamander::DxUi::Panel> control;
    RedConfigureRootController* controller = nullptr;
};

[[nodiscard]] RedConfigureRootCreateResult CreateRedConfigureRoot(HINSTANCE instance, RedConfigureSession& session, std::filesystem::path initialRoot);
}
