#pragma once

#include <atomic>
#include <cstddef>
#include <cstdint>
#include <string>

#include "PlugInterfaces/Host.h"
#include "SettingsStore.h"

class FolderWindow;

// Binds the process-lifetime plugin host to the composition root before any
// plugin can call it. Dependencies remain non-owning and outlive the host.
void ConfigureHostServices(FolderWindow& folderWindow, std::atomic<HWND>& folderWindowHwnd, Common::Settings::Settings& settings) noexcept;

// Returns a process-lifetime host services object that plugins can use via QueryInterface.
// The returned pointer is always non-null.
IHost* GetHostServices() noexcept;

// Internal convenience helpers for in-tree call sites.
HRESULT HostShowAlert(const HostAlertRequest& request, void* cookie = nullptr) noexcept;
HRESULT HostClearAlert(HostAlertScope scope, void* cookie = nullptr) noexcept;
HRESULT HostShowPrompt(const HostPromptRequest& request, void* cookie, HostPromptResult* result) noexcept;
void HostClearConnectionSessionState() noexcept;
// Fences new prompts and makes an active nested prompt return before queued owner teardown.
void HostBeginPromptShutdown() noexcept;
void HostResetPromptShutdown() noexcept;
[[nodiscard]] bool HostIsPromptShutdown() noexcept;

// Debug/testing hook: bypass prompts and accept the default result.
// Intended for automated self-tests that must not block on modal dialogs.
void HostSetAutoAcceptPrompts(bool enabled) noexcept;
bool HostGetAutoAcceptPrompts() noexcept;

#ifdef ENABLE_TESTS
struct HostPromptDebugSnapshot
{
    HostAlertSeverity severity          = HOST_ALERT_ERROR;
    HostPromptButtons buttons           = HOST_PROMPT_BUTTONS_OK;
    HostPromptResult defaultResult      = HOST_PROMPT_RESULT_NONE;
    HostPromptPresentation presentation = HOST_PROMPT_PRESENTATION_DEFAULT;
    std::wstring title;
    std::wstring message;
    bool hasFileOperationOptions = false;
    HostFileOperationPromptOptions fileOperationOptions{};
};

// Debug/testing hook: force the next prompts to return a specific supported result.
// Used by self-tests that need to exercise non-default prompt branches.
void HostSetTestPromptResultOverride(HostPromptResult result) noexcept;
HostPromptResult HostGetTestPromptResultOverride() noexcept;
void HostClearTestPromptResultOverride() noexcept;
void HostSetTestPromptFileOperationOptionsOverride(const HostFileOperationPromptOptions& options) noexcept;
void HostClearTestPromptFileOperationOptionsOverride() noexcept;
void HostResetTestPromptRequestCount() noexcept;
uint64_t HostGetTestPromptRequestCount() noexcept;
bool HostGetTestLastPromptDebugSnapshot(HostPromptDebugSnapshot& out) noexcept;
size_t HostGetTestSynchronousPayloadRegistrationCount() noexcept;
LPARAM HostGetTestLastRetiredSynchronousPayloadToken() noexcept;
#endif

// FolderWindow dispatch helper for cross-thread plugin calls.
// Returns true if the message was handled (and `result` is set).
bool TryHandleHostServicesWindowMessage(UINT message, WPARAM wParam, LPARAM lParam, LRESULT& result) noexcept;
