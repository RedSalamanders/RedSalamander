#pragma once

#include <atomic>
#include <cstddef>
#include <cstdint>
#include <initializer_list>
#include <string>

#include "PlugInterfaces/Host.h"
#include "SettingsStore.h"

class FolderWindow;

// A synchronous provider call can report an actual credential UI wait to its owning task.
// Capture this observer with the request before crossing to the UI thread. Its context
// remains borrowed until that synchronous request returns.
struct HostCredentialWaitObserver
{
    void* context                                        = nullptr;
    void (*notify)(void* context, bool waiting) noexcept = nullptr;
};

class ScopedHostCredentialWaitObserver final
{
public:
    explicit ScopedHostCredentialWaitObserver(HostCredentialWaitObserver observer) noexcept : _previous(_current)
    {
        _current = observer;
    }

    ScopedHostCredentialWaitObserver(const ScopedHostCredentialWaitObserver&)            = delete;
    ScopedHostCredentialWaitObserver& operator=(const ScopedHostCredentialWaitObserver&) = delete;
    ~ScopedHostCredentialWaitObserver() noexcept
    {
        _current = _previous;
    }

    [[nodiscard]] static HostCredentialWaitObserver Capture() noexcept
    {
        return _current;
    }

private:
    // Provider ABI calls do not carry this host-only observation. Nested work on a
    // scheduler thread restores its calling task's observer when the call unwinds.
    inline static thread_local HostCredentialWaitObserver _current{};
    HostCredentialWaitObserver _previous;
};

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
// A blocking host prompt's nested message loop is running on the UI thread.
[[nodiscard]] bool HostIsBlockingPromptActive() noexcept;

// Debug/testing hook: bypass prompts and accept the default result.
// Intended for automated self-tests that must not block on modal dialogs.
void HostSetAutoAcceptPrompts(bool enabled) noexcept;
bool HostGetAutoAcceptPrompts() noexcept;

#ifdef ENABLE_TESTS
struct HostPromptDebugSnapshot
{
    HostAlertScope scope                = HOST_ALERT_SCOPE_APPLICATION;
    HWND targetWindow                   = nullptr;
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
// Answers consumed one per prompt, ahead of the sticky override above (multi-step prompt flows).
void HostQueueTestPromptResults(std::initializer_list<HostPromptResult> results);
void HostClearTestPromptResultQueue() noexcept;
uint64_t HostGetTestPromptRequestCount() noexcept;
bool HostGetTestLastPromptDebugSnapshot(HostPromptDebugSnapshot& out) noexcept;
size_t HostGetTestSynchronousPayloadRegistrationCount() noexcept;
size_t HostGetTestWindowAlertOverlayCount() noexcept;
LPARAM HostGetTestLastRetiredSynchronousPayloadToken() noexcept;
#endif

// FolderWindow dispatch helper for cross-thread plugin calls.
// Returns true if the message was handled (and `result` is set).
bool TryHandleHostServicesWindowMessage(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam, LRESULT& result) noexcept;
