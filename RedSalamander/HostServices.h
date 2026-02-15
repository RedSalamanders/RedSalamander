#pragma once

#include "PlugInterfaces/Host.h"

// Returns a process-lifetime host services object that plugins can use via QueryInterface.
// The returned pointer is always non-null.
IHost* GetHostServices() noexcept;

// Internal convenience helpers for in-tree call sites.
HRESULT HostShowAlert(const HostAlertRequest& request, void* cookie = nullptr) noexcept;
HRESULT HostClearAlert(HostAlertScope scope, void* cookie = nullptr) noexcept;
HRESULT HostShowPrompt(const HostPromptRequest& request, void* cookie, HostPromptResult* result) noexcept;

// Debug/testing hook: bypass prompts and accept the default result.
// Intended for automated self-tests that must not block on modal dialogs.
void HostSetAutoAcceptPrompts(bool enabled) noexcept;
bool HostGetAutoAcceptPrompts() noexcept;

// FolderWindow dispatch helper for cross-thread plugin calls.
// Returns true if the message was handled (and `result` is set).
bool TryHandleHostServicesWindowMessage(UINT message, WPARAM wParam, LPARAM lParam, LRESULT& result) noexcept;
