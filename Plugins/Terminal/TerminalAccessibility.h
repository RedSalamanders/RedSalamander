#pragma once

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <memory>
#include <string>

class TerminalAccessibility final
{
public:
    TerminalAccessibility();
    ~TerminalAccessibility();

    TerminalAccessibility(const TerminalAccessibility&) = delete;
    TerminalAccessibility(TerminalAccessibility&&) = delete;
    TerminalAccessibility& operator=(const TerminalAccessibility&) = delete;
    TerminalAccessibility& operator=(TerminalAccessibility&&) = delete;

    [[nodiscard]] HRESULT Initialize(HWND window) noexcept;
    void Publish(std::wstring text, bool focused) noexcept;
    [[nodiscard]] LRESULT HandleGetObject(WPARAM wParam, LPARAM lParam) noexcept;
    void Retire() noexcept;

#if defined(ENABLE_TESTS)
    [[nodiscard]] HRESULT DebugGetName(std::wstring& name) noexcept;
#endif

private:
    struct Impl;
    std::unique_ptr<Impl> _impl;
};

[[nodiscard]] bool CanUnloadTerminalAccessibilityProviders() noexcept;
