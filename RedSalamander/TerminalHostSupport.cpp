#include "Framework.h"

#include "TerminalHostSupport.h"

#include "PathUtils.h"

#include <algorithm>
#include <objbase.h>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027) // WIL move-only resource wrappers.
#include <wil/resource.h>
#pragma warning(pop)

namespace TerminalHostSupport
{
TerminalUtf16Span TerminalSpan(std::wstring_view text) noexcept
{
    return TerminalUtf16Span{.data = text.data(), .length = static_cast<uint32_t>(text.size())};
}

TerminalLogicalLocation OwnedTerminalLocation::View() const noexcept
{
    TerminalLogicalLocation value{};
    value.sizeBytes                = sizeof(value);
    value.kind                     = kind;
    value.windowsPath              = TerminalSpan(windowsPath);
    value.wslDistribution          = TerminalSpan(wslDistribution);
    value.wslAbsolutePath          = TerminalSpan(wslAbsolutePath);
    value.pluginShortId            = TerminalSpan(pluginShortId);
    value.pluginBackingWindowsPath = TerminalSpan(pluginBackingWindowsPath);
    return value;
}

std::filesystem::path OwnedTerminalLocation::IdentityPath() const
{
    if (! windowsPath.empty())
    {
        return std::filesystem::path(windowsPath);
    }
    if (! pluginBackingWindowsPath.empty())
    {
        return std::filesystem::path(pluginBackingWindowsPath);
    }
    return std::filesystem::path(wslAbsolutePath);
}

OwnedTerminalLocation MakeTerminalLocation(const std::filesystem::path& path)
{
    OwnedTerminalLocation result{};
    const std::wstring text                     = path.wstring();
    constexpr std::wstring_view localhostPrefix = L"\\\\wsl.localhost\\";
    constexpr std::wstring_view legacyPrefix    = L"\\\\wsl$\\";
    const auto startsWithNoCase                 = [](std::wstring_view value, std::wstring_view prefix) noexcept
    {
        return value.size() >= prefix.size() &&
               CompareStringOrdinal(value.data(), static_cast<int>(prefix.size()), prefix.data(), static_cast<int>(prefix.size()), TRUE) == CSTR_EQUAL;
    };
    size_t prefixLength = 0u;
    if (startsWithNoCase(text, localhostPrefix))
    {
        prefixLength = localhostPrefix.size();
    }
    else if (startsWithNoCase(text, legacyPrefix))
    {
        prefixLength = legacyPrefix.size();
    }
    if (prefixLength != 0u)
    {
        const std::wstring_view tail(text.data() + prefixLength, text.size() - prefixLength);
        const size_t separator = tail.find(L'\\');
        result.wslDistribution.assign(tail.substr(0u, separator));
        if (! result.wslDistribution.empty())
        {
            result.wslAbsolutePath = separator == std::wstring_view::npos ? L"/" : std::wstring(tail.substr(separator));
            std::replace(result.wslAbsolutePath.begin(), result.wslAbsolutePath.end(), L'\\', L'/');
            result.kind = TerminalLocationKind::Wsl;
            return result;
        }
    }

    switch (Common::Paths::ClassifyWindowsPath(text))
    {
        case Common::Paths::WindowsPathClass::DriveAbsolute:
        case Common::Paths::WindowsPathClass::ExtendedDriveAbsolute:
            result.kind        = TerminalLocationKind::WindowsLocal;
            result.windowsPath = text;
            break;
        case Common::Paths::WindowsPathClass::Unc:
        case Common::Paths::WindowsPathClass::ExtendedUnc:
            result.kind        = TerminalLocationKind::WindowsUnc;
            result.windowsPath = text;
            break;
        case Common::Paths::WindowsPathClass::Relative:
        case Common::Paths::WindowsPathClass::Rooted:
        case Common::Paths::WindowsPathClass::DriveRelative:
        case Common::Paths::WindowsPathClass::ExtendedOther:
        case Common::Paths::WindowsPathClass::Device: break;
    }
    return result;
}

TerminalTheme BuildTerminalTheme(const AppTheme& appTheme, uint32_t dpi) noexcept
{
    TerminalTheme theme{};
    theme.sizeBytes               = sizeof(theme);
    theme.dpi                     = dpi;
    theme.dark                    = appTheme.dark ? 1u : 0u;
    theme.highContrast            = appTheme.highContrast ? 1u : 0u;
    theme.backgroundArgb          = ColorToArgb(appTheme.folderView.backgroundColor);
    theme.foregroundArgb          = ColorToArgb(appTheme.folderView.textNormal);
    theme.cursorArgb              = theme.foregroundArgb;
    theme.selectionBackgroundArgb = ColorToArgb(appTheme.folderView.itemBackgroundSelected);
    theme.selectionForegroundArgb = ColorToArgb(appTheme.folderView.textSelected);
    theme.hyperlinkArgb           = ColorToArgb(appTheme.accent);
    theme.inactiveStatusArgb      = ColorToArgb(appTheme.folderView.textDisabled);
    return theme;
}

#if defined(ENABLE_TESTS)
HRESULT DebugGetScreenText(ITerminal* terminal, std::wstring& text) noexcept
{
    text.clear();
    if (terminal == nullptr)
    {
        return E_INVALIDARG;
    }
    const HMODULE module = GetModuleHandleW(L"Terminal.dll");
    if (module == nullptr)
    {
        return HRESULT_FROM_WIN32(ERROR_MOD_NOT_FOUND);
    }
    using DebugGetScreenTextFn = HRESULT(__stdcall*)(ITerminal*, TerminalOwnedUtf16*) noexcept;
#pragma warning(push)
#pragma warning(disable : 4191) // GetProcAddress returns FARPROC; this test-only export has the declared ABI.
    const auto getText = reinterpret_cast<DebugGetScreenTextFn>(GetProcAddress(module, "RedSalamanderTerminalDebugGetScreenText"));
#pragma warning(pop)
    if (getText == nullptr)
    {
        return E_NOINTERFACE;
    }
    TerminalOwnedUtf16 owned{};
    const HRESULT hr = getText(terminal, &owned);
    wil::unique_cotaskmem_string storage(owned.data);
    if (SUCCEEDED(hr) && storage)
    {
        text.assign(storage.get(), owned.length);
    }
    return hr;
}

HRESULT DebugTerminateRootProcess(ITerminal* terminal, uint32_t exitCode) noexcept
{
    if (terminal == nullptr)
    {
        return E_INVALIDARG;
    }
    const HMODULE module = GetModuleHandleW(L"Terminal.dll");
    if (module == nullptr)
    {
        return HRESULT_FROM_WIN32(ERROR_MOD_NOT_FOUND);
    }
    using DebugTerminateRootProcessFn = HRESULT(__stdcall*)(ITerminal*, uint32_t);
#pragma warning(push)
#pragma warning(disable : 4191) // GetProcAddress returns FARPROC; this test-only export has the declared ABI.
    const auto terminate = reinterpret_cast<DebugTerminateRootProcessFn>(GetProcAddress(module, "RedSalamanderTerminalDebugTerminateRootProcess"));
#pragma warning(pop)
    return terminate != nullptr ? terminate(terminal, exitCode) : E_NOINTERFACE;
}

HRESULT DebugRunCommandExperiencePerfSelfTests(unsigned int* passedTests, unsigned int* failedTests) noexcept
{
    if (passedTests == nullptr || failedTests == nullptr)
    {
        return E_POINTER;
    }
    const HMODULE module = GetModuleHandleW(L"Terminal.dll");
    if (module == nullptr)
    {
        return HRESULT_FROM_WIN32(ERROR_MOD_NOT_FOUND);
    }
    using DebugPerfSelfTestsFn = HRESULT(__stdcall*)(unsigned int*, unsigned int*);
#pragma warning(push)
#pragma warning(disable : 4191) // GetProcAddress returns FARPROC; this test-only export has the declared ABI.
    const auto run = reinterpret_cast<DebugPerfSelfTestsFn>(GetProcAddress(module, "RedSalamanderTerminalDebugCommandExperiencePerfSelfTests"));
#pragma warning(pop)
    return run != nullptr ? run(passedTests, failedTests) : E_NOINTERFACE;
}
#endif
} // namespace TerminalHostSupport
