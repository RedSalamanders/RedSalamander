#include "pch.h"

#include "HostServices.h"
#include "PlugInterfaces/Host.h"
#include "SessionState.h"

PCWSTR REDSALAMANDER_TEXT_VERSION = L"test";

namespace
{
class NullHost final : public IHost
{
public:
    HRESULT __stdcall QueryInterface(REFIID riid, void** ppvObject) override
    {
        if (! ppvObject)
        {
            return E_POINTER;
        }

        *ppvObject = nullptr;
        if (riid == __uuidof(IUnknown) || riid == __uuidof(IHost))
        {
            *ppvObject = static_cast<IHost*>(this);
            AddRef();
            return S_OK;
        }

        return E_NOINTERFACE;
    }

    ULONG __stdcall AddRef() override
    {
        return 2;
    }

    ULONG __stdcall Release() override
    {
        return 1;
    }
};

NullHost g_nullHost;
} // namespace

IHost* GetHostServices() noexcept
{
    return &g_nullHost;
}

HRESULT HostShowAlert(const HostAlertRequest&, void*) noexcept
{
    return E_NOTIMPL;
}

HRESULT HostClearAlert(HostAlertScope, void*) noexcept
{
    return E_NOTIMPL;
}

HRESULT HostShowPrompt(const HostPromptRequest&, void*, HostPromptResult*) noexcept
{
    return E_NOTIMPL;
}

void HostSetAutoAcceptPrompts(bool) noexcept
{
}

bool HostGetAutoAcceptPrompts() noexcept
{
    return false;
}

bool TryHandleHostServicesWindowMessage(UINT, WPARAM, LPARAM, LRESULT&) noexcept
{
    return false;
}

namespace SessionState
{
void Clear() noexcept
{
}

void UpdateActiveFileSystemPluginIdsAndOperation(std::initializer_list<std::wstring_view>, OperationKind) noexcept
{
}

std::optional<State> TryRead() noexcept
{
    return std::nullopt;
}

std::filesystem::path GetSessionStatePath() noexcept
{
    return {};
}
} // namespace SessionState
