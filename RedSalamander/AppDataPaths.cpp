#include "AppDataPaths.h"

#include "Framework.h"

#include <string>

#include <shlobj_core.h>

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027)
#include <wil/resource.h>
#pragma warning(pop)

namespace AppDataPaths
{
std::filesystem::path GetLocalAppDataPath() noexcept
{
    wil::unique_cotaskmem_string localAppData;
    const HRESULT hr = SHGetKnownFolderPath(FOLDERID_LocalAppData, 0, nullptr, localAppData.put());
    if (SUCCEEDED(hr) && localAppData)
    {
        return std::filesystem::path(localAppData.get());
    }

    const DWORD required = GetEnvironmentVariableW(L"LOCALAPPDATA", nullptr, 0);
    if (required == 0)
    {
        return {};
    }

    std::wstring buffer(static_cast<size_t>(required), L'\0');
    const DWORD written = GetEnvironmentVariableW(L"LOCALAPPDATA", buffer.data(), required);
    if (written == 0 || written >= required)
    {
        return {};
    }

    buffer.resize(static_cast<size_t>(written));
    return std::filesystem::path(buffer);
}
} // namespace AppDataPaths

