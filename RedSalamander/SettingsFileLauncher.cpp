#include "Framework.h"

#include "SettingsFileLauncher.h"

#include "SettingsHotReload.h"

#include <shellapi.h>

namespace
{
[[nodiscard]] DWORD ShellExecuteError(const INT_PTR result) noexcept
{
    switch (result)
    {
        case 0: return ERROR_NOT_ENOUGH_MEMORY;
        case SE_ERR_FNF: return ERROR_FILE_NOT_FOUND;
        case SE_ERR_PNF: return ERROR_PATH_NOT_FOUND;
        case SE_ERR_ACCESSDENIED: return ERROR_ACCESS_DENIED;
        case SE_ERR_OOM: return ERROR_NOT_ENOUGH_MEMORY;
        case SE_ERR_DLLNOTFOUND: return ERROR_MOD_NOT_FOUND;
        case SE_ERR_SHARE: return ERROR_SHARING_VIOLATION;
        case SE_ERR_ASSOCINCOMPLETE:
        case SE_ERR_NOASSOC: return ERROR_NO_ASSOCIATION;
        case SE_ERR_DDETIMEOUT: return ERROR_TIMEOUT;
        case SE_ERR_DDEFAIL:
        case SE_ERR_DDEBUSY: return ERROR_BUSY;
        default: return ERROR_GEN_FAILURE;
    }
}
}

HRESULT SettingsFileLauncher::Open(HWND owner,
                                   std::wstring_view appId,
                                   Common::Settings::Settings settingsToCreate,
                                   std::filesystem::path& outPath) noexcept
{
    outPath.clear();
    if (appId.empty())
    {
        return E_INVALIDARG;
    }

    outPath = Common::Settings::GetSettingsPath(appId);
    if (outPath.empty() || ! outPath.has_parent_path())
    {
        return HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND);
    }
    const std::filesystem::path parent = outPath.parent_path();
    std::error_code error;
    std::filesystem::create_directories(parent, error);
    if (error)
    {
        return HRESULT_FROM_WIN32(error.value() > 0 ? static_cast<DWORD>(error.value()) : ERROR_CANNOT_MAKE);
    }

    const bool exists = std::filesystem::exists(outPath, error);
    if (error)
    {
        return HRESULT_FROM_WIN32(error.value() > 0 ? static_cast<DWORD>(error.value()) : ERROR_FILE_NOT_FOUND);
    }
    if (! exists)
    {
        const HRESULT saveHr = SettingsHotReload::SaveSettingsAndSchema(appId, settingsToCreate);
        if (FAILED(saveHr))
        {
            return saveHr;
        }
    }

    const INT_PTR result = reinterpret_cast<INT_PTR>(ShellExecuteW(owner, L"open", outPath.c_str(), nullptr, parent.c_str(), SW_SHOWNORMAL));
    return result > 32 ? S_OK : HRESULT_FROM_WIN32(ShellExecuteError(result));
}
