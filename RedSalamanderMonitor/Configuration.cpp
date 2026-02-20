#include "Configuration.h"
#include "Framework.h"

#pragma warning(push)
#pragma warning(                                                                                                                                               \
    disable : 4625 4626 5026 5027) // WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#include <wil/registry.h>
#include <wil/registry_helpers.h>
#pragma warning(pop)

#include <Objbase.h>
#include <WTypes.h>

static constexpr PCWSTR kRegistryRoot = L"Software\\RedSalamander\\Monitor";

Configuration g_config;

Configuration::Configuration()
{
    ; // Initialize the configuration with default values
}

bool Configuration::Load()
{
    wil::unique_hkey hKey;

    // Try new registry path first, fall back to legacy path for migration
    auto hRes = wil::reg::open_unique_key_nothrow(HKEY_CURRENT_USER, kRegistryRoot, hKey);
    if (hRes != S_OK)
    {
        hRes = wil::reg::open_unique_key_nothrow(HKEY_CURRENT_USER, L"Software\\RedSalamander\\Bug Report", hKey);
    }

    if (hRes == S_OK)
    {
        // Load filter settings
        DWORD filterMaskValue = 0x1F;
        hRes                  = wil::reg::get_value_nothrow<DWORD>(hKey.get(), L"FilterMask", &filterMaskValue);
        if (hRes == S_OK)
        {
            filterMask = filterMaskValue;
        }

        DWORD presetValue = static_cast<DWORD>(-1);
        hRes              = wil::reg::get_value_nothrow<DWORD>(hKey.get(), L"LastFilterPreset", &presetValue);
        if (hRes == S_OK)
        {
            lastFilterPreset = static_cast<int>(presetValue);
        }

        AddLine(L"Configuration loaded successfully.");
        return true;
    }

    // If the key does not exist, initialize with default values
    filterMask       = 0x1F;
    lastFilterPreset = -1;
    return false;
}

bool Configuration::Save()
{
    wil::unique_hkey hKey;
    auto hRes = wil::reg::create_unique_key_nothrow(HKEY_CURRENT_USER, kRegistryRoot, hKey);
    if (hRes != S_OK)
    {
        return false;
    }

    // Save filter settings
    hRes = wil::reg::set_value_nothrow<DWORD>(hKey.get(), L"FilterMask", static_cast<DWORD>(filterMask));
    if (hRes != S_OK)
    {
        return false;
    }

    hRes = wil::reg::set_value_nothrow<DWORD>(hKey.get(), L"LastFilterPreset", static_cast<DWORD>(lastFilterPreset));
    if (hRes != S_OK)
    {
        return false;
    }

    return true;
}
