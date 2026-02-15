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

PCWSTR DB_ROOT_KEY      = L"Software\\RedSalamander\\Bug Report";
PCWSTR CONFIG_EMAIL_REG = L"Email";

Configuration g_config;

Configuration::Configuration()
{
    ; // Initialize the configuration with default values
}

BOOL Configuration::Load()
{
    wil::unique_hkey hKey;
    auto hRes = wil::reg::open_unique_key_nothrow(HKEY_CURRENT_USER, DB_ROOT_KEY, hKey);
    if (hRes == S_OK)
    {
        std::wstring emailParam;
        hRes = wil::reg::get_value_nothrow<std::wstring>(hKey.get(), CONFIG_EMAIL_REG, &emailParam);
        if (hRes == S_OK)
        {
            email = emailParam;
        }

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
        return TRUE;
    }

    // If the key does not exist, initialize with default values
    email            = L"";
    filterMask       = 0x1F;
    lastFilterPreset = -1;
    return FALSE;
}

BOOL Configuration::Save()
{
    wil::unique_hkey hKey;
    auto hRes = wil::reg::create_unique_key_nothrow(HKEY_CURRENT_USER, DB_ROOT_KEY, hKey);
    if (hRes != S_OK)
    {
        // Handle error if setting the value fails
        // TODO: Log the error or notify the user
        return FALSE;
    }

    hRes = wil::reg::set_value_string_nothrow(hKey.get(), CONFIG_EMAIL_REG, email.c_str());
    if (hRes != S_OK)
    {
        // Handle error if setting the value fails
        // TODO: Log the error or notify the user
        return FALSE;
    }

    // Save filter settings
    hRes = wil::reg::set_value_nothrow<DWORD>(hKey.get(), L"FilterMask", static_cast<DWORD>(filterMask));
    if (hRes != S_OK)
    {
        return FALSE;
    }

    hRes = wil::reg::set_value_nothrow<DWORD>(hKey.get(), L"LastFilterPreset", static_cast<DWORD>(lastFilterPreset));
    if (hRes != S_OK)
    {
        return FALSE;
    }

    return TRUE;
}
