#include <cstddef>
#include <filesystem>
#include <fstream>
#include <iostream>
#include <iterator>
#include <string>
#include <vector>

#include <Windows.h>

#define REDSAL_DEFINE_TRACE_PROVIDER
#include "Helpers.h"
#include "LocalizationManager.h"
#include "SettingsStore.h"
#include "Win32CallbackHelpers.h"

#include "resource.h"

namespace
{
constexpr wchar_t kLocalizationTestCustomControlClass[] = L"LocalizationTestCustomControl";

INT_PTR CALLBACK LocalizationTestDialogProc(HWND, UINT message, WPARAM, LPARAM) noexcept
{
    return message == WM_INITDIALOG ? static_cast<INT_PTR>(TRUE) : static_cast<INT_PTR>(FALSE);
}

LRESULT CALLBACK LocalizationTestCustomControlProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
{
    return DefWindowProcW(hwnd, message, wParam, lParam);
}

void Check(bool condition, const wchar_t* message, bool& success) noexcept
{
    if (! condition)
    {
        std::wcerr << L"[ FAILED  ] " << message << L"\n";
        success = false;
        return;
    }

    std::wcout << L"[       OK ] " << message << L"\n";
}

bool EnsureLocalizationTestCustomControlClass(HINSTANCE instance) noexcept
{
    WNDCLASSW existing{};
    if (GetClassInfoW(instance, kLocalizationTestCustomControlClass, &existing) != FALSE)
    {
        return true;
    }

    WNDCLASSW wc{};
    wc.lpfnWndProc   = LocalizationTestCustomControlProc;
    wc.hInstance     = instance;
    wc.lpszClassName = kLocalizationTestCustomControlClass;
    return RegisterClassW(&wc) != 0 || GetLastError() == ERROR_CLASS_ALREADY_EXISTS;
}

void CheckDialogCustomChild(HWND dialog, const wchar_t* message, bool& success) noexcept
{
    Check(dialog != nullptr && GetDlgItem(dialog, IDC_LOCALIZATION_TEST_CUSTOM) != nullptr, message, success);
}

bool TestEmbeddedFallbackWithoutSatellite() noexcept
{
    bool success = true;

    HINSTANCE instance = GetModuleHandleW(nullptr);
    Check(instance != nullptr, L"test module handle resolves", success);

    const HRESULT registerHr = Localization::RegisterResourceOwner(L"LocalizationTests", instance);
    Check(SUCCEEDED(registerHr), L"resource owner registration succeeds", success);

    const std::wstring present = LoadStringResource(instance, IDS_LOCALIZATION_TEST_PRESENT);
    Check(present == L"Embedded English test string", L"existing embedded resource loads without satellite", success);

    const std::wstring missing = LoadStringResource(instance, IDS_LOCALIZATION_TEST_MISSING);
    Check(missing.empty(), L"missing embedded resource returns empty", success);

    Localization::UnregisterResourceOwner(instance);
    return success;
}

bool TestSatelliteOverridesEmbeddedString() noexcept
{
    bool success = true;

    HINSTANCE instance = GetModuleHandleW(nullptr);
    Check(instance != nullptr, L"test module handle resolves for satellite lookup", success);
    Check(EnsureLocalizationTestCustomControlClass(instance), L"test custom dialog control class registers on embedded module", success);

    const HRESULT registerHr = Localization::RegisterResourceOwner(L"LocalizationTests", instance);
    Check(SUCCEEDED(registerHr), L"resource owner registration succeeds for satellite lookup", success);

    Localization::LanguagePreference preference;
    preference.kind       = Localization::LanguagePreferenceKind::Culture;
    preference.culture    = L"fr-FR";
    const HRESULT applyHr = Localization::ApplyLanguagePreference(preference);
    Check(SUCCEEDED(applyHr), L"fr-FR language preference applies", success);

    const std::wstring present = LoadStringResource(instance, IDS_LOCALIZATION_TEST_PRESENT);
    Check(present == L"Chaine de test francaise", L"satellite string overrides embedded English", success);

    const std::wstring embeddedPresent = LoadEmbeddedStringResource(instance, IDS_LOCALIZATION_TEST_PRESENT);
    Check(embeddedPresent == L"Embedded English test string", L"embedded string helper bypasses satellite string", success);

    const std::wstring embeddedFormatted = FormatEmbeddedStringResource(instance, IDS_LOCALIZATION_TEST_EMBEDDED_FORMAT, std::wstring(L"token"), 42u);
    Check(embeddedFormatted == L"Embedded token 0x0000002A", L"embedded format helper bypasses satellite format string", success);

    wil::unique_hmenu menu(Localization::LoadMenuResource(instance, IDR_LOCALIZATION_TEST_MENU));
    Check(static_cast<bool>(menu), L"satellite menu resource loads", success);
    if (menu)
    {
        wchar_t text[64]{};
        const int chars = GetMenuStringW(menu.get(), 0, text, static_cast<int>(std::size(text)), MF_BYPOSITION);
        Check(chars > 0 && std::wstring_view(text) == L"&Fichier", L"satellite menu overrides embedded English", success);
    }

    const std::wstring missing = LoadStringResource(instance, IDS_LOCALIZATION_TEST_MISSING);
    Check(missing.empty(), L"missing satellite and embedded resource returns empty", success);

    const std::wstring embeddedOnly = LoadStringResource(instance, IDS_LOCALIZATION_TEST_EMBEDDED_ONLY);
    Check(embeddedOnly == L"Embedded only fallback string", L"missing satellite string falls back to embedded English", success);

    const Localization::ResourceLookupResult dialogResource =
        Localization::FindLocalizedResourceHandle(instance, MAKEINTRESOURCEW(IDD_LOCALIZATION_TEST_DIALOG), RT_DIALOG);
    Check(static_cast<bool>(dialogResource) && dialogResource.instance != instance, L"satellite dialog resource handle resolves", success);

    std::vector<std::byte> dialogBytes;
    Check(Localization::LoadResourceBytes(instance, MAKEINTRESOURCEW(IDD_LOCALIZATION_TEST_DIALOG), RT_DIALOG, dialogBytes) && ! dialogBytes.empty(),
          L"satellite dialog template bytes load",
          success);

    wil::unique_hwnd dialog(RedSalamander::Win32Callback::CreateDialogParamResourceNoThrow(
        instance, MAKEINTRESOURCEW(IDD_LOCALIZATION_TEST_DIALOG), nullptr, LocalizationTestDialogProc, 0));
    Check(static_cast<bool>(dialog), L"localized dialog template creates a modeless dialog", success);
    if (dialog)
    {
        CheckDialogCustomChild(dialog.get(), L"localized dialog creates custom child class registered by embedded module", success);
        wchar_t caption[64]{};
        const int chars = GetWindowTextW(dialog.get(), caption, static_cast<int>(std::size(caption)));
        Check(chars > 0 && std::wstring_view(caption) == L"Dialogue FR", L"satellite dialog overrides embedded English caption", success);
    }

    Localization::UnregisterResourceOwner(instance);
    const std::wstring unregistered = LoadStringResource(instance, IDS_LOCALIZATION_TEST_PRESENT);
    Check(unregistered == L"Embedded English test string", L"unregistered resource owner falls back to embedded English", success);

    Localization::LanguagePreference systemPreference;
    static_cast<void>(Localization::ApplyLanguagePreference(systemPreference));
    return success;
}

bool TestInvalidCultureFallsBackToEmbedded() noexcept
{
    bool success = true;

    HINSTANCE instance = GetModuleHandleW(nullptr);
    Check(instance != nullptr, L"test module handle resolves for invalid culture fallback", success);
    Check(EnsureLocalizationTestCustomControlClass(instance), L"test custom dialog control class registers for embedded fallback", success);

    const HRESULT registerHr = Localization::RegisterResourceOwner(L"LocalizationTests", instance);
    Check(SUCCEEDED(registerHr), L"resource owner registration succeeds for invalid culture fallback", success);

    Localization::LanguagePreference preference;
    preference.kind       = Localization::LanguagePreferenceKind::Culture;
    preference.culture    = L"zz-ZZ";
    const HRESULT applyHr = Localization::ApplyLanguagePreference(preference);
    Check(SUCCEEDED(applyHr), L"missing concrete culture preference applies", success);

    const std::wstring present = LoadStringResource(instance, IDS_LOCALIZATION_TEST_PRESENT);
    Check(present == L"Embedded English test string", L"missing culture satellite falls back to embedded English", success);

    wil::unique_hmenu menu(Localization::LoadMenuResource(instance, IDR_LOCALIZATION_TEST_MENU));
    Check(static_cast<bool>(menu), L"embedded menu fallback loads for missing culture satellite", success);
    if (menu)
    {
        wchar_t text[64]{};
        const int chars = GetMenuStringW(menu.get(), 0, text, static_cast<int>(std::size(text)), MF_BYPOSITION);
        Check(chars > 0 && std::wstring_view(text) == L"&File", L"missing culture menu falls back to embedded English", success);
    }

    const Localization::ResourceLookupResult dialogResource =
        Localization::FindLocalizedResourceHandle(instance, MAKEINTRESOURCEW(IDD_LOCALIZATION_TEST_DIALOG), RT_DIALOG);
    Check(static_cast<bool>(dialogResource) && dialogResource.instance == instance, L"missing culture dialog falls back to embedded English handle", success);

    wil::unique_hwnd dialog(RedSalamander::Win32Callback::CreateDialogParamResourceNoThrow(
        instance, MAKEINTRESOURCEW(IDD_LOCALIZATION_TEST_DIALOG), nullptr, LocalizationTestDialogProc, 0));
    Check(static_cast<bool>(dialog), L"embedded dialog fallback creates a modeless dialog", success);
    if (dialog)
    {
        CheckDialogCustomChild(dialog.get(), L"embedded dialog fallback creates custom child class registered by embedded module", success);
        wchar_t caption[64]{};
        const int chars = GetWindowTextW(dialog.get(), caption, static_cast<int>(std::size(caption)));
        Check(chars > 0 && std::wstring_view(caption) == L"Embedded Dialog", L"missing culture dialog falls back to embedded English caption", success);
    }

    Localization::LanguagePreference systemPreference;
    static_cast<void>(Localization::ApplyLanguagePreference(systemPreference));
    Localization::UnregisterResourceOwner(instance);
    return success;
}

std::wstring UniqueSettingsAppId(std::wstring_view suffix)
{
    std::wstring appId = L"RedSalamanderLocalizationTests-";
    appId.append(suffix);
    return appId;
}

void RemoveSettingsFiles(std::wstring_view appId) noexcept
{
    std::error_code ec;
    std::filesystem::remove(Common::Settings::GetSettingsPath(appId), ec);
    std::filesystem::remove(Common::Settings::GetSettingsSchemaPath(appId), ec);
}

bool WriteSettingsJson(std::wstring_view appId, std::string_view json) noexcept
{
    const std::filesystem::path path = Common::Settings::GetSettingsPath(appId);
    if (path.empty())
    {
        return false;
    }

    std::error_code ec;
    std::filesystem::create_directories(path.parent_path(), ec);
    if (ec)
    {
        return false;
    }

    std::ofstream stream(path, std::ios::binary | std::ios::trunc);
    if (! stream)
    {
        return false;
    }

    stream.write(json.data(), static_cast<std::streamsize>(json.size()));
    return stream.good();
}

bool ReadSettingsJson(std::wstring_view appId, std::string& out) noexcept
{
    out.clear();
    const std::filesystem::path path = Common::Settings::GetSettingsPath(appId);
    if (path.empty())
    {
        return false;
    }

    std::ifstream stream(path, std::ios::binary);
    if (! stream)
    {
        return false;
    }

    out.assign(std::istreambuf_iterator<char>(stream), std::istreambuf_iterator<char>());
    return stream.good() || stream.eof();
}

bool LoadSettingsForJson(std::wstring_view appId, std::string_view json, Common::Settings::Settings& out, bool& success, const wchar_t* writeMessage)
{
    RemoveSettingsFiles(appId);
    Check(WriteSettingsJson(appId, json), writeMessage, success);
    if (! success)
    {
        return false;
    }

    const HRESULT loadHr = Common::Settings::TryLoadSettingsNoRecovery(appId, out);
    Check(SUCCEEDED(loadHr), L"settings JSON loads without recovery", success);
    RemoveSettingsFiles(appId);
    return SUCCEEDED(loadHr);
}

bool TestUiLanguageSettings() noexcept
{
    bool success = true;

    {
        Common::Settings::Settings loaded;
        const std::wstring appId = UniqueSettingsAppId(L"MissingLanguage");
        if (LoadSettingsForJson(appId, R"({"schemaVersion":16,"ui":{}})", loaded, success, L"settings JSON with empty ui is written"))
        {
            Check(loaded.ui.has_value(), L"empty ui object loads ui settings", success);
            Check(loaded.ui.value().language == L"system", L"missing ui.language loads as system", success);
        }
    }

    {
        Common::Settings::Settings loaded;
        const std::wstring appId = UniqueSettingsAppId(L"SystemLanguage");
        if (LoadSettingsForJson(appId, R"({"schemaVersion":16,"ui":{"language":"system"}})", loaded, success, L"settings JSON with system language is written"))
        {
            Check(loaded.ui.has_value(), L"system language loads ui settings", success);
            Check(loaded.ui.value().language == L"system", L"ui.language system loads as system", success);
        }
    }

    {
        Common::Settings::Settings loaded;
        const std::wstring appId = UniqueSettingsAppId(L"FrenchLanguage");
        if (LoadSettingsForJson(appId, R"({"schemaVersion":16,"ui":{"language":"fr-FR"}})", loaded, success, L"settings JSON with fr-FR language is written"))
        {
            Check(loaded.ui.has_value(), L"fr-FR language loads ui settings", success);
            Check(loaded.ui.value().language == L"fr-FR", L"ui.language fr-FR loads as fr-FR", success);
        }
    }

    {
        Common::Settings::Settings loaded;
        const std::wstring appId = UniqueSettingsAppId(L"ParentFrenchLanguage");
        if (LoadSettingsForJson(appId, R"({"schemaVersion":16,"ui":{"language":"fr"}})", loaded, success, L"settings JSON with fr language is written"))
        {
            Check(loaded.ui.has_value(), L"fr language loads ui settings", success);
            Check(loaded.ui.value().language == L"fr", L"ui.language fr loads as fr", success);
        }
    }

    {
        const std::wstring appId = UniqueSettingsAppId(L"InvalidLanguage");
        Common::Settings::Settings loaded;
        if (LoadSettingsForJson(
                appId, R"({"schemaVersion":16,"ui":{"language":"..\\bad"}})", loaded, success, L"settings JSON with invalid language is written"))
        {
            Check(loaded.ui.has_value(), L"invalid language loads ui settings", success);
            Check(loaded.ui.value().language == L"system", L"invalid ui.language loads as system", success);
        }
    }

    {
        const std::wstring appId = UniqueSettingsAppId(L"SaveDefaultLanguage");
        RemoveSettingsFiles(appId);

        Common::Settings::Settings settings;
        settings.ui          = Common::Settings::UiSettings{};
        const HRESULT saveHr = Common::Settings::SaveSettings(appId, settings);
        Check(SUCCEEDED(saveHr), L"default language settings save", success);

        std::string json;
        Check(ReadSettingsJson(appId, json), L"default language settings JSON can be read", success);
        Check(json.find("\"language\"") == std::string::npos, L"default ui.language is omitted on save", success);
        RemoveSettingsFiles(appId);
    }

    {
        const std::wstring appId = UniqueSettingsAppId(L"SaveFrenchLanguage");
        RemoveSettingsFiles(appId);

        Common::Settings::Settings settings;
        settings.ui                  = Common::Settings::UiSettings{};
        settings.ui.value().language = L"fr-FR";
        const HRESULT saveHr         = Common::Settings::SaveSettings(appId, settings);
        Check(SUCCEEDED(saveHr), L"fr-FR language settings save", success);

        std::string json;
        Check(ReadSettingsJson(appId, json), L"fr-FR language settings JSON can be read", success);
        Check(
            json.find("\"language\"") != std::string::npos && json.find("\"fr-FR\"") != std::string::npos, L"concrete ui.language is written on save", success);

        Common::Settings::Settings loaded;
        const HRESULT loadHr = Common::Settings::TryLoadSettingsNoRecovery(appId, loaded);
        Check(SUCCEEDED(loadHr), L"saved fr-FR language settings reload", success);
        Check(loaded.ui.has_value(), L"saved fr-FR language reloads ui settings", success);
        Check(loaded.ui.value().language == L"fr-FR", L"saved ui.language reloads as fr-FR", success);
        RemoveSettingsFiles(appId);
    }

    return success;
}
} // namespace

int wmain()
{
    bool success = true;
    std::wcout << L"[ RUN      ] TestEmbeddedFallbackWithoutSatellite\n";
    success = TestEmbeddedFallbackWithoutSatellite() && success;
    std::wcout << L"[ RUN      ] TestSatelliteOverridesEmbeddedString\n";
    success = TestSatelliteOverridesEmbeddedString() && success;
    std::wcout << L"[ RUN      ] TestInvalidCultureFallsBackToEmbedded\n";
    success = TestInvalidCultureFallsBackToEmbedded() && success;
    std::wcout << L"[ RUN      ] TestUiLanguageSettings\n";
    success = TestUiLanguageSettings() && success;
    std::wcout << (success ? L"LocalizationTests passed.\n" : L"LocalizationTests failed.\n");
    return success ? 0 : 1;
}
