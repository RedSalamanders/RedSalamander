#include <array>
#include <format>
#include <iostream>
#include <span>
#include <string>
#include <string_view>

#include <Windows.h>
#include <unknwn.h>

#define REDSAL_DEFINE_TRACE_PROVIDER
#include "Helpers.h"
#include "FileSystemPathIdentity.h"
#include "PlugInterfaces/Factory.h"
#include "PlugInterfaces/FileSystem.h"
#include "PlugInterfaces/Host.h"
#include "PlugInterfaces/Informations.h"
#include "PlugInterfaces/Viewer.h"

#include <wil/com.h>
#include <wil/resource.h>
#include <yyjson.h>

namespace
{

// ---------------------------------------------------------------------------
// yyjson RAII
// ---------------------------------------------------------------------------
using unique_yyjson_doc = wil::unique_any<yyjson_doc*, decltype(&yyjson_doc_free), yyjson_doc_free>;

// ---------------------------------------------------------------------------
// Minimal IHost stub for filesystem plugin instantiation
// ---------------------------------------------------------------------------
class NullHost final : public IHost
{
public:
    HRESULT __stdcall QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (ppvObject == nullptr)
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

    ULONG __stdcall AddRef() noexcept override
    {
        return 2;
    }

    ULONG __stdcall Release() noexcept override
    {
        return 1;
    }
};

NullHost g_nullHost;

// ---------------------------------------------------------------------------
// Check helper (matches LocalizationTests pattern)
// ---------------------------------------------------------------------------
void Check(bool condition, const wchar_t* message, bool& success) noexcept
{
    if (!condition)
    {
        std::wcerr << L"[ FAILED  ] " << message << L"\n";
        success = false;
        return;
    }
    std::wcout << L"[       OK ] " << message << L"\n";
}

// ---------------------------------------------------------------------------
// Export function typedefs
// ---------------------------------------------------------------------------
using PfnEnumeratePlugins      = HRESULT(__stdcall*)(REFIID, const PluginMetaData**, unsigned int*);
using PfnCreate                = HRESULT(__stdcall*)(REFIID, const FactoryOptions*, IHost*, const wchar_t*, void**);
using PfnGetConfigurationSchema = HRESULT(__stdcall*)(REFIID, const wchar_t*, const char**);
using PfnRunDebugSelfTests     = HRESULT(__stdcall*)(unsigned int*, unsigned int*);

// ---------------------------------------------------------------------------
// Plugin sets (paths relative to the exe output dir; exe runs from .build\x64\Debug\)
// DLLs live in the Plugins\ subfolder per REVIEWER CORRECTION 1.
// ---------------------------------------------------------------------------
constexpr std::array<std::wstring_view, 7> kFilesystemDlls = {
    L"Plugins\\FileSystem.dll",
    L"Plugins\\FileSystem7z.dll",
    L"Plugins\\FileSystemCurl.dll",
    L"Plugins\\FileSystemDummy.dll",
    L"Plugins\\FileSystemGoogleDrive.dll",
    L"Plugins\\FileSystemMicrosoftDrive.dll",
    L"Plugins\\FileSystemS3.dll",
};

constexpr std::array<std::wstring_view, 7> kViewerDlls = {
    L"Plugins\\ViewerText.dll",
    L"Plugins\\ViewerSqlite.dll",
    L"Plugins\\ViewerSpace.dll",
    L"Plugins\\ViewerImgRaw.dll",
    L"Plugins\\ViewerVLC.dll",
    L"Plugins\\ViewerPE.dll",
    L"Plugins\\ViewerWeb.dll",
};

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------
[[nodiscard]] std::wstring GetExeDir() noexcept
{
    wchar_t path[MAX_PATH]{};
    const DWORD len = GetModuleFileNameW(nullptr, path, MAX_PATH);
    if (len == 0)
    {
        return {};
    }
    std::wstring result(path, len);
    const auto pos = result.rfind(L'\\');
    if (pos == std::wstring::npos)
    {
        return {};
    }
    result.resize(pos + 1); // keep trailing backslash
    return result;
}

// Parse a const char* as JSON and return the doc (null on failure).
[[nodiscard]] unique_yyjson_doc ParseJson(const char* utf8) noexcept
{
    if (utf8 == nullptr || utf8[0] == '\0')
    {
        return {};
    }
    yyjson_read_err err{};
    // yyjson_read takes a const char* in read-only mode (no INSITU flag).
    yyjson_doc* doc = yyjson_read_opts(
        const_cast<char*>(utf8),
        std::strlen(utf8),
        YYJSON_READ_NOFLAG,
        nullptr,
        &err);
    return unique_yyjson_doc(doc);
}

// ---------------------------------------------------------------------------
// Step 2: Enumerate-and-validate (metadata + schema, no instance creation)
// ---------------------------------------------------------------------------
bool TestEnumerateAndSchema(std::wstring_view relPath, const IID& expectedIid, bool& success) noexcept
{
    const std::wstring exeDir  = GetExeDir();
    const std::wstring absPath = exeDir + std::wstring(relPath);

    const std::wstring loadMsg = std::format(L"{}: DLL loads", relPath);
    wil::unique_hmodule mod(LoadLibraryExW(absPath.c_str(), nullptr, 0));
    Check(static_cast<bool>(mod), loadMsg.c_str(), success);
    if (!mod)
    {
        return false;
    }

    // Resolve exports
    const auto pfnEnumerate =
        reinterpret_cast<PfnEnumeratePlugins>(GetProcAddress(mod.get(), "RedSalamanderEnumeratePlugins"));
    const auto pfnCreate =
        reinterpret_cast<PfnCreate>(GetProcAddress(mod.get(), "RedSalamanderCreate"));
    const auto pfnSchema =
        reinterpret_cast<PfnGetConfigurationSchema>(GetProcAddress(mod.get(), "RedSalamanderGetConfigurationSchema"));

    Check(pfnEnumerate != nullptr,
          std::format(L"{}: RedSalamanderEnumeratePlugins export resolves", relPath).c_str(),
          success);
    Check(pfnCreate != nullptr,
          std::format(L"{}: RedSalamanderCreate export resolves", relPath).c_str(),
          success);
    Check(pfnSchema != nullptr,
          std::format(L"{}: RedSalamanderGetConfigurationSchema export resolves", relPath).c_str(),
          success);

    if (pfnEnumerate == nullptr || pfnCreate == nullptr || pfnSchema == nullptr)
    {
        return false;
    }

    // Happy-path enumerate with correct IID
    const PluginMetaData* metaData = nullptr;
    unsigned int count             = 0;
    const HRESULT hrEnum           = pfnEnumerate(expectedIid, &metaData, &count);
    Check(hrEnum == S_OK,
          std::format(L"{}: EnumeratePlugins(correct IID) returns S_OK", relPath).c_str(),
          success);
    Check(count >= 1,
          std::format(L"{}: EnumeratePlugins returns count >= 1", relPath).c_str(),
          success);
    Check(metaData != nullptr,
          std::format(L"{}: EnumeratePlugins returns non-null metaData", relPath).c_str(),
          success);

    if (hrEnum != S_OK || metaData == nullptr || count == 0)
    {
        return false;
    }

    // Validate each metadata entry
    for (unsigned int i = 0; i < count; ++i)
    {
        const PluginMetaData& md = metaData[i];
        Check(md.id != nullptr && md.id[0] != L'\0',
              std::format(L"{}: metaData[{}].id is non-empty", relPath, i).c_str(),
              success);
        Check(md.name != nullptr && md.name[0] != L'\0',
              std::format(L"{}: metaData[{}].name is non-empty", relPath, i).c_str(),
              success);
        // version is optional per spec (may be nullptr) — just note its presence
        (void)md.version;
    }

    // Wrong IID must return E_NOINTERFACE and count == 0
    const PluginMetaData* wrongMeta = nullptr;
    unsigned int wrongCount         = 0;
    // Bogus IID that matches neither IFileSystem nor IViewer
    static constexpr BYTE kBogusData4[8] = {0xde, 0xad, 0xde, 0xad, 0xde, 0xad, 0xde, 0xad};
    const IID bogusIid = {0xdeadbeef, 0xdead, 0xdead, {kBogusData4[0], kBogusData4[1], kBogusData4[2], kBogusData4[3],
                                                        kBogusData4[4], kBogusData4[5], kBogusData4[6], kBogusData4[7]}};
    const HRESULT hrWrong = pfnEnumerate(bogusIid, &wrongMeta, &wrongCount);
    Check(hrWrong == E_NOINTERFACE,
          std::format(L"{}: EnumeratePlugins(wrong IID) returns E_NOINTERFACE", relPath).c_str(),
          success);
    Check(wrongCount == 0,
          std::format(L"{}: EnumeratePlugins(wrong IID) returns count == 0", relPath).c_str(),
          success);

    // Schema validation for each plugin id
    for (unsigned int i = 0; i < count; ++i)
    {
        const wchar_t* pluginId    = metaData[i].id;
        const char* schemaJson     = nullptr;
        const HRESULT hrSchema     = pfnSchema(expectedIid, pluginId, &schemaJson);
        const bool schemaOk        = hrSchema == S_OK;
        const bool schemaNotFound  = hrSchema == static_cast<HRESULT>(HRESULT_FROM_WIN32(ERROR_NOT_FOUND));
        Check(schemaOk || schemaNotFound,
              std::format(L"{}: GetConfigurationSchema[{}] returns S_OK or ERROR_NOT_FOUND", relPath, i).c_str(),
              success);
        if (schemaOk && schemaJson != nullptr && schemaJson[0] != '\0')
        {
            unique_yyjson_doc doc = ParseJson(schemaJson);
            Check(static_cast<bool>(doc),
                  std::format(L"{}: GetConfigurationSchema[{}] JSON parses successfully", relPath, i).c_str(),
                  success);
        }
    }

    return true;
}

// ---------------------------------------------------------------------------
// Step 3: GetCapabilities pass (filesystem plugins only)
// ---------------------------------------------------------------------------
bool TestCapabilities(std::wstring_view relPath, bool& success) noexcept
{
    const std::wstring exeDir  = GetExeDir();
    const std::wstring absPath = exeDir + std::wstring(relPath);

    // Module must stay alive while we hold interface pointers from it.
    wil::unique_hmodule mod(LoadLibraryExW(absPath.c_str(), nullptr, 0));
    if (!mod)
    {
        // Already reported in enumerate pass; don't double-report.
        return false;
    }

    const auto pfnEnumerate =
        reinterpret_cast<PfnEnumeratePlugins>(GetProcAddress(mod.get(), "RedSalamanderEnumeratePlugins"));
    const auto pfnCreate =
        reinterpret_cast<PfnCreate>(GetProcAddress(mod.get(), "RedSalamanderCreate"));

    if (pfnEnumerate == nullptr || pfnCreate == nullptr)
    {
        return false;
    }

    const PluginMetaData* metaData = nullptr;
    unsigned int count             = 0;
    if (pfnEnumerate(__uuidof(IFileSystem), &metaData, &count) != S_OK || count == 0 || metaData == nullptr)
    {
        return false;
    }

    for (unsigned int i = 0; i < count; ++i)
    {
        const wchar_t* pluginId = metaData[i].id;

        void* raw = nullptr;
        const HRESULT hrCreate =
            pfnCreate(__uuidof(IFileSystem), nullptr, static_cast<IHost*>(&g_nullHost), pluginId, &raw);
        Check(hrCreate == S_OK,
              std::format(L"{}: Create(pluginId={}) returns S_OK", relPath, pluginId).c_str(),
              success);
        Check(raw != nullptr,
              std::format(L"{}: Create(pluginId={}) returns non-null instance", relPath, pluginId).c_str(),
              success);

        if (hrCreate != S_OK || raw == nullptr)
        {
            continue;
        }

        // Wrap in a smart pointer — IFileSystem inherits IUnknown
        wil::com_ptr_nothrow<IFileSystem> fs;
        fs.attach(static_cast<IFileSystem*>(raw));

        const char* capJson = nullptr;
        const HRESULT hrCap = fs->GetCapabilities(&capJson);
        Check(hrCap == S_OK,
              std::format(L"{}: GetCapabilities(pluginId={}) returns S_OK", relPath, pluginId).c_str(),
              success);
        Check(capJson != nullptr && capJson[0] != '\0',
              std::format(L"{}: GetCapabilities(pluginId={}) returns non-empty JSON", relPath, pluginId).c_str(),
              success);

        if (hrCap != S_OK || capJson == nullptr || capJson[0] == '\0')
        {
            continue;
        }

        unique_yyjson_doc doc = ParseJson(capJson);
        Check(static_cast<bool>(doc),
              std::format(L"{}: GetCapabilities(pluginId={}) JSON parses successfully", relPath, pluginId).c_str(),
              success);

        if (!doc)
        {
            continue;
        }

        yyjson_val* root = yyjson_doc_get_root(doc.get());
        Check(root != nullptr && yyjson_is_obj(root),
              std::format(L"{}: GetCapabilities(pluginId={}) root is object", relPath, pluginId).c_str(),
              success);

        if (root == nullptr || !yyjson_is_obj(root))
        {
            continue;
        }

        // Mandatory fields: "version": 1, "operations" (obj), "concurrency" (obj), "crossFileSystem" (obj)
        yyjson_val* verVal = yyjson_obj_get(root, "version");
        Check(verVal != nullptr && yyjson_is_int(verVal) && yyjson_get_int(verVal) == 1,
              std::format(L"{}: GetCapabilities(pluginId={}) has \"version\": 1", relPath, pluginId).c_str(),
              success);

        yyjson_val* opsVal = yyjson_obj_get(root, "operations");
        Check(opsVal != nullptr && yyjson_is_obj(opsVal),
              std::format(L"{}: GetCapabilities(pluginId={}) has \"operations\" object", relPath, pluginId).c_str(),
              success);

        yyjson_val* concVal = yyjson_obj_get(root, "concurrency");
        Check(concVal != nullptr && yyjson_is_obj(concVal),
              std::format(L"{}: GetCapabilities(pluginId={}) has \"concurrency\" object", relPath, pluginId).c_str(),
              success);

        yyjson_val* crossVal = yyjson_obj_get(root, "crossFileSystem");
        Check(crossVal != nullptr && yyjson_is_obj(crossVal),
              std::format(L"{}: GetCapabilities(pluginId={}) has \"crossFileSystem\" object", relPath, pluginId).c_str(),
              success);

        // pathIdentity is now a MANDATORY block: the host (ParseFileSystemCapabilitiesJson ->
        // TryParseFileSystemPathIdentityContract, FolderWindow.FileOperations.cpp:139) fails the WHOLE
        // capability set closed -- disabling copy/move/delete and cross-filesystem operations -- if it
        // is absent or fails the strict parser. Assert each plugin's block is present and parses under
        // the exact host contract so a typo (preferredSeparator not in acceptedSeparators, a misspelled
        // enum, a missing bool, normalization != "none", ...) can never silently brick the plugin in
        // production while this contract test stays green.
        yyjson_val* pathIdVal = yyjson_obj_get(root, "pathIdentity");
        Check(pathIdVal != nullptr && yyjson_is_obj(pathIdVal),
              std::format(L"{}: GetCapabilities(pluginId={}) has \"pathIdentity\" object", relPath, pluginId).c_str(),
              success);

        const std::optional<FileSystemPathIdentity> parsedPathIdentity = TryParseFileSystemPathIdentityContract(std::string_view(capJson), {});
        Check(parsedPathIdentity.has_value(),
              std::format(L"{}: GetCapabilities(pluginId={}) pathIdentity parses under the host contract", relPath, pluginId).c_str(),
              success);

        // capJson pointer is plugin-owned; do NOT free it.
        // fs (IFileSystem) is released here by wil::com_ptr_nothrow going out of scope.
        // doc (yyjson_doc) is released here.
        // mod stays alive (declared first in the enclosing scope).
    }

    return true;
}

// ---------------------------------------------------------------------------
// Step 4: Negative-contract test (bogus plugin id on FileSystemDummy only)
// ---------------------------------------------------------------------------
bool TestBogusPluginId(bool& success) noexcept
{
    const std::wstring exeDir  = GetExeDir();
    const std::wstring absPath = exeDir + L"Plugins\\FileSystemDummy.dll";

    wil::unique_hmodule mod(LoadLibraryExW(absPath.c_str(), nullptr, 0));
    Check(static_cast<bool>(mod), L"FileSystemDummy.dll: loads for negative-id test", success);
    if (!mod)
    {
        return false;
    }

    const auto pfnCreate =
        reinterpret_cast<PfnCreate>(GetProcAddress(mod.get(), "RedSalamanderCreate"));
    Check(pfnCreate != nullptr, L"FileSystemDummy.dll: RedSalamanderCreate resolves for negative-id test", success);
    if (pfnCreate == nullptr)
    {
        return false;
    }

    void* raw = nullptr;
    const HRESULT hr =
        pfnCreate(__uuidof(IFileSystem), nullptr, static_cast<IHost*>(&g_nullHost), L"no/such-plugin", &raw);
    Check(hr == HRESULT_FROM_WIN32(ERROR_NOT_FOUND),
          L"FileSystemDummy.dll: Create(bogus-id) returns HRESULT_FROM_WIN32(ERROR_NOT_FOUND)",
          success);
    Check(raw == nullptr, L"FileSystemDummy.dll: Create(bogus-id) returns null result pointer", success);

    return true;
}

bool TestPluginDebugSelfTests(std::wstring_view relativeDllPath, const char* exportName, std::wstring_view label, bool& success) noexcept
{
    const std::wstring exeDir  = GetExeDir();
    const std::wstring absPath = exeDir + std::wstring(relativeDllPath);

    wil::unique_hmodule mod(LoadLibraryExW(absPath.c_str(), nullptr, 0));
    const std::wstring loadMessage = std::format(L"{}: loads for debug selftests", label);
    Check(static_cast<bool>(mod), loadMessage.c_str(), success);
    if (!mod)
    {
        return false;
    }

    const auto pfnSelfTests =
        reinterpret_cast<PfnRunDebugSelfTests>(GetProcAddress(mod.get(), exportName));
    if (pfnSelfTests == nullptr)
    {
        std::wcout << L"[       OK ] " << label << L": debug selftest export absent in this configuration\n";
        return true;
    }

    unsigned int passed = 0;
    unsigned int failed = 0;
    const HRESULT hr    = pfnSelfTests(&passed, &failed);
    const std::wstring passMessage =
        std::format(L"{}: debug selftests pass (passed={}, failed={}, hr=0x{:08X})", label, passed, failed, static_cast<unsigned long>(hr));
    Check(SUCCEEDED(hr) && failed == 0 && passed > 0,
          passMessage.c_str(),
          success);
    return true;
}

// ---------------------------------------------------------------------------
// Run all tests for a DLL set
// ---------------------------------------------------------------------------
bool RunEnumerateAndSchemaPass(std::span<const std::wstring_view> dlls, const IID& iid) noexcept
{
    bool success = true;
    for (const std::wstring_view relPath : dlls)
    {
        TestEnumerateAndSchema(relPath, iid, success);
    }
    return success;
}

bool RunCapabilitiesPass(std::span<const std::wstring_view> dlls) noexcept
{
    bool success = true;
    for (const std::wstring_view relPath : dlls)
    {
        TestCapabilities(relPath, success);
    }
    return success;
}

} // namespace

int wmain()
{
    // Configure the DLL search path so that plugin DLLs find their transitive dependencies:
    // - Plugins\ contains aws-*.dll, libcurl-d.dll, sqlite3.dll, and so on.
    // - The exe dir (parent of Plugins\) contains Common.dll and other shared DLLs.
    // We use SetDllDirectoryW to add Plugins\ to the standard DLL search path (it is inserted
    // between the app dir and the system/PATH directories in the standard search order).
    const std::wstring exeDir     = GetExeDir();
    const std::wstring pluginsDir = exeDir + L"Plugins";
    SetDllDirectoryW(pluginsDir.c_str());
    // AddDllDirectory is also called for LOAD_LIBRARY_SEARCH_USER_DIRS coverage.
    const DLL_DIRECTORY_COOKIE pluginsDirCookie =
        AddDllDirectory(pluginsDir.c_str());

    bool success = true;

    std::wcout << L"[ RUN      ] Step2: EnumerateAndSchema - filesystem plugins\n";
    success = RunEnumerateAndSchemaPass(kFilesystemDlls, __uuidof(IFileSystem)) && success;

    std::wcout << L"[ RUN      ] Step2: EnumerateAndSchema - viewer plugins\n";
    success = RunEnumerateAndSchemaPass(kViewerDlls, __uuidof(IViewer)) && success;

    std::wcout << L"[ RUN      ] Step3: GetCapabilities - filesystem plugins\n";
    success = RunCapabilitiesPass(kFilesystemDlls) && success;

    std::wcout << L"[ RUN      ] Step4: Negative bogus plugin id\n";
    {
        bool negSuccess = true;
        TestBogusPluginId(negSuccess);
        success = negSuccess && success;
    }

    std::wcout << L"[ RUN      ] Step5: Optional plugin debug selftests\n";
    {
        bool debugSuccess = true;
        TestPluginDebugSelfTests(L"Plugins\\FileSystem.dll", "RedSalamanderFileSystemDebugSelfTests", L"FileSystem.dll", debugSuccess);
        TestPluginDebugSelfTests(
            L"Plugins\\FileSystemMicrosoftDrive.dll", "RedSalamanderMicrosoftDriveDebugSelfTests", L"FileSystemMicrosoftDrive.dll", debugSuccess);
        TestPluginDebugSelfTests(L"Plugins\\FileSystemS3.dll", "RedSalamanderS3DebugSelfTests", L"FileSystemS3.dll", debugSuccess);
        TestPluginDebugSelfTests(L"Plugins\\FileSystemCurl.dll", "RedSalamanderCurlDebugSelfTests", L"FileSystemCurl.dll", debugSuccess);
        success = debugSuccess && success;
    }

    if (pluginsDirCookie != nullptr)
    {
        RemoveDllDirectory(pluginsDirCookie);
    }

    std::wcout << (success ? L"PluginContractTests passed.\n" : L"PluginContractTests FAILED.\n");
    return success ? 0 : 1;
}
