#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <array>
#include <cstring>
#include <format>
#include <limits>
#include <mutex>
#include <new>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4514 28182) // WIL headers: deleted copy/move and unused inline helpers
#include <wil/com.h>
#include <wil/resource.h>
#pragma warning(pop)

#define PLUGFACTORY_EXPORTS
#include "PlugInterfaces/Factory.h"

#define REDSAL_DEFINE_TRACE_PROVIDER
#include "FileSystemMtpResources.h"
#include "Helpers.h"

#include "FileSystemMtp.h"

#include "PlugInterfaces/FactoryImpl.h"

extern HINSTANCE g_hInstance;

namespace
{
struct LocalizedPluginMetaDataSet
{
    std::wstring name;
    std::wstring description;
    std::array<PluginMetaData, 1> plugins{};

    LocalizedPluginMetaDataSet()
    {
        LoadEmbeddedStringResource(g_hInstance, IDS_FILESYSTEMMTP_NAME, name);
        LoadStringResource(g_hInstance, IDS_FILESYSTEMMTP_DESCRIPTION, description);

        plugins = {{
            {
                .id          = L"builtin/file-system-mtp",
                .shortId     = L"mtp",
                .name        = name.c_str(),
                .description = description.c_str(),
                .author      = L"RedSalamander",
                .version     = VERSINFO_PLUGIN_VERSION,
            },
        }};
    }
};

[[nodiscard]] const LocalizedPluginMetaDataSet& GetPluginMetaDataSet() noexcept
{
    static const LocalizedPluginMetaDataSet data;
    return data;
}

const PluginMetaData* GetMetaDataMtp() noexcept
{
    return &GetPluginMetaDataSet().plugins[0];
}

const char* GetSchemaMtp() noexcept
{
    return GetFileSystemMtpStaticConfigurationSchema();
}

[[nodiscard]] std::string BuildConnectionBrowseDevicesJson(const std::vector<FileSystemMtpInternal::MtpConnectionBrowseDevice>& devices)
{
    std::string json(R"json({"version":1,"devices":[)json");
    for (size_t index = 0u; index < devices.size(); ++index)
    {
        if (index != 0u)
        {
            json.push_back(',');
        }

        const auto& device = devices[index];
        json.append(std::format(R"json({{"pnpId":"{}","friendlyName":"{}","devicePuid":"{}"}})json",
                                FileSystemMtpInternal::JsonEscapeUtf8(FileSystemMtpInternal::Utf8FromUtf16(device.pnpId)),
                                FileSystemMtpInternal::JsonEscapeUtf8(FileSystemMtpInternal::Utf8FromUtf16(device.friendlyName)),
                                FileSystemMtpInternal::JsonEscapeUtf8(FileSystemMtpInternal::Utf8FromUtf16(device.devicePuid))));
    }
    json.append("]}");
    return json;
}

[[nodiscard]] std::string BuildConnectionBrowseStoragesJson(const std::vector<FileSystemMtpInternal::MtpConnectionBrowseStorage>& storages)
{
    std::string json(R"json({"version":1,"storages":[)json");
    for (size_t index = 0u; index < storages.size(); ++index)
    {
        if (index != 0u)
        {
            json.push_back(',');
        }

        const auto& storage = storages[index];
        json.append(std::format(R"json({{"name":"{}","persistentId":"{}","objectId":"{}","initialPath":"{}"}})json",
                                FileSystemMtpInternal::JsonEscapeUtf8(FileSystemMtpInternal::Utf8FromUtf16(storage.name)),
                                FileSystemMtpInternal::JsonEscapeUtf8(FileSystemMtpInternal::Utf8FromUtf16(storage.persistentId)),
                                FileSystemMtpInternal::JsonEscapeUtf8(FileSystemMtpInternal::Utf8FromUtf16(storage.objectId)),
                                FileSystemMtpInternal::JsonEscapeUtf8(FileSystemMtpInternal::Utf8FromUtf16(storage.initialPath))));
    }
    json.append("]}");
    return json;
}

[[nodiscard]] HRESULT CopyJsonToCoTaskMem(std::string_view jsonUtf8, char** out) noexcept
{
    if (! out)
    {
        return E_POINTER;
    }

    *out = nullptr;
    if (jsonUtf8.size() >= (std::numeric_limits<size_t>::max)())
    {
        return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
    }

    const size_t bytes = jsonUtf8.size() + 1u;
    auto* mem          = static_cast<char*>(CoTaskMemAlloc(bytes));
    if (! mem)
    {
        return E_OUTOFMEMORY;
    }

    if (! jsonUtf8.empty())
    {
        std::memcpy(mem, jsonUtf8.data(), jsonUtf8.size());
    }
    mem[jsonUtf8.size()] = '\0';
    *out                 = mem;
    return S_OK;
}

#ifdef _DEBUG
std::mutex g_pickerFakeBackendMutex;
std::optional<std::string> g_pickerFakeBackendJson;

[[nodiscard]] bool TryGetPickerFakeBackendJson(std::string& outJsonUtf8)
{
    std::lock_guard lock(g_pickerFakeBackendMutex);
    if (! g_pickerFakeBackendJson.has_value())
    {
        return false;
    }

    outJsonUtf8 = g_pickerFakeBackendJson.value();
    return true;
}
#endif

[[nodiscard]] HRESULT BuildConnectionBrowseJson(const FactoryConnectionBrowseRequest& request, std::string& outJsonUtf8, uint64_t& outCount) noexcept
{
    outJsonUtf8.clear();
    outCount = 0u;

#ifdef _DEBUG
    std::string fakeBackendJson;
    const bool useFakeBackend = TryGetPickerFakeBackendJson(fakeBackendJson);
    std::unique_ptr<FileSystemMtpInternal::IMtpBackend> fakeBackend;
    if (useFakeBackend)
    {
        fakeBackend = FileSystemMtpInternal::CreateFakeMtpBackend(fakeBackendJson);
        if (! fakeBackend)
        {
            return E_OUTOFMEMORY;
        }
    }
#endif

    if (request.kind == FACTORY_CONNECTION_BROWSE_DEVICES)
    {
        std::vector<FileSystemMtpInternal::MtpConnectionBrowseDevice> devices;
#ifdef _DEBUG
        const HRESULT hr =
            fakeBackend ? FileSystemMtpInternal::EnumerateMtpConnectionBrowseDevicesFromBackend(*fakeBackend, devices)
                        : FileSystemMtpInternal::EnumerateMtpConnectionBrowseDevices(devices);
#else
        const HRESULT hr = FileSystemMtpInternal::EnumerateMtpConnectionBrowseDevices(devices);
#endif
        if (FAILED(hr))
        {
            return hr;
        }

        outCount    = devices.size();
        outJsonUtf8 = BuildConnectionBrowseDevicesJson(devices);
        return S_OK;
    }

    if (request.kind == FACTORY_CONNECTION_BROWSE_STORAGES)
    {
        const std::wstring_view parentDeviceId(request.parentDeviceId ? request.parentDeviceId : L"");
        if (parentDeviceId.empty())
        {
            return E_INVALIDARG;
        }

        std::vector<FileSystemMtpInternal::MtpConnectionBrowseStorage> storages;
#ifdef _DEBUG
        const HRESULT hr =
            fakeBackend ? FileSystemMtpInternal::EnumerateMtpConnectionBrowseStoragesFromBackend(*fakeBackend, parentDeviceId, storages)
                        : FileSystemMtpInternal::EnumerateMtpConnectionBrowseStorages(parentDeviceId, storages);
#else
        const HRESULT hr = FileSystemMtpInternal::EnumerateMtpConnectionBrowseStorages(parentDeviceId, storages);
#endif
        if (FAILED(hr))
        {
            return hr;
        }

        outCount    = storages.size();
        outJsonUtf8 = BuildConnectionBrowseStoragesJson(storages);
        return S_OK;
    }

    return E_INVALIDARG;
}

HRESULT CreateInstanceMtp(const FactoryOptions* /*factoryOptions*/, IHost* host, void** result) noexcept
{
    auto* instance = new (std::nothrow) FileSystemMtp(host);
    if (! instance)
    {
        return E_OUTOFMEMORY;
    }

    const HRESULT hr = instance->QueryInterface(__uuidof(IFileSystem), result);
    instance->Release();
    return hr;
}

const PluginFactoryEntry kEntries[] = {
    {&GetMetaDataMtp, &GetSchemaMtp, &CreateInstanceMtp},
};
} // namespace

extern "C" HRESULT __stdcall RedSalamanderEnumeratePlugins(REFIID riid, const PluginMetaData** metaData, unsigned int* count)
{
    return FactoryEnumeratePlugins<IFileSystem>(kEntries, riid, metaData, count);
}

extern "C" HRESULT __stdcall RedSalamanderCreate(REFIID riid, const FactoryOptions* factoryOptions, IHost* host, const wchar_t* pluginId, void** result)
{
    return FactoryCreate<IFileSystem>(kEntries, riid, factoryOptions, host, pluginId, result);
}

extern "C" HRESULT __stdcall RedSalamanderGetConfigurationSchema(REFIID riid, const wchar_t* pluginId, const char** schemaJsonUtf8)
{
    return FactoryGetConfigurationSchema<IFileSystem>(kEntries, riid, pluginId, schemaJsonUtf8);
}

extern "C" PLUGFACTORY_API HRESULT __stdcall RedSalamanderBrowseConnectionTargets(
    REFIID riid, const wchar_t* pluginId, const FactoryConnectionBrowseRequest* request, FactoryConnectionBrowseResult* result) noexcept
{
    if (! result)
    {
        return E_POINTER;
    }

    if (result->sizeBytes != sizeof(FactoryConnectionBrowseResult))
    {
        return E_INVALIDARG;
    }
    result->jsonUtf8 = nullptr;
    if (! request)
    {
        return E_POINTER;
    }
    if (request->sizeBytes != sizeof(FactoryConnectionBrowseRequest))
    {
        return E_INVALIDARG;
    }
    if (riid != __uuidof(IFileSystem))
    {
        return E_NOINTERFACE;
    }
    if (pluginId && pluginId[0] != L'\0' && ! OrdinalString::EqualsNoCase(pluginId, GetMetaDataMtp()->id))
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
    }

    const bool devicesRequest = request->kind == FACTORY_CONNECTION_BROWSE_DEVICES;
    Debug::Perf::Scope perf(devicesRequest ? L"mtp.connection_browse.devices_us" : L"mtp.connection_browse.storages_us");

    std::string jsonUtf8;
    uint64_t resultCount = 0u;
    const HRESULT browseHr = BuildConnectionBrowseJson(*request, jsonUtf8, resultCount);
    perf.SetValue0(resultCount);
    perf.SetHr(browseHr);
    Debug::Perf::EmitValue(devicesRequest ? L"mtp.connection_browse.devices" : L"mtp.connection_browse.storages", resultCount, browseHr);
    if (FAILED(browseHr))
    {
        return browseHr;
    }

    return CopyJsonToCoTaskMem(jsonUtf8, &result->jsonUtf8);
}

extern "C" PLUGFACTORY_API void __stdcall RedSalamanderPluginShutdown() noexcept
{
    ShutdownFileSystemMtpModule();
}

extern "C" PLUGFACTORY_API BOOL __stdcall RedSalamanderPluginCanUnloadNow() noexcept
{
    return CanUnloadFileSystemMtpModule() ? TRUE : FALSE;
}

extern "C" PLUGFACTORY_API BOOL __stdcall RedSalamanderPluginRetainModuleUntilProcessExit() noexcept
{
    return RetainFileSystemMtpModuleUntilProcessExit() ? TRUE : FALSE;
}

#ifdef _DEBUG
extern "C" PLUGFACTORY_API HRESULT __stdcall RedSalamanderMtpSetPickerFakeBackendForSelfTest(const char* fakeBackendJsonUtf8) noexcept
{
    std::lock_guard lock(g_pickerFakeBackendMutex);
    if (! fakeBackendJsonUtf8 || fakeBackendJsonUtf8[0] == '\0')
    {
        g_pickerFakeBackendJson.reset();
        return S_OK;
    }

    g_pickerFakeBackendJson = std::string(fakeBackendJsonUtf8);
    return S_OK;
}

extern "C" PLUGFACTORY_API HRESULT __stdcall RedSalamanderMtpCreateForSelfTest(
    REFIID riid, const FactoryOptions* /*factoryOptions*/, IHost* host, const char* fakeBackendJsonUtf8, void** result) noexcept
{
    if (! result)
    {
        return E_POINTER;
    }

    *result = nullptr;
    if (riid != __uuidof(IFileSystem))
    {
        return E_NOINTERFACE;
    }

    auto backend = FileSystemMtpInternal::CreateFakeMtpBackend(fakeBackendJsonUtf8 ? std::string_view(fakeBackendJsonUtf8) : std::string_view());
    if (! backend)
    {
        return E_OUTOFMEMORY;
    }

    auto* instance = new (std::nothrow) FileSystemMtp(host, std::move(backend));
    if (! instance)
    {
        return E_OUTOFMEMORY;
    }

    const HRESULT hr = instance->QueryInterface(riid, result);
    instance->Release();
    return hr;
}

extern "C" PLUGFACTORY_API HRESULT __stdcall RedSalamanderMtpCreateWpdCacheForSelfTest(
    REFIID riid, const FactoryOptions* /*factoryOptions*/, IHost* host, const char* backendJsonUtf8, void** result) noexcept
{
    if (! result)
    {
        return E_POINTER;
    }

    *result = nullptr;
    if (riid != __uuidof(IFileSystem))
    {
        return E_NOINTERFACE;
    }

    std::unique_ptr<FileSystemMtpInternal::IMtpBackend> backend;
    const HRESULT backendHr =
        FileSystemMtpInternal::CreateSelfTestWpdMtpBackend(backendJsonUtf8 ? std::string_view(backendJsonUtf8) : std::string_view(), backend);
    if (FAILED(backendHr))
    {
        return backendHr;
    }

    auto* instance = new (std::nothrow) FileSystemMtp(host, std::move(backend));
    if (! instance)
    {
        return E_OUTOFMEMORY;
    }

    const HRESULT hr = instance->QueryInterface(riid, result);
    instance->Release();
    return hr;
}

extern "C" PLUGFACTORY_API BOOL __stdcall RedSalamanderMtpRunJournalGenerationSelfTest() noexcept
{
    return FileSystemMtpInternal::RunOverwriteJournalGenerationSelfTest() ? TRUE : FALSE;
}

extern "C" PLUGFACTORY_API uint64_t __stdcall RedSalamanderMtpResetJournalProbeCountForSelfTest() noexcept
{
    return FileSystemMtpInternal::ResetOverwriteJournalProbeCountForSelfTest();
}

extern "C" PLUGFACTORY_API uint64_t __stdcall RedSalamanderMtpGetJournalProbeCountForSelfTest() noexcept
{
    return FileSystemMtpInternal::GetOverwriteJournalProbeCountForSelfTest();
}

extern "C" PLUGFACTORY_API HRESULT __stdcall RedSalamanderMtpNotifyJournalInjectedForSelfTest(const wchar_t* deviceIdentity) noexcept
{
    if (! deviceIdentity || deviceIdentity[0] == L'\0')
    {
        return E_INVALIDARG;
    }
    FileSystemMtpInternal::NotifyOverwriteJournalInjectedForSelfTest(deviceIdentity);
    return S_OK;
}
#endif
