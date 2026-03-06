#include "CompareDirectoriesEngine.SelfTest.h"

#ifdef _DEBUG

#include "Framework.h"

#include "ConnectionProfileUtils.h"

#include <algorithm>
#include <array>
#include <atomic>
#include <chrono>
#include <condition_variable>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <filesystem>
#include <format>
#include <limits>
#include <mutex>
#include <new>
#include <optional>
#include <span>
#include <string>
#include <string_view>
#include <thread>
#include <vector>

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027 28182)
#include <wil/com.h>
#include <wil/resource.h>
#pragma warning(pop)

#include "CompareDirectoriesEngine.h"
#include "ConnectionSecrets.h"
#include "CrashHandler.h"
#include "CrashQuarantine.h"
#include "FileSystemPluginManager.h"
#include "Helpers.h"
#include "HostServices.h"
#include "PlugInterfaces/Factory.h"
#include "PlugInterfaces/Host.h"
#include "PlugInterfaces/Informations.h"
#include "SessionState.h"
#include "SelfTestCommon.h"
#include "SettingsStore.h"
#include "WindowsHello.h"

extern Common::Settings::Settings g_settings;

namespace
{
constexpr std::wstring_view kBuiltinLocalFileSystemId = L"builtin/file-system";
constexpr std::wstring_view kBuiltinDummyFileSystemId = L"builtin/file-system-dummy";
constexpr std::wstring_view kBuiltin7zFileSystemId    = L"builtin/file-system-7z";
constexpr std::wstring_view kBuiltinFtpFileSystemId   = L"builtin/file-system-ftp";
constexpr std::wstring_view kBuiltinS3FileSystemId    = L"builtin/file-system-s3";

constexpr std::wstring_view kSelfTestEnvConnFtp = L"REDSALAMANDER_SELFTEST_CONN_FTP";
constexpr std::wstring_view kSelfTestEnvConnS3  = L"REDSALAMANDER_SELFTEST_CONN_S3";

constexpr std::wstring_view kSelfTestDefaultConnFtp = L"FileOpsSelfTest FTP";
constexpr std::wstring_view kSelfTestDefaultConnS3  = L"FileOpsSelfTest S3";

std::atomic<uint32_t> g_windowsHelloVerifierCalls{0};

HRESULT TestWindowsHelloVerifier(HWND /*ownerWindow*/, std::wstring_view /*message*/) noexcept
{
    g_windowsHelloVerifierCalls.fetch_add(1u, std::memory_order_relaxed);
    return S_OK;
}

void Trace(std::wstring_view message) noexcept
{
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::CompareDirectories, message);
    SelfTest::AppendSelfTestTrace(message);
}

void AppendCompareSelfTestTraceLine(std::wstring_view message) noexcept;

void AppendCaseResult(SelfTest::SelfTestSuiteResult& suite,
                      std::wstring_view name,
                      SelfTest::SelfTestCaseResult::Status status,
                      std::wstring_view reason = {}) noexcept
{
    std::wstring caseLine;
    caseLine.reserve(6 + name.size());
    caseLine.append(L"Case: ");
    caseLine.append(name);
    Trace(caseLine);

    SelfTest::SelfTestCaseResult result{};
    result.name       = std::wstring(name);
    result.status     = status;
    result.durationMs = 0;
    result.reason     = std::wstring(reason);

    suite.cases.push_back(std::move(result));
    switch (status)
    {
        case SelfTest::SelfTestCaseResult::Status::passed: ++suite.passed; break;
        case SelfTest::SelfTestCaseResult::Status::failed:
        {
            ++suite.failed;
            if (suite.failureMessage.empty())
            {
                suite.failureMessage = suite.cases.back().reason;
            }
            break;
        }
        case SelfTest::SelfTestCaseResult::Status::skipped: ++suite.skipped; break;
    }
}

[[nodiscard]] bool EqualsIgnoreCase(std::wstring_view a, std::wstring_view b) noexcept
{
    if (a.size() != b.size())
    {
        return false;
    }

    for (size_t i = 0; i < a.size(); ++i)
    {
        if (std::towlower(a[i]) != std::towlower(b[i]))
        {
            return false;
        }
    }

    return true;
}

[[nodiscard]] bool ContainsIgnoreCase(std::wstring_view text, std::wstring_view needle) noexcept
{
    if (needle.empty())
    {
        return true;
    }

    if (text.size() < needle.size())
    {
        return false;
    }

    for (size_t i = 0; i + needle.size() <= text.size(); ++i)
    {
        bool match = true;
        for (size_t j = 0; j < needle.size(); ++j)
        {
            if (std::towlower(text[i + j]) != std::towlower(needle[j]))
            {
                match = false;
                break;
            }
        }

        if (match)
        {
            return true;
        }
    }

    return false;
}

[[nodiscard]] std::wstring TrimWhitespace(std::wstring_view text) noexcept
{
    size_t start = 0;
    while (start < text.size())
    {
        const wchar_t ch = text[start];
        if (ch != L' ' && ch != L'\t' && ch != L'\r' && ch != L'\n')
        {
            break;
        }
        ++start;
    }

    size_t end = text.size();
    while (end > start)
    {
        const wchar_t ch = text[end - 1u];
        if (ch != L' ' && ch != L'\t' && ch != L'\r' && ch != L'\n')
        {
            break;
        }
        --end;
    }

    return std::wstring(text.substr(start, end - start));
}

[[nodiscard]] std::wstring GetEnvVarTrimmed(std::wstring_view name) noexcept
{
    if (name.empty())
    {
        return {};
    }

    std::wstring key(name);
    const DWORD required = GetEnvironmentVariableW(key.c_str(), nullptr, 0);
    if (required == 0)
    {
        return {};
    }

    std::wstring value;
    value.resize(required);
    const DWORD written = GetEnvironmentVariableW(key.c_str(), value.data(), required);
    if (written == 0 || written >= required)
    {
        return {};
    }

    value.resize(written);
    return TrimWhitespace(value);
}

void SecureClearAndFreeSecret(wil::unique_cotaskmem_string& secret) noexcept
{
    if (wchar_t* text = secret.get())
    {
        const size_t len = wcslen(text);
        if (len > 0)
        {
            SecureZeroMemory(text, len * sizeof(wchar_t));
        }
    }
    secret.reset();
}

struct PhaseCheckResult
{
    SelfTest::SelfTestCaseResult::Status status = SelfTest::SelfTestCaseResult::Status::skipped;
    std::wstring reason;
};

[[nodiscard]] PhaseCheckResult CheckRemoteConnectionSecret(std::wstring_view protocolLabel,
                                                           std::wstring_view envVarName,
                                                           std::wstring_view defaultProfileName,
                                                           std::wstring_view expectedPluginId) noexcept
{
    const std::wstring overrideName = GetEnvVarTrimmed(envVarName);
    const std::wstring profileName  = ! overrideName.empty() ? overrideName : std::wstring(defaultProfileName);

    const Common::Settings::ConnectionProfile* profile = ConnectionProfileUtils::FindConnectionProfileByName(&g_settings, profileName);
    if (! profile)
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped,
                .reason = std::format(L"{}: connection profile not found (set {} or create '{}').", protocolLabel, envVarName, defaultProfileName)};
    }

    if (profile->pluginId.empty() || ! EqualsIgnoreCase(profile->pluginId, expectedPluginId))
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped, .reason = std::format(L"{}: profile targets a different plugin.", protocolLabel)};
    }

    if (profile->authMode == Common::Settings::ConnectionAuthMode::Anonymous)
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped, .reason = std::format(L"{}: authMode=anonymous (no secret needed).", protocolLabel)};
    }

    if (profile->id.empty())
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped, .reason = std::format(L"{}: profile is missing a stable id.", protocolLabel)};
    }

    const bool bypassHello = g_settings.connections ? g_settings.connections->bypassWindowsHello : false;
    if (profile->requireWindowsHello && ! bypassHello)
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped,
                .reason = std::format(L"{}: requireWindowsHello=true (enable bypassWindowsHello or disable the profile flag for automation).", protocolLabel)};
    }

    if (! profile->savePassword)
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped,
                .reason = std::format(L"{}: savePassword=false (secret is not persisted).", protocolLabel)};
    }

    HostConnectionSecretKind kind = HOST_CONNECTION_SECRET_PASSWORD;
    if (profile->authMode == Common::Settings::ConnectionAuthMode::SshKey)
    {
        kind = HOST_CONNECTION_SECRET_SSH_KEY_PASSPHRASE;
    }

    wil::com_ptr<IHostConnections> hostConnections;
    const HRESULT hrQI = GetHostServices()->QueryInterface(IID_PPV_ARGS(hostConnections.addressof()));
    if (FAILED(hrQI) || ! hostConnections)
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::failed,
                .reason = std::format(L"{}: missing IHostConnections. hr=0x{:08X}", protocolLabel, static_cast<unsigned long>(hrQI))};
    }

    wil::unique_cotaskmem_string secret;
    const HRESULT hrSecret = hostConnections->GetConnectionSecret(profileName.c_str(), kind, nullptr, secret.put());
    if (hrSecret == HRESULT_FROM_WIN32(ERROR_NOT_FOUND))
    {
        SecureClearAndFreeSecret(secret);
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped, .reason = std::format(L"{}: secret not found.", protocolLabel)};
    }

    if (FAILED(hrSecret))
    {
        SecureClearAndFreeSecret(secret);
        return {.status = SelfTest::SelfTestCaseResult::Status::failed,
                .reason = std::format(L"{}: GetConnectionSecret failed. hr=0x{:08X}", protocolLabel, static_cast<unsigned long>(hrSecret))};
    }

    SecureClearAndFreeSecret(secret);
    return {.status = SelfTest::SelfTestCaseResult::Status::passed};
}

[[nodiscard]] std::wstring NormalizePluginPathForSelfTest(std::wstring_view rawPath) noexcept
{
    std::wstring path = TrimWhitespace(rawPath);
    for (wchar_t& ch : path)
    {
        if (ch == L'\\')
        {
            ch = L'/';
        }
    }

    while (path.size() > 1u && path.back() == L'/')
    {
        path.pop_back();
    }

    return path;
}

[[nodiscard]] PhaseCheckResult CheckRemoteConnectionSandbox(std::wstring_view protocolLabel,
                                                            std::wstring_view envVarName,
                                                            std::wstring_view defaultProfileName,
                                                            std::wstring_view expectedPluginId) noexcept
{
    const std::wstring overrideName = GetEnvVarTrimmed(envVarName);
    const std::wstring profileName  = ! overrideName.empty() ? overrideName : std::wstring(defaultProfileName);

    const Common::Settings::ConnectionProfile* profile = ConnectionProfileUtils::FindConnectionProfileByName(&g_settings, profileName);
    if (! profile)
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped,
                .reason = std::format(L"{}: connection profile not found (set {} or create '{}').", protocolLabel, envVarName, defaultProfileName)};
    }

    if (profile->pluginId.empty() || ! EqualsIgnoreCase(profile->pluginId, expectedPluginId))
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped, .reason = std::format(L"{}: profile targets a different plugin.", protocolLabel)};
    }

    const std::wstring initialPath = NormalizePluginPathForSelfTest(profile->initialPath);
    if (initialPath.empty() || initialPath[0] != L'/')
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped,
                .reason = std::format(L"{}: HARD REQUIREMENT: initialPath must be an absolute plugin path (starting with '/').", protocolLabel)};
    }

    if (initialPath == L"/")
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped,
                .reason = std::format(L"{}: HARD REQUIREMENT: initialPath must point to a dedicated selftest folder/prefix (not '/').", protocolLabel)};
    }

    if (ContainsIgnoreCase(initialPath, L"/@conn:"))
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped,
                .reason = std::format(L"{}: HARD REQUIREMENT: initialPath must not include the host-reserved '/@conn:' prefix.", protocolLabel)};
    }

    size_t segmentCount = 0;
    std::wstring_view soleSegment;
    for (size_t i = 0; i < initialPath.size();)
    {
        while (i < initialPath.size() && initialPath[i] == L'/')
        {
            ++i;
        }

        const size_t start = i;
        while (i < initialPath.size() && initialPath[i] != L'/')
        {
            ++i;
        }

        if (i == start)
        {
            continue;
        }

        const std::wstring_view segment = std::wstring_view(initialPath).substr(start, i - start);
        if (segment == L"." || segment == L"..")
        {
            return {.status = SelfTest::SelfTestCaseResult::Status::skipped,
                    .reason = std::format(L"{}: HARD REQUIREMENT: initialPath must not contain '.' or '..' segments.", protocolLabel)};
        }

        ++segmentCount;
        if (segmentCount == 1u)
        {
            soleSegment = segment;
        }
    }

    const bool isS3 = EqualsIgnoreCase(expectedPluginId, kBuiltinS3FileSystemId);
    if (isS3 && segmentCount < 2)
    {
        const bool bucketRootIsSelfTestOnly = segmentCount == 1u && ContainsIgnoreCase(soleSegment, L"selftest");
        if (bucketRootIsSelfTestOnly)
        {
            return {.status = SelfTest::SelfTestCaseResult::Status::passed};
        }

        return {.status = SelfTest::SelfTestCaseResult::Status::skipped,
                .reason = std::format(
                    L"{}: HARD REQUIREMENT: initialPath must include a bucket and a dedicated selftest prefix (e.g. '/bucket/red-salamander-selftest'), "
                    L"or use a dedicated selftest bucket root whose bucket name includes 'selftest'.",
                    protocolLabel)};
    }

    if (! ContainsIgnoreCase(initialPath, L"selftest"))
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped,
                .reason =
                    std::format(L"{}: HARD REQUIREMENT: initialPath must include 'selftest' (case-insensitive) to prove it is test-only.", protocolLabel)};
    }

    return {.status = SelfTest::SelfTestCaseResult::Status::passed};
}

using CreateFactoryFunc   = HRESULT(__stdcall*)(REFIID, const FactoryOptions*, IHost*, void**);
using CreateFactoryExFunc = HRESULT(__stdcall*)(REFIID, const FactoryOptions*, IHost*, const wchar_t*, void**);

struct CreatedFileSystemInstance
{
    wil::unique_hmodule module;
    wil::com_ptr<IFileSystem> fileSystem;

    CreatedFileSystemInstance()                                            = default;
    CreatedFileSystemInstance(const CreatedFileSystemInstance&)            = delete;
    CreatedFileSystemInstance& operator=(const CreatedFileSystemInstance&) = delete;
    CreatedFileSystemInstance(CreatedFileSystemInstance&&)                 = default;
    CreatedFileSystemInstance& operator=(CreatedFileSystemInstance&&)      = default;
};

[[nodiscard]] const FileSystemPluginManager::PluginEntry* FindFileSystemPluginById(std::wstring_view pluginId) noexcept
{
    if (pluginId.empty())
    {
        return nullptr;
    }

    const auto& plugins = FileSystemPluginManager::GetInstance().GetPlugins();
    for (const FileSystemPluginManager::PluginEntry& entry : plugins)
    {
        if (entry.id.empty())
        {
            continue;
        }

        if (CompareStringOrdinal(entry.id.c_str(), -1, pluginId.data(), static_cast<int>(pluginId.size()), TRUE) == CSTR_EQUAL)
        {
            return &entry;
        }
    }

    return nullptr;
}

[[nodiscard]] HRESULT TryCreateFileSystemInstance(std::wstring_view pluginId, std::wstring_view instanceContext, CreatedFileSystemInstance& out) noexcept
{
    out = {};

    const FileSystemPluginManager::PluginEntry* entry = FindFileSystemPluginById(pluginId);
    if (! entry || entry->id.empty() || entry->disabled || ! entry->loadable || entry->path.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
    }

    wil::unique_hmodule module(LoadLibraryExW(entry->path.c_str(), nullptr, LOAD_WITH_ALTERED_SEARCH_PATH));
    if (! module)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

#pragma warning(push)
#pragma warning(disable : 4191) // unsafe conversion from FARPROC
    const auto createFactory   = reinterpret_cast<CreateFactoryFunc>(GetProcAddress(module.get(), "RedSalamanderCreate"));
    const auto createFactoryEx = reinterpret_cast<CreateFactoryExFunc>(GetProcAddress(module.get(), "RedSalamanderCreateEx"));
#pragma warning(pop)
    if (! createFactory)
    {
        return HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND);
    }

    FactoryOptions options{};
    options.debugLevel = DEBUG_LEVEL_NONE;

    wil::com_ptr<IFileSystem> fileSystem;
    HRESULT createHr = E_FAIL;
    if (entry->factoryPluginId.empty())
    {
        createHr = createFactory(__uuidof(IFileSystem), &options, GetHostServices(), fileSystem.put_void());
    }
    else if (createFactoryEx)
    {
        createHr = createFactoryEx(__uuidof(IFileSystem), &options, GetHostServices(), entry->factoryPluginId.c_str(), fileSystem.put_void());
    }
    else
    {
        return HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND);
    }

    if (FAILED(createHr) || ! fileSystem)
    {
        return createHr;
    }

    wil::com_ptr<IInformations> informations;
    const HRESULT qiInfos = fileSystem->QueryInterface(__uuidof(IInformations), informations.put_void());
    if (FAILED(qiInfos) || ! informations)
    {
        return qiInfos;
    }

    if (entry->informations)
    {
        const char* configuration = nullptr;
        static_cast<void>(entry->informations->GetConfiguration(&configuration));
        if (configuration && configuration[0] != '\0')
        {
            static_cast<void>(informations->SetConfiguration(configuration));
        }
    }

    if (! instanceContext.empty())
    {
        wil::com_ptr<IFileSystemInitialize> initializer;
        const HRESULT qiInit = fileSystem->QueryInterface(__uuidof(IFileSystemInitialize), initializer.put_void());
        if (FAILED(qiInit) || ! initializer)
        {
            return qiInit;
        }

        std::wstring contextText(instanceContext);
        const HRESULT initHr = initializer->Initialize(contextText.c_str(), nullptr);
        if (FAILED(initHr))
        {
            return initHr;
        }
    }

    out.module     = std::move(module);
    out.fileSystem = std::move(fileSystem);
    return S_OK;
}

[[nodiscard]] std::wstring MakeGuidText() noexcept
{
    GUID guid{};
    if (FAILED(::CoCreateGuid(&guid)))
    {
        return {};
    }

    wchar_t buffer[64]{};
    if (::StringFromGUID2(guid, buffer, static_cast<int>(std::size(buffer))) <= 0)
    {
        return {};
    }

    std::wstring text(buffer);
    text.erase(std::remove_if(text.begin(), text.end(), [](wchar_t ch) noexcept { return ch == L'{' || ch == L'}'; }), text.end());
    return text;
}

[[nodiscard]] bool SetFileLastWriteTime(const std::filesystem::path& path, const FILETIME& lastWriteTime) noexcept
{
    wil::unique_handle file(::CreateFileW(
        path.c_str(), FILE_WRITE_ATTRIBUTES, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr));
    if (! file)
    {
        return false;
    }

    return ::SetFileTime(file.get(), nullptr, nullptr, &lastWriteTime) != 0;
}

[[nodiscard]] wil::com_ptr<IFileSystem> GetLocalFileSystem() noexcept
{
    return SelfTest::GetFileSystem(kBuiltinLocalFileSystemId);
}

class ShortReadFileReader final : public IFileReader
{
public:
    ShortReadFileReader(wil::com_ptr<IFileReader> inner, unsigned long maxBytesPerRead, DWORD delayMs) noexcept
        : _inner(std::move(inner)),
          _maxBytesPerRead(std::max<unsigned long>(maxBytesPerRead, 1u)),
          _delayMs(delayMs)
    {
    }

    ShortReadFileReader(const ShortReadFileReader&)            = delete;
    ShortReadFileReader(ShortReadFileReader&&)                 = delete;
    ShortReadFileReader& operator=(const ShortReadFileReader&) = delete;
    ShortReadFileReader& operator=(ShortReadFileReader&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (ppvObject == nullptr)
        {
            return E_POINTER;
        }

        if (riid == __uuidof(IUnknown) || riid == __uuidof(IFileReader))
        {
            *ppvObject = static_cast<IFileReader*>(this);
            AddRef();
            return S_OK;
        }

        *ppvObject = nullptr;
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return _refCount.fetch_add(1, std::memory_order_relaxed) + 1;
    }

    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        const ULONG current = _refCount.fetch_sub(1, std::memory_order_acq_rel) - 1;
        if (current == 0)
        {
            delete this;
        }
        return current;
    }

    HRESULT STDMETHODCALLTYPE GetSize(uint64_t* sizeBytes) noexcept override
    {
        if (! _inner)
        {
            return E_FAIL;
        }
        return _inner->GetSize(sizeBytes);
    }

    HRESULT STDMETHODCALLTYPE Seek(__int64 offset, unsigned long origin, uint64_t* newPosition) noexcept override
    {
        if (! _inner)
        {
            return E_FAIL;
        }
        return _inner->Seek(offset, origin, newPosition);
    }

    HRESULT STDMETHODCALLTYPE Read(void* buffer, unsigned long bytesToRead, unsigned long* bytesRead) noexcept override
    {
        if (bytesRead == nullptr)
        {
            return E_POINTER;
        }

        *bytesRead = 0;

        if (bytesToRead == 0)
        {
            return S_OK;
        }

        if (buffer == nullptr)
        {
            return E_POINTER;
        }

        if (! _inner)
        {
            return E_FAIL;
        }

        const unsigned long capped = std::min(bytesToRead, _maxBytesPerRead);
        if (_delayMs != 0)
        {
            ::Sleep(_delayMs);
        }
        return _inner->Read(buffer, capped, bytesRead);
    }

private:
    ~ShortReadFileReader() = default;

    std::atomic_ulong _refCount{1};
    wil::com_ptr<IFileReader> _inner;
    unsigned long _maxBytesPerRead = 1;
    DWORD _delayMs                 = 0;
};

// ShortReadFileSystem wraps a real IFileSystem/IFileSystemIO and limits every Read()
// call to at most maxBytesPerRead bytes.  Used as a regression guard to verify that
// the content-comparison engine handles partial reads correctly (i.e. it never assumes
// a single Read() returns the full file).
class ShortReadFileSystem final : public IFileSystem, public IFileSystemIO
{
public:
    ShortReadFileSystem(wil::com_ptr<IFileSystem> base, std::filesystem::path shortReadRoot, unsigned long maxBytesPerRead, DWORD delayMs) noexcept
        : _base(std::move(base)),
          _shortReadRoot(std::move(shortReadRoot)),
          _maxBytesPerRead(std::max<unsigned long>(maxBytesPerRead, 1u)),
          _delayMs(delayMs)
    {
        if (_base)
        {
            static_cast<void>(_base->QueryInterface(__uuidof(IFileSystemIO), _baseIo.put_void()));
        }
    }

    ShortReadFileSystem(const ShortReadFileSystem&)            = delete;
    ShortReadFileSystem(ShortReadFileSystem&&)                 = delete;
    ShortReadFileSystem& operator=(const ShortReadFileSystem&) = delete;
    ShortReadFileSystem& operator=(ShortReadFileSystem&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (ppvObject == nullptr)
        {
            return E_POINTER;
        }

        if (riid == __uuidof(IUnknown) || riid == __uuidof(IFileSystem))
        {
            *ppvObject = static_cast<IFileSystem*>(this);
            AddRef();
            return S_OK;
        }

        if (riid == __uuidof(IFileSystemIO))
        {
            *ppvObject = static_cast<IFileSystemIO*>(this);
            AddRef();
            return S_OK;
        }

        if (_base)
        {
            return _base->QueryInterface(riid, ppvObject);
        }

        *ppvObject = nullptr;
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return _refCount.fetch_add(1, std::memory_order_relaxed) + 1;
    }

    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        const ULONG current = _refCount.fetch_sub(1, std::memory_order_acq_rel) - 1;
        if (current == 0)
        {
            delete this;
        }
        return current;
    }

    HRESULT STDMETHODCALLTYPE ReadDirectoryInfo(const wchar_t* path, IFilesInformation** ppFilesInformation) noexcept override
    {
        if (! _base)
        {
            return E_POINTER;
        }
        return _base->ReadDirectoryInfo(path, ppFilesInformation);
    }

    HRESULT STDMETHODCALLTYPE CopyItem(const wchar_t* sourcePath,
                                       const wchar_t* destinationPath,
                                       FileSystemFlags flags,
                                       const FileSystemOptions* options,
                                       IFileSystemCallback* callback,
                                       void* cookie) noexcept override
    {
        if (! _base)
        {
            return E_POINTER;
        }
        return _base->CopyItem(sourcePath, destinationPath, flags, options, callback, cookie);
    }

    HRESULT STDMETHODCALLTYPE MoveItem(const wchar_t* sourcePath,
                                       const wchar_t* destinationPath,
                                       FileSystemFlags flags,
                                       const FileSystemOptions* options,
                                       IFileSystemCallback* callback,
                                       void* cookie) noexcept override
    {
        if (! _base)
        {
            return E_POINTER;
        }
        return _base->MoveItem(sourcePath, destinationPath, flags, options, callback, cookie);
    }

    HRESULT STDMETHODCALLTYPE
    DeleteItem(const wchar_t* path, FileSystemFlags flags, const FileSystemOptions* options, IFileSystemCallback* callback, void* cookie) noexcept override
    {
        if (! _base)
        {
            return E_POINTER;
        }
        return _base->DeleteItem(path, flags, options, callback, cookie);
    }

    HRESULT STDMETHODCALLTYPE RenameItem(const wchar_t* sourcePath,
                                         const wchar_t* destinationPath,
                                         FileSystemFlags flags,
                                         const FileSystemOptions* options,
                                         IFileSystemCallback* callback,
                                         void* cookie) noexcept override
    {
        if (! _base)
        {
            return E_POINTER;
        }
        return _base->RenameItem(sourcePath, destinationPath, flags, options, callback, cookie);
    }

    HRESULT STDMETHODCALLTYPE CopyItems(const wchar_t* const* sourcePaths,
                                        unsigned long count,
                                        const wchar_t* destinationFolder,
                                        FileSystemFlags flags,
                                        const FileSystemOptions* options,
                                        IFileSystemCallback* callback,
                                        void* cookie) noexcept override
    {
        if (! _base)
        {
            return E_POINTER;
        }
        return _base->CopyItems(sourcePaths, count, destinationFolder, flags, options, callback, cookie);
    }

    HRESULT STDMETHODCALLTYPE MoveItems(const wchar_t* const* sourcePaths,
                                        unsigned long count,
                                        const wchar_t* destinationFolder,
                                        FileSystemFlags flags,
                                        const FileSystemOptions* options,
                                        IFileSystemCallback* callback,
                                        void* cookie) noexcept override
    {
        if (! _base)
        {
            return E_POINTER;
        }
        return _base->MoveItems(sourcePaths, count, destinationFolder, flags, options, callback, cookie);
    }

    HRESULT STDMETHODCALLTYPE DeleteItems(const wchar_t* const* paths,
                                          unsigned long count,
                                          FileSystemFlags flags,
                                          const FileSystemOptions* options,
                                          IFileSystemCallback* callback,
                                          void* cookie) noexcept override
    {
        if (! _base)
        {
            return E_POINTER;
        }
        return _base->DeleteItems(paths, count, flags, options, callback, cookie);
    }

    HRESULT STDMETHODCALLTYPE RenameItems(const FileSystemRenamePair* items,
                                          unsigned long count,
                                          FileSystemFlags flags,
                                          const FileSystemOptions* options,
                                          IFileSystemCallback* callback,
                                          void* cookie) noexcept override
    {
        if (! _base)
        {
            return E_POINTER;
        }
        return _base->RenameItems(items, count, flags, options, callback, cookie);
    }

    HRESULT STDMETHODCALLTYPE GetCapabilities(const char** jsonUtf8) noexcept override
    {
        if (! _base)
        {
            return E_POINTER;
        }
        return _base->GetCapabilities(jsonUtf8);
    }

    HRESULT STDMETHODCALLTYPE GetAttributes(const wchar_t* path, unsigned long* fileAttributes) noexcept override
    {
        if (! _baseIo)
        {
            return E_POINTER;
        }
        return _baseIo->GetAttributes(path, fileAttributes);
    }

    HRESULT STDMETHODCALLTYPE CreateFileReader(const wchar_t* path, IFileReader** reader) noexcept override
    {
        if (reader == nullptr)
        {
            return E_POINTER;
        }
        *reader = nullptr;

        if (! _baseIo)
        {
            return E_POINTER;
        }

        wil::com_ptr<IFileReader> inner;
        const HRESULT hr = _baseIo->CreateFileReader(path, inner.put());
        if (FAILED(hr) || ! inner)
        {
            return FAILED(hr) ? hr : E_FAIL;
        }

        const std::wstring_view pathText(path ? path : L"");
        const std::wstring rootText = _shortReadRoot.wstring();
        const bool shouldShortRead  = ! rootText.empty() && OrdinalString::StartsWithNoCase(pathText, rootText);
        if (! shouldShortRead)
        {
            *reader = inner.detach();
            return S_OK;
        }

        auto* wrapper = new (std::nothrow) ShortReadFileReader(std::move(inner), _maxBytesPerRead, _delayMs);
        if (! wrapper)
        {
            return E_OUTOFMEMORY;
        }

        *reader = wrapper;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE CreateFileWriter(const wchar_t* path, FileSystemFlags flags, IFileWriter** writer) noexcept override
    {
        if (! _baseIo)
        {
            return E_POINTER;
        }
        return _baseIo->CreateFileWriter(path, flags, writer);
    }

    HRESULT STDMETHODCALLTYPE GetFileBasicInformation(const wchar_t* path, FileSystemBasicInformation* info) noexcept override
    {
        if (! _baseIo)
        {
            return E_POINTER;
        }
        return _baseIo->GetFileBasicInformation(path, info);
    }

    HRESULT STDMETHODCALLTYPE SetFileBasicInformation(const wchar_t* path, const FileSystemBasicInformation* info) noexcept override
    {
        if (! _baseIo)
        {
            return E_POINTER;
        }
        return _baseIo->SetFileBasicInformation(path, info);
    }

    HRESULT STDMETHODCALLTYPE GetItemProperties(const wchar_t* path, const char** jsonUtf8) noexcept override
    {
        if (! _baseIo)
        {
            return E_POINTER;
        }
        return _baseIo->GetItemProperties(path, jsonUtf8);
    }

private:
    ~ShortReadFileSystem() = default;

    std::atomic_ulong _refCount{1};
    wil::com_ptr<IFileSystem> _base;
    wil::com_ptr<IFileSystemIO> _baseIo;
    std::filesystem::path _shortReadRoot;
    unsigned long _maxBytesPerRead = 1;
    DWORD _delayMs                 = 0;
};

[[nodiscard]] wil::com_ptr<IFileSystem> CreateShortReadFileSystem(const wil::com_ptr<IFileSystem>& base,
                                                                  const std::filesystem::path& shortReadRoot,
                                                                  unsigned long maxBytesPerRead,
                                                                  DWORD delayMs) noexcept
{
    wil::com_ptr<IFileSystem> wrapped;
    auto* wrapper = new (std::nothrow) ShortReadFileSystem(base, shortReadRoot, maxBytesPerRead, delayMs);
    if (! wrapper)
    {
        return {};
    }
    wrapped.attach(wrapper);
    return wrapped;
}

struct ReadDirectoryTestBehavior
{
    std::filesystem::path targetPath;
    DWORD delayMs           = 0;
    bool returnMalformed    = false;
    unsigned long fakeCount = 1;
};

[[nodiscard]] std::wstring NormalizePathForNoCaseCompare(std::filesystem::path value) noexcept
{
    value = value.lexically_normal();
    while (! value.empty() && ! value.has_filename() && value != value.root_path())
    {
        value = value.parent_path();
    }

    std::wstring text = value.wstring();
    std::replace(text.begin(), text.end(), L'/', L'\\');

    if (text.rfind(L"\\\\?\\UNC\\", 0) == 0)
    {
        text.erase(0, 8);
        text.insert(0, L"\\\\");
    }
    else if (text.rfind(L"\\\\?\\", 0) == 0)
    {
        text.erase(0, 4);
    }

    if (! text.empty())
    {
        ::CharLowerBuffW(text.data(), static_cast<DWORD>(text.size()));
    }

    return text;
}

class RawFilesInformation final : public IFilesInformation
{
public:
    RawFilesInformation(std::vector<unsigned char> buffer, unsigned long count) noexcept
        : _buffer(std::move(buffer)),
          _count(count)
    {
    }

    RawFilesInformation(const RawFilesInformation&)            = delete;
    RawFilesInformation(RawFilesInformation&&)                 = delete;
    RawFilesInformation& operator=(const RawFilesInformation&) = delete;
    RawFilesInformation& operator=(RawFilesInformation&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (ppvObject == nullptr)
        {
            return E_POINTER;
        }

        if (riid == __uuidof(IUnknown) || riid == __uuidof(IFilesInformation))
        {
            *ppvObject = static_cast<IFilesInformation*>(this);
            AddRef();
            return S_OK;
        }

        *ppvObject = nullptr;
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return _refCount.fetch_add(1, std::memory_order_relaxed) + 1;
    }

    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        const ULONG current = _refCount.fetch_sub(1, std::memory_order_acq_rel) - 1;
        if (current == 0)
        {
            delete this;
        }
        return current;
    }

    HRESULT STDMETHODCALLTYPE GetBuffer(FileInfo** ppFileInfo) noexcept override
    {
        if (ppFileInfo == nullptr)
        {
            return E_POINTER;
        }

        if (_buffer.empty())
        {
            *ppFileInfo = nullptr;
            return S_OK;
        }

        *ppFileInfo = reinterpret_cast<FileInfo*>(_buffer.data());
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetBufferSize(unsigned long* pSize) noexcept override
    {
        if (pSize == nullptr)
        {
            return E_POINTER;
        }

        *pSize = static_cast<unsigned long>(_buffer.size());
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetAllocatedSize(unsigned long* pSize) noexcept override
    {
        if (pSize == nullptr)
        {
            return E_POINTER;
        }

        *pSize = static_cast<unsigned long>(_buffer.capacity());
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetCount(unsigned long* pCount) noexcept override
    {
        if (pCount == nullptr)
        {
            return E_POINTER;
        }

        *pCount = _count;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Get(unsigned long index, FileInfo** ppEntry) noexcept override
    {
        if (ppEntry == nullptr)
        {
            return E_POINTER;
        }

        *ppEntry = nullptr;
        if (index != 0 || _buffer.empty())
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_INDEX);
        }

        *ppEntry = reinterpret_cast<FileInfo*>(_buffer.data());
        return S_OK;
    }

private:
    ~RawFilesInformation() = default;

    std::atomic_ulong _refCount{1};
    std::vector<unsigned char> _buffer;
    unsigned long _count = 0;
};

[[nodiscard]] std::vector<unsigned char> BuildMalformedDirectoryBuffer(std::wstring_view name) noexcept
{
    const size_t nameBytes = name.size() * sizeof(wchar_t);
    const size_t totalSize = offsetof(FileInfo, FileName) + nameBytes;

    std::vector<unsigned char> buffer(totalSize);
    if (buffer.empty())
    {
        return buffer;
    }

    auto* entry = reinterpret_cast<FileInfo*>(buffer.data());
    entry->NextEntryOffset = static_cast<unsigned long>(sizeof(FileInfo) + 16u); // Intentionally larger than the backing buffer.
    entry->FileIndex       = 1;
    entry->CreationTime    = 0;
    entry->LastAccessTime  = 0;
    entry->LastWriteTime   = 0;
    entry->ChangeTime      = 0;
    entry->EndOfFile       = 1;
    entry->AllocationSize  = 1;
    entry->FileAttributes  = FILE_ATTRIBUTE_ARCHIVE;
    entry->FileNameSize    = static_cast<unsigned long>(nameBytes);
    entry->EaSize          = 0;

    if (nameBytes != 0)
    {
        std::memcpy(entry->FileName, name.data(), nameBytes);
    }

    return buffer;
}

class ReadDirectoryBehaviorFileSystem final : public IFileSystem
{
public:
    ReadDirectoryBehaviorFileSystem(wil::com_ptr<IFileSystem> base, ReadDirectoryTestBehavior behavior) noexcept
        : _base(std::move(base)),
          _behavior(std::move(behavior))
    {
        _normalizedTargetPath = NormalizePathForNoCaseCompare(_behavior.targetPath);
    }

    ReadDirectoryBehaviorFileSystem(const ReadDirectoryBehaviorFileSystem&)            = delete;
    ReadDirectoryBehaviorFileSystem(ReadDirectoryBehaviorFileSystem&&)                 = delete;
    ReadDirectoryBehaviorFileSystem& operator=(const ReadDirectoryBehaviorFileSystem&) = delete;
    ReadDirectoryBehaviorFileSystem& operator=(ReadDirectoryBehaviorFileSystem&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (ppvObject == nullptr)
        {
            return E_POINTER;
        }

        if (riid == __uuidof(IUnknown) || riid == __uuidof(IFileSystem))
        {
            *ppvObject = static_cast<IFileSystem*>(this);
            AddRef();
            return S_OK;
        }

        if (_base)
        {
            return _base->QueryInterface(riid, ppvObject);
        }

        *ppvObject = nullptr;
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return _refCount.fetch_add(1, std::memory_order_relaxed) + 1;
    }

    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        const ULONG current = _refCount.fetch_sub(1, std::memory_order_acq_rel) - 1;
        if (current == 0)
        {
            delete this;
        }
        return current;
    }

    HRESULT STDMETHODCALLTYPE ReadDirectoryInfo(const wchar_t* path, IFilesInformation** ppFilesInformation) noexcept override
    {
        if (ppFilesInformation == nullptr)
        {
            return E_POINTER;
        }

        *ppFilesInformation = nullptr;
        if (! _base)
        {
            return E_POINTER;
        }

        const std::filesystem::path requestedPath(path ? path : L"");
        const std::wstring requested = NormalizePathForNoCaseCompare(requestedPath);
        const bool intercept         = _normalizedTargetPath.empty() || requested == _normalizedTargetPath;
        if (! intercept)
        {
            return _base->ReadDirectoryInfo(path, ppFilesInformation);
        }

        if (_behavior.delayMs != 0u)
        {
            ::Sleep(_behavior.delayMs);
        }

        if (! _behavior.returnMalformed)
        {
            return _base->ReadDirectoryInfo(path, ppFilesInformation);
        }

        auto* info = new (std::nothrow) RawFilesInformation(BuildMalformedDirectoryBuffer(L"corrupt-entry.txt"), _behavior.fakeCount);
        if (! info)
        {
            return E_OUTOFMEMORY;
        }

        *ppFilesInformation = info;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE CopyItem(const wchar_t* sourcePath,
                                       const wchar_t* destinationPath,
                                       FileSystemFlags flags,
                                       const FileSystemOptions* options,
                                       IFileSystemCallback* callback,
                                       void* cookie) noexcept override
    {
        return _base ? _base->CopyItem(sourcePath, destinationPath, flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE MoveItem(const wchar_t* sourcePath,
                                       const wchar_t* destinationPath,
                                       FileSystemFlags flags,
                                       const FileSystemOptions* options,
                                       IFileSystemCallback* callback,
                                       void* cookie) noexcept override
    {
        return _base ? _base->MoveItem(sourcePath, destinationPath, flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE
    DeleteItem(const wchar_t* itemPath, FileSystemFlags flags, const FileSystemOptions* options, IFileSystemCallback* callback, void* cookie) noexcept override
    {
        return _base ? _base->DeleteItem(itemPath, flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE RenameItem(const wchar_t* sourcePath,
                                         const wchar_t* destinationPath,
                                         FileSystemFlags flags,
                                         const FileSystemOptions* options,
                                         IFileSystemCallback* callback,
                                         void* cookie) noexcept override
    {
        return _base ? _base->RenameItem(sourcePath, destinationPath, flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE CopyItems(const wchar_t* const* sourcePaths,
                                        unsigned long count,
                                        const wchar_t* destinationFolder,
                                        FileSystemFlags flags,
                                        const FileSystemOptions* options,
                                        IFileSystemCallback* callback,
                                        void* cookie) noexcept override
    {
        return _base ? _base->CopyItems(sourcePaths, count, destinationFolder, flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE MoveItems(const wchar_t* const* sourcePaths,
                                        unsigned long count,
                                        const wchar_t* destinationFolder,
                                        FileSystemFlags flags,
                                        const FileSystemOptions* options,
                                        IFileSystemCallback* callback,
                                        void* cookie) noexcept override
    {
        return _base ? _base->MoveItems(sourcePaths, count, destinationFolder, flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE DeleteItems(const wchar_t* const* paths,
                                          unsigned long count,
                                          FileSystemFlags flags,
                                          const FileSystemOptions* options,
                                          IFileSystemCallback* callback,
                                          void* cookie) noexcept override
    {
        return _base ? _base->DeleteItems(paths, count, flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE RenameItems(const FileSystemRenamePair* items,
                                          unsigned long count,
                                          FileSystemFlags flags,
                                          const FileSystemOptions* options,
                                          IFileSystemCallback* callback,
                                          void* cookie) noexcept override
    {
        return _base ? _base->RenameItems(items, count, flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE GetCapabilities(const char** jsonUtf8) noexcept override
    {
        return _base ? _base->GetCapabilities(jsonUtf8) : E_POINTER;
    }

private:
    ~ReadDirectoryBehaviorFileSystem() = default;

    std::atomic_ulong _refCount{1};
    wil::com_ptr<IFileSystem> _base;
    ReadDirectoryTestBehavior _behavior;
    std::wstring _normalizedTargetPath;
};

[[nodiscard]] wil::com_ptr<IFileSystem> CreateReadDirectoryBehaviorFileSystem(const wil::com_ptr<IFileSystem>& base,
                                                                               ReadDirectoryTestBehavior behavior) noexcept
{
    wil::com_ptr<IFileSystem> wrapped;
    auto* wrapper = new (std::nothrow) ReadDirectoryBehaviorFileSystem(base, std::move(behavior));
    if (! wrapper)
    {
        return {};
    }

    wrapped.attach(wrapper);
    return wrapped;
}

class PluginPathMappedRootFileSystem final : public IFileSystem, public IFileSystemIO
{
public:
    PluginPathMappedRootFileSystem(wil::com_ptr<IFileSystem> base, std::filesystem::path windowsRoot) noexcept
        : _base(std::move(base)),
          _windowsRoot(std::move(windowsRoot))
    {
        if (_base)
        {
            static_cast<void>(_base->QueryInterface(__uuidof(IFileSystemIO), _baseIo.put_void()));
        }
    }

    PluginPathMappedRootFileSystem(const PluginPathMappedRootFileSystem&)            = delete;
    PluginPathMappedRootFileSystem(PluginPathMappedRootFileSystem&&)                 = delete;
    PluginPathMappedRootFileSystem& operator=(const PluginPathMappedRootFileSystem&) = delete;
    PluginPathMappedRootFileSystem& operator=(PluginPathMappedRootFileSystem&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (ppvObject == nullptr)
        {
            return E_POINTER;
        }

        if (riid == __uuidof(IUnknown) || riid == __uuidof(IFileSystem))
        {
            *ppvObject = static_cast<IFileSystem*>(this);
            AddRef();
            return S_OK;
        }

        // Do not expose IInformations: DirectoryInfoCache uses it to decide path semantics and would
        // otherwise treat this wrapper as the file plugin, breaking plugin-path mapping.
        if (riid == __uuidof(IInformations))
        {
            *ppvObject = nullptr;
            return E_NOINTERFACE;
        }

        if (riid == __uuidof(IFileSystemIO))
        {
            if (! _baseIo)
            {
                *ppvObject = nullptr;
                return E_NOINTERFACE;
            }

            *ppvObject = static_cast<IFileSystemIO*>(this);
            AddRef();
            return S_OK;
        }

        if (_base)
        {
            return _base->QueryInterface(riid, ppvObject);
        }

        *ppvObject = nullptr;
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return _refCount.fetch_add(1, std::memory_order_relaxed) + 1;
    }

    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        const ULONG current = _refCount.fetch_sub(1, std::memory_order_acq_rel) - 1;
        if (current == 0)
        {
            delete this;
        }
        return current;
    }

    // IFileSystem
    HRESULT STDMETHODCALLTYPE ReadDirectoryInfo(const wchar_t* path, IFilesInformation** ppFilesInformation) noexcept override
    {
        if (! _base)
        {
            return E_POINTER;
        }
        return _base->ReadDirectoryInfo(MapPath(path).c_str(), ppFilesInformation);
    }

    HRESULT STDMETHODCALLTYPE CopyItem(const wchar_t* sourcePath,
                                       const wchar_t* destinationPath,
                                       FileSystemFlags flags,
                                       const FileSystemOptions* options,
                                       IFileSystemCallback* callback,
                                       void* cookie) noexcept override
    {
        return _base ? _base->CopyItem(MapPath(sourcePath).c_str(), MapPath(destinationPath).c_str(), flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE MoveItem(const wchar_t* sourcePath,
                                       const wchar_t* destinationPath,
                                       FileSystemFlags flags,
                                       const FileSystemOptions* options,
                                       IFileSystemCallback* callback,
                                       void* cookie) noexcept override
    {
        return _base ? _base->MoveItem(MapPath(sourcePath).c_str(), MapPath(destinationPath).c_str(), flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE
    DeleteItem(const wchar_t* path, FileSystemFlags flags, const FileSystemOptions* options, IFileSystemCallback* callback, void* cookie) noexcept override
    {
        return _base ? _base->DeleteItem(MapPath(path).c_str(), flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE RenameItem(const wchar_t* sourcePath,
                                         const wchar_t* destinationPath,
                                         FileSystemFlags flags,
                                         const FileSystemOptions* options,
                                         IFileSystemCallback* callback,
                                         void* cookie) noexcept override
    {
        return _base ? _base->RenameItem(MapPath(sourcePath).c_str(), MapPath(destinationPath).c_str(), flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE CopyItems(const wchar_t* const* sourcePaths,
                                        unsigned long count,
                                        const wchar_t* destinationFolder,
                                        FileSystemFlags flags,
                                        const FileSystemOptions* options,
                                        IFileSystemCallback* callback,
                                        void* cookie) noexcept override
    {
        if (! _base)
        {
            return E_POINTER;
        }

        std::vector<std::wstring> mapped;
        std::vector<const wchar_t*> mappedPtrs;
        mapped.reserve(count);
        mappedPtrs.reserve(count);

        for (unsigned long i = 0; i < count; ++i)
        {
            mapped.emplace_back(MapPath(sourcePaths ? sourcePaths[i] : nullptr).wstring());
            mappedPtrs.emplace_back(mapped.back().c_str());
        }

        const std::filesystem::path mappedDestination = MapPath(destinationFolder);
        return _base->CopyItems(mappedPtrs.data(), count, mappedDestination.c_str(), flags, options, callback, cookie);
    }

    HRESULT STDMETHODCALLTYPE MoveItems(const wchar_t* const* sourcePaths,
                                        unsigned long count,
                                        const wchar_t* destinationFolder,
                                        FileSystemFlags flags,
                                        const FileSystemOptions* options,
                                        IFileSystemCallback* callback,
                                        void* cookie) noexcept override
    {
        if (! _base)
        {
            return E_POINTER;
        }

        std::vector<std::wstring> mapped;
        std::vector<const wchar_t*> mappedPtrs;
        mapped.reserve(count);
        mappedPtrs.reserve(count);

        for (unsigned long i = 0; i < count; ++i)
        {
            mapped.emplace_back(MapPath(sourcePaths ? sourcePaths[i] : nullptr).wstring());
            mappedPtrs.emplace_back(mapped.back().c_str());
        }

        const std::filesystem::path mappedDestination = MapPath(destinationFolder);
        return _base->MoveItems(mappedPtrs.data(), count, mappedDestination.c_str(), flags, options, callback, cookie);
    }

    HRESULT STDMETHODCALLTYPE DeleteItems(const wchar_t* const* paths,
                                          unsigned long count,
                                          FileSystemFlags flags,
                                          const FileSystemOptions* options,
                                          IFileSystemCallback* callback,
                                          void* cookie) noexcept override
    {
        if (! _base)
        {
            return E_POINTER;
        }

        std::vector<std::wstring> mapped;
        std::vector<const wchar_t*> mappedPtrs;
        mapped.reserve(count);
        mappedPtrs.reserve(count);

        for (unsigned long i = 0; i < count; ++i)
        {
            mapped.emplace_back(MapPath(paths ? paths[i] : nullptr).wstring());
            mappedPtrs.emplace_back(mapped.back().c_str());
        }

        return _base->DeleteItems(mappedPtrs.data(), count, flags, options, callback, cookie);
    }

    HRESULT STDMETHODCALLTYPE RenameItems(const FileSystemRenamePair* items,
                                          unsigned long count,
                                          FileSystemFlags flags,
                                          const FileSystemOptions* options,
                                          IFileSystemCallback* callback,
                                          void* cookie) noexcept override
    {
        if (! _base)
        {
            return E_POINTER;
        }

        std::vector<std::wstring> mappedSources;
        std::vector<FileSystemRenamePair> mappedPairs;
        mappedSources.reserve(count);
        mappedPairs.reserve(count);

        for (unsigned long i = 0; i < count; ++i)
        {
            const FileSystemRenamePair& src = items ? items[i] : FileSystemRenamePair{};
            mappedSources.emplace_back(MapPath(src.sourcePath).wstring());

            FileSystemRenamePair pair{};
            pair.sizeBytes  = sizeof(FileSystemRenamePair);
            pair.sourcePath = mappedSources.back().c_str();
            pair.newName    = src.newName;
            mappedPairs.emplace_back(pair);
        }

        return _base->RenameItems(mappedPairs.data(), count, flags, options, callback, cookie);
    }

    HRESULT STDMETHODCALLTYPE GetCapabilities(const char** jsonUtf8) noexcept override
    {
        return _base ? _base->GetCapabilities(jsonUtf8) : E_POINTER;
    }

    // IFileSystemIO
    HRESULT STDMETHODCALLTYPE GetAttributes(const wchar_t* path, unsigned long* fileAttributes) noexcept override
    {
        if (! _baseIo)
        {
            return E_POINTER;
        }
        return _baseIo->GetAttributes(MapPath(path).c_str(), fileAttributes);
    }

    HRESULT STDMETHODCALLTYPE CreateFileReader(const wchar_t* path, IFileReader** reader) noexcept override
    {
        if (! _baseIo)
        {
            return E_POINTER;
        }
        return _baseIo->CreateFileReader(MapPath(path).c_str(), reader);
    }

    HRESULT STDMETHODCALLTYPE CreateFileWriter(const wchar_t* path, FileSystemFlags flags, IFileWriter** writer) noexcept override
    {
        if (! _baseIo)
        {
            return E_POINTER;
        }
        return _baseIo->CreateFileWriter(MapPath(path).c_str(), flags, writer);
    }

    HRESULT STDMETHODCALLTYPE GetFileBasicInformation(const wchar_t* path, FileSystemBasicInformation* info) noexcept override
    {
        if (! _baseIo)
        {
            return E_POINTER;
        }
        return _baseIo->GetFileBasicInformation(MapPath(path).c_str(), info);
    }

    HRESULT STDMETHODCALLTYPE SetFileBasicInformation(const wchar_t* path, const FileSystemBasicInformation* info) noexcept override
    {
        if (! _baseIo)
        {
            return E_POINTER;
        }
        return _baseIo->SetFileBasicInformation(MapPath(path).c_str(), info);
    }

    HRESULT STDMETHODCALLTYPE GetItemProperties(const wchar_t* path, const char** jsonUtf8) noexcept override
    {
        if (! _baseIo)
        {
            return E_POINTER;
        }
        return _baseIo->GetItemProperties(MapPath(path).c_str(), jsonUtf8);
    }

private:
    ~PluginPathMappedRootFileSystem() = default;

    [[nodiscard]] std::filesystem::path MapPath(const wchar_t* path) const noexcept
    {
        if (! path || path[0] == L'\0')
        {
            return _windowsRoot;
        }

        std::wstring text(path);
        std::replace(text.begin(), text.end(), L'\\', L'/');

        const size_t start = text.find_first_not_of(L"/");
        if (start == std::wstring::npos)
        {
            return _windowsRoot;
        }
        text.erase(0, start);
        if (text.empty())
        {
            return _windowsRoot;
        }

        std::replace(text.begin(), text.end(), L'/', L'\\');
        return (_windowsRoot / std::filesystem::path(text)).lexically_normal();
    }

    std::atomic_ulong _refCount{1};
    wil::com_ptr<IFileSystem> _base;
    wil::com_ptr<IFileSystemIO> _baseIo;
    std::filesystem::path _windowsRoot;
};

class PluginPathMappedRootFileSystemNoIO final : public IFileSystem
{
public:
    PluginPathMappedRootFileSystemNoIO(wil::com_ptr<IFileSystem> base, std::filesystem::path windowsRoot) noexcept
        : _base(std::move(base)),
          _windowsRoot(std::move(windowsRoot))
    {
    }

    PluginPathMappedRootFileSystemNoIO(const PluginPathMappedRootFileSystemNoIO&)            = delete;
    PluginPathMappedRootFileSystemNoIO(PluginPathMappedRootFileSystemNoIO&&)                 = delete;
    PluginPathMappedRootFileSystemNoIO& operator=(const PluginPathMappedRootFileSystemNoIO&) = delete;
    PluginPathMappedRootFileSystemNoIO& operator=(PluginPathMappedRootFileSystemNoIO&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (ppvObject == nullptr)
        {
            return E_POINTER;
        }

        if (riid == __uuidof(IUnknown) || riid == __uuidof(IFileSystem))
        {
            *ppvObject = static_cast<IFileSystem*>(this);
            AddRef();
            return S_OK;
        }

        // Do not expose IInformations: DirectoryInfoCache uses it to decide path semantics and would
        // otherwise treat this wrapper as the file plugin, breaking plugin-path mapping.
        if (riid == __uuidof(IInformations))
        {
            *ppvObject = nullptr;
            return E_NOINTERFACE;
        }

        if (riid == __uuidof(IFileSystemIO))
        {
            *ppvObject = nullptr;
            return E_NOINTERFACE;
        }

        if (_base)
        {
            return _base->QueryInterface(riid, ppvObject);
        }

        *ppvObject = nullptr;
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return _refCount.fetch_add(1, std::memory_order_relaxed) + 1;
    }

    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        const ULONG current = _refCount.fetch_sub(1, std::memory_order_acq_rel) - 1;
        if (current == 0)
        {
            delete this;
        }
        return current;
    }

    // IFileSystem
    HRESULT STDMETHODCALLTYPE ReadDirectoryInfo(const wchar_t* path, IFilesInformation** ppFilesInformation) noexcept override
    {
        if (! _base)
        {
            return E_POINTER;
        }
        return _base->ReadDirectoryInfo(MapPath(path).c_str(), ppFilesInformation);
    }

    HRESULT STDMETHODCALLTYPE CopyItem(const wchar_t* sourcePath,
                                       const wchar_t* destinationPath,
                                       FileSystemFlags flags,
                                       const FileSystemOptions* options,
                                       IFileSystemCallback* callback,
                                       void* cookie) noexcept override
    {
        return _base ? _base->CopyItem(MapPath(sourcePath).c_str(), MapPath(destinationPath).c_str(), flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE MoveItem(const wchar_t* sourcePath,
                                       const wchar_t* destinationPath,
                                       FileSystemFlags flags,
                                       const FileSystemOptions* options,
                                       IFileSystemCallback* callback,
                                       void* cookie) noexcept override
    {
        return _base ? _base->MoveItem(MapPath(sourcePath).c_str(), MapPath(destinationPath).c_str(), flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE
    DeleteItem(const wchar_t* path, FileSystemFlags flags, const FileSystemOptions* options, IFileSystemCallback* callback, void* cookie) noexcept override
    {
        return _base ? _base->DeleteItem(MapPath(path).c_str(), flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE RenameItem(const wchar_t* sourcePath,
                                         const wchar_t* destinationPath,
                                         FileSystemFlags flags,
                                         const FileSystemOptions* options,
                                         IFileSystemCallback* callback,
                                         void* cookie) noexcept override
    {
        return _base ? _base->RenameItem(MapPath(sourcePath).c_str(), MapPath(destinationPath).c_str(), flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE CopyItems(const wchar_t* const* sourcePaths,
                                        unsigned long count,
                                        const wchar_t* destinationFolder,
                                        FileSystemFlags flags,
                                        const FileSystemOptions* options,
                                        IFileSystemCallback* callback,
                                        void* cookie) noexcept override
    {
        if (! _base)
        {
            return E_POINTER;
        }

        std::vector<std::wstring> mapped;
        std::vector<const wchar_t*> mappedPtrs;
        mapped.reserve(count);
        mappedPtrs.reserve(count);

        for (unsigned long i = 0; i < count; ++i)
        {
            mapped.emplace_back(MapPath(sourcePaths ? sourcePaths[i] : nullptr).wstring());
            mappedPtrs.emplace_back(mapped.back().c_str());
        }

        const std::filesystem::path mappedDestination = MapPath(destinationFolder);
        return _base->CopyItems(mappedPtrs.data(), count, mappedDestination.c_str(), flags, options, callback, cookie);
    }

    HRESULT STDMETHODCALLTYPE MoveItems(const wchar_t* const* sourcePaths,
                                        unsigned long count,
                                        const wchar_t* destinationFolder,
                                        FileSystemFlags flags,
                                        const FileSystemOptions* options,
                                        IFileSystemCallback* callback,
                                        void* cookie) noexcept override
    {
        if (! _base)
        {
            return E_POINTER;
        }

        std::vector<std::wstring> mapped;
        std::vector<const wchar_t*> mappedPtrs;
        mapped.reserve(count);
        mappedPtrs.reserve(count);

        for (unsigned long i = 0; i < count; ++i)
        {
            mapped.emplace_back(MapPath(sourcePaths ? sourcePaths[i] : nullptr).wstring());
            mappedPtrs.emplace_back(mapped.back().c_str());
        }

        const std::filesystem::path mappedDestination = MapPath(destinationFolder);
        return _base->MoveItems(mappedPtrs.data(), count, mappedDestination.c_str(), flags, options, callback, cookie);
    }

    HRESULT STDMETHODCALLTYPE DeleteItems(const wchar_t* const* paths,
                                          unsigned long count,
                                          FileSystemFlags flags,
                                          const FileSystemOptions* options,
                                          IFileSystemCallback* callback,
                                          void* cookie) noexcept override
    {
        if (! _base)
        {
            return E_POINTER;
        }

        std::vector<std::wstring> mapped;
        std::vector<const wchar_t*> mappedPtrs;
        mapped.reserve(count);
        mappedPtrs.reserve(count);

        for (unsigned long i = 0; i < count; ++i)
        {
            mapped.emplace_back(MapPath(paths ? paths[i] : nullptr).wstring());
            mappedPtrs.emplace_back(mapped.back().c_str());
        }

        return _base->DeleteItems(mappedPtrs.data(), count, flags, options, callback, cookie);
    }

    HRESULT STDMETHODCALLTYPE RenameItems(const FileSystemRenamePair* items,
                                          unsigned long count,
                                          FileSystemFlags flags,
                                          const FileSystemOptions* options,
                                          IFileSystemCallback* callback,
                                          void* cookie) noexcept override
    {
        if (! _base)
        {
            return E_POINTER;
        }

        std::vector<std::wstring> mappedSources;
        std::vector<FileSystemRenamePair> mappedPairs;
        mappedSources.reserve(count);
        mappedPairs.reserve(count);

        for (unsigned long i = 0; i < count; ++i)
        {
            const FileSystemRenamePair& src = items ? items[i] : FileSystemRenamePair{};
            mappedSources.emplace_back(MapPath(src.sourcePath).wstring());

            FileSystemRenamePair pair{};
            pair.sizeBytes  = sizeof(FileSystemRenamePair);
            pair.sourcePath = mappedSources.back().c_str();
            pair.newName    = src.newName;
            mappedPairs.emplace_back(pair);
        }

        return _base->RenameItems(mappedPairs.data(), count, flags, options, callback, cookie);
    }

    HRESULT STDMETHODCALLTYPE GetCapabilities(const char** jsonUtf8) noexcept override
    {
        return _base ? _base->GetCapabilities(jsonUtf8) : E_POINTER;
    }

private:
    ~PluginPathMappedRootFileSystemNoIO() = default;

    [[nodiscard]] std::filesystem::path MapPath(const wchar_t* path) const noexcept
    {
        if (! path || path[0] == L'\0')
        {
            return _windowsRoot;
        }

        std::wstring text(path);
        std::replace(text.begin(), text.end(), L'\\', L'/');

        const size_t start = text.find_first_not_of(L"/");
        if (start == std::wstring::npos)
        {
            return _windowsRoot;
        }
        text.erase(0, start);
        if (text.empty())
        {
            return _windowsRoot;
        }

        std::replace(text.begin(), text.end(), L'/', L'\\');
        return (_windowsRoot / std::filesystem::path(text)).lexically_normal();
    }

    std::atomic_ulong _refCount{1};
    wil::com_ptr<IFileSystem> _base;
    std::filesystem::path _windowsRoot;
};

class CountingReadDirectoryFileSystem final : public IFileSystem
{
public:
    CountingReadDirectoryFileSystem(wil::com_ptr<IFileSystem> base, std::atomic_uint32_t* counter) noexcept : _base(std::move(base)), _counter(counter)
    {
    }

    CountingReadDirectoryFileSystem(const CountingReadDirectoryFileSystem&)            = delete;
    CountingReadDirectoryFileSystem(CountingReadDirectoryFileSystem&&)                 = delete;
    CountingReadDirectoryFileSystem& operator=(const CountingReadDirectoryFileSystem&) = delete;
    CountingReadDirectoryFileSystem& operator=(CountingReadDirectoryFileSystem&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (ppvObject == nullptr)
        {
            return E_POINTER;
        }

        if (riid == __uuidof(IUnknown) || riid == __uuidof(IFileSystem))
        {
            *ppvObject = static_cast<IFileSystem*>(this);
            AddRef();
            return S_OK;
        }

        if (_base)
        {
            return _base->QueryInterface(riid, ppvObject);
        }

        *ppvObject = nullptr;
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return _refCount.fetch_add(1, std::memory_order_relaxed) + 1;
    }

    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        const ULONG current = _refCount.fetch_sub(1, std::memory_order_acq_rel) - 1;
        if (current == 0)
        {
            delete this;
        }
        return current;
    }

    // IFileSystem
    HRESULT STDMETHODCALLTYPE ReadDirectoryInfo(const wchar_t* path, IFilesInformation** ppFilesInformation) noexcept override
    {
        if (! _base)
        {
            return E_POINTER;
        }

        if (_counter)
        {
            static_cast<void>(_counter->fetch_add(1u, std::memory_order_relaxed));
        }

        return _base->ReadDirectoryInfo(path, ppFilesInformation);
    }

    HRESULT STDMETHODCALLTYPE CopyItem(const wchar_t* sourcePath,
                                       const wchar_t* destinationPath,
                                       FileSystemFlags flags,
                                       const FileSystemOptions* options,
                                       IFileSystemCallback* callback,
                                       void* cookie) noexcept override
    {
        return _base ? _base->CopyItem(sourcePath, destinationPath, flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE MoveItem(const wchar_t* sourcePath,
                                       const wchar_t* destinationPath,
                                       FileSystemFlags flags,
                                       const FileSystemOptions* options,
                                       IFileSystemCallback* callback,
                                       void* cookie) noexcept override
    {
        return _base ? _base->MoveItem(sourcePath, destinationPath, flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE
    DeleteItem(const wchar_t* path, FileSystemFlags flags, const FileSystemOptions* options, IFileSystemCallback* callback, void* cookie) noexcept override
    {
        return _base ? _base->DeleteItem(path, flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE RenameItem(const wchar_t* sourcePath,
                                         const wchar_t* destinationPath,
                                         FileSystemFlags flags,
                                         const FileSystemOptions* options,
                                         IFileSystemCallback* callback,
                                         void* cookie) noexcept override
    {
        return _base ? _base->RenameItem(sourcePath, destinationPath, flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE CopyItems(const wchar_t* const* sourcePaths,
                                        unsigned long count,
                                        const wchar_t* destinationFolder,
                                        FileSystemFlags flags,
                                        const FileSystemOptions* options,
                                        IFileSystemCallback* callback,
                                        void* cookie) noexcept override
    {
        return _base ? _base->CopyItems(sourcePaths, count, destinationFolder, flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE MoveItems(const wchar_t* const* sourcePaths,
                                        unsigned long count,
                                        const wchar_t* destinationFolder,
                                        FileSystemFlags flags,
                                        const FileSystemOptions* options,
                                        IFileSystemCallback* callback,
                                        void* cookie) noexcept override
    {
        return _base ? _base->MoveItems(sourcePaths, count, destinationFolder, flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE DeleteItems(const wchar_t* const* paths,
                                          unsigned long count,
                                          FileSystemFlags flags,
                                          const FileSystemOptions* options,
                                          IFileSystemCallback* callback,
                                          void* cookie) noexcept override
    {
        return _base ? _base->DeleteItems(paths, count, flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE RenameItems(const FileSystemRenamePair* items,
                                          unsigned long count,
                                          FileSystemFlags flags,
                                          const FileSystemOptions* options,
                                          IFileSystemCallback* callback,
                                          void* cookie) noexcept override
    {
        return _base ? _base->RenameItems(items, count, flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE GetCapabilities(const char** jsonUtf8) noexcept override
    {
        return _base ? _base->GetCapabilities(jsonUtf8) : E_POINTER;
    }

private:
    ~CountingReadDirectoryFileSystem() = default;

    std::atomic_ulong _refCount{1};
    wil::com_ptr<IFileSystem> _base;
    std::atomic_uint32_t* _counter = nullptr;
};

[[nodiscard]] wil::com_ptr<IFileSystem> CreatePluginPathMappedRootFileSystem(const wil::com_ptr<IFileSystem>& base,
                                                                             const std::filesystem::path& windowsRoot) noexcept
{
    wil::com_ptr<IFileSystem> wrapped;
    auto* wrapper = new (std::nothrow) PluginPathMappedRootFileSystem(base, windowsRoot);
    if (! wrapper)
    {
        return {};
    }
    wrapped.attach(wrapper);
    return wrapped;
}

[[nodiscard]] wil::com_ptr<IFileSystem> CreatePluginPathMappedRootFileSystemNoIO(const wil::com_ptr<IFileSystem>& base,
                                                                                 const std::filesystem::path& windowsRoot) noexcept
{
    wil::com_ptr<IFileSystem> wrapped;
    auto* wrapper = new (std::nothrow) PluginPathMappedRootFileSystemNoIO(base, windowsRoot);
    if (! wrapper)
    {
        return {};
    }
    wrapped.attach(wrapper);
    return wrapped;
}

[[nodiscard]] wil::com_ptr<IFileSystem> CreateCountingReadDirectoryFileSystem(const wil::com_ptr<IFileSystem>& base, std::atomic_uint32_t* counter) noexcept
{
    wil::com_ptr<IFileSystem> wrapped;
    auto* wrapper = new (std::nothrow) CountingReadDirectoryFileSystem(base, counter);
    if (! wrapper)
    {
        return {};
    }
    wrapped.attach(wrapper);
    return wrapped;
}

[[nodiscard]] wil::com_ptr<IFileSystem> GetDummyFileSystem() noexcept
{
    return SelfTest::GetFileSystem(kBuiltinDummyFileSystemId);
}

[[nodiscard]] bool CreateFileSystemIo(const wil::com_ptr<IFileSystem>& fs, wil::com_ptr<IFileSystemIO>& outIo) noexcept
{
    outIo.reset();
    if (! fs)
    {
        return false;
    }

    const HRESULT hr = fs->QueryInterface(__uuidof(IFileSystemIO), outIo.put_void());
    return SUCCEEDED(hr) && static_cast<bool>(outIo);
}

[[nodiscard]] bool CreateInformations(const wil::com_ptr<IFileSystem>& fs, wil::com_ptr<IInformations>& outInfo) noexcept
{
    outInfo.reset();
    if (! fs)
    {
        return false;
    }

    const HRESULT hr = fs->QueryInterface(__uuidof(IInformations), outInfo.put_void());
    return SUCCEEDED(hr) && static_cast<bool>(outInfo);
}

[[nodiscard]] bool CreateFileSystemDirectoryOperations(const wil::com_ptr<IFileSystem>& fs, wil::com_ptr<IFileSystemDirectoryOperations>& outOps) noexcept
{
    outOps.reset();
    if (! fs)
    {
        return false;
    }

    const HRESULT hr = fs->QueryInterface(__uuidof(IFileSystemDirectoryOperations), outOps.put_void());
    return SUCCEEDED(hr) && static_cast<bool>(outOps);
}

[[nodiscard]] bool EnsureDirectoryExistsFsOps(const wil::com_ptr<IFileSystemDirectoryOperations>& ops, const std::filesystem::path& path) noexcept
{
    if (! ops)
    {
        return false;
    }

    const std::filesystem::path normalized = path.lexically_normal();
    std::filesystem::path current          = normalized.root_path();
    for (const auto& part : normalized.relative_path())
    {
        current /= part;
        const HRESULT hr = ops->CreateDirectory(current.c_str());
        if (SUCCEEDED(hr) || hr == HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS))
        {
            continue;
        }
        return false;
    }

    return true;
}

[[nodiscard]] bool WriteFileBytesFsIo(const wil::com_ptr<IFileSystemIO>& io, const std::filesystem::path& path, const void* data, size_t sizeBytes) noexcept
{
    if (! io || ! data)
    {
        return false;
    }

    if (sizeBytes > static_cast<size_t>(std::numeric_limits<unsigned long>::max()))
    {
        return false;
    }

    wil::com_ptr<IFileWriter> writer;
    const HRESULT createHr = io->CreateFileWriter(path.c_str(), FILESYSTEM_FLAG_NONE, writer.put());
    if (FAILED(createHr) || ! writer)
    {
        return false;
    }

    unsigned long written = 0;
    const HRESULT writeHr = writer->Write(data, static_cast<unsigned long>(sizeBytes), &written);
    if (FAILED(writeHr) || written != static_cast<unsigned long>(sizeBytes))
    {
        return false;
    }

    return SUCCEEDED(writer->Commit());
}

[[nodiscard]] bool WriteFileTextFsIo(const wil::com_ptr<IFileSystemIO>& io, const std::filesystem::path& path, std::string_view text) noexcept
{
    return WriteFileBytesFsIo(io, path, text.data(), text.size());
}

[[nodiscard]] std::wstring ToPluginPathText(const std::filesystem::path& path) noexcept
{
    return path.generic_wstring();
}

[[nodiscard]] std::wstring JoinPluginPathForSelfTest(std::wstring_view base, std::wstring_view leaf) noexcept
{
    std::wstring result = NormalizePluginPathForSelfTest(base);
    if (result.empty())
    {
        return {};
    }

    if (! leaf.empty())
    {
        if (result.back() != L'/')
        {
            result.push_back(L'/');
        }
        result.append(leaf);
    }

    return NormalizePluginPathForSelfTest(result);
}

[[nodiscard]] std::wstring MakeConnectionPathForSelfTest(std::wstring_view profileName, std::wstring_view pluginPath) noexcept
{
    if (profileName.empty() || pluginPath.empty() || pluginPath.front() != L'/')
    {
        return {};
    }

    return std::format(L"/@conn:{}{}", profileName, pluginPath);
}

[[nodiscard]] bool PathExistsFsIo(const wil::com_ptr<IFileSystemIO>& io, const std::filesystem::path& path, unsigned long* attributes = nullptr) noexcept
{
    if (! io)
    {
        return false;
    }

    unsigned long attrs            = 0;
    const std::wstring pluginPath  = ToPluginPathText(path);
    const HRESULT hr               = io->GetAttributes(pluginPath.c_str(), &attrs);
    if (FAILED(hr))
    {
        return false;
    }

    if (attributes)
    {
        *attributes = attrs;
    }
    return true;
}

[[nodiscard]] std::filesystem::path GetWorkspaceRootFromSourcePath() noexcept
{
    return std::filesystem::path(__FILE__).parent_path().parent_path();
}

[[nodiscard]] std::optional<std::wstring> FindFirstRegularEntryPath(const wil::com_ptr<IFileSystem>& fs, std::wstring_view rootPath) noexcept
{
    if (! fs || rootPath.empty())
    {
        return std::nullopt;
    }

    std::vector<std::wstring> pending{std::wstring(rootPath)};
    while (! pending.empty())
    {
        const std::wstring current = std::move(pending.back());
        pending.pop_back();

        wil::com_ptr<IFilesInformation> info;
        const HRESULT hr = fs->ReadDirectoryInfo(current.c_str(), info.put());
        if (FAILED(hr) || ! info)
        {
            continue;
        }

        FileInfo* head = nullptr;
        if (FAILED(info->GetBuffer(&head)) || head == nullptr)
        {
            continue;
        }

        for (FileInfo* entry = head; entry != nullptr;)
        {
            if (entry->FileNameSize >= sizeof(wchar_t))
            {
                const size_t charCount = entry->FileNameSize / sizeof(wchar_t);
                std::wstring path(current);
                if (! path.empty() && path.back() != L'/')
                {
                    path.push_back(L'/');
                }
                path.append(entry->FileName, entry->FileName + charCount);

                const bool isDirectory = (entry->FileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
                if (! isDirectory)
                {
                    return path;
                }

                pending.push_back(std::move(path));
            }

            if (entry->NextEntryOffset == 0)
            {
                break;
            }

            entry = reinterpret_cast<FileInfo*>(reinterpret_cast<std::byte*>(entry) + entry->NextEntryOffset);
        }
    }

    return std::nullopt;
}

enum class DirectorySizeCallbackMode : uint8_t
{
    Success,
    FailProgress,
    FailCancel,
    Cancel,
};

constexpr HRESULT kDirectorySizeProgressFailureHr = HRESULT_FROM_WIN32(ERROR_RETRY);
constexpr HRESULT kDirectorySizeCancelFailureHr   = E_ACCESSDENIED;

class RecordingDirectorySizeCallback final : public IFileSystemDirectorySizeCallback
{
public:
    explicit RecordingDirectorySizeCallback(DirectorySizeCallbackMode mode) noexcept : _mode(mode)
    {
    }

    RecordingDirectorySizeCallback(const RecordingDirectorySizeCallback&)            = delete;
    RecordingDirectorySizeCallback(RecordingDirectorySizeCallback&&)                 = delete;
    RecordingDirectorySizeCallback& operator=(const RecordingDirectorySizeCallback&) = delete;
    RecordingDirectorySizeCallback& operator=(RecordingDirectorySizeCallback&&)      = delete;

    HRESULT STDMETHODCALLTYPE DirectorySizeProgress(
        uint64_t scannedEntries, uint64_t totalBytes, uint64_t fileCount, uint64_t directoryCount, const wchar_t* currentPath, [[maybe_unused]] void* cookie) noexcept override
    {
        _progressCalls.fetch_add(1u, std::memory_order_relaxed);
        _lastScannedEntries.store(scannedEntries, std::memory_order_release);
        _lastTotalBytes.store(totalBytes, std::memory_order_release);
        _lastFileCount.store(fileCount, std::memory_order_release);
        _lastDirectoryCount.store(directoryCount, std::memory_order_release);

        if (currentPath == nullptr)
        {
            _nullPathProgressCalls.fetch_add(1u, std::memory_order_relaxed);
        }

        if (_mode == DirectorySizeCallbackMode::FailProgress)
        {
            return kDirectorySizeProgressFailureHr;
        }

        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE DirectorySizeShouldCancel(BOOL* pCancel, [[maybe_unused]] void* cookie) noexcept override
    {
        if (pCancel == nullptr)
        {
            return E_POINTER;
        }

        _cancelCalls.fetch_add(1u, std::memory_order_relaxed);

        if (_mode == DirectorySizeCallbackMode::FailCancel)
        {
            return kDirectorySizeCancelFailureHr;
        }

        *pCancel = (_mode == DirectorySizeCallbackMode::Cancel) ? TRUE : FALSE;
        return S_OK;
    }

    [[nodiscard]] uint32_t ProgressCalls() const noexcept
    {
        return _progressCalls.load(std::memory_order_acquire);
    }

    [[nodiscard]] uint32_t CancelCalls() const noexcept
    {
        return _cancelCalls.load(std::memory_order_acquire);
    }

    [[nodiscard]] uint32_t NullPathProgressCalls() const noexcept
    {
        return _nullPathProgressCalls.load(std::memory_order_acquire);
    }

    [[nodiscard]] uint64_t LastTotalBytes() const noexcept
    {
        return _lastTotalBytes.load(std::memory_order_acquire);
    }

    [[nodiscard]] uint64_t LastFileCount() const noexcept
    {
        return _lastFileCount.load(std::memory_order_acquire);
    }

    [[nodiscard]] uint64_t LastDirectoryCount() const noexcept
    {
        return _lastDirectoryCount.load(std::memory_order_acquire);
    }

private:
    DirectorySizeCallbackMode _mode = DirectorySizeCallbackMode::Success;
    std::atomic<uint32_t> _progressCalls{0};
    std::atomic<uint32_t> _cancelCalls{0};
    std::atomic<uint32_t> _nullPathProgressCalls{0};
    std::atomic<uint64_t> _lastScannedEntries{0};
    std::atomic<uint64_t> _lastTotalBytes{0};
    std::atomic<uint64_t> _lastFileCount{0};
    std::atomic<uint64_t> _lastDirectoryCount{0};
};

[[nodiscard]] bool ValidateDirectorySizeCallbackContract(SelfTest::CaseState& state,
                                                         IFileSystemDirectoryOperations* dirOps,
                                                         const std::wstring& path,
                                                         FileSystemFlags flags = FILESYSTEM_FLAG_NONE) noexcept
{
    if (dirOps == nullptr || path.empty())
    {
        state.Require(false, L"Directory size callback contract: invalid test inputs.");
        return false;
    }

    FileSystemDirectorySizeResult baseline{};
    baseline.sizeBytes = sizeof(FileSystemDirectorySizeResult);
    const HRESULT baselineHr = dirOps->GetDirectorySize(path.c_str(), flags, nullptr, nullptr, &baseline);
    state.Require(SUCCEEDED(baselineHr) && SUCCEEDED(baseline.status),
                  std::format(L"Directory size callback contract: baseline GetDirectorySize failed. hr=0x{:08X} status=0x{:08X}",
                              static_cast<unsigned long>(baselineHr),
                              static_cast<unsigned long>(baseline.status)));
    if (FAILED(baselineHr) || FAILED(baseline.status))
    {
        return false;
    }

    const auto runCase = [&](DirectorySizeCallbackMode mode, HRESULT expectedHr, std::wstring_view label) noexcept
    {
        RecordingDirectorySizeCallback callback(mode);
        FileSystemDirectorySizeResult result{};
        result.sizeBytes = sizeof(FileSystemDirectorySizeResult);

        const HRESULT hr = dirOps->GetDirectorySize(path.c_str(), flags, &callback, nullptr, &result);
        state.Require(hr == expectedHr && result.status == expectedHr,
                      std::format(L"Directory size callback contract ({}): expected hr/status 0x{:08X}, got hr=0x{:08X} status=0x{:08X}.",
                                  label,
                                  static_cast<unsigned long>(expectedHr),
                                  static_cast<unsigned long>(hr),
                                  static_cast<unsigned long>(result.status)));
        state.Require(callback.ProgressCalls() >= 1u,
                      std::format(L"Directory size callback contract ({}): expected at least one progress callback.", label));
        if (mode != DirectorySizeCallbackMode::FailProgress)
        {
            state.Require(callback.CancelCalls() >= 1u,
                          std::format(L"Directory size callback contract ({}): expected at least one cancel callback.", label));
        }
    };

    runCase(DirectorySizeCallbackMode::FailProgress, kDirectorySizeProgressFailureHr, L"progress failure");
    if (! state.failure.empty())
    {
        return false;
    }

    runCase(DirectorySizeCallbackMode::FailCancel, kDirectorySizeCancelFailureHr, L"cancel failure");
    if (! state.failure.empty())
    {
        return false;
    }

    runCase(DirectorySizeCallbackMode::Cancel, HRESULT_FROM_WIN32(ERROR_CANCELLED), L"cancel");
    if (! state.failure.empty())
    {
        return false;
    }

    RecordingDirectorySizeCallback successCallback(DirectorySizeCallbackMode::Success);
    FileSystemDirectorySizeResult success{};
    success.sizeBytes = sizeof(FileSystemDirectorySizeResult);
    const HRESULT successHr = dirOps->GetDirectorySize(path.c_str(), flags, &successCallback, nullptr, &success);
    state.Require(SUCCEEDED(successHr) && SUCCEEDED(success.status),
                  std::format(L"Directory size callback contract (success): hr=0x{:08X} status=0x{:08X}.",
                              static_cast<unsigned long>(successHr),
                              static_cast<unsigned long>(success.status)));
    state.Require(success.totalBytes == baseline.totalBytes,
                  std::format(L"Directory size callback contract (success): totalBytes mismatch. expected={} got={}.",
                              baseline.totalBytes,
                              success.totalBytes));
    state.Require(success.fileCount == baseline.fileCount,
                  std::format(L"Directory size callback contract (success): fileCount mismatch. expected={} got={}.",
                              baseline.fileCount,
                              success.fileCount));
    state.Require(success.directoryCount == baseline.directoryCount,
                  std::format(L"Directory size callback contract (success): directoryCount mismatch. expected={} got={}.",
                              baseline.directoryCount,
                              success.directoryCount));
    state.Require(successCallback.NullPathProgressCalls() >= 1u,
                  L"Directory size callback contract (success): expected a final progress callback with currentPath=nullptr.");

    return state.failure.empty();
}

struct RecordedFileSystemItem final
{
    bool seen = false;
    HRESULT status = S_OK;
    std::wstring sourcePath;
    std::wstring destinationPath;
};

class RecordingFileSystemCallback final : public IFileSystemCallback
{
public:
    explicit RecordingFileSystemCallback(unsigned long expectedItemCount) : _items(expectedItemCount)
    {
    }

    RecordingFileSystemCallback(const RecordingFileSystemCallback&)            = delete;
    RecordingFileSystemCallback(RecordingFileSystemCallback&&)                 = delete;
    RecordingFileSystemCallback& operator=(const RecordingFileSystemCallback&) = delete;
    RecordingFileSystemCallback& operator=(RecordingFileSystemCallback&&)      = delete;

    HRESULT STDMETHODCALLTYPE FileSystemProgress([[maybe_unused]] FileSystemOperation operationType,
                                                 [[maybe_unused]] unsigned long totalItems,
                                                 [[maybe_unused]] unsigned long completedItems,
                                                 [[maybe_unused]] uint64_t totalBytes,
                                                 [[maybe_unused]] uint64_t completedBytes,
                                                 [[maybe_unused]] const wchar_t* currentSourcePath,
                                                 [[maybe_unused]] const wchar_t* currentDestinationPath,
                                                 [[maybe_unused]] uint64_t currentItemTotalBytes,
                                                 [[maybe_unused]] uint64_t currentItemCompletedBytes,
                                                 [[maybe_unused]] FileSystemOptions* options,
                                                 [[maybe_unused]] uint64_t progressStreamId,
                                                 [[maybe_unused]] void* cookie) noexcept override
    {
        _progressCalls.fetch_add(1u, std::memory_order_relaxed);
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE FileSystemItemCompleted([[maybe_unused]] FileSystemOperation operationType,
                                                      unsigned long itemIndex,
                                                      const wchar_t* sourcePath,
                                                      const wchar_t* destinationPath,
                                                      HRESULT status,
                                                      [[maybe_unused]] FileSystemOptions* options,
                                                      [[maybe_unused]] void* cookie) noexcept override
    {
        {
            std::lock_guard lock(_mutex);
            if (itemIndex < _items.size())
            {
                RecordedFileSystemItem& item = _items[itemIndex];
                item.seen                    = true;
                item.status                  = status;
                item.sourcePath              = sourcePath ? sourcePath : L"";
                item.destinationPath         = destinationPath ? destinationPath : L"";
            }
        }

        _completedCalls.fetch_add(1u, std::memory_order_relaxed);
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE FileSystemShouldCancel(BOOL* pCancel, [[maybe_unused]] void* cookie) noexcept override
    {
        if (pCancel == nullptr)
        {
            return E_POINTER;
        }

        *pCancel = FALSE;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE FileSystemIssue([[maybe_unused]] FileSystemOperation operationType,
                                              [[maybe_unused]] const wchar_t* sourcePath,
                                              [[maybe_unused]] const wchar_t* destinationPath,
                                              [[maybe_unused]] HRESULT status,
                                              FileSystemIssueAction* action,
                                              [[maybe_unused]] FileSystemOptions* options,
                                              [[maybe_unused]] void* cookie) noexcept override
    {
        _unexpectedIssue.store(true, std::memory_order_release);
        if (action != nullptr)
        {
            *action = FileSystemIssueAction::Cancel;
        }
        return E_UNEXPECTED;
    }

    [[nodiscard]] bool TryGetItem(unsigned long itemIndex, RecordedFileSystemItem& outItem) const noexcept
    {
        std::lock_guard lock(_mutex);
        if (itemIndex >= _items.size() || ! _items[itemIndex].seen)
        {
            return false;
        }

        outItem = _items[itemIndex];
        return true;
    }

    [[nodiscard]] unsigned long CompletedCount() const noexcept
    {
        return _completedCalls.load(std::memory_order_acquire);
    }

    [[nodiscard]] bool SawUnexpectedIssue() const noexcept
    {
        return _unexpectedIssue.load(std::memory_order_acquire);
    }

private:
    mutable std::mutex _mutex;
    std::vector<RecordedFileSystemItem> _items;
    std::atomic<unsigned long> _progressCalls{0};
    std::atomic<unsigned long> _completedCalls{0};
    std::atomic<bool> _unexpectedIssue{false};
};

#ifndef SYMBOLIC_LINK_FLAG_ALLOW_UNPRIVILEGED_CREATE
#define SYMBOLIC_LINK_FLAG_ALLOW_UNPRIVILEGED_CREATE 0x2u
#endif

[[nodiscard]] bool TryCreateDirectorySymlink(const std::filesystem::path& linkPath, const std::filesystem::path& targetPath) noexcept
{
    const DWORD flags = SYMBOLIC_LINK_FLAG_DIRECTORY;

    if (::CreateSymbolicLinkW(linkPath.c_str(), targetPath.c_str(), flags | SYMBOLIC_LINK_FLAG_ALLOW_UNPRIVILEGED_CREATE) != 0)
    {
        return true;
    }

    if (::CreateSymbolicLinkW(linkPath.c_str(), targetPath.c_str(), flags) != 0)
    {
        return true;
    }

    return false;
}

[[nodiscard]] bool WriteFileFill(const std::filesystem::path& path, char ch, size_t sizeBytes) noexcept
{
    if (sizeBytes == 0)
    {
        return SelfTest::WriteBinaryFile(path, {});
    }

    const std::string text(sizeBytes, ch);
    const std::span<const char> textBytes(text.data(), text.size());
    return SelfTest::WriteBinaryFile(path, std::as_bytes(textBytes));
}

struct CaseFolders
{
    std::filesystem::path left;
    std::filesystem::path right;
};

[[nodiscard]] std::optional<CaseFolders> CreateCaseFolders(const std::filesystem::path& base, std::wstring_view caseName) noexcept
{
    std::filesystem::path caseRoot = base / std::filesystem::path(caseName);
    std::filesystem::path left     = caseRoot / L"left";
    std::filesystem::path right    = caseRoot / L"right";

    SelfTest::EnsureDirectory(left);
    SelfTest::EnsureDirectory(right);
    if (! SelfTest::PathExists(left) || ! SelfTest::PathExists(right))
    {
        return std::nullopt;
    }

    return CaseFolders{std::move(left), std::move(right)};
}

void AppendCompareSelfTestTraceLine(std::wstring_view message) noexcept
{
    Trace(message);
}

[[nodiscard]] std::vector<std::wstring> EnumerateDirectoryNames(const wil::com_ptr<IFileSystem>& fs,
                                                                const std::filesystem::path& folder,
                                                                SelfTest::CaseState& state) noexcept
{
    if (! fs)
    {
        state.Require(false, L"EnumerateDirectoryNames: file system is null.");
        return {};
    }

    wil::com_ptr<IFilesInformation> info;
    const HRESULT hr = fs->ReadDirectoryInfo(folder.c_str(), info.put());
    state.Require(SUCCEEDED(hr), L"EnumerateDirectoryNames: ReadDirectoryInfo failed.");
    if (FAILED(hr) || ! info)
    {
        return {};
    }

    FileInfo* head         = nullptr;
    const HRESULT hrBuffer = info->GetBuffer(&head);
    state.Require(SUCCEEDED(hrBuffer), L"EnumerateDirectoryNames: GetBuffer failed.");
    if (FAILED(hrBuffer) || head == nullptr)
    {
        return {};
    }

    std::vector<std::wstring> result;
    for (FileInfo* entry = head; entry != nullptr;)
    {
        const size_t nameChars = static_cast<size_t>(entry->FileNameSize) / sizeof(wchar_t);
        result.emplace_back(entry->FileName, nameChars);

        if (entry->NextEntryOffset == 0)
        {
            break;
        }
        entry = reinterpret_cast<FileInfo*>(reinterpret_cast<unsigned char*>(entry) + entry->NextEntryOffset);
    }

    return result;
}

[[nodiscard]] bool ContainsName(const std::vector<std::wstring>& names, std::wstring_view name) noexcept
{
    return std::any_of(names.begin(), names.end(), [&](const std::wstring& value) noexcept { return value == name; });
}

struct GetDecisionSehContext
{
    CompareDirectoriesSession* session                                   = nullptr;
    std::shared_ptr<const CompareDirectoriesFolderDecision>* outDecision = nullptr;
};

void InvokeGetRootDecision(void* rawContext) noexcept
{
    auto* ctx = static_cast<GetDecisionSehContext*>(rawContext);
    if (! ctx || ! ctx->session || ! ctx->outDecision)
    {
        return;
    }

    *ctx->outDecision = ctx->session->GetOrComputeDecision(std::filesystem::path{});
}

[[nodiscard]] bool TryGetRootDecisionWithSeh(CompareDirectoriesSession& session, std::shared_ptr<const CompareDirectoriesFolderDecision>& outDecision) noexcept
{
    GetDecisionSehContext ctx{};
    ctx.session     = &session;
    ctx.outDecision = &outDecision;

    __try
    {
        InvokeGetRootDecision(&ctx);
        return true;
    }
    __except (CrashHandler::WriteDumpForException(GetExceptionInformation()))
    {
        return false;
    }
}

[[nodiscard]] std::shared_ptr<const CompareDirectoriesFolderDecision> ComputeRootDecision(wil::com_ptr<IFileSystem> baseFs,
                                                                                          const CaseFolders& folders,
                                                                                          Common::Settings::CompareDirectoriesSettings settings,
                                                                                          SelfTest::CaseState& state) noexcept
{
    if (! baseFs)
    {
        state.Require(false, L"Base file system is null.");
        return {};
    }

    auto session = std::make_shared<CompareDirectoriesSession>(baseFs, baseFs, folders.left, folders.right, std::move(settings));
    std::shared_ptr<const CompareDirectoriesFolderDecision> decision;
    if (! TryGetRootDecisionWithSeh(*session, decision))
    {
        state.Require(false, L"GetOrComputeDecision crashed.");
        return {};
    }
    state.Require(static_cast<bool>(decision), L"GetOrComputeDecision returned null.");
    if (! decision)
    {
        return {};
    }

    state.Require(SUCCEEDED(decision->hr), L"Decision hr is failure.");
    return decision;
}

[[nodiscard]] bool StartScanAndWaitForIdle(const std::shared_ptr<CompareDirectoriesSession>& session, std::chrono::milliseconds timeout) noexcept
{
    if (! session)
    {
        return false;
    }

    std::mutex mutex;
    std::condition_variable cv;
    bool started = false;
    bool done    = false;

    session->SetScanProgressCallback([&](const std::filesystem::path&, std::wstring_view, uint64_t, uint64_t, uint32_t activeScans, uint64_t, uint64_t) noexcept
    {
        std::lock_guard lock(mutex);
        if (activeScans != 0u)
        {
            started = true;
        }
        if (started && activeScans == 0u)
        {
            done = true;
            cv.notify_all();
        }
    });

    session->StartScan();

    {
        std::unique_lock lock(mutex);
        static_cast<void>(cv.wait_for(lock, timeout, [&] { return done; }));
    }

    session->SetScanProgressCallback({});
    return done;
}

[[nodiscard]] bool DrainPendingSubdirUpdates(const std::shared_ptr<CompareDirectoriesSession>& session, size_t maxIterations) noexcept
{
    if (! session)
    {
        return false;
    }

    for (size_t i = 0; i < maxIterations; ++i)
    {
        if (! session->FlushPendingSubdirUpdatesBudgeted(64))
        {
            return true;
        }
        std::this_thread::sleep_for(std::chrono::milliseconds(1));
    }

    return false;
}

[[nodiscard]] const CompareDirectoriesItemDecision* FindItem(const CompareDirectoriesFolderDecision& decision, std::wstring_view name) noexcept
{
    const auto it = decision.items.find(name);
    if (it == decision.items.end())
    {
        return nullptr;
    }
    return &it->second;
}

[[nodiscard]] std::shared_ptr<const CompareDirectoriesFolderDecision> WaitForContentCompare(const std::shared_ptr<CompareDirectoriesSession>& session,
                                                                                            const std::filesystem::path& relativeFolder,
                                                                                            std::wstring_view itemName,
                                                                                            SelfTest::CaseState& state) noexcept
{
    if (! session)
    {
        state.Require(false, L"WaitForContentCompare: session is null.");
        return {};
    }

    const auto deadline =
        std::chrono::steady_clock::now() + std::chrono::milliseconds{static_cast<std::chrono::milliseconds::rep>(SelfTest::ScaleTimeout(5000))};
    while (std::chrono::steady_clock::now() < deadline)
    {
        auto decision = session->GetOrComputeDecision(relativeFolder);
        state.Require(static_cast<bool>(decision), L"WaitForContentCompare: decision is null.");
        if (! decision)
        {
            return {};
        }

        const auto* item = FindItem(*decision, itemName);
        // In differences-only mode, pending content placeholders may be elided to keep memory bounded.
        // Allow the item to appear later once content compare determines it is actually different.
        if (! item)
        {
            std::this_thread::sleep_for(std::chrono::milliseconds(10));
            continue;
        }

        if (! HasFlag(item->differenceMask, CompareDirectoriesDiffBit::ContentPending))
        {
            return decision;
        }

        std::this_thread::sleep_for(std::chrono::milliseconds(10));
    }

    state.Require(false, std::format(L"Timed out waiting for content compare: {}.", itemName));
    return session->GetOrComputeDecision(relativeFolder);
}
} // namespace

bool CompareDirectoriesSelfTest::Run(const SelfTest::SelfTestOptions& options, SelfTest::SelfTestSuiteResult* outResult) noexcept
{
    const auto startedAt = std::chrono::steady_clock::now();
    Debug::Info(L"CompareSelfTest: begin");
    AppendCompareSelfTestTraceLine(L"Run: begin");

    SelfTest::SelfTestSuiteResult suite{};
    suite.suite = SelfTest::SelfTestSuite::CompareDirectories;

    std::wstring fatalSetupFailure;

    wil::com_ptr<IFileSystem> baseFs = GetLocalFileSystem();
    if (! baseFs)
    {
        fatalSetupFailure = L"CompareSelfTest: local file system plugin not available.";
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::CompareDirectories);
    if (fatalSetupFailure.empty() && suiteRoot.empty())
    {
        fatalSetupFailure = L"CompareSelfTest: suite artifact root not available.";
    }

    const std::filesystem::path root = suiteRoot / L"work";
    if (fatalSetupFailure.empty() && ! SelfTest::EnsureDirectory(root))
    {
        fatalSetupFailure = L"CompareSelfTest: failed to create work root folder.";
    }

    if (! fatalSetupFailure.empty())
    {
        AppendCompareSelfTestTraceLine(L"Run: aborting due to setup failure");
        SelfTest::SelfTestCaseResult setup{};
        setup.name       = L"setup";
        setup.status     = SelfTest::SelfTestCaseResult::Status::failed;
        setup.durationMs = 0;
        setup.reason     = fatalSetupFailure;
        suite.cases.push_back(std::move(setup));
        ++suite.failed;
        suite.failureMessage = fatalSetupFailure;
    }
    else
    {
        AppendCompareSelfTestTraceLine(L"Run: root created");
    }

    std::wstring guid = MakeGuidText();
    if (guid.empty())
    {
        guid = L"0";
    }

    wil::com_ptr<IFileSystem> dummyFs = GetDummyFileSystem();
    wil::com_ptr<IInformations> dummyInfo;
    wil::com_ptr<IFileSystemIO> dummyIo;
    wil::com_ptr<IFileSystemDirectoryOperations> dummyOps;

    if (fatalSetupFailure.empty())
    {
        std::wstring setupFailure;
        if (! dummyFs)
        {
            setupFailure = L"CompareSelfTest: FileSystemDummy plugin not available.";
        }
        else
        {
            AppendCompareSelfTestTraceLine(L"Run: dummy plugin setup");
            if (! CreateInformations(dummyFs, dummyInfo))
            {
                setupFailure = L"CompareSelfTest: FileSystemDummy missing IInformations.";
            }
            else
            {
                const HRESULT setHr =
                    dummyInfo->SetConfiguration("{\"maxChildrenPerDirectory\":0,\"maxDepth\":0,\"seed\":1,\"latencyMs\":0,\"virtualSpeedLimit\":\"0\"}");
                if (FAILED(setHr))
                {
                    setupFailure = L"CompareSelfTest: FileSystemDummy SetConfiguration failed.";
                }
            }

            if (setupFailure.empty() && ! CreateFileSystemIo(dummyFs, dummyIo))
            {
                setupFailure = L"CompareSelfTest: FileSystemDummy missing IFileSystemIO.";
            }
            if (setupFailure.empty() && ! CreateFileSystemDirectoryOperations(dummyFs, dummyOps))
            {
                setupFailure = L"CompareSelfTest: FileSystemDummy missing IFileSystemDirectoryOperations.";
            }
        }

        if (! setupFailure.empty())
        {
            SelfTest::SelfTestCaseResult setup{};
            setup.name       = L"setup";
            setup.status     = SelfTest::SelfTestCaseResult::Status::failed;
            setup.durationMs = 0;
            setup.reason     = setupFailure;
            suite.cases.push_back(std::move(setup));
            ++suite.failed;
            if (suite.failureMessage.empty())
            {
                suite.failureMessage = setupFailure;
            }
        }
    }

    if (fatalSetupFailure.empty())
    {
        SelfTest::RunCase(options,
                          suite,
                          L"windows_hello_cache",
                          [&](SelfTest::CaseState& state) noexcept
        {
            const auto restoreVerifier = wil::scope_exit([&] noexcept { static_cast<void>(RedSalamander::Security::SetWindowsHelloTestVerifier(nullptr)); });
            g_windowsHelloVerifierCalls.store(0u, std::memory_order_relaxed);
            RedSalamander::Security::SetWindowsHelloTestVerifier(&TestWindowsHelloVerifier);

            std::wstring id = MakeGuidText();
            state.Require(! id.empty(), L"Failed to generate a connection profile id.");

            Common::Settings::ConnectionProfile profile;
            profile.id                  = id;
            profile.name                = std::format(L"SelfTest WindowsHello {}", id);
            profile.pluginId            = std::wstring(kBuiltinFtpFileSystemId);
            profile.host                = L"example.invalid";
            profile.initialPath         = L"/";
            profile.userName            = L"user";
            profile.authMode            = Common::Settings::ConnectionAuthMode::Password;
            profile.savePassword        = true;
            profile.requireWindowsHello = true;

            const bool hadConnectionsSettings = g_settings.connections.has_value();
            if (! hadConnectionsSettings)
            {
                g_settings.connections.emplace();
            }

            const bool previousBypassHello     = g_settings.connections->bypassWindowsHello;
            const uint32_t previousTimeoutMins = g_settings.connections->windowsHelloReauthTimeoutMinute;

            g_settings.connections->bypassWindowsHello              = false;
            g_settings.connections->windowsHelloReauthTimeoutMinute = 10;
            g_settings.connections->items.push_back(profile);

            const auto restoreSettings = wil::scope_exit([&] noexcept
            {
                RedSalamander::Connections::ClearSecretAccessAuthorization(id);

                if (! g_settings.connections)
                {
                    return;
                }

                auto& items = g_settings.connections->items;
                items.erase(std::remove_if(items.begin(), items.end(), [&](const Common::Settings::ConnectionProfile& item) noexcept { return item.id == id; }),
                            items.end());

                g_settings.connections->bypassWindowsHello              = previousBypassHello;
                g_settings.connections->windowsHelloReauthTimeoutMinute = previousTimeoutMins;

                if (! hadConnectionsSettings)
                {
                    g_settings.connections.reset();
                }
            });

            const std::wstring targetName = RedSalamander::Connections::BuildCredentialTargetName(profile.id, RedSalamander::Connections::SecretKind::Password);
            state.Require(! targetName.empty(), L"Failed to build WinCred target name.");

            const std::wstring password = L"pw";
            const HRESULT saveHr        = RedSalamander::Connections::SaveGenericCredential(targetName, profile.userName, password);
            state.Require(SUCCEEDED(saveHr), std::format(L"SaveGenericCredential failed. hr=0x{:08X}", static_cast<unsigned long>(saveHr)));
            const auto deleteCredential = wil::scope_exit([&] noexcept
            {
                const HRESULT delHr = RedSalamander::Connections::DeleteGenericCredential(targetName);
                if (FAILED(delHr) && delHr != HRESULT_FROM_WIN32(ERROR_NOT_FOUND))
                {
                    Debug::Warning(L"SelfTest: DeleteGenericCredential failed target='{}' hr=0x{:08X}", targetName, static_cast<unsigned long>(delHr));
                }
            });

            RedSalamander::Connections::ClearSecretAccessAuthorization(profile.id);

            wil::com_ptr<IHostConnections> hostConnections;
            const HRESULT qiHr = GetHostServices()->QueryInterface(IID_PPV_ARGS(hostConnections.put()));
            state.Require(SUCCEEDED(qiHr) && hostConnections, std::format(L"Missing IHostConnections. hr=0x{:08X}", static_cast<unsigned long>(qiHr)));

            wil::unique_cotaskmem_string secret;
            HRESULT hr = hostConnections->GetConnectionSecret(profile.name.c_str(), HOST_CONNECTION_SECRET_PASSWORD, nullptr, secret.put());
            state.Require(SUCCEEDED(hr), std::format(L"GetConnectionSecret failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
            state.Require(secret && std::wstring_view(secret.get()) == password, L"Unexpected secret value.");
            SecureClearAndFreeSecret(secret);
            state.Require(g_windowsHelloVerifierCalls.load(std::memory_order_relaxed) == 1u, L"Expected Windows Hello to be requested once.");

            wil::unique_cotaskmem_string secret2;
            hr = hostConnections->GetConnectionSecret(profile.name.c_str(), HOST_CONNECTION_SECRET_PASSWORD, nullptr, secret2.put());
            state.Require(SUCCEEDED(hr), std::format(L"GetConnectionSecret (second call) failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
            SecureClearAndFreeSecret(secret2);
            state.Require(g_windowsHelloVerifierCalls.load(std::memory_order_relaxed) == 1u, L"Expected Windows Hello to be cached (no second prompt).");

#ifdef _DEBUG
            // Simulate an expired authorization timestamp to ensure long-running operations (compare/copy) won't re-prompt.
            RedSalamander::Connections::SetSecretAccessAuthorizationTickForTesting(profile.id, 0);

            wil::unique_cotaskmem_string secretExpired;
            hr = hostConnections->GetConnectionSecret(profile.name.c_str(), HOST_CONNECTION_SECRET_PASSWORD, nullptr, secretExpired.put());
            state.Require(SUCCEEDED(hr), std::format(L"GetConnectionSecret (expired auth) failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
            SecureClearAndFreeSecret(secretExpired);
            state.Require(g_windowsHelloVerifierCalls.load(std::memory_order_relaxed) == 1u, L"Expected Windows Hello not to re-prompt after session auth.");
#endif

            RedSalamander::Connections::ClearSecretAccessAuthorization(profile.id);
            RedSalamander::Connections::NoteSecretAccessAuthorized(profile.id);
            g_windowsHelloVerifierCalls.store(0u, std::memory_order_relaxed);

            wil::unique_cotaskmem_string secret3;
            hr = hostConnections->GetConnectionSecret(profile.name.c_str(), HOST_CONNECTION_SECRET_PASSWORD, nullptr, secret3.put());
            state.Require(SUCCEEDED(hr), std::format(L"GetConnectionSecret (manual auth) failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
            SecureClearAndFreeSecret(secret3);
            state.Require(g_windowsHelloVerifierCalls.load(std::memory_order_relaxed) == 0u, L"Manual secret entry should suppress Windows Hello prompts.");

            return state.failure.empty();
        });

        SelfTest::RunCase(options,
                          suite,
                          L"unique",
                          [&](SelfTest::CaseState& state) noexcept
        {
            // Case: Unique files/dirs selected; identical excluded by default.
            if (const auto foldersOpt = CreateCaseFolders(root, L"unique"))
            {
                const auto& folders = *foldersOpt;
                state.Require(SelfTest::WriteTextFile(folders.left / L"only_left.txt", "L"), L"Failed to create only_left.txt (left).");
                state.Require(SelfTest::WriteTextFile(folders.right / L"only_right.txt", "R"), L"Failed to create only_right.txt (right).");
                state.Require(SelfTest::EnsureDirectory(folders.left / L"only_left_dir"), L"Failed to create only_left_dir (left).");
                state.Require(SelfTest::WriteTextFile(folders.left / L"same.txt", "S"), L"Failed to create same.txt (left).");
                state.Require(SelfTest::WriteTextFile(folders.right / L"same.txt", "S"), L"Failed to create same.txt (right).");

                AppendCompareSelfTestTraceLine(L"Case: unique: computing decision");
                auto decision = ComputeRootDecision(baseFs, folders, Common::Settings::CompareDirectoriesSettings{}, state);
                AppendCompareSelfTestTraceLine(L"Case: unique: decision returned");
                if (decision)
                {
                    {
                        const auto* item = FindItem(*decision, L"only_left.txt");
                        state.Require(item != nullptr, L"only_left.txt missing from decision.");
                        if (item)
                        {
                            state.Require(item->isDifferent, L"only_left.txt expected isDifferent.");
                            state.Require(item->selectLeft && ! item->selectRight, L"only_left.txt expected selectLeft only.");
                            state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::OnlyInLeft),
                                          L"only_left.txt expected differenceMask=OnlyInLeft.");
                        }
                    }
                    {
                        const auto* item = FindItem(*decision, L"only_right.txt");
                        state.Require(item != nullptr, L"only_right.txt missing from decision.");
                        if (item)
                        {
                            state.Require(item->isDifferent, L"only_right.txt expected isDifferent.");
                            state.Require(! item->selectLeft && item->selectRight, L"only_right.txt expected selectRight only.");
                            state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::OnlyInRight),
                                          L"only_right.txt expected differenceMask=OnlyInRight.");
                        }
                    }
                    {
                        const auto* item = FindItem(*decision, L"only_left_dir");
                        state.Require(item != nullptr, L"only_left_dir missing from decision.");
                        if (item)
                        {
                            state.Require(item->isDirectory, L"only_left_dir expected isDirectory.");
                            state.Require(item->isDifferent, L"only_left_dir expected isDifferent.");
                            state.Require(item->selectLeft && ! item->selectRight, L"only_left_dir expected selectLeft only.");
                            state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::OnlyInLeft),
                                          L"only_left_dir expected differenceMask=OnlyInLeft.");
                        }
                    }
                    {
                        const auto* item = FindItem(*decision, L"same.txt");
                        state.Require(item == nullptr, L"same.txt expected elided from decision in differences-only mode.");
                    }

                    auto session = std::make_shared<CompareDirectoriesSession>(
                        baseFs, baseFs, folders.left, folders.right, Common::Settings::CompareDirectoriesSettings{});
                    const auto fsLeft  = CreateCompareDirectoriesFileSystem(ComparePane::Left, session);
                    const auto fsRight = CreateCompareDirectoriesFileSystem(ComparePane::Right, session);

                    const auto scanTimeout = std::chrono::milliseconds(SelfTest::ScaleTimeout(5'000));
                    state.Require(StartScanAndWaitForIdle(session, scanTimeout), L"unique: scan did not go idle.");
                    state.Require(DrainPendingSubdirUpdates(session, 256), L"unique: subdir updates did not drain.");

                    const auto leftNames  = EnumerateDirectoryNames(fsLeft, folders.left, state);
                    const auto rightNames = EnumerateDirectoryNames(fsRight, folders.right, state);
                    AppendCompareSelfTestTraceLine(L"Case: unique: enumeration done");

                    state.Require(ContainsName(leftNames, L"only_left.txt"), L"only_left.txt expected in left enumeration.");
                    state.Require(! ContainsName(leftNames, L"only_right.txt"), L"only_right.txt unexpected in left enumeration.");
                    state.Require(! ContainsName(leftNames, L"same.txt"), L"same.txt expected excluded in left enumeration.");

                    state.Require(ContainsName(rightNames, L"only_right.txt"), L"only_right.txt expected in right enumeration.");
                    state.Require(! ContainsName(rightNames, L"only_left.txt"), L"only_left.txt unexpected in right enumeration.");
                    state.Require(! ContainsName(rightNames, L"same.txt"), L"same.txt expected excluded in right enumeration.");

                    AppendCompareSelfTestTraceLine(L"Case: unique: done");
                }
            }
            else
            {
                state.Require(false, L"Failed to create case folders: unique.");
            }

            return state.failure.empty();
        });

        SelfTest::RunCase(options,
                          suite,
                          L"typemismatch",
                          [&](SelfTest::CaseState& state) noexcept
        {
            // Case: File vs directory mismatch selects both sides.
            if (const auto foldersOpt = CreateCaseFolders(root, L"typemismatch"))
            {
                const auto& folders = *foldersOpt;
                state.Require(SelfTest::WriteTextFile(folders.left / L"mix", "F"), L"Failed to create mix file (left).");
                state.Require(SelfTest::EnsureDirectory(folders.right / L"mix"), L"Failed to create mix directory (right).");

                auto decision = ComputeRootDecision(baseFs, folders, Common::Settings::CompareDirectoriesSettings{}, state);
                if (decision)
                {
                    const auto* item = FindItem(*decision, L"mix");
                    state.Require(item != nullptr, L"mix missing from decision.");
                    if (item)
                    {
                        state.Require(item->isDifferent, L"mix expected isDifferent on type mismatch.");
                        state.Require(item->selectLeft && item->selectRight, L"mix expected select both on type mismatch.");
                        state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::TypeMismatch), L"mix expected differenceMask=TypeMismatch.");
                    }
                }
            }
            else
            {
                state.Require(false, L"Failed to create case folders: typemismatch.");
            }

            return state.failure.empty();
        });

        SelfTest::RunCase(options,
                          suite,
                          L"size",
                          [&](SelfTest::CaseState& state) noexcept
        {
            // Case: Size compare selects bigger file.
            if (const auto foldersOpt = CreateCaseFolders(root, L"size"))
            {
                const auto& folders = *foldersOpt;
                state.Require(WriteFileFill(folders.left / L"a.bin", 'A', 200), L"Failed to create a.bin (left).");
                state.Require(WriteFileFill(folders.right / L"a.bin", 'B', 100), L"Failed to create a.bin (right).");

                Common::Settings::CompareDirectoriesSettings settings{};
                settings.compareSize = true;

                auto decision = ComputeRootDecision(baseFs, folders, settings, state);
                if (decision)
                {
                    const auto* item = FindItem(*decision, L"a.bin");
                    state.Require(item != nullptr, L"a.bin missing from decision.");
                    if (item)
                    {
                        state.Require(item->isDifferent, L"a.bin expected isDifferent with compareSize.");
                        state.Require(item->selectLeft && ! item->selectRight, L"a.bin expected selectLeft only when left is bigger.");
                        state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::Size), L"a.bin expected differenceMask=Size.");
                    }
                }
            }
            else
            {
                state.Require(false, L"Failed to create case folders: size.");
            }

            return state.failure.empty();
        });

        SelfTest::RunCase(options,
                          suite,
                          L"time",
                          [&](SelfTest::CaseState& state) noexcept
        {
            // Case: Date/time compare selects newer file.
            if (const auto foldersOpt = CreateCaseFolders(root, L"time"))
            {
                const auto& folders = *foldersOpt;
                state.Require(SelfTest::WriteTextFile(folders.left / L"a.txt", "T"), L"Failed to create a.txt (left).");
                state.Require(SelfTest::WriteTextFile(folders.right / L"a.txt", "T"), L"Failed to create a.txt (right).");

                FILETIME now{};
                ::GetSystemTimeAsFileTime(&now);
                ULARGE_INTEGER newer{};
                newer.LowPart  = now.dwLowDateTime;
                newer.HighPart = now.dwHighDateTime;
                newer.QuadPart += 60ull * 10'000'000ull;

                FILETIME leftFt{};
                leftFt.dwLowDateTime  = newer.LowPart;
                leftFt.dwHighDateTime = newer.HighPart;

                state.Require(SetFileLastWriteTime(folders.left / L"a.txt", leftFt), L"Failed to set a.txt last write time (left).");
                state.Require(SetFileLastWriteTime(folders.right / L"a.txt", now), L"Failed to set a.txt last write time (right).");

                Common::Settings::CompareDirectoriesSettings settings{};
                settings.compareDateTime = true;

                auto decision = ComputeRootDecision(baseFs, folders, settings, state);
                if (decision)
                {
                    const auto* item = FindItem(*decision, L"a.txt");
                    state.Require(item != nullptr, L"a.txt missing from decision.");
                    if (item)
                    {
                        state.Require(item->isDifferent, L"a.txt expected isDifferent with compareDateTime.");
                        state.Require(item->selectLeft && ! item->selectRight, L"a.txt expected selectLeft only when left is newer.");
                        state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::DateTime), L"a.txt expected differenceMask=DateTime.");
                    }
                }
            }
            else
            {
                state.Require(false, L"Failed to create case folders: time.");
            }

            return state.failure.empty();
        });

        SelfTest::RunCase(options,
                          suite,
                          L"attributes",
                          [&](SelfTest::CaseState& state) noexcept
        {
            // Case: Attribute compare selects both sides.
            if (const auto foldersOpt = CreateCaseFolders(root, L"attributes"))
            {
                const auto& folders = *foldersOpt;
                state.Require(SelfTest::WriteTextFile(folders.left / L"a.txt", "A"), L"Failed to create a.txt (left).");
                state.Require(SelfTest::WriteTextFile(folders.right / L"a.txt", "A"), L"Failed to create a.txt (right).");

                const std::filesystem::path leftPath = folders.left / L"a.txt";
                const DWORD leftAttrs                = ::GetFileAttributesW(leftPath.c_str());
                state.Require(leftAttrs != INVALID_FILE_ATTRIBUTES, L"GetFileAttributesW failed for a.txt (left).");
                if (leftAttrs != INVALID_FILE_ATTRIBUTES)
                {
                    state.Require(::SetFileAttributesW(leftPath.c_str(), leftAttrs | FILE_ATTRIBUTE_HIDDEN) != 0,
                                  L"SetFileAttributesW failed for a.txt (left).");
                }

                Common::Settings::CompareDirectoriesSettings settings{};
                settings.compareAttributes = true;

                auto decision = ComputeRootDecision(baseFs, folders, settings, state);
                if (decision)
                {
                    const auto* item = FindItem(*decision, L"a.txt");
                    state.Require(item != nullptr, L"a.txt missing from decision.");
                    if (item)
                    {
                        state.Require(item->isDifferent, L"a.txt expected isDifferent with compareAttributes.");
                        state.Require(item->selectLeft && item->selectRight, L"a.txt expected select both when attributes differ.");
                        state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::Attributes), L"a.txt expected differenceMask=Attributes.");
                    }
                }
            }
            else
            {
                state.Require(false, L"Failed to create case folders: attributes.");
            }

            return state.failure.empty();
        });

        SelfTest::RunCase(options,
                          suite,
                          L"content",
                          [&](SelfTest::CaseState& state) noexcept
        {
            // Case: Content compare selects both sides.
            if (const auto foldersOpt = CreateCaseFolders(root, L"content"))
            {
                const auto& folders = *foldersOpt;
                state.Require(WriteFileFill(folders.left / L"a.bin", 'X', 64), L"Failed to create a.bin (left).");
                state.Require(WriteFileFill(folders.right / L"a.bin", 'Y', 64), L"Failed to create a.bin (right).");

                Common::Settings::CompareDirectoriesSettings settings{};
                settings.compareContent = true;

                auto session  = std::make_shared<CompareDirectoriesSession>(baseFs, baseFs, folders.left, folders.right, settings);
                auto decision = WaitForContentCompare(session, std::filesystem::path{}, L"a.bin", state);
                if (decision)
                {
                    const auto* item = FindItem(*decision, L"a.bin");
                    state.Require(item != nullptr, L"a.bin missing from decision.");
                    if (item)
                    {
                        state.Require(item->isDifferent, L"a.bin expected isDifferent with compareContent.");
                        state.Require(item->selectLeft && item->selectRight, L"a.bin expected select both when content differs.");
                        state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::Content), L"a.bin expected differenceMask=Content.");
                        state.Require(! HasFlag(item->differenceMask, CompareDirectoriesDiffBit::ContentPending),
                                      L"a.bin expected ContentPending cleared after compare completes.");
                    }
                }
            }
            else
            {
                state.Require(false, L"Failed to create case folders: content.");
            }

            return state.failure.empty();
        });

        SelfTest::RunCase(options,
                          suite,
                          L"content_dual_io",
                          [&](SelfTest::CaseState& state) noexcept
        {
            // Case: Content compare uses the correct per-pane IFileSystemIO (dual-IO regression guard).
            if (const auto foldersOpt = CreateCaseFolders(root, L"content_dual_io"))
            {
                const auto& folders = *foldersOpt;
                state.Require(SelfTest::WriteTextFile(folders.left / L"a.txt", "AAAA"), L"Failed to create a.txt (left).");
                state.Require(SelfTest::WriteTextFile(folders.right / L"a.txt", "BBBB"), L"Failed to create a.txt (right).");

                const wil::com_ptr<IFileSystem> leftFs  = CreatePluginPathMappedRootFileSystem(baseFs, folders.left);
                const wil::com_ptr<IFileSystem> rightFs = CreatePluginPathMappedRootFileSystem(baseFs, folders.right);
                state.Require(static_cast<bool>(leftFs), L"Failed to create left mapped filesystem.");
                state.Require(static_cast<bool>(rightFs), L"Failed to create right mapped filesystem.");

                Common::Settings::CompareDirectoriesSettings settings{};
                settings.compareContent = true;

                const std::filesystem::path pluginRoot(L"/");
                auto session = std::make_shared<CompareDirectoriesSession>(leftFs, rightFs, pluginRoot, pluginRoot, settings);
                state.Require(session->IsContentCompareSupported(), L"content_dual_io: content compare should be supported.");

                auto decision = WaitForContentCompare(session, std::filesystem::path{}, L"a.txt", state);
                if (decision)
                {
                    const auto* item = FindItem(*decision, L"a.txt");
                    state.Require(item != nullptr, L"a.txt missing from decision.");
                    if (item)
                    {
                        state.Require(item->isDifferent, L"a.txt expected isDifferent with compareContent.");
                        state.Require(item->selectLeft && item->selectRight, L"a.txt expected select both when content differs.");
                        state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::Content), L"a.txt expected differenceMask=Content.");
                        state.Require(! HasFlag(item->differenceMask, CompareDirectoriesDiffBit::ContentPending),
                                      L"a.txt expected ContentPending cleared after compare completes.");
                    }
                }
            }
            else
            {
                state.Require(false, L"Failed to create case folders: content_dual_io.");
            }

            return state.failure.empty();
        });

        SelfTest::RunCase(options,
                          suite,
                          L"content_no_io_disables_compareContent",
                          [&](SelfTest::CaseState& state) noexcept
        {
            // Case: compareContent is treated as disabled when either side lacks IFileSystemIO.
            if (const auto foldersOpt = CreateCaseFolders(root, L"content_no_io_disables_compareContent"))
            {
                const auto& folders = *foldersOpt;
                state.Require(SelfTest::WriteTextFile(folders.left / L"a.txt", "AAAA"), L"Failed to create a.txt (left).");
                state.Require(SelfTest::WriteTextFile(folders.right / L"a.txt", "BBBB"), L"Failed to create a.txt (right).");

                const wil::com_ptr<IFileSystem> leftFs  = CreatePluginPathMappedRootFileSystem(baseFs, folders.left);
                const wil::com_ptr<IFileSystem> rightFs = CreatePluginPathMappedRootFileSystemNoIO(baseFs, folders.right);
                state.Require(static_cast<bool>(leftFs), L"Failed to create left mapped filesystem.");
                state.Require(static_cast<bool>(rightFs), L"Failed to create right mapped filesystem (no IO).");

                Common::Settings::CompareDirectoriesSettings settings{};
                settings.compareContent     = true;
                settings.keepIdenticalItems = true;

                const std::filesystem::path pluginRoot(L"/");
                const auto session = std::make_shared<CompareDirectoriesSession>(leftFs, rightFs, pluginRoot, pluginRoot, settings);
                state.Require(! session->IsContentCompareSupported(), L"content_no_io: content compare should be unsupported.");

                const auto decision = session->GetOrComputeDecision(std::filesystem::path{});
                state.Require(static_cast<bool>(decision), L"content_no_io: decision is null.");
                if (decision)
                {
                    state.Require(SUCCEEDED(decision->hr), L"content_no_io: decision hr is failure.");

                    const auto* item = FindItem(*decision, L"a.txt");
                    state.Require(item != nullptr, L"content_no_io: a.txt missing from decision.");
                    if (item)
                    {
                        state.Require(! item->isDifferent, L"content_no_io: a.txt expected identical (content criterion disabled).");
                        state.Require(item->differenceMask == 0u, L"content_no_io: a.txt expected differenceMask=0.");
                        state.Require(! HasFlag(item->differenceMask, CompareDirectoriesDiffBit::ContentPending),
                                      L"content_no_io: a.txt expected no ContentPending when unsupported.");
                    }
                }
            }
            else
            {
                state.Require(false, L"Failed to create case folders: content_no_io_disables_compareContent.");
            }

            return state.failure.empty();
        });

        SelfTest::RunCase(options,
                          suite,
                          L"content_size_mismatch_no_pending",
                          [&](SelfTest::CaseState& state) noexcept
        {
            // Case: Content compare with different sizes does not mark ContentPending.
            if (const auto foldersOpt = CreateCaseFolders(root, L"content_size_mismatch_no_pending"))
            {
                const auto& folders = *foldersOpt;
                state.Require(WriteFileFill(folders.left / L"a.bin", 'X', 64), L"Failed to create a.bin (left).");
                state.Require(WriteFileFill(folders.right / L"a.bin", 'X', 32), L"Failed to create a.bin (right).");

                Common::Settings::CompareDirectoriesSettings settings{};
                settings.compareContent = true;

                auto decision = ComputeRootDecision(baseFs, folders, settings, state);
                if (decision)
                {
                    const auto* item = FindItem(*decision, L"a.bin");
                    state.Require(item != nullptr, L"a.bin missing from decision.");
                    if (item)
                    {
                        state.Require(item->isDifferent, L"a.bin expected isDifferent with compareContent and size mismatch.");
                        state.Require(item->selectLeft && item->selectRight, L"a.bin expected select both when compareContent and sizes differ.");
                        state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::Content), L"a.bin expected differenceMask=Content.");
                        state.Require(! HasFlag(item->differenceMask, CompareDirectoriesDiffBit::ContentPending),
                                      L"a.bin expected ContentPending not set when sizes differ.");
                    }
                }
            }
            else
            {
                state.Require(false, L"Failed to create case folders: content_size_mismatch_no_pending.");
            }

            return state.failure.empty();
        });

        SelfTest::RunCase(options,
                          suite,
                          L"zero_vs_nonzero_content",
                          [&](SelfTest::CaseState& state) noexcept
        {
            // Case: Content compare with zero vs non-zero size does not mark ContentPending.
            if (const auto foldersOpt = CreateCaseFolders(root, L"zero_vs_nonzero_content"))
            {
                const auto& folders = *foldersOpt;
                state.Require(SelfTest::WriteBinaryFile(folders.left / L"a.bin", {}), L"Failed to create a.bin (left).");
                state.Require(WriteFileFill(folders.right / L"a.bin", 'Z', 1), L"Failed to create a.bin (right).");

                Common::Settings::CompareDirectoriesSettings settings{};
                settings.compareContent = true;

                auto decision = ComputeRootDecision(baseFs, folders, settings, state);
                if (decision)
                {
                    const auto* item = FindItem(*decision, L"a.bin");
                    state.Require(item != nullptr, L"a.bin missing from decision.");
                    if (item)
                    {
                        state.Require(item->isDifferent, L"a.bin expected isDifferent with compareContent and size mismatch.");
                        state.Require(item->selectLeft && item->selectRight, L"a.bin expected select both when compareContent and sizes differ.");
                        state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::Content), L"a.bin expected differenceMask=Content.");
                        state.Require(! HasFlag(item->differenceMask, CompareDirectoriesDiffBit::ContentPending),
                                      L"a.bin expected ContentPending not set when sizes differ.");
                    }
                }
            }
            else
            {
                state.Require(false, L"Failed to create case folders: zero_vs_nonzero_content.");
            }

            return state.failure.empty();
        });

        SelfTest::RunCase(options,
                          suite,
                          L"unicode_filenames",
                          [&](SelfTest::CaseState& state) noexcept
        {
            // Case: Unicode filenames (CJK + emoji) appear in decisions and content compare works.
            if (const auto foldersOpt = CreateCaseFolders(root, L"unicode_filenames"))
            {
                const auto& folders = *foldersOpt;
                state.Require(SelfTest::WriteTextFile(folders.left / L"\u3053\u3093\u306B\u3061\u306F.txt", "JP"), L"Failed to create こんにちは.txt (left).");
                state.Require(SelfTest::WriteTextFile(folders.right / L"\u3053\u3093\u306B\u3061\u306F.txt", "JP"),
                              L"Failed to create こんにちは.txt (right).");

                state.Require(WriteFileFill(folders.left / L"emoji_\U0001F600.bin", 'A', 64), L"Failed to create emoji file (left).");
                state.Require(WriteFileFill(folders.right / L"emoji_\U0001F600.bin", 'B', 64), L"Failed to create emoji file (right).");

                Common::Settings::CompareDirectoriesSettings settings{};
                settings.compareContent     = true;
                settings.keepIdenticalItems = true;

                auto session  = std::make_shared<CompareDirectoriesSession>(baseFs, baseFs, folders.left, folders.right, settings);
                auto decision = WaitForContentCompare(session, std::filesystem::path{}, L"emoji_\U0001F600.bin", state);
                if (decision)
                {
                    {
                        const auto* item = FindItem(*decision, L"\u3053\u3093\u306B\u3061\u306F.txt");
                        state.Require(item != nullptr, L"Unicode file missing from decision.");
                        if (item)
                        {
                            state.Require(! item->isDifferent, L"Unicode identical file expected not different.");
                        }
                    }
                    {
                        const auto* item = FindItem(*decision, L"emoji_\U0001F600.bin");
                        state.Require(item != nullptr, L"Emoji file missing from decision.");
                        if (item)
                        {
                            state.Require(item->isDifferent, L"Emoji file expected different with compareContent.");
                            state.Require(item->selectLeft && item->selectRight, L"Emoji file expected select both when content differs.");
                            state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::Content), L"Emoji file expected differenceMask=Content.");
                            state.Require(! HasFlag(item->differenceMask, CompareDirectoriesDiffBit::ContentPending),
                                          L"Emoji file expected ContentPending cleared after compare completes.");
                        }
                    }
                }
            }
            else
            {
                state.Require(false, L"Failed to create case folders: unicode_filenames.");
            }

            return state.failure.empty();
        });

        SelfTest::RunCase(options,
                          suite,
                          L"content short reads",
                          [&](SelfTest::CaseState& state) noexcept
        {
            // Case: Content compare tolerates short reads for equal files.
            if (const auto foldersOpt = CreateCaseFolders(root, L"content_shortreads"))
            {
                const auto& folders = *foldersOpt;
                state.Require(WriteFileFill(folders.left / L"a.bin", 'Z', 4096), L"Failed to create a.bin (left).");
                state.Require(WriteFileFill(folders.right / L"a.bin", 'Z', 4096), L"Failed to create a.bin (right).");

                Common::Settings::CompareDirectoriesSettings settings{};
                settings.compareContent     = true;
                settings.keepIdenticalItems = true;

                wil::com_ptr<IFileSystem> wrapped = CreateShortReadFileSystem(baseFs, folders.left, 1u, 0u);
                state.Require(static_cast<bool>(wrapped), L"Failed to create short-read file system wrapper.");

                const wil::com_ptr<IFileSystem> compareFs = wrapped ? wrapped : baseFs;
                auto session  = std::make_shared<CompareDirectoriesSession>(compareFs, compareFs, folders.left, folders.right, settings);
                auto decision = WaitForContentCompare(session, std::filesystem::path{}, L"a.bin", state);
                if (decision)
                {
                    const auto* item = FindItem(*decision, L"a.bin");
                    state.Require(item != nullptr, L"a.bin missing from decision.");
                    if (item)
                    {
                        state.Require(! item->isDifferent, L"a.bin expected not different for equal content with short reads.");
                        state.Require(! HasFlag(item->differenceMask, CompareDirectoriesDiffBit::Content),
                                      L"a.bin expected Content bit cleared for equal content with short reads.");
                        state.Require(! HasFlag(item->differenceMask, CompareDirectoriesDiffBit::ContentPending),
                                      L"a.bin expected ContentPending cleared after compare completes (short reads).");
                        state.Require(! item->selectLeft && ! item->selectRight, L"a.bin expected no selection when equal.");
                    }
                }
            }
            else
            {
                state.Require(false, L"Failed to create case folders: content_shortreads.");
            }

            return state.failure.empty();
        });

        SelfTest::RunCase(options,
                          suite,
                          L"subdir pending",
                          [&](SelfTest::CaseState& state) noexcept
        {
            // Case: Subdirectory pending state + flush updates ancestors without navigation.
            if (const auto foldersOpt = CreateCaseFolders(root, L"subdir_pending"))
            {
                const auto& folders = *foldersOpt;
                state.Require(SelfTest::EnsureDirectory(folders.left / L"sub"), L"Failed to create sub (left).");
                state.Require(SelfTest::EnsureDirectory(folders.right / L"sub"), L"Failed to create sub (right).");
                state.Require(WriteFileFill(folders.left / L"sub" / L"a.bin", 'A', 512 * 1024), L"Failed to create sub\\a.bin (left).");
                state.Require(WriteFileFill(folders.right / L"sub" / L"a.bin", 'A', 512 * 1024), L"Failed to create sub\\a.bin (right).");

                Common::Settings::CompareDirectoriesSettings settings{};
                settings.compareContent        = true;
                settings.compareSubdirectories = true;
                settings.keepIdenticalItems    = true;

                wil::com_ptr<IFileSystem> wrapped = CreateShortReadFileSystem(baseFs, folders.left, 1024u, 1u);
                state.Require(static_cast<bool>(wrapped), L"Failed to create short-read file system wrapper (subdir pending).");

                const wil::com_ptr<IFileSystem> compareFs = wrapped ? wrapped : baseFs;
                auto session = std::make_shared<CompareDirectoriesSession>(compareFs, compareFs, folders.left, folders.right, settings);

                std::mutex progressMutex;
                std::condition_variable progressCv;
                bool contentDone = false;

                session->SetContentProgressCallback([&](uint32_t,
                                                        const std::filesystem::path&,
                                                        std::wstring_view,
                                                        uint64_t,
                                                        uint64_t,
                                                        uint64_t,
                                                        uint64_t,
                                                        uint64_t pendingContentCompares,
                                                        uint64_t totalContentCompares,
                                                        uint64_t completedContentCompares) noexcept
                {
                    if (pendingContentCompares != 0u || totalContentCompares == 0u || completedContentCompares != totalContentCompares)
                    {
                        return;
                    }

                    std::lock_guard lock(progressMutex);
                    contentDone = true;
                    progressCv.notify_all();
                });

                auto rootDecision = session->GetOrComputeDecision(std::filesystem::path{});
                state.Require(static_cast<bool>(rootDecision), L"subdir pending: root decision is null.");
                if (rootDecision)
                {
                    const auto* subItem = FindItem(*rootDecision, L"sub");
                    state.Require(subItem != nullptr, L"subdir pending: sub missing from root decision.");
                    if (subItem)
                    {
                        state.Require(subItem->isDirectory, L"subdir pending: sub expected isDirectory.");
                        state.Require(HasFlag(subItem->differenceMask, CompareDirectoriesDiffBit::SubdirPending),
                                      L"subdir pending: sub expected SubdirPending while content compare is running.");
                        state.Require(! HasFlag(subItem->differenceMask, CompareDirectoriesDiffBit::SubdirContent),
                                      L"subdir pending: sub expected no SubdirContent while only content compares are pending.");
                        state.Require(! subItem->isDifferent, L"subdir pending: sub expected not different while pending.");
                        state.Require(! subItem->selectLeft && ! subItem->selectRight, L"subdir pending: sub expected not selected while pending.");
                    }
                }

                auto subDecision = session->GetOrComputeDecision(std::filesystem::path(L"sub"));
                state.Require(static_cast<bool>(subDecision), L"subdir pending: sub decision is null.");
                if (subDecision)
                {
                    const auto* fileItem = FindItem(*subDecision, L"a.bin");
                    state.Require(fileItem != nullptr, L"subdir pending: a.bin missing from sub decision.");
                    if (fileItem)
                    {
                        state.Require(HasFlag(fileItem->differenceMask, CompareDirectoriesDiffBit::ContentPending),
                                      L"subdir pending: a.bin expected ContentPending while content compare is running.");
                        state.Require(! HasFlag(fileItem->differenceMask, CompareDirectoriesDiffBit::Content),
                                      L"subdir pending: a.bin expected no Content bit while pending.");
                        state.Require(! fileItem->isDifferent, L"subdir pending: a.bin expected not different while pending.");
                        state.Require(! fileItem->selectLeft && ! fileItem->selectRight, L"subdir pending: a.bin expected not selected while pending.");
                    }
                }

                {
                    std::unique_lock lock(progressMutex);
                    static_cast<void>(progressCv.wait_for(lock, std::chrono::milliseconds(SelfTest::ScaleTimeout(30'000)), [&] { return contentDone; }));
                }
                state.Require(contentDone, L"subdir pending: timed out waiting for content compare to finish.");

                // Root decision remains in pending state until pending updates are flushed.
                auto rootBeforeFlush = session->GetOrComputeDecision(std::filesystem::path{});
                if (rootBeforeFlush)
                {
                    const auto* subItem = FindItem(*rootBeforeFlush, L"sub");
                    if (subItem)
                    {
                        state.Require(HasFlag(subItem->differenceMask, CompareDirectoriesDiffBit::SubdirPending),
                                      L"subdir pending: expected SubdirPending to remain until FlushPendingContentCompareUpdates.");
                    }
                }

                session->FlushPendingContentCompareUpdates();
                session->SetContentProgressCallback({});

                auto rootAfterFlush = session->GetOrComputeDecision(std::filesystem::path{});
                state.Require(static_cast<bool>(rootAfterFlush), L"subdir pending: root decision missing after flush.");
                if (rootAfterFlush)
                {
                    const auto* subItem = FindItem(*rootAfterFlush, L"sub");
                    state.Require(subItem != nullptr, L"subdir pending: sub missing after flush.");
                    if (subItem)
                    {
                        state.Require(subItem->differenceMask == 0u, L"subdir pending: sub expected no difference mask after flush (equal subtree).");
                        state.Require(! subItem->isDifferent, L"subdir pending: sub expected not different after flush (equal subtree).");
                        state.Require(! subItem->selectLeft && ! subItem->selectRight,
                                      L"subdir pending: sub expected not selected after flush (equal subtree).");
                    }
                }

                auto subAfterFlush = session->GetOrComputeDecision(std::filesystem::path(L"sub"));
                if (subAfterFlush)
                {
                    const auto* fileItem = FindItem(*subAfterFlush, L"a.bin");
                    state.Require(fileItem != nullptr, L"subdir pending: a.bin missing after flush.");
                    if (fileItem)
                    {
                        state.Require(fileItem->differenceMask == 0u, L"subdir pending: a.bin expected no difference mask after flush (equal).");
                        state.Require(! fileItem->isDifferent, L"subdir pending: a.bin expected not different after flush (equal).");
                        state.Require(! fileItem->selectLeft && ! fileItem->selectRight, L"subdir pending: a.bin expected not selected after flush (equal).");
                    }
                }
            }
            else
            {
                state.Require(false, L"Failed to create case folders: subdir_pending.");
            }

            return state.failure.empty();
        });

        SelfTest::RunCase(options,
                          suite,
                          L"subdirs",
                          [&](SelfTest::CaseState& state) noexcept
        {
            // Case: Subdirectory content compare selects both directories.
            if (const auto foldersOpt = CreateCaseFolders(root, L"subdirs"))
            {
                const auto& folders = *foldersOpt;
                state.Require(SelfTest::EnsureDirectory(folders.left / L"sub"), L"Failed to create sub (left).");
                state.Require(SelfTest::EnsureDirectory(folders.right / L"sub"), L"Failed to create sub (right).");
                state.Require(SelfTest::WriteTextFile(folders.left / L"sub" / L"child.txt", "C"), L"Failed to create sub\\child.txt (left).");

                Common::Settings::CompareDirectoriesSettings settings{};
                settings.compareSubdirectories = true;

                auto session = std::make_shared<CompareDirectoriesSession>(baseFs, baseFs, folders.left, folders.right, settings);
                state.Require(static_cast<bool>(session), L"subdirs: failed to create session.");
                state.Require(StartScanAndWaitForIdle(session, std::chrono::milliseconds{SelfTest::ScaleTimeout(10'000)}),
                              L"subdirs: scan did not become idle within timeout.");
                state.Require(DrainPendingSubdirUpdates(session, 256), L"subdirs: failed to drain pending subtree updates.");

                auto decision = session->GetOrComputeDecision(std::filesystem::path{});
                if (decision)
                {
                    const auto* item = FindItem(*decision, L"sub");
                    state.Require(item != nullptr, L"sub missing from decision.");
                    if (item)
                    {
                        state.Require(item->isDirectory, L"sub expected isDirectory.");
                        state.Require(item->isDifferent, L"sub expected isDifferent with compareSubdirectories.");
                        state.Require(item->selectLeft && item->selectRight, L"sub expected select both when content differs.");
                        state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::SubdirContent), L"sub expected differenceMask=SubdirContent.");
                    }
                }
            }
            else
            {
                state.Require(false, L"Failed to create case folders: subdirs.");
            }

            return state.failure.empty();
        });

        SelfTest::RunCase(options,
                          suite,
                          L"no_sync_deep_scan",
                          [&](SelfTest::CaseState& state) noexcept
        {
            // Case: GetOrComputeDecision() must not perform a synchronous deep subtree traversal.
            if (const auto foldersOpt = CreateCaseFolders(root, L"no_sync_deep_scan"))
            {
                const auto& folders = *foldersOpt;

                const std::filesystem::path leftRoot  = folders.left;
                const std::filesystem::path rightRoot = folders.right;

                state.Require(SelfTest::EnsureDirectory(leftRoot / L"sub" / L"sub2"), L"no_sync_deep_scan: failed to create sub tree (left).");
                state.Require(SelfTest::EnsureDirectory(rightRoot / L"sub" / L"sub2"), L"no_sync_deep_scan: failed to create sub tree (right).");
                state.Require(SelfTest::WriteTextFile(leftRoot / L"sub" / L"sub2" / L"leaf.txt", "L"), L"no_sync_deep_scan: failed to create leaf.txt (left).");

                std::atomic_uint32_t readDirCalls{0};

                const wil::com_ptr<IFileSystem> leftFs  = CreateCountingReadDirectoryFileSystem(baseFs, &readDirCalls);
                const wil::com_ptr<IFileSystem> rightFs = CreateCountingReadDirectoryFileSystem(baseFs, &readDirCalls);
                state.Require(static_cast<bool>(leftFs), L"no_sync_deep_scan: failed to create left counting fs.");
                state.Require(static_cast<bool>(rightFs), L"no_sync_deep_scan: failed to create right counting fs.");

                Common::Settings::CompareDirectoriesSettings settings{};
                settings.compareSubdirectories = true;

                const uint32_t before = readDirCalls.load(std::memory_order_acquire);
                auto session          = std::make_shared<CompareDirectoriesSession>(leftFs, rightFs, leftRoot, rightRoot, settings);
                auto decision         = session->GetOrComputeDecision(std::filesystem::path{});
                state.Require(static_cast<bool>(decision), L"no_sync_deep_scan: root decision is null.");
                const uint32_t after = readDirCalls.load(std::memory_order_acquire);

                state.Require((after - before) == 2u,
                              std::format(L"no_sync_deep_scan: expected exactly 2 ReadDirectoryInfo calls (root left+right), got {}.", (after - before)));
            }
            else
            {
                state.Require(false, L"Failed to create case folders: no_sync_deep_scan.");
            }

            return state.failure.empty();
        });

        SelfTest::RunCase(options,
                          suite,
                          L"subdirattrs",
                          [&](SelfTest::CaseState& state) noexcept
        {
            // Case: Compare attributes of subdirectories selects both.
            if (const auto foldersOpt = CreateCaseFolders(root, L"subdirattrs"))
            {
                const auto& folders = *foldersOpt;
                state.Require(SelfTest::EnsureDirectory(folders.left / L"sub"), L"Failed to create sub (left).");
                state.Require(SelfTest::EnsureDirectory(folders.right / L"sub"), L"Failed to create sub (right).");

                const std::filesystem::path leftDir = folders.left / L"sub";
                const DWORD leftAttrs               = ::GetFileAttributesW(leftDir.c_str());
                state.Require(leftAttrs != INVALID_FILE_ATTRIBUTES, L"GetFileAttributesW failed for sub (left).");
                if (leftAttrs != INVALID_FILE_ATTRIBUTES)
                {
                    state.Require(::SetFileAttributesW(leftDir.c_str(), leftAttrs | FILE_ATTRIBUTE_HIDDEN) != 0, L"SetFileAttributesW failed for sub (left).");
                }

                Common::Settings::CompareDirectoriesSettings settings{};
                settings.compareSubdirectoryAttributes = true;

                auto decision = ComputeRootDecision(baseFs, folders, settings, state);
                if (decision)
                {
                    const auto* item = FindItem(*decision, L"sub");
                    state.Require(item != nullptr, L"sub missing from decision.");
                    if (item)
                    {
                        state.Require(item->isDirectory, L"sub expected isDirectory.");
                        state.Require(item->isDifferent, L"sub expected isDifferent with compareSubdirectoryAttributes.");
                        state.Require(item->selectLeft && item->selectRight, L"sub expected select both when attributes differ.");
                        state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::SubdirAttributes),
                                      L"sub expected differenceMask=SubdirAttributes.");
                    }
                }
            }
            else
            {
                state.Require(false, L"Failed to create case folders: subdirattrs.");
            }

            return state.failure.empty();
        });

        SelfTest::RunCase(options,
                          suite,
                          L"missing folder",
                          [&](SelfTest::CaseState& state) noexcept
        {
            // Case: Missing folder is reported without failing the decision.
            if (const auto foldersOpt = CreateCaseFolders(root, L"missing_folder"))
            {
                const auto& folders = *foldersOpt;
                state.Require(SelfTest::EnsureDirectory(folders.left / L"sub"), L"Failed to create sub (left).");
                state.Require(SelfTest::WriteTextFile(folders.left / L"sub" / L"a.txt", "A"), L"Failed to create sub\\a.txt (left).");

                Common::Settings::CompareDirectoriesSettings settings{};
                auto session  = std::make_shared<CompareDirectoriesSession>(baseFs, baseFs, folders.left, folders.right, settings);
                auto decision = session->GetOrComputeDecision(std::filesystem::path(L"sub"));
                state.Require(static_cast<bool>(decision), L"missing folder: decision is null.");
                if (decision)
                {
                    state.Require(SUCCEEDED(decision->hr), L"missing folder: expected decision hr success.");
                    state.Require(! decision->leftFolderMissing, L"missing folder: expected leftFolderMissing=false.");
                    state.Require(decision->rightFolderMissing, L"missing folder: expected rightFolderMissing=true.");
                }
            }
            else
            {
                state.Require(false, L"Failed to create case folders: missing_folder.");
            }

            return state.failure.empty();
        });

        SelfTest::RunCase(options,
                          suite,
                          L"reparse",
                          [&](SelfTest::CaseState& state) noexcept
        {
            // Case: Reparse points are not traversed for subdirectory comparison.
            if (const auto foldersOpt = CreateCaseFolders(root, L"reparse"))
            {
                const auto& folders                = *foldersOpt;
                const std::filesystem::path target = folders.left / L"target";
                state.Require(SelfTest::EnsureDirectory(target), L"Failed to create reparse target (left).");
                state.Require(SelfTest::WriteTextFile(target / L"child.txt", "C"), L"Failed to create target\\child.txt (left).");

                const std::filesystem::path linkPath = folders.left / L"sub";
                const bool linkCreated               = TryCreateDirectorySymlink(linkPath, target);
                if (! linkCreated)
                {
                    const DWORD err = ::GetLastError();
                    if (err == ERROR_PRIVILEGE_NOT_HELD || err == ERROR_ACCESS_DENIED || err == ERROR_INVALID_PARAMETER)
                    {
                        Debug::Warning(L"CompareSelfTest: skipping reparse point test (CreateSymbolicLinkW failed: {0}).", err);
                    }
                    else
                    {
                        state.Require(false, std::format(L"CreateSymbolicLinkW failed unexpectedly: {}.", err));
                    }
                }
                else
                {
                    state.Require(SelfTest::EnsureDirectory(folders.right / L"sub"), L"Failed to create sub directory (right).");

                    Common::Settings::CompareDirectoriesSettings settings{};
                    settings.compareSubdirectories = true;
                    settings.keepIdenticalItems    = true;

                    auto decision = ComputeRootDecision(baseFs, folders, settings, state);
                    if (decision)
                    {
                        const auto* item = FindItem(*decision, L"sub");
                        state.Require(item != nullptr, L"sub missing from decision.");
                        if (item)
                        {
                            state.Require(item->isDirectory, L"sub expected isDirectory.");
                            state.Require(! HasFlag(item->differenceMask, CompareDirectoriesDiffBit::SubdirContent),
                                          L"sub expected SubdirContent not set for reparse points.");
                        }
                    }
                }
            }
            else
            {
                state.Require(false, L"Failed to create case folders: reparse.");
            }

            return state.failure.empty();
        });

        SelfTest::RunCase(options,
                          suite,
                          L"dummy_content",
                          [&](SelfTest::CaseState& state) noexcept
        {
            // Case: Dummy filesystem paths use plugin I/O for content compare (cross-filesystem support).
            if (dummyFs && dummyIo && dummyOps)
            {
                const std::filesystem::path baseRoot  = std::filesystem::path(L"Y:\\") / (L"CompareSelfTest_" + guid) / L"compare";
                const std::filesystem::path leftRoot  = baseRoot / L"left";
                const std::filesystem::path rightRoot = baseRoot / L"right";
                state.Require(EnsureDirectoryExistsFsOps(dummyOps, leftRoot), L"Dummy: failed to create left root.");
                state.Require(EnsureDirectoryExistsFsOps(dummyOps, rightRoot), L"Dummy: failed to create right root.");

                state.Require(WriteFileTextFsIo(dummyIo, leftRoot / L"a.bin", "SAME"), L"Dummy: failed to write a.bin (left).");
                state.Require(WriteFileTextFsIo(dummyIo, rightRoot / L"a.bin", "SAME"), L"Dummy: failed to write a.bin (right).");

                Common::Settings::CompareDirectoriesSettings settings{};
                settings.compareContent     = true;
                settings.keepIdenticalItems = true;

                auto session  = std::make_shared<CompareDirectoriesSession>(dummyFs, dummyFs, leftRoot, rightRoot, settings);
                auto decision = WaitForContentCompare(session, std::filesystem::path{}, L"a.bin", state);
                if (decision)
                {
                    const auto* item = FindItem(*decision, L"a.bin");
                    state.Require(item != nullptr, L"Dummy: a.bin missing from decision.");
                    if (item)
                    {
                        state.Require(! item->isDifferent, L"Dummy: a.bin expected identical after content compare.");
                        state.Require(item->differenceMask == 0u, L"Dummy: a.bin expected differenceMask=0 after content compare.");
                    }
                }
            }
            else
            {
                state.Require(false, L"CompareSelfTest: FileSystemDummy unavailable for cross-filesystem content compare test.");
            }

            return state.failure.empty();
        });

        SelfTest::RunCase(options,
                          suite,
                          L"deep_tree",
                          [&](SelfTest::CaseState& state) noexcept
        {
            // Case: Deep directory trees do not overflow the stack (iterative traversal).
            if (dummyFs && dummyIo && dummyOps)
            {
                const std::filesystem::path baseRoot  = std::filesystem::path(L"Z:\\") / (L"CompareSelfTest_" + guid) / L"deep";
                const std::filesystem::path leftRoot  = baseRoot / L"left";
                const std::filesystem::path rightRoot = baseRoot / L"right";
                state.Require(EnsureDirectoryExistsFsOps(dummyOps, leftRoot), L"Dummy: failed to create deep left root.");
                state.Require(EnsureDirectoryExistsFsOps(dummyOps, rightRoot), L"Dummy: failed to create deep right root.");

                constexpr size_t kDepth = 1024;

                std::filesystem::path leftPath  = leftRoot;
                std::filesystem::path rightPath = rightRoot;
                for (size_t i = 0; i < kDepth; ++i)
                {
                    const std::wstring name = std::format(L"d{:04}", i);
                    leftPath /= name;
                    rightPath /= name;
                    const HRESULT leftHr  = dummyOps->CreateDirectory(leftPath.c_str());
                    const HRESULT rightHr = dummyOps->CreateDirectory(rightPath.c_str());
                    state.Require(SUCCEEDED(leftHr) || leftHr == HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS),
                                  std::format(L"Dummy: failed to create left dir at depth {}.", i));
                    state.Require(SUCCEEDED(rightHr) || rightHr == HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS),
                                  std::format(L"Dummy: failed to create right dir at depth {}.", i));
                }

                state.Require(WriteFileTextFsIo(dummyIo, leftPath / L"leaf.txt", "L"), L"Dummy: failed to create leaf.txt (left).");

                Common::Settings::CompareDirectoriesSettings settings{};
                settings.compareSubdirectories = true;

                auto session = std::make_shared<CompareDirectoriesSession>(dummyFs, dummyFs, leftRoot, rightRoot, settings);
                state.Require(static_cast<bool>(session), L"Dummy: failed to create session (deep_tree).");
                state.Require(StartScanAndWaitForIdle(session, std::chrono::milliseconds{SelfTest::ScaleTimeout(60'000)}),
                              L"Dummy: scan did not become idle within timeout (deep_tree).");
                state.Require(DrainPendingSubdirUpdates(session, 512), L"Dummy: failed to drain pending subtree updates (deep_tree).");

                const auto decision = session->GetOrComputeDecision(std::filesystem::path{});
                if (decision)
                {
                    const auto* item = FindItem(*decision, L"d0000");
                    state.Require(item != nullptr, L"Dummy: d0000 missing from decision.");
                    if (item)
                    {
                        state.Require(item->isDirectory, L"Dummy: d0000 expected isDirectory.");
                        state.Require(item->isDifferent, L"Dummy: d0000 expected isDifferent from deep leaf mismatch.");
                        state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::SubdirContent),
                                      L"Dummy: d0000 expected differenceMask=SubdirContent from deep leaf mismatch.");
                    }
                }
            }
            else
            {
                state.Require(false, L"CompareSelfTest: FileSystemDummy unavailable for deep tree test.");
            }

            return state.failure.empty();
        });

        SelfTest::RunCase(options,
                          suite,
                          L"invalidate",
                          [&](SelfTest::CaseState& state) noexcept
        {
            // Case: Version invalidation mid-scan does not cache stale results.
            if (dummyFs && dummyIo && dummyOps)
            {
                const std::filesystem::path baseRoot  = std::filesystem::path(L"W:\\") / (L"CompareSelfTest_" + guid) / L"invalidate";
                const std::filesystem::path leftRoot  = baseRoot / L"left";
                const std::filesystem::path rightRoot = baseRoot / L"right";
                state.Require(EnsureDirectoryExistsFsOps(dummyOps, leftRoot), L"Dummy: failed to create invalidate left root.");
                state.Require(EnsureDirectoryExistsFsOps(dummyOps, rightRoot), L"Dummy: failed to create invalidate right root.");

                constexpr size_t kDepth         = 256;
                std::filesystem::path leftPath  = leftRoot;
                std::filesystem::path rightPath = rightRoot;
                for (size_t i = 0; i < kDepth; ++i)
                {
                    const std::wstring name = std::format(L"d{}", i);
                    leftPath /= name;
                    rightPath /= name;
                    static_cast<void>(dummyOps->CreateDirectory(leftPath.c_str()));
                    static_cast<void>(dummyOps->CreateDirectory(rightPath.c_str()));
                }
                state.Require(WriteFileTextFsIo(dummyIo, leftPath / L"leaf.txt", "X"), L"Dummy: failed to create invalidate leaf.txt (left).");

                Common::Settings::CompareDirectoriesSettings settings{};
                settings.compareSubdirectories = true;

                auto session                 = std::make_shared<CompareDirectoriesSession>(dummyFs, dummyFs, leftRoot, rightRoot, settings);
                const uint64_t versionBefore = session->GetVersion();

                const auto decisionBefore = session->GetOrComputeDecision(std::filesystem::path{});
                state.Require(static_cast<bool>(decisionBefore), L"Invalidate: initial decision missing.");

                std::mutex progressMutex;
                std::condition_variable progressCv;
                bool scanInProgress = false;

                session->SetScanProgressCallback(
                    [&](const std::filesystem::path&, std::wstring_view, uint64_t scannedFolders, uint64_t, uint32_t activeScans, uint64_t, uint64_t) noexcept
                {
                    if (scannedFolders == 0u || activeScans == 0u)
                    {
                        return;
                    }

                    std::lock_guard lock(progressMutex);
                    scanInProgress = true;
                    progressCv.notify_all();
                });

                session->StartScan();

                {
                    std::unique_lock lock(progressMutex);
                    static_cast<void>(progressCv.wait_for(lock, std::chrono::milliseconds{SelfTest::ScaleTimeout(5'000)}, [&] { return scanInProgress; }));
                }
                session->SetScanProgressCallback({});

                state.Require(scanInProgress, L"Invalidate: scan did not start within timeout.");

                session->Invalidate();
                state.Require(session->GetVersion() == versionBefore + 1u, L"Invalidate: expected version bump.");

                const auto decisionAfter = session->GetOrComputeDecision(std::filesystem::path{});
                state.Require(static_cast<bool>(decisionAfter), L"Invalidate: decision missing after invalidation.");
                if (decisionBefore && decisionAfter)
                {
                    state.Require(decisionAfter != decisionBefore, L"Invalidate: expected a new decision after invalidation (stale result cached).");
                }

                state.Require(StartScanAndWaitForIdle(session, std::chrono::milliseconds{SelfTest::ScaleTimeout(30'000)}),
                              L"Invalidate: scan did not become idle within timeout after invalidation.");

                const CompareDirectoriesPerfStats stats = session->GetPerfStats();
                state.Require(stats.scanActiveScans == 0u, L"Invalidate: expected scanActiveScans == 0 after idle.");
                state.Require(stats.scanQueueSize == 0u, L"Invalidate: expected scanQueueSize == 0 after idle.");
                state.Require(stats.scanScheduledKeys == 0u, L"Invalidate: expected scanScheduledKeys == 0 after idle.");
                state.Require(stats.scanInFlightKeys == 0u, L"Invalidate: expected scanInFlightKeys == 0 after idle.");
            }
            else
            {
                state.Require(false, L"CompareSelfTest: FileSystemDummy unavailable for invalidation test.");
            }

            return state.failure.empty();
        });

        SelfTest::RunCase(options,
                          suite,
                          L"concurrent_get_or_compute_decision",
                          [&](SelfTest::CaseState& state) noexcept
        {
            // Case: Concurrent GetOrComputeDecision and Invalidate does not crash and never returns null.
            if (const auto foldersOpt = CreateCaseFolders(root, L"concurrent_get_or_compute_decision"))
            {
                const auto& folders = *foldersOpt;
                state.Require(SelfTest::WriteTextFile(folders.left / L"a.txt", "A"), L"Failed to create a.txt (left).");
                state.Require(SelfTest::WriteTextFile(folders.right / L"a.txt", "A"), L"Failed to create a.txt (right).");
                state.Require(SelfTest::WriteTextFile(folders.left / L"b.txt", "L"), L"Failed to create b.txt (left).");
                state.Require(SelfTest::WriteTextFile(folders.right / L"b.txt", "R"), L"Failed to create b.txt (right).");

                Common::Settings::CompareDirectoriesSettings settings{};
                const auto session = std::make_shared<CompareDirectoriesSession>(baseFs, baseFs, folders.left, folders.right, settings);

                std::atomic<uint32_t> nullDecisions{0};
                constexpr int kWorkerCount     = 4;
                constexpr int kWorkerIters     = 50;
                constexpr int kInvalidateIters = 10;

                std::vector<std::jthread> workers;
                workers.reserve(kWorkerCount);
                for (int i = 0; i < kWorkerCount; ++i)
                {
                    workers.emplace_back([session, &nullDecisions](std::stop_token) noexcept
                    {
                        for (int j = 0; j < kWorkerIters; ++j)
                        {
                            auto decision = session->GetOrComputeDecision(std::filesystem::path{});
                            if (! decision)
                            {
                                nullDecisions.fetch_add(1u, std::memory_order_relaxed);
                            }
                        }
                    });
                }

                std::jthread invalidator([session](std::stop_token) noexcept
                {
                    for (int j = 0; j < kInvalidateIters; ++j)
                    {
                        session->Invalidate();
                    }
                });

                for (auto& worker : workers)
                {
                    worker.join();
                }
                invalidator.join();

                state.Require(nullDecisions.load(std::memory_order_relaxed) == 0u, L"Concurrent GetOrComputeDecision returned null.");
            }
            else
            {
                state.Require(false, L"Failed to create case folders: concurrent_get_or_compute_decision.");
            }

            return state.failure.empty();
        });

        SelfTest::RunCase(options,
                          suite,
                          L"empty_directories",
                          [&](SelfTest::CaseState& state) noexcept
        {
            // Case: Empty trees produce an empty decision and scan reaches idle.
            if (const auto foldersOpt = CreateCaseFolders(root, L"empty_directories"))
            {
                const auto& folders = *foldersOpt;

                Common::Settings::CompareDirectoriesSettings settings{};
                const auto session = std::make_shared<CompareDirectoriesSession>(baseFs, baseFs, folders.left, folders.right, settings);

                session->StartScan();
                state.Require(StartScanAndWaitForIdle(session, std::chrono::milliseconds{SelfTest::ScaleTimeout(10'000)}),
                              L"Empty directories: scan did not become idle within timeout.");

                const auto decision = session->GetOrComputeDecision(std::filesystem::path{});
                state.Require(static_cast<bool>(decision), L"Empty directories: decision is null.");
                if (decision)
                {
                    state.Require(SUCCEEDED(decision->hr), L"Empty directories: expected decision hr success.");
                    state.Require(decision->items.empty(), L"Empty directories: expected no items.");
                }
            }
            else
            {
                state.Require(false, L"Failed to create case folders: empty_directories.");
            }

            return state.failure.empty();
        });

        SelfTest::RunCase(options,
                          suite,
                          L"ignore",
                          [&](SelfTest::CaseState& state) noexcept
        {
            // Case: Ignore patterns exclude files/directories.
            if (const auto foldersOpt = CreateCaseFolders(root, L"ignore"))
            {
                const auto& folders = *foldersOpt;
                state.Require(SelfTest::WriteTextFile(folders.left / L"ignore.log", "I"), L"Failed to create ignore.log (left).");
                state.Require(SelfTest::WriteTextFile(folders.left / L"keep.txt", "K"), L"Failed to create keep.txt (left).");
                state.Require(SelfTest::EnsureDirectory(folders.left / L"ignore_dir"), L"Failed to create ignore_dir (left).");
                state.Require(SelfTest::EnsureDirectory(folders.left / L"keep_dir"), L"Failed to create keep_dir (left).");

                Common::Settings::CompareDirectoriesSettings settings{};
                settings.ignoreFiles               = true;
                settings.ignoreFilesPatterns       = L"*.log";
                settings.ignoreDirectories         = true;
                settings.ignoreDirectoriesPatterns = L"ignore*";

                auto decision = ComputeRootDecision(baseFs, folders, settings, state);
                if (decision)
                {
                    state.Require(FindItem(*decision, L"keep.txt") != nullptr, L"keep.txt expected in decision.");
                    state.Require(FindItem(*decision, L"ignore.log") == nullptr, L"ignore.log expected to be ignored.");
                    state.Require(FindItem(*decision, L"keep_dir") != nullptr, L"keep_dir expected in decision.");
                    state.Require(FindItem(*decision, L"ignore_dir") == nullptr, L"ignore_dir expected to be ignored.");
                }
            }
            else
            {
                state.Require(false, L"Failed to create case folders: ignore.");
            }

            return state.failure.empty();
        });

        SelfTest::RunCase(options,
                          suite,
                          L"ignore_multiple_patterns",
                          [&](SelfTest::CaseState& state) noexcept
        {
            // Case: Multiple ignore patterns exclude all matching files.
            if (const auto foldersOpt = CreateCaseFolders(root, L"ignore_multiple_patterns"))
            {
                const auto& folders = *foldersOpt;
                state.Require(SelfTest::WriteTextFile(folders.left / L"ignore.log", "I"), L"Failed to create ignore.log (left).");
                state.Require(SelfTest::WriteTextFile(folders.left / L"foo.tmp", "T"), L"Failed to create foo.tmp (left).");
                state.Require(SelfTest::WriteTextFile(folders.left / L"keep.txt", "K"), L"Failed to create keep.txt (left).");

                Common::Settings::CompareDirectoriesSettings settings{};
                settings.ignoreFiles         = true;
                settings.ignoreFilesPatterns = L"*.log;*.tmp";

                auto decision = ComputeRootDecision(baseFs, folders, settings, state);
                if (decision)
                {
                    state.Require(FindItem(*decision, L"keep.txt") != nullptr, L"keep.txt expected in decision.");
                    state.Require(FindItem(*decision, L"ignore.log") == nullptr, L"ignore.log expected to be ignored.");
                    state.Require(FindItem(*decision, L"foo.tmp") == nullptr, L"foo.tmp expected to be ignored.");
                }
            }
            else
            {
                state.Require(false, L"Failed to create case folders: ignore_multiple_patterns.");
            }

            return state.failure.empty();
        });

        SelfTest::RunCase(options,
                          suite,
                          L"ignore_pattern_length_cap",
                          [&](SelfTest::CaseState& state) noexcept
        {
            // Case: Overly long ignore patterns are dropped (harden against pathological inputs).
            if (const auto foldersOpt = CreateCaseFolders(root, L"ignore_pattern_length_cap"))
            {
                const auto& folders = *foldersOpt;
                state.Require(SelfTest::WriteTextFile(folders.left / L"cap.txt", "C"), L"Failed to create cap.txt (left).");

                Common::Settings::CompareDirectoriesSettings settings{};
                settings.ignoreFiles         = true;
                settings.ignoreFilesPatterns = std::wstring(129, L'*'); // exceeds cap (128) => dropped

                auto decision = ComputeRootDecision(baseFs, folders, settings, state);
                if (decision)
                {
                    state.Require(FindItem(*decision, L"cap.txt") != nullptr, L"cap.txt expected in decision (pattern too long must not ignore).");
                }
            }
            else
            {
                state.Require(false, L"Failed to create case folders: ignore_pattern_length_cap.");
            }

            return state.failure.empty();
        });

        SelfTest::RunCase(options,
                          suite,
                          L"ignore_pattern_count_cap",
                          [&](SelfTest::CaseState& state) noexcept
        {
            // Case: Ignore patterns are capped to a bounded count to keep matching work bounded.
            if (const auto foldersOpt = CreateCaseFolders(root, L"ignore_pattern_count_cap"))
            {
                const auto& folders = *foldersOpt;
                state.Require(SelfTest::WriteTextFile(folders.left / L"cap.log", "C"), L"Failed to create cap.log (left).");

                std::wstring patterns;
                for (int i = 0; i < 32; ++i)
                {
                    patterns.append(std::format(L"p{:02};", i)); // non-matching
                }
                patterns.append(L"*.log"); // would match, but is beyond the cap

                Common::Settings::CompareDirectoriesSettings settings{};
                settings.ignoreFiles         = true;
                settings.ignoreFilesPatterns = patterns;

                auto decision = ComputeRootDecision(baseFs, folders, settings, state);
                if (decision)
                {
                    state.Require(FindItem(*decision, L"cap.log") != nullptr, L"cap.log expected in decision (pattern beyond cap must not ignore).");
                }
            }
            else
            {
                state.Require(false, L"Failed to create case folders: ignore_pattern_count_cap.");
            }

            return state.failure.empty();
        });

        SelfTest::RunCase(options,
                          suite,
                          L"ignore_wildcard_pathology_runtime_bound",
                          [&](SelfTest::CaseState& state) noexcept
        {
            // Case: A pathological wildcard pattern list must not hang the scan (work is bounded by pattern caps + matcher budgets).
            if (const auto foldersOpt = CreateCaseFolders(root, L"ignore_wildcard_pathology_runtime_bound"))
            {
                const auto& folders = *foldersOpt;

                constexpr int kFileCount = 256;
                for (int i = 0; i < kFileCount; ++i)
                {
                    const std::wstring name = std::format(L"file{:04}.txt", i);
                    state.Require(SelfTest::WriteTextFile(folders.left / name, "X"), L"Failed to create file (left).");
                    if (i != 0)
                    {
                        state.Require(SelfTest::WriteTextFile(folders.right / name, "X"), L"Failed to create file (right).");
                    }
                }

                std::wstring longWildcard;
                longWildcard.reserve(128);
                for (int i = 0; i < 64; ++i)
                {
                    longWildcard.append(L"*a"); // length 128; requires many 'a' (won't match our file names)
                }

                std::wstring patterns;
                patterns.reserve((longWildcard.size() + 1) * 32);
                for (int i = 0; i < 32; ++i)
                {
                    patterns.append(longWildcard);
                    patterns.push_back(L';');
                }

                Common::Settings::CompareDirectoriesSettings settings{};
                settings.ignoreFiles         = true;
                settings.ignoreFilesPatterns = patterns;

                const auto session = std::make_shared<CompareDirectoriesSession>(baseFs, baseFs, folders.left, folders.right, settings);
                state.Require(StartScanAndWaitForIdle(session, std::chrono::milliseconds{SelfTest::ScaleTimeout(10'000)}),
                              L"Wildcard patterns: scan did not become idle within timeout.");

                const auto decision = session->GetOrComputeDecision(std::filesystem::path{});
                state.Require(static_cast<bool>(decision), L"Wildcard patterns: decision is null.");
                if (decision)
                {
                    state.Require(SUCCEEDED(decision->hr), L"Wildcard patterns: expected decision hr success.");
                    const auto* item = FindItem(*decision, L"file0000.txt");
                    state.Require(item != nullptr, L"file0000.txt expected in decision.");
                    if (item)
                    {
                        state.Require(item->isDifferent, L"file0000.txt should remain surfaced as a differing item.");
                        state.Require(item->existsLeft && ! item->existsRight,
                                      L"file0000.txt should be present only on the left side.");
                    }
                }
            }
            else
            {
                state.Require(false, L"Failed to create case folders: ignore_wildcard_pathology_runtime_bound.");
            }

            return state.failure.empty();
        });

        SelfTest::RunCase(options,
                          suite,
                          L"crash_quarantine_synthetic_marker",
                          [&](SelfTest::CaseState& state) noexcept
        {
            // Case: A synthetic crash marker should trigger crash-quarantine to disable the last-active filesystem plugin.
            constexpr std::wstring_view kPluginId = L"builtin/file-system-s3";

            const std::filesystem::path sessionPath = SessionState::GetSessionStatePath();
            state.Require(! sessionPath.empty(), L"SessionState::GetSessionStatePath returned empty path.");
            const std::filesystem::path baseDir = sessionPath.parent_path();
            state.Require(! baseDir.empty(), L"SessionState base directory is empty.");

            const std::filesystem::path crashDir   = baseDir / L"Crashes";
            const std::filesystem::path markerPath = crashDir / L"last_crash.txt";

            struct Backup
            {
                bool hadFile = false;
                std::vector<std::byte> bytes;
            };

            const auto readAllBytes = [&](const std::filesystem::path& path, Backup& backup) noexcept -> bool
            {
                backup.hadFile = false;
                backup.bytes.clear();

                std::error_code ec;
                if (! std::filesystem::exists(path, ec))
                {
                    return true;
                }

                backup.hadFile = true;

                wil::unique_handle file(CreateFileW(path.c_str(),
                                                    GENERIC_READ,
                                                    FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                                    nullptr,
                                                    OPEN_EXISTING,
                                                    FILE_ATTRIBUTE_NORMAL,
                                                    nullptr));
                if (! file)
                {
                    state.Require(false, L"Failed to open backup file for read.");
                    return false;
                }

                LARGE_INTEGER size{};
                if (GetFileSizeEx(file.get(), &size) == 0)
                {
                    state.Require(false, L"Failed to get backup file size.");
                    return false;
                }

                if (size.QuadPart < 0 || size.QuadPart > static_cast<LONGLONG>((std::numeric_limits<DWORD>::max)()))
                {
                    state.Require(false, L"Backup file too large.");
                    return false;
                }

                const DWORD bytesToRead = static_cast<DWORD>(size.QuadPart);
                backup.bytes.resize(bytesToRead);

                DWORD readBytes = 0;
                if (bytesToRead > 0)
                {
                    if (ReadFile(file.get(), backup.bytes.data(), bytesToRead, &readBytes, nullptr) == 0 || readBytes != bytesToRead)
                    {
                        state.Require(false, L"Failed to read backup file bytes.");
                        return false;
                    }
                }

                return true;
            };

            Backup sessionBackup{};
            Backup markerBackup{};
            if (! readAllBytes(sessionPath, sessionBackup) || ! readAllBytes(markerPath, markerBackup))
            {
                return state.failure.empty();
            }

            const auto restore = wil::scope_exit([&] noexcept
            {
                const auto restoreFile = [&](const std::filesystem::path& path, const Backup& backup) noexcept
                {
                    std::error_code ec;
                    if (! backup.hadFile)
                    {
                        static_cast<void>(std::filesystem::remove(path, ec));
                        return;
                    }

                    static_cast<void>(std::filesystem::create_directories(path.parent_path(), ec));
                    wil::unique_handle file(CreateFileW(path.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
                    if (! file)
                    {
                        return;
                    }

                    if (! backup.bytes.empty())
                    {
                        DWORD written = 0;
                        static_cast<void>(WriteFile(file.get(),
                                                    backup.bytes.data(),
                                                    static_cast<DWORD>(backup.bytes.size()),
                                                    &written,
                                                    nullptr));
                    }

                    static_cast<void>(FlushFileBuffers(file.get()));
                };

                restoreFile(sessionPath, sessionBackup);
                restoreFile(markerPath, markerBackup);
            });

            const auto writeUtf16File = [&](const std::filesystem::path& path, std::wstring_view content) noexcept -> bool
            {
                std::error_code ec;
                std::filesystem::create_directories(path.parent_path(), ec);
                if (ec)
                {
                    state.Require(false, L"Failed to create directory for marker file.");
                    return false;
                }

                wil::unique_handle file(CreateFileW(path.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
                if (! file)
                {
                    state.Require(false, L"Failed to create marker file.");
                    return false;
                }

                const wchar_t bom = 0xFEFF;
                DWORD written     = 0;
                if (! WriteFile(file.get(), &bom, sizeof(bom), &written, nullptr))
                {
                    state.Require(false, L"Failed to write marker BOM.");
                    return false;
                }

                const DWORD bytes = static_cast<DWORD>(content.size() * sizeof(wchar_t));
                if (bytes > 0)
                {
                    if (! WriteFile(file.get(), content.data(), bytes, &written, nullptr))
                    {
                        state.Require(false, L"Failed to write marker content.");
                        return false;
                    }
                }

                static_cast<void>(FlushFileBuffers(file.get()));
                return true;
            };

            std::wstring sessionText;
            sessionText.reserve(32 + kPluginId.size());
            sessionText.append(L"fsPlugin=");
            sessionText.append(kPluginId);
            sessionText.append(L"\r\nop=browse\r\n");

            if (! writeUtf16File(sessionPath, sessionText) || ! writeUtf16File(markerPath, L"synthetic"))
            {
                return state.failure.empty();
            }

            Common::Settings::Settings settings{};
            settings.plugins.currentFileSystemPluginId = std::wstring(kPluginId);

            const bool oldAutoAccept = HostGetAutoAcceptPrompts();
            HostSetAutoAcceptPrompts(true);
            const auto restoreAutoAccept = wil::scope_exit([&] noexcept { HostSetAutoAcceptPrompts(oldAutoAccept); });

            CrashQuarantine::OfferPluginDisableIfPreviousCrashDetected(settings);

            const auto isDisabled = [&](std::wstring_view id) noexcept -> bool
            {
                for (const std::wstring& disabled : settings.plugins.disabledPluginIds)
                {
                    if (OrdinalString::EqualsNoCase(disabled, id))
                    {
                        return true;
                    }
                }
                return false;
            };

            state.Require(isDisabled(kPluginId), L"Crash quarantine did not disable the expected plugin id.");
            state.Require(settings.plugins.currentFileSystemPluginId.empty(), L"Crash quarantine did not clear currentFileSystemPluginId.");

            return state.failure.empty();
        });

        SelfTest::RunCase(options,
                          suite,
                          L"showIdentical",
                          [&](SelfTest::CaseState& state) noexcept
        {
            // Case: keepIdenticalItems retains identical files; showIdenticalItems toggles view without invalidating decisions.
            if (const auto foldersOpt = CreateCaseFolders(root, L"identical"))
            {
                const auto& folders = *foldersOpt;
                state.Require(SelfTest::WriteTextFile(folders.left / L"same.txt", "SAME"), L"Failed to create same.txt (left).");
                state.Require(SelfTest::WriteTextFile(folders.right / L"same.txt", "SAME"), L"Failed to create same.txt (right).");

                Common::Settings::CompareDirectoriesSettings settings{};
                settings.keepIdenticalItems = true;
                auto session                = std::make_shared<CompareDirectoriesSession>(baseFs, baseFs, folders.left, folders.right, settings);
                const auto fsLeft           = CreateCompareDirectoriesFileSystem(ComparePane::Left, session);
                const auto fsRight          = CreateCompareDirectoriesFileSystem(ComparePane::Right, session);

                const uint64_t versionBefore = session->GetVersion();
                const auto decisionBefore    = session->GetOrComputeDecision(std::filesystem::path{});
                state.Require(static_cast<bool>(decisionBefore), L"Decision missing (before showIdentical).");
                if (decisionBefore)
                {
                    const auto* item = FindItem(*decisionBefore, L"same.txt");
                    state.Require(item != nullptr, L"same.txt missing from cached decision when keepIdenticalItems is enabled.");
                    if (item)
                    {
                        state.Require(! item->isDifferent, L"same.txt expected identical (before showIdentical).");
                        state.Require(item->differenceMask == 0u, L"same.txt expected differenceMask=0 (before showIdentical).");
                    }
                }

                state.Require(! ContainsName(EnumerateDirectoryNames(fsLeft, folders.left, state), L"same.txt"),
                              L"same.txt expected excluded from left enumeration (before showIdentical).");
                state.Require(! ContainsName(EnumerateDirectoryNames(fsRight, folders.right, state), L"same.txt"),
                              L"same.txt expected excluded from right enumeration (before showIdentical).");

                settings.showIdenticalItems = true;
                session->SetSettings(settings);

                const uint64_t versionAfter = session->GetVersion();
                state.Require(versionAfter == versionBefore, L"SetSettings(showIdenticalItems) must not invalidate decisions.");

                const auto decisionAfter = session->GetOrComputeDecision(std::filesystem::path{});
                state.Require(decisionAfter == decisionBefore, L"Decision should remain cached across showIdenticalItems toggle.");
                state.Require(static_cast<bool>(decisionAfter), L"Decision missing (after showIdentical).");
                if (decisionAfter)
                {
                    const auto* item = FindItem(*decisionAfter, L"same.txt");
                    state.Require(item != nullptr, L"same.txt missing from decision (after showIdentical).");
                    if (item)
                    {
                        state.Require(! item->isDifferent, L"same.txt expected identical (after showIdentical).");
                        state.Require(item->differenceMask == 0u, L"same.txt expected differenceMask=0 (after showIdentical).");
                    }
                }

                state.Require(ContainsName(EnumerateDirectoryNames(fsLeft, folders.left, state), L"same.txt"),
                              L"same.txt expected included in left enumeration (after showIdentical).");
                state.Require(ContainsName(EnumerateDirectoryNames(fsRight, folders.right, state), L"same.txt"),
                              L"same.txt expected included in right enumeration (after showIdentical).");
            }
            else
            {
                state.Require(false, L"Failed to create case folders: identical.");
            }

            return state.failure.empty();
        });

        SelfTest::RunCase(options,
                          suite,
                          L"content_pending_elided",
                          [&](SelfTest::CaseState& state) noexcept
        {
            // Case: When keepIdenticalItems is off, file-level ContentPending placeholders are elided (tracked per-folder)
            // so content compare does not explode memory on very large folders.
            if (const auto foldersOpt = CreateCaseFolders(root, L"content_pending_elided"))
            {
                const auto& folders = *foldersOpt;

                constexpr int kFileCount = 200;

                std::vector<std::byte> payload(16 * 1024);
                for (size_t i = 0; i < payload.size(); ++i)
                {
                    payload[i] = static_cast<std::byte>(i & 0xFF);
                }

                const std::span<const std::byte> payloadSpan(payload.data(), payload.size());
                for (int i = 0; i < kFileCount; ++i)
                {
                    const std::wstring name = std::format(L"f_{:04}.bin", i);
                    state.Require(SelfTest::WriteBinaryFile(folders.left / name, payloadSpan), L"Failed to write test file (left).");
                    state.Require(SelfTest::WriteBinaryFile(folders.right / name, payloadSpan), L"Failed to write test file (right).");
                }

                Common::Settings::CompareDirectoriesSettings settings{};
                settings.compareSize        = false;
                settings.compareDateTime    = false;
                settings.compareAttributes  = false;
                settings.compareContent     = true;
                settings.keepIdenticalItems = false;

                auto session = std::make_shared<CompareDirectoriesSession>(baseFs, baseFs, folders.left, folders.right, settings);

                std::mutex mutex;
                std::condition_variable cv;
                bool contentDone      = false;
                uint64_t lastPending  = 0;
                uint64_t lastTotal    = 0;
                uint64_t lastComplete = 0;

                session->SetContentProgressCallback([&](uint32_t /*workerIndex*/,
                                                        const std::filesystem::path& /*relativeFolder*/,
                                                        std::wstring_view /*entryName*/,
                                                        uint64_t /*fileTotalBytes*/,
                                                        uint64_t /*fileCompletedBytes*/,
                                                        uint64_t /*overallTotalBytes*/,
                                                        uint64_t /*overallCompletedBytes*/,
                                                        uint64_t pendingContentCompares,
                                                        uint64_t totalContentCompares,
                                                        uint64_t completedContentCompares) noexcept
                {
                    std::lock_guard guard(mutex);
                    lastPending  = pendingContentCompares;
                    lastTotal    = totalContentCompares;
                    lastComplete = completedContentCompares;
                    if (pendingContentCompares == 0u && totalContentCompares != 0u && completedContentCompares == totalContentCompares)
                    {
                        contentDone = true;
                    }
                    cv.notify_all();
                });

                const auto decisionInitial = session->GetOrComputeDecision(std::filesystem::path{});
                state.Require(static_cast<bool>(decisionInitial), L"content_pending_elided: decision missing.");
                if (decisionInitial)
                {
                    state.Require(decisionInitial->items.empty(),
                                  L"content_pending_elided: expected no per-file ContentPending items in differences-only mode.");
                    state.Require(decisionInitial->pendingContentCompareCount == static_cast<uint32_t>(kFileCount),
                                  L"content_pending_elided: expected pendingContentCompareCount to match the file count.");
                    state.Require(decisionInitial->anyPending, L"content_pending_elided: expected anyPending=true while compares are queued.");
                }

                const auto deadline =
                    std::chrono::steady_clock::now() + std::chrono::milliseconds{static_cast<std::chrono::milliseconds::rep>(SelfTest::ScaleTimeout(20'000))};
                {
                    std::unique_lock lock(mutex);
                    cv.wait_until(lock, deadline, [&] { return contentDone; });
                }

                state.Require(
                    contentDone,
                    std::format(
                        L"content_pending_elided: content compares did not complete. pending={} total={} completed={}", lastPending, lastTotal, lastComplete));
                state.Require(lastTotal == static_cast<uint64_t>(kFileCount), L"content_pending_elided: unexpected totalContentCompares.");
                state.Require(lastComplete == static_cast<uint64_t>(kFileCount), L"content_pending_elided: unexpected completedContentCompares.");

                const auto decisionFinal = session->GetOrComputeDecision(std::filesystem::path{});
                state.Require(static_cast<bool>(decisionFinal), L"content_pending_elided: final decision missing.");
                if (decisionFinal)
                {
                    state.Require(decisionFinal->items.empty(), L"content_pending_elided: expected no surfaced items after equal content compares.");
                    state.Require(decisionFinal->pendingContentCompareCount == 0u,
                                  L"content_pending_elided: expected pendingContentCompareCount=0 after completion.");
                    state.Require(! decisionFinal->anyPending, L"content_pending_elided: expected anyPending=false after completion.");
                }
            }
            else
            {
                state.Require(false, L"Failed to create case folders: content_pending_elided.");
            }

            return state.failure.empty();
        });

        SelfTest::RunCase(options,
                          suite,
                          L"setCompareEnabled",
                          [&](SelfTest::CaseState& state) noexcept
        {
            // Case: SetCompareEnabled(false) stops producing decisions; re-enabling resumes.
            if (const auto foldersOpt = CreateCaseFolders(root, L"setCompareEnabled"))
            {
                const auto& folders = *foldersOpt;
                state.Require(SelfTest::WriteTextFile(folders.left / L"a.txt", "A"), L"Failed to create a.txt (left).");
                state.Require(SelfTest::WriteTextFile(folders.right / L"b.txt", "B"), L"Failed to create b.txt (right).");

                auto session =
                    std::make_shared<CompareDirectoriesSession>(baseFs, baseFs, folders.left, folders.right, Common::Settings::CompareDirectoriesSettings{});

                state.Require(session->IsCompareEnabled(), L"IsCompareEnabled should be true by default.");

                // When compare is disabled, ReadDirectoryInfo falls back to the base filesystem
                // and shows all files (no comparison filtering applied).
                session->SetCompareEnabled(false);
                state.Require(! session->IsCompareEnabled(), L"IsCompareEnabled should be false after SetCompareEnabled(false).");

                {
                    const auto fsLeft  = CreateCompareDirectoriesFileSystem(ComparePane::Left, session);
                    const auto fsRight = CreateCompareDirectoriesFileSystem(ComparePane::Right, session);

                    const auto leftNames  = EnumerateDirectoryNames(fsLeft, folders.left, state);
                    const auto rightNames = EnumerateDirectoryNames(fsRight, folders.right, state);

                    // Disabled compare: both sides should see their own files unfiltered.
                    state.Require(ContainsName(leftNames, L"a.txt"), L"setCompareEnabled: a.txt should be visible in left when compare is disabled.");
                    state.Require(ContainsName(rightNames, L"b.txt"), L"setCompareEnabled: b.txt should be visible in right when compare is disabled.");
                    // a.txt only exists on the left, b.txt only exists on the right — in enabled mode
                    // they would be filtered to their own pane; disabled should expose them as-is.
                    state.Require(! ContainsName(leftNames, L"b.txt"), L"setCompareEnabled: b.txt should not appear in the left pane.");
                    state.Require(! ContainsName(rightNames, L"a.txt"), L"setCompareEnabled: a.txt should not appear in the right pane.");
                }

                session->SetCompareEnabled(true);
                state.Require(session->IsCompareEnabled(), L"IsCompareEnabled should be true after re-enabling.");

                // After re-enabling, decisions should be obtainable and filtering should be back.
                auto decision = session->GetOrComputeDecision(std::filesystem::path{});
                state.Require(static_cast<bool>(decision), L"GetOrComputeDecision should succeed after re-enabling compare.");

                {
                    const auto fsLeft  = CreateCompareDirectoriesFileSystem(ComparePane::Left, session);
                    const auto fsRight = CreateCompareDirectoriesFileSystem(ComparePane::Right, session);

                    const auto leftNames  = EnumerateDirectoryNames(fsLeft, folders.left, state);
                    const auto rightNames = EnumerateDirectoryNames(fsRight, folders.right, state);

                    // Re-enabled compare: only pane-relevant different items are shown.
                    state.Require(ContainsName(leftNames, L"a.txt"), L"setCompareEnabled: a.txt should be shown in left pane after re-enable (only in left).");
                    state.Require(! ContainsName(leftNames, L"b.txt"), L"setCompareEnabled: b.txt should not appear in left pane after re-enable.");
                    state.Require(ContainsName(rightNames, L"b.txt"),
                                  L"setCompareEnabled: b.txt should be shown in right pane after re-enable (only in right).");
                    state.Require(! ContainsName(rightNames, L"a.txt"), L"setCompareEnabled: a.txt should not appear in right pane after re-enable.");
                }
            }
            else
            {
                state.Require(false, L"Failed to create case folders: setCompareEnabled.");
            }

            return state.failure.empty();
        });

        SelfTest::RunCase(options,
                          suite,
                          L"invalidateForPath",
                          [&](SelfTest::CaseState& state) noexcept
        {
            // Case: InvalidateForAbsolutePath invalidates only the targeted subtree.
            if (const auto foldersOpt = CreateCaseFolders(root, L"invalidateForPath"))
            {
                const auto& folders = *foldersOpt;
                state.Require(SelfTest::EnsureDirectory(folders.left / L"sub1"), L"Failed to create sub1 (left).");
                state.Require(SelfTest::EnsureDirectory(folders.right / L"sub1"), L"Failed to create sub1 (right).");
                state.Require(SelfTest::WriteTextFile(folders.left / L"sub1" / L"f.txt", "X"), L"Failed to create sub1/f.txt (left).");
                state.Require(SelfTest::WriteTextFile(folders.right / L"sub1" / L"f.txt", "X"), L"Failed to create sub1/f.txt (right).");
                state.Require(SelfTest::EnsureDirectory(folders.left / L"sub2"), L"Failed to create sub2 (left).");
                state.Require(SelfTest::EnsureDirectory(folders.right / L"sub2"), L"Failed to create sub2 (right).");
                state.Require(SelfTest::WriteTextFile(folders.left / L"sub2" / L"g.txt", "Y"), L"Failed to create sub2/g.txt (left).");
                state.Require(SelfTest::WriteTextFile(folders.right / L"sub2" / L"g.txt", "Y"), L"Failed to create sub2/g.txt (right).");

                auto session =
                    std::make_shared<CompareDirectoriesSession>(baseFs, baseFs, folders.left, folders.right, Common::Settings::CompareDirectoriesSettings{});

                // Warm up both subtrees.
                const auto decisionSub1Before = session->GetOrComputeDecision(std::filesystem::path(L"sub1"));
                const auto decisionSub2Before = session->GetOrComputeDecision(std::filesystem::path(L"sub2"));
                state.Require(static_cast<bool>(decisionSub1Before), L"sub1 decision missing before invalidate.");
                state.Require(static_cast<bool>(decisionSub2Before), L"sub2 decision missing before invalidate.");

                // Invalidate only sub1's absolute path.
                session->InvalidateForAbsolutePath(folders.left / L"sub1", /*includeSubtree=*/true);

                const auto decisionSub1After = session->GetOrComputeDecision(std::filesystem::path(L"sub1"));
                const auto decisionSub2After = session->GetOrComputeDecision(std::filesystem::path(L"sub2"));

                state.Require(static_cast<bool>(decisionSub1After), L"sub1 decision missing after invalidate.");
                state.Require(static_cast<bool>(decisionSub2After), L"sub2 decision missing after invalidate.");

                // Sub1 must be a different (newly computed) decision object.
                state.Require(decisionSub1After != decisionSub1Before, L"sub1 decision should be new after InvalidateForAbsolutePath.");
                // Sub2 must be the same cached object — it was not invalidated.
                state.Require(decisionSub2After == decisionSub2Before, L"sub2 decision should remain cached (not invalidated).");
            }
            else
            {
                state.Require(false, L"Failed to create case folders: invalidateForPath.");
            }

            return state.failure.empty();
        });

        SelfTest::RunCase(options,
                          suite,
                          L"decisionUpdatedCallback",
                          [&](SelfTest::CaseState& state) noexcept
        {
            // Case: SetDecisionUpdatedCallback fires after Invalidate().
            if (const auto foldersOpt = CreateCaseFolders(root, L"decisionUpdatedCallback"))
            {
                const auto& folders = *foldersOpt;
                // Use compareContent=true with same-size but byte-different files so a content-compare
                // job is enqueued and dispatched to a worker thread.  The callback fires on that worker
                // thread when the compare job completes (size-different files are short-circuited without
                // an async job and would never fire the callback).
                state.Require(SelfTest::WriteTextFile(folders.left / L"a.txt", "AAAA"), L"Failed to create a.txt (left).");
                state.Require(SelfTest::WriteTextFile(folders.right / L"a.txt", "BBBB"), L"Failed to create a.txt (right).");

                Common::Settings::CompareDirectoriesSettings settings{};
                settings.compareContent = true;

                auto session = std::make_shared<CompareDirectoriesSession>(baseFs, baseFs, folders.left, folders.right, settings);

                std::atomic<int> callbackCount{0};
                session->SetDecisionUpdatedCallback([&]() noexcept { callbackCount.fetch_add(1, std::memory_order_relaxed); });

                // Trigger a scan so content-compare workers are started.
                static_cast<void>(session->GetOrComputeDecision(std::filesystem::path{}));

                // Wait up to 10 s for the callback to fire at least once, polling GetOrComputeDecision
                // to keep the scan driving (consistent with the WaitForContentCompare pattern).
                const auto deadline = std::chrono::steady_clock::now() + std::chrono::milliseconds(SelfTest::ScaleTimeout(10'000));
                while (callbackCount.load(std::memory_order_relaxed) == 0 && std::chrono::steady_clock::now() < deadline)
                {
                    static_cast<void>(session->GetOrComputeDecision(std::filesystem::path{}));
                    std::this_thread::sleep_for(std::chrono::milliseconds(10));
                }

                state.Require(callbackCount.load(std::memory_order_relaxed) > 0,
                              L"DecisionUpdatedCallback must fire at least once after content compare completes.");

                // Unregister before session is destroyed to avoid dangling reference.
                session->SetDecisionUpdatedCallback(nullptr);
            }
            else
            {
                state.Require(false, L"Failed to create case folders: decisionUpdatedCallback.");
            }

            return state.failure.empty();
        });

        SelfTest::RunCase(options,
                          suite,
                          L"uiVersion",
                          [&](SelfTest::CaseState& state) noexcept
        {
            // Case: GetUiVersion increments on Invalidate() and after FlushPendingContentCompareUpdates().
            if (const auto foldersOpt = CreateCaseFolders(root, L"uiVersion"))
            {
                const auto& folders = *foldersOpt;
                state.Require(SelfTest::WriteTextFile(folders.left / L"a.txt", "A"), L"Failed to create a.txt (left).");
                state.Require(SelfTest::WriteTextFile(folders.right / L"a.txt", "A"), L"Failed to create a.txt (right).");

                auto session =
                    std::make_shared<CompareDirectoriesSession>(baseFs, baseFs, folders.left, folders.right, Common::Settings::CompareDirectoriesSettings{});

                const uint64_t uiV0 = session->GetUiVersion();
                const uint64_t ver0 = session->GetVersion();

                session->Invalidate();
                const uint64_t uiV1 = session->GetUiVersion();
                const uint64_t ver1 = session->GetVersion();

                state.Require(uiV1 != uiV0, L"GetUiVersion should change after Invalidate().");
                state.Require(ver1 != ver0, L"GetVersion should change after Invalidate().");
            }
            else
            {
                state.Require(false, L"Failed to create case folders: uiVersion.");
            }

            return state.failure.empty();
        });

        SelfTest::RunCase(options,
                          suite,
                          L"accessors",
                          [&](SelfTest::CaseState& state) noexcept
        {
            // Case: Accessor getters return correct values after construction.
            if (const auto foldersOpt = CreateCaseFolders(root, L"accessors"))
            {
                const auto& folders = *foldersOpt;
                Common::Settings::CompareDirectoriesSettings settings{};
                settings.compareSize = true;

                auto session = std::make_shared<CompareDirectoriesSession>(baseFs, baseFs, folders.left, folders.right, settings);

                state.Require(session->GetRoot(ComparePane::Left) == folders.left, L"GetRoot(Left) should match the left root passed to constructor.");
                state.Require(session->GetRoot(ComparePane::Right) == folders.right, L"GetRoot(Right) should match the right root passed to constructor.");
                state.Require(session->GetSettings().compareSize == settings.compareSize,
                              L"GetSettings().compareSize should match the value passed to constructor.");

                // TryMakeRelative / ResolveAbsolute round-trip.
                const std::filesystem::path sub(L"subdir");
                const std::filesystem::path absLeft = folders.left / sub;
                const auto relOpt                   = session->TryMakeRelative(ComparePane::Left, absLeft);
                state.Require(relOpt.has_value(), L"TryMakeRelative should succeed for a path under the left root.");
                if (relOpt.has_value())
                {
                    state.Require(relOpt.value() == sub, L"TryMakeRelative should return the expected relative path.");
                    const std::filesystem::path resolved = session->ResolveAbsolute(ComparePane::Left, relOpt.value());
                    state.Require(resolved == absLeft, L"ResolveAbsolute round-trip should match the original absolute path.");
                }
            }
            else
            {
                state.Require(false, L"Failed to create case folders: accessors.");
            }

            return state.failure.empty();
        });

        SelfTest::RunCase(options,
                          suite,
                          L"plugin_path_math",
                          [&](SelfTest::CaseState& state) noexcept
        {
            // Case: TryMakeRelative / ResolveAbsolute behave correctly for plugin-style (forward-slash) paths.
            const std::filesystem::path pluginRoot(L"/a");
            auto session = std::make_shared<CompareDirectoriesSession>(baseFs, baseFs, pluginRoot, pluginRoot, Common::Settings::CompareDirectoriesSettings{});

            const auto relRoot = session->TryMakeRelative(ComparePane::Left, std::filesystem::path(L"/a"));
            state.Require(relRoot.has_value(), L"plugin_path_math: TryMakeRelative should succeed for the root itself.");
            if (relRoot.has_value())
            {
                state.Require(relRoot.value().empty(), L"plugin_path_math: TryMakeRelative(root, root) should return empty relative path.");
            }

            const auto relOpt = session->TryMakeRelative(ComparePane::Left, std::filesystem::path(L"/a/b"));
            state.Require(relOpt.has_value(), L"plugin_path_math: TryMakeRelative should succeed for /a/b.");
            if (relOpt.has_value())
            {
                state.Require(relOpt.value().generic_wstring() == L"b", L"plugin_path_math: TryMakeRelative(/a, /a/b) should return 'b'.");

                const std::filesystem::path resolved = session->ResolveAbsolute(ComparePane::Left, relOpt.value());
                state.Require(resolved.generic_wstring() == L"/a/b", L"plugin_path_math: ResolveAbsolute(/a, b) should return /a/b.");
            }

            const std::filesystem::path resolvedRoot = session->ResolveAbsolute(ComparePane::Left, std::filesystem::path{});
            state.Require(resolvedRoot.generic_wstring() == L"/a", L"plugin_path_math: ResolveAbsolute(/a, '') should return /a.");

            const auto outside = session->TryMakeRelative(ComparePane::Left, std::filesystem::path(L"/x"));
            state.Require(! outside.has_value(), L"plugin_path_math: TryMakeRelative should return nullopt for an out-of-scope plugin path.");

            return state.failure.empty();
        });

        SelfTest::RunCase(options,
                          suite,
                          L"try_make_relative_outside_root",
                          [&](SelfTest::CaseState& state) noexcept
        {
            // Case: TryMakeRelative returns nullopt for an absolute path not under the root.
            if (const auto foldersOpt = CreateCaseFolders(root, L"try_make_relative_outside_root"))
            {
                const auto& folders = *foldersOpt;
                auto session =
                    std::make_shared<CompareDirectoriesSession>(baseFs, baseFs, folders.left, folders.right, Common::Settings::CompareDirectoriesSettings{});

                const std::filesystem::path outsideLeft = folders.left.parent_path();
                const auto relLeft                      = session->TryMakeRelative(ComparePane::Left, outsideLeft);
                state.Require(! relLeft.has_value(), L"TryMakeRelative should return nullopt for a path outside the left root.");

                const std::filesystem::path outsideRight = folders.right.parent_path();
                const auto relRight                      = session->TryMakeRelative(ComparePane::Right, outsideRight);
                state.Require(! relRight.has_value(), L"TryMakeRelative should return nullopt for a path outside the right root.");
            }
            else
            {
                state.Require(false, L"Failed to create case folders: try_make_relative_outside_root.");
            }

            return state.failure.empty();
        });

        SelfTest::RunCase(options,
                          suite,
                          L"baseInterfaces",
                          [&](SelfTest::CaseState& state) noexcept
        {
            // Case: Base interface accessors return non-null objects after construction.
            if (const auto foldersOpt = CreateCaseFolders(root, L"baseInterfaces"))
            {
                const auto& folders = *foldersOpt;
                auto session =
                    std::make_shared<CompareDirectoriesSession>(baseFs, baseFs, folders.left, folders.right, Common::Settings::CompareDirectoriesSettings{});

                state.Require(static_cast<bool>(session->GetFileSystem(ComparePane::Left)), L"GetFileSystem(Left) should return non-null.");
                state.Require(static_cast<bool>(session->GetFileSystem(ComparePane::Right)), L"GetFileSystem(Right) should return non-null.");
                state.Require(static_cast<bool>(session->GetInformations(ComparePane::Left)), L"GetInformations(Left) should return non-null.");
                state.Require(static_cast<bool>(session->GetInformations(ComparePane::Right)), L"GetInformations(Right) should return non-null.");
                state.Require(static_cast<bool>(session->GetFileSystemIO(ComparePane::Left)), L"GetFileSystemIO(Left) should return non-null.");
                state.Require(static_cast<bool>(session->GetFileSystemIO(ComparePane::Right)), L"GetFileSystemIO(Right) should return non-null.");
                state.Require(session->IsContentCompareSupported(), L"IsContentCompareSupported() should return true.");
            }
            else
            {
                state.Require(false, L"Failed to create case folders: baseInterfaces.");
            }

            return state.failure.empty();
        });

        SelfTest::RunCase(options,
                          suite,
                          L"contentCacheHit",
                          [&](SelfTest::CaseState& state) noexcept
        {
            // Case: Repeated GetOrComputeDecision without invalidation returns the same cached object.
            if (const auto foldersOpt = CreateCaseFolders(root, L"contentCacheHit"))
            {
                const auto& folders = *foldersOpt;
                state.Require(SelfTest::WriteTextFile(folders.left / L"a.txt", "CacheA"), L"Failed to create a.txt (left).");
                state.Require(SelfTest::WriteTextFile(folders.right / L"a.txt", "CacheA"), L"Failed to create a.txt (right).");

                auto session =
                    std::make_shared<CompareDirectoriesSession>(baseFs, baseFs, folders.left, folders.right, Common::Settings::CompareDirectoriesSettings{});

                const auto decision1 = session->GetOrComputeDecision(std::filesystem::path{});
                state.Require(static_cast<bool>(decision1), L"First call should return a valid decision.");
                const auto decision2 = session->GetOrComputeDecision(std::filesystem::path{});
                state.Require(static_cast<bool>(decision2), L"Second call should return a valid decision.");

                // Without any intervening Invalidate(), both calls must return the identical cached shared_ptr.
                state.Require(decision1 == decision2, L"Repeated GetOrComputeDecision without invalidation must return the same cached decision.");
            }
            else
            {
                state.Require(false, L"Failed to create case folders: contentCacheHit.");
            }

            return state.failure.empty();
        });

        SelfTest::RunCase(options,
                          suite,
                          L"empty_directories",
                          [&](SelfTest::CaseState& state) noexcept
        {
            // Case: Empty directory roots produce an empty decision.
            if (const auto foldersOpt = CreateCaseFolders(root, L"empty_directories"))
            {
                const auto& folders = *foldersOpt;

                auto decision = ComputeRootDecision(baseFs, folders, Common::Settings::CompareDirectoriesSettings{}, state);
                if (decision)
                {
                    state.Require(decision->items.empty(), L"Empty roots expected decision.items empty.");
                    state.Require(! decision->anyDifferent, L"Empty roots expected anyDifferent=false.");
                }
            }
            else
            {
                state.Require(false, L"Failed to create case folders: empty_directories.");
            }

            return state.failure.empty();
        });

        SelfTest::RunCase(options,
                          suite,
                          L"zeroByteContent",
                          [&](SelfTest::CaseState& state) noexcept
        {
            // Case: compareContent=true on two zero-byte files reports them as identical.
            if (const auto foldersOpt = CreateCaseFolders(root, L"zeroByteContent"))
            {
                const auto& folders = *foldersOpt;
                // Create empty files on both sides.
                state.Require(SelfTest::WriteBinaryFile(folders.left / L"empty.txt", {}), L"Failed to create empty.txt (left).");
                state.Require(SelfTest::WriteBinaryFile(folders.right / L"empty.txt", {}), L"Failed to create empty.txt (right).");

                Common::Settings::CompareDirectoriesSettings settings{};
                settings.compareContent     = true;
                settings.keepIdenticalItems = true;

                auto decision = ComputeRootDecision(baseFs, folders, settings, state);
                if (decision)
                {
                    const auto* item = FindItem(*decision, L"empty.txt");
                    state.Require(item != nullptr, L"empty.txt should appear in the decision.");
                    if (item)
                    {
                        state.Require(! item->isDifferent, L"Zero-byte files on both sides must be identical.");
                        state.Require(! HasFlag(item->differenceMask, CompareDirectoriesDiffBit::Content),
                                      L"Zero-byte files must not have the Content diff bit set.");
                    }
                }
            }
            else
            {
                state.Require(false, L"Failed to create case folders: zeroByteContent.");
            }

            return state.failure.empty();
        });

        SelfTest::RunCase(options,
                          suite,
                          L"setSettingsInvalidates",
                          [&](SelfTest::CaseState& state) noexcept
        {
            // Case: SetSettings with a comparison-changing setting increments GetVersion(); a view-only toggle does not.
            if (const auto foldersOpt = CreateCaseFolders(root, L"setSettingsInvalidates"))
            {
                const auto& folders = *foldersOpt;
                state.Require(SelfTest::WriteTextFile(folders.left / L"a.txt", "V"), L"Failed to create a.txt (left).");
                state.Require(SelfTest::WriteTextFile(folders.right / L"a.txt", "V"), L"Failed to create a.txt (right).");

                Common::Settings::CompareDirectoriesSettings settings{};
                settings.compareContent = false;
                auto session            = std::make_shared<CompareDirectoriesSession>(baseFs, baseFs, folders.left, folders.right, settings);

                const uint64_t v0 = session->GetVersion();

                // keepIdenticalItems affects cached decision shape (pruning), so it must invalidate.
                settings.keepIdenticalItems = true;
                session->SetSettings(settings);
                const uint64_t v1 = session->GetVersion();
                state.Require(v1 != v0, L"SetSettings with keepIdenticalItems toggled must increment GetVersion().");

                // showIdenticalItems is view-only and must NOT invalidate when keepIdenticalItems is unchanged.
                settings.showIdenticalItems = true;
                session->SetSettings(settings);
                const uint64_t v2 = session->GetVersion();
                state.Require(v2 == v1, L"SetSettings with showIdenticalItems toggled must not increment GetVersion().");

                // Changing compareContent must invalidate the cache (version bump).
                settings.compareContent = true;
                session->SetSettings(settings);
                const uint64_t v3 = session->GetVersion();
                state.Require(v3 != v2, L"SetSettings with compareContent toggled must increment GetVersion().");

                // Setting the same value again must NOT bump the version.
                session->SetSettings(settings);
                const uint64_t v4 = session->GetVersion();
                state.Require(v4 == v3, L"SetSettings with identical settings must not increment GetVersion().");

                // Changing compareSize must also invalidate.
                settings.compareSize = ! settings.compareSize;
                session->SetSettings(settings);
                const uint64_t v5 = session->GetVersion();
                state.Require(v5 != v4, L"SetSettings with compareSize toggled must increment GetVersion().");
            }
            else
            {
                state.Require(false, L"Failed to create case folders: setSettingsInvalidates.");
            }

            return state.failure.empty();
        });

        SelfTest::RunCase(options,
                          suite,
                          L"dircache_not_polluted_by_compare_scan",
                          [&](SelfTest::CaseState& state) noexcept
        {
            // Case: Compare scans must not populate DirectoryInfoCache (memory blow-up regression guard).
            if (const auto foldersOpt = CreateCaseFolders(root, L"dircache_not_polluted_by_compare_scan"))
            {
                const auto& folders = *foldersOpt;

                constexpr size_t kDirCount  = 64;
                constexpr size_t kFileCount = 10;

                for (size_t i = 0; i < kDirCount; ++i)
                {
                    const std::wstring dirName = std::format(L"d{:03}", i);

                    const std::filesystem::path leftSub  = folders.left / dirName / L"sub";
                    const std::filesystem::path rightSub = folders.right / dirName / L"sub";

                    state.Require(SelfTest::EnsureDirectory(leftSub), std::format(L"Failed to create {} (left).", leftSub.wstring()));
                    state.Require(SelfTest::EnsureDirectory(rightSub), std::format(L"Failed to create {} (right).", rightSub.wstring()));

                    for (size_t j = 0; j < kFileCount; ++j)
                    {
                        const std::wstring fileName = std::format(L"f{:03}.txt", j);
                        state.Require(SelfTest::WriteTextFile(leftSub / fileName, "X"), std::format(L"Failed to create {} (left).", fileName));
                        state.Require(SelfTest::WriteTextFile(rightSub / fileName, "X"), std::format(L"Failed to create {} (right).", fileName));
                    }
                }

                DirectoryInfoCache& cache = DirectoryInfoCache::GetInstance();
                cache.ClearForFileSystem(baseFs.get());

                const DirectoryInfoCache::Stats before = cache.GetStats();

                Common::Settings::CompareDirectoriesSettings settings{};
                settings.compareSubdirectories = true;

                auto session = std::make_shared<CompareDirectoriesSession>(baseFs, baseFs, folders.left, folders.right, settings);
                state.Require(StartScanAndWaitForIdle(session, std::chrono::milliseconds{SelfTest::ScaleTimeout(30'000)}),
                              L"DirectoryInfoCache regression: scan did not become idle within timeout.");

                const DirectoryInfoCache::Stats after = cache.GetStats();

                const uint64_t deltaEnumerations = after.enumerations - before.enumerations;
                state.Require(deltaEnumerations == 0u,
                              std::format(L"DirectoryInfoCache regression: expected enumerations delta=0, got {}.", deltaEnumerations));
            }
            else
            {
                state.Require(false, L"Failed to create case folders: dircache_not_polluted_by_compare_scan.");
            }

            return state.failure.empty();
        });

        SelfTest::RunCase(options,
                          suite,
                          L"content_queue_bounded_hi_lo",
                          [&](SelfTest::CaseState& state) noexcept
        {
            // Case: Content compare job queues are bounded (hi/lo) under load; scan workers backpressure instead of OOM.
            if (const auto foldersOpt = CreateCaseFolders(root, L"content_queue_bounded_hi_lo"))
            {
                const auto& folders = *foldersOpt;

                constexpr size_t kRootFiles = 1000;
                constexpr size_t kHotFiles  = 200;

                state.Require(SelfTest::EnsureDirectory(folders.left / L"hot"), L"Failed to create hot (left).");
                state.Require(SelfTest::EnsureDirectory(folders.right / L"hot"), L"Failed to create hot (right).");

                for (size_t i = 0; i < kRootFiles; ++i)
                {
                    const std::wstring name = std::format(L"r{:04}.bin", i);
                    state.Require(SelfTest::WriteTextFile(folders.left / name, "AAAA"), std::format(L"Failed to write {} (left).", name));
                    state.Require(SelfTest::WriteTextFile(folders.right / name, "BBBB"), std::format(L"Failed to write {} (right).", name));
                }

                for (size_t i = 0; i < kHotFiles; ++i)
                {
                    const std::wstring name = std::format(L"h{:04}.bin", i);
                    state.Require(SelfTest::WriteTextFile(folders.left / L"hot" / name, "AAAA"), std::format(L"Failed to write hot\\{} (left).", name));
                    state.Require(SelfTest::WriteTextFile(folders.right / L"hot" / name, "BBBB"), std::format(L"Failed to write hot\\{} (right).", name));
                }

                Common::Settings::CompareDirectoriesSettings settings{};
                settings.compareContent    = true;
                settings.compareSize       = false;
                settings.compareDateTime   = false;
                settings.compareAttributes = false;
                settings.keepIdenticalItems = true;

                auto session = std::make_shared<CompareDirectoriesSession>(baseFs, baseFs, folders.left, folders.right, settings);

                std::mutex mutex;
                std::condition_variable cv;
                bool started = false;
                bool done    = false;

                session->SetScanProgressCallback(
                    [&](const std::filesystem::path&, std::wstring_view, uint64_t, uint64_t, uint32_t activeScans, uint64_t, uint64_t) noexcept
                {
                    std::lock_guard lock(mutex);
                    if (activeScans != 0u)
                    {
                        started = true;
                    }
                    if (started && activeScans == 0u)
                    {
                        done = true;
                        cv.notify_all();
                    }
                });

                session->RequestScanForFolder(std::filesystem::path(L"hot"));
                session->StartScan();

                {
                    std::unique_lock lock(mutex);
                    static_cast<void>(cv.wait_for(lock, std::chrono::milliseconds{SelfTest::ScaleTimeout(60'000)}, [&] { return done; }));
                }

                session->SetScanProgressCallback({});
                state.Require(done, L"content_queue_bounded_hi_lo: scan did not become idle within timeout.");

                const CompareDirectoriesPerfStats stats = session->GetPerfStats();

                constexpr size_t kMaxHighJobs = 128;
                constexpr size_t kMaxLowJobs  = 896;
                constexpr size_t kMaxTotal    = kMaxHighJobs + kMaxLowJobs;

                state.Require(stats.contentQueueHighHighWater <= kMaxHighJobs,
                              std::format(L"High content queue exceeded cap: {} > {}.", stats.contentQueueHighHighWater, kMaxHighJobs));
                state.Require(stats.contentQueueLowHighWater <= kMaxLowJobs,
                              std::format(L"Low content queue exceeded cap: {} > {}.", stats.contentQueueLowHighWater, kMaxLowJobs));
                state.Require(stats.contentQueueHighWater <= kMaxTotal,
                              std::format(L"Total content queue exceeded cap: {} > {}.", stats.contentQueueHighWater, kMaxTotal));
            }
            else
            {
                state.Require(false, L"Failed to create case folders: content_queue_bounded_hi_lo.");
            }

            return state.failure.empty();
        });

        SelfTest::RunCase(options,
                          suite,
                          L"decision_cache_eviction_budget_pins_visible",
                          [&](SelfTest::CaseState& state) noexcept
        {
            // Case: Decision cache eviction respects pinned visible folders (prevents UI thrash).
            if (const auto foldersOpt = CreateCaseFolders(root, L"decision_cache_eviction_budget_pins_visible"))
            {
                const auto& folders = *foldersOpt;

                state.Require(SelfTest::EnsureDirectory(folders.left / L"keep"), L"Failed to create keep (left).");
                state.Require(SelfTest::EnsureDirectory(folders.right / L"keep"), L"Failed to create keep (right).");
                state.Require(SelfTest::WriteTextFile(folders.left / L"keep" / L"keep.txt", "K"), L"Failed to write keep.txt (left).");
                state.Require(SelfTest::WriteTextFile(folders.right / L"keep" / L"keep.txt", "K"), L"Failed to write keep.txt (right).");

                state.Require(SelfTest::EnsureDirectory(folders.left / L"spill"), L"Failed to create spill (left).");
                state.Require(SelfTest::EnsureDirectory(folders.right / L"spill"), L"Failed to create spill (right).");

                constexpr size_t kSpillDirs   = 120;
                constexpr size_t kFilesPerDir = 10;

                for (size_t i = 0; i < kSpillDirs; ++i)
                {
                    const std::wstring dirName           = std::format(L"d{:04}", i);
                    const std::filesystem::path leftDir  = folders.left / L"spill" / dirName;
                    const std::filesystem::path rightDir = folders.right / L"spill" / dirName;

                    state.Require(SelfTest::EnsureDirectory(leftDir), std::format(L"Failed to create {} (left).", leftDir.wstring()));
                    state.Require(SelfTest::EnsureDirectory(rightDir), std::format(L"Failed to create {} (right).", rightDir.wstring()));

                    for (size_t j = 0; j < kFilesPerDir; ++j)
                    {
                        const std::wstring fileName = std::format(L"f{:03}.txt", j);
                        state.Require(SelfTest::WriteTextFile(leftDir / fileName, "X"), std::format(L"Failed to write spill file {} (left).", fileName));
                        state.Require(SelfTest::WriteTextFile(rightDir / fileName, "X"), std::format(L"Failed to write spill file {} (right).", fileName));
                    }
                }

                Common::Settings::CompareDirectoriesSettings settings{};
                settings.keepIdenticalItems = true;

                auto session = std::make_shared<CompareDirectoriesSession>(baseFs, baseFs, folders.left, folders.right, settings);
                session->SetDecisionCacheBudgetBytesForSelfTest(64u * 1024u);
                session->SetPinnedFolders(std::filesystem::path(L"keep"), std::filesystem::path(L"keep"));

                std::mutex mutex;
                std::condition_variable cv;
                bool started = false;
                bool done    = false;

                session->SetScanProgressCallback(
                    [&](const std::filesystem::path&, std::wstring_view, uint64_t, uint64_t, uint32_t activeScans, uint64_t, uint64_t) noexcept
                {
                    std::lock_guard lock(mutex);
                    if (activeScans != 0u)
                    {
                        started = true;
                    }
                    if (started && activeScans == 0u)
                    {
                        done = true;
                        cv.notify_all();
                    }
                });

                session->RequestScanForFolder(std::filesystem::path(L"keep"));
                for (size_t i = 0; i < kSpillDirs; ++i)
                {
                    const std::wstring dirName = std::format(L"d{:04}", i);
                    session->RequestScanForFolder(std::filesystem::path(L"spill") / dirName);
                }
                session->StartScan();

                {
                    std::unique_lock lock(mutex);
                    static_cast<void>(cv.wait_for(lock, std::chrono::milliseconds{SelfTest::ScaleTimeout(60'000)}, [&] { return done; }));
                }

                session->SetScanProgressCallback({});
                state.Require(done, L"decision_cache_eviction_budget_pins_visible: scan did not become idle within timeout.");

                const CompareDirectoriesPerfStats stats = session->GetPerfStats();
                state.Require(stats.decisionCacheEntriesHighWater > stats.decisionCacheEntries,
                              L"decision cache eviction expected (high-water should exceed current entries).");

                const auto keepDecision = session->TryGetCachedDecision(std::filesystem::path(L"keep"));
                state.Require(static_cast<bool>(keepDecision), L"Pinned folder decision (keep) must remain cached.");

                size_t evicted = 0;
                for (size_t i = 0; i < 16; ++i)
                {
                    const std::wstring dirName = std::format(L"d{:04}", i);
                    if (! session->TryGetCachedDecision(std::filesystem::path(L"spill") / dirName))
                    {
                        ++evicted;
                    }
                }
                state.Require(evicted != 0u, L"Expected at least one spill folder decision to be evicted under small budget.");
            }
            else
            {
                state.Require(false, L"Failed to create case folders: decision_cache_eviction_budget_pins_visible.");
            }

            return state.failure.empty();
        });

        SelfTest::RunCase(options,
                          suite,
                          L"cancel_completes_bounded",
                          [&](SelfTest::CaseState& state) noexcept
        {
            // Case: Disabling background work cancels scan/content work promptly (exit/cancel regression guard).
            if (const auto foldersOpt = CreateCaseFolders(root, L"cancel_completes_bounded"))
            {
                const auto& folders = *foldersOpt;

                constexpr size_t kBytes = 16u * 1024u * 1024u;
                state.Require(WriteFileFill(folders.left / L"big.bin", 'A', kBytes), L"Failed to create big.bin (left).");
                state.Require(WriteFileFill(folders.right / L"big.bin", 'B', kBytes), L"Failed to create big.bin (right).");

                Common::Settings::CompareDirectoriesSettings settings{};
                settings.compareContent    = true;
                settings.compareSize       = false;
                settings.compareDateTime   = false;
                settings.compareAttributes = false;

                wil::com_ptr<IFileSystem> wrapped = CreateShortReadFileSystem(baseFs, folders.left, 4096u, 1u);
                state.Require(static_cast<bool>(wrapped), L"Failed to create short-read file system wrapper (cancel).");

                const wil::com_ptr<IFileSystem> compareFs = wrapped ? wrapped : baseFs;
                auto session = std::make_shared<CompareDirectoriesSession>(compareFs, compareFs, folders.left, folders.right, settings);

                static_cast<void>(session->GetOrComputeDecision(std::filesystem::path{}));

                const auto startedDeadline =
                    std::chrono::steady_clock::now() + std::chrono::milliseconds{static_cast<std::chrono::milliseconds::rep>(SelfTest::ScaleTimeout(5'000))};
                bool sawPending = false;
                while (std::chrono::steady_clock::now() < startedDeadline)
                {
                    const CompareDirectoriesPerfStats stats = session->GetPerfStats();
                    if (stats.contentPendingCompares != 0u || stats.contentInFlightSize != 0u || stats.contentQueueSize != 0u)
                    {
                        sawPending = true;
                        break;
                    }
                    std::this_thread::sleep_for(std::chrono::milliseconds(10));
                }

                state.Require(sawPending, L"cancel_completes_bounded: expected pending content compare work.");

                session->SetBackgroundWorkEnabled(false);

                const auto doneDeadline =
                    std::chrono::steady_clock::now() + std::chrono::milliseconds{static_cast<std::chrono::milliseconds::rep>(SelfTest::ScaleTimeout(10'000))};
                bool canceled = false;
                while (std::chrono::steady_clock::now() < doneDeadline)
                {
                    const CompareDirectoriesPerfStats stats = session->GetPerfStats();
                    if (stats.scanActiveScans == 0u && stats.contentPendingCompares == 0u && stats.contentQueueSize == 0u && stats.contentInFlightSize == 0u)
                    {
                        canceled = true;
                        break;
                    }
                    std::this_thread::sleep_for(std::chrono::milliseconds(10));
                }

                state.Require(canceled, L"cancel_completes_bounded: background work did not cancel/drain within timeout.");
            }
            else
            {
                state.Require(false, L"Failed to create case folders: cancel_completes_bounded.");
            }

            return state.failure.empty();
        });

        SelfTest::RunCase(options,
                          suite,
                          L"invalid_directory_entry_buffer",
                          [&](SelfTest::CaseState& state) noexcept
        {
            // Case: Corrupt FileInfo buffers are rejected with ERROR_INVALID_DATA (no crash / OOB walk).
            if (const auto foldersOpt = CreateCaseFolders(root, L"invalid_directory_entry_buffer"))
            {
                const auto& folders = *foldersOpt;
                state.Require(SelfTest::WriteTextFile(folders.left / L"a.txt", "L"), L"Failed to create a.txt (left).");
                state.Require(SelfTest::WriteTextFile(folders.right / L"a.txt", "R"), L"Failed to create a.txt (right).");

                ReadDirectoryTestBehavior behavior{};
                behavior.targetPath      = folders.left;
                behavior.returnMalformed = true;

                wil::com_ptr<IFileSystem> malformedLeft = CreateReadDirectoryBehaviorFileSystem(baseFs, behavior);
                state.Require(static_cast<bool>(malformedLeft), L"Failed to create malformed ReadDirectoryInfo wrapper.");

                auto session = std::make_shared<CompareDirectoriesSession>(
                    malformedLeft ? malformedLeft : baseFs, baseFs, folders.left, folders.right, Common::Settings::CompareDirectoriesSettings{});

                const auto decision = session->GetOrComputeDecision(std::filesystem::path{});
                state.Require(static_cast<bool>(decision), L"invalid_directory_entry_buffer: decision is null.");
                if (decision)
                {
                    state.Require(FAILED(decision->hr), L"invalid_directory_entry_buffer: expected failure HRESULT.");
                    state.Require(decision->hr == HRESULT_FROM_WIN32(ERROR_INVALID_DATA),
                                  std::format(L"invalid_directory_entry_buffer: expected ERROR_INVALID_DATA, got 0x{:08X}.",
                                              static_cast<unsigned int>(decision->hr)));
                }
            }
            else
            {
                state.Require(false, L"Failed to create case folders: invalid_directory_entry_buffer.");
            }

            return state.failure.empty();
        });

        SelfTest::RunCase(options,
                          suite,
                          L"scan_inflight_stamp_guards_restart",
                          [&](SelfTest::CaseState& state) noexcept
        {
            // Case: Stale scan worker completion must not clear active tracking for a restarted run.
            if (const auto foldersOpt = CreateCaseFolders(root, L"scan_inflight_stamp_guards_restart"))
            {
                const auto& folders = *foldersOpt;
                state.Require(SelfTest::WriteTextFile(folders.left / L"a.txt", "L"), L"Failed to create a.txt (left).");
                state.Require(SelfTest::WriteTextFile(folders.right / L"a.txt", "R"), L"Failed to create a.txt (right).");

                ReadDirectoryTestBehavior behavior{};
                behavior.targetPath = folders.left;
                behavior.delayMs    = static_cast<DWORD>(SelfTest::ScaleTimeout(1200));

                wil::com_ptr<IFileSystem> delayedLeft = CreateReadDirectoryBehaviorFileSystem(baseFs, behavior);
                state.Require(static_cast<bool>(delayedLeft), L"Failed to create delayed ReadDirectoryInfo wrapper.");

                auto session = std::make_shared<CompareDirectoriesSession>(
                    delayedLeft ? delayedLeft : baseFs, baseFs, folders.left, folders.right, Common::Settings::CompareDirectoriesSettings{});

                session->StartScan();

                const auto startDeadline =
                    std::chrono::steady_clock::now() + std::chrono::milliseconds{static_cast<std::chrono::milliseconds::rep>(SelfTest::ScaleTimeout(5'000))};
                bool sawActive = false;
                while (std::chrono::steady_clock::now() < startDeadline)
                {
                    if (session->GetPerfStats().scanActiveScans != 0u)
                    {
                        sawActive = true;
                        break;
                    }
                    std::this_thread::sleep_for(std::chrono::milliseconds(10));
                }
                state.Require(sawActive, L"scan_inflight_stamp_guards_restart: initial scan did not become active.");

                std::this_thread::sleep_for(std::chrono::milliseconds(SelfTest::ScaleTimeout(800)));

                session->SetBackgroundWorkEnabled(false);
                session->InvalidateForAbsolutePath(folders.left, true);
                session->SetBackgroundWorkEnabled(true);
                session->StartScan();

                // This point is after stale run completion should have happened, but before the restarted run can finish.
                std::this_thread::sleep_for(std::chrono::milliseconds(SelfTest::ScaleTimeout(1000)));
                const CompareDirectoriesPerfStats midStats = session->GetPerfStats();
                state.Require(midStats.scanActiveScans != 0u,
                              L"scan_inflight_stamp_guards_restart: restarted scan lost active tracking before completion.");

                const auto doneDeadline =
                    std::chrono::steady_clock::now() + std::chrono::milliseconds{static_cast<std::chrono::milliseconds::rep>(SelfTest::ScaleTimeout(15'000))};
                bool idle = false;
                while (std::chrono::steady_clock::now() < doneDeadline)
                {
                    const CompareDirectoriesPerfStats stats = session->GetPerfStats();
                    if (stats.scanActiveScans == 0u && stats.scanQueueSize == 0u && stats.scanInFlightKeys == 0u)
                    {
                        idle = true;
                        break;
                    }
                    std::this_thread::sleep_for(std::chrono::milliseconds(10));
                }
                state.Require(idle, L"scan_inflight_stamp_guards_restart: restarted scan did not become idle within timeout.");
            }
            else
            {
                state.Require(false, L"Failed to create case folders: scan_inflight_stamp_guards_restart.");
            }

            return state.failure.empty();
        });

        SelfTest::RunCase(options,
                          suite,
                          L"content_inflight_stamp_guards_restart",
                          [&](SelfTest::CaseState& state) noexcept
        {
            // Case: Stale content worker completion must not consume a restarted run's in-flight slot.
            if (const auto foldersOpt = CreateCaseFolders(root, L"content_inflight_stamp_guards_restart"))
            {
                const auto& folders = *foldersOpt;
                constexpr size_t kBytes = 8u * 1024u * 1024u;
                state.Require(WriteFileFill(folders.left / L"a.bin", 'A', kBytes), L"Failed to create a.bin (left).");
                state.Require(WriteFileFill(folders.right / L"a.bin", 'B', kBytes), L"Failed to create a.bin (right).");

                Common::Settings::CompareDirectoriesSettings settings{};
                settings.compareContent    = true;
                settings.compareSize       = false;
                settings.compareDateTime   = false;
                settings.compareAttributes = false;
                settings.keepIdenticalItems = true;

                wil::com_ptr<IFileSystem> wrapped = CreateShortReadFileSystem(baseFs, folders.left, 4096u, static_cast<DWORD>(SelfTest::ScaleTimeout(2)));
                state.Require(static_cast<bool>(wrapped), L"Failed to create short-read file system wrapper (content restart stamp).");

                const wil::com_ptr<IFileSystem> compareFs = wrapped ? wrapped : baseFs;
                auto session = std::make_shared<CompareDirectoriesSession>(compareFs, compareFs, folders.left, folders.right, settings);

                static_cast<void>(session->GetOrComputeDecision(std::filesystem::path{}));

                const auto startDeadline =
                    std::chrono::steady_clock::now() + std::chrono::milliseconds{static_cast<std::chrono::milliseconds::rep>(SelfTest::ScaleTimeout(10'000))};
                bool sawInFlight = false;
                while (std::chrono::steady_clock::now() < startDeadline)
                {
                    const CompareDirectoriesPerfStats stats = session->GetPerfStats();
                    if (stats.contentPendingCompares != 0u && stats.contentInFlightSize != 0u)
                    {
                        sawInFlight = true;
                        break;
                    }
                    std::this_thread::sleep_for(std::chrono::milliseconds(10));
                }
                state.Require(sawInFlight, L"content_inflight_stamp_guards_restart: initial content compare did not become active.");

                session->SetBackgroundWorkEnabled(false);
                session->InvalidateForAbsolutePath(folders.left, true);
                session->SetBackgroundWorkEnabled(true);
                session->StartScan();

                const auto restartedDecision = session->GetOrComputeDecision(std::filesystem::path{});
                state.Require(static_cast<bool>(restartedDecision),
                              L"content_inflight_stamp_guards_restart: restarted decision is null.");
                if (restartedDecision)
                {
                    const auto* restartedItem = FindItem(*restartedDecision, L"a.bin");
                    state.Require(restartedItem != nullptr,
                                  L"content_inflight_stamp_guards_restart: a.bin missing from restarted decision.");
                    if (restartedItem)
                    {
                        state.Require(HasFlag(restartedItem->differenceMask, CompareDirectoriesDiffBit::ContentPending),
                                      L"content_inflight_stamp_guards_restart: restarted decision did not re-enter ContentPending.");
                    }
                }

                const auto restartedDeadline =
                    std::chrono::steady_clock::now() + std::chrono::milliseconds{static_cast<std::chrono::milliseconds::rep>(SelfTest::ScaleTimeout(10'000))};
                bool restartedObserved = false;
                while (std::chrono::steady_clock::now() < restartedDeadline)
                {
                    const CompareDirectoriesPerfStats stats = session->GetPerfStats();
                    if (stats.contentPendingCompares != 0u || stats.contentInFlightSize != 0u)
                    {
                        restartedObserved = true;
                        break;
                    }
                    std::this_thread::sleep_for(std::chrono::milliseconds(10));
                }
                state.Require(restartedObserved,
                              L"content_inflight_stamp_guards_restart: restarted content compare did not become pending or in-flight.");

                const auto doneDeadline =
                    std::chrono::steady_clock::now() + std::chrono::milliseconds{static_cast<std::chrono::milliseconds::rep>(SelfTest::ScaleTimeout(30'000))};
                bool done = false;
                while (std::chrono::steady_clock::now() < doneDeadline)
                {
                    const CompareDirectoriesPerfStats stats = session->GetPerfStats();
                    if (stats.scanActiveScans == 0u && stats.contentPendingCompares == 0u && stats.contentQueueSize == 0u && stats.contentInFlightSize == 0u)
                    {
                        done = true;
                        break;
                    }
                    std::this_thread::sleep_for(std::chrono::milliseconds(10));
                }
                state.Require(done, L"content_inflight_stamp_guards_restart: compare did not complete within timeout.");

                session->FlushPendingContentCompareUpdates();
                const auto decision = session->GetOrComputeDecision(std::filesystem::path{});
                state.Require(static_cast<bool>(decision), L"content_inflight_stamp_guards_restart: final decision is null.");
                if (decision)
                {
                    const auto* item = FindItem(*decision, L"a.bin");
                    state.Require(item != nullptr, L"content_inflight_stamp_guards_restart: a.bin missing from final decision.");
                    if (item)
                    {
                        state.Require(! HasFlag(item->differenceMask, CompareDirectoriesDiffBit::ContentPending),
                                      L"content_inflight_stamp_guards_restart: ContentPending should be cleared after completion.");
                        state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::Content),
                                      L"content_inflight_stamp_guards_restart: expected Content difference after completion.");
                        state.Require(item->isDifferent, L"content_inflight_stamp_guards_restart: a.bin should be marked different.");
                    }
                }
            }
            else
            {
                state.Require(false, L"Failed to create case folders: content_inflight_stamp_guards_restart.");
            }

            return state.failure.empty();
        });

        SelfTest::RunCase(options,
                          suite,
                          L"directory_size_local_callback_contract",
                          [&](SelfTest::CaseState& state) noexcept
        {
            if (const auto foldersOpt = CreateCaseFolders(root, L"directory_size_local_callback_contract"))
            {
                const auto& folders = *foldersOpt;
                const std::filesystem::path filePath = folders.left / L"probe.bin";
                state.Require(WriteFileFill(filePath, 'L', 13u), L"Directory size local callback: failed to create probe file.");

                wil::com_ptr<IFileSystemDirectoryOperations> dirOps;
                state.Require(CreateFileSystemDirectoryOperations(baseFs, dirOps), L"Directory size local callback: missing IFileSystemDirectoryOperations.");
                if (dirOps)
                {
                    static_cast<void>(ValidateDirectorySizeCallbackContract(state, dirOps.get(), filePath.wstring()));
                }
            }
            else
            {
                state.Require(false, L"Directory size local callback: failed to create case folders.");
            }

            return state.failure.empty();
        });

        AppendCaseResult(
            suite, L"directory_size_local_overflow", SelfTest::SelfTestCaseResult::Status::skipped, L"no deterministic overflow harness for the local filesystem");

        SelfTest::RunCase(options,
                          suite,
                          L"directory_size_dummy_callback_contract",
                          [&](SelfTest::CaseState& state) noexcept
        {
            state.Require(dummyFs && dummyIo && dummyOps, L"Directory size dummy callback: dummy filesystem setup is unavailable.");
            if (! dummyFs || ! dummyIo || ! dummyOps)
            {
                return false;
            }

            const std::filesystem::path dirPath(std::format(L"/directory-size-dummy-{}", guid));
            const std::filesystem::path filePath = dirPath / L"probe.txt";
            const std::wstring fileText          = ToPluginPathText(filePath);
            const auto cleanup = wil::scope_exit([&]() noexcept
            {
                static_cast<void>(
                    dummyFs->DeleteItem(dirPath.c_str(), static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_CONTINUE_ON_ERROR), nullptr, nullptr, nullptr));
            });

            state.Require(EnsureDirectoryExistsFsOps(dummyOps, dirPath), L"Directory size dummy callback: failed to create dummy sandbox directory.");
            state.Require(WriteFileTextFsIo(dummyIo, filePath, "dummy-directory-size"), L"Directory size dummy callback: failed to write probe file.");
            if (! state.failure.empty())
            {
                return false;
            }

            static_cast<void>(ValidateDirectorySizeCallbackContract(state, dummyOps.get(), fileText));
            return state.failure.empty();
        });

        SelfTest::RunCase(options,
                          suite,
                          L"directory_size_7z_callback_contract",
                          [&](SelfTest::CaseState& state) noexcept
        {
            CreatedFileSystemInstance archiveCreated{};
            const HRESULT createHr = TryCreateFileSystemInstance(kBuiltin7zFileSystemId, {}, archiveCreated);
            state.Require(SUCCEEDED(createHr) && archiveCreated.fileSystem,
                          std::format(L"Directory size 7z callback: failed to create filesystem instance. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
            if (FAILED(createHr) || ! archiveCreated.fileSystem)
            {
                return false;
            }

            wil::com_ptr<IFileSystemInitialize> init;
            const HRESULT initQiHr = archiveCreated.fileSystem->QueryInterface(IID_PPV_ARGS(init.put()));
            state.Require(SUCCEEDED(initQiHr) && init, L"Directory size 7z callback: missing IFileSystemInitialize.");
            if (FAILED(initQiHr) || ! init)
            {
                return false;
            }

            wil::com_ptr<IFileSystemDirectoryOperations> dirOps;
            const HRESULT dirQiHr = archiveCreated.fileSystem->QueryInterface(IID_PPV_ARGS(dirOps.put()));
            state.Require(SUCCEEDED(dirQiHr) && dirOps, L"Directory size 7z callback: missing IFileSystemDirectoryOperations.");
            if (FAILED(dirQiHr) || ! dirOps)
            {
                return false;
            }

            const std::filesystem::path archivePath = GetWorkspaceRootFromSourcePath() / L"Plugins" / L"FileSystem7z" / L"Tests" / L"Tests.zip";
            state.Require(SelfTest::PathExists(archivePath), std::format(L"Directory size 7z callback: fixture archive missing: {}", archivePath.wstring()));
            if (! SelfTest::PathExists(archivePath))
            {
                return false;
            }

            const HRESULT initHr = init->Initialize(archivePath.c_str(), nullptr);
            state.Require(SUCCEEDED(initHr),
                          std::format(L"Directory size 7z callback: Initialize failed. hr=0x{:08X}", static_cast<unsigned long>(initHr)));
            if (FAILED(initHr))
            {
                return false;
            }

            const auto filePathOpt = FindFirstRegularEntryPath(archiveCreated.fileSystem, L"/");
            state.Require(filePathOpt.has_value(), L"Directory size 7z callback: no regular file found in fixture archive.");
            if (! filePathOpt.has_value())
            {
                return false;
            }

            static_cast<void>(ValidateDirectorySizeCallbackContract(state, dirOps.get(), filePathOpt.value()));
            return state.failure.empty();
        });

        const auto runRemoteFileCompare = [&](std::wstring_view caseName,
                                              std::wstring_view protocolLabel,
                                              std::wstring_view envVarName,
                                              std::wstring_view defaultProfileName,
                                              std::wstring_view pluginId) noexcept
        {
            if (options.failFast && suite.failed != 0)
            {
                AppendCaseResult(suite, caseName, SelfTest::SelfTestCaseResult::Status::skipped, L"not executed (fail-fast)");
                return;
            }

            const PhaseCheckResult secretOutcome = CheckRemoteConnectionSecret(protocolLabel, envVarName, defaultProfileName, pluginId);
            if (secretOutcome.status != SelfTest::SelfTestCaseResult::Status::passed)
            {
                AppendCaseResult(suite, caseName, secretOutcome.status, secretOutcome.reason);
                return;
            }

            const PhaseCheckResult sandboxOutcome = CheckRemoteConnectionSandbox(protocolLabel, envVarName, defaultProfileName, pluginId);
            if (sandboxOutcome.status != SelfTest::SelfTestCaseResult::Status::passed)
            {
                AppendCaseResult(suite, caseName, sandboxOutcome.status, sandboxOutcome.reason);
                return;
            }

            SelfTest::RunCase(options,
                              suite,
                              caseName,
                              [&](SelfTest::CaseState& state) noexcept
            {
                const std::wstring overrideName                    = GetEnvVarTrimmed(envVarName);
                const std::wstring profileName                     = ! overrideName.empty() ? overrideName : std::wstring(defaultProfileName);
    const Common::Settings::ConnectionProfile* profile = ConnectionProfileUtils::FindConnectionProfileByName(&g_settings, profileName);
                state.Require(profile != nullptr, L"Remote compare: profile missing after preconditions passed.");
                if (! profile)
                {
                    return false;
                }

                const std::wstring initialPath = NormalizePluginPathForSelfTest(profile->initialPath);
                state.Require(! initialPath.empty() && initialPath[0] == L'/', L"Remote compare: initialPath must be an absolute plugin path.");
                if (initialPath.empty() || initialPath[0] != L'/')
                {
                    return false;
                }

                const std::filesystem::path remoteRoot = std::filesystem::path(std::format(L"/@conn:{}{}", profileName, initialPath));

                const std::filesystem::path localRoot = root / std::wstring(caseName) / L"left";
                state.Require(SelfTest::EnsureDirectory(localRoot), L"Remote compare: failed to create local root folder.");

                const std::wstring uniqueName = std::format(L"only_left_{}.txt", guid);
                state.Require(SelfTest::WriteTextFile(localRoot / uniqueName, "L"), L"Remote compare: failed to write local test file.");

                CreatedFileSystemInstance remoteCreated{};
                const HRESULT createHr = TryCreateFileSystemInstance(pluginId, {}, remoteCreated);
                state.Require(SUCCEEDED(createHr),
                              std::format(L"Remote compare: failed to create filesystem instance. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
                if (FAILED(createHr) || ! remoteCreated.fileSystem)
                {
                    return false;
                }

                Common::Settings::CompareDirectoriesSettings settings{};
                settings.compareContent = true;

                auto session  = std::make_shared<CompareDirectoriesSession>(baseFs, remoteCreated.fileSystem, localRoot, remoteRoot, settings);
                auto decision = session->GetOrComputeDecision(std::filesystem::path{});
                state.Require(static_cast<bool>(decision), L"Remote compare: decision is null.");
                if (! decision)
                {
                    return false;
                }

                state.Require(SUCCEEDED(decision->hr),
                              std::format(L"Remote compare: decision hr failed. hr=0x{:08X}", static_cast<unsigned long>(decision->hr)));
                state.Require(! decision->rightFolderMissing, L"Remote compare: remote root reported missing.");
                if (FAILED(decision->hr) || decision->rightFolderMissing)
                {
                    return false;
                }

                const auto relLeftRoot = session->TryMakeRelative(ComparePane::Left, localRoot);
                state.Require(relLeftRoot.has_value() && relLeftRoot.value().empty(), L"Remote compare: TryMakeRelative(Left, root) should return empty.");

                const auto relRightRoot = session->TryMakeRelative(ComparePane::Right, remoteRoot);
                state.Require(relRightRoot.has_value() && relRightRoot.value().empty(), L"Remote compare: TryMakeRelative(Right, root) should return empty.");

                const auto* item = FindItem(*decision, uniqueName);
                state.Require(item != nullptr, L"Remote compare: unique local file missing from decision.");
                if (item)
                {
                    state.Require(item->isDifferent, L"Remote compare: unique local file expected different.");
                    state.Require(item->selectLeft && ! item->selectRight, L"Remote compare: unique local file expected selectLeft only.");
                    state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::OnlyInLeft),
                                  L"Remote compare: unique local file expected differenceMask=OnlyInLeft.");
                }

                return state.failure.empty();
            });
        };

        const auto runRemoteDirectorySizeCallbackContract = [&](std::wstring_view caseName,
                                                                std::wstring_view protocolLabel,
                                                                std::wstring_view envVarName,
                                                                std::wstring_view defaultProfileName,
                                                                std::wstring_view pluginId) noexcept
        {
            if (options.failFast && suite.failed != 0)
            {
                AppendCaseResult(suite, caseName, SelfTest::SelfTestCaseResult::Status::skipped, L"not executed (fail-fast)");
                return;
            }

            const PhaseCheckResult secretOutcome = CheckRemoteConnectionSecret(protocolLabel, envVarName, defaultProfileName, pluginId);
            if (secretOutcome.status != SelfTest::SelfTestCaseResult::Status::passed)
            {
                AppendCaseResult(suite, caseName, secretOutcome.status, secretOutcome.reason);
                return;
            }

            const PhaseCheckResult sandboxOutcome = CheckRemoteConnectionSandbox(protocolLabel, envVarName, defaultProfileName, pluginId);
            if (sandboxOutcome.status != SelfTest::SelfTestCaseResult::Status::passed)
            {
                AppendCaseResult(suite, caseName, sandboxOutcome.status, sandboxOutcome.reason);
                return;
            }

            SelfTest::RunCase(options,
                              suite,
                              caseName,
                              [&](SelfTest::CaseState& state) noexcept
            {
                const std::wstring overrideName                    = GetEnvVarTrimmed(envVarName);
                const std::wstring profileName                     = ! overrideName.empty() ? overrideName : std::wstring(defaultProfileName);
                const Common::Settings::ConnectionProfile* profile = ConnectionProfileUtils::FindConnectionProfileByName(&g_settings, profileName);
                state.Require(profile != nullptr, std::format(L"Remote {} directory size: profile missing after preconditions passed.", protocolLabel));
                if (! profile)
                {
                    return false;
                }

                const std::wstring initialPath = NormalizePluginPathForSelfTest(profile->initialPath);
                state.Require(! initialPath.empty() && initialPath[0] == L'/',
                              std::format(L"Remote {} directory size: initialPath must be an absolute plugin path.", protocolLabel));
                if (initialPath.empty() || initialPath[0] != L'/')
                {
                    return false;
                }

                CreatedFileSystemInstance remoteCreated{};
                const HRESULT createHr = TryCreateFileSystemInstance(pluginId, {}, remoteCreated);
                state.Require(SUCCEEDED(createHr) && remoteCreated.fileSystem,
                              std::format(L"Remote {} directory size: failed to create filesystem instance. hr=0x{:08X}",
                                          protocolLabel,
                                          static_cast<unsigned long>(createHr)));
                if (FAILED(createHr) || ! remoteCreated.fileSystem)
                {
                    return false;
                }

                wil::com_ptr<IFileSystemIO> io;
                const HRESULT ioHr = remoteCreated.fileSystem->QueryInterface(IID_PPV_ARGS(io.put()));
                state.Require(SUCCEEDED(ioHr) && io, std::format(L"Remote {} directory size: filesystem instance missing IFileSystemIO.", protocolLabel));
                if (FAILED(ioHr) || ! io)
                {
                    return false;
                }

                wil::com_ptr<IFileSystemDirectoryOperations> dirOps;
                const HRESULT dirHr = remoteCreated.fileSystem->QueryInterface(IID_PPV_ARGS(dirOps.put()));
                state.Require(SUCCEEDED(dirHr) && dirOps,
                              std::format(L"Remote {} directory size: filesystem instance missing IFileSystemDirectoryOperations.", protocolLabel));
                if (FAILED(dirHr) || ! dirOps)
                {
                    return false;
                }

                const std::wstring caseRootPlugin =
                    JoinPluginPathForSelfTest(initialPath, std::format(L"compare-selftest-dirsize-{}-{}", protocolLabel, guid));
                state.Require(! caseRootPlugin.empty(), std::format(L"Remote {} directory size: failed to build sandbox path.", protocolLabel));
                if (caseRootPlugin.empty())
                {
                    return false;
                }

                const std::wstring caseRootText = MakeConnectionPathForSelfTest(profileName, caseRootPlugin);
                state.Require(! caseRootText.empty(), std::format(L"Remote {} directory size: failed to build connection path.", protocolLabel));
                if (caseRootText.empty())
                {
                    return false;
                }

                const std::filesystem::path caseRoot(caseRootText);
                const auto cleanup = wil::scope_exit([&]() noexcept
                {
                    static_cast<void>(remoteCreated.fileSystem->DeleteItem(caseRootText.c_str(),
                                                                          static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_CONTINUE_ON_ERROR),
                                                                          nullptr,
                                                                          nullptr,
                                                                          nullptr));
                });

                state.Require(EnsureDirectoryExistsFsOps(dirOps, caseRoot),
                              std::format(L"Remote {} directory size: failed to prepare sandbox prefix.", protocolLabel));
                if (! state.failure.empty())
                {
                    return false;
                }

                const std::filesystem::path objectPath = caseRoot / L"probe.txt";
                const std::wstring objectText          = ToPluginPathText(objectPath);
                state.Require(WriteFileTextFsIo(io, objectPath, "directory-size-remote"),
                              std::format(L"Remote {} directory size: failed to write probe object.", protocolLabel));
                if (! state.failure.empty())
                {
                    return false;
                }

                static_cast<void>(ValidateDirectorySizeCallbackContract(state, dirOps.get(), objectText));
                return state.failure.empty();
            });
        };

        const auto runRemoteS3Pagination = [&](std::wstring_view caseName) noexcept
        {
            if (options.failFast && suite.failed != 0)
            {
                AppendCaseResult(suite, caseName, SelfTest::SelfTestCaseResult::Status::skipped, L"not executed (fail-fast)");
                return;
            }

            const PhaseCheckResult secretOutcome = CheckRemoteConnectionSecret(L"S3", kSelfTestEnvConnS3, kSelfTestDefaultConnS3, kBuiltinS3FileSystemId);
            if (secretOutcome.status != SelfTest::SelfTestCaseResult::Status::passed)
            {
                AppendCaseResult(suite, caseName, secretOutcome.status, secretOutcome.reason);
                return;
            }

            const PhaseCheckResult sandboxOutcome = CheckRemoteConnectionSandbox(L"S3", kSelfTestEnvConnS3, kSelfTestDefaultConnS3, kBuiltinS3FileSystemId);
            if (sandboxOutcome.status != SelfTest::SelfTestCaseResult::Status::passed)
            {
                AppendCaseResult(suite, caseName, sandboxOutcome.status, sandboxOutcome.reason);
                return;
            }

            SelfTest::RunCase(options,
                              suite,
                              caseName,
                              [&](SelfTest::CaseState& state) noexcept
            {
                const std::wstring overrideName                    = GetEnvVarTrimmed(kSelfTestEnvConnS3);
                const std::wstring profileName                     = ! overrideName.empty() ? overrideName : std::wstring(kSelfTestDefaultConnS3);
    const Common::Settings::ConnectionProfile* profile = ConnectionProfileUtils::FindConnectionProfileByName(&g_settings, profileName);
                state.Require(profile != nullptr, L"Remote S3 pagination: profile missing after preconditions passed.");
                if (! profile)
                {
                    return false;
                }

                const std::wstring initialPath = NormalizePluginPathForSelfTest(profile->initialPath);
                state.Require(! initialPath.empty() && initialPath[0] == L'/', L"Remote S3 pagination: initialPath must be an absolute plugin path.");
                if (initialPath.empty() || initialPath[0] != L'/')
                {
                    return false;
                }

                CreatedFileSystemInstance remoteCreated{};
                const HRESULT createHr = TryCreateFileSystemInstance(kBuiltinS3FileSystemId, {}, remoteCreated);
                state.Require(SUCCEEDED(createHr),
                              std::format(L"Remote S3 pagination: failed to create filesystem instance. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
                if (FAILED(createHr) || ! remoteCreated.fileSystem)
                {
                    return false;
                }

                wil::com_ptr<IFileSystemIO> io;
                const HRESULT ioHr = remoteCreated.fileSystem->QueryInterface(IID_PPV_ARGS(io.put()));
                state.Require(SUCCEEDED(ioHr) && io, L"Remote S3 pagination: filesystem instance missing IFileSystemIO.");
                if (FAILED(ioHr) || ! io)
                {
                    return false;
                }

                wil::com_ptr<IFileSystemDirectoryOperations> dirOps;
                const HRESULT hrQI = remoteCreated.fileSystem->QueryInterface(__uuidof(IFileSystemDirectoryOperations), dirOps.put_void());
                state.Require(SUCCEEDED(hrQI) && dirOps, L"Remote S3 pagination: filesystem instance missing IFileSystemDirectoryOperations.");
                if (FAILED(hrQI) || ! dirOps)
                {
                    return false;
                }

                const std::wstring caseRootPlugin = JoinPluginPathForSelfTest(initialPath, std::format(L"compare-selftest-s3-pagination-{}", guid));
                state.Require(! caseRootPlugin.empty(), L"Remote S3 pagination: failed to build sandbox path.");
                if (caseRootPlugin.empty())
                {
                    return false;
                }

                const std::wstring caseRootText = MakeConnectionPathForSelfTest(profileName, caseRootPlugin);
                state.Require(! caseRootText.empty(), L"Remote S3 pagination: failed to build connection path.");
                if (caseRootText.empty())
                {
                    return false;
                }

                const std::filesystem::path caseRoot(caseRootText);
                const auto cleanup = wil::scope_exit([&]() noexcept
                {
                    static_cast<void>(remoteCreated.fileSystem->DeleteItem(caseRootText.c_str(),
                                                                          static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_CONTINUE_ON_ERROR),
                                                                          nullptr,
                                                                          nullptr,
                                                                          nullptr));
                });

                state.Require(EnsureDirectoryExistsFsOps(dirOps, caseRoot), L"Remote S3 pagination: failed to prepare sandbox prefix.");
                if (! state.failure.empty())
                {
                    return false;
                }

                constexpr unsigned long kObjectCount = 1005u;
                constexpr std::string_view kPayload  = "p";
                for (unsigned long index = 0; index < kObjectCount; ++index)
                {
                    const std::filesystem::path objectPath = caseRoot / std::format(L"object_{:04}.txt", index);
                    state.Require(WriteFileTextFsIo(io, objectPath, kPayload),
                                  std::format(L"Remote S3 pagination: failed to write object {}", objectPath.wstring()));
                    if (! state.failure.empty())
                    {
                        return false;
                    }
                }

                FileSystemDirectorySizeResult sizeResult{};
                sizeResult.sizeBytes = sizeof(FileSystemDirectorySizeResult);
                const HRESULT sizeHr = dirOps->GetDirectorySize(caseRootText.c_str(), FileSystemFlags{}, nullptr, nullptr, &sizeResult);
                state.Require(SUCCEEDED(sizeHr) && SUCCEEDED(sizeResult.status),
                              std::format(L"Remote S3 pagination: GetDirectorySize failed. hr=0x{:08X} status=0x{:08X}",
                                          static_cast<unsigned long>(sizeHr),
                                          static_cast<unsigned long>(sizeResult.status)));
                if (FAILED(sizeHr) || FAILED(sizeResult.status))
                {
                    return false;
                }

                state.Require(sizeResult.fileCount == kObjectCount,
                              std::format(L"Remote S3 pagination: fileCount mismatch. expected={} got={}", kObjectCount, sizeResult.fileCount));
                state.Require(sizeResult.directoryCount == 0u,
                              std::format(L"Remote S3 pagination: expected directoryCount=0, got {}.", sizeResult.directoryCount));
                state.Require(sizeResult.totalBytes == static_cast<uint64_t>(kObjectCount * kPayload.size()),
                              std::format(L"Remote S3 pagination: totalBytes mismatch. expected={} got={}.",
                                          static_cast<uint64_t>(kObjectCount * kPayload.size()),
                                          sizeResult.totalBytes));
                if (! state.failure.empty())
                {
                    return false;
                }

                wil::com_ptr<IFilesInformation> listing;
                const HRESULT listHr = remoteCreated.fileSystem->ReadDirectoryInfo(caseRootText.c_str(), listing.put());
                state.Require(SUCCEEDED(listHr) && listing,
                              std::format(L"Remote S3 pagination: ReadDirectoryInfo failed. hr=0x{:08X}", static_cast<unsigned long>(listHr)));
                if (FAILED(listHr) || ! listing)
                {
                    return false;
                }

                unsigned long got     = 0;
                const HRESULT countHr = listing->GetCount(&got);
                state.Require(SUCCEEDED(countHr), std::format(L"Remote S3 pagination: GetCount failed. hr=0x{:08X}", static_cast<unsigned long>(countHr)));
                if (FAILED(countHr))
                {
                    return false;
                }

                state.Require(got == kObjectCount,
                              std::format(L"Remote S3 pagination: listing count mismatch. expected={} got={}", kObjectCount, got));
                return state.failure.empty();
            });
        };

        const auto runRemoteFtpPartialContinue = [&](std::wstring_view caseName) noexcept
        {
            if (options.failFast && suite.failed != 0)
            {
                AppendCaseResult(suite, caseName, SelfTest::SelfTestCaseResult::Status::skipped, L"not executed (fail-fast)");
                return;
            }

            const PhaseCheckResult secretOutcome = CheckRemoteConnectionSecret(L"FTP", kSelfTestEnvConnFtp, kSelfTestDefaultConnFtp, kBuiltinFtpFileSystemId);
            if (secretOutcome.status != SelfTest::SelfTestCaseResult::Status::passed)
            {
                AppendCaseResult(suite, caseName, secretOutcome.status, secretOutcome.reason);
                return;
            }

            const PhaseCheckResult sandboxOutcome = CheckRemoteConnectionSandbox(L"FTP", kSelfTestEnvConnFtp, kSelfTestDefaultConnFtp, kBuiltinFtpFileSystemId);
            if (sandboxOutcome.status != SelfTest::SelfTestCaseResult::Status::passed)
            {
                AppendCaseResult(suite, caseName, sandboxOutcome.status, sandboxOutcome.reason);
                return;
            }

            SelfTest::RunCase(options,
                              suite,
                              caseName,
                              [&](SelfTest::CaseState& state) noexcept
            {
                const std::wstring overrideName                    = GetEnvVarTrimmed(kSelfTestEnvConnFtp);
                const std::wstring profileName                     = ! overrideName.empty() ? overrideName : std::wstring(kSelfTestDefaultConnFtp);
                const Common::Settings::ConnectionProfile* profile = ConnectionProfileUtils::FindConnectionProfileByName(&g_settings, profileName);
                state.Require(profile != nullptr, L"Remote FTP partial: profile missing after preconditions passed.");
                if (! profile)
                {
                    return false;
                }

                const std::wstring initialPath = NormalizePluginPathForSelfTest(profile->initialPath);
                state.Require(! initialPath.empty() && initialPath[0] == L'/', L"Remote FTP partial: initialPath must be an absolute plugin path.");
                if (initialPath.empty() || initialPath[0] != L'/')
                {
                    return false;
                }

                CreatedFileSystemInstance remoteCreated{};
                const HRESULT createHr = TryCreateFileSystemInstance(kBuiltinFtpFileSystemId, {}, remoteCreated);
                state.Require(SUCCEEDED(createHr),
                              std::format(L"Remote FTP partial: failed to create filesystem instance. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
                if (FAILED(createHr) || ! remoteCreated.fileSystem)
                {
                    return false;
                }

                wil::com_ptr<IFileSystemIO> io;
                const HRESULT ioHr = remoteCreated.fileSystem->QueryInterface(IID_PPV_ARGS(io.put()));
                state.Require(SUCCEEDED(ioHr) && io, L"Remote FTP partial: filesystem instance missing IFileSystemIO.");
                if (FAILED(ioHr) || ! io)
                {
                    return false;
                }

                wil::com_ptr<IFileSystemDirectoryOperations> dirOps;
                const HRESULT dirHr = remoteCreated.fileSystem->QueryInterface(IID_PPV_ARGS(dirOps.put()));
                state.Require(SUCCEEDED(dirHr) && dirOps, L"Remote FTP partial: filesystem instance missing IFileSystemDirectoryOperations.");
                if (FAILED(dirHr) || ! dirOps)
                {
                    return false;
                }

                const std::wstring caseRootPlugin = JoinPluginPathForSelfTest(initialPath, std::format(L"compare-selftest-ftp-partial-{}", guid));
                state.Require(! caseRootPlugin.empty(), L"Remote FTP partial: failed to build sandbox path.");
                if (caseRootPlugin.empty())
                {
                    return false;
                }

                const std::wstring caseRootText = MakeConnectionPathForSelfTest(profileName, caseRootPlugin);
                state.Require(! caseRootText.empty(), L"Remote FTP partial: failed to build connection path.");
                if (caseRootText.empty())
                {
                    return false;
                }

                const std::filesystem::path caseRoot(caseRootText);
                const auto cleanup = wil::scope_exit([&]() noexcept
                {
                    const FileSystemFlags cleanupFlags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_CONTINUE_ON_ERROR);
                    const std::wstring cleanupPath     = ToPluginPathText(caseRoot);
                    static_cast<void>(remoteCreated.fileSystem->DeleteItem(cleanupPath.c_str(), cleanupFlags, nullptr, nullptr, nullptr));
                });

                const std::filesystem::path srcDir     = caseRoot / L"src";
                const std::filesystem::path copyDst    = caseRoot / L"copy-dst";
                const std::filesystem::path moveDst    = caseRoot / L"move-dst";
                const std::filesystem::path copyGood   = srcDir / L"copy-good.txt";
                const std::filesystem::path copyMiss   = srcDir / L"copy-missing.txt";
                const std::filesystem::path moveGood   = srcDir / L"move-good.txt";
                const std::filesystem::path moveMiss   = srcDir / L"move-missing.txt";
                const std::filesystem::path renameGood = srcDir / L"rename-good.txt";
                const std::filesystem::path renameMiss = srcDir / L"rename-missing.txt";

                state.Require(EnsureDirectoryExistsFsOps(dirOps, srcDir), L"Remote FTP partial: failed to create source directory.");
                state.Require(EnsureDirectoryExistsFsOps(dirOps, copyDst), L"Remote FTP partial: failed to create copy destination.");
                state.Require(EnsureDirectoryExistsFsOps(dirOps, moveDst), L"Remote FTP partial: failed to create move destination.");
                state.Require(WriteFileTextFsIo(io, copyGood, "copy-good"), L"Remote FTP partial: failed to write copy-good.txt.");
                state.Require(WriteFileTextFsIo(io, moveGood, "move-good"), L"Remote FTP partial: failed to write move-good.txt.");
                state.Require(WriteFileTextFsIo(io, renameGood, "rename-good"), L"Remote FTP partial: failed to write rename-good.txt.");
                if (! state.failure.empty())
                {
                    return false;
                }

                const std::wstring copyGoodText = ToPluginPathText(copyGood);
                const std::wstring copyMissText = ToPluginPathText(copyMiss);
                const std::wstring copyDstText  = ToPluginPathText(copyDst);
                const wchar_t* copySources[]    = {copyGoodText.c_str(), copyMissText.c_str()};
                RecordingFileSystemCallback copyCallback(2);
                const HRESULT copyHr = remoteCreated.fileSystem->CopyItems(copySources,
                                                                           2,
                                                                           copyDstText.c_str(),
                                                                           static_cast<FileSystemFlags>(FILESYSTEM_FLAG_CONTINUE_ON_ERROR),
                                                                           nullptr,
                                                                           &copyCallback,
                                                                           nullptr);
                const HRESULT partialHr = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
                state.Require(copyHr == partialHr,
                              std::format(L"Remote FTP partial: CopyItems expected 0x{:08X}, got 0x{:08X}.",
                                          static_cast<unsigned long>(partialHr),
                                          static_cast<unsigned long>(copyHr)));
                state.Require(! copyCallback.SawUnexpectedIssue(), L"Remote FTP partial: CopyItems triggered an unexpected issue callback.");
                state.Require(copyCallback.CompletedCount() == 2u,
                              std::format(L"Remote FTP partial: CopyItems expected 2 completion callbacks, got {}.", copyCallback.CompletedCount()));
                RecordedFileSystemItem copyItem0{};
                RecordedFileSystemItem copyItem1{};
                state.Require(copyCallback.TryGetItem(0, copyItem0), L"Remote FTP partial: CopyItems missing callback for item 0.");
                state.Require(copyCallback.TryGetItem(1, copyItem1), L"Remote FTP partial: CopyItems missing callback for item 1.");
                if (! state.failure.empty())
                {
                    return false;
                }

                state.Require(SUCCEEDED(copyItem0.status),
                              std::format(L"Remote FTP partial: CopyItems item 0 expected success, got 0x{:08X}.",
                                          static_cast<unsigned long>(copyItem0.status)));
                state.Require(copyItem1.status == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND),
                              std::format(L"Remote FTP partial: CopyItems item 1 expected ERROR_FILE_NOT_FOUND, got 0x{:08X}.",
                                          static_cast<unsigned long>(copyItem1.status)));
                state.Require(PathExistsFsIo(io, copyGood), L"Remote FTP partial: CopyItems removed the successful source unexpectedly.");
                state.Require(PathExistsFsIo(io, copyDst / L"copy-good.txt"), L"Remote FTP partial: CopyItems did not materialize the successful destination file.");
                if (! state.failure.empty())
                {
                    return false;
                }

                const std::wstring moveGoodText = ToPluginPathText(moveGood);
                const std::wstring moveMissText = ToPluginPathText(moveMiss);
                const std::wstring moveDstText  = ToPluginPathText(moveDst);
                const wchar_t* moveSources[]    = {moveGoodText.c_str(), moveMissText.c_str()};
                RecordingFileSystemCallback moveCallback(2);
                const HRESULT moveHr = remoteCreated.fileSystem->MoveItems(moveSources,
                                                                           2,
                                                                           moveDstText.c_str(),
                                                                           static_cast<FileSystemFlags>(FILESYSTEM_FLAG_CONTINUE_ON_ERROR),
                                                                           nullptr,
                                                                           &moveCallback,
                                                                           nullptr);
                state.Require(moveHr == partialHr,
                              std::format(L"Remote FTP partial: MoveItems expected 0x{:08X}, got 0x{:08X}.",
                                          static_cast<unsigned long>(partialHr),
                                          static_cast<unsigned long>(moveHr)));
                state.Require(! moveCallback.SawUnexpectedIssue(), L"Remote FTP partial: MoveItems triggered an unexpected issue callback.");
                state.Require(moveCallback.CompletedCount() == 2u,
                              std::format(L"Remote FTP partial: MoveItems expected 2 completion callbacks, got {}.", moveCallback.CompletedCount()));
                RecordedFileSystemItem moveItem0{};
                RecordedFileSystemItem moveItem1{};
                state.Require(moveCallback.TryGetItem(0, moveItem0), L"Remote FTP partial: MoveItems missing callback for item 0.");
                state.Require(moveCallback.TryGetItem(1, moveItem1), L"Remote FTP partial: MoveItems missing callback for item 1.");
                if (! state.failure.empty())
                {
                    return false;
                }

                state.Require(SUCCEEDED(moveItem0.status),
                              std::format(L"Remote FTP partial: MoveItems item 0 expected success, got 0x{:08X}.",
                                          static_cast<unsigned long>(moveItem0.status)));
                state.Require(moveItem1.status == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND),
                              std::format(L"Remote FTP partial: MoveItems item 1 expected ERROR_FILE_NOT_FOUND, got 0x{:08X}.",
                                          static_cast<unsigned long>(moveItem1.status)));
                state.Require(! PathExistsFsIo(io, moveGood), L"Remote FTP partial: MoveItems left the successful source in place.");
                state.Require(PathExistsFsIo(io, moveDst / L"move-good.txt"), L"Remote FTP partial: MoveItems did not create the successful destination file.");
                if (! state.failure.empty())
                {
                    return false;
                }

                const std::wstring renameGoodText = ToPluginPathText(renameGood);
                const std::wstring renameMissText = ToPluginPathText(renameMiss);
                FileSystemRenamePair renamePairs[2]{};
                renamePairs[0].sizeBytes  = sizeof(FileSystemRenamePair);
                renamePairs[0].sourcePath = renameGoodText.c_str();
                renamePairs[0].newName    = L"rename-good-dst.txt";
                renamePairs[1].sizeBytes  = sizeof(FileSystemRenamePair);
                renamePairs[1].sourcePath = renameMissText.c_str();
                renamePairs[1].newName    = L"rename-missing-dst.txt";

                RecordingFileSystemCallback renameCallback(2);
                const HRESULT renameHr = remoteCreated.fileSystem->RenameItems(renamePairs,
                                                                               2,
                                                                               static_cast<FileSystemFlags>(FILESYSTEM_FLAG_CONTINUE_ON_ERROR),
                                                                               nullptr,
                                                                               &renameCallback,
                                                                               nullptr);
                state.Require(renameHr == partialHr,
                              std::format(L"Remote FTP partial: RenameItems expected 0x{:08X}, got 0x{:08X}.",
                                          static_cast<unsigned long>(partialHr),
                                          static_cast<unsigned long>(renameHr)));
                state.Require(! renameCallback.SawUnexpectedIssue(), L"Remote FTP partial: RenameItems triggered an unexpected issue callback.");
                state.Require(renameCallback.CompletedCount() == 2u,
                              std::format(L"Remote FTP partial: RenameItems expected 2 completion callbacks, got {}.", renameCallback.CompletedCount()));
                RecordedFileSystemItem renameItem0{};
                RecordedFileSystemItem renameItem1{};
                state.Require(renameCallback.TryGetItem(0, renameItem0), L"Remote FTP partial: RenameItems missing callback for item 0.");
                state.Require(renameCallback.TryGetItem(1, renameItem1), L"Remote FTP partial: RenameItems missing callback for item 1.");
                if (! state.failure.empty())
                {
                    return false;
                }

                state.Require(SUCCEEDED(renameItem0.status),
                              std::format(L"Remote FTP partial: RenameItems item 0 expected success, got 0x{:08X}.",
                                          static_cast<unsigned long>(renameItem0.status)));
                state.Require(renameItem1.status == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND),
                              std::format(L"Remote FTP partial: RenameItems item 1 expected ERROR_FILE_NOT_FOUND, got 0x{:08X}.",
                                          static_cast<unsigned long>(renameItem1.status)));
                state.Require(! PathExistsFsIo(io, renameGood), L"Remote FTP partial: RenameItems left the successful source in place.");
                state.Require(PathExistsFsIo(io, srcDir / L"rename-good-dst.txt"),
                              L"Remote FTP partial: RenameItems did not create the successful destination name.");
                return state.failure.empty();
            });
        };

        const auto runRemoteS3MetadataSmoke = [&](std::wstring_view caseName) noexcept
        {
            if (options.failFast && suite.failed != 0)
            {
                AppendCaseResult(suite, caseName, SelfTest::SelfTestCaseResult::Status::skipped, L"not executed (fail-fast)");
                return;
            }

            const PhaseCheckResult secretOutcome = CheckRemoteConnectionSecret(L"S3", kSelfTestEnvConnS3, kSelfTestDefaultConnS3, kBuiltinS3FileSystemId);
            if (secretOutcome.status != SelfTest::SelfTestCaseResult::Status::passed)
            {
                AppendCaseResult(suite, caseName, secretOutcome.status, secretOutcome.reason);
                return;
            }

            const PhaseCheckResult sandboxOutcome = CheckRemoteConnectionSandbox(L"S3", kSelfTestEnvConnS3, kSelfTestDefaultConnS3, kBuiltinS3FileSystemId);
            if (sandboxOutcome.status != SelfTest::SelfTestCaseResult::Status::passed)
            {
                AppendCaseResult(suite, caseName, sandboxOutcome.status, sandboxOutcome.reason);
                return;
            }

            SelfTest::RunCase(options,
                              suite,
                              caseName,
                              [&](SelfTest::CaseState& state) noexcept
            {
                const std::wstring overrideName                    = GetEnvVarTrimmed(kSelfTestEnvConnS3);
                const std::wstring profileName                     = ! overrideName.empty() ? overrideName : std::wstring(kSelfTestDefaultConnS3);
                const Common::Settings::ConnectionProfile* profile = ConnectionProfileUtils::FindConnectionProfileByName(&g_settings, profileName);
                state.Require(profile != nullptr, L"Remote S3 metadata: profile missing after preconditions passed.");
                if (! profile)
                {
                    return false;
                }

                const std::wstring initialPath = NormalizePluginPathForSelfTest(profile->initialPath);
                state.Require(! initialPath.empty() && initialPath[0] == L'/', L"Remote S3 metadata: initialPath must be an absolute plugin path.");
                if (initialPath.empty() || initialPath[0] != L'/')
                {
                    return false;
                }

                CreatedFileSystemInstance remoteCreated{};
                const HRESULT createHr = TryCreateFileSystemInstance(kBuiltinS3FileSystemId, {}, remoteCreated);
                state.Require(SUCCEEDED(createHr),
                              std::format(L"Remote S3 metadata: failed to create filesystem instance. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
                if (FAILED(createHr) || ! remoteCreated.fileSystem)
                {
                    return false;
                }

                wil::com_ptr<IFileSystemIO> io;
                const HRESULT ioHr = remoteCreated.fileSystem->QueryInterface(IID_PPV_ARGS(io.put()));
                state.Require(SUCCEEDED(ioHr) && io, L"Remote S3 metadata: filesystem instance missing IFileSystemIO.");
                if (FAILED(ioHr) || ! io)
                {
                    return false;
                }

                wil::com_ptr<IFileSystemDirectoryOperations> dirOps;
                const HRESULT dirHr = remoteCreated.fileSystem->QueryInterface(IID_PPV_ARGS(dirOps.put()));
                state.Require(SUCCEEDED(dirHr) && dirOps, L"Remote S3 metadata: filesystem instance missing IFileSystemDirectoryOperations.");
                if (FAILED(dirHr) || ! dirOps)
                {
                    return false;
                }

                const std::wstring caseRootPlugin = JoinPluginPathForSelfTest(initialPath, std::format(L"compare-selftest-s3-metadata-{}", guid));
                state.Require(! caseRootPlugin.empty(), L"Remote S3 metadata: failed to build sandbox path.");
                if (caseRootPlugin.empty())
                {
                    return false;
                }

                const std::wstring caseRootText = MakeConnectionPathForSelfTest(profileName, caseRootPlugin);
                state.Require(! caseRootText.empty(), L"Remote S3 metadata: failed to build connection path.");
                if (caseRootText.empty())
                {
                    return false;
                }

                const std::filesystem::path caseRoot(caseRootText);
                state.Require(EnsureDirectoryExistsFsOps(dirOps, caseRoot), L"Remote S3 metadata: failed to prepare sandbox prefix.");
                if (! state.failure.empty())
                {
                    return false;
                }

                const std::filesystem::path objectPath = caseRoot / L"metadata.txt";
                const std::wstring objectText          = ToPluginPathText(objectPath);
                const auto cleanup = wil::scope_exit([&]() noexcept
                {
                    static_cast<void>(remoteCreated.fileSystem->DeleteItem(objectText.c_str(),
                                                                          static_cast<FileSystemFlags>(FILESYSTEM_FLAG_CONTINUE_ON_ERROR),
                                                                          nullptr,
                                                                          nullptr,
                                                                          nullptr));
                });

                constexpr std::string_view kPayload = "s3-metadata-smoke";
                state.Require(WriteFileTextFsIo(io, objectPath, kPayload), L"Remote S3 metadata: failed to write test object.");
                if (! state.failure.empty())
                {
                    return false;
                }

                FileSystemBasicInformation info{};
                info.sizeBytes = sizeof(FileSystemBasicInformation);
                const HRESULT infoHr = io->GetFileBasicInformation(objectText.c_str(), &info);
                state.Require(SUCCEEDED(infoHr),
                              std::format(L"Remote S3 metadata: GetFileBasicInformation failed. hr=0x{:08X}", static_cast<unsigned long>(infoHr)));
                if (FAILED(infoHr))
                {
                    return false;
                }

                state.Require((info.attributes & FILE_ATTRIBUTE_DIRECTORY) == 0, L"Remote S3 metadata: object reported as a directory.");
                state.Require(info.lastWriteTime != 0, L"Remote S3 metadata: object lastWriteTime was not populated.");
                if (! state.failure.empty())
                {
                    return false;
                }

                wil::com_ptr<IFileReader> reader;
                const HRESULT readerHr = io->CreateFileReader(objectText.c_str(), reader.put());
                state.Require(SUCCEEDED(readerHr) && reader,
                              std::format(L"Remote S3 metadata: CreateFileReader failed. hr=0x{:08X}", static_cast<unsigned long>(readerHr)));
                if (FAILED(readerHr) || ! reader)
                {
                    return false;
                }

                uint64_t sizeBytes = 0;
                const HRESULT sizeHr = reader->GetSize(&sizeBytes);
                state.Require(SUCCEEDED(sizeHr),
                              std::format(L"Remote S3 metadata: reader->GetSize failed. hr=0x{:08X}", static_cast<unsigned long>(sizeHr)));
                state.Require(sizeBytes == static_cast<uint64_t>(kPayload.size()),
                              std::format(L"Remote S3 metadata: expected size {} but got {}.", kPayload.size(), sizeBytes));
                if (! state.failure.empty())
                {
                    return false;
                }

                std::string readBack;
                readBack.resize(kPayload.size());
                size_t totalRead = 0;
                while (totalRead < readBack.size())
                {
                    unsigned long chunkRead = 0;
                    const unsigned long requestBytes = static_cast<unsigned long>((std::min)(readBack.size() - totalRead,
                                                                                               static_cast<size_t>((std::numeric_limits<unsigned long>::max)())));
                    const HRESULT readHr = reader->Read(readBack.data() + totalRead, requestBytes, &chunkRead);
                    state.Require(SUCCEEDED(readHr),
                                  std::format(L"Remote S3 metadata: reader->Read failed after {} bytes. hr=0x{:08X}",
                                              totalRead,
                                              static_cast<unsigned long>(readHr)));
                    if (FAILED(readHr))
                    {
                        return false;
                    }
                    if (chunkRead == 0)
                    {
                        break;
                    }
                    totalRead += chunkRead;
                }

                state.Require(totalRead == kPayload.size(),
                              std::format(L"Remote S3 metadata: expected to read {} bytes but got {}.", kPayload.size(), totalRead));
                state.Require(std::string_view(readBack.data(), totalRead) == kPayload,
                              L"Remote S3 metadata: reader contents did not match the uploaded payload.");
                return state.failure.empty();
            });
        };

        const auto runRemoteS3DeleteMissing = [&](std::wstring_view caseName) noexcept
        {
            if (options.failFast && suite.failed != 0)
            {
                AppendCaseResult(suite, caseName, SelfTest::SelfTestCaseResult::Status::skipped, L"not executed (fail-fast)");
                return;
            }

            const PhaseCheckResult secretOutcome = CheckRemoteConnectionSecret(L"S3", kSelfTestEnvConnS3, kSelfTestDefaultConnS3, kBuiltinS3FileSystemId);
            if (secretOutcome.status != SelfTest::SelfTestCaseResult::Status::passed)
            {
                AppendCaseResult(suite, caseName, secretOutcome.status, secretOutcome.reason);
                return;
            }

            const PhaseCheckResult sandboxOutcome = CheckRemoteConnectionSandbox(L"S3", kSelfTestEnvConnS3, kSelfTestDefaultConnS3, kBuiltinS3FileSystemId);
            if (sandboxOutcome.status != SelfTest::SelfTestCaseResult::Status::passed)
            {
                AppendCaseResult(suite, caseName, sandboxOutcome.status, sandboxOutcome.reason);
                return;
            }

            SelfTest::RunCase(options,
                              suite,
                              caseName,
                              [&](SelfTest::CaseState& state) noexcept
            {
                const std::wstring overrideName                    = GetEnvVarTrimmed(kSelfTestEnvConnS3);
                const std::wstring profileName                     = ! overrideName.empty() ? overrideName : std::wstring(kSelfTestDefaultConnS3);
                const Common::Settings::ConnectionProfile* profile = ConnectionProfileUtils::FindConnectionProfileByName(&g_settings, profileName);
                state.Require(profile != nullptr, L"Remote S3 delete: profile missing after preconditions passed.");
                if (! profile)
                {
                    return false;
                }

                const std::wstring initialPath = NormalizePluginPathForSelfTest(profile->initialPath);
                state.Require(! initialPath.empty() && initialPath[0] == L'/', L"Remote S3 delete: initialPath must be an absolute plugin path.");
                if (initialPath.empty() || initialPath[0] != L'/')
                {
                    return false;
                }

                CreatedFileSystemInstance remoteCreated{};
                const HRESULT createHr = TryCreateFileSystemInstance(kBuiltinS3FileSystemId, {}, remoteCreated);
                state.Require(SUCCEEDED(createHr),
                              std::format(L"Remote S3 delete: failed to create filesystem instance. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
                if (FAILED(createHr) || ! remoteCreated.fileSystem)
                {
                    return false;
                }

                wil::com_ptr<IFileSystemIO> io;
                const HRESULT ioHr = remoteCreated.fileSystem->QueryInterface(IID_PPV_ARGS(io.put()));
                state.Require(SUCCEEDED(ioHr) && io, L"Remote S3 delete: filesystem instance missing IFileSystemIO.");
                if (FAILED(ioHr) || ! io)
                {
                    return false;
                }

                wil::com_ptr<IFileSystemDirectoryOperations> dirOps;
                const HRESULT dirHr = remoteCreated.fileSystem->QueryInterface(IID_PPV_ARGS(dirOps.put()));
                state.Require(SUCCEEDED(dirHr) && dirOps, L"Remote S3 delete: filesystem instance missing IFileSystemDirectoryOperations.");
                if (FAILED(dirHr) || ! dirOps)
                {
                    return false;
                }

                const std::wstring caseRootPlugin = JoinPluginPathForSelfTest(initialPath, std::format(L"compare-selftest-s3-delete-{}", guid));
                state.Require(! caseRootPlugin.empty(), L"Remote S3 delete: failed to build sandbox path.");
                if (caseRootPlugin.empty())
                {
                    return false;
                }

                const std::wstring caseRootText = MakeConnectionPathForSelfTest(profileName, caseRootPlugin);
                state.Require(! caseRootText.empty(), L"Remote S3 delete: failed to build connection path.");
                if (caseRootText.empty())
                {
                    return false;
                }

                const std::filesystem::path caseRoot(caseRootText);
                state.Require(EnsureDirectoryExistsFsOps(dirOps, caseRoot), L"Remote S3 delete: failed to prepare sandbox prefix.");
                if (! state.failure.empty())
                {
                    return false;
                }

                const std::filesystem::path singlePath    = caseRoot / L"single-delete.txt";
                const std::filesystem::path batchGoodPath = caseRoot / L"batch-existing.txt";
                const std::filesystem::path batchMissPath = caseRoot / L"batch-missing.txt";
                const std::wstring singleText             = ToPluginPathText(singlePath);
                const std::wstring batchGoodText          = ToPluginPathText(batchGoodPath);
                const std::wstring batchMissText          = ToPluginPathText(batchMissPath);
                const auto cleanup = wil::scope_exit([&]() noexcept
                {
                    const wchar_t* cleanupPaths[] = {singleText.c_str(), batchGoodText.c_str()};
                    static_cast<void>(remoteCreated.fileSystem->DeleteItems(cleanupPaths,
                                                                            2,
                                                                            static_cast<FileSystemFlags>(FILESYSTEM_FLAG_CONTINUE_ON_ERROR),
                                                                            nullptr,
                                                                            nullptr,
                                                                            nullptr));
                });

                state.Require(WriteFileTextFsIo(io, singlePath, "single-delete"), L"Remote S3 delete: failed to write single-delete.txt.");
                state.Require(WriteFileTextFsIo(io, batchGoodPath, "batch-delete"), L"Remote S3 delete: failed to write batch-existing.txt.");
                if (! state.failure.empty())
                {
                    return false;
                }

                const HRESULT deleteHr = remoteCreated.fileSystem->DeleteItem(singleText.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr);
                state.Require(SUCCEEDED(deleteHr),
                              std::format(L"Remote S3 delete: first DeleteItem failed. hr=0x{:08X}", static_cast<unsigned long>(deleteHr)));
                state.Require(! PathExistsFsIo(io, singlePath), L"Remote S3 delete: first DeleteItem did not remove the object.");
                if (! state.failure.empty())
                {
                    return false;
                }

                const HRESULT deleteMissingHr = remoteCreated.fileSystem->DeleteItem(singleText.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr);
                const HRESULT notFoundHr      = HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
                state.Require(deleteMissingHr == notFoundHr,
                              std::format(L"Remote S3 delete: second DeleteItem expected 0x{:08X}, got 0x{:08X}.",
                                          static_cast<unsigned long>(notFoundHr),
                                          static_cast<unsigned long>(deleteMissingHr)));
                if (! state.failure.empty())
                {
                    return false;
                }

                RecordingFileSystemCallback deleteCallback(2);
                const wchar_t* deletePaths[] = {batchGoodText.c_str(), batchMissText.c_str()};
                const HRESULT batchDeleteHr  = remoteCreated.fileSystem->DeleteItems(deletePaths,
                                                                                    2,
                                                                                    static_cast<FileSystemFlags>(FILESYSTEM_FLAG_CONTINUE_ON_ERROR),
                                                                                    nullptr,
                                                                                    &deleteCallback,
                                                                                    nullptr);
                const HRESULT partialHr = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
                state.Require(batchDeleteHr == partialHr,
                              std::format(L"Remote S3 delete: DeleteItems expected 0x{:08X}, got 0x{:08X}.",
                                          static_cast<unsigned long>(partialHr),
                                          static_cast<unsigned long>(batchDeleteHr)));
                state.Require(! deleteCallback.SawUnexpectedIssue(), L"Remote S3 delete: DeleteItems triggered an unexpected issue callback.");
                state.Require(deleteCallback.CompletedCount() == 2u,
                              std::format(L"Remote S3 delete: DeleteItems expected 2 completion callbacks, got {}.", deleteCallback.CompletedCount()));
                RecordedFileSystemItem deleteItem0{};
                RecordedFileSystemItem deleteItem1{};
                state.Require(deleteCallback.TryGetItem(0, deleteItem0), L"Remote S3 delete: DeleteItems missing callback for item 0.");
                state.Require(deleteCallback.TryGetItem(1, deleteItem1), L"Remote S3 delete: DeleteItems missing callback for item 1.");
                if (! state.failure.empty())
                {
                    return false;
                }

                state.Require(SUCCEEDED(deleteItem0.status),
                              std::format(L"Remote S3 delete: DeleteItems item 0 expected success, got 0x{:08X}.",
                                          static_cast<unsigned long>(deleteItem0.status)));
                state.Require(deleteItem1.status == notFoundHr,
                              std::format(L"Remote S3 delete: DeleteItems item 1 expected 0x{:08X}, got 0x{:08X}.",
                                          static_cast<unsigned long>(notFoundHr),
                                          static_cast<unsigned long>(deleteItem1.status)));
                state.Require(! PathExistsFsIo(io, batchGoodPath), L"Remote S3 delete: DeleteItems did not remove the existing object.");
                return state.failure.empty();
            });
        };

        // Optional remote smoke: runs only when Connection Manager profiles + secrets exist.
        runRemoteFileCompare(L"remote_file_s3", L"S3", kSelfTestEnvConnS3, kSelfTestDefaultConnS3, kBuiltinS3FileSystemId);
        runRemoteDirectorySizeCallbackContract(
            L"remote_s3_directory_size_callback_contract", L"S3", kSelfTestEnvConnS3, kSelfTestDefaultConnS3, kBuiltinS3FileSystemId);
        runRemoteS3Pagination(L"remote_s3_pagination");
        runRemoteFileCompare(L"remote_file_ftp", L"FTP", kSelfTestEnvConnFtp, kSelfTestDefaultConnFtp, kBuiltinFtpFileSystemId);
        runRemoteDirectorySizeCallbackContract(
            L"remote_ftp_directory_size_callback_contract", L"FTP", kSelfTestEnvConnFtp, kSelfTestDefaultConnFtp, kBuiltinFtpFileSystemId);
        runRemoteFtpPartialContinue(L"remote_ftp_continue_on_error_partial");
        runRemoteS3MetadataSmoke(L"remote_s3_metadata_smoke");
        runRemoteS3DeleteMissing(L"remote_s3_delete_missing");
    }

    AppendCompareSelfTestTraceLine(L"Run: finalizing");

    const auto endedAt = std::chrono::steady_clock::now();
    suite.durationMs   = static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::milliseconds>(endedAt - startedAt).count());

    if (outResult)
    {
        *outResult = suite;
    }

    if (options.writeJsonSummary)
    {
        const std::filesystem::path jsonPath = SelfTest::GetSuiteArtifactPath(SelfTest::SelfTestSuite::CompareDirectories, L"results.json");
        SelfTest::WriteSuiteJson(suite, jsonPath);
    }

    if (suite.failed != 0)
    {
        AppendCompareSelfTestTraceLine(L"Run: failed");
        Debug::Error(L"CompareSelfTest: failed.");
        return false;
    }

    AppendCompareSelfTestTraceLine(L"Run: passed");
    Debug::Info(L"CompareSelfTest: passed.");
    return true;
}

#endif // _DEBUG
