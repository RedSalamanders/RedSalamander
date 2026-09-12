#include <algorithm>
#include <array>
#include <cstring>
#include <filesystem>
#include <fstream>
#include <format>
#include <iostream>
#include <latch>
#include <limits>
#include <memory>
#include <optional>
#include <span>
#include <string>
#include <string_view>
#include <thread>
#include <unordered_set>
#include <vector>

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <unknwn.h>

#define REDSAL_DEFINE_TRACE_PROVIDER
#include "DeleteOnCloseTemporaryFile.h"
#include "FileSystemPathIdentity.h"
#include "FileSystemRouteContract.h"
#include "Helpers.h"
#include "LocalizationManager.h"
#include "PackedFileInfoBuffer.h"
#include "RegistryUtils.h"
#include "PluginConfiguration.h"
#include "WslDistributionCatalog.h"
#include "YyjsonHelpers.h"
#include "TestSupport.h"
#include "PlugInterfaces/Factory.h"
#include "PlugInterfaces/FileSystem.h"
#include "PlugInterfaces/Host.h"
#include "PlugInterfaces/Informations.h"
#include "PlugInterfaces/Terminal.h"
#include "PlugInterfaces/Viewer.h"
#include "PlugInterfaces/FactoryImpl.h"
#include "TerminalVtUpgradeTestContract.h"

#include <wil/com.h>
#include <wil/resource.h>
#include <yyjson.h>

namespace
{

// ---------------------------------------------------------------------------
// yyjson RAII
// ---------------------------------------------------------------------------
using unique_yyjson_doc = Common::Json::UniqueDocument;

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

enum class ScriptedRouteFault : uint8_t
{
    None,
    OversizedPrefix,
    InvalidRecordSize,
    InvalidAvailability,
    InvalidCancellation,
    InvalidNamespace,
    InvalidComparison,
    InvalidNormalization,
    InvalidCaseRename,
    InvalidProofFlags,
    NonStrictBoolean,
    WatchdogContradiction,
    EmptyProviderId,
    OutsideArenaPointer,
    MisalignedArenaPointer,
    RequiredArenaOverflow,
    RequiresArenaFallback,
    UnsupportedIdentity,
    NonStrictPeer,
    PeerDenied,
    PeerFailure,
    InvalidNameResult,
    ProviderInvalidName,
    UnsupportedName,
    CollisionOutsideArena,
    EmptyCollisionKey,
    ChildStringArenaFallback,
    JoinRequiredMismatch,
    JoinLeafMismatch,
};

class ScriptedRouteCapabilities final : public IFileSystemRouteCapabilities
{
public:
    explicit ScriptedRouteCapabilities(ScriptedRouteFault fault) noexcept : _fault(fault) {}

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** object) noexcept override
    {
        if (object == nullptr)
        {
            return E_POINTER;
        }
        *object = nullptr;
        if (riid == __uuidof(IUnknown) || riid == __uuidof(IFileSystemRouteCapabilities))
        {
            *object = static_cast<IFileSystemRouteCapabilities*>(this);
            return S_OK;
        }
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() noexcept override { return 1u; }
    ULONG STDMETHODCALLTYPE Release() noexcept override { return 1u; }

    HRESULT STDMETHODCALLTYPE GetRouteFacts(const wchar_t* path,
                                            FileSystemOperation operation,
                                            FileSystemArena* arena,
                                            FileSystemRouteFacts* facts) noexcept override
    {
        if (path == nullptr || path[0] == L'\0' || operation > FILESYSTEM_CREATE_DIRECTORY || arena == nullptr || facts == nullptr)
        {
            return E_INVALIDARG;
        }
        if (_fault == ScriptedRouteFault::RequiredArenaOverflow)
        {
            facts->requiredArenaBytes = FileSystemRouteContract::kMaximumArenaBytes + 1u;
            return HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER);
        }

        const std::wstring_view provider = _fault == ScriptedRouteFault::EmptyProviderId ? std::wstring_view{} : L"selftest/typed-route";
        static const std::wstring largeProfile(
            FileSystemRouteContract::kNormalArenaBytes / sizeof(wchar_t) + 32u, L'p');
        const std::wstring_view profile = _fault == ScriptedRouteFault::RequiresArenaFallback
            ? std::wstring_view(largeProfile)
            : std::wstring_view(L"selftest-profile");
        constexpr std::wstring_view root = L"selftest-root";
        constexpr std::wstring_view separators = L"/";
        const uint64_t required64 = Bytes(provider) + Bytes(profile) + Bytes(root) + Bytes(separators);
        if (required64 > (std::numeric_limits<unsigned long>::max)())
        {
            return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
        }
        facts->requiredArenaBytes = static_cast<unsigned long>(required64);
        if (arena->buffer == nullptr || arena->usedBytes > arena->capacityBytes ||
            facts->requiredArenaBytes > arena->capacityBytes - arena->usedBytes)
        {
            return HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER);
        }

        const wchar_t* accepted = Copy(separators, arena);
        const wchar_t* providerId = Copy(provider, arena);
        const wchar_t* pathProfileId = Copy(profile, arena);
        const wchar_t* rootId = Copy(root, arena);
        if (accepted == nullptr || providerId == nullptr || pathProfileId == nullptr || rootId == nullptr)
        {
            return E_OUTOFMEMORY;
        }

        facts->sizeBytes = _fault == ScriptedRouteFault::OversizedPrefix
            ? sizeof(FileSystemRouteFacts) + 16u
            : (_fault == ScriptedRouteFault::InvalidRecordSize ? sizeof(FileSystemRouteFacts) - 1u : sizeof(FileSystemRouteFacts));
        facts->availability = _fault == ScriptedRouteFault::InvalidAvailability
            ? static_cast<FileSystemRouteAvailability>(99u)
            : FILESYSTEM_ROUTE_AVAILABLE;
        facts->cancellationRoute = _fault == ScriptedRouteFault::InvalidCancellation
            ? static_cast<FileSystemCancellationRoute>(99u)
            : FILESYSTEM_CANCELLATION_BOUNDED;
        facts->namespaceKind = _fault == ScriptedRouteFault::InvalidNamespace
            ? static_cast<FileSystemNamespaceKind>(99u)
            : FILESYSTEM_NAMESPACE_REAL_CONTAINER;
        facts->componentComparison = _fault == ScriptedRouteFault::InvalidComparison
            ? static_cast<FileSystemRouteComponentComparison>(99u)
            : FILESYSTEM_ROUTE_COMPONENT_ORDINAL_CASE_SENSITIVE;
        facts->normalization = _fault == ScriptedRouteFault::InvalidNormalization
            ? static_cast<FileSystemRouteNormalization>(99u)
            : FILESYSTEM_ROUTE_NORMALIZATION_NONE;
        facts->caseOnlyRename = _fault == ScriptedRouteFault::InvalidCaseRename
            ? static_cast<FileSystemRouteCaseOnlyRename>(99u)
            : FILESYSTEM_ROUTE_CASE_ONLY_SUPPORTED;
        facts->proofFlags = _fault == ScriptedRouteFault::InvalidProofFlags ? 0x80000000u : FILESYSTEM_ROUTE_PROOF_NONE;
        facts->providerWatchdogTimeoutMs = _fault == ScriptedRouteFault::WatchdogContradiction ? 30'000u : 0u;
        facts->copyMoveMaxConcurrency = 1u;
        facts->deleteMaxConcurrency = 1u;
        facts->deleteRecycleBinMaxConcurrency = 1u;
        facts->maxComponentUtf16 = 255u;
        facts->copyOperation = _fault == ScriptedRouteFault::NonStrictBoolean ? 2 : TRUE;
        facts->moveOperation = TRUE;
        facts->nativeMoveOperation = TRUE;
        facts->deleteOperation = TRUE;
        facts->renameOperation = TRUE;
        facts->createDirectoryOperation = TRUE;
        facts->propertiesOperation = TRUE;
        facts->readOperation = TRUE;
        facts->writeOperation = TRUE;
        facts->recycleOperation = TRUE;
        facts->boundDelete = TRUE;
        facts->conditionalDelete = TRUE;
        facts->exclusiveStage = TRUE;
        facts->conditionalPublish = TRUE;
        facts->committedSize = TRUE;
        facts->preserveFileLink = TRUE;
        facts->preserveDirectoryLink = TRUE;
        facts->retargetInTree = TRUE;
        facts->exactLinkRemoval = TRUE;
        facts->cancellationAbort = TRUE;
        facts->cancellationDeadline = TRUE;
        facts->pathTextStableIdentity = _fault == ScriptedRouteFault::UnsupportedIdentity ? FALSE : TRUE;
        facts->casePreserving = TRUE;
        facts->preferredSeparator = L'/';
        facts->acceptedSeparators = accepted;
        facts->providerId = _fault == ScriptedRouteFault::OutsideArenaPointer
            ? L"selftest/typed-route"
            : (_fault == ScriptedRouteFault::MisalignedArenaPointer
                   ? reinterpret_cast<const wchar_t*>(arena->buffer + 1u)
                   : providerId);
        facts->pathProfileId = pathProfileId;
        facts->rootId = rootId;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE IsTransferPeerAllowed(const wchar_t*,
                                                    FileSystemOperation,
                                                    FileSystemTransferPeerRole,
                                                    const wchar_t*,
                                                    BOOL* allowed) noexcept override
    {
        if (allowed == nullptr)
        {
            return E_POINTER;
        }
        if (_fault == ScriptedRouteFault::PeerFailure)
        {
            return E_FAIL;
        }
        *allowed = _fault == ScriptedRouteFault::NonStrictPeer ? 2 : (_fault == ScriptedRouteFault::PeerDenied ? FALSE : TRUE);
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE ValidateChildName(const wchar_t*,
                                                const wchar_t*,
                                                FileSystemOperation,
                                                FileSystemChildNameValidation* validation) noexcept override
    {
        if (validation == nullptr)
        {
            return E_POINTER;
        }
        validation->sizeBytes = sizeof(*validation);
        validation->status = _fault == ScriptedRouteFault::InvalidNameResult
            ? static_cast<FileSystemChildNameStatus>(99u)
            : (_fault == ScriptedRouteFault::ProviderInvalidName
                   ? FILESYSTEM_CHILD_NAME_INVALID
                   : (_fault == ScriptedRouteFault::UnsupportedName ? FILESYSTEM_CHILD_NAME_UNSUPPORTED : FILESYSTEM_CHILD_NAME_VALID));
        validation->failureStatus = _fault == ScriptedRouteFault::ProviderInvalidName
            ? HRESULT_FROM_WIN32(ERROR_INVALID_NAME)
            : (_fault == ScriptedRouteFault::UnsupportedName ? HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED) : S_OK);
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetChildNameCollisionKey(const wchar_t*,
                                                       const wchar_t* childName,
                                                       FileSystemOperation,
                                                       FileSystemArena* arena,
                                                       const wchar_t** key,
                                                       unsigned long* requiredArenaBytes) noexcept override
    {
        if (arena == nullptr || key == nullptr || requiredArenaBytes == nullptr || childName == nullptr)
        {
            return E_INVALIDARG;
        }
        if (_fault == ScriptedRouteFault::CollisionOutsideArena)
        {
            *requiredArenaBytes = 4u * sizeof(wchar_t);
            arena->usedBytes = *requiredArenaBytes;
            *key = L"bad";
            return S_OK;
        }
        if (_fault == ScriptedRouteFault::EmptyCollisionKey)
        {
            return CopyOutput({}, arena, key, requiredArenaBytes);
        }
        if (_fault == ScriptedRouteFault::ChildStringArenaFallback)
        {
            static const std::wstring largeKey(
                FileSystemRouteContract::kNormalArenaBytes / sizeof(wchar_t) + 32u, L'k');
            return CopyOutput(largeKey, arena, key, requiredArenaBytes);
        }
        return CopyOutput(childName, arena, key, requiredArenaBytes);
    }

    HRESULT STDMETHODCALLTYPE JoinPath(const wchar_t* parentPath,
                                       const wchar_t* childName,
                                       FileSystemOperation,
                                       FileSystemArena* arena,
                                       const wchar_t** joinedPath,
                                       unsigned long* requiredArenaBytes) noexcept override
    {
        if (arena == nullptr || joinedPath == nullptr || requiredArenaBytes == nullptr || parentPath == nullptr || childName == nullptr)
        {
            return E_INVALIDARG;
        }
        std::wstring joined(parentPath);
        if (! joined.empty() && joined.back() != L'/')
        {
            joined.push_back(L'/');
        }
        joined.append(_fault == ScriptedRouteFault::JoinLeafMismatch ? L"different.txt" : childName);
        const HRESULT hr = CopyOutput(joined, arena, joinedPath, requiredArenaBytes);
        if (SUCCEEDED(hr) && _fault == ScriptedRouteFault::JoinRequiredMismatch)
        {
            *requiredArenaBytes += sizeof(wchar_t);
        }
        return hr;
    }

private:
    [[nodiscard]] static uint64_t Bytes(std::wstring_view text) noexcept
    {
        return (static_cast<uint64_t>(text.size()) + 1u) * sizeof(wchar_t);
    }

    [[nodiscard]] static const wchar_t* Copy(std::wstring_view text, FileSystemArena* arena) noexcept
    {
        const uint64_t bytes64 = Bytes(text);
        if (bytes64 > (std::numeric_limits<unsigned long>::max)())
        {
            return nullptr;
        }
        wchar_t* output = static_cast<wchar_t*>(
            AllocateFromFileSystemArena(arena, static_cast<unsigned long>(bytes64), alignof(wchar_t)));
        if (output == nullptr)
        {
            return nullptr;
        }
        if (! text.empty())
        {
            memcpy(output, text.data(), text.size() * sizeof(wchar_t));
        }
        output[text.size()] = L'\0';
        return output;
    }

    [[nodiscard]] static HRESULT CopyOutput(std::wstring_view text,
                                            FileSystemArena* arena,
                                            const wchar_t** output,
                                            unsigned long* requiredArenaBytes) noexcept
    {
        const uint64_t bytes64 = Bytes(text);
        if (bytes64 > (std::numeric_limits<unsigned long>::max)())
        {
            return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
        }
        *requiredArenaBytes = static_cast<unsigned long>(bytes64);
        if (arena->buffer == nullptr || arena->usedBytes > arena->capacityBytes ||
            *requiredArenaBytes > arena->capacityBytes - arena->usedBytes)
        {
            return HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER);
        }
        *output = Copy(text, arena);
        return *output != nullptr ? S_OK : E_OUTOFMEMORY;
    }

    ScriptedRouteFault _fault = ScriptedRouteFault::None;
};

std::array<PluginMetaData, 2> g_factoryContractMetaData{};
PluginMetaData g_nonContiguousFactoryMetaData{};

const PluginMetaData* GetFactoryContractMetaData0() noexcept
{
    return &g_factoryContractMetaData[0];
}

const PluginMetaData* GetFactoryContractMetaData1() noexcept
{
    return &g_factoryContractMetaData[1];
}

const PluginMetaData* GetNonContiguousFactoryMetaData() noexcept
{
    return &g_nonContiguousFactoryMetaData;
}

const char* GetEmptyFactoryContractSchema() noexcept
{
    return nullptr;
}

HRESULT CreateFactoryContractInstance(const FactoryOptions*, IHost*, void**) noexcept
{
    return E_NOTIMPL;
}

// ---------------------------------------------------------------------------
// Check helper (matches LocalizationTests pattern)
// ---------------------------------------------------------------------------
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

[[nodiscard]] bool SetRegistryStringValue(HKEY key,
                                          const wchar_t* valueName,
                                          DWORD type,
                                          std::wstring_view value,
                                          bool includeTerminator) noexcept
{
    const std::wstring stored(value);
    const size_t byteCount = (stored.size() + (includeTerminator ? 1u : 0u)) * sizeof(wchar_t);
    if (byteCount > static_cast<size_t>((std::numeric_limits<DWORD>::max)()))
    {
        return false;
    }

    return RegSetValueExW(key,
                          valueName,
                          0u,
                          type,
                          reinterpret_cast<const BYTE*>(stored.c_str()),
                          static_cast<DWORD>(byteCount)) == ERROR_SUCCESS;
}

[[nodiscard]] wil::unique_hkey CreateVolatileRegistryKey(HKEY parent, const wchar_t* subKey) noexcept
{
    wil::unique_hkey key;
    DWORD disposition = 0u;
    if (RegCreateKeyExW(parent,
                        subKey,
                        0u,
                        nullptr,
                        REG_OPTION_VOLATILE,
                        KEY_ALL_ACCESS,
                        nullptr,
                        key.put(),
                        &disposition) != ERROR_SUCCESS)
    {
        return {};
    }
    return key;
}

struct RegistryMutationContext
{
    std::wstring replacement;
    bool mutationSucceeded = false;
};

void ReplaceRegistryStringAfterSizeRead(HKEY key, const wchar_t* valueName, unsigned int attempt, void* cookie) noexcept
{
    auto* context = static_cast<RegistryMutationContext*>(cookie);
    if (context == nullptr || attempt != 0u)
    {
        return;
    }

    context->mutationSucceeded = SetRegistryStringValue(key, valueName, REG_SZ, context->replacement, true);
}

void TestRegistryAndWslCatalogContracts(bool& success) noexcept
{
    const std::wstring testPath = std::format(L"Software\\RedSalamander\\Tests\\PluginContract\\RegistryWsl-{}-{}",
                                               GetCurrentProcessId(),
                                               GetTickCount64());
    static_cast<void>(RegDeleteTreeW(HKEY_CURRENT_USER, testPath.c_str()));

    wil::unique_hkey testRoot = CreateVolatileRegistryKey(HKEY_CURRENT_USER, testPath.c_str());
    Check(testRoot.is_valid(), L"volatile per-user registry fixture opens", success);
    if (! testRoot)
    {
        return;
    }
    const auto cleanup = wil::scope_exit([&]() noexcept
    {
        testRoot.reset();
        static_cast<void>(RegDeleteTreeW(HKEY_CURRENT_USER, testPath.c_str()));
    });

    Check(SetRegistryStringValue(testRoot.get(), L"Terminated", REG_SZ, L"terminated", true),
          L"terminated registry fixture writes",
          success);
    const auto terminated = Common::Registry::ReadBoundedStringValue(testRoot.get(), L"Terminated");
    Check(terminated.has_value() && terminated.value() == L"terminated", L"bounded registry reader accepts terminated REG_SZ", success);

    Check(SetRegistryStringValue(testRoot.get(), L"NonTerminated", REG_SZ, L"nonterminated", false),
          L"nonterminated registry fixture writes",
          success);
    const auto nonterminated = Common::Registry::ReadBoundedStringValue(testRoot.get(), L"NonTerminated");
    Check(nonterminated.has_value() && nonterminated.value() == L"nonterminated",
          L"bounded registry reader uses returned length for nonterminated REG_SZ",
          success);

    const std::wstring longValue(4096u, L'x');
    Check(SetRegistryStringValue(testRoot.get(), L"Long", REG_SZ, longValue, true), L"long registry fixture writes", success);
    const auto longRead = Common::Registry::ReadBoundedStringValue(testRoot.get(), L"Long");
    Check(longRead.has_value() && longRead.value() == longValue, L"bounded registry reader accepts values longer than legacy buffers", success);

    Check(SetRegistryStringValue(testRoot.get(), L"Empty", REG_SZ, L"", true), L"empty registry fixture writes", success);
    Check(! Common::Registry::ReadBoundedStringValue(testRoot.get(), L"Empty").has_value(), L"bounded registry reader rejects empty values by policy", success);

    const std::array<BYTE, 3> oddBytes = {static_cast<BYTE>('o'), 0u, static_cast<BYTE>('d')};
    Check(RegSetValueExW(testRoot.get(), L"OddBytes", 0u, REG_SZ, oddBytes.data(), static_cast<DWORD>(oddBytes.size())) == ERROR_SUCCESS,
          L"odd-byte registry fixture writes",
          success);
    Check(! Common::Registry::ReadBoundedStringValue(testRoot.get(), L"OddBytes").has_value(), L"bounded registry reader rejects odd byte counts", success);

    const DWORD wrongTypeValue = 7u;
    Check(RegSetValueExW(testRoot.get(),
                         L"WrongType",
                         0u,
                         REG_DWORD,
                         reinterpret_cast<const BYTE*>(&wrongTypeValue),
                         sizeof(wrongTypeValue)) == ERROR_SUCCESS,
          L"wrong-type registry fixture writes",
          success);
    Check(! Common::Registry::ReadBoundedStringValue(testRoot.get(), L"WrongType").has_value(), L"bounded registry reader rejects wrong types", success);

    Common::Registry::StringReadOptions cappedOptions;
    cappedOptions.maxBytes = 32u;
    Check(! Common::Registry::ReadBoundedStringValue(testRoot.get(), L"Long", cappedOptions).has_value(),
          L"bounded registry reader rejects values above the caller cap",
          success);

    Check(SetRegistryStringValue(testRoot.get(), L"Changing", REG_SZ, L"small", true), L"changing registry fixture writes", success);
    RegistryMutationContext mutationContext{.replacement = std::wstring(2048u, L'm')};
    Common::Registry::StringReadOptions mutationOptions;
    mutationOptions.afterSizeReadObserver = &ReplaceRegistryStringAfterSizeRead;
    mutationOptions.observerCookie        = &mutationContext;
    const auto changedRead                = Common::Registry::ReadBoundedStringValue(testRoot.get(), L"Changing", mutationOptions);
    Check(mutationContext.mutationSucceeded && changedRead.has_value() && changedRead.value() == mutationContext.replacement,
          L"bounded registry reader retries a value that grows between reads",
          success);

    wil::unique_hkey catalogKey = CreateVolatileRegistryKey(testRoot.get(), L"Lxss");
    Check(catalogKey.is_valid(), L"volatile WSL catalog fixture opens", success);
    if (! catalogKey)
    {
        return;
    }

    const auto addDistribution = [&](const wchar_t* guid, std::wstring_view name, std::wstring_view basePath, std::optional<DWORD> modern) noexcept
    {
        wil::unique_hkey distributionKey = CreateVolatileRegistryKey(catalogKey.get(), guid);
        if (! distributionKey || ! SetRegistryStringValue(distributionKey.get(), L"DistributionName", REG_SZ, name, true) ||
            ! SetRegistryStringValue(distributionKey.get(), L"BasePath", REG_SZ, basePath, true))
        {
            return false;
        }
        if (! modern.has_value())
        {
            return true;
        }

        const DWORD value = modern.value();
        return RegSetValueExW(distributionKey.get(),
                              L"Modern",
                              0u,
                              REG_DWORD,
                              reinterpret_cast<const BYTE*>(&value),
                              sizeof(value)) == ERROR_SUCCESS;
    };

    Check(addDistribution(L"{00000000-0000-0000-0000-000000000001}", L"zulu", L"C:\\WSL\\Zulu", std::nullopt),
          L"WSL1 catalog fixture writes",
          success);
    Check(addDistribution(L"{00000000-0000-0000-0000-000000000002}", L"Alpha", L"C:\\WSL\\Alpha", 1u),
          L"WSL2 catalog fixture writes",
          success);
    Check(addDistribution(L"{00000000-0000-0000-0000-000000000003}", L"DOCKER-DESKTOP-data", L"C:\\WSL\\Docker", 1u),
          L"Docker utility catalog fixture writes",
          success);
    Check(addDistribution(L"{00000000-0000-0000-0000-000000000004}", L"Rancher-Desktop", L"C:\\WSL\\Rancher", 1u),
          L"Rancher utility catalog fixture writes",
          success);
    Check(addDistribution(L"not-a-guid", L"ignored", L"C:\\WSL\\Ignored", 1u), L"non-GUID catalog fixture writes", success);

    const auto distributions = Common::Wsl::EnumerateDistributions(testRoot.get(), L"Lxss");
    Check(distributions.size() == 2u, L"canonical WSL catalog excludes utility and malformed subkeys", success);
    if (distributions.size() == 2u)
    {
        Check(distributions[0].name == L"Alpha" && distributions[0].basePath == L"C:\\WSL\\Alpha" && distributions[0].isWsl2,
              L"canonical WSL catalog sorts names without case and preserves WSL2 metadata",
              success);
        Check(distributions[1].name == L"zulu" && distributions[1].basePath == L"C:\\WSL\\Zulu" && ! distributions[1].isWsl2,
              L"canonical WSL catalog preserves WSL1 default metadata",
              success);
    }
}

void TestDeleteOnCloseTemporaryFileContracts(bool& success) noexcept
{
    std::error_code error;
    const std::filesystem::path root = RedSalamander::TestSupport::AcquireTestDirectory(
        {.harnessSegment      = L"PluginContractTests",
         .leafSegment         = L"delete-on-close-temporary-file",
         .fallbackRunIdPrefix = L"plugin-contract",
         .kind                = RedSalamander::TestSupport::TestDirectoryKind::Scratch},
        error);
    Check(! error && ! root.empty(), L"delete-on-close helper acquires TestSandbox scratch", success);
    if (error || root.empty())
    {
        return;
    }

    std::filesystem::remove_all(root, error);
    error.clear();
    std::filesystem::create_directories(root, error);
    Check(! error, L"delete-on-close helper creates isolated scratch directory", success);
    if (error)
    {
        return;
    }
    const auto cleanup = wil::scope_exit([&]() noexcept
    {
        std::error_code cleanupError;
        std::filesystem::remove_all(root, cleanupError);
    });

    const Common::Files::DeleteOnCloseTemporaryFileOptions options{
        .prefix             = L"rst",
        .directory          = root.native(),
        .flagsAndAttributes = FILE_ATTRIBUTE_TEMPORARY | FILE_FLAG_SEQUENTIAL_SCAN,
    };

#if defined(ENABLE_TESTS)
    Common::Files::Testing::FailNextDeleteOnCloseTemporaryFileOpen(HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED));
    wil::unique_hfile failedFile;
    const HRESULT failedHr = Common::Files::CreateDeleteOnCloseTemporaryFile(options, failedFile);
    const bool reservationRemoved = std::filesystem::directory_iterator(root, error) == std::filesystem::directory_iterator{};
    Check(failedHr == HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED) && ! failedFile && ! error && reservationRemoved,
          L"failed reopen deletes the GetTempFileName reservation",
          success);
#endif

    struct ConcurrentResult
    {
        ConcurrentResult() = default;
        ConcurrentResult(const ConcurrentResult&)            = delete;
        ConcurrentResult& operator=(const ConcurrentResult&) = delete;
        ConcurrentResult(ConcurrentResult&&) noexcept        = default;
        ConcurrentResult& operator=(ConcurrentResult&&) noexcept = default;

        HRESULT hr = E_PENDING;
        wil::unique_hfile file;
        std::wstring path;
    };

    constexpr size_t kConcurrentFiles = 12u;
    std::array<ConcurrentResult, kConcurrentFiles> results{};
    std::latch ready(static_cast<ptrdiff_t>(kConcurrentFiles));
    std::latch start(1);
    std::vector<std::jthread> workers;
    workers.reserve(kConcurrentFiles);
    for (size_t index = 0u; index < kConcurrentFiles; ++index)
    {
        workers.emplace_back([&, index]() noexcept
        {
            ready.count_down();
            start.wait();
            ConcurrentResult& result = results[index];
            result.hr = Common::Files::CreateDeleteOnCloseTemporaryFile(options, result.file);
            if (FAILED(result.hr) || ! result.file)
            {
                return;
            }

            std::array<wchar_t, 32768u> path{};
            const DWORD length = GetFinalPathNameByHandleW(result.file.get(), path.data(), static_cast<DWORD>(path.size()), FILE_NAME_NORMALIZED);
            if (length == 0u || length >= path.size())
            {
                result.hr = HRESULT_FROM_WIN32(GetLastError());
                return;
            }
            result.path.assign(path.data(), length);
        });
    }
    ready.wait();
    start.count_down();
    workers.clear();

    std::unordered_set<std::wstring> paths;
    bool allCreated = true;
    bool policyPreserved = true;
    for (const ConcurrentResult& result : results)
    {
        allCreated = allCreated && SUCCEEDED(result.hr) && result.file && ! result.path.empty() && paths.insert(result.path).second;
        FILE_BASIC_INFO basicInfo{};
        policyPreserved = policyPreserved && result.file &&
                          GetFileInformationByHandleEx(result.file.get(), FileBasicInfo, &basicInfo, sizeof(basicInfo)) != FALSE &&
                          (basicInfo.FileAttributes & FILE_ATTRIBUTE_TEMPORARY) != 0u;
    }
    Check(allCreated && paths.size() == kConcurrentFiles, L"concurrent delete-on-close allocations are unique", success);
    Check(policyPreserved, L"delete-on-close helper preserves caller temporary-file attributes", success);

    for (ConcurrentResult& result : results)
    {
        result.file.reset();
    }
    error.clear();
    const bool emptyAfterClose = std::filesystem::directory_iterator(root, error) == std::filesystem::directory_iterator{};
    Check(! error && emptyAfterClose, L"closing every returned handle deletes every reserved pathname", success);
}

struct PackedFileInfoTestEntry
{
    std::wstring name;
    unsigned long fileIndex  = 0;
    unsigned long attributes = 0;
    uint64_t sizeBytes       = 0;
};

void TestPackedFileInfoBuffer(bool& success) noexcept
{
    Common::Plugins::PackedFileInfoBuffer buffer;
    Check(buffer.Build(std::vector<PackedFileInfoTestEntry>{}, [](const PackedFileInfoTestEntry&, FileInfo&) noexcept {}) == S_OK,
          L"packed FileInfo owner accepts an empty result",
          success);

    FileInfo* first        = reinterpret_cast<FileInfo*>(static_cast<uintptr_t>(1u));
    unsigned long byteSize = 1;
    unsigned long count    = 1;
    Check(buffer.GetBuffer(&first) == S_OK && first == nullptr, L"packed FileInfo empty result exposes a null buffer", success);
    Check(buffer.GetBufferSize(&byteSize) == S_OK && byteSize == 0, L"packed FileInfo empty result has zero used bytes", success);
    Check(buffer.GetCount(&count) == S_OK && count == 0, L"packed FileInfo empty result has zero entries", success);
    Check(buffer.Get(0, &first) == HRESULT_FROM_WIN32(ERROR_NO_MORE_FILES) && first == nullptr,
          L"packed FileInfo empty result rejects indexed access",
          success);

    const std::vector<PackedFileInfoTestEntry> entries = {
        {L"a", 7u, FILE_ATTRIBUTE_NORMAL, 11u},
        {L"longer-name.bin", 9u, FILE_ATTRIBUTE_ARCHIVE, 42u},
        {L"folder", 10u, FILE_ATTRIBUTE_DIRECTORY, 0u},
    };
    const HRESULT buildHr = buffer.Build(entries,
                                         [](const PackedFileInfoTestEntry& source, FileInfo& entry) noexcept
    {
        entry.FileIndex      = source.fileIndex;
        entry.FileAttributes = source.attributes;
        entry.EndOfFile      = static_cast<__int64>(source.sizeBytes);
        entry.AllocationSize = static_cast<__int64>(source.sizeBytes);
    });
    Check(buildHr == S_OK, L"packed FileInfo owner builds a multi-entry result", success);
    Check(buffer.GetBuffer(&first) == S_OK && first != nullptr, L"packed FileInfo multi-entry result exposes its buffer", success);
    Check(buffer.GetBufferSize(&byteSize) == S_OK && byteSize >= sizeof(FileInfo),
          L"packed FileInfo multi-entry result exposes its used byte size",
          success);
    Check(first != nullptr && (reinterpret_cast<uintptr_t>(first) % alignof(FileInfo)) == 0u,
          L"packed FileInfo buffer base honors FileInfo alignment",
          success);
    Check(buffer.GetCount(&count) == S_OK && count == entries.size(), L"packed FileInfo count matches the source entries", success);

    for (unsigned long index = 0; index < entries.size(); ++index)
    {
        FileInfo* entry = nullptr;
        const HRESULT getHr = buffer.Get(index, &entry);
        const std::wstring_view actualName = entry ? std::wstring_view(entry->FileName, entry->FileNameSize / sizeof(wchar_t)) : std::wstring_view{};
        Check(getHr == S_OK && entry != nullptr && actualName == entries[index].name && entry->FileIndex == entries[index].fileIndex &&
                  entry->FileAttributes == entries[index].attributes && entry->EndOfFile == static_cast<__int64>(entries[index].sizeBytes),
              std::format(L"packed FileInfo entry {} preserves name and provider metadata", index).c_str(),
              success);
        Check(entry != nullptr && (reinterpret_cast<uintptr_t>(entry) % alignof(FileInfo)) == 0u,
              std::format(L"packed FileInfo entry {} honors FileInfo alignment", index).c_str(),
              success);
    }

    const unsigned long originalOffset = first ? first->NextEntryOffset : 0u;
    if (first)
    {
        first->NextEntryOffset = 1u;
    }
    FileInfo* malformed = nullptr;
    Check(buffer.Get(1, &malformed) == HRESULT_FROM_WIN32(ERROR_INVALID_DATA) && malformed == nullptr,
          L"packed FileInfo traversal rejects a short unaligned next offset",
          success);
    if (first)
    {
        first->NextEntryOffset = originalOffset;
        ++first->FileNameSize;
    }
    Check(buffer.Get(0, &malformed) == HRESULT_FROM_WIN32(ERROR_INVALID_DATA) && malformed == nullptr,
          L"packed FileInfo traversal rejects an odd UTF-16 name byte count",
          success);
    if (first)
    {
        --first->FileNameSize;
    }

    const FileInfo* located = nullptr;
    constexpr unsigned long headerBytes = static_cast<unsigned long>(offsetof(FileInfo, FileName));
    Check(Common::Plugins::LocatePackedFileInfoRecord(first, headerBytes - 1u, count, 0u, &located) ==
              HRESULT_FROM_WIN32(ERROR_INVALID_DATA) && located == nullptr,
          L"packed FileInfo traversal rejects a truncated first header before reading fields",
          success);
    Check(originalOffset > headerBytes &&
              Common::Plugins::LocatePackedFileInfoRecord(first, originalOffset + headerBytes - 1u, count, 0u, &located) ==
                  HRESULT_FROM_WIN32(ERROR_INVALID_DATA) && located == nullptr,
          L"packed FileInfo traversal rejects a truncated later header even when index zero was requested",
          success);
    if (first)
    {
        const unsigned long savedNameSize = first->FileNameSize;
        first->FileNameSize = byteSize;
        Check(Common::Plugins::LocatePackedFileInfoRecord(first, byteSize, count, 0u, &located) ==
                  HRESULT_FROM_WIN32(ERROR_INVALID_DATA),
              L"packed FileInfo traversal rejects a name extending beyond its record",
              success);
        first->FileNameSize = savedNameSize;

        first->NextEntryOffset = 0u;
        Check(Common::Plugins::LocatePackedFileInfoRecord(first, byteSize, count, 0u, &located) ==
                  HRESULT_FROM_WIN32(ERROR_INVALID_DATA),
              L"packed FileInfo traversal rejects an early terminal record",
              success);
        first->NextEntryOffset = originalOffset;
    }

    FileInfo* finalEntry = nullptr;
    Check(buffer.Get(2u, &finalEntry) == S_OK && finalEntry != nullptr,
          L"packed FileInfo malformed corpus locates the final valid record",
          success);
    if (finalEntry)
    {
        finalEntry->NextEntryOffset = static_cast<unsigned long>(alignof(FileInfo));
        located = nullptr;
        Check(Common::Plugins::LocatePackedFileInfoRecord(first, byteSize, count, 0u, &located) ==
                  HRESULT_FROM_WIN32(ERROR_INVALID_DATA) && located == nullptr,
              L"packed FileInfo traversal rejects a nonzero terminal offset before returning an earlier record",
              success);
        finalEntry->NextEntryOffset = 0u;
    }

    Check(buffer.GetBuffer(nullptr) == E_POINTER && buffer.GetBufferSize(nullptr) == E_POINTER && buffer.GetAllocatedSize(nullptr) == E_POINTER &&
              buffer.GetCount(nullptr) == E_POINTER && buffer.Get(0, nullptr) == E_POINTER,
          L"packed FileInfo query methods reject null out-parameters",
          success);
}

// ---------------------------------------------------------------------------
// Export function typedefs
// ---------------------------------------------------------------------------
using PfnEnumeratePlugins       = HRESULT(__stdcall*)(REFIID, const PluginMetaData**, unsigned int*);
using PfnCreate                 = HRESULT(__stdcall*)(REFIID, const FactoryOptions*, IHost*, const wchar_t*, void**);
using PfnGetConfigurationSchema = HRESULT(__stdcall*)(REFIID, const wchar_t*, const char**);
using PfnRunDebugSelfTests      = HRESULT(__stdcall*)(unsigned int*, unsigned int*);
using PfnRunTerminalVtUpgradeCorpus = HRESULT(__stdcall*)(TerminalVtUpgradeTestContract::Evidence*);
using PfnPluginShutdown         = void(__stdcall*)() noexcept;
using PfnPluginCanUnloadNow     = BOOL(__stdcall*)() noexcept;
using PfnTerminalAccessibilityLocalizationSelfTests = HRESULT(__stdcall*)(unsigned int*, unsigned int*) noexcept;
using PfnTerminalDebugGetDiagnosticText = HRESULT(__stdcall*)(ITerminal*, TerminalOwnedUtf16*) noexcept;
#if defined(_DEBUG)
using PfnDebugCurlRuntimeProbe  = HRESULT(__stdcall*)() noexcept;
#endif

#if defined(_DEBUG)
constexpr bool kDebugSelfTestExportsRequired = true;
#else
constexpr bool kDebugSelfTestExportsRequired = false;
#endif

// ---------------------------------------------------------------------------
// Plugin sets (paths relative to the exe output dir; exe runs from .build\x64\Debug\)
// DLLs live in the Plugins\ subfolder per REVIEWER CORRECTION 1.
// ---------------------------------------------------------------------------
constexpr std::array<std::wstring_view, 8> kFilesystemDlls = {
    L"Plugins\\FileSystem.dll",
    L"Plugins\\FileSystem7z.dll",
    L"Plugins\\FileSystemCurl.dll",
    L"Plugins\\FileSystemDummy.dll",
    L"Plugins\\FileSystemGoogleDrive.dll",
    L"Plugins\\FileSystemMicrosoftDrive.dll",
    L"Plugins\\FileSystemMtp.dll",
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

constexpr std::array<std::wstring_view, 1> kTerminalDlls = {
    L"Plugins\\Terminal.dll",
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
    yyjson_doc* doc = yyjson_read_opts(const_cast<char*>(utf8), std::strlen(utf8), YYJSON_READ_NOFLAG, nullptr, &err);
    return unique_yyjson_doc(doc);
}

[[nodiscard]] std::string_view GetJsonStringMember(yyjson_val* object, const char* key) noexcept
{
    yyjson_val* value = object != nullptr ? yyjson_obj_get(object, key) : nullptr;
    const char* text = value != nullptr && yyjson_is_str(value) ? yyjson_get_str(value) : nullptr;
    return text != nullptr ? std::string_view(text) : std::string_view{};
}

[[nodiscard]] bool JsonScalarMemberEqual(yyjson_val* left, yyjson_val* right, const char* key) noexcept
{
    yyjson_val* leftValue = left != nullptr ? yyjson_obj_get(left, key) : nullptr;
    yyjson_val* rightValue = right != nullptr ? yyjson_obj_get(right, key) : nullptr;
    if (leftValue == nullptr || rightValue == nullptr)
    {
        return leftValue == rightValue;
    }
    if (yyjson_is_str(leftValue) && yyjson_is_str(rightValue))
    {
        return GetJsonStringMember(left, key) == GetJsonStringMember(right, key);
    }
    if (yyjson_is_uint(leftValue) && yyjson_is_uint(rightValue))
    {
        return yyjson_get_uint(leftValue) == yyjson_get_uint(rightValue);
    }
    if (yyjson_is_bool(leftValue) && yyjson_is_bool(rightValue))
    {
        return yyjson_get_bool(leftValue) == yyjson_get_bool(rightValue);
    }
    return false;
}

[[nodiscard]] bool TerminalSchemaInvariantsEqual(yyjson_doc* leftDocument, yyjson_doc* rightDocument) noexcept
{
    yyjson_val* leftRoot = leftDocument != nullptr ? yyjson_doc_get_root(leftDocument) : nullptr;
    yyjson_val* rightRoot = rightDocument != nullptr ? yyjson_doc_get_root(rightDocument) : nullptr;
    if (! yyjson_is_obj(leftRoot) || ! yyjson_is_obj(rightRoot) ||
        ! JsonScalarMemberEqual(leftRoot, rightRoot, "version"))
    {
        return false;
    }

    yyjson_val* leftFields = yyjson_obj_get(leftRoot, "fields");
    yyjson_val* rightFields = yyjson_obj_get(rightRoot, "fields");
    if (! yyjson_is_arr(leftFields) || ! yyjson_is_arr(rightFields) ||
        yyjson_arr_size(leftFields) != yyjson_arr_size(rightFields))
    {
        return false;
    }

    constexpr std::array scalarMembers = {"key", "type", "default", "min", "max"};
    for (size_t fieldIndex = 0u; fieldIndex < yyjson_arr_size(leftFields); ++fieldIndex)
    {
        yyjson_val* leftField = yyjson_arr_get(leftFields, fieldIndex);
        yyjson_val* rightField = yyjson_arr_get(rightFields, fieldIndex);
        if (! yyjson_is_obj(leftField) || ! yyjson_is_obj(rightField) ||
            ! std::ranges::all_of(scalarMembers, [leftField, rightField](const char* member) noexcept {
                return JsonScalarMemberEqual(leftField, rightField, member);
            }))
        {
            return false;
        }

        yyjson_val* leftOptions = yyjson_obj_get(leftField, "options");
        yyjson_val* rightOptions = yyjson_obj_get(rightField, "options");
        if ((leftOptions == nullptr) != (rightOptions == nullptr))
        {
            return false;
        }
        if (leftOptions == nullptr)
        {
            continue;
        }
        if (! yyjson_is_arr(leftOptions) || ! yyjson_is_arr(rightOptions) ||
            yyjson_arr_size(leftOptions) != yyjson_arr_size(rightOptions))
        {
            return false;
        }
        for (size_t optionIndex = 0u; optionIndex < yyjson_arr_size(leftOptions); ++optionIndex)
        {
            if (! JsonScalarMemberEqual(
                    yyjson_arr_get(leftOptions, optionIndex), yyjson_arr_get(rightOptions, optionIndex), "value"))
            {
                return false;
            }
        }
    }
    return true;
}

[[nodiscard]] bool HasExpectedTerminalSchemaShape(yyjson_doc* document) noexcept
{
    yyjson_val* root = document != nullptr ? yyjson_doc_get_root(document) : nullptr;
    yyjson_val* fields = yyjson_is_obj(root) ? yyjson_obj_get(root, "fields") : nullptr;
    constexpr std::array<std::string_view, 10> expectedKeys = {
        "defaultShell",
        "fontFamily",
        "fontSizeDip",
        "maxFormattedMiB",
        "pasteMaxBytes",
        "warnOnUnsafePaste",
        "followPathWhenIdle",
        "hyperlinkPolicy",
        "osc52Policy",
        "osc52MaxBytes",
    };
    if (! yyjson_is_arr(fields) || yyjson_arr_size(fields) != expectedKeys.size())
    {
        return false;
    }
    for (size_t index = 0u; index < expectedKeys.size(); ++index)
    {
        if (GetJsonStringMember(yyjson_arr_get(fields, index), "key") != expectedKeys[index])
        {
            return false;
        }
    }
    return true;
}

[[nodiscard]] bool WaitForPluginQuiet(PfnPluginCanUnloadNow canUnload, DWORD timeoutMilliseconds = 5000u) noexcept
{
    if (canUnload == nullptr)
    {
        return false;
    }
    const ULONGLONG deadline = GetTickCount64() + timeoutMilliseconds;
    do
    {
        if (canUnload() == TRUE)
        {
            return true;
        }
        Sleep(10u);
    } while (GetTickCount64() < deadline);
    return canUnload() == TRUE;
}

void TestTerminalLocalizedContracts(bool& success, bool runTestOnlyContracts = true) noexcept
{
    constexpr std::wstring_view relativePath = L"Plugins\\Terminal.dll";
    const std::filesystem::path executableDirectory(GetExeDir());
    const std::filesystem::path terminalPath = executableDirectory / relativePath;

    {
        wil::unique_hmodule module(LoadLibraryExW(terminalPath.c_str(), nullptr, 0));
        Check(static_cast<bool>(module), L"Terminal.dll: localization contract loads DLL", success);
        if (! module)
        {
            return;
        }

        const auto enumerate = reinterpret_cast<PfnEnumeratePlugins>(GetProcAddress(module.get(), "RedSalamanderEnumeratePlugins"));
        const auto schema = reinterpret_cast<PfnGetConfigurationSchema>(GetProcAddress(module.get(), "RedSalamanderGetConfigurationSchema"));
        const auto accessibilityTests = reinterpret_cast<PfnTerminalAccessibilityLocalizationSelfTests>(
            GetProcAddress(module.get(), "RedSalamanderTerminalAccessibilityLocalizationSelfTests"));
        const auto shutdown = reinterpret_cast<PfnPluginShutdown>(GetProcAddress(module.get(), "RedSalamanderPluginShutdown"));
        const auto canUnload = reinterpret_cast<PfnPluginCanUnloadNow>(GetProcAddress(module.get(), "RedSalamanderPluginCanUnloadNow"));
        Check(enumerate != nullptr && schema != nullptr && shutdown != nullptr && canUnload != nullptr,
              L"Terminal.dll: localization contract resolves schema and lifecycle exports",
              success);
        if (enumerate == nullptr || schema == nullptr || shutdown == nullptr || canUnload == nullptr)
        {
            return;
        }
        if (runTestOnlyContracts)
        {
            Check(accessibilityTests != nullptr,
                  L"Terminal.dll: test-enabled localization contract resolves the accessibility selftest export",
                  success);
            if (accessibilityTests == nullptr)
            {
                return;
            }
        }

        const HRESULT registerHr = Localization::RegisterResourceOwner(L"Terminal", module.get());
        Check(registerHr == S_OK, L"Terminal.dll: localization contract registers its resource owner", success);
        if (FAILED(registerHr))
        {
            return;
        }
        const auto cleanupLocalization = wil::scope_exit([&]() noexcept
        {
            static_cast<void>(Localization::ApplyLanguagePreference({.kind = Localization::LanguagePreferenceKind::System}));
            Localization::UnregisterResourceOwner(module.get());
            shutdown();
        });

        const PluginMetaData* metadata = nullptr;
        unsigned int count = 0u;
        const HRESULT enumerateHr = enumerate(__uuidof(ITerminal), &metadata, &count);
        Check(enumerateHr == S_OK && metadata != nullptr && count == 1u,
              L"Terminal.dll: localization contract enumerates the embedded terminal",
              success);
        if (FAILED(enumerateHr) || metadata == nullptr || count != 1u)
        {
            return;
        }

        const auto readSchema = [&](std::wstring_view culture, std::string& output) noexcept
        {
            const HRESULT languageHr = Localization::ApplyLanguagePreference(
                {.kind = Localization::LanguagePreferenceKind::Culture, .culture = std::wstring(culture)});
            const char* json = nullptr;
            const HRESULT schemaHr = SUCCEEDED(languageHr)
                ? schema(__uuidof(ITerminal), metadata[0].id, &json)
                : languageHr;
            output = SUCCEEDED(schemaHr) && json != nullptr ? json : "";
            return languageHr == S_OK && schemaHr == S_OK && ! output.empty();
        };

        std::string englishSchema;
        std::string frenchSchema;
        std::string japaneseSchema;
        Check(readSchema(L"en-US", englishSchema) && readSchema(L"fr-FR", frenchSchema) &&
                  readSchema(L"ja-JP", japaneseSchema),
              L"Terminal.dll: embedded English and French/Japanese satellite schemas load live",
              success);

        unique_yyjson_doc englishDocument = ParseJson(englishSchema.c_str());
        unique_yyjson_doc frenchDocument = ParseJson(frenchSchema.c_str());
        unique_yyjson_doc japaneseDocument = ParseJson(japaneseSchema.c_str());
        Check(englishDocument && frenchDocument && japaneseDocument,
              L"Terminal.dll: every localized schema is valid UTF-8 JSON",
              success);
        if (englishDocument && frenchDocument && japaneseDocument)
        {
            yyjson_val* englishRoot = yyjson_doc_get_root(englishDocument.get());
            yyjson_val* frenchRoot = yyjson_doc_get_root(frenchDocument.get());
            yyjson_val* japaneseRoot = yyjson_doc_get_root(japaneseDocument.get());
            Check(GetJsonStringMember(englishRoot, "title") == "Embedded Terminal" &&
                      GetJsonStringMember(frenchRoot, "title") == "Terminal intégré" &&
                      GetJsonStringMember(japaneseRoot, "title") == "組み込みターミナル",
                  L"Terminal.dll: schema display text follows the selected culture",
                  success);
            Check(HasExpectedTerminalSchemaShape(englishDocument.get()) &&
                      TerminalSchemaInvariantsEqual(englishDocument.get(), frenchDocument.get()) &&
                      TerminalSchemaInvariantsEqual(englishDocument.get(), japaneseDocument.get()),
                  L"Terminal.dll: localized schemas preserve invariant keys, types, defaults, bounds, and option values",
                  success);

            yyjson_val* frenchFields = yyjson_obj_get(frenchRoot, "fields");
            Check(GetJsonStringMember(yyjson_arr_get(frenchFields, 0u), "label") == "Interpréteur Windows par défaut",
                  L"Terminal.dll: localized field labels come from the French satellite",
                  success);
        }

        if (runTestOnlyContracts)
        {
            unsigned int accessibilityPassed = 0u;
            unsigned int accessibilityFailed = 0u;
            const HRESULT accessibilityHr = accessibilityTests(&accessibilityPassed, &accessibilityFailed);
            Check(accessibilityHr == S_OK && accessibilityPassed == 4u && accessibilityFailed == 0u,
                  L"Terminal.dll: live UI Automation Name resolves through the active Japanese resource satellite",
                  success);
        }
        Check(WaitForPluginQuiet(canUnload), L"Terminal.dll: localization-only characterization reaches its unload quiet point", success);
    }

    if (! runTestOnlyContracts)
    {
        return;
    }

    std::error_code error;
    const std::filesystem::path fixtureRoot = RedSalamander::TestSupport::AcquireTestDirectory(
        {.harnessSegment = L"PluginContractTests",
         .leafSegment = L"terminal-localized-runtime-diagnostic",
         .fallbackRunIdPrefix = L"plugin-contract",
         .kind = RedSalamander::TestSupport::TestDirectoryKind::Scratch},
        error);
    Check(! error && ! fixtureRoot.empty(), L"Terminal.dll: localized runtime diagnostic acquires TestSandbox scratch", success);
    if (error || fixtureRoot.empty())
    {
        return;
    }

    const std::filesystem::path fixtureLang = fixtureRoot / L"Lang";
    std::filesystem::create_directories(fixtureLang, error);
    const std::filesystem::path fixtureModulePath = fixtureRoot / L"Terminal.dll";
    const std::filesystem::path sourceSatellitePath = executableDirectory / L"Lang" / L"Terminal-ja-JP.dll";
    const std::filesystem::path fixtureSatellitePath = fixtureLang / L"Terminal-ja-JP.dll";
    if (! error)
    {
        std::filesystem::copy_file(terminalPath, fixtureModulePath, std::filesystem::copy_options::overwrite_existing, error);
    }
    if (! error)
    {
        std::filesystem::copy_file(
            sourceSatellitePath, fixtureSatellitePath, std::filesystem::copy_options::overwrite_existing, error);
    }
    Check(! error, L"Terminal.dll: isolated fixture contains the plugin and Japanese satellite but no private runtime", success);
    if (error)
    {
        std::filesystem::remove_all(fixtureRoot, error);
        return;
    }

    {
        wil::unique_hmodule module(LoadLibraryExW(fixtureModulePath.c_str(), nullptr, 0));
        Check(static_cast<bool>(module), L"Terminal.dll: isolated missing-runtime fixture loads", success);
        if (! module)
        {
            std::filesystem::remove_all(fixtureRoot, error);
            return;
        }

        const auto enumerate = reinterpret_cast<PfnEnumeratePlugins>(GetProcAddress(module.get(), "RedSalamanderEnumeratePlugins"));
        const auto create = reinterpret_cast<PfnCreate>(GetProcAddress(module.get(), "RedSalamanderCreate"));
        const auto getDiagnostic = reinterpret_cast<PfnTerminalDebugGetDiagnosticText>(
            GetProcAddress(module.get(), "RedSalamanderTerminalDebugGetDiagnosticText"));
        const auto shutdown = reinterpret_cast<PfnPluginShutdown>(GetProcAddress(module.get(), "RedSalamanderPluginShutdown"));
        const auto canUnload = reinterpret_cast<PfnPluginCanUnloadNow>(GetProcAddress(module.get(), "RedSalamanderPluginCanUnloadNow"));
        Check(enumerate != nullptr && create != nullptr && getDiagnostic != nullptr && shutdown != nullptr && canUnload != nullptr,
              L"Terminal.dll: isolated diagnostic fixture resolves factory, diagnostic, and lifecycle exports",
              success);
        if (enumerate == nullptr || create == nullptr || getDiagnostic == nullptr || shutdown == nullptr || canUnload == nullptr)
        {
            return;
        }

        const HRESULT registerHr = Localization::RegisterResourceOwner(L"Terminal", module.get());
        const HRESULT languageHr = SUCCEEDED(registerHr)
            ? Localization::ApplyLanguagePreference(
                  {.kind = Localization::LanguagePreferenceKind::Culture, .culture = L"ja-JP"})
            : registerHr;
        Check(registerHr == S_OK && languageHr == S_OK,
              L"Terminal.dll: isolated diagnostic fixture activates its Japanese resource owner",
              success);
        if (FAILED(registerHr) || FAILED(languageHr))
        {
            return;
        }
        const auto cleanupLocalization = wil::scope_exit([&]() noexcept
        {
            static_cast<void>(Localization::ApplyLanguagePreference({.kind = Localization::LanguagePreferenceKind::System}));
            Localization::UnregisterResourceOwner(module.get());
        });

        const PluginMetaData* metadata = nullptr;
        unsigned int count = 0u;
        const HRESULT enumerateHr = enumerate(__uuidof(ITerminal), &metadata, &count);
        wil::unique_hwnd parent(CreateWindowExW(
            0u, L"STATIC", L"", WS_OVERLAPPED, 0, 0, 1, 1, nullptr, nullptr, GetModuleHandleW(nullptr), nullptr));
        Check(enumerateHr == S_OK && metadata != nullptr && count == 1u && parent != nullptr,
              L"Terminal.dll: isolated diagnostic fixture enumerates and creates a valid parent window",
              success);
        if (FAILED(enumerateHr) || metadata == nullptr || count != 1u || ! parent)
        {
            return;
        }

        FactoryOptions options{};
        wil::com_ptr_nothrow<ITerminal> terminal;
        const HRESULT createHr = create(__uuidof(ITerminal), &options, &g_nullHost, metadata[0].id, terminal.put_void());
        TerminalOpenContext context{};
        context.sizeBytes = sizeof(context);
        context.parentWindow = parent.get();
        context.sourceGeneration = 1u;
        context.sourceLocation.sizeBytes = sizeof(context.sourceLocation);
        context.sourceLocation.kind = TerminalLocationKind::Unsupported;
        context.launchLocation.sizeBytes = sizeof(context.launchLocation);
        context.launchLocation.kind = TerminalLocationKind::Unsupported;
        const HRESULT openHr = SUCCEEDED(createHr) && terminal ? terminal->Open(&context) : createHr;
        Check(createHr == S_OK && terminal != nullptr && openHr == S_OK,
              L"Terminal.dll: missing private runtime enters the supported diagnostic lifecycle",
              success);

        TerminalOwnedUtf16 diagnostic{};
        const HRESULT diagnosticHr = terminal ? getDiagnostic(terminal.get(), &diagnostic) : E_UNEXPECTED;
        wil::unique_cotaskmem_string diagnosticOwner(diagnostic.data);
        const std::wstring_view diagnosticText = diagnostic.data != nullptr
            ? std::wstring_view(diagnostic.data, diagnostic.length)
            : std::wstring_view{};
        Check(diagnosticHr == S_OK &&
                  diagnosticText.find(L"組み込みターミナル エンジンを利用できません。") != std::wstring_view::npos &&
                  diagnosticText.find(L"プライベート Ghostty ランタイムを検証用に排他的に開けませんでした") != std::wstring_view::npos &&
                  diagnosticText.find(L"The private Ghostty runtime") == std::wstring_view::npos,
              L"Terminal.dll: production missing-runtime diagnostic is formatted from Japanese resources",
              success);

        if (terminal)
        {
            static_cast<void>(terminal->Close());
            terminal.reset();
        }
        parent.reset();
        Check(WaitForPluginQuiet(canUnload), L"Terminal.dll: isolated diagnostic fixture reaches its unload quiet point", success);
        shutdown();
    }

    std::filesystem::remove_all(fixtureRoot, error);
    Check(! error, L"Terminal.dll: isolated localized diagnostic fixture cleans its scratch directory", success);
}

// ---------------------------------------------------------------------------
// Step 2: Enumerate-and-validate (metadata + schema, no instance creation)
// ---------------------------------------------------------------------------
void TestPackagedGhosttyRuntimeLoad(bool& success) noexcept
{
    const std::wstring runtimePath = GetExeDir() + L"Plugins\\TerminalRuntime\\ghostty-vt.dll";
    wil::unique_hmodule runtime(LoadLibraryExW(
        runtimePath.c_str(),
        nullptr,
        LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR | LOAD_LIBRARY_SEARCH_APPLICATION_DIR | LOAD_LIBRARY_SEARCH_SYSTEM32));
    Check(static_cast<bool>(runtime), L"package Ghostty runtime loads natively", success);
    if (runtime)
    {
        Check(GetProcAddress(runtime.get(), "ghostty_build_info") != nullptr,
              L"package Ghostty runtime resolves its ABI identity export",
              success);
    }
}

// A plugin declares its configuration surface through GetConfigurationSchema, and the host renders
// and serializes that surface through Common::PluginConfiguration. Nothing forced the two halves to
// agree: a plugin could declare an integer-typed `value` field and then reject the JSON integer the
// shared codec emits for it, which makes SetConfiguration fail wholesale and silently revert every
// other field to compiled defaults. This drives the schema's own declared defaults back through the
// plugin exactly the way the Preferences page would, and requires the plugin to accept them.
void TestSchemaDefaultsRoundTrip(PfnCreate create,
                                 const IID& expectedIid,
                                 const wchar_t* pluginId,
                                 const char* schemaJsonUtf8,
                                 std::wstring_view relPath,
                                 bool& success) noexcept
{
    const Common::PluginConfiguration::SchemaParseResult schema =
        Common::PluginConfiguration::ParseSchema(schemaJsonUtf8);
    Check(! schema.HasErrors(),
          std::format(L"{}: configuration schema parses without errors for pluginId={}", relPath, pluginId).c_str(),
          success);
    if (schema.HasErrors() || schema.fields.empty())
    {
        return;
    }

    std::vector<Common::PluginConfiguration::FieldValue> defaults;
    defaults.reserve(schema.fields.size());
    for (const Common::PluginConfiguration::Field& field : schema.fields)
    {
        defaults.push_back(Common::PluginConfiguration::MakeDefaultValue(field));
    }

    std::string defaultsJson;
    const HRESULT serializeHr =
        Common::PluginConfiguration::SerializeConfiguration("", schema.fields, defaults, defaultsJson);
    Check(serializeHr == S_OK && ! defaultsJson.empty(),
          std::format(L"{}: shared codec serializes the schema defaults for pluginId={}", relPath, pluginId).c_str(),
          success);
    if (FAILED(serializeHr) || defaultsJson.empty())
    {
        return;
    }

    FactoryOptions options{};
    wil::com_ptr_nothrow<IUnknown> instance;
    const HRESULT createHr = create(expectedIid, &options, &g_nullHost, pluginId, instance.put_void());
    if (FAILED(createHr) || ! instance)
    {
        // Instance creation is covered by the dedicated per-interface tests; a plugin that cannot be
        // created here has already failed those and must not also fail this one.
        return;
    }

    wil::com_ptr_nothrow<IInformations> information;
    const HRESULT qiHr = instance->QueryInterface(__uuidof(IInformations), information.put_void());
    if (FAILED(qiHr) || ! information)
    {
        return;
    }

    const HRESULT acceptHr = information->SetConfiguration(defaultsJson.c_str());
    Check(acceptHr == S_OK,
          std::format(L"{}: SetConfiguration accepts its own schema defaults for pluginId={}", relPath, pluginId).c_str(),
          success);
    if (FAILED(acceptHr))
    {
        return;
    }

    // The accepted configuration comes back out through GetConfiguration and gets persisted; the
    // returned form must therefore still satisfy the plugin's own schema and be re-acceptable, or
    // the setting degrades on the next launch instead of at the moment it was saved.
    const char* persisted = nullptr;
    const HRESULT getHr = information->GetConfiguration(&persisted);
    Check(getHr == S_OK && persisted != nullptr,
          std::format(L"{}: GetConfiguration returns the accepted configuration for pluginId={}", relPath, pluginId).c_str(),
          success);
    if (FAILED(getHr) || persisted == nullptr)
    {
        return;
    }

    const std::string persistedJson(persisted);
    const Common::PluginConfiguration::ConfigurationParseResult reparsed =
        Common::PluginConfiguration::ParseConfiguration(schema.fields, persistedJson);
    Check(! reparsed.HasErrors(),
          std::format(L"{}: persisted configuration still matches the declared schema for pluginId={}", relPath, pluginId).c_str(),
          success);

    const HRESULT reacceptHr = information->SetConfiguration(persistedJson.c_str());
    Check(reacceptHr == S_OK,
          std::format(L"{}: SetConfiguration round-trips its own GetConfiguration output for pluginId={}", relPath, pluginId).c_str(),
          success);
}

bool TestEnumerateAndSchema(std::wstring_view relPath, const IID& expectedIid, bool& success) noexcept
{
    const std::wstring exeDir  = GetExeDir();
    const std::wstring absPath = exeDir + std::wstring(relPath);

    const std::wstring loadMsg = std::format(L"{}: DLL loads", relPath);
    wil::unique_hmodule mod(LoadLibraryExW(absPath.c_str(), nullptr, 0));
    Check(static_cast<bool>(mod), loadMsg.c_str(), success);
    if (! mod)
    {
        return false;
    }

    // Resolve exports
    const auto pfnEnumerate = reinterpret_cast<PfnEnumeratePlugins>(GetProcAddress(mod.get(), "RedSalamanderEnumeratePlugins"));
    const auto pfnCreate    = reinterpret_cast<PfnCreate>(GetProcAddress(mod.get(), "RedSalamanderCreate"));
    const auto pfnSchema    = reinterpret_cast<PfnGetConfigurationSchema>(GetProcAddress(mod.get(), "RedSalamanderGetConfigurationSchema"));

    Check(pfnEnumerate != nullptr, std::format(L"{}: RedSalamanderEnumeratePlugins export resolves", relPath).c_str(), success);
    Check(pfnCreate != nullptr, std::format(L"{}: RedSalamanderCreate export resolves", relPath).c_str(), success);
    Check(pfnSchema != nullptr, std::format(L"{}: RedSalamanderGetConfigurationSchema export resolves", relPath).c_str(), success);

    if (pfnEnumerate == nullptr || pfnCreate == nullptr || pfnSchema == nullptr)
    {
        return false;
    }

    // Happy-path enumerate with correct IID
    const PluginMetaData* metaData = nullptr;
    unsigned int count             = 0;
    const HRESULT hrEnum           = pfnEnumerate(expectedIid, &metaData, &count);
    Check(hrEnum == S_OK, std::format(L"{}: EnumeratePlugins(correct IID) returns S_OK", relPath).c_str(), success);
    Check(count >= 1, std::format(L"{}: EnumeratePlugins returns count >= 1", relPath).c_str(), success);
    Check(IsValidEnumeratedPluginCount(count), std::format(L"{}: EnumeratePlugins count stays within the host contract", relPath).c_str(), success);
    Check(metaData != nullptr, std::format(L"{}: EnumeratePlugins returns non-null metaData", relPath).c_str(), success);

    if (hrEnum != S_OK || metaData == nullptr || ! IsValidEnumeratedPluginCount(count))
    {
        return false;
    }

    // Validate each metadata entry
    for (unsigned int i = 0; i < count; ++i)
    {
        const PluginMetaData& md = metaData[i];
        Check(md.id != nullptr && md.id[0] != L'\0', std::format(L"{}: metaData[{}].id is non-empty", relPath, i).c_str(), success);
        Check(md.name != nullptr && md.name[0] != L'\0', std::format(L"{}: metaData[{}].name is non-empty", relPath, i).c_str(), success);
        // version is optional per spec (may be nullptr) — just note its presence
        (void)md.version;
    }

    // Wrong IID must return E_NOINTERFACE and count == 0
    const PluginMetaData* wrongMeta = nullptr;
    unsigned int wrongCount         = 0;
    // Bogus IID that matches neither IFileSystem nor IViewer
    static constexpr BYTE kBogusData4[8] = {0xde, 0xad, 0xde, 0xad, 0xde, 0xad, 0xde, 0xad};
    const IID bogusIid    = {0xdeadbeef,
                             0xdead,
                             0xdead,
                             {kBogusData4[0], kBogusData4[1], kBogusData4[2], kBogusData4[3], kBogusData4[4], kBogusData4[5], kBogusData4[6], kBogusData4[7]}};
    const HRESULT hrWrong = pfnEnumerate(bogusIid, &wrongMeta, &wrongCount);
    Check(hrWrong == E_NOINTERFACE, std::format(L"{}: EnumeratePlugins(wrong IID) returns E_NOINTERFACE", relPath).c_str(), success);
    Check(wrongCount == 0, std::format(L"{}: EnumeratePlugins(wrong IID) returns count == 0", relPath).c_str(), success);

    // Schema validation for each plugin id
    for (unsigned int i = 0; i < count; ++i)
    {
        const wchar_t* pluginId   = metaData[i].id;
        const char* schemaJson    = nullptr;
        const HRESULT hrSchema    = pfnSchema(expectedIid, pluginId, &schemaJson);
        const bool schemaOk       = hrSchema == S_OK;
        const bool schemaNotFound = hrSchema == static_cast<HRESULT>(HRESULT_FROM_WIN32(ERROR_NOT_FOUND));
        Check(schemaOk || schemaNotFound, std::format(L"{}: GetConfigurationSchema[{}] returns S_OK or ERROR_NOT_FOUND", relPath, i).c_str(), success);
        if (schemaOk && schemaJson != nullptr && schemaJson[0] != '\0')
        {
            unique_yyjson_doc doc = ParseJson(schemaJson);
            Check(static_cast<bool>(doc), std::format(L"{}: GetConfigurationSchema[{}] JSON parses successfully", relPath, i).c_str(), success);
            TestSchemaDefaultsRoundTrip(pfnCreate, expectedIid, pluginId, schemaJson, relPath, success);
        }
    }

    return true;
}

void TestViewerSizedRecords(std::wstring_view relPath, bool& success) noexcept
{
    const std::wstring absPath = GetExeDir() + std::wstring(relPath);
    wil::unique_hmodule module(LoadLibraryExW(absPath.c_str(), nullptr, 0));
    Check(static_cast<bool>(module), std::format(L"{}: sized-record test loads DLL", relPath).c_str(), success);
    if (! module)
    {
        return;
    }

    const auto enumerate = reinterpret_cast<PfnEnumeratePlugins>(GetProcAddress(module.get(), "RedSalamanderEnumeratePlugins"));
    const auto create = reinterpret_cast<PfnCreate>(GetProcAddress(module.get(), "RedSalamanderCreate"));
    Check(enumerate != nullptr && create != nullptr,
          std::format(L"{}: sized-record test resolves factory exports", relPath).c_str(),
          success);
    if (enumerate == nullptr || create == nullptr)
    {
        return;
    }

    const PluginMetaData* metadata = nullptr;
    unsigned int count = 0u;
    if (enumerate(__uuidof(IViewer), &metadata, &count) != S_OK || metadata == nullptr || count == 0u)
    {
        Check(false, std::format(L"{}: sized-record test enumerates a viewer", relPath).c_str(), success);
        return;
    }

    FactoryOptions options{};
    wil::com_ptr_nothrow<IViewer> viewer;
    const HRESULT createHr = create(__uuidof(IViewer), &options, &g_nullHost, metadata[0].id, viewer.put_void());
    Check(createHr == S_OK && viewer != nullptr,
          std::format(L"{}: sized-record test creates a viewer", relPath).c_str(),
          success);
    if (FAILED(createHr) || ! viewer)
    {
        return;
    }

    ViewerOpenContext shortOpen{};
    shortOpen.sizeBytes = sizeof(ViewerOpenContext) - 1u;
    Check(viewer->Open(&shortOpen) == E_INVALIDARG,
          std::format(L"{}: Open rejects an undersized ViewerOpenContext", relPath).c_str(),
          success);

    ViewerTheme shortTheme{};
    shortTheme.sizeBytes = sizeof(ViewerTheme) - 1u;
    Check(viewer->SetTheme(nullptr) == E_INVALIDARG && viewer->SetTheme(&shortTheme) == E_INVALIDARG,
          std::format(L"{}: SetTheme rejects null and undersized ViewerTheme records", relPath).c_str(),
          success);

    ViewerTheme currentTheme{};
    currentTheme.sizeBytes = sizeof(currentTheme);
    currentTheme.dpi = USER_DEFAULT_SCREEN_DPI;
    Check(viewer->SetTheme(&currentTheme) == S_OK,
          std::format(L"{}: SetTheme accepts the current ViewerTheme record", relPath).c_str(),
          success);

    struct ExtendedViewerTheme final
    {
        ViewerTheme base{};
        uint64_t unknownTail[4]{};
    };
    ExtendedViewerTheme extendedTheme{};
    extendedTheme.base = currentTheme;
    extendedTheme.base.sizeBytes = sizeof(extendedTheme);
    Check(viewer->SetTheme(&extendedTheme.base) == S_OK,
          std::format(L"{}: SetTheme ignores an unknown ViewerTheme tail", relPath).c_str(),
          success);

    static_cast<void>(viewer->SetCallback(nullptr, nullptr));
    static_cast<void>(viewer->Close());
}

// Defined below alongside the file system provider proofs; the Terminal reuses it verbatim.
void TestTransactionalConfiguration(IInformations& information, std::wstring_view relativePath, const wchar_t* pluginId, bool& success) noexcept;

void TestTerminalSizedRecords(bool& success, bool runTestOnlyContracts = true) noexcept
{
    constexpr std::wstring_view relativePath = L"Plugins\\Terminal.dll";
    const std::wstring absPath = GetExeDir() + std::wstring(relativePath);
    wil::unique_hmodule module(LoadLibraryExW(absPath.c_str(), nullptr, 0));
    Check(static_cast<bool>(module), L"Terminal.dll: sized-record test loads DLL", success);
    if (! module)
    {
        return;
    }

    const auto enumerate = reinterpret_cast<PfnEnumeratePlugins>(GetProcAddress(module.get(), "RedSalamanderEnumeratePlugins"));
    const auto create = reinterpret_cast<PfnCreate>(GetProcAddress(module.get(), "RedSalamanderCreate"));
    if (enumerate == nullptr || create == nullptr)
    {
        Check(false, L"Terminal.dll: sized-record test resolves factory exports", success);
        return;
    }

    const PluginMetaData* metadata = nullptr;
    unsigned int count = 0u;
    if (enumerate(__uuidof(ITerminal), &metadata, &count) != S_OK || metadata == nullptr || count == 0u)
    {
        Check(false, L"Terminal.dll: sized-record test enumerates a terminal", success);
        return;
    }

    FactoryOptions options{};
    wil::com_ptr_nothrow<ITerminal> terminal;
    const HRESULT createHr = create(__uuidof(ITerminal), &options, &g_nullHost, metadata[0].id, terminal.put_void());
    Check(createHr == S_OK && terminal != nullptr, L"Terminal.dll: sized-record test creates a terminal", success);
    if (FAILED(createHr) || ! terminal)
    {
        return;
    }

    TerminalOpenContext shortOpen{};
    shortOpen.sizeBytes = sizeof(TerminalOpenContext) - 1u;
    Check(terminal->Open(&shortOpen) == E_INVALIDARG,
          L"Terminal.dll: Open rejects an undersized TerminalOpenContext",
          success);

    {
        // The Terminal persists a font family and size through the shared plugin-configuration
        // codec, so it owes the same transactional guarantees as the file system providers.
        wil::com_ptr_nothrow<IInformations> information;
        const HRESULT informationHr = terminal->QueryInterface(__uuidof(IInformations), information.put_void());
        Check(informationHr == S_OK && information != nullptr,
              L"Terminal.dll: transactional configuration surface exposes IInformations",
              success);
        if (information)
        {
            TestTransactionalConfiguration(*information.get(), L"Plugins\\Terminal.dll", metadata[0].id, success);
        }
    }

    TerminalTheme shortTheme{};
    shortTheme.sizeBytes = sizeof(TerminalTheme) - 1u;
    Check(terminal->SetTheme(nullptr) == E_INVALIDARG && terminal->SetTheme(&shortTheme) == E_INVALIDARG,
          L"Terminal.dll: SetTheme rejects null and undersized TerminalTheme records",
          success);

    TerminalTheme currentTheme{};
    currentTheme.sizeBytes = sizeof(currentTheme);
    Check(terminal->SetTheme(&currentTheme) == S_OK,
          L"Terminal.dll: SetTheme accepts the current TerminalTheme record",
          success);

    struct ExtendedTerminalTheme final
    {
        TerminalTheme base{};
        uint64_t unknownTail[4]{};
    };
    ExtendedTerminalTheme extendedTheme{};
    extendedTheme.base = currentTheme;
    extendedTheme.base.sizeBytes = sizeof(extendedTheme);
    Check(terminal->SetTheme(&extendedTheme.base) == S_OK,
          L"Terminal.dll: SetTheme ignores an unknown TerminalTheme tail",
          success);

    TerminalViewState shortState{};
    shortState.sizeBytes = sizeof(TerminalViewState) - 1u;
    Check(terminal->GetViewState(&shortState) == E_INVALIDARG,
          L"Terminal.dll: GetViewState rejects an undersized TerminalViewState",
          success);

    TerminalViewState currentState{};
    currentState.sizeBytes = sizeof(currentState);
    const HRESULT currentStateHr = terminal->GetViewState(&currentState);
    wil::unique_cotaskmem_string currentTitle(currentState.title.data);
    wil::unique_cotaskmem_string currentStatus(currentState.status.data);
    Check(currentStateHr == S_OK && currentState.sizeBytes == sizeof(currentState) &&
              currentState.activity.sizeBytes == sizeof(currentState.activity),
          L"Terminal.dll: GetViewState returns current nested sizeBytes values",
          success);

    struct ExtendedTerminalViewState final
    {
        TerminalViewState base{};
        uint64_t unknownTail[4]{0xA55AA55AA55AA55Aull, 0xA55AA55AA55AA55Aull, 0xA55AA55AA55AA55Aull, 0xA55AA55AA55AA55Aull};
    };
    ExtendedTerminalViewState extendedState{};
    extendedState.base.sizeBytes = sizeof(extendedState);
    const HRESULT extendedStateHr = terminal->GetViewState(&extendedState.base);
    wil::unique_cotaskmem_string extendedTitle(extendedState.base.title.data);
    wil::unique_cotaskmem_string extendedStatus(extendedState.base.status.data);
    const bool tailPreserved = std::ranges::all_of(extendedState.unknownTail, [](uint64_t value) noexcept {
        return value == 0xA55AA55AA55AA55Aull;
    });
    Check(extendedStateHr == S_OK && tailPreserved,
          L"Terminal.dll: GetViewState accepts and preserves an unknown caller tail",
          success);

    wil::com_ptr_nothrow<ITerminalActions> actions;
    const HRESULT actionsHr = terminal->QueryInterface(__uuidof(ITerminalActions), actions.put_void());
    Check(actionsHr == S_OK && actions != nullptr,
          L"Terminal.dll: ITerminal exposes the release-lockstep ITerminalActions capability",
          success);
    if (actions)
    {
        constexpr std::wstring_view unknownCommand = L"cmd/app/contractProbe";
        TerminalShortcutRequest shortcut{};
        shortcut.sizeBytes = sizeof(shortcut);
        shortcut.commandId = {unknownCommand.data(), static_cast<uint32_t>(unknownCommand.size())};
        shortcut.message = WM_KEYDOWN;
        shortcut.virtualKey = 'P';
        shortcut.scanCode = 0x19u;
        shortcut.repeatCount = 1u;
        shortcut.instanceId = currentState.instanceId;
        shortcut.sessionGeneration = currentState.sessionGeneration;
        TerminalShortcutRoute route = TerminalShortcutRoute::Blocked;
        TerminalShortcutRequest shortShortcut = shortcut;
        shortShortcut.sizeBytes = sizeof(shortShortcut) - 1u;
        Check(actions->RouteShortcut(&shortShortcut, &route) == E_INVALIDARG && route == TerminalShortcutRoute::PassThrough,
              L"Terminal.dll: RouteShortcut rejects undersized input and initializes fail-open output",
              success);
        route = TerminalShortcutRoute::Blocked;
        Check(actions->RouteShortcut(&shortcut, &route) == S_OK && route == TerminalShortcutRoute::InvokeHostCommand,
              L"Terminal.dll: an eligible non-plugin command routes back to the host",
              success);
        shortcut.modifierFlags = TerminalShortcutModifierCtrl | TerminalShortcutModifierAlt | TerminalShortcutModifierRightAlt;
        route = TerminalShortcutRoute::Blocked;
        Check(actions->RouteShortcut(&shortcut, &route) == S_OK && route == TerminalShortcutRoute::PassThrough,
              L"Terminal.dll: printable AltGr input passes through before host-command arbitration",
              success);

        constexpr std::wstring_view findCommand = L"cmd/terminal/find";
        TerminalActionRequest action{};
        action.sizeBytes = sizeof(action);
        action.commandId = {findCommand.data(), static_cast<uint32_t>(findCommand.size())};
        action.instanceId = currentState.instanceId;
        action.sessionGeneration = currentState.sessionGeneration;
        TerminalActionState actionState{};
        actionState.sizeBytes = sizeof(actionState);
        Check(actions->GetActionState(&action, nullptr) == E_POINTER,
              L"Terminal.dll: GetActionState rejects a null output record",
              success);

        struct ActionStateWithCanary final
        {
            TerminalActionState base{};
            std::array<uint64_t, 4u> tail{};
        };
        ActionStateWithCanary undersizedState{};
        std::memset(&undersizedState, 0xA5, sizeof(undersizedState));
        undersizedState.base.sizeBytes = sizeof(TerminalActionState) - 1u;
        const ActionStateWithCanary undersizedBefore = undersizedState;
        Check(actions->GetActionState(&action, &undersizedState.base) == E_INVALIDARG &&
                  std::memcmp(&undersizedState, &undersizedBefore, sizeof(undersizedState)) == 0,
              L"Terminal.dll: GetActionState rejects an undersized output without touching its canary",
              success);

        Check(actions->GetActionState(&action, &actionState) == S_OK && actionState.enabled == 0u &&
                  actions->ExecuteAction(&action) == E_NOTIMPL,
              L"Terminal.dll: unavailable Find is disabled and yields instead of consuming input",
              success);

        ActionStateWithCanary extendedActionState{};
        extendedActionState.base.sizeBytes = sizeof(extendedActionState);
        extendedActionState.tail.fill(0xA55AA55AA55AA55Aull);
        Check(actions->GetActionState(&action, &extendedActionState.base) == S_OK &&
                  std::ranges::all_of(extendedActionState.tail, [](uint64_t value) noexcept {
                      return value == 0xA55AA55AA55AA55Aull;
                  }),
              L"Terminal.dll: GetActionState accepts and preserves an unknown caller tail",
              success);

        action.commandId = {unknownCommand.data(), static_cast<uint32_t>(unknownCommand.size())};
        Check(actions->GetActionState(&action, &actionState) == E_NOTIMPL && actions->ExecuteAction(&action) == E_NOTIMPL,
              L"Terminal.dll: unknown direct actions fail open with E_NOTIMPL",
              success);

        constexpr std::wstring_view copyOrPassthroughCommand = L"cmd/terminal/copySelectionOrPassthrough";
        action.commandId = {copyOrPassthroughCommand.data(), static_cast<uint32_t>(copyOrPassthroughCommand.size())};
        Check(actions->GetActionState(&action, &actionState) == S_OK && actionState.enabled == 0u &&
                  actions->ExecuteAction(&action) == S_FALSE,
              L"Terminal.dll: copySelectionOrPassthrough is a plugin action that is disabled without a selection and does not pass through from ExecuteAction",
              success);
    }

    struct CallbackProbe final : ITerminalEventCallback
    {
        void STDMETHODCALLTYPE OnTerminalEvent(const TerminalEvent* /*event*/, void* /*cookie*/) noexcept override {}
    } callbackProbe;
    int callbackCookie = 0;
    Check(terminal->SetCallback(nullptr, &callbackCookie) == E_INVALIDARG &&
              terminal->SetCallback(&callbackProbe, nullptr) == E_INVALIDARG &&
              terminal->SetCallback(&callbackProbe, &callbackCookie) == S_OK &&
              terminal->SetCallback(nullptr, nullptr) == S_OK,
          L"Terminal.dll: weak event callback registration accepts only both-null or both-non-null pairs",
          success);

    if (actions)
    {
        constexpr std::wstring_view findCommand = L"cmd/terminal/find";
        TerminalActionRequest findAction{};
        findAction.sizeBytes = sizeof(findAction);
        findAction.commandId = {findCommand.data(), static_cast<uint32_t>(findCommand.size())};
        TerminalViewState liveState{};
        liveState.sizeBytes = sizeof(liveState);
        if (SUCCEEDED(terminal->GetViewState(&liveState)))
        {
            CoTaskMemFree(liveState.title.data);
            CoTaskMemFree(liveState.status.data);
            findAction.instanceId = liveState.instanceId;
            findAction.sessionGeneration = liveState.sessionGeneration;
        }
        static_cast<void>(actions->ExecuteAction(&findAction));
    }

    Check(SUCCEEDED(terminal->Close()), L"Terminal.dll: Close succeeds", success);
    if (! runTestOnlyContracts)
    {
        return;
    }

    using PfnJoinTid = DWORD(__stdcall*)(ITerminal*);
    const auto joinTid = reinterpret_cast<PfnJoinTid>(GetProcAddress(module.get(), "RedSalamanderTerminalDebugCommandSurfaceJoinTid"));
    const DWORD closeCallerTid = GetCurrentThreadId();
    DWORD observedJoinTid = 0u;
    for (int attempt = 0; attempt < 50; ++attempt)
    {
        observedJoinTid = joinTid != nullptr ? joinTid(terminal.get()) : 0u;
        if (observedJoinTid != 0u)
        {
            break;
        }
        Sleep(20);
    }
    Check(joinTid != nullptr && observedJoinTid != 0u && observedJoinTid != closeCallerTid,
          L"Terminal.dll: Close does not join the Find/Suggestions worker on the calling UI thread",
          success);
}

// ---------------------------------------------------------------------------
// Step 3: GetCapabilities pass (filesystem plugins only)
// ---------------------------------------------------------------------------
[[nodiscard]] bool RequiresTransactionalConfigurationProof(std::wstring_view relativePath) noexcept
{
    constexpr std::array<std::wstring_view, 6> providers = {
        L"Plugins\\FileSystem.dll",
        L"Plugins\\FileSystem7z.dll",
        L"Plugins\\FileSystemCurl.dll",
        L"Plugins\\FileSystemGoogleDrive.dll",
        L"Plugins\\FileSystemMicrosoftDrive.dll",
        L"Plugins\\FileSystemS3.dll",
    };
    return std::ranges::find(providers, relativePath) != providers.end();
}

void TestTransactionalConfiguration(IInformations& information, std::wstring_view relativePath, const wchar_t* pluginId, bool& success) noexcept
{
    constexpr char kForwardConfiguration[] = R"json({"observatoryUnknown":{"value":17}})json";
    const HRESULT forwardHr = information.SetConfiguration(kForwardConfiguration);
    const char* configuration = nullptr;
    const HRESULT getForwardHr = information.GetConfiguration(&configuration);
    const std::string preserved = configuration != nullptr ? configuration : "";
    Check(forwardHr == S_OK && getForwardHr == S_OK && preserved.find("\"observatoryUnknown\"") != std::string::npos,
          std::format(L"{}: SetConfiguration(pluginId={}) preserves unknown members", relativePath, pluginId).c_str(),
          success);

    const HRESULT malformedHr = information.SetConfiguration("{");
    configuration = nullptr;
    const HRESULT getAfterMalformedHr = information.GetConfiguration(&configuration);
    Check(malformedHr == HRESULT_FROM_WIN32(ERROR_INVALID_DATA) && getAfterMalformedHr == S_OK && configuration != nullptr &&
              std::string_view(configuration) == preserved,
          std::format(L"{}: malformed configuration preserves live state for pluginId={}", relativePath, pluginId).c_str(),
          success);

    const HRESULT wrongRootHr = information.SetConfiguration("[]");
    configuration = nullptr;
    const HRESULT getAfterWrongRootHr = information.GetConfiguration(&configuration);
    Check(wrongRootHr == HRESULT_FROM_WIN32(ERROR_INVALID_DATA) && getAfterWrongRootHr == S_OK && configuration != nullptr &&
              std::string_view(configuration) == preserved,
          std::format(L"{}: wrong-root configuration preserves live state for pluginId={}", relativePath, pluginId).c_str(),
          success);

    if (relativePath == L"Plugins\\FileSystem7z.dll" || relativePath == L"Plugins\\FileSystemCurl.dll")
    {
        constexpr char kLegacyPasswordConfiguration[] =
            R"json({"defaultPassword":"observatory-password-sentinel","observatoryUnknown":17})json";
        constexpr char kLegacyCurlSecretConfiguration[] =
            R"json({"defaultPassword":"observatory-password-sentinel","sshKeyPassphrase":"observatory-passphrase-sentinel","observatoryUnknown":17})json";
        const char* legacyConfiguration = relativePath == L"Plugins\\FileSystemCurl.dll" ? kLegacyCurlSecretConfiguration : kLegacyPasswordConfiguration;
        const HRESULT secretHr          = information.SetConfiguration(legacyConfiguration);
        configuration          = nullptr;
        const HRESULT getSecretHr = information.GetConfiguration(&configuration);
        const std::string_view sanitized = configuration != nullptr ? configuration : "";
        Check(secretHr == S_OK && getSecretHr == S_OK && sanitized.find("observatoryUnknown") != std::string_view::npos &&
                  sanitized.find("observatory-password-sentinel") == std::string_view::npos &&
                  (relativePath != L"Plugins\\FileSystemCurl.dll" || sanitized.find("observatory-passphrase-sentinel") == std::string_view::npos),
              std::format(L"{}: legacy configuration secrets are imported but not persisted for pluginId={}", relativePath, pluginId).c_str(),
              success);
    }

    Check(information.SetConfiguration("{}") == S_OK,
          std::format(L"{}: transactional configuration proof restores defaults for pluginId={}", relativePath, pluginId).c_str(),
          success);
}

[[nodiscard]] bool CreateEmptyProviderFile(IFileSystemIO& io, const std::wstring& path) noexcept
{
    wil::com_ptr_nothrow<IFileWriter> writer;
    return io.CreateFileWriter(path.c_str(), FILESYSTEM_FLAG_NONE, writer.put()) == S_OK && writer && writer->Commit() == S_OK;
}

[[nodiscard]] bool ProviderFileExists(IFileSystemIO& io, const std::wstring& path) noexcept
{
    wil::com_ptr_nothrow<IFileReader> reader;
    return io.CreateFileReader(path.c_str(), reader.put()) == S_OK && reader != nullptr;
}

[[nodiscard]] bool SetProviderReadOnly(IFileSystemIO& io, const std::wstring& path) noexcept
{
    FileSystemBasicInformation information{};
    information.sizeBytes = sizeof(information);
    if (FAILED(io.GetFileBasicInformation(path.c_str(), &information)))
    {
        return false;
    }
    information.attributes |= FILE_ATTRIBUTE_READONLY;
    return io.SetFileBasicInformation(path.c_str(), &information) == S_OK;
}

void TestLocalWriterFlagContract(IFileSystem& fileSystem, bool& success) noexcept
{
    wil::com_ptr_nothrow<IFileSystemIO> io;
    if (FAILED(fileSystem.QueryInterface(__uuidof(IFileSystemIO), io.put_void())) || ! io)
    {
        Check(false, L"local provider exposes IFileSystemIO for writer flag proof", success);
        return;
    }

    std::error_code error;
    const std::filesystem::path root = RedSalamander::TestSupport::AcquireTestDirectory(
        {.harnessSegment     = L"PluginContractTests",
         .leafSegment        = L"observatory-track13-local-writer",
         .fallbackRunIdPrefix = L"plugin-contract",
         .kind                = RedSalamander::TestSupport::TestDirectoryKind::Scratch},
        error);
    Check(! error && ! root.empty(), L"local writer flag proof acquires TestSandbox scratch", success);
    if (error || root.empty())
    {
        return;
    }

    const std::filesystem::path target = root / L"readonly-existing.txt";
    wil::unique_handle seed(::CreateFileW(target.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_NEW, FILE_ATTRIBUTE_NORMAL, nullptr));
    const bool seeded = seed != nullptr;
    seed.reset();
    const bool madeReadOnly = seeded && ::SetFileAttributesW(target.c_str(), FILE_ATTRIBUTE_READONLY) != 0;
    Check(madeReadOnly, L"local writer flag proof seeds a read-only destination", success);

    wil::com_ptr_nothrow<IFileWriter> writer;
    const HRESULT createHr = madeReadOnly
                                 ? io->CreateFileWriter(target.c_str(), FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY, writer.put())
                                 : E_UNEXPECTED;
    const DWORD attributes = ::GetFileAttributesW(target.c_str());
    Check(createHr == E_INVALIDARG && ! writer && attributes != INVALID_FILE_ATTRIBUTES && (attributes & FILE_ATTRIBUTE_READONLY) != 0u,
          L"local writer rejects replace-readonly without overwrite and preserves destination attributes",
          success);

    if (attributes != INVALID_FILE_ATTRIBUTES)
    {
        static_cast<void>(::SetFileAttributesW(target.c_str(), attributes & ~FILE_ATTRIBUTE_READONLY));
    }
    std::filesystem::remove_all(root, error);
}

void TestDummyTransactionalMutationContracts(IFileSystem& fileSystem, bool& success) noexcept
{
    wil::com_ptr_nothrow<IFileSystemIO> io;
    wil::com_ptr_nothrow<IFileSystemDirectoryOperations> directoryOperations;
    if (FAILED(fileSystem.QueryInterface(__uuidof(IFileSystemIO), io.put_void())) || ! io ||
        FAILED(fileSystem.QueryInterface(__uuidof(IFileSystemDirectoryOperations), directoryOperations.put_void())) || ! directoryOperations)
    {
        Check(false, L"dummy provider exposes mutation interfaces for Track 13 proof", success);
        return;
    }

    const std::wstring root = std::format(L"/observatory-track13-{}", GetTickCount64());
    const auto makeDirectory = [&](std::wstring_view suffix) noexcept
    {
        return directoryOperations->CreateDirectory((root + std::wstring(suffix)).c_str()) == S_OK;
    };
    const auto makeFile = [&](std::wstring_view suffix) noexcept
    {
        return CreateEmptyProviderFile(*io.get(), root + std::wstring(suffix));
    };

    const bool copySeeded = makeDirectory(L"") && makeDirectory(L"/copy-source") && makeDirectory(L"/copy-destination") &&
                            makeFile(L"/copy-source/first.txt") && makeFile(L"/copy-source/late.txt") &&
                            makeFile(L"/copy-destination/late.txt");
    Check(copySeeded, L"dummy copy rollback proof seeds source and late destination collision", success);
    const HRESULT copyHr = copySeeded
                               ? fileSystem.CopyItem((root + L"/copy-source").c_str(),
                                                     (root + L"/copy-destination").c_str(),
                                                     FILESYSTEM_FLAG_RECURSIVE,
                                                     nullptr,
                                                     nullptr,
                                                     nullptr)
                               : E_UNEXPECTED;
    Check(copyHr == HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS) && ! ProviderFileExists(*io.get(), root + L"/copy-destination/first.txt"),
          L"dummy directory copy preflights a late collision without partial destination mutation",
          success);

    const bool moveSeeded = makeDirectory(L"/move-source") && makeDirectory(L"/move-destination") && makeFile(L"/move-source/first.txt") &&
                            makeFile(L"/move-source/late.txt") && makeFile(L"/move-destination/late.txt");
    Check(moveSeeded, L"dummy move rollback proof seeds source and late destination collision", success);
    const HRESULT moveHr = moveSeeded
                               ? fileSystem.MoveItem((root + L"/move-source").c_str(),
                                                     (root + L"/move-destination").c_str(),
                                                     FILESYSTEM_FLAG_RECURSIVE,
                                                     nullptr,
                                                     nullptr,
                                                     nullptr)
                               : E_UNEXPECTED;
    Check(moveHr == HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS) && ProviderFileExists(*io.get(), root + L"/move-source/first.txt") &&
              ! ProviderFileExists(*io.get(), root + L"/move-destination/first.txt"),
          L"dummy directory move preflights a late collision without partial source or destination mutation",
          success);

    const bool deleteSeeded = makeDirectory(L"/delete-source") && makeFile(L"/delete-source/readonly-child.txt") &&
                              SetProviderReadOnly(*io.get(), root + L"/delete-source/readonly-child.txt");
    Check(deleteSeeded, L"dummy recursive delete proof seeds a read-only descendant", success);
    const HRESULT deleteHr = deleteSeeded
                                 ? fileSystem.DeleteItem(
                                       (root + L"/delete-source").c_str(), FILESYSTEM_FLAG_RECURSIVE, nullptr, nullptr, nullptr)
                                 : E_UNEXPECTED;
    Check(deleteHr == HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED) && ProviderFileExists(*io.get(), root + L"/delete-source/readonly-child.txt"),
          L"dummy recursive delete applies read-only policy to descendants before mutation",
          success);

    const FileSystemFlags cleanupFlags =
        static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);
    Check(fileSystem.DeleteItem(root.c_str(), cleanupFlags, nullptr, nullptr, nullptr) == S_OK,
          L"dummy Track 13 proof cleans its provider fixture",
          success);
}

void TestTrack13ProviderContracts(IFileSystem& fileSystem, std::wstring_view relativePath, bool& success) noexcept
{
    if (relativePath == L"Plugins\\FileSystem.dll")
    {
        TestLocalWriterFlagContract(fileSystem, success);
    }
    else if (relativePath == L"Plugins\\FileSystemDummy.dll")
    {
        TestDummyTransactionalMutationContracts(fileSystem, success);
    }
}

void SkipTrack13ProviderContracts(std::wstring_view relativePath) noexcept
{
    if (relativePath != L"Plugins\\FileSystem.dll" && relativePath != L"Plugins\\FileSystemDummy.dll")
    {
        return;
    }
    std::wcout << L"[  SKIPPED ] " << relativePath
               << L": provider mutation proofs need a TestSandbox scratch root that a packaged extraction does not have\n";
}

void TestTypedRouteValidatorContracts(bool& success) noexcept
{
    {
        ScriptedRouteCapabilities route(ScriptedRouteFault::None);
        const FileSystemRouteContract::QueryResult result =
            FileSystemRouteContract::Query(&route, L"/", FILESYSTEM_COPY, L"selftest/typed-route");
        Check(result.state == FileSystemRouteContract::QueryState::Available && result.status == S_OK &&
                  ! result.usedArenaFallback && result.snapshot.pathIdentity.has_value(),
              L"typed route validator accepts a current complete fact set",
              success);
        const FileSystemRouteContract::QueryResult mismatch =
            FileSystemRouteContract::Query(&route, L"/", FILESYSTEM_COPY, L"selftest/other-provider");
        Check(mismatch.state == FileSystemRouteContract::QueryState::ContractViolation,
              L"typed route validator rejects a mismatched full provider ID",
              success);
    }
    {
        ScriptedRouteCapabilities route(ScriptedRouteFault::OversizedPrefix);
        const FileSystemRouteContract::QueryResult result =
            FileSystemRouteContract::Query(&route, L"/", FILESYSTEM_COPY, L"selftest/typed-route");
        Check(result.state == FileSystemRouteContract::QueryState::Available,
              L"typed route consumer accepts a provider record with a valid current prefix",
              success);
    }
    {
        ScriptedRouteCapabilities route(ScriptedRouteFault::UnsupportedIdentity);
        const FileSystemRouteContract::QueryResult result =
            FileSystemRouteContract::Query(&route, L"/", FILESYSTEM_COPY, L"selftest/typed-route");
        Check(result.state == FileSystemRouteContract::QueryState::Unsupported,
              L"typed route validator classifies an uncomparable path identity as Unsupported",
              success);
    }
    {
        ScriptedRouteCapabilities route(ScriptedRouteFault::RequiresArenaFallback);
        const FileSystemRouteContract::QueryResult result =
            FileSystemRouteContract::Query(&route, L"/", FILESYSTEM_COPY, L"selftest/typed-route");
        const std::wstring copiedProfile = result.snapshot.pathProfileId;
        const FileSystemRouteContract::QueryResult second =
            FileSystemRouteContract::Query(&route, L"/", FILESYSTEM_COPY, L"selftest/typed-route");
        Check(result.state == FileSystemRouteContract::QueryState::Available && result.usedArenaFallback &&
                  copiedProfile.size() > FileSystemRouteContract::kNormalArenaBytes / sizeof(wchar_t) &&
                  result.snapshot.pathProfileId == copiedProfile && second.snapshot.pathProfileId == copiedProfile,
              L"typed route validator performs one bounded arena retry and copies returned strings immediately",
              success);
#ifdef ENABLE_TESTS
        const FileSystemRouteContract::QueryResult allocationFailure =
            FileSystemRouteContract::QueryWithAllocationFailureForSelfTest(
                &route, L"/", FILESYSTEM_COPY, L"selftest/typed-route");
        Check(allocationFailure.state == FileSystemRouteContract::QueryState::ContractViolation &&
                  allocationFailure.status == E_OUTOFMEMORY,
              L"typed route validator fails closed on deterministic fallback allocation failure",
              success);
#endif
    }

    constexpr std::array malformedFaults{
        ScriptedRouteFault::InvalidRecordSize,
        ScriptedRouteFault::InvalidAvailability,
        ScriptedRouteFault::InvalidCancellation,
        ScriptedRouteFault::InvalidNamespace,
        ScriptedRouteFault::InvalidComparison,
        ScriptedRouteFault::InvalidNormalization,
        ScriptedRouteFault::InvalidCaseRename,
        ScriptedRouteFault::InvalidProofFlags,
        ScriptedRouteFault::NonStrictBoolean,
        ScriptedRouteFault::WatchdogContradiction,
        ScriptedRouteFault::EmptyProviderId,
        ScriptedRouteFault::OutsideArenaPointer,
        ScriptedRouteFault::MisalignedArenaPointer,
        ScriptedRouteFault::RequiredArenaOverflow,
    };
    for (const ScriptedRouteFault fault : malformedFaults)
    {
        ScriptedRouteCapabilities route(fault);
        const FileSystemRouteContract::QueryResult result =
            FileSystemRouteContract::Query(&route, L"/", FILESYSTEM_COPY, L"selftest/typed-route");
        Check(result.state == FileSystemRouteContract::QueryState::ContractViolation,
              std::format(L"typed route validator rejects malformed fact fault {}", static_cast<unsigned int>(fault)).c_str(),
              success);
    }
    {
        ScriptedRouteCapabilities route(ScriptedRouteFault::NonStrictPeer);
        const FileSystemRouteContract::BooleanResult peer = FileSystemRouteContract::QueryTransferPeerAllowed(
            &route, L"/", FILESYSTEM_COPY, FILESYSTEM_TRANSFER_PEER_EXPORT, L"selftest/peer");
        Check(peer.state == FileSystemRouteContract::QueryState::ContractViolation,
              L"typed route validator rejects a non-strict peer BOOL",
              success);
    }
    {
        ScriptedRouteCapabilities route(ScriptedRouteFault::PeerDenied);
        const FileSystemRouteContract::BooleanResult denied = FileSystemRouteContract::QueryTransferPeerAllowed(
            &route, L"/", FILESYSTEM_COPY, FILESYSTEM_TRANSFER_PEER_EXPORT, L"selftest/peer");
        const FileSystemRouteContract::BooleanResult invalidRole = FileSystemRouteContract::QueryTransferPeerAllowed(
            &route, L"/", FILESYSTEM_COPY, static_cast<FileSystemTransferPeerRole>(99u), L"selftest/peer");
        ScriptedRouteCapabilities failedRoute(ScriptedRouteFault::PeerFailure);
        const FileSystemRouteContract::BooleanResult failed = FileSystemRouteContract::QueryTransferPeerAllowed(
            &failedRoute, L"/", FILESYSTEM_COPY, FILESYSTEM_TRANSFER_PEER_EXPORT, L"selftest/peer");
        Check(denied.state == FileSystemRouteContract::QueryState::Available && ! denied.value &&
                  invalidRole.state == FileSystemRouteContract::QueryState::ContractViolation &&
                  failed.state == FileSystemRouteContract::QueryState::ContractViolation,
              L"typed route validator preserves peer deny and rejects invalid-role or failed peer output",
              success);
    }
    {
        ScriptedRouteCapabilities route(ScriptedRouteFault::InvalidNameResult);
        const FileSystemRouteContract::ChildNameResult name =
            FileSystemRouteContract::ValidateChildName(&route, L"/", L"child.txt", FILESYSTEM_RENAME);
        Check(name.state == FileSystemRouteContract::QueryState::ContractViolation,
              L"typed route validator rejects an invalid child-name result enum",
              success);
    }
    {
        ScriptedRouteCapabilities route(ScriptedRouteFault::CollisionOutsideArena);
        const FileSystemRouteContract::StringResult collision = FileSystemRouteContract::QueryChildNameCollisionKey(
            &route, L"/", L"child.txt", FILESYSTEM_RENAME);
        ScriptedRouteCapabilities joinRoute(ScriptedRouteFault::JoinRequiredMismatch);
        const FileSystemRouteContract::StringResult joined =
            FileSystemRouteContract::QueryJoinedPath(&joinRoute, L"/", L"child.txt", FILESYSTEM_RENAME);
        Check(collision.state == FileSystemRouteContract::QueryState::ContractViolation &&
                  joined.state == FileSystemRouteContract::QueryState::ContractViolation,
              L"typed route validator rejects malformed collision-key and joined-path output",
              success);
    }
    {
        ScriptedRouteCapabilities route(ScriptedRouteFault::None);
        const FileSystemRouteContract::ChildNameContractResult name = FileSystemRouteContract::QueryChildNameContract(
            &route, L"/parent", L"CON", FILESYSTEM_RENAME, L"selftest/typed-route");
        Check(name.state == FileSystemRouteContract::QueryState::Available && name.status == S_OK &&
                  name.nameStatus == FILESYSTEM_CHILD_NAME_VALID && name.failureStatus == S_OK &&
                  name.joinedPath == L"/parent/CON" && name.collisionKey == L"CON" && name.arenaFallbackCount == 0u,
              L"composite child-name contract accepts provider-valid names without applying Windows reserved-name rules",
              success);
    }
    {
        ScriptedRouteCapabilities route(ScriptedRouteFault::ProviderInvalidName);
        const FileSystemRouteContract::ChildNameContractResult name = FileSystemRouteContract::QueryChildNameContract(
            &route, L"/parent", L"ordinary.txt", FILESYSTEM_RENAME, L"selftest/typed-route");
        Check(name.state == FileSystemRouteContract::QueryState::Available &&
                  name.nameStatus == FILESYSTEM_CHILD_NAME_INVALID &&
                  name.failureStatus == HRESULT_FROM_WIN32(ERROR_INVALID_NAME) && name.joinedPath.empty() && name.collisionKey.empty(),
              L"composite child-name contract preserves provider Invalid HRESULT and returns no executable path/key",
              success);
    }
    {
        ScriptedRouteCapabilities unsupportedRoute(ScriptedRouteFault::UnsupportedName);
        const FileSystemRouteContract::ChildNameContractResult unsupported = FileSystemRouteContract::QueryChildNameContract(
            &unsupportedRoute, L"/parent", L"child.txt", FILESYSTEM_RENAME, L"selftest/typed-route");
        ScriptedRouteCapabilities joinRoute(ScriptedRouteFault::JoinLeafMismatch);
        const FileSystemRouteContract::ChildNameContractResult mismatchedJoin = FileSystemRouteContract::QueryChildNameContract(
            &joinRoute, L"/parent", L"child.txt", FILESYSTEM_RENAME, L"selftest/typed-route");
        ScriptedRouteCapabilities emptyKeyRoute(ScriptedRouteFault::EmptyCollisionKey);
        const FileSystemRouteContract::ChildNameContractResult emptyKey = FileSystemRouteContract::QueryChildNameContract(
            &emptyKeyRoute, L"/parent", L"child.txt", FILESYSTEM_RENAME, L"selftest/typed-route");
        Check(unsupported.state == FileSystemRouteContract::QueryState::Unsupported &&
                  mismatchedJoin.state == FileSystemRouteContract::QueryState::ContractViolation &&
                  emptyKey.state == FileSystemRouteContract::QueryState::ContractViolation,
              L"composite child-name contract fails closed on unsupported, mismatched-join, and empty-key output",
              success);
    }
    {
        ScriptedRouteCapabilities route(ScriptedRouteFault::ChildStringArenaFallback);
        const FileSystemRouteContract::ChildNameContractResult name = FileSystemRouteContract::QueryChildNameContract(
            &route, L"/parent", L"child.txt", FILESYSTEM_RENAME, L"selftest/typed-route");
        Check(name.state == FileSystemRouteContract::QueryState::Available && name.arenaFallbackCount == 1u &&
                  name.collisionKey.size() > FileSystemRouteContract::kNormalArenaBytes / sizeof(wchar_t),
              L"composite child-name contract performs one bounded string-arena fallback and copies the canonical key",
              success);
#ifdef ENABLE_TESTS
        const FileSystemRouteContract::ChildNameContractResult allocationFailure =
            FileSystemRouteContract::QueryChildNameContractWithAllocationFailureForSelfTest(
                &route, L"/parent", L"child.txt", FILESYSTEM_RENAME, L"selftest/typed-route");
        Check(allocationFailure.state == FileSystemRouteContract::QueryState::ContractViolation &&
                  allocationFailure.status == E_OUTOFMEMORY,
              L"composite child-name contract fails closed on deterministic string-arena allocation failure",
              success);
#endif
    }

    FileSystemItemMutationResult retryable{
        sizeof(FileSystemItemMutationResult), TRUE, FALSE, TRUE, FileSystemOwnedStageDisposition::Removed};
    FileSystemItemMutationResult committed{
        sizeof(FileSystemItemMutationResult), TRUE, TRUE, FALSE, FileSystemOwnedStageDisposition::Published};
    FileSystemItemMutationResult unknown{
        sizeof(FileSystemItemMutationResult), FALSE, FALSE, TRUE, FileSystemOwnedStageDisposition::Unknown};
    FileSystemItemMutationResult retained{
        sizeof(FileSystemItemMutationResult), TRUE, FALSE, TRUE, FileSystemOwnedStageDisposition::Retained};
    FileSystemItemMutationResult malformed = retryable;
    malformed.outcomeKnown = 2;
    Check(FileSystemRouteContract::ClassifyFailedMutation(false, E_FAIL, &retryable) ==
              FileSystemRouteContract::MutationClassification::Unsupported &&
              FileSystemRouteContract::ClassifyFailedMutation(true, E_FAIL, nullptr) ==
                  FileSystemRouteContract::MutationClassification::Indeterminate &&
              FileSystemRouteContract::ClassifyFailedMutation(true, E_FAIL, &retryable) ==
                  FileSystemRouteContract::MutationClassification::RetryableNoCommit &&
              FileSystemRouteContract::ClassifyFailedMutation(true, E_FAIL, &committed) ==
                  FileSystemRouteContract::MutationClassification::FailedKnown &&
              FileSystemRouteContract::ClassifyFailedMutation(true, E_FAIL, &unknown) ==
                  FileSystemRouteContract::MutationClassification::Indeterminate &&
              FileSystemRouteContract::ClassifyFailedMutation(true, E_FAIL, &retained) ==
                  FileSystemRouteContract::MutationClassification::FailedKnown &&
              FileSystemRouteContract::ClassifyFailedMutation(true, E_FAIL, &malformed) ==
                  FileSystemRouteContract::MutationClassification::ContractViolation,
          L"typed mutation classifier preserves unsupported, retryable, known, indeterminate, and contract-violation states",
          success);
}

bool TestCapabilities(std::wstring_view relPath, bool runProviderMutationProofs, bool& success) noexcept
{
    const std::wstring exeDir  = GetExeDir();
    const std::wstring absPath = exeDir + std::wstring(relPath);

    // Module must stay alive while we hold interface pointers from it.
    wil::unique_hmodule mod(LoadLibraryExW(absPath.c_str(), nullptr, 0));
    if (! mod)
    {
        // Already reported in enumerate pass; don't double-report.
        return false;
    }

    const auto pfnEnumerate = reinterpret_cast<PfnEnumeratePlugins>(GetProcAddress(mod.get(), "RedSalamanderEnumeratePlugins"));
    const auto pfnCreate    = reinterpret_cast<PfnCreate>(GetProcAddress(mod.get(), "RedSalamanderCreate"));

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

        void* raw              = nullptr;
        const HRESULT hrCreate = pfnCreate(__uuidof(IFileSystem), nullptr, static_cast<IHost*>(&g_nullHost), pluginId, &raw);
        Check(hrCreate == S_OK, std::format(L"{}: Create(pluginId={}) returns S_OK", relPath, pluginId).c_str(), success);
        Check(raw != nullptr, std::format(L"{}: Create(pluginId={}) returns non-null instance", relPath, pluginId).c_str(), success);

        if (hrCreate != S_OK || raw == nullptr)
        {
            continue;
        }

        // Wrap in a smart pointer — IFileSystem inherits IUnknown
        wil::com_ptr_nothrow<IFileSystem> fs;
        fs.attach(static_cast<IFileSystem*>(raw));

        wil::com_ptr_nothrow<IFileSystemRouteCapabilities> routeCapabilities;
        const HRESULT routeCapabilitiesQi =
            fs->QueryInterface(__uuidof(IFileSystemRouteCapabilities), routeCapabilities.put_void());
        Check(routeCapabilitiesQi == S_OK && routeCapabilities != nullptr,
              std::format(L"{}: pluginId={} exposes mandatory IFileSystemRouteCapabilities", relPath, pluginId).c_str(),
              success);
        if (routeCapabilities)
        {
            alignas(wchar_t) std::array<unsigned char, FileSystemRouteContract::kNormalArenaBytes> routeStorage{};
            FileSystemArena routeArena{routeStorage.data(), static_cast<unsigned long>(routeStorage.size()), 0u};
            FileSystemRouteFacts routeFacts{};
            routeFacts.sizeBytes = sizeof(routeFacts) - 1u;
            Check(routeCapabilities->GetRouteFacts(L"/", FILESYSTEM_COPY, &routeArena, &routeFacts) == E_INVALIDARG,
                  std::format(L"{}: pluginId={} rejects an undersized typed route record", relPath, pluginId).c_str(),
                  success);
            routeFacts = {};
            routeFacts.sizeBytes = sizeof(routeFacts) + 1u;
            Check(routeCapabilities->GetRouteFacts(L"/", FILESYSTEM_COPY, &routeArena, &routeFacts) == E_INVALIDARG,
                  std::format(L"{}: pluginId={} rejects a caller record from a different ABI generation", relPath, pluginId).c_str(),
                  success);
            routeFacts = {};
            routeFacts.sizeBytes = sizeof(routeFacts);
            routeArena.usedBytes = 0u;
            const HRESULT currentFactsHr =
                routeCapabilities->GetRouteFacts(L"/", FILESYSTEM_COPY, &routeArena, &routeFacts);
            Check(currentFactsHr == S_OK && routeFacts.sizeBytes == sizeof(routeFacts) &&
                      routeFacts.requiredArenaBytes == routeArena.usedBytes && routeArena.usedBytes != 0u,
                  std::format(L"{}: pluginId={} fills the exact current typed route prefix", relPath, pluginId).c_str(),
                  success);

            const unsigned long exactArenaBytes = routeFacts.requiredArenaBytes;
            FileSystemArena zeroArena{};
            routeFacts = {};
            routeFacts.sizeBytes = sizeof(routeFacts);
            const HRESULT zeroArenaHr =
                routeCapabilities->GetRouteFacts(L"/", FILESYSTEM_COPY, &zeroArena, &routeFacts);
            const bool zeroArenaValid = zeroArenaHr == HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER) &&
                routeFacts.requiredArenaBytes == exactArenaBytes;

            std::vector<unsigned char> undersizedArenaStorage(exactArenaBytes - 1u);
            FileSystemArena undersizedArena{undersizedArenaStorage.data(), exactArenaBytes - 1u, 0u};
            routeFacts = {};
            routeFacts.sizeBytes = sizeof(routeFacts);
            const HRESULT undersizedArenaHr =
                routeCapabilities->GetRouteFacts(L"/", FILESYSTEM_COPY, &undersizedArena, &routeFacts);

            std::vector<unsigned char> exactArenaStorage(exactArenaBytes);
            FileSystemArena exactArena{exactArenaStorage.data(), exactArenaBytes, 0u};
            routeFacts = {};
            routeFacts.sizeBytes = sizeof(routeFacts);
            const HRESULT exactArenaHr =
                routeCapabilities->GetRouteFacts(L"/", FILESYSTEM_COPY, &exactArena, &routeFacts);

            std::vector<unsigned char> oversizedArenaStorage(exactArenaBytes + 64u);
            FileSystemArena oversizedArena{oversizedArenaStorage.data(), exactArenaBytes + 64u, 0u};
            routeFacts = {};
            routeFacts.sizeBytes = sizeof(routeFacts);
            const HRESULT oversizedArenaHr =
                routeCapabilities->GetRouteFacts(L"/", FILESYSTEM_COPY, &oversizedArena, &routeFacts);
            Check(zeroArenaValid && undersizedArenaHr == HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER) &&
                      exactArenaHr == S_OK && exactArena.usedBytes == exactArenaBytes &&
                      oversizedArenaHr == S_OK && oversizedArena.usedBytes == exactArenaBytes,
                  std::format(L"{}: pluginId={} enforces zero, undersized, exact, and oversized typed arenas", relPath, pluginId).c_str(),
                  success);

            const FileSystemRouteContract::QueryResult typed =
                FileSystemRouteContract::Query(fs.get(), L"/", FILESYSTEM_COPY, pluginId);
            const FileSystemRouteContract::QueryState expectedTypedState = routeFacts.pathTextStableIdentity == TRUE
                ? FileSystemRouteContract::QueryState::Available
                : FileSystemRouteContract::QueryState::Unsupported;
            Check(typed.state == expectedTypedState && ! typed.usedArenaFallback && typed.snapshot.providerId == pluginId &&
                      ! typed.snapshot.pathProfileId.empty() && ! typed.snapshot.rootId.empty() &&
                      typed.snapshot.pathIdentity.has_value(),
                  std::format(L"{}: pluginId={} passes the canonical typed route validator without arena fallback", relPath, pluginId).c_str(),
                  success);

            BOOL peerAllowed = 2;
            const HRESULT peerHr = routeCapabilities->IsTransferPeerAllowed(
                L"/", FILESYSTEM_COPY, FILESYSTEM_TRANSFER_PEER_EXPORT, L"selftest/peer", &peerAllowed);
            Check(peerHr == S_OK && (peerAllowed == FALSE || peerAllowed == TRUE),
                  std::format(L"{}: pluginId={} returns a strict typed export-peer decision", relPath, pluginId).c_str(),
                  success);

            FileSystemChildNameValidation nameValidation{.sizeBytes = sizeof(nameValidation)};
            const HRESULT nameHr = routeCapabilities->ValidateChildName(
                L"/", L"r2-valid-child.txt", FILESYSTEM_RENAME, &nameValidation);
            Check(nameHr == S_OK && nameValidation.sizeBytes == sizeof(nameValidation) &&
                      nameValidation.status == FILESYSTEM_CHILD_NAME_VALID && nameValidation.failureStatus == S_OK,
                  std::format(L"{}: pluginId={} validates an ordinary child name through the typed provider surface", relPath, pluginId).c_str(),
                  success);

            routeArena.usedBytes = 0u;
            const wchar_t* joinedPath = nullptr;
            unsigned long joinedRequired = 0u;
            const HRESULT joinHr = routeCapabilities->JoinPath(
                L"/", L"r2-valid-child.txt", FILESYSTEM_RENAME, &routeArena, &joinedPath, &joinedRequired);
            Check(joinHr == S_OK && joinedPath != nullptr && joinedPath[0] != L'\0' &&
                      joinedRequired == routeArena.usedBytes && joinedRequired != 0u,
                  std::format(L"{}: pluginId={} returns a bounded typed provider-owned joined path", relPath, pluginId).c_str(),
                  success);
        }

        if (runProviderMutationProofs)
        {
            TestTrack13ProviderContracts(*fs.get(), relPath, success);
        }
        else
        {
            SkipTrack13ProviderContracts(relPath);
        }

        if (RequiresTransactionalConfigurationProof(relPath))
        {
            wil::com_ptr_nothrow<IInformations> information;
            const HRESULT informationHr = fs->QueryInterface(__uuidof(IInformations), information.put_void());
            Check(informationHr == S_OK && information != nullptr,
                  std::format(L"{}: transactional configuration provider exposes IInformations for pluginId={}", relPath, pluginId).c_str(),
                  success);
            if (information)
            {
                TestTransactionalConfiguration(*information.get(), relPath, pluginId, success);
            }
        }

        wil::com_ptr_nothrow<IFileSystemPathCapabilities2> pathCapabilities;
        const HRESULT pathCapabilitiesQi =
            fs->QueryInterface(__uuidof(IFileSystemPathCapabilities2), pathCapabilities.put_void());
        Check(pathCapabilitiesQi == S_OK && pathCapabilities != nullptr,
              std::format(L"{}: pluginId={} exposes mandatory IFileSystemPathCapabilities2", relPath, pluginId).c_str(),
              success);
        if (! pathCapabilities)
        {
            continue;
        }

        const char* capJson = nullptr;
        const HRESULT hrCap = pathCapabilities->GetPathCapabilities(L"/", FILESYSTEM_COPY, &capJson);
        Check(hrCap == S_OK, std::format(L"{}: GetPathCapabilities(pluginId={}) returns S_OK", relPath, pluginId).c_str(), success);
        Check(capJson != nullptr && capJson[0] != '\0',
              std::format(L"{}: GetPathCapabilities(pluginId={}) returns non-empty JSON", relPath, pluginId).c_str(),
              success);

        if (hrCap != S_OK || capJson == nullptr || capJson[0] == '\0')
        {
            continue;
        }

        unique_yyjson_doc doc = ParseJson(capJson);
        Check(static_cast<bool>(doc), std::format(L"{}: GetPathCapabilities(pluginId={}) JSON parses successfully", relPath, pluginId).c_str(), success);

        if (! doc)
        {
            continue;
        }

        yyjson_val* root = yyjson_doc_get_root(doc.get());
        Check(root != nullptr && yyjson_is_obj(root), std::format(L"{}: GetPathCapabilities(pluginId={}) root is object", relPath, pluginId).c_str(), success);

        if (root == nullptr || ! yyjson_is_obj(root))
        {
            continue;
        }

        // Capability v2 is path scoped and has no v1 compatibility shape.
        yyjson_val* verVal = yyjson_obj_get(root, "version");
        Check(verVal != nullptr && yyjson_is_int(verVal) && yyjson_get_int(verVal) == 2,
              std::format(L"{}: GetPathCapabilities(pluginId={}) has \"version\": 2", relPath, pluginId).c_str(),
              success);

        constexpr std::array<std::string_view, 10> requiredObjects{
            "operations", "concurrency", "transfer", "identity", "publication", "links", "metadata", "verification", "cancellation", "names"};
        for (const std::string_view required : requiredObjects)
        {
            yyjson_val* value = yyjson_obj_getn(root, required.data(), required.size());
            Check(value != nullptr && yyjson_is_obj(value),
                  std::format(L"{}: GetPathCapabilities(pluginId={}) has required '{}' object",
                              relPath,
                              pluginId,
                              std::wstring(required.begin(), required.end()))
                      .c_str(),
                  success);
        }
        yyjson_val* pathProfile = yyjson_obj_get(root, "pathProfile");
        Check(pathProfile != nullptr && yyjson_is_str(pathProfile) && yyjson_get_len(pathProfile) != 0u,
              std::format(L"{}: GetPathCapabilities(pluginId={}) has a non-empty pathProfile", relPath, pluginId).c_str(),
              success);
        yyjson_val* rootId = yyjson_obj_get(root, "rootId");
        Check(rootId != nullptr && yyjson_is_str(rootId) && yyjson_get_len(rootId) != 0u,
              std::format(L"{}: GetPathCapabilities(pluginId={}) has a non-empty path-scoped rootId", relPath, pluginId).c_str(),
              success);
        yyjson_val* directories = yyjson_obj_get(root, "directories");
        Check(directories != nullptr && yyjson_is_obj(directories),
              std::format(L"{}: GetPathCapabilities(pluginId={}) has required directories object", relPath, pluginId).c_str(),
              success);

        yyjson_val* operations = yyjson_obj_get(root, "operations");
        constexpr std::array<std::string_view, 10> requiredOperationBooleans{
            "copy", "move", "nativeMove", "delete", "rename", "createDirectory", "properties", "read", "write", "recycle"};
        for (const std::string_view required : requiredOperationBooleans)
        {
            yyjson_val* value = operations && yyjson_is_obj(operations) ? yyjson_obj_getn(operations, required.data(), required.size()) : nullptr;
            Check(value != nullptr && yyjson_is_bool(value),
                  std::format(L"{}: GetPathCapabilities(pluginId={}) operations.{} is a required boolean",
                              relPath,
                              pluginId,
                              std::wstring(required.begin(), required.end()))
                      .c_str(),
                  success);
        }

        // The diagnostic JSON is hand-written per provider while IFileSystemRouteCapabilities is
        // the executable authority. Every route section the JSON repeats must equal the typed facts.
        if (routeCapabilities)
        {
            const FileSystemRouteContract::QueryResult typed =
                FileSystemRouteContract::Query(routeCapabilities.get(), L"/", FILESYSTEM_COPY, pluginId);
            const bool typedFactsCopied = typed.snapshot.pathIdentity.has_value() &&
                ((typed.state == FileSystemRouteContract::QueryState::Available && typed.status == S_OK) ||
                 (typed.state == FileSystemRouteContract::QueryState::Unsupported && typed.status == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED)));
            Check(typedFactsCopied,
                  std::format(L"{}: pluginId={} typed route facts are copied for the JSON drift check (state={}, status=0x{:08X})",
                              relPath,
                              pluginId,
                              static_cast<int>(typed.state),
                              static_cast<unsigned long>(typed.status))
                      .c_str(),
                  success);
            if (typedFactsCopied)
            {
                const FileSystemRouteContract::Snapshot& facts = typed.snapshot;
                const auto utf8 = [](std::wstring_view text) -> std::string
                {
                    if (text.empty())
                    {
                        return {};
                    }
                    const int required = WideCharToMultiByte(CP_UTF8, 0, text.data(), static_cast<int>(text.size()), nullptr, 0, nullptr, nullptr);
                    std::string result(required > 0 ? static_cast<size_t>(required) : 0u, '\0');
                    if (required > 0)
                    {
                        WideCharToMultiByte(CP_UTF8, 0, text.data(), static_cast<int>(text.size()), result.data(), required, nullptr, nullptr);
                    }
                    return result;
                };
                const auto jsonObject = [&](const char* key) noexcept -> yyjson_val*
                {
                    yyjson_val* value = yyjson_obj_get(root, key);
                    return value != nullptr && yyjson_is_obj(value) ? value : nullptr;
                };
                const auto jsonBool = [](yyjson_val* object, const char* key, bool expected) noexcept -> bool
                {
                    yyjson_val* value = object != nullptr ? yyjson_obj_get(object, key) : nullptr;
                    return value != nullptr && yyjson_is_bool(value) && yyjson_get_bool(value) == expected;
                };
                const auto jsonUnsigned = [](yyjson_val* object, const char* key, uint64_t expected) noexcept -> bool
                {
                    yyjson_val* value = object != nullptr ? yyjson_obj_get(object, key) : nullptr;
                    return value != nullptr && yyjson_is_uint(value) && yyjson_get_uint(value) == expected;
                };
                const auto jsonString = [](yyjson_val* object, const char* key, std::string_view expected) noexcept -> bool
                {
                    yyjson_val* value = object != nullptr ? yyjson_obj_get(object, key) : nullptr;
                    return value != nullptr && yyjson_is_str(value) &&
                           std::string_view(yyjson_get_str(value), yyjson_get_len(value)) == expected;
                };
                const auto checkDrift = [&](bool matches, const wchar_t* section) noexcept
                {
                    Check(matches,
                          std::format(L"{}: GetPathCapabilities(pluginId={}) '{}' equals the typed route facts", relPath, pluginId, section).c_str(),
                          success);
                };

                yyjson_val* const concurrency  = jsonObject("concurrency");
                yyjson_val* const publication  = jsonObject("publication");
                yyjson_val* const identity     = jsonObject("identity");
                yyjson_val* const links        = jsonObject("links");
                yyjson_val* const verification = jsonObject("verification");
                yyjson_val* const cancellation = jsonObject("cancellation");
                yyjson_val* const names        = jsonObject("names");

                checkDrift(jsonString(root, "pathProfile", utf8(facts.pathProfileId)) && jsonString(root, "rootId", utf8(facts.rootId)),
                           L"pathProfile/rootId");
                checkDrift(jsonBool(operations, "copy", facts.copyOperation) && jsonBool(operations, "move", facts.moveOperation) &&
                               jsonBool(operations, "nativeMove", facts.nativeMoveOperation) &&
                               jsonBool(operations, "delete", facts.deleteOperation) && jsonBool(operations, "rename", facts.renameOperation) &&
                               jsonBool(operations, "createDirectory", facts.createDirectoryOperation) &&
                               jsonBool(operations, "properties", facts.properties) && jsonBool(operations, "read", facts.read) &&
                               jsonBool(operations, "write", facts.write) && jsonBool(operations, "recycle", facts.recycleOperation),
                           L"operations");
                // deleteRecycleBinMax is intentionally not compared: the typed record keeps every bound
                // positive while a provider without Recycle may publish 0 in its diagnostics.
                checkDrift(jsonUnsigned(concurrency, "copyMoveMax", facts.copyMoveMaxConcurrency) &&
                               jsonUnsigned(concurrency, "deleteMax", facts.deleteMaxConcurrency),
                           L"concurrency");
                checkDrift(jsonBool(publication, "exclusiveStage", facts.exclusiveStage) &&
                               jsonBool(publication, "conditionalPublish", facts.conditionalPublish) &&
                               jsonBool(publication, "committedSize", facts.committedSize),
                           L"publication");
                checkDrift(jsonBool(identity, "boundDelete", facts.boundDelete) && jsonBool(identity, "conditionalDelete", facts.conditionalDelete),
                           L"identity");
                checkDrift(jsonBool(links, "preserveFileLink", facts.preserveFileLink) &&
                               jsonBool(links, "preserveDirectoryLink", facts.preserveDirectoryLink) &&
                               jsonBool(links, "retargetInTree", facts.retargetInTree) &&
                               jsonBool(links, "exactLinkRemoval", facts.exactLinkRemoval),
                           L"links");
                checkDrift(jsonBool(verification, "hostReadback", (facts.proofFlags & FILESYSTEM_ROUTE_PROOF_HOST_READBACK) != 0u) &&
                               jsonString(verification,
                                          "providerProof",
                                          (facts.proofFlags & FILESYSTEM_ROUTE_PROOF_PROVIDER_BLAKE3) != 0u  ? "blake3-bound-object"
                                          : (facts.proofFlags & FILESYSTEM_ROUTE_PROOF_WRITER_DIGEST) != 0u ? "writer-digest"
                                                                                                           : "none"),
                           L"verification");
                const std::string_view expectedRouteClass = facts.cancellationRoute == FILESYSTEM_CANCELLATION_BOUNDED ? "bounded"
                    : facts.cancellationRoute == FILESYSTEM_CANCELLATION_PROVIDER_WATCHDOG            ? "providerWatchdog"
                                                                                                       : "uncontained";
                checkDrift(jsonBool(cancellation, "abort", facts.cancellationAbort) &&
                               jsonBool(cancellation, "deadline", facts.cancellationDeadline) &&
                               jsonString(cancellation, "routeClass", expectedRouteClass) &&
                               jsonUnsigned(cancellation, "providerWatchdogTimeoutMs", facts.providerWatchdogTimeoutMs),
                           L"cancellation");
                const std::optional<FileSystemPathIdentity> jsonIdentity =
                    ParseDiagnosticFileSystemPathIdentityContractFromRoot(root, pluginId);
                checkDrift(jsonIdentity.has_value() && facts.pathIdentity.has_value() &&
                               jsonIdentity->pathTextStableIdentity == facts.pathIdentity->pathTextStableIdentity &&
                               jsonIdentity->componentComparison == facts.pathIdentity->componentComparison &&
                               jsonIdentity->preferredSeparator == facts.pathIdentity->preferredSeparator &&
                               jsonIdentity->acceptedSeparators == facts.pathIdentity->acceptedSeparators &&
                               jsonIdentity->casePreserving == facts.pathIdentity->casePreserving &&
                               jsonIdentity->caseOnlyRename == facts.pathIdentity->caseOnlyRename &&
                               jsonUnsigned(names, "maxComponentUtf16", facts.maxComponentUtf16),
                           L"names");
            }
        }

        struct ExpectedDestructiveCapabilities
        {
            bool deleteOperation  = false;
            bool renameOperation  = false;
            bool recycleOperation = false;
        };
        const std::wstring_view pluginIdView(pluginId);
        std::optional<ExpectedDestructiveCapabilities> expectedDestructive;
        if (pluginIdView == L"builtin/file-system-dummy" || pluginIdView == L"builtin/file-system-mtp" ||
            pluginIdView == L"builtin/file-system-imap" || pluginIdView == L"builtin/file-system-s3table")
        {
            expectedDestructive = ExpectedDestructiveCapabilities{};
        }
        else if (pluginIdView == L"builtin/file-system-ftp" || pluginIdView == L"builtin/file-system-sftp" ||
                 pluginIdView == L"builtin/file-system-scp")
        {
            // R0f-Curl: FTP/SFTP/SCP are full file-manager destinations; only Recycle stays absent.
            expectedDestructive = ExpectedDestructiveCapabilities{.deleteOperation = true, .renameOperation = true};
        }
        else if (pluginIdView == L"builtin/file-system-onedrive-personal" ||
                 pluginIdView == L"builtin/file-system-onedrive-business" || pluginIdView == L"builtin/file-system-sharepoint")
        {
            // R0f-Graph: OneDrive/SharePoint rename through Graph; Delete stays Recycle by stable item ID.
            expectedDestructive = ExpectedDestructiveCapabilities{.renameOperation = true, .recycleOperation = true};
        }
        else if (pluginIdView == L"builtin/file-system-s3")
        {
            expectedDestructive = ExpectedDestructiveCapabilities{.deleteOperation = true};
        }
        if (expectedDestructive.has_value() && operations && yyjson_is_obj(operations))
        {
            const auto checkOperation = [&](const char* key, bool expected) noexcept
            {
                yyjson_val* value = yyjson_obj_get(operations, key);
                Check(value != nullptr && yyjson_is_bool(value) && yyjson_get_bool(value) == expected,
                      std::format(L"{}: pluginId={} operations.{} truthfully reports {}",
                                  relPath,
                                  pluginId,
                                  std::wstring(key, key + std::char_traits<char>::length(key)),
                                  expected ? L"true" : L"false")
                          .c_str(),
                      success);
            };
            checkOperation("delete", expectedDestructive->deleteOperation);
            checkOperation("rename", expectedDestructive->renameOperation);
            checkOperation("recycle", expectedDestructive->recycleOperation);
        }

        yyjson_val* verification = yyjson_obj_get(root, "verification");
        yyjson_val* hostReadback = verification && yyjson_is_obj(verification)
            ? yyjson_obj_get(verification, "hostReadback")
            : nullptr;
        Check(hostReadback != nullptr && yyjson_is_bool(hostReadback),
              std::format(L"{}: GetPathCapabilities(pluginId={}) verification.hostReadback is a required boolean", relPath, pluginId).c_str(),
              success);
        yyjson_val* providerProof = verification && yyjson_is_obj(verification)
            ? yyjson_obj_get(verification, "providerProof")
            : nullptr;
        const std::string_view providerProofValue = providerProof != nullptr && yyjson_is_str(providerProof)
            ? std::string_view(yyjson_get_str(providerProof), yyjson_get_len(providerProof))
            : std::string_view{};
        Check(providerProof != nullptr && yyjson_is_str(providerProof) &&
                  (providerProofValue == "none" || providerProofValue == "blake3-bound-object" || providerProofValue == "writer-digest"),
              std::format(L"{}: GetPathCapabilities(pluginId={}) verification.providerProof is a supported required string",
                          relPath,
                          pluginId)
                  .c_str(),
              success);

        yyjson_val* cancellation = yyjson_obj_get(root, "cancellation");
        yyjson_val* routeClass = cancellation && yyjson_is_obj(cancellation)
            ? yyjson_obj_get(cancellation, "routeClass")
            : nullptr;
        const std::string_view routeClassValue = routeClass && yyjson_is_str(routeClass)
            ? std::string_view(yyjson_get_str(routeClass), yyjson_get_len(routeClass))
            : std::string_view{};
        Check(routeClass != nullptr && yyjson_is_str(routeClass) &&
                  (routeClassValue == "bounded" || routeClassValue == "providerWatchdog" || routeClassValue == "uncontained"),
              std::format(L"{}: GetPathCapabilities(pluginId={}) cancellation.routeClass is a supported required string",
                          relPath,
                          pluginId)
                  .c_str(),
              success);
        yyjson_val* providerWatchdogTimeout = cancellation && yyjson_is_obj(cancellation)
            ? yyjson_obj_get(cancellation, "providerWatchdogTimeoutMs")
            : nullptr;
        Check(providerWatchdogTimeout != nullptr && yyjson_is_uint(providerWatchdogTimeout),
              std::format(L"{}: GetPathCapabilities(pluginId={}) cancellation.providerWatchdogTimeoutMs is a required unsigned integer",
                          relPath,
                          pluginId)
                  .c_str(),
              success);
        const uint64_t providerWatchdogTimeoutValue = providerWatchdogTimeout && yyjson_is_uint(providerWatchdogTimeout)
            ? yyjson_get_uint(providerWatchdogTimeout)
            : 0u;
        Check((routeClassValue == "providerWatchdog" && providerWatchdogTimeoutValue > 0u &&
               providerWatchdogTimeoutValue <= (std::numeric_limits<uint32_t>::max)()) ||
                  ((routeClassValue == "bounded" || routeClassValue == "uncontained") && providerWatchdogTimeoutValue == 0u),
              std::format(L"{}: GetPathCapabilities(pluginId={}) cancellation route/watchdog combination is fail-closed",
                          relPath,
                          pluginId)
                  .c_str(),
              success);
        yyjson_val* retiredHostQuietPointTimeout = cancellation && yyjson_is_obj(cancellation)
            ? yyjson_obj_get(cancellation, "quietPointTimeoutMs")
            : nullptr;
        Check(retiredHostQuietPointTimeout == nullptr,
              std::format(L"{}: GetPathCapabilities(pluginId={}) does not advertise the retired unenforced quietPointTimeoutMs",
                          relPath,
                          pluginId)
                  .c_str(),
              success);

        const std::optional<FileSystemPathIdentity> parsedPathIdentity =
            ParseDiagnosticFileSystemPathIdentityContract(std::string_view(capJson), {});
        Check(parsedPathIdentity.has_value(),
              std::format(L"{}: GetPathCapabilities(pluginId={}) names contract parses under the host contract", relPath, pluginId).c_str(),
              success);

        const char* invalidOutput = reinterpret_cast<const char*>(static_cast<uintptr_t>(1u));
        Check(pathCapabilities->GetPathCapabilities(nullptr, FILESYSTEM_COPY, &invalidOutput) == E_INVALIDARG && invalidOutput == nullptr,
              std::format(L"{}: GetPathCapabilities(pluginId={}) rejects null path and clears output", relPath, pluginId).c_str(),
              success);

        const char* createDirectoryJson = nullptr;
        const HRESULT createDirectoryCapabilityHr =
            pathCapabilities->GetPathCapabilities(L"/", FILESYSTEM_CREATE_DIRECTORY, &createDirectoryJson);
        Check(createDirectoryCapabilityHr == S_OK && createDirectoryJson != nullptr && createDirectoryJson[0] != '\0',
              std::format(L"{}: GetPathCapabilities(pluginId={}) accepts the Create Directory capability operation", relPath, pluginId).c_str(),
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
    if (! mod)
    {
        return false;
    }

    const auto pfnCreate = reinterpret_cast<PfnCreate>(GetProcAddress(mod.get(), "RedSalamanderCreate"));
    Check(pfnCreate != nullptr, L"FileSystemDummy.dll: RedSalamanderCreate resolves for negative-id test", success);
    if (pfnCreate == nullptr)
    {
        return false;
    }

    void* raw        = nullptr;
    const HRESULT hr = pfnCreate(__uuidof(IFileSystem), nullptr, static_cast<IHost*>(&g_nullHost), L"no/such-plugin", &raw);
    Check(hr == HRESULT_FROM_WIN32(ERROR_NOT_FOUND), L"FileSystemDummy.dll: Create(bogus-id) returns HRESULT_FROM_WIN32(ERROR_NOT_FOUND)", success);
    Check(raw == nullptr, L"FileSystemDummy.dll: Create(bogus-id) returns null result pointer", success);

    return true;
}

[[nodiscard]] constexpr bool IsDebugSelfTestExportPresenceAccepted(bool exportPresent, bool exportRequired) noexcept
{
    return exportPresent || ! exportRequired;
}

[[nodiscard]] bool ValidateDebugSelfTestExportPresence(bool exportPresent, bool exportRequired, std::wstring_view label, bool& success) noexcept
{
    if (exportPresent)
    {
        return true;
    }

    if (exportRequired)
    {
        Check(false, std::format(L"{}: required debug selftest export resolves", label).c_str(), success);
        return false;
    }

    std::wcout << L"[  SKIPPED ] " << label << L": debug selftest export is not expected in this configuration\n";
    return false;
}

bool TestPluginDebugSelfTests(
    std::wstring_view relativeDllPath, const char* exportName, std::wstring_view label, bool exportRequired, bool& success) noexcept
{
    const std::wstring exeDir  = GetExeDir();
    const std::wstring absPath = exeDir + std::wstring(relativeDllPath);

    wil::unique_hmodule mod(LoadLibraryExW(absPath.c_str(), nullptr, 0));
    const std::wstring loadMessage = std::format(L"{}: loads for debug selftests", label);
    Check(static_cast<bool>(mod), loadMessage.c_str(), success);
    if (! mod)
    {
        return false;
    }

    const auto pfnSelfTests = reinterpret_cast<PfnRunDebugSelfTests>(GetProcAddress(mod.get(), exportName));
    if (! ValidateDebugSelfTestExportPresence(pfnSelfTests != nullptr, exportRequired, label, success))
    {
        return ! exportRequired;
    }

    unsigned int passed = 0;
    unsigned int failed = 0;
    const HRESULT hr    = pfnSelfTests(&passed, &failed);
    const std::wstring passMessage =
        std::format(L"{}: debug selftests pass (passed={}, failed={}, hr=0x{:08X})", label, passed, failed, static_cast<unsigned long>(hr));
    Check(SUCCEEDED(hr) && failed == 0 && passed > 0, passMessage.c_str(), success);

    const auto shutdown = reinterpret_cast<PfnPluginShutdown>(GetProcAddress(mod.get(), "RedSalamanderPluginShutdown"));
    const auto canUnload = reinterpret_cast<PfnPluginCanUnloadNow>(GetProcAddress(mod.get(), "RedSalamanderPluginCanUnloadNow"));
    if (shutdown && canUnload)
    {
        shutdown();
        Check(canUnload() == TRUE, std::format(L"{}: debug selftest module reaches its unload quiet point", label).c_str(), success);
    }
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

bool RunCapabilitiesPass(std::span<const std::wstring_view> dlls, bool runProviderMutationProofs) noexcept
{
    bool success = true;
    for (const std::wstring_view relPath : dlls)
    {
        TestCapabilities(relPath, runProviderMutationProofs, success);
    }
    return success;
}

void TestS3RuntimeUnloadContract(bool& success) noexcept
{
    const std::wstring absPath = GetExeDir() + L"Plugins\\FileSystemS3.dll";
    constexpr unsigned int kRefreshCycles = 8u;
    for (unsigned int cycle = 0u; cycle < kRefreshCycles; ++cycle)
    {
        wil::unique_hmodule module(LoadLibraryExW(absPath.c_str(), nullptr, 0));
        Check(static_cast<bool>(module), std::format(L"FileSystemS3.dll: refresh cycle {} loads the module", cycle + 1u).c_str(), success);
        if (! module)
        {
            return;
        }

        const auto create = reinterpret_cast<PfnCreate>(GetProcAddress(module.get(), "RedSalamanderCreate"));
        const auto shutdown = reinterpret_cast<PfnPluginShutdown>(GetProcAddress(module.get(), "RedSalamanderPluginShutdown"));
        const auto canUnload = reinterpret_cast<PfnPluginCanUnloadNow>(GetProcAddress(module.get(), "RedSalamanderPluginCanUnloadNow"));
        Check(create != nullptr && shutdown != nullptr && canUnload != nullptr,
              std::format(L"FileSystemS3.dll: refresh cycle {} resolves lifecycle exports", cycle + 1u).c_str(),
              success);
        if (create == nullptr || shutdown == nullptr || canUnload == nullptr)
        {
            return;
        }

        void* raw              = nullptr;
        const HRESULT createHr = create(__uuidof(IFileSystem), nullptr, &g_nullHost, L"builtin/file-system-s3", &raw);
        wil::com_ptr_nothrow<IFileSystem> fileSystem;
        fileSystem.attach(static_cast<IFileSystem*>(raw));
        Check(createHr == S_OK && fileSystem,
              std::format(L"FileSystemS3.dll: refresh cycle {} creates an initialized owner", cycle + 1u).c_str(),
              success);
        if (FAILED(createHr) || ! fileSystem)
        {
            return;
        }

        Check(canUnload() == FALSE,
              std::format(L"FileSystemS3.dll: refresh cycle {} is closed before the quiet point", cycle + 1u).c_str(),
              success);
        shutdown();
        Check(canUnload() == FALSE,
              std::format(L"FileSystemS3.dll: refresh cycle {} stays closed while an AWS owner is alive", cycle + 1u).c_str(),
              success);
        fileSystem.reset();
        Check(canUnload() == TRUE,
              std::format(L"FileSystemS3.dll: refresh cycle {} opens after AWS shutdown", cycle + 1u).c_str(),
              success);
        shutdown();
        Check(canUnload() == TRUE,
              std::format(L"FileSystemS3.dll: refresh cycle {} keeps shutdown idempotent", cycle + 1u).c_str(),
              success);
        module.reset();
        Check(GetModuleHandleW(absPath.c_str()) == nullptr,
              std::format(L"FileSystemS3.dll: refresh cycle {} releases the module", cycle + 1u).c_str(),
              success);
    }
}

#if defined(_DEBUG)
void TestCrossPluginCurlRuntimeUnloadContract(bool& success) noexcept
{
    struct RuntimeModule
    {
        RuntimeModule() = default;
        RuntimeModule(const RuntimeModule&)            = delete;
        RuntimeModule& operator=(const RuntimeModule&) = delete;
        RuntimeModule(RuntimeModule&&)                 = default;
        RuntimeModule& operator=(RuntimeModule&&)      = default;

        std::wstring label;
        std::wstring path;
        wil::unique_hmodule module;
        PfnPluginShutdown shutdown       = nullptr;
        PfnPluginCanUnloadNow canUnload  = nullptr;
        PfnDebugCurlRuntimeProbe probe   = nullptr;
    };

    const auto loadModule = [&](std::wstring_view fileName, std::wstring_view label) -> RuntimeModule
    {
        RuntimeModule loaded;
        loaded.label  = label;
        loaded.path   = GetExeDir() + L"Plugins\\" + std::wstring(fileName);
        loaded.module.reset(LoadLibraryExW(loaded.path.c_str(), nullptr, 0));
        Check(static_cast<bool>(loaded.module), std::format(L"{}: loads for shared libcurl runtime proof", label).c_str(), success);
        if (loaded.module)
        {
            loaded.shutdown = reinterpret_cast<PfnPluginShutdown>(GetProcAddress(loaded.module.get(), "RedSalamanderPluginShutdown"));
            loaded.canUnload = reinterpret_cast<PfnPluginCanUnloadNow>(GetProcAddress(loaded.module.get(), "RedSalamanderPluginCanUnloadNow"));
            loaded.probe = reinterpret_cast<PfnDebugCurlRuntimeProbe>(GetProcAddress(loaded.module.get(), "RedSalamanderDebugCurlRuntimeProbe"));
            Check(loaded.shutdown && loaded.canUnload && loaded.probe,
                  std::format(L"{}: resolves shared libcurl lifecycle and probe exports", label).c_str(),
                  success);
        }
        return loaded;
    };

    const auto unloadOne = [&](RuntimeModule& target, unsigned int cycle)
    {
        target.shutdown();
        Check(target.canUnload() == TRUE,
              std::format(L"{}: refresh cycle {} reaches its quiet point while its peer survives", target.label, cycle).c_str(),
              success);
        target.module.reset();
        Check(GetModuleHandleW(target.path.c_str()) == nullptr,
              std::format(L"{}: refresh cycle {} physically unloads", target.label, cycle).c_str(),
              success);
    };

    constexpr unsigned int kRefreshCycles = 8u;
    for (unsigned int cycle = 1u; cycle <= kRefreshCycles; ++cycle)
    {
        RuntimeModule curl   = loadModule(L"FileSystemCurl.dll", L"FileSystemCurl.dll");
        RuntimeModule google = loadModule(L"FileSystemGoogleDrive.dll", L"FileSystemGoogleDrive.dll");
        if (! curl.module || ! google.module || ! curl.shutdown || ! google.shutdown || ! curl.canUnload || ! google.canUnload || ! curl.probe ||
            ! google.probe)
        {
            return;
        }

        Check(curl.probe() == S_OK && google.probe() == S_OK,
              std::format(L"shared libcurl refresh cycle {} initializes both plugin participants", cycle).c_str(),
              success);

        RuntimeModule& first    = (cycle % 2u) != 0u ? curl : google;
        RuntimeModule& survivor = (cycle % 2u) != 0u ? google : curl;
        unloadOne(first, cycle);
        Check(survivor.probe() == S_OK,
              std::format(L"{}: refresh cycle {} still creates a libcurl easy handle after peer unload", survivor.label, cycle).c_str(),
              success);
        unloadOne(survivor, cycle);
    }
}
#endif

[[nodiscard]] std::wstring GetRequiredEnvironmentValue(const wchar_t* name)
{
    const DWORD required = GetEnvironmentVariableW(name, nullptr, 0u);
    if (required <= 1u)
    {
        return {};
    }
    std::wstring value(required, L'\0');
    const DWORD written = GetEnvironmentVariableW(name, value.data(), required);
    if (written != required - 1u)
    {
        return {};
    }
    value.resize(written);
    return value;
}

[[nodiscard]] std::string NarrowAscii(std::wstring_view value)
{
    std::string result;
    result.reserve(value.size());
    for (const wchar_t character : value)
    {
        if (character < 0x20 || character > 0x7Eu)
        {
            return {};
        }
        result.push_back(static_cast<char>(character));
    }
    return result;
}

[[nodiscard]] std::string EscapeJsonAscii(std::string_view value)
{
    std::string escaped;
    escaped.reserve(value.size());
    for (const char character : value)
    {
        switch (character)
        {
        case '\\':
            escaped.append("\\\\");
            break;
        case '"':
            escaped.append("\\\"");
            break;
        default:
            escaped.push_back(character);
            break;
        }
    }
    return escaped;
}

[[nodiscard]] bool IsLowerHex(std::wstring_view value, size_t length) noexcept
{
    return value.size() == length && std::ranges::all_of(value, [](wchar_t character) noexcept
    { return (character >= L'0' && character <= L'9') || (character >= L'a' && character <= L'f'); });
}

[[nodiscard]] uint64_t Percentile(
    const std::array<uint64_t, TerminalVtUpgradeTestContract::kSampleCount>& samples,
    size_t numerator,
    size_t denominator)
{
    std::array<uint64_t, TerminalVtUpgradeTestContract::kSampleCount> sorted = samples;
    std::ranges::sort(sorted);
    const size_t index = std::min(
        sorted.size() - 1u, (sorted.size() * numerator + denominator - 1u) / denominator - 1u);
    return sorted[index];
}

void WriteSamples(
    std::ofstream& output,
    const std::array<uint64_t, TerminalVtUpgradeTestContract::kSampleCount>& samples)
{
    output << '[';
    for (size_t index = 0u; index < samples.size(); ++index)
    {
        if (index != 0u)
        {
            output << ',';
        }
        output << samples[index];
    }
    output << ']';
}

[[nodiscard]] bool WriteTerminalVtUpgradeArchive(
    const std::filesystem::path& outputDirectory,
    const TerminalVtUpgradeTestContract::Evidence& evidence)
{
    const std::wstring repositoryCommit = GetRequiredEnvironmentValue(L"RS_VT_REPOSITORY_COMMIT");
    const std::wstring branch = GetRequiredEnvironmentValue(L"RS_VT_BRANCH");
    const std::wstring machineHash = GetRequiredEnvironmentValue(L"RS_VT_MACHINE_HASH");
    const std::wstring runId = GetRequiredEnvironmentValue(L"RS_VT_RUN_ID");
    const std::wstring buildReceipt = GetRequiredEnvironmentValue(L"RS_VT_BUILD_RECEIPT");
    const std::wstring runtimeIdentity = GetRequiredEnvironmentValue(L"RS_VT_RUNTIME_IDENTITY");
    const std::wstring windowsBuild = GetRequiredEnvironmentValue(L"RS_VT_WINDOWS_BUILD");
    const std::wstring cpu = GetRequiredEnvironmentValue(L"RS_VT_CPU");
    const std::wstring compiler = GetRequiredEnvironmentValue(L"RS_VT_COMPILER");
    const std::wstring sdk = GetRequiredEnvironmentValue(L"RS_VT_SDK");
    const std::wstring zig = GetRequiredEnvironmentValue(L"RS_VT_ZIG");
    const std::wstring vcpkg = GetRequiredEnvironmentValue(L"RS_VT_VCPKG");
    if (! IsLowerHex(repositoryCommit, 40u) || branch.empty() || ! IsLowerHex(machineHash, 12u) ||
        runId.empty() || ! IsLowerHex(buildReceipt, 64u) || ! IsLowerHex(runtimeIdentity, 64u) ||
        windowsBuild.empty() || cpu.empty() || compiler.empty() || sdk.empty() || zig.empty() || vcpkg.empty())
    {
        std::wcerr << L"VT upgrade evidence metadata environment is missing or invalid.\n";
        return false;
    }
    const std::array<std::wstring_view, 7u> asciiValues{
        branch, runId, windowsBuild, cpu, compiler, sdk, zig};
    if (std::ranges::any_of(asciiValues, [](std::wstring_view value) { return NarrowAscii(value).empty(); }) ||
        NarrowAscii(vcpkg).empty())
    {
        std::wcerr << L"VT upgrade evidence metadata must be nonempty printable ASCII.\n";
        return false;
    }

    std::error_code error;
    std::filesystem::create_directories(outputDirectory / L"perf", error);
    if (error)
    {
        std::wcerr << L"Unable to create the VT upgrade evidence directory.\n";
        return false;
    }

    const std::filesystem::path perfPath = outputDirectory / L"perf" / L"perf_metrics.jsonl";
    Debug::Perf::ConfigureJsonlOutput(
        perfPath, L"terminal-vt-upgrade-corpus-v1", L"Release", branch, repositoryCommit, machineHash, runId);
    for (size_t index = 0u; index < evidence.sampleCount; ++index)
    {
        Debug::Perf::EmitDurationUs(
            L"terminal.vt_upgrade.vt_write_us", evidence.vtWriteSamplesUs[index], evidence.inputByteCount, index);
        Debug::Perf::EmitDurationUs(
            L"terminal.vt_upgrade.snapshot_us", evidence.snapshotSamplesUs[index], evidence.inputByteCount, index);
        Debug::Perf::EmitDurationUs(
            L"terminal.vt_upgrade.hosted_render_us", evidence.hostedRenderSamplesUs[index], 120u * 40u, index);
    }
    Debug::Perf::EmitValue(L"terminal.vt_upgrade.kitty_placement_count", evidence.kittyPlacementCount);
    Debug::Perf::EmitValue(L"terminal.vt_upgrade.kitty_placement_bytes", evidence.kittyPlacementBytes);
    Debug::Perf::EmitValue(L"terminal.vt_upgrade.kitty_memory_high_water_bytes", evidence.kittyMemoryHighWaterBytes);
    Debug::Perf::EmitValue(L"terminal.vt_upgrade.kitty_cache_bytes", evidence.kittyCacheBytes);
    Debug::Perf::EmitValue(L"terminal.vt_upgrade.kitty_pinned_bytes", evidence.kittyPinnedBytes);
    Debug::Perf::EmitValue(L"terminal.vt_upgrade.peak_private_bytes", evidence.peakPrivateBytes);
    Debug::Perf::EmitValue(L"terminal.vt_upgrade.peak_working_set_bytes", evidence.peakWorkingSetBytes);
    Debug::Perf::ClearJsonlOutput();

    const uint64_t vtP50 = Percentile(evidence.vtWriteSamplesUs, 50u, 100u);
    const uint64_t vtP95 = Percentile(evidence.vtWriteSamplesUs, 95u, 100u);
    const uint64_t snapshotP50 = Percentile(evidence.snapshotSamplesUs, 50u, 100u);
    const uint64_t snapshotP95 = Percentile(evidence.snapshotSamplesUs, 95u, 100u);
    const uint64_t renderP50 = Percentile(evidence.hostedRenderSamplesUs, 50u, 100u);
    const uint64_t renderP95 = Percentile(evidence.hostedRenderSamplesUs, 95u, 100u);
    const auto threshold = [](uint64_t baseline) noexcept { return (baseline * 110u + 99u) / 100u; };

    SYSTEMTIME created{};
    GetSystemTime(&created);
    const std::string createdUtc = std::format("{:04}-{:02}-{:02}T{:02}:{:02}:{:02}.{:03}Z",
                                                created.wYear,
                                                created.wMonth,
                                                created.wDay,
                                                created.wHour,
                                                created.wMinute,
                                                created.wSecond,
                                                created.wMilliseconds);
    std::ofstream results(outputDirectory / L"results.json", std::ios::binary | std::ios::trunc);
    if (! results)
    {
        return false;
    }
    results << "{\n"
            << "  \"schema\": \"red-salamander.terminal-vt-upgrade-evidence.v1\",\n"
            << "  \"runId\": \"" << EscapeJsonAscii(NarrowAscii(runId)) << "\",\n"
            << "  \"createdUtc\": \"" << createdUtc << "\",\n"
            << "  \"repositoryCommit\": \"" << NarrowAscii(repositoryCommit) << "\",\n"
            << "  \"repositoryBranch\": \"" << EscapeJsonAscii(NarrowAscii(branch)) << "\",\n"
            << "  \"machineHash\": \"" << NarrowAscii(machineHash) << "\",\n"
            << "  \"configuration\": \"Release\",\n"
            << "  \"buildReceiptId\": \"" << NarrowAscii(buildReceipt) << "\",\n"
            << "  \"runtimeIdentitySha256\": \"" << NarrowAscii(runtimeIdentity) << "\",\n"
            << "  \"environment\": {\n"
            << "    \"windowsBuild\": \"" << EscapeJsonAscii(NarrowAscii(windowsBuild)) << "\",\n"
            << "    \"cpu\": \"" << EscapeJsonAscii(NarrowAscii(cpu)) << "\",\n"
            << "    \"architecture\": \"AMD64\",\n"
            << "    \"compiler\": \"" << EscapeJsonAscii(NarrowAscii(compiler)) << "\",\n"
            << "    \"windowsSdk\": \"" << EscapeJsonAscii(NarrowAscii(sdk)) << "\",\n"
            << "    \"zig\": \"" << EscapeJsonAscii(NarrowAscii(zig)) << "\",\n"
            << "    \"vcpkgCommit\": \"" << EscapeJsonAscii(NarrowAscii(vcpkg)) << "\"\n"
            << "  },\n"
            << "  \"corpus\": {\n"
            << "    \"formatId\": \"red-salamander-terminal-vt-upgrade-corpus\",\n"
            << "    \"version\": 1,\n"
            << "    \"sha256\": \"" << evidence.corpusSha256.data() << "\",\n"
            << "    \"inputByteCount\": " << evidence.inputByteCount << "\n"
            << "  },\n"
            << "  \"behavior\": {\n"
            << "    \"ptyOutputSha256\": \"" << evidence.ptyOutputSha256.data() << "\",\n"
            << "    \"snapshotSha256\": \"" << evidence.snapshotSha256.data() << "\",\n"
            << "    \"withoutExpectedDeltas\": {\"ptyOutputSha256\": \""
            << evidence.ptyOutputWithoutExpectedDeltaSha256.data()
            << "\", \"snapshotSha256\": \"" << evidence.snapshotWithoutExpectedDeltaSha256.data() << "\"},\n"
            << "    \"expectedDeltas\": [{\"id\": \"dcs-high-byte-preservation\", \"disposition\": \"candidate-only-if-reviewed\"}]\n"
            << "  },\n"
            << "  \"measurements\": {\n"
            << "    \"sampleCount\": " << evidence.sampleCount << ",\n"
            << "    \"percentileRule\": \"nearest-rank\",\n"
            << "    \"vtWrite\": {\"unit\": \"us\", \"p50\": " << vtP50 << ", \"p95\": " << vtP95
            << ", \"candidateMaximumP95\": " << threshold(vtP95) << ", \"samples\": ";
    WriteSamples(results, evidence.vtWriteSamplesUs);
    results << "},\n"
            << "    \"snapshot\": {\"unit\": \"us\", \"p50\": " << snapshotP50 << ", \"p95\": " << snapshotP95
            << ", \"candidateMaximumP95\": " << threshold(snapshotP95) << ", \"samples\": ";
    WriteSamples(results, evidence.snapshotSamplesUs);
    results << "},\n"
            << "    \"hostedRender\": {\"unit\": \"us\", \"p50\": " << renderP50 << ", \"p95\": " << renderP95
            << ", \"candidateMaximumP95\": " << threshold(renderP95) << ", \"samples\": ";
    WriteSamples(results, evidence.hostedRenderSamplesUs);
    results << "}\n"
            << "  },\n"
            << "  \"kitty\": {\"pendingProbeCount\": " << evidence.kittyPendingProbeCount
            << ", \"completedProbeCount\": " << evidence.kittyCompletedProbeCount
            << ", \"placementCount\": " << evidence.kittyPlacementCount
            << ", \"placementBytes\": " << evidence.kittyPlacementBytes
            << ", \"queuedBytes\": " << evidence.kittyQueuedBytes
            << ", \"activeSourceBytes\": " << evidence.kittyActiveSourceBytes
            << ", \"activeConvertedBytes\": " << evidence.kittyActiveConvertedBytes
            << ", \"readyBytes\": " << evidence.kittyReadyBytes
            << ", \"residentBytes\": " << evidence.kittyResidentBytes
            << ", \"memoryHighWaterBytes\": " << evidence.kittyMemoryHighWaterBytes
            << ", \"cacheBytes\": " << evidence.kittyCacheBytes
            << ", \"pinnedBytes\": " << evidence.kittyPinnedBytes
            << ", \"frameUploadBytes\": " << evidence.kittyFrameUploadBytes
            << ", \"frameUploadCount\": " << evidence.kittyFrameUploadCount << "},\n"
            << "  \"memory\": {\"peakPrivateBytes\": " << evidence.peakPrivateBytes
            << ", \"peakWorkingSetBytes\": " << evidence.peakWorkingSetBytes << "},\n"
            << "  \"assertionCount\": " << evidence.assertionCount << ",\n"
            << "  \"outcome\": \"pass\"\n"
            << "}\n";
    results.close();
    if (! results)
    {
        return false;
    }

    std::ofstream trace(outputDirectory / L"trace.txt", std::ios::binary | std::ios::trunc);
    if (! trace)
    {
        return false;
    }
    trace << "Terminal VT upgrade corpus: PASS\n"
          << "Samples per p95: " << evidence.sampleCount << "\n"
          << "Corpus SHA-256: " << evidence.corpusSha256.data() << "\n"
          << "PTY output SHA-256: " << evidence.ptyOutputSha256.data() << "\n"
          << "Snapshot SHA-256: " << evidence.snapshotSha256.data() << "\n"
          << "VT write p50/p95 us: " << vtP50 << '/' << vtP95 << "\n"
          << "Snapshot p50/p95 us: " << snapshotP50 << '/' << snapshotP95 << "\n"
          << "Hosted render p50/p95 us: " << renderP50 << '/' << renderP95 << "\n";
    trace.close();
    return static_cast<bool>(trace) && std::filesystem::is_regular_file(perfPath, error) && ! error;
}

[[nodiscard]] bool RunTerminalVtUpgradeCorpus(
    const std::filesystem::path& outputDirectory,
    bool& success) noexcept
{
    const std::wstring absPath = GetExeDir() + L"Plugins\\Terminal.dll";
    wil::unique_hmodule module(LoadLibraryExW(absPath.c_str(), nullptr, 0));
    Check(static_cast<bool>(module), L"Terminal.dll: loads for VT upgrade corpus", success);
    if (! module)
    {
        return false;
    }
    const auto run = reinterpret_cast<PfnRunTerminalVtUpgradeCorpus>(
        GetProcAddress(module.get(), "RedSalamanderTerminalDebugVtUpgradeCorpus"));
    Check(run != nullptr, L"Terminal.dll: resolves the test-enabled VT upgrade corpus export", success);
    if (run == nullptr)
    {
        return false;
    }
    TerminalVtUpgradeTestContract::Evidence evidence;
    const HRESULT hr = run(&evidence);
    const bool evidenceReady = SUCCEEDED(hr) && evidence.version == TerminalVtUpgradeTestContract::kVersion &&
        evidence.sampleCount == TerminalVtUpgradeTestContract::kSampleCount && evidence.assertionCount != 0u;
    Check(evidenceReady, L"Terminal.dll: VT upgrade corpus behavior and 200-sample measurement pass", success);
    const bool archiveWritten = evidenceReady && WriteTerminalVtUpgradeArchive(outputDirectory, evidence);
    Check(archiveWritten, L"Terminal.dll: writes bounded digest-only VT upgrade evidence", success);

    const auto shutdown = reinterpret_cast<PfnPluginShutdown>(
        GetProcAddress(module.get(), "RedSalamanderPluginShutdown"));
    const auto canUnload = reinterpret_cast<PfnPluginCanUnloadNow>(
        GetProcAddress(module.get(), "RedSalamanderPluginCanUnloadNow"));
    if (shutdown != nullptr && canUnload != nullptr)
    {
        shutdown();
        Check(canUnload() == TRUE, L"Terminal.dll: VT upgrade corpus reaches its unload quiet point", success);
    }
    return archiveWritten;
}

} // namespace

int wmain(int argc, wchar_t** argv)
{
    if (argc == 2 && argv != nullptr && argv[1] != nullptr && std::wstring_view(argv[1]) == L"--temporary-file-contract")
    {
        bool focusedSuccess = true;
        std::wcout << L"[ RUN      ] Delete-on-close temporary-file contracts\n";
        TestDeleteOnCloseTemporaryFileContracts(focusedSuccess);
        std::wcout << (focusedSuccess ? L"PluginContractTests focused temporary-file contracts passed.\n"
                                      : L"PluginContractTests focused temporary-file contracts failed.\n");
        return focusedSuccess ? 0 : 1;
    }

    if (argc == 2 && argv != nullptr && argv[1] != nullptr && std::wstring_view(argv[1]) == L"--registry-wsl-catalog")
    {
        bool focusedSuccess = true;
        std::wcout << L"[ RUN      ] Bounded registry and canonical WSL catalog contracts\n";
        TestRegistryAndWslCatalogContracts(focusedSuccess);
        std::wcout << (focusedSuccess ? L"PluginContractTests focused registry/WSL catalog contracts passed.\n"
                                      : L"PluginContractTests focused registry/WSL catalog contracts failed.\n");
        return focusedSuccess ? 0 : 1;
    }

#if defined(RS_ASAN_DEBUG_BUILD)
    if (argc == 2 && argv != nullptr && argv[1] != nullptr && std::wstring_view(argv[1]) == L"--asan-seed-heap-overflow")
    {
        auto storage                = std::make_unique<unsigned char[]>(8u);
        volatile unsigned char* raw = storage.get();
        raw[16]                     = 0x5Au;
        return raw[16] == 0x5Au ? 0 : 1;
    }
#else
    if (argc == 2 && argv != nullptr && argv[1] != nullptr && std::wstring_view(argv[1]) == L"--asan-seed-heap-overflow")
    {
        std::wcerr << L"The seeded heap probe requires the ASan Debug configuration.\n";
        return 2;
    }
#endif

    const bool hasSingleArgument = argc == 2 && argv != nullptr && argv[1] != nullptr;
    const bool packageSmoke = hasSingleArgument && std::wstring_view(argv[1]) == L"--package-smoke";
    const bool terminalSelfTests = hasSingleArgument && std::wstring_view(argv[1]) == L"--terminal-selftests";
    const bool terminalCommandExperiencePerfSelfTests =
        hasSingleArgument && std::wstring_view(argv[1]) == L"--terminal-command-experience-perf";
    const bool terminalVtUpgradeCorpus = argc == 3 && argv != nullptr && argv[1] != nullptr && argv[2] != nullptr &&
        std::wstring_view(argv[1]) == L"--terminal-vt-upgrade-corpus" && std::wstring_view(argv[2]).size() != 0u;
    const bool s3DirectoryTransferSelfTests = hasSingleArgument && std::wstring_view(argv[1]) == L"--s3-directory-transfer-selftests";
    const bool s3R0bDeleteSelfTests = hasSingleArgument && std::wstring_view(argv[1]) == L"--s3-r0b-delete-selftests";
    const bool s3R0fContainmentSelfTests = hasSingleArgument && std::wstring_view(argv[1]) == L"--s3-r0f-containment-selftests";
    const bool graphR0fContainmentSelfTests = hasSingleArgument && std::wstring_view(argv[1]) == L"--microsoft-drive-r0f-containment-selftests";
    const bool googleDriveR0fContainmentSelfTests = hasSingleArgument && std::wstring_view(argv[1]) == L"--google-drive-r0f-containment-selftests";
    const bool curlSelfTests = hasSingleArgument && std::wstring_view(argv[1]) == L"--curl-selftests";

    // Configure the DLL search path so that plugin DLLs find their transitive dependencies:
    // - Plugins\ contains aws-*.dll, libcurl-d.dll, sqlite3.dll, and so on.
    // - The exe dir (parent of Plugins\) contains Common.dll and other shared DLLs.
    // We use SetDllDirectoryW to add Plugins\ to the standard DLL search path (it is inserted
    // between the app dir and the system/PATH directories in the standard search order).
    const std::wstring exeDir     = GetExeDir();
    const std::wstring pluginsDir = exeDir + L"Plugins";
    SetDllDirectoryW(pluginsDir.c_str());
    // AddDllDirectory is also called for LOAD_LIBRARY_SEARCH_USER_DIRS coverage.
    const DLL_DIRECTORY_COOKIE pluginsDirCookie = AddDllDirectory(pluginsDir.c_str());

    if (terminalVtUpgradeCorpus)
    {
        bool focusedSuccess = true;
        std::wcout << L"[ RUN      ] Terminal VT upgrade deterministic corpus and Release performance baseline\n";
        static_cast<void>(RunTerminalVtUpgradeCorpus(std::filesystem::path(argv[2]), focusedSuccess));
        if (pluginsDirCookie != nullptr)
        {
            RemoveDllDirectory(pluginsDirCookie);
        }
        std::wcout << (focusedSuccess ? L"PluginContractTests focused Terminal VT upgrade corpus passed.\n"
                                      : L"PluginContractTests focused Terminal VT upgrade corpus failed.\n");
        return focusedSuccess ? 0 : 1;
    }

    if (s3DirectoryTransferSelfTests)
    {
        bool focusedSuccess = true;
        std::wcout << L"[ RUN      ] S3 CopySource and directory-transfer selftests\n";
        TestPluginDebugSelfTests(L"Plugins\\FileSystemS3.dll",
                                 "RedSalamanderS3DirectoryTransferProbeSelfTests",
                                 L"FileSystemS3.dll CopySource and directory-transfer",
                                 true,
                                 focusedSuccess);
        if (pluginsDirCookie != nullptr)
        {
            RemoveDllDirectory(pluginsDirCookie);
        }
        std::wcout << (focusedSuccess ? L"PluginContractTests focused S3 CopySource/directory-transfer selftests passed.\n"
                                      : L"PluginContractTests focused S3 CopySource/directory-transfer selftests failed.\n");
        return focusedSuccess ? 0 : 1;
    }

    if (s3R0bDeleteSelfTests)
    {
        bool focusedSuccess = true;
        std::wcout << L"[ RUN      ] S3 R0b exact-generation virtual-folder Delete selftests\n";
        TestPluginDebugSelfTests(L"Plugins\\FileSystemS3.dll",
                                 "RedSalamanderS3R0bDeleteSelfTests",
                                 L"FileSystemS3.dll R0b exact-generation Delete",
                                 true,
                                 focusedSuccess);
        if (pluginsDirCookie != nullptr)
        {
            RemoveDllDirectory(pluginsDirCookie);
        }
        std::wcout << (focusedSuccess ? L"PluginContractTests focused S3 R0b Delete selftests passed.\n"
                                      : L"PluginContractTests focused S3 R0b Delete selftests failed.\n");
        return focusedSuccess ? 0 : 1;
    }

    if (s3R0fContainmentSelfTests)
    {
        bool focusedSuccess = true;
        std::wcout << L"[ RUN      ] S3 R0f containment selftests (fake S3 endpoint: streaming cancel, stalled request, provider bound)\n";
        TestPluginDebugSelfTests(L"Plugins\\FileSystemS3.dll",
                                 "RedSalamanderS3R0fContainmentSelfTests",
                                 L"FileSystemS3.dll R0f containment",
                                 true,
                                 focusedSuccess);
        if (pluginsDirCookie != nullptr)
        {
            RemoveDllDirectory(pluginsDirCookie);
        }
        std::wcout << (focusedSuccess ? L"PluginContractTests focused S3 R0f containment selftests passed.\n"
                                      : L"PluginContractTests focused S3 R0f containment selftests failed.\n");
        return focusedSuccess ? 0 : 1;
    }

    if (graphR0fContainmentSelfTests)
    {
        bool focusedSuccess = true;
        std::wcout << L"[ RUN      ] Microsoft Drive R0f containment selftests (fake Graph endpoint: streaming cancel, stalled request, provider bound)\n";
        TestPluginDebugSelfTests(L"Plugins\\FileSystemMicrosoftDrive.dll",
                                 "RedSalamanderMicrosoftDriveR0fContainmentSelfTests",
                                 L"FileSystemMicrosoftDrive.dll R0f containment",
                                 true,
                                 focusedSuccess);
        if (pluginsDirCookie != nullptr)
        {
            RemoveDllDirectory(pluginsDirCookie);
        }
        std::wcout << (focusedSuccess ? L"PluginContractTests focused Microsoft Drive R0f containment selftests passed.\n"
                                      : L"PluginContractTests focused Microsoft Drive R0f containment selftests failed.\n");
        return focusedSuccess ? 0 : 1;
    }

    if (googleDriveR0fContainmentSelfTests)
    {
        bool focusedSuccess = true;
        std::wcout << L"[ RUN      ] Google Drive R0f containment selftests (fake Drive endpoint: streaming cancel, stalled request, provider bound)\n";
        TestPluginDebugSelfTests(L"Plugins\\FileSystemGoogleDrive.dll",
                                 "RedSalamanderGoogleDriveR0fContainmentSelfTests",
                                 L"FileSystemGoogleDrive.dll R0f containment",
                                 true,
                                 focusedSuccess);
        if (pluginsDirCookie != nullptr)
        {
            RemoveDllDirectory(pluginsDirCookie);
        }
        std::wcout << (focusedSuccess ? L"PluginContractTests focused Google Drive R0f containment selftests passed.\n"
                                      : L"PluginContractTests focused Google Drive R0f containment selftests failed.\n");
        return focusedSuccess ? 0 : 1;
    }

    if (curlSelfTests)
    {
        bool focusedSuccess = true;
        std::wcout << L"[ RUN      ] Curl transfer and reader integrity selftests\n";
        TestPluginDebugSelfTests(L"Plugins\\FileSystemCurl.dll",
                                 "RedSalamanderCurlDebugSelfTests",
                                 L"FileSystemCurl.dll",
                                 true,
                                 focusedSuccess);
        if (pluginsDirCookie != nullptr)
        {
            RemoveDllDirectory(pluginsDirCookie);
        }
        std::wcout << (focusedSuccess ? L"PluginContractTests focused Curl selftests passed.\n"
                                      : L"PluginContractTests focused Curl selftests failed.\n");
        return focusedSuccess ? 0 : 1;
    }

    if (terminalSelfTests)
    {
        bool focusedSuccess = true;
        std::wcout << L"[ RUN      ] Terminal live localization contracts\n";
        TestTerminalLocalizedContracts(focusedSuccess);
        std::wcout << L"[ RUN      ] Terminal action ABI and sized-record contracts\n";
        TestTerminalSizedRecords(focusedSuccess);
        std::wcout << L"[ RUN      ] Terminal plugin deterministic selftests\n";
        TestPluginDebugSelfTests(L"Plugins\\Terminal.dll",
                                 "RedSalamanderTerminalDebugSelfTests",
                                 L"Terminal.dll",
                                 true,
                                 focusedSuccess);
        if (pluginsDirCookie != nullptr)
        {
            RemoveDllDirectory(pluginsDirCookie);
        }
        std::wcout << (focusedSuccess ? L"PluginContractTests focused Terminal selftests passed.\n"
                                      : L"PluginContractTests focused Terminal selftests failed.\n");
        return focusedSuccess ? 0 : 1;
    }

    if (terminalCommandExperiencePerfSelfTests)
    {
        bool focusedSuccess = true;
        std::wcout << L"[ RUN      ] Terminal command-experience Release performance selftests\n";
        TestPluginDebugSelfTests(L"Plugins\\Terminal.dll",
                                 "RedSalamanderTerminalDebugCommandExperiencePerfSelfTests",
                                 L"Terminal.dll command experience performance",
                                 true,
                                 focusedSuccess);
        if (pluginsDirCookie != nullptr)
        {
            RemoveDllDirectory(pluginsDirCookie);
        }
        std::wcout << (focusedSuccess ? L"PluginContractTests focused Terminal command-experience performance passed.\n"
                                      : L"PluginContractTests focused Terminal command-experience performance failed.\n");
        return focusedSuccess ? 0 : 1;
    }

    bool success = true;

    if (packageSmoke)
    {
        std::wcout << L"[ RUN      ] Step0: Native packaged Ghostty runtime load/free\n";
        TestPackagedGhosttyRuntimeLoad(success);
    }

    std::wcout << L"[ RUN      ] Step1: Enumerated metadata count contract\n";
    Check(! IsValidEnumeratedPluginCount(0u), L"enumerated metadata count rejects zero", success);
    Check(IsValidEnumeratedPluginCount(1u), L"enumerated metadata count accepts one", success);
    Check(IsValidEnumeratedPluginCount(kMaxEnumeratedPluginsPerModule), L"enumerated metadata count accepts the documented maximum", success);
    Check(! IsValidEnumeratedPluginCount(kMaxEnumeratedPluginsPerModule + 1u), L"enumerated metadata count rejects an implausible range", success);
    PluginMetaData metadata{};
    Check(! IsValidEnumeratedPluginRange(nullptr, 1u), L"enumerated metadata range rejects a null pointer with a nonzero count", success);
    Check(! IsValidEnumeratedPluginRange(&metadata, 0u), L"enumerated metadata range rejects a nonnull pointer with a zero count", success);
    Check(IsValidEnumeratedPluginRange(&metadata, 1u), L"enumerated metadata range accepts a nonnull pointer with a valid count", success);

    const PluginFactoryEntry validFactoryEntries[] = {
        {&GetFactoryContractMetaData0, &GetEmptyFactoryContractSchema, &CreateFactoryContractInstance},
        {&GetFactoryContractMetaData1, &GetEmptyFactoryContractSchema, &CreateFactoryContractInstance},
    };
    const PluginMetaData* enumeratedMetaData = nullptr;
    unsigned int enumeratedCount             = 0u;
    Check(FactoryEnumeratePlugins<IViewer>(validFactoryEntries, __uuidof(IViewer), &enumeratedMetaData, &enumeratedCount) == S_OK &&
              enumeratedMetaData == g_factoryContractMetaData.data() &&
              enumeratedCount == static_cast<unsigned int>(std::size(validFactoryEntries)),
          L"factory enumeration accepts a valid contiguous bounded metadata array",
          success);

    const PluginFactoryEntry nonContiguousEntries[] = {
        {&GetFactoryContractMetaData0, &GetEmptyFactoryContractSchema, &CreateFactoryContractInstance},
        {&GetNonContiguousFactoryMetaData, &GetEmptyFactoryContractSchema, &CreateFactoryContractInstance},
    };
    Check(FactoryEnumeratePlugins<IViewer>(nonContiguousEntries, __uuidof(IViewer), &enumeratedMetaData, &enumeratedCount) ==
              HRESULT_FROM_WIN32(ERROR_INVALID_DATA),
          L"factory enumeration rejects non-contiguous producer metadata",
          success);

    const PluginFactoryEntry nullMetadataEntry[] = {{nullptr, &GetEmptyFactoryContractSchema, &CreateFactoryContractInstance}};
    Check(FactoryEnumeratePlugins<IViewer>(nullMetadataEntry, __uuidof(IViewer), &enumeratedMetaData, &enumeratedCount) ==
              HRESULT_FROM_WIN32(ERROR_INVALID_DATA),
          L"factory enumeration rejects a null producer metadata callback",
          success);

    std::array<PluginFactoryEntry, kMaxEnumeratedPluginsPerModule + 1u> excessiveFactoryEntries{};
    excessiveFactoryEntries.fill(validFactoryEntries[0]);
    Check(FactoryEnumeratePlugins<IViewer>(excessiveFactoryEntries, __uuidof(IViewer), &enumeratedMetaData, &enumeratedCount) ==
              HRESULT_FROM_WIN32(ERROR_TOO_MANY_NAMES),
          L"factory enumeration rejects a producer count above the ABI maximum",
          success);

    std::wcout << L"[ RUN      ] Step1b: Packed FileInfo owner contracts\n";
    TestPackedFileInfoBuffer(success);

    std::wcout << L"[ RUN      ] Step1c: Bounded registry and canonical WSL catalog contracts\n";
    TestRegistryAndWslCatalogContracts(success);

    std::wcout << L"[ RUN      ] Step2: EnumerateAndSchema - filesystem plugins\n";
    success = RunEnumerateAndSchemaPass(kFilesystemDlls, __uuidof(IFileSystem)) && success;

    std::wcout << L"[ RUN      ] Step2: EnumerateAndSchema - viewer plugins\n";
    success = RunEnumerateAndSchemaPass(kViewerDlls, __uuidof(IViewer)) && success;

    std::wcout << L"[ RUN      ] Step2: EnumerateAndSchema - terminal plugins\n";
    success = RunEnumerateAndSchemaPass(kTerminalDlls, __uuidof(ITerminal)) && success;

    std::wcout << L"[ RUN      ] Step2b: sizeBytes-only viewer and terminal records\n";
    for (const std::wstring_view viewerDll : kViewerDlls)
    {
        TestViewerSizedRecords(viewerDll, success);
    }
    TestTerminalSizedRecords(success, ! packageSmoke);

    std::wcout << L"[ RUN      ] Step2c: Terminal live localization and diagnostic contracts\n";
    TestTerminalLocalizedContracts(success, ! packageSmoke);

    std::wcout << L"[ RUN      ] Step3: typed route validator and provider capabilities\n";
    TestTypedRouteValidatorContracts(success);
    success = RunCapabilitiesPass(kFilesystemDlls, ! packageSmoke) && success;

    std::wcout << L"[ RUN      ] Step4: Negative bogus plugin id\n";
    {
        bool negSuccess = true;
        TestBogusPluginId(negSuccess);
        success = negSuccess && success;
    }

    if (! packageSmoke)
    {
        std::wcout << L"[ RUN      ] Step5: Configuration-gated plugin debug selftests\n";
        {
            bool debugSuccess = true;
            Check(! IsDebugSelfTestExportPresenceAccepted(false, true),
                  L"missing required debug selftest export turns the step red",
                  debugSuccess);
            Check(IsDebugSelfTestExportPresenceAccepted(false, false),
                  L"missing configuration-gated debug selftest export is an explicit skip",
                  debugSuccess);

            TestPluginDebugSelfTests(L"Plugins\\FileSystem.dll",
                                     "RedSalamanderFileSystemDebugSelfTests",
                                     L"FileSystem.dll",
                                     kDebugSelfTestExportsRequired,
                                     debugSuccess);
            TestPluginDebugSelfTests(L"Plugins\\FileSystem7z.dll",
                                     "RedSalamander7zDebugSelfTests",
                                     L"FileSystem7z.dll",
                                     kDebugSelfTestExportsRequired,
                                     debugSuccess);
            TestPluginDebugSelfTests(L"Plugins\\FileSystemMicrosoftDrive.dll",
                                     "RedSalamanderMicrosoftDriveDebugSelfTests",
                                     L"FileSystemMicrosoftDrive.dll",
                                     kDebugSelfTestExportsRequired,
                                     debugSuccess);
            TestPluginDebugSelfTests(L"Plugins\\FileSystemGoogleDrive.dll",
                                     "RedSalamanderGoogleDriveDebugSelfTests",
                                     L"FileSystemGoogleDrive.dll",
                                     kDebugSelfTestExportsRequired,
                                     debugSuccess);
            TestPluginDebugSelfTests(L"Plugins\\FileSystemS3.dll",
                                     "RedSalamanderS3DebugSelfTests",
                                     L"FileSystemS3.dll",
                                     kDebugSelfTestExportsRequired,
                                     debugSuccess);
            TestPluginDebugSelfTests(L"Plugins\\FileSystemCurl.dll",
                                     "RedSalamanderCurlDebugSelfTests",
                                     L"FileSystemCurl.dll",
                                     kDebugSelfTestExportsRequired,
                                     debugSuccess);
            TestPluginDebugSelfTests(L"Plugins\\Terminal.dll",
                                     "RedSalamanderTerminalDebugSelfTests",
                                     L"Terminal.dll",
                                     kDebugSelfTestExportsRequired,
                                     debugSuccess);
            success = debugSuccess && success;
        }

#if defined(ENABLE_TESTS)
        std::wcout << L"[ RUN      ] Step5b: S3 multipart performance and teardown selftests\n";
        {
            bool multipartSuccess = true;
            TestPluginDebugSelfTests(L"Plugins\\FileSystemS3.dll",
                                     "RedSalamanderS3MultipartSelfTests",
                                     L"FileSystemS3.dll multipart",
                                     true,
                                     multipartSuccess);
            success = multipartSuccess && success;
        }

        std::wcout << L"[ RUN      ] Step5c: S3 directory-transfer destination-probe selftests\n";
        {
            bool directoryTransferSuccess = true;
            TestPluginDebugSelfTests(L"Plugins\\FileSystemS3.dll",
                                     "RedSalamanderS3DirectoryTransferProbeSelfTests",
                                     L"FileSystemS3.dll directory-transfer probes",
                                     true,
                                     directoryTransferSuccess);
            success = directoryTransferSuccess && success;
        }
#endif

        std::wcout << L"[ RUN      ] Step6: S3 runtime-refresh unload quiet point\n";
        TestS3RuntimeUnloadContract(success);

#if defined(_DEBUG)
        std::wcout << L"[ RUN      ] Step7: Cross-plugin libcurl runtime-refresh survivor proof\n";
        TestCrossPluginCurlRuntimeUnloadContract(success);
#else
        std::wcout << L"[  SKIPPED ] Step7: Debug curl runtime probes are absent from Release plugins\n";
#endif
    }

    if (pluginsDirCookie != nullptr)
    {
        RemoveDllDirectory(pluginsDirCookie);
    }

    std::wcout << (success ? (packageSmoke ? L"PluginContractTests package smoke passed.\n" : L"PluginContractTests passed.\n")
                            : L"PluginContractTests FAILED.\n");
    return success ? 0 : 1;
}
