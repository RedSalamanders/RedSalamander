#pragma once

#include "PlugInterfaces/FileSystem.h"

#include <algorithm>
#include <cstring>
#include <cwctype>
#include <limits>
#include <string>
#include <string_view>

// Provider-side implementation helper for IFileSystemRouteCapabilities. Providers
// build one typed descriptor directly from their own state; this helper performs the
// ABI copy, peer, child-name, collision-key, and join mechanics without consulting
// capability JSON.
struct FileSystemRouteDescriptor final
{
    std::wstring providerId;
    std::wstring pathProfileId;
    std::wstring rootId;
    std::wstring acceptedSeparators = L"/";
    std::wstring forbiddenChildCharacters;

    FileSystemRouteAvailability availability               = FILESYSTEM_ROUTE_UNSUPPORTED;
    FileSystemCancellationRoute cancellationRoute          = FILESYSTEM_CANCELLATION_UNCONTAINED;
    FileSystemNamespaceKind namespaceKind                  = FILESYSTEM_NAMESPACE_PROVIDER_VIRTUAL_FOLDER;
    FileSystemRouteComponentComparison componentComparison = FILESYSTEM_ROUTE_COMPONENT_ORDINAL_CASE_SENSITIVE;
    FileSystemRouteNormalization normalization             = FILESYSTEM_ROUTE_NORMALIZATION_NONE;
    FileSystemRouteCaseOnlyRename caseOnlyRename           = FILESYSTEM_ROUTE_CASE_ONLY_NOT_APPLICABLE;
    uint32_t proofFlags                                    = FILESYSTEM_ROUTE_PROOF_NONE;
    uint32_t providerWatchdogTimeoutMs                     = 0u;
    uint32_t copyMoveMaxConcurrency                        = 1u;
    uint32_t deleteMaxConcurrency                          = 1u;
    uint32_t deleteRecycleBinMaxConcurrency                = 1u;
    uint64_t maxComponentUtf16                             = 255u;

    bool copyOperation            = false;
    bool moveOperation            = false;
    bool nativeMoveOperation      = false;
    bool deleteOperation          = false;
    bool renameOperation          = false;
    bool createDirectoryOperation = false;
    bool propertiesOperation      = false;
    bool readOperation            = false;
    bool writeOperation           = false;
    bool recycleOperation         = false;
    bool boundDelete              = false;
    bool conditionalDelete        = false;
    bool exclusiveStage           = false;
    bool conditionalPublish       = false;
    bool committedSize            = false;
    bool preserveFileLink         = false;
    bool preserveDirectoryLink    = false;
    bool retargetInTree           = false;
    bool exactLinkRemoval         = false;
    bool cancellationAbort        = false;
    bool cancellationDeadline     = false;
    bool pathTextStableIdentity   = true;
    bool casePreserving           = true;
    bool windowsChildNames        = false;
    bool exportCopyAll            = false;
    bool exportMoveAll            = false;
    bool importCopyAll            = false;
    bool importMoveAll            = false;
    wchar_t preferredSeparator    = L'/';
};

class FileSystemRouteCapabilitiesBase : public IFileSystemRouteCapabilities
{
public:
    HRESULT STDMETHODCALLTYPE GetRouteFacts(const wchar_t* path,
                                            FileSystemOperation operation,
                                            FileSystemArena* arena,
                                            FileSystemRouteFacts* facts) noexcept override
    {
        if (! IsValidRequest(path, operation) || arena == nullptr || facts == nullptr)
        {
            return E_INVALIDARG;
        }
        if (facts->sizeBytes != sizeof(FileSystemRouteFacts))
        {
            return E_INVALIDARG;
        }
        if (arena->buffer != nullptr && reinterpret_cast<uintptr_t>(arena->buffer) % alignof(wchar_t) != 0u)
        {
            return E_INVALIDARG;
        }

        FileSystemRouteDescriptor descriptor{};
        const HRESULT descriptorHr = BuildFileSystemRouteDescriptor(path, operation, descriptor);
        if (FAILED(descriptorHr))
        {
            return descriptorHr;
        }
        if (! IsDescriptorValid(descriptor))
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        const uint64_t requiredBytes64 =
            WideBytes(descriptor.acceptedSeparators) + WideBytes(descriptor.providerId) + WideBytes(descriptor.pathProfileId) + WideBytes(descriptor.rootId);
        if (requiredBytes64 == 0u || requiredBytes64 > (std::numeric_limits<uint32_t>::max)())
        {
            return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
        }
        facts->requiredArenaBytes = static_cast<uint32_t>(requiredBytes64);
        if (arena->buffer == nullptr || arena->usedBytes > arena->capacityBytes ||
            requiredBytes64 > static_cast<uint64_t>(arena->capacityBytes - arena->usedBytes))
        {
            return HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER);
        }

        const wchar_t* acceptedSeparators = CopyToArena(descriptor.acceptedSeparators, arena);
        const wchar_t* providerId         = CopyToArena(descriptor.providerId, arena);
        const wchar_t* pathProfileId      = CopyToArena(descriptor.pathProfileId, arena);
        const wchar_t* rootId             = CopyToArena(descriptor.rootId, arena);
        if (acceptedSeparators == nullptr || providerId == nullptr || pathProfileId == nullptr || rootId == nullptr)
        {
            return E_OUTOFMEMORY;
        }

        facts->availability                   = descriptor.availability;
        facts->cancellationRoute              = descriptor.cancellationRoute;
        facts->namespaceKind                  = descriptor.namespaceKind;
        facts->componentComparison            = descriptor.componentComparison;
        facts->normalization                  = descriptor.normalization;
        facts->caseOnlyRename                 = descriptor.caseOnlyRename;
        facts->proofFlags                     = descriptor.proofFlags;
        facts->providerWatchdogTimeoutMs      = descriptor.providerWatchdogTimeoutMs;
        facts->copyMoveMaxConcurrency         = descriptor.copyMoveMaxConcurrency;
        facts->deleteMaxConcurrency           = descriptor.deleteMaxConcurrency;
        facts->deleteRecycleBinMaxConcurrency = descriptor.deleteRecycleBinMaxConcurrency;
        facts->maxComponentUtf16              = descriptor.maxComponentUtf16;
        facts->copyOperation                  = ToBool(descriptor.copyOperation);
        facts->moveOperation                  = ToBool(descriptor.moveOperation);
        facts->nativeMoveOperation            = ToBool(descriptor.nativeMoveOperation);
        facts->deleteOperation                = ToBool(descriptor.deleteOperation);
        facts->renameOperation                = ToBool(descriptor.renameOperation);
        facts->createDirectoryOperation       = ToBool(descriptor.createDirectoryOperation);
        facts->propertiesOperation            = ToBool(descriptor.propertiesOperation);
        facts->readOperation                  = ToBool(descriptor.readOperation);
        facts->writeOperation                 = ToBool(descriptor.writeOperation);
        facts->recycleOperation               = ToBool(descriptor.recycleOperation);
        facts->boundDelete                    = ToBool(descriptor.boundDelete);
        facts->conditionalDelete              = ToBool(descriptor.conditionalDelete);
        facts->exclusiveStage                 = ToBool(descriptor.exclusiveStage);
        facts->conditionalPublish             = ToBool(descriptor.conditionalPublish);
        facts->committedSize                  = ToBool(descriptor.committedSize);
        facts->preserveFileLink               = ToBool(descriptor.preserveFileLink);
        facts->preserveDirectoryLink          = ToBool(descriptor.preserveDirectoryLink);
        facts->retargetInTree                 = ToBool(descriptor.retargetInTree);
        facts->exactLinkRemoval               = ToBool(descriptor.exactLinkRemoval);
        facts->cancellationAbort              = ToBool(descriptor.cancellationAbort);
        facts->cancellationDeadline           = ToBool(descriptor.cancellationDeadline);
        facts->pathTextStableIdentity         = ToBool(descriptor.pathTextStableIdentity);
        facts->casePreserving                 = ToBool(descriptor.casePreserving);
        facts->preferredSeparator             = descriptor.preferredSeparator;
        facts->acceptedSeparators             = acceptedSeparators;
        facts->providerId                     = providerId;
        facts->pathProfileId                  = pathProfileId;
        facts->rootId                         = rootId;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE IsTransferPeerAllowed(
        const wchar_t* path, FileSystemOperation operation, FileSystemTransferPeerRole role, const wchar_t* peerPluginId, BOOL* allowed) noexcept override
    {
        if (allowed == nullptr)
        {
            return E_POINTER;
        }
        *allowed = FALSE;
        if (! IsValidRequest(path, operation) || (operation != FILESYSTEM_COPY && operation != FILESYSTEM_MOVE) ||
            (role != FILESYSTEM_TRANSFER_PEER_EXPORT && role != FILESYSTEM_TRANSFER_PEER_IMPORT) || peerPluginId == nullptr || peerPluginId[0] == L'\0')
        {
            return E_INVALIDARG;
        }

        FileSystemRouteDescriptor descriptor{};
        const HRESULT descriptorHr = BuildFileSystemRouteDescriptor(path, operation, descriptor);
        if (FAILED(descriptorHr))
        {
            return descriptorHr;
        }
        if (! IsDescriptorValid(descriptor))
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }
        const bool isCopy  = operation == FILESYSTEM_COPY;
        const bool permits = role == FILESYSTEM_TRANSFER_PEER_EXPORT ? (isCopy ? descriptor.exportCopyAll : descriptor.exportMoveAll)
                                                                     : (isCopy ? descriptor.importCopyAll : descriptor.importMoveAll);
        *allowed           = ToBool(permits);
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE ValidateChildName(const wchar_t* parentPath,
                                                const wchar_t* childName,
                                                FileSystemOperation operation,
                                                FileSystemChildNameValidation* validation) noexcept override
    {
        if (validation == nullptr)
        {
            return E_POINTER;
        }
        if (validation->sizeBytes != sizeof(FileSystemChildNameValidation) || ! IsValidRequest(parentPath, operation) || childName == nullptr)
        {
            return E_INVALIDARG;
        }

        FileSystemRouteDescriptor descriptor{};
        const HRESULT descriptorHr = BuildFileSystemRouteDescriptor(parentPath, operation, descriptor);
        if (FAILED(descriptorHr))
        {
            return descriptorHr;
        }
        if (! IsDescriptorValid(descriptor))
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        const HRESULT nameHr      = ValidateName(descriptor, childName);
        validation->status        = SUCCEEDED(nameHr) ? FILESYSTEM_CHILD_NAME_VALID : FILESYSTEM_CHILD_NAME_INVALID;
        validation->failureStatus = nameHr;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetChildNameCollisionKey(const wchar_t* parentPath,
                                                       const wchar_t* childName,
                                                       FileSystemOperation operation,
                                                       FileSystemArena* arena,
                                                       const wchar_t** key,
                                                       unsigned long* requiredArenaBytes) noexcept override
    {
        if (key == nullptr || requiredArenaBytes == nullptr)
        {
            return E_POINTER;
        }
        *key                = nullptr;
        *requiredArenaBytes = 0u;
        if (arena == nullptr || (arena->buffer != nullptr && reinterpret_cast<uintptr_t>(arena->buffer) % alignof(wchar_t) != 0u) ||
            ! IsValidRequest(parentPath, operation) || childName == nullptr)
        {
            return E_INVALIDARG;
        }

        FileSystemRouteDescriptor descriptor{};
        const HRESULT descriptorHr = BuildFileSystemRouteDescriptor(parentPath, operation, descriptor);
        if (FAILED(descriptorHr))
        {
            return descriptorHr;
        }
        if (! IsDescriptorValid(descriptor))
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }
        const HRESULT nameHr = ValidateNameShape(descriptor, childName);
        if (FAILED(nameHr))
        {
            return nameHr;
        }

        std::wstring collisionKey(childName);
        if (descriptor.componentComparison == FILESYSTEM_ROUTE_COMPONENT_ORDINAL_IGNORE_CASE && ! collisionKey.empty())
        {
            if (collisionKey.size() > static_cast<size_t>((std::numeric_limits<int>::max)()))
            {
                return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
            }
            const int length   = static_cast<int>(collisionKey.size());
            const int required = LCMapStringEx(LOCALE_NAME_INVARIANT, LCMAP_LOWERCASE, collisionKey.data(), length, nullptr, 0, nullptr, nullptr, 0);
            if (required <= 0)
            {
                return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            }
            std::wstring folded(static_cast<size_t>(required), L'\0');
            const int written =
                LCMapStringEx(LOCALE_NAME_INVARIANT, LCMAP_LOWERCASE, collisionKey.data(), length, folded.data(), required, nullptr, nullptr, 0);
            if (written != required)
            {
                return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            }
            folded.resize(static_cast<size_t>(written));
            collisionKey = std::move(folded);
        }
        return CopyOutputString(collisionKey, arena, key, requiredArenaBytes);
    }

    HRESULT STDMETHODCALLTYPE JoinPath(const wchar_t* parentPath,
                                       const wchar_t* childName,
                                       FileSystemOperation operation,
                                       FileSystemArena* arena,
                                       const wchar_t** joinedPath,
                                       unsigned long* requiredArenaBytes) noexcept override
    {
        if (joinedPath == nullptr || requiredArenaBytes == nullptr)
        {
            return E_POINTER;
        }
        *joinedPath         = nullptr;
        *requiredArenaBytes = 0u;
        if (arena == nullptr || (arena->buffer != nullptr && reinterpret_cast<uintptr_t>(arena->buffer) % alignof(wchar_t) != 0u) ||
            ! IsValidRequest(parentPath, operation) || childName == nullptr)
        {
            return E_INVALIDARG;
        }

        FileSystemRouteDescriptor descriptor{};
        const HRESULT descriptorHr = BuildFileSystemRouteDescriptor(parentPath, operation, descriptor);
        if (FAILED(descriptorHr))
        {
            return descriptorHr;
        }
        if (! IsDescriptorValid(descriptor))
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }
        const HRESULT nameHr = ValidateNameShape(descriptor, childName);
        if (FAILED(nameHr))
        {
            return nameHr;
        }

        std::wstring joined(parentPath);
        if (! joined.empty() && descriptor.acceptedSeparators.find(joined.back()) == std::wstring::npos)
        {
            joined.push_back(descriptor.preferredSeparator);
        }
        joined.append(childName);
        return CopyOutputString(joined, arena, joinedPath, requiredArenaBytes);
    }

protected:
    virtual ~FileSystemRouteCapabilitiesBase() = default;

    virtual HRESULT BuildFileSystemRouteDescriptor(const wchar_t* path, FileSystemOperation operation, FileSystemRouteDescriptor& descriptor) noexcept = 0;

private:
    [[nodiscard]] static bool IsValidRequest(const wchar_t* path, FileSystemOperation operation) noexcept
    {
        return path != nullptr && path[0] != L'\0' && operation <= FILESYSTEM_CREATE_DIRECTORY;
    }

    [[nodiscard]] static BOOL ToBool(bool value) noexcept
    {
        return value ? TRUE : FALSE;
    }

    [[nodiscard]] static uint64_t WideBytes(std::wstring_view text) noexcept
    {
        return (static_cast<uint64_t>(text.size()) + 1u) * sizeof(wchar_t);
    }

    [[nodiscard]] static bool IsDescriptorValid(const FileSystemRouteDescriptor& descriptor) noexcept
    {
        const bool availabilityValid = descriptor.availability == FILESYSTEM_ROUTE_UNSUPPORTED || descriptor.availability == FILESYSTEM_ROUTE_AVAILABLE;
        const bool routeValid        = descriptor.cancellationRoute <= FILESYSTEM_CANCELLATION_PROVIDER_WATCHDOG;
        const bool namespaceValid =
            descriptor.namespaceKind >= FILESYSTEM_NAMESPACE_REAL_CONTAINER && descriptor.namespaceKind <= FILESYSTEM_NAMESPACE_FIXED_OBJECT_SET;
        const bool comparisonValid    = descriptor.componentComparison == FILESYSTEM_ROUTE_COMPONENT_ORDINAL_IGNORE_CASE ||
                                        descriptor.componentComparison == FILESYSTEM_ROUTE_COMPONENT_ORDINAL_CASE_SENSITIVE;
        const bool normalizationValid = descriptor.normalization == FILESYSTEM_ROUTE_NORMALIZATION_NONE;
        const bool caseRenameValid =
            descriptor.caseOnlyRename >= FILESYSTEM_ROUTE_CASE_ONLY_SUPPORTED && descriptor.caseOnlyRename <= FILESYSTEM_ROUTE_CASE_ONLY_NOT_APPLICABLE;
        constexpr uint32_t knownProofFlags = FILESYSTEM_ROUTE_PROOF_HOST_READBACK | FILESYSTEM_ROUTE_PROOF_PROVIDER_BLAKE3 |
                                             FILESYSTEM_ROUTE_PROOF_PROVIDER_REREAD | FILESYSTEM_ROUTE_PROOF_WRITER_DIGEST;
        const bool cancellationValid =
            (descriptor.cancellationRoute == FILESYSTEM_CANCELLATION_PROVIDER_WATCHDOG && descriptor.providerWatchdogTimeoutMs != 0u) ||
            (descriptor.cancellationRoute != FILESYSTEM_CANCELLATION_PROVIDER_WATCHDOG && descriptor.providerWatchdogTimeoutMs == 0u);
        return ! descriptor.providerId.empty() && ! descriptor.pathProfileId.empty() && ! descriptor.rootId.empty() &&
               ! descriptor.acceptedSeparators.empty() && descriptor.acceptedSeparators.find(descriptor.preferredSeparator) != std::wstring::npos &&
               descriptor.maxComponentUtf16 != 0u && descriptor.copyMoveMaxConcurrency != 0u && descriptor.deleteMaxConcurrency != 0u &&
               descriptor.deleteRecycleBinMaxConcurrency != 0u && availabilityValid && routeValid && namespaceValid && comparisonValid && normalizationValid &&
               caseRenameValid && (descriptor.proofFlags & ~knownProofFlags) == 0u && cancellationValid &&
               (! descriptor.nativeMoveOperation || descriptor.moveOperation) &&
               (! descriptor.retargetInTree || descriptor.preserveFileLink || descriptor.preserveDirectoryLink);
    }

    [[nodiscard]] static const wchar_t* CopyToArena(std::wstring_view text, FileSystemArena* arena) noexcept
    {
        const uint64_t bytes64 = WideBytes(text);
        if (bytes64 > (std::numeric_limits<unsigned long>::max)())
        {
            return nullptr;
        }
        void* storage = AllocateFromFileSystemArena(arena, static_cast<unsigned long>(bytes64), alignof(wchar_t));
        if (storage == nullptr)
        {
            return nullptr;
        }
        wchar_t* output = static_cast<wchar_t*>(storage);
        if (! text.empty())
        {
            memcpy(output, text.data(), text.size() * sizeof(wchar_t));
        }
        output[text.size()] = L'\0';
        return output;
    }

    // Shape rules every route shares: one non-empty component that is not "." or ".." and
    // contains no accepted separator. Collision keys and joins apply only these rules, so both
    // stay computable for every name the provider can already enumerate; proposed names use
    // ValidateName.
    [[nodiscard]] static HRESULT ValidateNameShape(const FileSystemRouteDescriptor& descriptor, std::wstring_view name) noexcept
    {
        if (name.empty() || name == L"." || name == L".." || name.find_first_of(descriptor.acceptedSeparators) != std::wstring_view::npos)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
        }
        return S_OK;
    }

    [[nodiscard]] static bool IsReservedWindowsLeaf(std::wstring_view name) noexcept
    {
        const size_t dot             = name.find(L'.');
        const std::wstring_view stem = name.substr(0u, dot);
        if (stem.empty() || stem.size() > 7u)
        {
            return false;
        }
        std::wstring folded(stem);
        for (wchar_t& ch : folded)
        {
            ch = static_cast<wchar_t>(towupper(ch));
        }
        if (folded == L"CON" || folded == L"PRN" || folded == L"AUX" || folded == L"NUL" || folded == L"CONIN$" || folded == L"CONOUT$")
        {
            return true;
        }
        if (folded.size() != 4u || ! (folded.starts_with(L"COM") || folded.starts_with(L"LPT")))
        {
            return false;
        }
        const wchar_t digit = folded[3];
        return (digit >= L'0' && digit <= L'9') || digit == L'\u00B9' || digit == L'\u00B2' || digit == L'\u00B3';
    }

    // Proposed-name rules: shape, provider-forbidden characters, the component length bound, and
    // the Windows rules when the route declares them. The distinct HRESULTs let a preview explain
    // which rule failed without re-deriving provider policy.
    [[nodiscard]] static HRESULT ValidateName(const FileSystemRouteDescriptor& descriptor, std::wstring_view name) noexcept
    {
        const HRESULT shapeHr = ValidateNameShape(descriptor, name);
        if (FAILED(shapeHr))
        {
            return shapeHr;
        }
        if (name.size() > descriptor.maxComponentUtf16)
        {
            return HRESULT_FROM_WIN32(ERROR_FILENAME_EXCED_RANGE);
        }
        for (const wchar_t ch : name)
        {
            if (descriptor.forbiddenChildCharacters.find(ch) != std::wstring::npos ||
                (descriptor.windowsChildNames && (ch < 32 || std::wstring_view(L"<>:\"/\\|?*").find(ch) != std::wstring_view::npos)))
            {
                return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
            }
        }
        if (descriptor.windowsChildNames)
        {
            if (name.back() == L'.' || name.back() == L' ')
            {
                return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
            }
            if (IsReservedWindowsLeaf(name))
            {
                return HRESULT_FROM_WIN32(ERROR_BAD_DEVICE);
            }
        }
        return S_OK;
    }

    [[nodiscard]] static HRESULT CopyOutputString(std::wstring_view value,
                                                  FileSystemArena* arena,
                                                  const wchar_t** output,
                                                  unsigned long* requiredArenaBytes) noexcept
    {
        const uint64_t bytes64 = WideBytes(value);
        if (bytes64 > (std::numeric_limits<unsigned long>::max)())
        {
            return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
        }
        *requiredArenaBytes = static_cast<unsigned long>(bytes64);
        if (arena->buffer == nullptr || arena->usedBytes > arena->capacityBytes || bytes64 > static_cast<uint64_t>(arena->capacityBytes - arena->usedBytes))
        {
            return HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER);
        }
        *output = CopyToArena(value, arena);
        return *output != nullptr ? S_OK : E_OUTOFMEMORY;
    }
};
