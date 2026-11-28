#pragma once

#include "HandleIo.h"
#include "PathUtils.h"

#include <array>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <filesystem>
#include <format>
#include <limits>
#include <optional>
#include <span>
#include <string>
#include <string_view>
#include <system_error>
#include <utility>
#include <vector>

#ifdef ENABLE_TESTS
#include <atomic>
#endif

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027) // WIL move-only wrappers intentionally delete copy operations.
#include <wil/resource.h>
#pragma warning(pop)

namespace Common::Files
{
enum class ExistingTargetPolicy : uint8_t
{
    FailIfExists,
    Replace,
};

struct LocalFileVersion
{
    FILE_ID_INFO identity{};
    int64_t writeTime  = 0;
    int64_t changeTime = 0;
    int64_t size       = 0;
    DWORD attributes   = 0;

    [[nodiscard]] bool operator==(const LocalFileVersion& other) const noexcept
    {
        return identity.VolumeSerialNumber == other.identity.VolumeSerialNumber &&
               std::memcmp(identity.FileId.Identifier, other.identity.FileId.Identifier, sizeof(identity.FileId.Identifier)) == 0 &&
               writeTime == other.writeTime && changeTime == other.changeTime && size == other.size && attributes == other.attributes;
    }
};

[[nodiscard]] inline HRESULT QueryLocalFileVersion(HANDLE file, LocalFileVersion& version) noexcept
{
    LocalFileVersion observed{};
    FILE_BASIC_INFO basic{};
    FILE_STANDARD_INFO standard{};
    if (GetFileInformationByHandleEx(file, FileBasicInfo, &basic, sizeof(basic)) == FALSE ||
        GetFileInformationByHandleEx(file, FileStandardInfo, &standard, sizeof(standard)) == FALSE ||
        GetFileInformationByHandleEx(file, FileIdInfo, &observed.identity, sizeof(observed.identity)) == FALSE)
    {
        const DWORD error = GetLastError();
        return HRESULT_FROM_WIN32(error == ERROR_SUCCESS ? ERROR_READ_FAULT : error);
    }
    if (standard.Directory || standard.DeletePending || (basic.FileAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0u)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }
    constexpr FILE_ID_128 unknownIdentity{};
    if (std::memcmp(observed.identity.FileId.Identifier, unknownIdentity.Identifier, sizeof(unknownIdentity.Identifier)) == 0)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    observed.writeTime  = basic.LastWriteTime.QuadPart;
    observed.changeTime = basic.ChangeTime.QuadPart;
    observed.size       = standard.EndOfFile.QuadPart;
    observed.attributes = basic.FileAttributes;
    version             = observed;
    return S_OK;
}

// A successful absent observation grants create-only publication.
[[nodiscard]] inline HRESULT ObserveLocalFileDestination(const std::filesystem::path& path, std::optional<LocalFileVersion>& version) noexcept
{
    version.reset();
    if (path.empty())
    {
        return E_INVALIDARG;
    }
    const auto extendedPath = Common::Paths::ToExtendedWin32Path(path.native());
    wil::unique_hfile file(CreateFileW(extendedPath.c_str(),
                                       FILE_READ_ATTRIBUTES,
                                       FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                       nullptr,
                                       OPEN_EXISTING,
                                       FILE_FLAG_OPEN_REPARSE_POINT | FILE_FLAG_BACKUP_SEMANTICS,
                                       nullptr));
    if (! file)
    {
        const DWORD error = GetLastError();
        return error == ERROR_FILE_NOT_FOUND || error == ERROR_PATH_NOT_FOUND ? S_OK : HRESULT_FROM_WIN32(error == ERROR_SUCCESS ? ERROR_OPEN_FAILED : error);
    }
    LocalFileVersion observed{};
    const HRESULT hr = QueryLocalFileVersion(file.get(), observed);
    if (SUCCEEDED(hr))
    {
        version = observed;
    }
    return hr;
}

// Runs after the staged file is finalized, immediately before each rename attempt.
struct TargetRenameCheck
{
    void* context                                                = nullptr;
    HRESULT (*verify)(void*, FILE_INFO_BY_HANDLE_CLASS) noexcept = nullptr;
};

struct GuardedReplacementOptions
{
    const std::filesystem::path* backupPath                                   = nullptr;
    void* context                                                             = nullptr;
    HRESULT (*prepareBackup)(void*, HANDLE, const LocalFileVersion&) noexcept = nullptr;
    BY_HANDLE_FILE_INFORMATION* committedInformation                          = nullptr;
};

#ifdef ENABLE_TESTS
namespace Testing
{
inline thread_local void (*g_beforeGuardedFilePublish)(const std::filesystem::path&) noexcept = nullptr;

// One-shot deterministic seams for focused output-transaction fault coverage.
// They are compiled only into test-enabled binaries and never alter Release behavior.
inline std::atomic<HRESULT> g_nextWriteFailure{S_OK};
inline std::atomic<HRESULT> g_nextFlushFailure{S_OK};
inline std::atomic<HRESULT> g_nextPublishFailure{S_OK};
inline std::atomic<HRESULT> g_nextRenameExFailure{S_OK};
inline std::atomic<HRESULT> g_nextLegacyRenameFailure{S_OK};

inline void FailNextLocalFileTransactionRenameEx(HRESULT hr) noexcept
{
    g_nextRenameExFailure.store(FAILED(hr) ? hr : E_FAIL, std::memory_order_relaxed);
}

inline void FailNextLocalFileTransactionLegacyRename(HRESULT hr) noexcept
{
    g_nextLegacyRenameFailure.store(FAILED(hr) ? hr : E_FAIL, std::memory_order_relaxed);
}

inline void FailNextLocalFileTransactionWrite(HRESULT hr) noexcept
{
    g_nextWriteFailure.store(FAILED(hr) ? hr : E_FAIL, std::memory_order_relaxed);
}

inline void FailNextLocalFileTransactionFlush(HRESULT hr) noexcept
{
    g_nextFlushFailure.store(FAILED(hr) ? hr : E_FAIL, std::memory_order_relaxed);
}

inline void FailNextLocalFileTransactionPublish(HRESULT hr) noexcept
{
    g_nextPublishFailure.store(FAILED(hr) ? hr : E_FAIL, std::memory_order_relaxed);
}

[[nodiscard]] inline HRESULT TakeNextLocalFileTransactionWriteFailure() noexcept
{
    return g_nextWriteFailure.exchange(S_OK, std::memory_order_relaxed);
}

[[nodiscard]] inline HRESULT TakeNextLocalFileTransactionFlushFailure() noexcept
{
    return g_nextFlushFailure.exchange(S_OK, std::memory_order_relaxed);
}

[[nodiscard]] inline HRESULT TakeNextLocalFileTransactionPublishFailure() noexcept
{
    return g_nextPublishFailure.exchange(S_OK, std::memory_order_relaxed);
}

struct LocalFileTransactionRollbackSnapshot final
{
    bool ownedHandle         = false;
    bool posixDisposition    = false;
    DWORD posixError         = ERROR_SUCCESS;
    bool legacyDisposition   = false;
    DWORD legacyError        = ERROR_SUCCESS;
    bool pathDeleteAttempted = false;
    DWORD pathDeleteError    = ERROR_SUCCESS;
};
inline thread_local LocalFileTransactionRollbackSnapshot g_lastRollbackSnapshot{};

[[nodiscard]] inline LocalFileTransactionRollbackSnapshot LastLocalFileTransactionRollback() noexcept
{
    return g_lastRollbackSnapshot;
}
} // namespace Testing
#endif

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027) // WIL member makes the transaction intentionally move-only.
// Local regular-file output transaction. It exclusively creates a unique sibling,
// accepts complete writes, flushes and optionally verifies the byte count, then
// publishes with same-volume rename/replace semantics. Abandoning or failing the
// transaction deletes only the temporary sibling and preserves the previous target.
class LocalFileTransaction final
{
public:
    LocalFileTransaction() noexcept                              = default;
    LocalFileTransaction(const LocalFileTransaction&)            = delete;
    LocalFileTransaction& operator=(const LocalFileTransaction&) = delete;

    LocalFileTransaction(LocalFileTransaction&& other) noexcept
        : _file(std::move(other._file)),
          _targetPath(std::move(other._targetPath)),
          _temporaryPath(std::move(other._temporaryPath)),
          _policy(other._policy),
          _committed(other._committed)
    {
        other._temporaryPath.clear();
        other._committed = true;
    }

    LocalFileTransaction& operator=(LocalFileTransaction&& other) noexcept
    {
        if (this != &other)
        {
            Abort();
            _file          = std::move(other._file);
            _targetPath    = std::move(other._targetPath);
            _temporaryPath = std::move(other._temporaryPath);
            _policy        = other._policy;
            _committed     = other._committed;
            other._temporaryPath.clear();
            other._committed = true;
        }
        return *this;
    }

    ~LocalFileTransaction() noexcept
    {
        Abort();
    }

    [[nodiscard]] static HRESULT Create(const std::filesystem::path& targetPath,
                                        ExistingTargetPolicy policy,
                                        bool createParentDirectories,
                                        LocalFileTransaction& out,
                                        std::wstring_view temporaryPrefix = L".$rs-transaction-",
                                        bool readStage                    = false) noexcept
    {
        out.Abort();
        if (targetPath.empty())
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
        }

        std::error_code ec;
        std::filesystem::path normalizedTarget = std::filesystem::absolute(targetPath, ec).lexically_normal();
        if (ec || normalizedTarget.empty() || normalizedTarget.parent_path().empty())
        {
            return ec ? HRESULT_FROM_WIN32(static_cast<DWORD>(ec.value())) : HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
        }

        const std::filesystem::path parent = normalizedTarget.parent_path();
        const std::filesystem::path extendedParent(Common::Paths::ToExtendedWin32Path(parent.native()));
        const std::filesystem::path extendedTarget(Common::Paths::ToExtendedWin32Path(normalizedTarget.native()));
        if (createParentDirectories)
        {
            std::filesystem::create_directories(extendedParent, ec);
            if (ec)
            {
                return HRESULT_FROM_WIN32(static_cast<DWORD>(ec.value()));
            }
        }

        const DWORD parentAttributes = GetFileAttributesW(extendedParent.c_str());
        if (parentAttributes == INVALID_FILE_ATTRIBUTES)
        {
            const DWORD error = GetLastError();
            return HRESULT_FROM_WIN32(error == ERROR_SUCCESS ? ERROR_PATH_NOT_FOUND : error);
        }
        if ((parentAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0u)
        {
            return HRESULT_FROM_WIN32(ERROR_DIRECTORY);
        }

        if (policy == ExistingTargetPolicy::FailIfExists)
        {
            const DWORD targetAttributes = GetFileAttributesW(extendedTarget.c_str());
            if (targetAttributes != INVALID_FILE_ATTRIBUTES)
            {
                return HRESULT_FROM_WIN32(ERROR_FILE_EXISTS);
            }
            const DWORD error = GetLastError();
            if (error != ERROR_FILE_NOT_FOUND && error != ERROR_PATH_NOT_FOUND)
            {
                return HRESULT_FROM_WIN32(error == ERROR_SUCCESS ? ERROR_GEN_FAILURE : error);
            }
        }

        Common::Paths::UniqueSiblingFileOptions options{};
        options.prefix             = temporaryPrefix;
        options.suffix             = L".tmp";
        options.desiredAccess      = GENERIC_WRITE | DELETE | (readStage ? GENERIC_READ : 0u);
        options.shareMode          = FILE_SHARE_READ;
        options.flagsAndAttributes = FILE_ATTRIBUTE_NORMAL | FILE_FLAG_SEQUENTIAL_SCAN;

        std::wstring temporaryPath;
        wil::unique_hfile file;
        const HRESULT createHr = Common::Paths::CreateUniqueSiblingFile(extendedTarget.native(), options, temporaryPath, file);
        if (FAILED(createHr))
        {
            return createHr;
        }

        out._file          = std::move(file);
        out._targetPath    = extendedTarget;
        out._temporaryPath = std::filesystem::path(std::move(temporaryPath));
        out._policy        = policy;
        out._committed     = false;
        return S_OK;
    }

    [[nodiscard]] HRESULT Write(std::span<const std::byte> bytes) noexcept
    {
#ifdef ENABLE_TESTS
        const HRESULT injectedFailure = Testing::TakeNextLocalFileTransactionWriteFailure();
        if (FAILED(injectedFailure))
        {
            return injectedFailure;
        }
#endif
        return Common::HandleIo::WriteAll(_file.get(), bytes);
    }

    [[nodiscard]] HRESULT Write(const void* data, size_t byteCount) noexcept
    {
#ifdef ENABLE_TESTS
        const HRESULT injectedFailure = Testing::TakeNextLocalFileTransactionWriteFailure();
        if (FAILED(injectedFailure))
        {
            return injectedFailure;
        }
#endif
        return Common::HandleIo::WriteAll(_file.get(), data, byteCount);
    }

    [[nodiscard]] HRESULT Write(std::string_view bytes) noexcept
    {
        return Write(bytes.data(), bytes.size());
    }

    [[nodiscard]] HRESULT Commit(std::optional<uint64_t> expectedSize                 = std::nullopt,
                                 BY_HANDLE_FILE_INFORMATION* committedFileInformation = nullptr,
                                 wil::unique_hfile* retainedPublishedFile             = nullptr,
                                 TargetRenameCheck targetCheck                        = {}) noexcept
    {
        if (! _file || _temporaryPath.empty() || _targetPath.empty() || _committed)
        {
            return E_UNEXPECTED;
        }

        if (expectedSize.has_value())
        {
            LARGE_INTEGER actualSize{};
            if (GetFileSizeEx(_file.get(), &actualSize) == FALSE)
            {
                const DWORD error = GetLastError();
                return HRESULT_FROM_WIN32(error == ERROR_SUCCESS ? ERROR_READ_FAULT : error);
            }
            if (actualSize.QuadPart < 0 || static_cast<uint64_t>(actualSize.QuadPart) != expectedSize.value())
            {
                return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            }
        }

#ifdef ENABLE_TESTS
        const HRESULT injectedFlushFailure = Testing::TakeNextLocalFileTransactionFlushFailure();
        if (FAILED(injectedFlushFailure))
        {
            return injectedFlushFailure;
        }
#endif
        if (FlushFileBuffers(_file.get()) == FALSE)
        {
            const DWORD error = GetLastError();
            return HRESULT_FROM_WIN32(error == ERROR_SUCCESS ? ERROR_WRITE_FAULT : error);
        }
        BY_HANDLE_FILE_INFORMATION finalizedInformation{};
        if (committedFileInformation)
        {
            if (GetFileInformationByHandle(_file.get(), &finalizedInformation) == FALSE)
            {
                const DWORD error = GetLastError();
                return HRESULT_FROM_WIN32(error == ERROR_SUCCESS ? ERROR_READ_FAULT : error);
            }
        }

#ifdef ENABLE_TESTS
        const HRESULT injectedPublishFailure = Testing::TakeNextLocalFileTransactionPublishFailure();
        if (FAILED(injectedPublishFailure))
        {
            return injectedPublishFailure;
        }
#endif
        std::vector<std::byte> renameStorage;
        const HRESULT prepareHr = PrepareRenameInfo(_targetPath.native(), renameStorage);
        if (FAILED(prepareHr))
        {
            return prepareHr;
        }
        auto* const rename = reinterpret_cast<FILE_RENAME_INFO*>(renameStorage.data());
        const bool replace = _policy == ExistingTargetPolicy::Replace;
        rename->Flags      = FILE_RENAME_FLAG_POSIX_SEMANTICS | (replace ? static_cast<DWORD>(FILE_RENAME_FLAG_REPLACE_IF_EXISTS) : 0u);

        const auto verifyTarget = [&](FILE_INFO_BY_HANDLE_CLASS informationClass) noexcept
        { return targetCheck.verify ? targetCheck.verify(targetCheck.context, informationClass) : S_OK; };
        if (const HRESULT targetHr = verifyTarget(FileRenameInfoEx); FAILED(targetHr))
            return targetHr;

        HRESULT renameHr = RenameOwnedFile(FileRenameInfoEx, rename, static_cast<DWORD>(renameStorage.size()));
        if (CanRetryLegacyRename(renameHr))
        {
            // Legacy handle rename preserves exact source authority on filesystems without Ex.
            // Closing and using a pathname here could publish a replacement of our private stage.
            rename->Flags           = 0u;
            rename->ReplaceIfExists = static_cast<BOOLEAN>(replace);
            if (const HRESULT targetHr = verifyTarget(FileRenameInfo); FAILED(targetHr))
                return targetHr;
            renameHr = RenameOwnedFile(FileRenameInfo, rename, static_cast<DWORD>(renameStorage.size()));
        }
        if (FAILED(renameHr))
            return renameHr;

        if (retainedPublishedFile)
            *retainedPublishedFile = std::move(_file);
        else
            _file.reset();
        _temporaryPath.clear();
        _committed = true;
        if (committedFileInformation)
        {
            *committedFileInformation = finalizedInformation;
        }
        return S_OK;
    }

    // Publication always reserves the destination name without replacement. An accepted
    // version is first moved through its own handle and restored only into a vacant name.
    [[nodiscard]] HRESULT CommitGuarded(const std::optional<LocalFileVersion>& acceptedVersion,
                                        std::filesystem::path& retainedBackup,
                                        std::optional<uint64_t> expectedSize     = std::nullopt,
                                        wil::unique_hfile* retainedPublishedFile = nullptr,
                                        GuardedReplacementOptions options        = {}) noexcept
    {
        retainedBackup.clear();
        _policy = ExistingTargetPolicy::FailIfExists;
        if (! acceptedVersion.has_value())
        {
            return DestinationConflict(Commit(expectedSize, options.committedInformation, retainedPublishedFile));
        }

        struct Displacement
        {
            Displacement(const std::filesystem::path& targetPath, const LocalFileVersion& version, GuardedReplacementOptions publicationOptions) noexcept
                : target(targetPath),
                  accepted(version),
                  options(publicationOptions)
            {
            }
            Displacement(const Displacement&)            = delete;
            Displacement& operator=(const Displacement&) = delete;
            Displacement(Displacement&&)                 = delete;
            Displacement& operator=(Displacement&&)      = delete;

            const std::filesystem::path& target;
            const LocalFileVersion& accepted;
            GuardedReplacementOptions options;
            wil::unique_hfile original;
            std::filesystem::path backup;
            std::vector<std::byte> restoreRename;
        };
        Displacement displacement{_targetPath, acceptedVersion.value(), options};
        const TargetRenameCheck check{.context = &displacement,
                                      .verify  = [](void* context, FILE_INFO_BY_HANDLE_CLASS) noexcept -> HRESULT
        {
            auto& state = *static_cast<Displacement*>(context);
            if (state.original)
            {
                return S_OK;
            }

            wil::unique_hfile original(CreateFileW(state.target.c_str(),
                                                   FILE_READ_ATTRIBUTES | DELETE,
                                                   FILE_SHARE_READ,
                                                   nullptr,
                                                   OPEN_EXISTING,
                                                   FILE_FLAG_OPEN_REPARSE_POINT | FILE_FLAG_BACKUP_SEMANTICS,
                                                   nullptr));
            if (! original)
            {
                const DWORD error = GetLastError();
                if (error == ERROR_FILE_NOT_FOUND || error == ERROR_PATH_NOT_FOUND || error == ERROR_SHARING_VIOLATION)
                {
                    return HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
                }
                return HRESULT_FROM_WIN32(error == ERROR_SUCCESS ? ERROR_OPEN_FAILED : error);
            }

            LocalFileVersion observed{};
            const HRESULT versionHr = QueryLocalFileVersion(original.get(), observed);
            if (FAILED(versionHr))
            {
                return versionHr;
            }
            if (observed != state.accepted)
            {
                return HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
            }
            if ((observed.attributes & FILE_ATTRIBUTE_READONLY) != 0u)
            {
                return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
            }

            const HRESULT restoreHr = PrepareRenameInfo(state.target.native(), state.restoreRename);
            if (FAILED(restoreHr))
            {
                return restoreHr;
            }
            if (state.options.backupPath)
            {
                state.backup = *state.options.backupPath;
                if (state.backup.parent_path() != state.target.parent_path() || state.backup == state.target)
                {
                    return E_INVALIDARG;
                }
            }
            else
            {
                GUID id{};
                const HRESULT guidHr = CoCreateGuid(&id);
                if (FAILED(guidHr))
                {
                    return guidHr;
                }
                std::array<wchar_t, 40u> guidText{};
                if (StringFromGUID2(id, guidText.data(), static_cast<int>(guidText.size())) <= 0)
                {
                    return E_FAIL;
                }
                state.backup = state.target.parent_path() / std::format(L".$rs-guard-{}.bak", guidText.data());
            }
            std::vector<std::byte> backupRename;
            const HRESULT prepareHr = PrepareRenameInfo(state.backup.native(), backupRename);
            if (FAILED(prepareHr))
            {
                return prepareHr;
            }
            if (state.options.prepareBackup)
            {
                const HRESULT journalHr = state.options.prepareBackup(state.options.context, original.get(), observed);
                if (FAILED(journalHr))
                {
                    return journalHr;
                }
            }
            const HRESULT moveHr = RenameBoundFileIfFree(original.get(), backupRename);
            if (FAILED(moveHr))
            {
                return moveHr;
            }
            state.original = std::move(original);
#ifdef ENABLE_TESTS
            if (const auto beforePublish = std::exchange(Testing::g_beforeGuardedFilePublish, nullptr))
            {
                beforePublish(state.target);
            }
#endif
            return S_OK;
        }};
        const HRESULT result = Commit(expectedSize, options.committedInformation, retainedPublishedFile, check);
        if (! displacement.original)
        {
            return DestinationConflict(result);
        }
        if (SUCCEEDED(result))
        {
            FILE_DISPOSITION_INFO disposition{.DeleteFile = TRUE};
            if (SetFileInformationByHandle(displacement.original.get(), FileDispositionInfo, &disposition, sizeof(disposition)) == FALSE)
            {
                retainedBackup = std::move(displacement.backup);
            }
            return S_OK;
        }

        const HRESULT restoreHr = RenameBoundFileIfFree(displacement.original.get(), displacement.restoreRename);
        if (FAILED(restoreHr))
        {
            retainedBackup         = std::move(displacement.backup);
            const HRESULT conflict = DestinationConflict(restoreHr);
            return conflict == HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH) ? conflict : HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
        }
        return DestinationConflict(result);
    }

    // The caller must already hold the verified, delete-capable source handle.
    [[nodiscard]] static HRESULT RestoreBoundFileIfFree(HANDLE file, const std::filesystem::path& target) noexcept
    {
        std::vector<std::byte> storage;
        const HRESULT prepareHr = PrepareRenameInfo(Common::Paths::ToExtendedWin32Path(target.native()), storage);
        return FAILED(prepareHr) ? prepareHr : DestinationConflict(RenameBoundFileIfFree(file, storage));
    }

    [[nodiscard]] HANDLE GetHandle() const noexcept
    {
        return _file.get();
    }

    [[nodiscard]] const std::filesystem::path& TargetPath() const noexcept
    {
        return _targetPath;
    }

    void Abort() noexcept
    {
#ifdef ENABLE_TESTS
        Testing::LocalFileTransactionRollbackSnapshot rollback{.ownedHandle = static_cast<bool>(_file)};
        const bool observedRollback = ! _committed && ! _temporaryPath.empty();
#endif
        bool markedForDeletion = false;
        if (_file && ! _committed && ! _temporaryPath.empty())
        {
            // Retire the private entry through its owned DELETE-capable handle.
            // Closing first opens a race with a reader that denies a path delete.
            FILE_DISPOSITION_INFO_EX disposition{.Flags = FILE_DISPOSITION_FLAG_DELETE | FILE_DISPOSITION_FLAG_POSIX_SEMANTICS};
            markedForDeletion = SetFileInformationByHandle(_file.get(), FileDispositionInfoEx, &disposition, sizeof(disposition)) != FALSE;
#ifdef ENABLE_TESTS
            rollback.posixDisposition = markedForDeletion;
            rollback.posixError       = markedForDeletion ? ERROR_SUCCESS : GetLastError();
#endif
            if (! markedForDeletion)
            {
                FILE_DISPOSITION_INFO legacyDisposition{.DeleteFile = TRUE};
                markedForDeletion = SetFileInformationByHandle(_file.get(), FileDispositionInfo, &legacyDisposition, sizeof(legacyDisposition)) != FALSE;
#ifdef ENABLE_TESTS
                rollback.legacyDisposition = markedForDeletion;
                rollback.legacyError       = markedForDeletion ? ERROR_SUCCESS : GetLastError();
#endif
            }
        }
        _file.reset();
        // No pathname fallback: after close the same spelling can belong to another writer.
        // If both disposition APIs refuse cleanup, preserve the uncertain entry for diagnostics.
#ifdef ENABLE_TESTS
        if (observedRollback)
            Testing::g_lastRollbackSnapshot = rollback;
#endif
        _temporaryPath.clear();
        _targetPath.clear();
        _committed = false;
    }

private:
    [[nodiscard]] static HRESULT DestinationConflict(HRESULT hr) noexcept
    {
        return hr == HRESULT_FROM_WIN32(ERROR_FILE_EXISTS) || hr == HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS) ? HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH) : hr;
    }

    [[nodiscard]] static HRESULT PrepareRenameInfo(std::wstring_view target, std::vector<std::byte>& storage) noexcept
    {
        if (target.size() > (std::numeric_limits<size_t>::max)() / sizeof(wchar_t))
        {
            return HRESULT_FROM_WIN32(ERROR_FILENAME_EXCED_RANGE);
        }
        const size_t fileNameBytes    = target.size() * sizeof(wchar_t);
        constexpr size_t kHeaderBytes = offsetof(FILE_RENAME_INFO, FileName);
        if (fileNameBytes > (std::numeric_limits<DWORD>::max)() - kHeaderBytes - sizeof(wchar_t) ||
            fileNameBytes > (std::numeric_limits<size_t>::max)() - kHeaderBytes - sizeof(wchar_t))
        {
            return HRESULT_FROM_WIN32(ERROR_FILENAME_EXCED_RANGE);
        }
        storage.assign(kHeaderBytes + fileNameBytes + sizeof(wchar_t), std::byte{});
        auto* const rename     = reinterpret_cast<FILE_RENAME_INFO*>(storage.data());
        rename->FileNameLength = static_cast<DWORD>(fileNameBytes);
        std::memcpy(rename->FileName, target.data(), fileNameBytes);
        return S_OK;
    }

    [[nodiscard]] static bool CanRetryLegacyRename(HRESULT hr) noexcept
    {
        return hr == HRESULT_FROM_WIN32(ERROR_INVALID_PARAMETER) || hr == HRESULT_FROM_WIN32(ERROR_INVALID_FUNCTION) ||
               hr == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED) || hr == HRESULT_FROM_WIN32(ERROR_CALL_NOT_IMPLEMENTED) || hr == E_ACCESSDENIED;
    }

    [[nodiscard]] static HRESULT RenameBoundFileIfFree(HANDLE file, std::vector<std::byte>& storage) noexcept
    {
        auto* const rename = reinterpret_cast<FILE_RENAME_INFO*>(storage.data());
        rename->Flags      = 0u;
        const DWORD bytes  = static_cast<DWORD>(storage.size());
        if (SetFileInformationByHandle(file, FileRenameInfoEx, rename, bytes) != FALSE)
        {
            return S_OK;
        }
        const DWORD error     = GetLastError();
        const HRESULT primary = HRESULT_FROM_WIN32(error == ERROR_SUCCESS ? ERROR_WRITE_FAULT : error);
        if (! CanRetryLegacyRename(primary))
        {
            return primary;
        }
        rename->ReplaceIfExists = FALSE;
        if (SetFileInformationByHandle(file, FileRenameInfo, rename, bytes) == FALSE)
        {
            const DWORD legacyError = GetLastError();
            return HRESULT_FROM_WIN32(legacyError == ERROR_SUCCESS ? ERROR_WRITE_FAULT : legacyError);
        }
        return S_OK;
    }

    [[nodiscard]] HRESULT RenameOwnedFile(FILE_INFO_BY_HANDLE_CLASS informationClass, FILE_RENAME_INFO* rename, DWORD bytes) noexcept
    {
#ifdef ENABLE_TESTS
        auto& nextFailure             = informationClass == FileRenameInfoEx ? Testing::g_nextRenameExFailure : Testing::g_nextLegacyRenameFailure;
        const HRESULT injectedFailure = nextFailure.exchange(S_OK, std::memory_order_relaxed);
        if (FAILED(injectedFailure))
            return injectedFailure;
#endif
        if (SetFileInformationByHandle(_file.get(), informationClass, rename, bytes) == FALSE)
        {
            const DWORD error = GetLastError();
            return HRESULT_FROM_WIN32(error == ERROR_SUCCESS ? ERROR_WRITE_FAULT : error);
        }
        return S_OK;
    }

    wil::unique_hfile _file;
    std::filesystem::path _targetPath;
    std::filesystem::path _temporaryPath;
    ExistingTargetPolicy _policy = ExistingTargetPolicy::FailIfExists;
    bool _committed              = false;
};
#pragma warning(pop)
} // namespace Common::Files
