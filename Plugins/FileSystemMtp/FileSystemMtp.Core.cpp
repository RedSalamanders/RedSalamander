#include "FileSystemMtp.h"

#include <algorithm>
#include <chrono>
#include <cstring>
#include <deque>
#include <format>
#include <limits>
#include <mutex>
#include <new>
#include <system_error>
#include <unordered_map>
#include <unordered_set>
#include <utility>

#pragma warning(push)
// (C6297) Arithmetic overflow. Results might not be an expected value.
// (C28182) Dereferencing NULL pointer.
#pragma warning(disable : 6297 28182)
#include <yyjson.h>
#pragma warning(pop)

#include "FileSystemMtpResources.h"
#include "HandleIo.h"
#include "Helpers.h"
#include "YyjsonHelpers.h"

extern HINSTANCE g_hInstance;

using namespace FileSystemMtpInternal;

namespace
{
static const int kFileSystemMtpModuleAnchor = 0;

[[nodiscard]] const wchar_t* LocalizedPluginName() noexcept
{
    static const std::wstring name = LoadEmbeddedStringResource(g_hInstance, IDS_FILESYSTEMMTP_NAME);
    return name.c_str();
}

[[nodiscard]] const wchar_t* LocalizedPluginDescription() noexcept
{
    static const std::wstring description = LoadStringResource(g_hInstance, IDS_FILESYSTEMMTP_DESCRIPTION);
    return description.c_str();
}

[[nodiscard]] const wchar_t* LocalizedFileSystemName() noexcept
{
    static const std::wstring name = LoadEmbeddedStringResource(g_hInstance, IDS_FILESYSTEMMTP_FSNAME);
    return name.c_str();
}

[[nodiscard]] const char* BoolText(bool value) noexcept
{
    return value ? "true" : "false";
}

[[nodiscard]] uint64_t HashBytes(std::span<const std::byte> bytes) noexcept;

struct OverwriteJournalContext
{
    bool failWrites      = false;
    bool mutatingCommand = false;
    std::wstring deviceIdentity;
};

struct OverwriteJournalToken
{
    std::wstring path;
    std::wstring deviceIdentity;
    uint64_t cacheGeneration = 0u;
    bool recorded            = false;
};

struct OverwriteJournalEntry
{
    std::string phase;
    std::wstring destinationPath;
    std::wstring tempPath;
    bool hasDeclaredSizeBytes           = false;
    uint64_t declaredSizeBytes          = 0;
    bool hasJournalTimestampFileTimeUtc = false;
    __int64 journalTimestampFileTimeUtc = 0;
    std::string sourceTransmitHashHex;
    std::wstring tempPuid;
    uint32_t replayAttemptCount = 0;
};

constexpr uint32_t kOverwriteJournalReplayRenameRetryLimit = 3u;
constexpr uint32_t kOverwriteJournalRetainedReplayLimit    = 3u;

struct OverwriteJournalCacheState
{
    uint64_t generation = 0u;
    bool absent         = false;
};

std::mutex g_overwriteJournalAbsentMutex;
std::unordered_map<std::wstring, OverwriteJournalCacheState> g_overwriteJournalCacheStates;
#ifdef _DEBUG
std::atomic_uint64_t g_overwriteJournalFilesystemProbes{0};
#endif

[[nodiscard]] std::string HashHex(uint64_t value);
[[nodiscard]] std::wstring OverwriteJournalDeviceIdentity(std::wstring_view connectionDevicePuid,
                                                          std::wstring_view connectionHost,
                                                          std::wstring_view normalizedPath);
[[nodiscard]] HRESULT GetOverwriteJournalPath(const OverwriteJournalContext& context, std::wstring& path) noexcept;
[[nodiscard]] HRESULT RecordOverwriteJournalIntent(const OverwriteJournalContext& context,
                                                   std::wstring_view sourcePath,
                                                   std::wstring_view destinationPath,
                                                   std::wstring_view tempPath,
                                                   uint64_t declaredSizeBytes,
                                                   std::string_view sourceTransmitHashHex,
                                                   OverwriteJournalToken& token) noexcept;
void ClearOverwriteJournalIntent(OverwriteJournalToken& token) noexcept;
[[nodiscard]] HRESULT ReplayOverwriteJournal(IMtpBackend& backend, const OverwriteJournalContext& context) noexcept;

HRESULT PrepareDeviceSourceOverwriteVerify(IMtpBackend& backend,
                                           std::wstring_view sourcePath,
                                           std::string_view verifyLevel,
                                           std::optional<std::vector<std::byte>>& sourceBytes) noexcept
{
    sourceBytes.reset();
    if (verifyLevel != "deviceReread")
    {
        return S_OK;
    }

    unsigned long sourceAttributes = 0;
    HRESULT hr                     = backend.GetAttributes(sourcePath, sourceAttributes);
    if (FAILED(hr))
    {
        return hr;
    }

    if ((sourceAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
    {
        return S_OK;
    }

    std::vector<std::byte> bytes;
    hr = backend.ReadFile(sourcePath, bytes);
    if (FAILED(hr))
    {
        return hr;
    }

    Debug::Perf::EmitValue(L"mtp.verify.device_source_bytes", static_cast<uint64_t>(bytes.size()), S_OK);
    sourceBytes = std::move(bytes);
    return S_OK;
}

HRESULT VerifyDeviceSourceOverwrite(IMtpBackend& backend,
                                    std::wstring_view destinationPath,
                                    std::string_view verifyLevel,
                                    const std::optional<std::vector<std::byte>>& sourceBytes) noexcept
{
    if (verifyLevel == "deviceReread" && sourceBytes.has_value())
    {
        std::vector<std::byte> reread;
        HRESULT hr = backend.ReadFile(destinationPath, reread);
        if (FAILED(hr))
        {
            return hr;
        }

        Debug::Perf::EmitValue(L"mtp.verify.device_source_reread_bytes", static_cast<uint64_t>(reread.size()), S_OK);
        const std::vector<std::byte>& expected = sourceBytes.value();
        if (reread.size() != expected.size() || ! std::equal(reread.begin(), reread.end(), expected.begin(), expected.end()))
        {
            Debug::Perf::EmitValue(L"mtp.verify.mismatch", 1u, HRESULT_FROM_WIN32(ERROR_CRC));
            return HRESULT_FROM_WIN32(ERROR_CRC);
        }

        return S_OK;
    }

    Debug::Perf::EmitValue(L"mtp.verify.device_source_trust_commit", 1u, S_OK);
    return S_OK;
}

HRESULT VerifyCommittedOverwrite(IMtpBackend& backend,
                                 std::wstring_view normalizedPath,
                                 std::string_view verifyLevel,
                                 std::span<const std::byte> bytes) noexcept
{
    uint64_t committedSize = 0;
    HRESULT hr             = backend.GetFileSize(normalizedPath, committedSize);
    if (FAILED(hr))
    {
        return hr;
    }

    Debug::Perf::EmitValue(L"mtp.verify.size_bytes", committedSize, S_OK);
    if (committedSize != static_cast<uint64_t>(bytes.size()))
    {
        Debug::Perf::EmitValue(L"mtp.verify.mismatch", 1u, HRESULT_FROM_WIN32(ERROR_CRC));
        return HRESULT_FROM_WIN32(ERROR_CRC);
    }

    if (verifyLevel == "deviceReread")
    {
        std::vector<std::byte> reread;
        hr = backend.ReadFile(normalizedPath, reread);
        if (FAILED(hr))
        {
            return hr;
        }

        Debug::Perf::EmitValue(L"mtp.verify.device_reread_bytes", static_cast<uint64_t>(reread.size()), S_OK);
        if (reread.size() != bytes.size() || ! std::equal(reread.begin(), reread.end(), bytes.begin(), bytes.end()))
        {
            Debug::Perf::EmitValue(L"mtp.verify.mismatch", 1u, HRESULT_FROM_WIN32(ERROR_CRC));
            return HRESULT_FROM_WIN32(ERROR_CRC);
        }

        return S_OK;
    }

    if (verifyLevel == "sizeOnly")
    {
        Debug::Perf::EmitValue(L"mtp.verify.size_only", committedSize, S_OK);
        return S_OK;
    }

    Debug::Perf::EmitValue(L"mtp.verify.transmit_hash", HashBytes(bytes), S_OK);
    return S_OK;
}

[[nodiscard]] HRESULT MakeOverwriteTempPath(std::wstring_view normalizedPath, std::wstring& tempPath) noexcept
{
    const std::wstring normalized = NormalizeMtpPath(normalizedPath);
    if (normalized == L"/")
    {
        return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
    }

    const std::wstring parent         = ParentPath(normalized);
    const std::wstring leaf           = LeafName(normalized);
    const std::wstring readablePrefix = std::wstring(L".") + leaf;
    std::wstring tempLeaf;
    const HRESULT hr = Common::Paths::BuildUniqueSiblingName(
        std::wstring_view(readablePrefix), std::wstring_view(L".rs-mtp-overwrite-"), std::wstring_view(L".tmp"), 255u, tempLeaf);
    if (FAILED(hr))
    {
        return hr;
    }

    tempPath = JoinPath(parent, tempLeaf);
    return S_OK;
}

[[nodiscard]] HRESULT ReadPersistentIdFromItemProperties(IMtpBackend& backend, std::wstring_view path, std::wstring& persistentId) noexcept
{
    persistentId.clear();

    std::string jsonUtf8;
    HRESULT hr = backend.GetItemProperties(path, jsonUtf8);
    if (FAILED(hr))
    {
        return hr;
    }

    std::string jsonCopy(jsonUtf8);
    yyjson_doc* doc = yyjson_read_opts(jsonCopy.data(), jsonCopy.size(), YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM, nullptr, nullptr);
    if (! doc)
    {
        return E_FAIL;
    }
    auto freeDoc = wil::scope_exit([&]() noexcept { yyjson_doc_free(doc); });

    yyjson_val* root = yyjson_doc_get_root(doc);
    if (! root || ! yyjson_is_obj(root))
    {
        return E_FAIL;
    }

    yyjson_val* persistentIdValue = yyjson_obj_get(root, "persistentId");
    if (! persistentIdValue || ! yyjson_is_str(persistentIdValue))
    {
        return E_FAIL;
    }

    const char* persistentIdUtf8 = yyjson_get_str(persistentIdValue);
    if (! persistentIdUtf8 || persistentIdUtf8[0] == '\0')
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    persistentId = Utf16FromUtf8(persistentIdUtf8);
    if (persistentId.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    return S_OK;
}

HRESULT CommitWriterOverwriteWithTempSwap(IMtpBackend& backend,
                                          std::wstring_view normalizedPath,
                                          std::span<const std::byte> bytes,
                                          std::string_view verifyLevel,
                                          const OverwriteJournalContext& journalContext,
                                          const std::shared_ptr<std::atomic_bool>& tempPuidMissing,
                                          const std::shared_ptr<std::atomic_bool>& tempPuidPresent) noexcept
{
    std::wstring tempPath;
    HRESULT hr = MakeOverwriteTempPath(normalizedPath, tempPath);
    if (FAILED(hr))
    {
        return hr;
    }

    OverwriteJournalToken journalToken;
    hr = RecordOverwriteJournalIntent(
        journalContext, L"", normalizedPath, tempPath, static_cast<uint64_t>(bytes.size()), HashHex(HashBytes(bytes)), journalToken);
    if (FAILED(hr))
    {
        return hr;
    }

    bool clearJournal       = true;
    auto clearJournalOnExit = wil::scope_exit([&]() noexcept
    {
        if (clearJournal)
        {
            ClearOverwriteJournalIntent(journalToken);
        }
    });

    hr = backend.WriteFile(tempPath, bytes, false);
    if (FAILED(hr))
    {
        Debug::Perf::EmitValue(L"mtp.overwrite.temp_upload_failed", 1u, hr);
        return hr;
    }

    bool cleanupTemp = true;
    auto cleanup     = wil::scope_exit([&]() noexcept
    {
        if (cleanupTemp)
        {
            static_cast<void>(backend.DeleteItem(tempPath, false));
        }
    });

    hr = VerifyCommittedOverwrite(backend, tempPath, verifyLevel, bytes);
    if (FAILED(hr))
    {
        return hr;
    }

    std::wstring tempPersistentId;
    hr = ReadPersistentIdFromItemProperties(backend, tempPath, tempPersistentId);
    if (FAILED(hr))
    {
        if (tempPuidMissing)
        {
            tempPuidMissing->store(true, std::memory_order_release);
        }
        Debug::Perf::EmitValue(L"mtp.overwrite.temp_puid_missing", 1u, hr);
        return hr;
    }
    if (tempPuidPresent)
    {
        tempPuidPresent->store(true, std::memory_order_release);
    }
    Debug::Perf::EmitValue(L"mtp.overwrite.temp_puid_present", 1u, S_OK);

    cleanupTemp = false;
    hr          = backend.DeleteItem(normalizedPath, false);
    if (FAILED(hr))
    {
        cleanupTemp = true;
        Debug::Perf::EmitValue(L"mtp.overwrite.delete_original_failed", 1u, hr);
        return hr;
    }

    hr = backend.RenameItem(tempPath, normalizedPath, false);
    if (FAILED(hr))
    {
        clearJournal = false;
        Debug::Perf::EmitValue(L"mtp.overwrite.rename_temp_failed", 1u, hr);
        return hr;
    }

    Debug::Perf::EmitValue(L"mtp.overwrite.temp_swap_committed", 1u, S_OK);
    return S_OK;
}

HRESULT CommitDeviceSourceOverwriteWithTempSwap(IMtpBackend& backend,
                                                std::wstring_view sourcePath,
                                                std::wstring_view destinationPath,
                                                bool moveSource,
                                                std::string_view verifyLevel,
                                                const OverwriteJournalContext& journalContext,
                                                const std::shared_ptr<std::atomic_bool>& tempPuidMissing,
                                                const std::shared_ptr<std::atomic_bool>& tempPuidPresent) noexcept
{
    std::wstring tempPath;
    HRESULT hr = MakeOverwriteTempPath(destinationPath, tempPath);
    if (FAILED(hr))
    {
        return hr;
    }

    OverwriteJournalToken journalToken;
    hr = RecordOverwriteJournalIntent(journalContext, sourcePath, destinationPath, tempPath, 0u, "device-source", journalToken);
    if (FAILED(hr))
    {
        return hr;
    }

    bool clearJournal       = true;
    auto clearJournalOnExit = wil::scope_exit([&]() noexcept
    {
        if (clearJournal)
        {
            ClearOverwriteJournalIntent(journalToken);
        }
    });

    std::optional<std::vector<std::byte>> sourceVerifyBytes;
    hr = PrepareDeviceSourceOverwriteVerify(backend, sourcePath, verifyLevel, sourceVerifyBytes);
    if (FAILED(hr))
    {
        return hr;
    }

    hr = backend.CopyItem(sourcePath, tempPath, false);
    if (FAILED(hr))
    {
        Debug::Perf::EmitValue(L"mtp.overwrite.temp_copy_failed", 1u, hr);
        return hr;
    }

    bool cleanupTemp = true;
    auto cleanup     = wil::scope_exit([&]() noexcept
    {
        if (cleanupTemp)
        {
            static_cast<void>(backend.DeleteItem(tempPath, true));
        }
    });

    hr = VerifyDeviceSourceOverwrite(backend, tempPath, verifyLevel, sourceVerifyBytes);
    if (FAILED(hr))
    {
        return hr;
    }

    std::wstring tempPersistentId;
    hr = ReadPersistentIdFromItemProperties(backend, tempPath, tempPersistentId);
    if (FAILED(hr))
    {
        if (tempPuidMissing)
        {
            tempPuidMissing->store(true, std::memory_order_release);
        }
        Debug::Perf::EmitValue(L"mtp.overwrite.temp_puid_missing", 1u, hr);
        return hr;
    }
    if (tempPuidPresent)
    {
        tempPuidPresent->store(true, std::memory_order_release);
    }
    Debug::Perf::EmitValue(L"mtp.overwrite.temp_puid_present", 1u, S_OK);

    cleanupTemp = false;
    hr          = backend.DeleteItem(destinationPath, false);
    if (FAILED(hr))
    {
        cleanupTemp = true;
        Debug::Perf::EmitValue(L"mtp.overwrite.delete_original_failed", 1u, hr);
        return hr;
    }

    hr = backend.RenameItem(tempPath, destinationPath, false);
    if (FAILED(hr))
    {
        clearJournal = false;
        Debug::Perf::EmitValue(L"mtp.overwrite.rename_temp_failed", 1u, hr);
        return hr;
    }

    Debug::Perf::EmitValue(L"mtp.overwrite.temp_swap_committed", 1u, S_OK);
    Debug::Perf::EmitValue(L"mtp.overwrite.device_source_temp_swap_committed", 1u, S_OK);

    if (moveSource)
    {
        hr = backend.DeleteItem(sourcePath, true);
        if (FAILED(hr))
        {
            Debug::Perf::EmitValue(L"mtp.transfer.move_fallback_delete_source_failed", 1u, hr);
            return hr;
        }
    }

    return S_OK;
}

[[nodiscard]] bool TryGetJsonBool(yyjson_val* root, const char* key, bool& out) noexcept
{
    const Common::Json::MemberResult<bool> value = Common::Json::GetBoolMember(root, key, Common::Json::MemberRequirement::Required);
    if (! value.HasValue())
    {
        return false;
    }
    out = value.value;
    return true;
}

[[nodiscard]] bool TryGetJsonUInt32(yyjson_val* root, const char* key, uint32_t minValue, uint32_t maxValue, uint32_t& out) noexcept
{
    const Common::Json::MemberResult<uint64_t> value = Common::Json::GetUInt64Member(root,
                                                                                     key,
                                                                                     Common::Json::MemberRequirement::Required,
                                                                                     Common::Json::NumericStringPolicy::Reject,
                                                                                     Common::Json::UnsignedIntegerPolicy::RequireUnsignedStorage);
    if (! value.HasValue())
    {
        return false;
    }
    const uint64_t raw = value.value;
    if (raw < minValue || raw > maxValue)
    {
        return false;
    }

    out = static_cast<uint32_t>(raw);
    return true;
}

[[nodiscard]] bool TryGetJsonString(yyjson_val* root, const char* key, std::string& out) noexcept
{
    const Common::Json::MemberResult<std::string_view> value = Common::Json::GetStringMember(root, key, Common::Json::MemberRequirement::Required);
    if (! value.HasValue())
    {
        return false;
    }
    out.assign(value.value);
    return true;
}

[[nodiscard]] bool TryGetJsonStringWide(yyjson_val* root, const char* key, std::wstring& out) noexcept
{
    std::string text;
    if (! TryGetJsonString(root, key, text))
    {
        return false;
    }

    std::wstring wide = Utf16FromUtf8(text);
    if (wide.empty() && ! text.empty())
    {
        return false;
    }

    out = std::move(wide);
    return true;
}

[[nodiscard]] std::wstring DeviceRootFromConnectionSettings(std::wstring_view connectionHost, std::wstring_view connectionFriendlyName)
{
    if (connectionHost.empty())
    {
        return L"/";
    }

    const std::wstring suffix   = MtpDeviceIdentitySuffix(connectionHost);
    const std::wstring friendly = SanitizeMtpPathComponent(std::wstring(connectionFriendlyName));
    if (friendly.empty())
    {
        return std::format(L"/{}", suffix);
    }

    return std::format(L"/{} {}", friendly, suffix);
}

[[nodiscard]] bool IsConnectionRootPath(std::wstring_view normalizedPath) noexcept
{
    constexpr std::wstring_view kConnPrefix = L"/@conn:";
    return normalizedPath.size() >= kConnPrefix.size() && OrdinalString::EqualsNoCase(normalizedPath.substr(0, kConnPrefix.size()), kConnPrefix);
}

[[nodiscard]] std::wstring ConnectionPathSuffix(std::wstring_view normalizedPath)
{
    if (! IsConnectionRootPath(normalizedPath))
    {
        return NormalizeMtpPath(normalizedPath);
    }

    constexpr std::wstring_view kConnPrefix = L"/@conn:";
    const std::wstring_view rest            = normalizedPath.substr(kConnPrefix.size());
    const size_t slash                      = rest.find(L'/');
    if (slash == std::wstring_view::npos)
    {
        return L"/";
    }

    return NormalizeMtpPath(rest.substr(slash));
}

[[nodiscard]] std::wstring JoinRootAndChildPath(std::wstring_view rootPath, std::wstring_view childPath)
{
    const std::wstring root  = NormalizeMtpPath(rootPath);
    const std::wstring child = NormalizeMtpPath(childPath);
    if (root == L"/" || child == L"/")
    {
        return root == L"/" ? child : root;
    }

    return JoinPath(root, std::wstring_view(child).substr(1u));
}

[[nodiscard]] std::string NormalizeVerifyLevel(std::string value)
{
    if (value == "deviceReread" || value == "sizeOnly" || value == "transmitHash")
    {
        return value;
    }

    return "transmitHash";
}

[[nodiscard]] uint64_t HashBytes(std::span<const std::byte> bytes) noexcept
{
    uint64_t hash = 1469598103934665603ull;
    for (const std::byte value : bytes)
    {
        hash ^= static_cast<uint64_t>(std::to_integer<unsigned char>(value));
        hash *= 1099511628211ull;
    }

    return hash;
}

[[nodiscard]] std::string HashHex(uint64_t value)
{
    return std::format("{:016X}", value);
}

[[nodiscard]] std::wstring FirstPathSegment(std::wstring_view normalizedPath)
{
    const std::wstring normalized = NormalizeMtpPath(normalizedPath);
    if (normalized.size() <= 1u)
    {
        return L"root";
    }

    const std::wstring_view rest(normalized.data() + 1u, normalized.size() - 1u);
    const size_t slash = rest.find(L'/');
    return std::wstring(slash == std::wstring_view::npos ? rest : rest.substr(0u, slash));
}

[[nodiscard]] std::wstring OverwriteJournalDeviceIdentity(std::wstring_view connectionDevicePuid,
                                                          std::wstring_view connectionHost,
                                                          std::wstring_view normalizedPath)
{
    if (! connectionDevicePuid.empty())
    {
        return std::wstring(connectionDevicePuid);
    }
    if (! connectionHost.empty())
    {
        return std::wstring(connectionHost);
    }

    std::wstring segment = FirstPathSegment(normalizedPath);
    if (segment.empty())
    {
        segment = L"root";
    }
    return segment;
}

[[nodiscard]] HRESULT GetLocalAppDataPath(std::wstring& localAppData) noexcept
{
    localAppData.clear();
    const DWORD required = GetEnvironmentVariableW(L"LOCALAPPDATA", nullptr, 0);
    if (required == 0)
    {
        const DWORD error = GetLastError();
        return HRESULT_FROM_WIN32(error == ERROR_SUCCESS ? ERROR_PATH_NOT_FOUND : error);
    }

    std::wstring buffer(required, L'\0');
    const DWORD written = GetEnvironmentVariableW(L"LOCALAPPDATA", buffer.data(), required);
    if (written == 0 || written >= required)
    {
        const DWORD error = GetLastError();
        return HRESULT_FROM_WIN32(error == ERROR_SUCCESS ? ERROR_INSUFFICIENT_BUFFER : error);
    }

    buffer.resize(written);
    localAppData = std::move(buffer);
    return S_OK;
}

[[nodiscard]] HRESULT EnsureDirectoryExists(std::wstring_view directoryPath) noexcept
{
    if (directoryPath.empty())
    {
        return E_INVALIDARG;
    }

    std::wstring path(directoryPath);
    std::replace(path.begin(), path.end(), L'/', L'\\');

    size_t pos = 0;
    if (path.size() >= 3u && path[1] == L':' && (path[2] == L'\\' || path[2] == L'/'))
    {
        pos = 3u;
    }

    while (pos <= path.size())
    {
        const size_t next          = path.find(L'\\', pos);
        const std::wstring current = next == std::wstring::npos ? path : path.substr(0u, next);
        if (! current.empty() && ! (current.size() == 2u && current[1] == L':'))
        {
            if (CreateDirectoryW(current.c_str(), nullptr) == 0)
            {
                const DWORD error = GetLastError();
                if (error != ERROR_ALREADY_EXISTS)
                {
                    return HRESULT_FROM_WIN32(error);
                }

                const DWORD attributes = GetFileAttributesW(current.c_str());
                if (attributes == INVALID_FILE_ATTRIBUTES || (attributes & FILE_ATTRIBUTE_DIRECTORY) == 0)
                {
                    return HRESULT_FROM_WIN32(ERROR_DIRECTORY);
                }
            }
        }

        if (next == std::wstring::npos)
        {
            break;
        }
        pos = next + 1u;
    }

    return S_OK;
}

[[nodiscard]] HRESULT WriteUtf8FileAtomic(std::wstring_view path, std::string_view bytes) noexcept
{
    std::wstring target(path);
    if (target.empty())
    {
        return E_INVALIDARG;
    }

    const std::wstring temp = target + L".tmp";
    wil::unique_hfile file(CreateFileW(temp.c_str(), GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
    if (! file)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    const HRESULT writeHr = Common::HandleIo::WriteAll(file.get(), bytes.data(), bytes.size());
    if (FAILED(writeHr))
    {
        file.reset();
        DeleteFileW(temp.c_str());
        return writeHr;
    }

    if (FlushFileBuffers(file.get()) == 0)
    {
        const DWORD error = GetLastError();
        file.reset();
        DeleteFileW(temp.c_str());
        return HRESULT_FROM_WIN32(error);
    }

    file.reset();
    if (MoveFileExW(temp.c_str(), target.c_str(), MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH) == 0)
    {
        const DWORD error = GetLastError();
        DeleteFileW(temp.c_str());
        return HRESULT_FROM_WIN32(error);
    }

    return S_OK;
}

[[nodiscard]] bool IsMissingItemHr(HRESULT hr) noexcept
{
    return hr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) || hr == HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND);
}

[[nodiscard]] HRESULT ReadUtf8File(std::wstring_view path, std::string& bytes) noexcept
{
    bytes.clear();
    std::wstring target(path);
    if (target.empty())
    {
        return E_INVALIDARG;
    }

    wil::unique_hfile file(CreateFileW(
        target.c_str(), GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr));
    if (! file)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    uint64_t fileSize    = 0u;
    const HRESULT sizeHr = Common::HandleIo::GetFileSizeBounded(file.get(), 256u * 1024u, fileSize);
    if (FAILED(sizeHr))
    {
        return sizeHr;
    }

    bytes.resize(static_cast<size_t>(fileSize));
    const HRESULT readHr = Common::HandleIo::ReadExact(file.get(), bytes.data(), bytes.size());
    if (readHr == HRESULT_FROM_WIN32(ERROR_HANDLE_EOF))
    {
        return HRESULT_FROM_WIN32(ERROR_READ_FAULT);
    }
    return readHr;
}

[[nodiscard]] HRESULT ParseOverwriteJournalEntry(std::string_view jsonUtf8, OverwriteJournalEntry& entry) noexcept
{
    entry = {};
    if (jsonUtf8.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    std::string jsonCopy(jsonUtf8);
    yyjson_doc* doc = yyjson_read_opts(jsonCopy.data(), jsonCopy.size(), YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM, nullptr, nullptr);
    if (! doc)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }
    auto freeDoc = wil::scope_exit([&]() noexcept { yyjson_doc_free(doc); });

    yyjson_val* root = yyjson_doc_get_root(doc);
    if (! root || ! yyjson_is_obj(root))
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    yyjson_val* schemaVersion = yyjson_obj_get(root, "schemaVersion");
    if (! schemaVersion || ! yyjson_is_uint(schemaVersion) || yyjson_get_uint(schemaVersion) != 1u)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    yyjson_val* entries = yyjson_obj_get(root, "entries");
    if (! entries || ! yyjson_is_arr(entries) || yyjson_arr_size(entries) == 0u)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    yyjson_val* firstEntry = yyjson_arr_get(entries, 0);
    if (! firstEntry || ! yyjson_is_obj(firstEntry))
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    if (! TryGetJsonString(firstEntry, "phase", entry.phase) || ! TryGetJsonStringWide(firstEntry, "destinationPath", entry.destinationPath) ||
        ! TryGetJsonStringWide(firstEntry, "tempPath", entry.tempPath))
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    yyjson_val* declaredSizeBytes = yyjson_obj_get(firstEntry, "declaredSizeBytes");
    if (declaredSizeBytes)
    {
        if (! yyjson_is_uint(declaredSizeBytes))
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }
        entry.hasDeclaredSizeBytes = true;
        entry.declaredSizeBytes    = yyjson_get_uint(declaredSizeBytes);
    }

    yyjson_val* journalTimestamp = yyjson_obj_get(firstEntry, "journalTimestampFileTimeUtc");
    if (journalTimestamp)
    {
        if (! yyjson_is_uint(journalTimestamp) || yyjson_get_uint(journalTimestamp) > static_cast<uint64_t>((std::numeric_limits<__int64>::max)()))
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }
        entry.hasJournalTimestampFileTimeUtc = true;
        entry.journalTimestampFileTimeUtc    = static_cast<__int64>(yyjson_get_uint(journalTimestamp));
    }

    static_cast<void>(TryGetJsonString(firstEntry, "sourceTransmitHashHex", entry.sourceTransmitHashHex));
    static_cast<void>(TryGetJsonStringWide(firstEntry, "tempPuid", entry.tempPuid));

    yyjson_val* replayAttemptCount = yyjson_obj_get(firstEntry, "replayAttemptCount");
    if (replayAttemptCount)
    {
        if (! yyjson_is_uint(replayAttemptCount) || yyjson_get_uint(replayAttemptCount) > std::numeric_limits<uint32_t>::max())
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }
        entry.replayAttemptCount = static_cast<uint32_t>(yyjson_get_uint(replayAttemptCount));
    }

    entry.destinationPath = NormalizeMtpPath(entry.destinationPath);
    entry.tempPath        = NormalizeMtpPath(entry.tempPath);
    if (entry.destinationPath == L"/" || entry.tempPath == L"/" || entry.destinationPath.empty() || entry.tempPath.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    return S_OK;
}

[[nodiscard]] std::wstring OverwriteJournalCacheIdentity(std::wstring_view deviceIdentity)
{
    // The journal root is part of the cache identity. Production normally has one stable
    // LOCALAPPDATA root, while selftests and portable hosts can redirect it between sessions;
    // an "absent" observation from one root must not suppress replay in another root.
    std::wstring identity;
    std::wstring localAppData;
    if (SUCCEEDED(GetLocalAppDataPath(localAppData)) && ! localAppData.empty())
    {
        identity.reserve(localAppData.size() + deviceIdentity.size() + 1u);
        identity.append(localAppData);
        identity.push_back(L'|');
    }
    identity.append(deviceIdentity.empty() ? std::wstring_view(L"unknown") : deviceIdentity);
    std::replace(identity.begin(), identity.end(), L'/', L'\\');
    std::transform(identity.begin(), identity.end(), identity.begin(), [](wchar_t ch) noexcept { return static_cast<wchar_t>(std::towlower(ch)); });
    return identity;
}

[[nodiscard]] OverwriteJournalCacheState CaptureOverwriteJournalCacheState(std::wstring_view deviceIdentity)
{
    const std::wstring identity = OverwriteJournalCacheIdentity(deviceIdentity);
    std::lock_guard lock(g_overwriteJournalAbsentMutex);
    return g_overwriteJournalCacheStates[identity];
}

[[nodiscard]] uint64_t BeginOverwriteJournalWrite(std::wstring_view deviceIdentity)
{
    const std::wstring identity = OverwriteJournalCacheIdentity(deviceIdentity);
    std::lock_guard lock(g_overwriteJournalAbsentMutex);
    OverwriteJournalCacheState& state = g_overwriteJournalCacheStates[identity];
    ++state.generation;
    state.absent = false;
    return state.generation;
}

[[nodiscard]] bool MarkOverwriteJournalAbsent(std::wstring_view deviceIdentity, uint64_t observedGeneration)
{
    const std::wstring identity = OverwriteJournalCacheIdentity(deviceIdentity);
    std::lock_guard lock(g_overwriteJournalAbsentMutex);
    OverwriteJournalCacheState& state = g_overwriteJournalCacheStates[identity];
    if (state.generation != observedGeneration)
    {
        return false;
    }
    state.absent = true;
    return true;
}

[[nodiscard]] HRESULT WriteOverwriteJournalEntry(const OverwriteJournalContext& context,
                                                 const OverwriteJournalEntry& entry,
                                                 uint32_t replayAttemptCount) noexcept
{
    static_cast<void>(BeginOverwriteJournalWrite(context.deviceIdentity));
    std::wstring journalPath;
    HRESULT hr = GetOverwriteJournalPath(context, journalPath);
    if (FAILED(hr))
    {
        return hr;
    }

    const std::string phase                 = entry.phase.empty() ? "planned" : entry.phase;
    const uint64_t declaredSizeBytes        = entry.hasDeclaredSizeBytes ? entry.declaredSizeBytes : 0u;
    const __int64 journalTimestamp          = entry.hasJournalTimestampFileTimeUtc ? entry.journalTimestampFileTimeUtc : NowFileTime64();
    const std::string sourceTransmitHashHex = entry.sourceTransmitHashHex.empty() ? "replay-retry" : entry.sourceTransmitHashHex;
    std::string tempPuidJson;
    if (! entry.tempPuid.empty())
    {
        tempPuidJson = std::format(R"json(,"tempPuid":"{}")json", JsonEscapeUtf8(Utf8FromUtf16(entry.tempPuid)));
    }

    const std::string json = std::format(
        R"json({{"schemaVersion":1,"entries":[{{"phase":"{}","devicePuid":"{}","sourcePath":"","destinationPath":"{}","tempPath":"{}","declaredSizeBytes":{},"sourceTransmitHashHex":"{}","journalTimestampFileTimeUtc":{},"replayAttemptCount":{}{} }}]}})json",
        JsonEscapeUtf8(phase),
        JsonEscapeUtf8(Utf8FromUtf16(context.deviceIdentity)),
        JsonEscapeUtf8(Utf8FromUtf16(entry.destinationPath)),
        JsonEscapeUtf8(Utf8FromUtf16(entry.tempPath)),
        declaredSizeBytes,
        JsonEscapeUtf8(sourceTransmitHashHex),
        static_cast<long long>(journalTimestamp),
        replayAttemptCount,
        tempPuidJson);

    return WriteUtf8FileAtomic(journalPath, json);
}

[[nodiscard]] HRESULT GetOverwriteJournalPath(const OverwriteJournalContext& context, std::wstring& path) noexcept
{
    path.clear();

    std::wstring localAppData;
    HRESULT hr = GetLocalAppDataPath(localAppData);
    if (FAILED(hr))
    {
        return hr;
    }

    const std::wstring identity   = context.deviceIdentity.empty() ? std::wstring(L"unknown") : context.deviceIdentity;
    const std::wstring deviceHash = FormatMtpIdentityHash(identity);
    const std::wstring directory  = localAppData + L"\\RedSalamander\\PluginState\\FileSystemMtp\\" + deviceHash;
    hr                            = EnsureDirectoryExists(directory);
    if (FAILED(hr))
    {
        return hr;
    }

    path = directory + L"\\overwrite-journal.json";
    return S_OK;
}

[[nodiscard]] HRESULT RecordOverwriteJournalIntent(const OverwriteJournalContext& context,
                                                   std::wstring_view sourcePath,
                                                   std::wstring_view destinationPath,
                                                   std::wstring_view tempPath,
                                                   uint64_t declaredSizeBytes,
                                                   std::string_view sourceTransmitHashHex,
                                                   OverwriteJournalToken& token) noexcept
{
    token                 = {};
    token.deviceIdentity  = context.deviceIdentity;
    token.cacheGeneration = BeginOverwriteJournalWrite(context.deviceIdentity);
    if (context.failWrites)
    {
        constexpr HRESULT hr = HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        Debug::Perf::EmitValue(L"mtp.overwrite.journal_write_failed", 1u, hr);
        return hr;
    }

    std::wstring journalPath;
    HRESULT hr = GetOverwriteJournalPath(context, journalPath);
    if (FAILED(hr))
    {
        Debug::Perf::EmitValue(L"mtp.overwrite.journal_write_failed", 1u, hr);
        return hr;
    }

    const std::string json = std::format(
        R"json({{"schemaVersion":1,"entries":[{{"phase":"planned","devicePuid":"{}","sourcePath":"{}","destinationPath":"{}","tempPath":"{}","declaredSizeBytes":{},"sourceTransmitHashHex":"{}","journalTimestampFileTimeUtc":{}}}]}})json",
        JsonEscapeUtf8(Utf8FromUtf16(context.deviceIdentity)),
        JsonEscapeUtf8(Utf8FromUtf16(sourcePath)),
        JsonEscapeUtf8(Utf8FromUtf16(destinationPath)),
        JsonEscapeUtf8(Utf8FromUtf16(tempPath)),
        declaredSizeBytes,
        JsonEscapeUtf8(sourceTransmitHashHex),
        static_cast<long long>(NowFileTime64()));

    hr = WriteUtf8FileAtomic(journalPath, json);
    if (FAILED(hr))
    {
        Debug::Perf::EmitValue(L"mtp.overwrite.journal_write_failed", 1u, hr);
        return hr;
    }

    token.path     = std::move(journalPath);
    token.recorded = true;
    Debug::Perf::EmitValue(L"mtp.overwrite.journal_intent_recorded", 1u, S_OK);
    return S_OK;
}

enum class DeviceSourceOperation : uint8_t
{
    Copy,
    Move,
    Rename,
};

HRESULT ExecuteDeviceSourceOperation(IMtpBackend& backend,
                                     DeviceSourceOperation operation,
                                     std::wstring_view sourcePath,
                                     std::wstring_view destinationPath,
                                     bool allowOverwrite,
                                     std::string_view verifyLevel,
                                     bool createdObjectPuidUnsupported,
                                     const OverwriteJournalContext& journalContext,
                                     const std::shared_ptr<std::atomic_bool>& tempPuidMissing,
                                     const std::shared_ptr<std::atomic_bool>& tempPuidPresent) noexcept
{
    unsigned long destinationAttributes = 0;
    bool destinationExisted             = false;
    if (allowOverwrite)
    {
        const HRESULT destinationHr = backend.GetAttributes(destinationPath, destinationAttributes);
        if (SUCCEEDED(destinationHr))
        {
            destinationExisted = true;
        }
        else if (destinationHr != HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) && destinationHr != HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND))
        {
            return destinationHr;
        }
    }

    if (destinationExisted)
    {
        if (operation == DeviceSourceOperation::Rename && sourcePath == destinationPath)
        {
            return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
        }

        unsigned long sourceAttributes = 0;
        RETURN_IF_FAILED(backend.GetAttributes(sourcePath, sourceAttributes));
        const bool sourceDirectory      = (sourceAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
        const bool destinationDirectory = (destinationAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
        if (destinationDirectory || sourceDirectory)
        {
            // The temp-swap overwrite protocol is file-only; reject type-changing directory overwrites before any mutation.
            return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        }
        if (createdObjectPuidUnsupported)
        {
            const HRESULT blockedHr = HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
            Debug::Perf::EmitValue(L"mtp.overwrite.temp_puid_policy_blocked", 1u, blockedHr);
            return blockedHr;
        }

        return CommitDeviceSourceOverwriteWithTempSwap(
            backend, sourcePath, destinationPath, operation != DeviceSourceOperation::Copy, verifyLevel, journalContext, tempPuidMissing, tempPuidPresent);
    }

    switch (operation)
    {
        case DeviceSourceOperation::Copy: return backend.CopyItem(sourcePath, destinationPath, false);
        case DeviceSourceOperation::Move: return backend.MoveItem(sourcePath, destinationPath, false);
        case DeviceSourceOperation::Rename: return backend.RenameItem(sourcePath, destinationPath, false);
    }
    return E_UNEXPECTED;
}

void ClearOverwriteJournalIntent(OverwriteJournalToken& token) noexcept
{
    if (! token.recorded || token.path.empty())
    {
        return;
    }

    if (DeleteFileW(token.path.c_str()) == 0)
    {
        const DWORD error = GetLastError();
        if (error != ERROR_FILE_NOT_FOUND && error != ERROR_PATH_NOT_FOUND)
        {
            Debug::Perf::EmitValue(L"mtp.overwrite.journal_clear_failed", 1u, HRESULT_FROM_WIN32(error));
            return;
        }
    }

    token.recorded = false;
    static_cast<void>(MarkOverwriteJournalAbsent(token.deviceIdentity, token.cacheGeneration));
    Debug::Perf::EmitValue(L"mtp.overwrite.journal_cleared", 1u, S_OK);
}

[[nodiscard]] HRESULT QuarantineOverwriteJournalIntent(OverwriteJournalToken& token) noexcept
{
    if (! token.recorded || token.path.empty())
    {
        return S_OK;
    }

    const std::wstring stalePath = token.path + L".stale";
    if (MoveFileExW(token.path.c_str(), stalePath.c_str(), MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH) == 0)
    {
        const DWORD error = GetLastError();
        if (error != ERROR_FILE_NOT_FOUND && error != ERROR_PATH_NOT_FOUND)
        {
            const HRESULT hr = HRESULT_FROM_WIN32(error);
            Debug::Perf::EmitValue(L"mtp.overwrite.journal_replay_stale_quarantine_failed", 1u, hr);
            return hr;
        }
    }

    token.recorded = false;
    static_cast<void>(MarkOverwriteJournalAbsent(token.deviceIdentity, token.cacheGeneration));
    Debug::Perf::EmitValue(L"mtp.overwrite.journal_replay_stale_quarantined", 1u, S_OK);
    return S_OK;
}

[[nodiscard]] HRESULT BackendItemExists(IMtpBackend& backend, std::wstring_view path, bool& exists) noexcept
{
    exists = false;

    unsigned long attributes = 0;
    const HRESULT hr         = backend.GetAttributes(path, attributes);
    if (SUCCEEDED(hr))
    {
        exists = true;
        return S_OK;
    }
    if (IsMissingItemHr(hr))
    {
        return S_OK;
    }

    return hr;
}

[[nodiscard]] HRESULT DestinationMatchesDeclaredSize(IMtpBackend& backend, const OverwriteJournalEntry& entry, bool& matches) noexcept
{
    matches = false;
    if (! entry.hasDeclaredSizeBytes)
    {
        return S_OK;
    }

    uint64_t destinationSize = 0;
    const HRESULT hr         = backend.GetFileSize(entry.destinationPath, destinationSize);
    if (FAILED(hr))
    {
        Debug::Perf::EmitValue(L"mtp.overwrite.journal_replay_completed_swap_size_probe_failed", 1u, hr);
        return S_OK;
    }

    matches = destinationSize == entry.declaredSizeBytes;
    return S_OK;
}

[[nodiscard]] bool IsOverwriteTempLeaf(std::wstring_view leaf) noexcept
{
    return leaf.find(L".rs-mtp-overwrite-") != std::wstring_view::npos;
}

[[nodiscard]] bool NameMatchesTempLeafToken(std::wstring_view itemName, std::wstring_view tempLeaf) noexcept
{
    if (itemName == tempLeaf)
    {
        return true;
    }

    return itemName.size() > tempLeaf.size() + 2u && itemName.rfind(tempLeaf, 0) == 0 && itemName[tempLeaf.size()] == L' ' &&
           itemName[tempLeaf.size() + 1u] == L'[';
}

[[nodiscard]] bool ShouldRunNoTempPuidOrphanSweep(const OverwriteJournalEntry& entry) noexcept
{
    if (! entry.tempPuid.empty() || ! entry.hasDeclaredSizeBytes || ! entry.hasJournalTimestampFileTimeUtc)
    {
        return false;
    }
    if (entry.sourceTransmitHashHex == "device-source")
    {
        return false;
    }

    const std::wstring tempLeaf = LeafName(entry.tempPath);
    if (! IsOverwriteTempLeaf(tempLeaf))
    {
        return false;
    }

    return entry.phase == "planned" || entry.phase == "uploaded" || entry.phase == "committed";
}

[[nodiscard]] HRESULT SweepCommittedTempWithoutTempPuid(IMtpBackend& backend,
                                                        const OverwriteJournalEntry& entry,
                                                        bool& journalHandled,
                                                        uint32_t& retainedCandidateCount) noexcept
{
    journalHandled         = false;
    retainedCandidateCount = 0;
    if (! ShouldRunNoTempPuidOrphanSweep(entry))
    {
        return S_OK;
    }

    const std::wstring parentPath = ParentPath(entry.tempPath);
    const std::wstring tempLeaf   = LeafName(entry.tempPath);
    if (parentPath.empty() || parentPath == entry.tempPath || tempLeaf.empty())
    {
        return S_OK;
    }

    std::vector<MtpItem> items;
    HRESULT hr = backend.EnumerateDirectory(parentPath, items);
    if (FAILED(hr))
    {
        Debug::Perf::EmitValue(L"mtp.overwrite.journal_orphan_sweep_failed", 1u, hr);
        return hr;
    }

    std::vector<std::wstring> candidatePaths;
    candidatePaths.reserve(items.size());
    for (const MtpItem& item : items)
    {
        if ((item.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0 || ! item.streamable)
        {
            continue;
        }
        if (item.sizeBytes != entry.declaredSizeBytes)
        {
            continue;
        }
        if (item.lastWriteTime != 0 && item.lastWriteTime < entry.journalTimestampFileTimeUtc)
        {
            continue;
        }
        if (! NameMatchesTempLeafToken(item.name, tempLeaf))
        {
            continue;
        }

        candidatePaths.push_back(JoinPath(parentPath, item.name));
    }

    if (candidatePaths.size() != 1u)
    {
        retainedCandidateCount = static_cast<uint32_t>((std::min)(candidatePaths.size(), static_cast<size_t>((std::numeric_limits<uint32_t>::max)())));
        Debug::Perf::EmitValue(L"mtp.overwrite.journal_orphan_sweep_retained", static_cast<uint64_t>(candidatePaths.size()), S_OK);
        return S_OK;
    }

    hr = backend.DeleteItem(candidatePaths.front(), false);
    if (FAILED(hr))
    {
        Debug::Perf::EmitValue(L"mtp.overwrite.journal_orphan_sweep_failed", 1u, hr);
        return hr;
    }

    Debug::Perf::EmitValue(L"mtp.overwrite.journal_orphan_sweep_temp_removed", 1u, S_OK);
    journalHandled = true;
    return S_OK;
}

[[nodiscard]] HRESULT RetainOrQuarantineOverwriteJournal(const OverwriteJournalContext& context,
                                                         OverwriteJournalToken& token,
                                                         const OverwriteJournalEntry& entry,
                                                         uint32_t retainedCandidateCount) noexcept
{
    const uint32_t nextAttempt =
        entry.replayAttemptCount >= kOverwriteJournalRetainedReplayLimit ? kOverwriteJournalRetainedReplayLimit : entry.replayAttemptCount + 1u;
    if (nextAttempt >= kOverwriteJournalRetainedReplayLimit)
    {
        Debug::Perf::EmitValue(L"mtp.overwrite.journal_replay_stale_retained_terminal", retainedCandidateCount, S_OK);
        return QuarantineOverwriteJournalIntent(token);
    }

    const HRESULT hr = WriteOverwriteJournalEntry(context, entry, nextAttempt);
    if (FAILED(hr))
    {
        Debug::Perf::EmitValue(L"mtp.overwrite.journal_replay_failed", 1u, hr);
        return hr;
    }

    Debug::Perf::EmitValue(L"mtp.overwrite.journal_replay_stale_retained", retainedCandidateCount, S_OK);
    return S_OK;
}

[[nodiscard]] HRESULT ReplayOverwriteJournal(IMtpBackend& backend, const OverwriteJournalContext& context) noexcept
{
    const OverwriteJournalCacheState cacheState = CaptureOverwriteJournalCacheState(context.deviceIdentity);
    if (cacheState.absent)
    {
        return S_OK;
    }
    if (context.mutatingCommand)
    {
        Debug::Perf::EmitValue(L"mtp.overwrite.journal_mutating_replay_probe", 1u, S_OK);
    }

    std::wstring journalPath;
    HRESULT hr = GetOverwriteJournalPath(context, journalPath);
    if (FAILED(hr))
    {
        Debug::Perf::EmitValue(L"mtp.overwrite.journal_replay_unavailable", 1u, hr);
        return S_OK;
    }

    std::string jsonUtf8;
#ifdef _DEBUG
    static_cast<void>(g_overwriteJournalFilesystemProbes.fetch_add(1u, std::memory_order_acq_rel));
#endif
    hr = ReadUtf8File(journalPath, jsonUtf8);
    if (IsMissingItemHr(hr))
    {
        static_cast<void>(MarkOverwriteJournalAbsent(context.deviceIdentity, cacheState.generation));
        return S_OK;
    }
    if (FAILED(hr))
    {
        Debug::Perf::EmitValue(L"mtp.overwrite.journal_replay_failed", 1u, hr);
        return hr;
    }

    OverwriteJournalToken token{.path = journalPath, .deviceIdentity = context.deviceIdentity, .recorded = true};
    OverwriteJournalEntry entry;
    hr = ParseOverwriteJournalEntry(jsonUtf8, entry);
    if (FAILED(hr))
    {
        Debug::Perf::EmitValue(L"mtp.overwrite.journal_replay_failed", 1u, hr);
        ClearOverwriteJournalIntent(token);
        return S_OK;
    }

    bool destinationExists = false;
    hr                     = BackendItemExists(backend, entry.destinationPath, destinationExists);
    if (FAILED(hr))
    {
        Debug::Perf::EmitValue(L"mtp.overwrite.journal_replay_failed", 1u, hr);
        return hr;
    }

    bool tempExists = false;
    hr              = BackendItemExists(backend, entry.tempPath, tempExists);
    if (FAILED(hr))
    {
        Debug::Perf::EmitValue(L"mtp.overwrite.journal_replay_failed", 1u, hr);
        return hr;
    }

    if (destinationExists && tempExists)
    {
        hr = backend.DeleteItem(entry.tempPath, true);
        if (FAILED(hr))
        {
            Debug::Perf::EmitValue(L"mtp.overwrite.journal_replay_failed", 1u, hr);
            return hr;
        }

        Debug::Perf::EmitValue(L"mtp.overwrite.journal_replay_temp_removed", 1u, S_OK);
        ClearOverwriteJournalIntent(token);
        return S_OK;
    }

    if (destinationExists)
    {
        bool journalHandled             = false;
        uint32_t retainedCandidateCount = 0;
        hr                              = SweepCommittedTempWithoutTempPuid(backend, entry, journalHandled, retainedCandidateCount);
        if (FAILED(hr))
        {
            Debug::Perf::EmitValue(L"mtp.overwrite.journal_replay_failed", 1u, hr);
            return hr;
        }
        if (journalHandled)
        {
            ClearOverwriteJournalIntent(token);
            return S_OK;
        }
        if (ShouldRunNoTempPuidOrphanSweep(entry))
        {
            bool completedSwapInferred = false;
            hr                         = DestinationMatchesDeclaredSize(backend, entry, completedSwapInferred);
            if (FAILED(hr))
            {
                Debug::Perf::EmitValue(L"mtp.overwrite.journal_replay_failed", 1u, hr);
                return hr;
            }
            if (retainedCandidateCount == 0u && completedSwapInferred)
            {
                Debug::Perf::EmitValue(L"mtp.overwrite.journal_replay_completed_swap_inferred", 1u, S_OK);
                ClearOverwriteJournalIntent(token);
                return S_OK;
            }

            return RetainOrQuarantineOverwriteJournal(context, token, entry, retainedCandidateCount);
        }

        Debug::Perf::EmitValue(L"mtp.overwrite.journal_replay_noop", 1u, S_OK);
        ClearOverwriteJournalIntent(token);
        return S_OK;
    }

    if (tempExists)
    {
        hr = backend.RenameItem(entry.tempPath, entry.destinationPath, false);
        if (FAILED(hr))
        {
            Debug::Perf::EmitValue(L"mtp.overwrite.journal_replay_failed", 1u, hr);
            const uint32_t nextAttempt =
                entry.replayAttemptCount >= kOverwriteJournalReplayRenameRetryLimit ? kOverwriteJournalReplayRenameRetryLimit : entry.replayAttemptCount + 1u;
            if (nextAttempt >= kOverwriteJournalReplayRenameRetryLimit)
            {
                Debug::Perf::EmitValue(L"mtp.overwrite.journal_replay_rename_terminal", 1u, hr);
                ClearOverwriteJournalIntent(token);
                return S_OK;
            }

            const HRESULT rewriteHr = WriteOverwriteJournalEntry(context, entry, nextAttempt);
            if (FAILED(rewriteHr))
            {
                Debug::Perf::EmitValue(L"mtp.overwrite.journal_replay_failed", 1u, rewriteHr);
                return rewriteHr;
            }

            Debug::Perf::EmitValue(L"mtp.overwrite.journal_replay_retry_scheduled", nextAttempt, hr);
            return hr;
        }

        Debug::Perf::EmitValue(L"mtp.overwrite.journal_replay_temp_committed", 1u, S_OK);
        ClearOverwriteJournalIntent(token);
        return S_OK;
    }

    Debug::Perf::EmitValue(L"mtp.overwrite.journal_replay_missing_objects", 1u, HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND));
    ClearOverwriteJournalIntent(token);
    return S_OK;
}

class MtpBufferedWriter final : public IFileWriter
{
public:
    MtpBufferedWriter(FileSystemMtp* owner, std::wstring path, FileSystemFlags flags) noexcept : _owner(owner), _path(std::move(path)), _flags(flags)
    {
    }

    MtpBufferedWriter(const MtpBufferedWriter&)            = delete;
    MtpBufferedWriter(MtpBufferedWriter&&)                 = delete;
    MtpBufferedWriter& operator=(const MtpBufferedWriter&) = delete;
    MtpBufferedWriter& operator=(MtpBufferedWriter&&)      = delete;

    ~MtpBufferedWriter() = default;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (! ppvObject)
        {
            return E_POINTER;
        }

        if (riid == __uuidof(IUnknown) || riid == __uuidof(IFileWriter))
        {
            *ppvObject = static_cast<IFileWriter*>(this);
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
        const ULONG result = _refCount.fetch_sub(1, std::memory_order_acq_rel) - 1;
        if (result == 0)
        {
            delete this;
        }
        return result;
    }

    HRESULT STDMETHODCALLTYPE GetPosition(uint64_t* positionBytes) noexcept override
    {
        if (! positionBytes)
        {
            return E_POINTER;
        }

        *positionBytes = static_cast<uint64_t>(_bytes.size());
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Write(const void* buffer, unsigned long bytesToWrite, unsigned long* bytesWritten) noexcept override
    {
        if (! bytesWritten)
        {
            return E_POINTER;
        }
        *bytesWritten = 0;
        if (! buffer && bytesToWrite > 0)
        {
            return E_POINTER;
        }
        if (_committed)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_STATE);
        }
        if (FAILED(_failedHr))
        {
            return _failedHr;
        }
        if (bytesToWrite > (std::numeric_limits<size_t>::max)() - _bytes.size())
        {
            _failedHr = E_OUTOFMEMORY;
            return _failedHr;
        }

        const auto* first = static_cast<const std::byte*>(buffer);
        _bytes.insert(_bytes.end(), first, first + bytesToWrite);
        *bytesWritten = bytesToWrite;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Commit() noexcept override
    {
        if (_committed)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_STATE);
        }
        if (! _owner)
        {
            return E_FAIL;
        }

        _committed = true;
        return _owner->CommitFileWriter(_path, _flags, _bytes);
    }

private:
    std::atomic_ulong _refCount{1};
    wil::com_ptr<FileSystemMtp> _owner;
    std::wstring _path;
    FileSystemFlags _flags = FILESYSTEM_FLAG_NONE;
    std::vector<std::byte> _bytes;
    HRESULT _failedHr = S_OK;
    bool _committed   = false;
};

class SingleItemIndexCallback final : public IFileSystemCallback
{
public:
    SingleItemIndexCallback(IFileSystemCallback* inner, unsigned long itemIndex) noexcept : _inner(inner), _itemIndex(itemIndex)
    {
    }

    HRESULT STDMETHODCALLTYPE FileSystemProgress(FileSystemOperation operationType,
                                                 unsigned long totalItems,
                                                 unsigned long completedItems,
                                                 uint64_t totalBytes,
                                                 uint64_t completedBytes,
                                                 const wchar_t* currentSourcePath,
                                                 const wchar_t* currentDestinationPath,
                                                 uint64_t currentItemTotalBytes,
                                                 uint64_t currentItemCompletedBytes,
                                                 FileSystemOptions* options,
                                                 uint64_t progressStreamId,
                                                 void* cookie) noexcept override
    {
        return _inner ? _inner->FileSystemProgress(operationType,
                                                   totalItems,
                                                   completedItems,
                                                   totalBytes,
                                                   completedBytes,
                                                   currentSourcePath,
                                                   currentDestinationPath,
                                                   currentItemTotalBytes,
                                                   currentItemCompletedBytes,
                                                   options,
                                                   progressStreamId,
                                                   cookie)
                      : S_OK;
    }

    HRESULT STDMETHODCALLTYPE FileSystemItemCompleted(FileSystemOperation operationType,
                                                      unsigned long itemIndex,
                                                      const wchar_t* sourcePath,
                                                      const wchar_t* destinationPath,
                                                      HRESULT status,
                                                      FileSystemOptions* options,
                                                      void* cookie) noexcept override
    {
        static_cast<void>(itemIndex);
        return _inner ? _inner->FileSystemItemCompleted(operationType, _itemIndex, sourcePath, destinationPath, status, options, cookie) : S_OK;
    }

    HRESULT STDMETHODCALLTYPE FileSystemShouldCancel(BOOL* pCancel, void* cookie) noexcept override
    {
        return _inner ? _inner->FileSystemShouldCancel(pCancel, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE FileSystemIssue(FileSystemOperation operationType,
                                              const wchar_t* sourcePath,
                                              const wchar_t* destinationPath,
                                              HRESULT status,
                                              FileSystemIssueAction* action,
                                              FileSystemOptions* options,
                                              void* cookie) noexcept override
    {
        return _inner ? _inner->FileSystemIssue(operationType, sourcePath, destinationPath, status, action, options, cookie) : E_POINTER;
    }

private:
    IFileSystemCallback* _inner = nullptr;
    unsigned long _itemIndex    = 0;
};

struct MtpBackendCommandStatus
{
    MtpBackendCommandStatus() = default;

    MtpBackendCommandStatus(const MtpBackendCommandStatus&)            = delete;
    MtpBackendCommandStatus(MtpBackendCommandStatus&&)                 = delete;
    MtpBackendCommandStatus& operator=(const MtpBackendCommandStatus&) = delete;
    MtpBackendCommandStatus& operator=(MtpBackendCommandStatus&&)      = delete;

    std::mutex mutex;
    std::condition_variable cv;
    bool completed = false;
    HRESULT hr     = E_PENDING;
};

void CompleteMtpBackendCommandStatus(const std::shared_ptr<MtpBackendCommandStatus>& status, HRESULT hr) noexcept
{
    if (! status)
    {
        return;
    }

    {
        std::lock_guard statusLock(status->mutex);
        status->hr        = hr;
        status->completed = true;
    }
    status->cv.notify_all();
}

struct MtpQuarantinedCommand
{
    MtpQuarantinedCommand(std::jthread commandWorker, std::shared_ptr<MtpBackendCommandStatus> commandStatus) noexcept
        : worker(std::move(commandWorker)),
          status(std::move(commandStatus))
    {
    }

    MtpQuarantinedCommand(const MtpQuarantinedCommand&)                = delete;
    MtpQuarantinedCommand(MtpQuarantinedCommand&&) noexcept            = default;
    MtpQuarantinedCommand& operator=(const MtpQuarantinedCommand&)     = delete;
    MtpQuarantinedCommand& operator=(MtpQuarantinedCommand&&) noexcept = default;

    std::jthread worker;
    std::shared_ptr<MtpBackendCommandStatus> status;
};

[[nodiscard]] std::atomic_uint32_t& MtpPendingCancelRequests() noexcept;

struct MtpBackendCancelRequest
{
    MtpBackendCancelRequest()                                          = default;
    MtpBackendCancelRequest(const MtpBackendCancelRequest&)            = delete;
    MtpBackendCancelRequest& operator=(const MtpBackendCancelRequest&) = delete;
    MtpBackendCancelRequest(MtpBackendCancelRequest&&)                 = delete;
    MtpBackendCancelRequest& operator=(MtpBackendCancelRequest&&)      = delete;

    std::shared_ptr<IMtpBackend> backend;
    wil::unique_hmodule moduleKeepAlive;

    void Execute(PTP_CALLBACK_INSTANCE instance) noexcept
    {
        if (backend)
        {
            backend->RequestCancel();
            backend.reset();
        }

        static_cast<void>(MtpPendingCancelRequests().fetch_sub(1u, std::memory_order_acq_rel));
        if (moduleKeepAlive)
        {
            TransferModulePinToCallbackReturn(instance, moduleKeepAlive);
        }
    }
};

[[nodiscard]] std::atomic_uint32_t& MtpPendingCancelRequests() noexcept
{
    static auto* pending = new std::atomic_uint32_t{0u};
    return *pending;
}

[[nodiscard]] std::mutex& MtpQuarantineMutex() noexcept
{
    static auto* mutex = new std::mutex();
    return *mutex;
}

[[nodiscard]] std::vector<MtpQuarantinedCommand>& MtpQuarantineCommands() noexcept
{
    static auto* commands = new std::vector<MtpQuarantinedCommand>();
    return *commands;
}

void SweepMtpQuarantinedCommands() noexcept
{
    std::lock_guard lock(MtpQuarantineMutex());
    std::erase_if(MtpQuarantineCommands(),
                  [](const MtpQuarantinedCommand& command) noexcept
    {
        std::lock_guard statusLock(command.status->mutex);
        return command.status->completed;
    });
}

void QueueBackendCancel(std::shared_ptr<IMtpBackend> backend) noexcept
{
    if (! backend)
    {
        return;
    }

    auto request             = std::make_unique<MtpBackendCancelRequest>();
    request->backend         = std::move(backend);
    request->moduleKeepAlive = AcquireModuleReferenceFromAddress(&kFileSystemMtpModuleAnchor);
    if (! request->moduleKeepAlive)
    {
        request->backend->RequestCancel();
        request->backend.reset();
        Debug::Error(L"FileSystemMtp: failed to pin the module for an asynchronous backend cancel request.");
        return;
    }

    static_cast<void>(MtpPendingCancelRequests().fetch_add(1u, std::memory_order_acq_rel));
    if (! SubmitOwnedThreadpoolCallbackWithInstance(request))
    {
        const DWORD submitError = GetLastError();
        static_cast<void>(MtpPendingCancelRequests().fetch_sub(1u, std::memory_order_acq_rel));
        request->backend->RequestCancel();
        request->backend.reset();

        // A failed submission is executing inside this DLL. Keep the extra
        // reference until process exit instead of releasing the module while
        // the failing path can still return through plugin code.
        static_cast<void>(request->moduleKeepAlive.release());
        static std::atomic_bool loggedSubmitFailure{false};
        if (! loggedSubmitFailure.exchange(true, std::memory_order_acq_rel))
        {
            Debug::Error(L"FileSystemMtp: failed to queue a backend cancel request (error {}). The module pin is retained until process exit.", submitError);
        }
    }
}
} // namespace

namespace FileSystemMtpInternal
{
namespace
{
struct MtpDetachedBackendWorker
{
    MtpDetachedBackendWorker()  = default;
    ~MtpDetachedBackendWorker() = default;

    MtpDetachedBackendWorker(const MtpDetachedBackendWorker&)                = delete;
    MtpDetachedBackendWorker(MtpDetachedBackendWorker&&) noexcept            = default;
    MtpDetachedBackendWorker& operator=(const MtpDetachedBackendWorker&)     = delete;
    MtpDetachedBackendWorker& operator=(MtpDetachedBackendWorker&&) noexcept = default;

    std::jthread worker;
    std::shared_ptr<MtpBackendCommandStatus> exitStatus;
};
} // namespace

#ifdef _DEBUG
bool RunOverwriteJournalGenerationSelfTest() noexcept
{
    constexpr std::wstring_view identity = L"firebreak-journal-generation-selftest";
    {
        std::lock_guard lock(g_overwriteJournalAbsentMutex);
        g_overwriteJournalCacheStates.erase(OverwriteJournalCacheIdentity(identity));
    }

    const OverwriteJournalCacheState observedAbsent = CaptureOverwriteJournalCacheState(identity);
    const uint64_t writeGeneration                  = BeginOverwriteJournalWrite(identity);
    const bool staleMarkAccepted                    = MarkOverwriteJournalAbsent(identity, observedAbsent.generation);
    const OverwriteJournalCacheState afterWrite     = CaptureOverwriteJournalCacheState(identity);
    return ! staleMarkAccepted && writeGeneration == observedAbsent.generation + 1u && afterWrite.generation == writeGeneration && ! afterWrite.absent;
}

uint64_t ResetOverwriteJournalProbeCountForSelfTest() noexcept
{
    {
        std::lock_guard lock(g_overwriteJournalAbsentMutex);
        g_overwriteJournalCacheStates.clear();
    }
    return g_overwriteJournalFilesystemProbes.exchange(0u, std::memory_order_acq_rel);
}

uint64_t GetOverwriteJournalProbeCountForSelfTest() noexcept
{
    return g_overwriteJournalFilesystemProbes.load(std::memory_order_acquire);
}

void NotifyOverwriteJournalInjectedForSelfTest(std::wstring_view deviceIdentity) noexcept
{
    // Selftests inject retained journal files directly to reproduce crash windows. Mirror
    // the generation invalidation performed by production journal writers so the absent
    // cache cannot hide that externally staged fixture.
    static_cast<void>(BeginOverwriteJournalWrite(deviceIdentity));
}
#endif

class MtpBackendCommandQueue final
{
public:
    explicit MtpBackendCommandQueue(std::shared_ptr<IMtpBackend> backend) noexcept : _state(std::make_shared<State>(std::move(backend)))
    {
        try
        {
            _worker = std::jthread([state = _state]() noexcept { WorkerMain(std::move(state)); });
        }
        // Thread creation is an ABI-boundary failure: surface it as a command error instead of terminating this noexcept constructor.
        catch (const std::system_error& error)
        {
            _startHr          = HRESULT_FROM_WIN32(ERROR_NOT_ENOUGH_MEMORY);
            _state->accepting = false;
            _state->failureHr = _startHr;
            Debug::Error(L"FileSystemMtp: failed to start the backend command worker: {}", Utf16FromUtf8(error.what()));
        }
    }

    ~MtpBackendCommandQueue()
    {
        StopAndCancelPending(HRESULT_FROM_WIN32(ERROR_OPERATION_ABORTED));
        if (_worker.joinable())
        {
            _worker.join();
        }
    }

    MtpBackendCommandQueue(const MtpBackendCommandQueue&)            = delete;
    MtpBackendCommandQueue(MtpBackendCommandQueue&&)                 = delete;
    MtpBackendCommandQueue& operator=(const MtpBackendCommandQueue&) = delete;
    MtpBackendCommandQueue& operator=(MtpBackendCommandQueue&&)      = delete;

    [[nodiscard]] HRESULT StartResult() const noexcept
    {
        return _startHr;
    }

    [[nodiscard]] std::shared_ptr<MtpBackendCommandStatus> Submit(std::function<HRESULT(IMtpBackend&)> command,
                                                                  std::wstring recoveryDeviceIdentity,
                                                                  MtpBackendCommandKind kind) noexcept
    {
        auto status = std::make_shared<MtpBackendCommandStatus>();
        if (! command)
        {
            CompleteMtpBackendCommandStatus(status, E_INVALIDARG);
            return status;
        }

        {
            std::lock_guard lock(_state->mutex);
            if (! _state->accepting)
            {
                CompleteMtpBackendCommandStatus(status, _state->failureHr);
                return status;
            }

            _state->queue.push_back(QueuedCommand{
                .command                = std::move(command),
                .recoveryDeviceIdentity = std::move(recoveryDeviceIdentity),
                .kind                   = kind,
                .status                 = status,
            });
        }
        _state->cv.notify_one();
        return status;
    }

    void StopAndCancelPending(HRESULT pendingHr) noexcept
    {
        std::deque<QueuedCommand> pending;
        {
            std::lock_guard lock(_state->mutex);
            _state->accepting = false;
            _state->stop      = true;
            pending.swap(_state->queue);
        }

        for (const QueuedCommand& command : pending)
        {
            CompleteMtpBackendCommandStatus(command.status, pendingHr);
        }
        _state->cv.notify_all();
    }

    [[nodiscard]] MtpDetachedBackendWorker DetachForQuarantine(HRESULT pendingHr) noexcept
    {
        auto exitStatus = std::make_shared<MtpBackendCommandStatus>();
        std::deque<QueuedCommand> pending;
        {
            std::lock_guard lock(_state->mutex);
            _state->accepting  = false;
            _state->stop       = true;
            _state->exitStatus = exitStatus;
            pending.swap(_state->queue);
        }

        for (const QueuedCommand& command : pending)
        {
            CompleteMtpBackendCommandStatus(command.status, pendingHr);
        }
        _state->cv.notify_all();

        MtpDetachedBackendWorker detached;
        detached.worker     = std::move(_worker);
        detached.exitStatus = std::move(exitStatus);
        return detached;
    }

private:
    struct QueuedCommand
    {
        std::function<HRESULT(IMtpBackend&)> command;
        std::wstring recoveryDeviceIdentity;
        MtpBackendCommandKind kind = MtpBackendCommandKind::ReadOnly;
        std::shared_ptr<MtpBackendCommandStatus> status;
    };

    struct State
    {
        explicit State(std::shared_ptr<IMtpBackend> stateBackend) noexcept : backend(std::move(stateBackend))
        {
        }

        State(const State&)            = delete;
        State(State&&)                 = delete;
        State& operator=(const State&) = delete;
        State& operator=(State&&)      = delete;

        std::shared_ptr<IMtpBackend> backend;
        std::mutex mutex;
        std::condition_variable cv;
        std::deque<QueuedCommand> queue;
        std::shared_ptr<MtpBackendCommandStatus> exitStatus;
        bool accepting    = true;
        bool stop         = false;
        HRESULT failureHr = HRESULT_FROM_WIN32(ERROR_DEVICE_NOT_CONNECTED);
    };

    static void WorkerMain(std::shared_ptr<State> state) noexcept
    {
        const HRESULT coinitHr = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
        if (FAILED(coinitHr))
        {
            std::deque<QueuedCommand> pending;
            std::shared_ptr<MtpBackendCommandStatus> exitStatus;
            {
                std::lock_guard lock(state->mutex);
                state->accepting = false;
                state->stop      = true;
                state->failureHr = coinitHr;
                pending.swap(state->queue);
                exitStatus = state->exitStatus;
            }
            for (const QueuedCommand& queued : pending)
            {
                CompleteMtpBackendCommandStatus(queued.status, coinitHr);
            }
            Debug::Error(L"FileSystemMtp: backend worker MTA initialization failed (hr=0x{:08X}).", static_cast<unsigned long>(coinitHr));
            CompleteMtpBackendCommandStatus(exitStatus, coinitHr);
            return;
        }
        [[maybe_unused]] const wil::unique_couninitialize_call coUninit;

        for (;;)
        {
            QueuedCommand queued;
            {
                std::unique_lock lock(state->mutex);
                state->cv.wait(lock, [&]() noexcept { return state->stop || ! state->queue.empty(); });
                if (state->queue.empty())
                {
                    if (state->stop)
                    {
                        break;
                    }
                    continue;
                }

                queued = std::move(state->queue.front());
                state->queue.pop_front();
            }

            HRESULT hr = E_FAIL;
            if (! state->backend)
            {
                hr = E_FAIL;
            }
            else
            {
                if (! queued.recoveryDeviceIdentity.empty())
                {
                    const OverwriteJournalContext recoveryContext{
                        .failWrites      = false,
                        .mutatingCommand = queued.kind == MtpBackendCommandKind::Mutating,
                        .deviceIdentity  = queued.recoveryDeviceIdentity,
                    };
                    hr = ReplayOverwriteJournal(*state->backend, recoveryContext);
                }
                else
                {
                    hr = S_OK;
                }

                if (SUCCEEDED(hr))
                {
                    hr = queued.command(*state->backend);
                }
            }

            CompleteMtpBackendCommandStatus(queued.status, hr);
        }

        std::shared_ptr<MtpBackendCommandStatus> exitStatus;
        {
            std::lock_guard lock(state->mutex);
            exitStatus = state->exitStatus;
        }
        CompleteMtpBackendCommandStatus(exitStatus, S_OK);
    }

    std::shared_ptr<State> _state;
    std::jthread _worker;
    HRESULT _startHr = S_OK;
};
} // namespace FileSystemMtpInternal

class MtpBackendReader final : public IFileReader
{
public:
    MtpBackendReader(FileSystemMtp* owner, std::shared_ptr<IMtpBackendFileReader> backendReader, uint64_t backendGeneration) noexcept
        : _owner(owner),
          _backendReader(std::move(backendReader)),
          _backendGeneration(backendGeneration)
    {
    }

    MtpBackendReader(const MtpBackendReader&)            = delete;
    MtpBackendReader(MtpBackendReader&&)                 = delete;
    MtpBackendReader& operator=(const MtpBackendReader&) = delete;
    MtpBackendReader& operator=(MtpBackendReader&&)      = delete;

    ~MtpBackendReader() = default;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (! ppvObject)
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
        const ULONG result = _refCount.fetch_sub(1, std::memory_order_acq_rel) - 1;
        if (result == 0)
        {
            delete this;
        }
        return result;
    }

    HRESULT STDMETHODCALLTYPE GetSize(uint64_t* sizeBytes) noexcept override
    {
        if (! sizeBytes)
        {
            return E_POINTER;
        }
        *sizeBytes = 0;
        if (! _owner || ! _backendReader)
        {
            return E_FAIL;
        }

        auto result                                                = std::make_shared<uint64_t>(0);
        const std::shared_ptr<IMtpBackendFileReader> backendReader = _backendReader;
        const HRESULT hr                                           = _owner->RunBackendCommand([backendReader, result](IMtpBackend&) noexcept {
            return backendReader->GetSize(*result);
        }, {}, MtpBackendCommandKind::ReadOnly, _backendGeneration);
        if (FAILED(hr))
        {
            return hr;
        }

        *sizeBytes = *result;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Seek(__int64 offset, unsigned long origin, uint64_t* newPosition) noexcept override
    {
        if (! newPosition)
        {
            return E_POINTER;
        }
        *newPosition = 0;
        if (! _owner || ! _backendReader)
        {
            return E_FAIL;
        }

        auto result                                                = std::make_shared<uint64_t>(0);
        const std::shared_ptr<IMtpBackendFileReader> backendReader = _backendReader;
        const HRESULT hr                                           = _owner->RunBackendCommand([backendReader, offset, origin, result](IMtpBackend&) noexcept {
            return backendReader->Seek(offset, origin, *result);
        }, {}, MtpBackendCommandKind::ReadOnly, _backendGeneration);
        if (FAILED(hr))
        {
            return hr;
        }

        *newPosition = *result;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Read(void* buffer, unsigned long bytesToRead, unsigned long* bytesRead) noexcept override
    {
        if (! bytesRead)
        {
            return E_POINTER;
        }
        *bytesRead = 0;
        if (bytesToRead == 0u)
        {
            return S_OK;
        }
        if (! buffer)
        {
            return E_POINTER;
        }
        if (! _owner || ! _backendReader)
        {
            return E_FAIL;
        }

        auto scratch                                               = std::make_shared<std::vector<std::byte>>(bytesToRead);
        auto result                                                = std::make_shared<unsigned long>(0);
        const std::shared_ptr<IMtpBackendFileReader> backendReader = _backendReader;
        const HRESULT hr = _owner->RunBackendCommand([backendReader, scratch, bytesToRead, result](IMtpBackend&) noexcept {
            return backendReader->Read(std::span<std::byte>(*scratch), bytesToRead, *result);
        }, {}, MtpBackendCommandKind::ReadOnly, _backendGeneration);
        if (FAILED(hr))
        {
            return hr;
        }
        if (*result > bytesToRead)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        std::memcpy(buffer, scratch->data(), *result);
        *bytesRead = *result;
        return S_OK;
    }

private:
    std::atomic_ulong _refCount{1};
    wil::com_ptr<FileSystemMtp> _owner;
    std::shared_ptr<IMtpBackendFileReader> _backendReader;
    uint64_t _backendGeneration = 0u;
};

FileSystemMtp::FileSystemMtp(IHost* host) noexcept : FileSystemMtp(host, CreateWpdMtpBackend())
{
}

FileSystemMtp::FileSystemMtp(IHost* host, std::unique_ptr<IMtpBackend> backend) noexcept
{
    _backend = std::shared_ptr<IMtpBackend>(std::move(backend));
    if (! _backend)
    {
        _backend = CreateWpdMtpBackend();
    }
    static_cast<void>(CreateBackendWorkerLocked());

    _settings.readOnly = _backend ? _backend->GetInfo().readOnly : true;

    _metaData = {
        .id          = kPluginId,
        .shortId     = kPluginShortId,
        .name        = LocalizedPluginName(),
        .description = LocalizedPluginDescription(),
        .author      = kPluginAuthor,
        .version     = kPluginVersion,
    };

    _driveDisplayName = L"MTP";
    _driveFileSystem  = LocalizedFileSystemName();

    if (host)
    {
        static_cast<void>(host->QueryInterface(__uuidof(IHostConnections), _hostConnections.put_void()));
    }
}

void FileSystemMtp::AbandonBackendSessionLocked(const std::shared_ptr<IMtpBackend>& abandonedBackend) noexcept
{
    ++_backendGeneration;
    if (abandonedBackend && _backend == abandonedBackend && abandonedBackend->GetInfo().liveWpd)
    {
        _backend = CreateWpdMtpBackend();
    }
}

HRESULT FileSystemMtp::CreateBackendWorkerLocked() noexcept
{
    _backendWorker.reset();
    if (! _backend)
    {
        _backendWorkerCreationHr = E_FAIL;
        return _backendWorkerCreationHr;
    }

    auto worker              = std::make_shared<MtpBackendCommandQueue>(_backend);
    _backendWorkerCreationHr = worker->StartResult();
    if (FAILED(_backendWorkerCreationHr))
    {
        return _backendWorkerCreationHr;
    }
    _backendWorker = std::move(worker);
    return S_OK;
}

FileSystemMtp::~FileSystemMtp() = default;

std::wstring FileSystemMtp::OverwriteJournalIdentityForPath(std::wstring_view normalizedPath) const noexcept
{
    Settings settings;
    {
        std::lock_guard lock(_stateMutex);
        settings = _settings;
    }

    return OverwriteJournalDeviceIdentity(settings.connectionDevicePuid, settings.connectionHost, normalizedPath);
}

HRESULT FileSystemMtp::RunBackendCommand(std::function<HRESULT(IMtpBackend&)> command,
                                         std::wstring recoveryDeviceIdentity,
                                         MtpBackendCommandKind kind,
                                         uint64_t requiredBackendGeneration) noexcept
{
    if (! command)
    {
        return E_INVALIDARG;
    }

    std::shared_ptr<IMtpBackend> backend;
    std::shared_ptr<MtpBackendCommandQueue> backendWorker;
    uint32_t timeoutMs = 0;
    {
        std::lock_guard lock(_stateMutex);
        if (_disconnected)
        {
            return HRESULT_FROM_WIN32(ERROR_DEVICE_NOT_CONNECTED);
        }
        if (requiredBackendGeneration != 0u && requiredBackendGeneration != _backendGeneration)
        {
            return HRESULT_FROM_WIN32(ERROR_DEVICE_NOT_CONNECTED);
        }

        backend       = _backend;
        backendWorker = _backendWorker;
        timeoutMs     = _settings.commandTimeoutMs;
    }

    if (! backend || ! backendWorker)
    {
        return FAILED(_backendWorkerCreationHr) ? _backendWorkerCreationHr : E_FAIL;
    }

    auto status = backendWorker->Submit(std::move(command), std::move(recoveryDeviceIdentity), kind);

    {
        std::unique_lock statusLock(status->mutex);
        if (status->cv.wait_for(statusLock, std::chrono::milliseconds(timeoutMs), [&]() noexcept { return status->completed; }))
        {
            const HRESULT hr = status->hr;
            return hr;
        }
    }

    MtpDetachedBackendWorker detachedWorker;
    {
        std::lock_guard lock(_stateMutex);
        _disconnected = true;
        if (_backendWorker == backendWorker)
        {
            detachedWorker = backendWorker->DetachForQuarantine(HRESULT_FROM_WIN32(ERROR_DEVICE_NOT_CONNECTED));
            _backendWorker.reset();
        }
        AbandonBackendSessionLocked(backend);
    }

    constexpr HRESULT timedOutHr = HRESULT_FROM_WIN32(ERROR_DEVICE_NOT_CONNECTED);
    Debug::Perf::EmitValue(L"mtp.device.watchdog_trips", 1u, timedOutHr);
    QueueBackendCancel(backend);

    if (detachedWorker.worker.joinable())
    {
        std::lock_guard lock(MtpQuarantineMutex());
        MtpQuarantineCommands().emplace_back(std::move(detachedWorker.worker), std::move(detachedWorker.exitStatus));
    }

    return timedOutHr;
}

HRESULT STDMETHODCALLTYPE FileSystemMtp::QueryInterface(REFIID riid, void** ppvObject) noexcept
{
    if (! ppvObject)
    {
        return E_POINTER;
    }

    if (riid == __uuidof(IUnknown) || riid == __uuidof(IFileSystem))
    {
        *ppvObject = static_cast<IFileSystem*>(this);
    }
    else if (riid == __uuidof(IFileSystemIO))
    {
        *ppvObject = static_cast<IFileSystemIO*>(this);
    }
    else if (riid == __uuidof(IFileSystemDirectoryOperations))
    {
        *ppvObject = static_cast<IFileSystemDirectoryOperations*>(this);
    }
    else if (riid == __uuidof(IFileSystemInitialize))
    {
        *ppvObject = static_cast<IFileSystemInitialize*>(this);
    }
    else if (riid == __uuidof(IInformations))
    {
        *ppvObject = static_cast<IInformations*>(this);
    }
    else if (riid == __uuidof(INavigationMenu))
    {
        *ppvObject = static_cast<INavigationMenu*>(this);
    }
    else if (riid == __uuidof(IDriveInfo))
    {
        *ppvObject = static_cast<IDriveInfo*>(this);
    }
    else
    {
        *ppvObject = nullptr;
        return E_NOINTERFACE;
    }

    AddRef();
    return S_OK;
}

ULONG STDMETHODCALLTYPE FileSystemMtp::AddRef() noexcept
{
    return _refCount.fetch_add(1, std::memory_order_relaxed) + 1;
}

ULONG STDMETHODCALLTYPE FileSystemMtp::Release() noexcept
{
    const ULONG result = _refCount.fetch_sub(1, std::memory_order_acq_rel) - 1;
    if (result == 0)
    {
        delete this;
    }
    return result;
}

HRESULT STDMETHODCALLTYPE FileSystemMtp::GetMetaData(const PluginMetaData** metaData) noexcept
{
    if (! metaData)
    {
        return E_POINTER;
    }

    *metaData = &_metaData;
    return S_OK;
}

const char* FileSystemMtp::StaticConfigurationSchema() noexcept
{
    return R"json(
{
  "version": 1,
  "title": "MTP",
  "fields": [
    {
      "key": "readOnly",
      "type": "bool",
      "label": "Read only",
      "description": "Open the device without mutation support.",
      "default": true
    },
    {
      "key": "byteVerifyOnOverwrite",
      "type": "option",
      "label": "Overwrite verification",
      "default": "transmitHash",
      "choices": [
        { "label": "Transmit hash", "value": "transmitHash" },
        { "label": "Device re-read", "value": "deviceReread" },
        { "label": "Size only", "value": "sizeOnly" }
      ]
    },
    {
      "key": "commandTimeoutMs",
      "type": "value",
      "label": "Command timeout (ms)",
      "default": 120000,
      "min": 5000,
      "max": 600000
    },
    {
      "key": "enumerationPageSize",
      "type": "value",
      "label": "Enumeration page size",
      "default": 32,
      "min": 1,
      "max": 256
    }
  ]
}
)json";
}

HRESULT STDMETHODCALLTYPE FileSystemMtp::GetConfigurationSchema(const char** schemaJsonUtf8) noexcept
{
    if (! schemaJsonUtf8)
    {
        return E_POINTER;
    }

    *schemaJsonUtf8 = StaticConfigurationSchema();
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemMtp::SetConfiguration(const char* configurationJsonUtf8) noexcept
{
    Settings nextSettings;
    bool backendLiveWpd = true;
    {
        std::lock_guard lock(_stateMutex);
        nextSettings   = _settings;
        backendLiveWpd = ! _backend || _backend->GetInfo().liveWpd;
    }
    std::string source                 = configurationJsonUtf8 ? configurationJsonUtf8 : "{}";
    const uint32_t commandTimeoutMinMs = backendLiveWpd ? 5'000u : 25u;

    if (configurationJsonUtf8 && configurationJsonUtf8[0] != '\0')
    {
        yyjson_doc* doc = yyjson_read(configurationJsonUtf8, std::strlen(configurationJsonUtf8), YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM);
        if (! doc)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }
        auto freeDoc = wil::scope_exit([&]() noexcept { yyjson_doc_free(doc); });

        yyjson_val* root = yyjson_doc_get_root(doc);
        if (! root || ! yyjson_is_obj(root))
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        bool readOnly = nextSettings.readOnly;
        if (TryGetJsonBool(root, "readOnly", readOnly))
        {
            nextSettings.readOnly = readOnly;
        }
        if (! backendLiveWpd)
        {
            static_cast<void>(TryGetJsonBool(root, "failOverwriteJournalWrites", nextSettings.failOverwriteJournalWrites));
        }

        static_cast<void>(TryGetJsonStringWide(root, "host", nextSettings.connectionHost));
        static_cast<void>(TryGetJsonStringWide(root, "initialPath", nextSettings.connectionInitialPath));

        if (yyjson_val* extra = yyjson_obj_get(root, "extra"); extra && yyjson_is_obj(extra))
        {
            if (TryGetJsonBool(extra, "readOnly", readOnly))
            {
                nextSettings.readOnly = readOnly;
            }
            if (! backendLiveWpd)
            {
                static_cast<void>(TryGetJsonBool(extra, "failOverwriteJournalWrites", nextSettings.failOverwriteJournalWrites));
            }
            static_cast<void>(TryGetJsonStringWide(extra, "devicePuid", nextSettings.connectionDevicePuid));
            static_cast<void>(TryGetJsonStringWide(extra, "friendlyName", nextSettings.connectionFriendlyName));
        }

        std::string verifyLevel;
        if (TryGetJsonString(root, "byteVerifyOnOverwrite", verifyLevel))
        {
            nextSettings.byteVerifyOnOverwrite = NormalizeVerifyLevel(std::move(verifyLevel));
        }
        if (yyjson_val* extra = yyjson_obj_get(root, "extra"); extra && yyjson_is_obj(extra) && TryGetJsonString(extra, "byteVerifyOnOverwrite", verifyLevel))
        {
            nextSettings.byteVerifyOnOverwrite = NormalizeVerifyLevel(std::move(verifyLevel));
        }

        static_cast<void>(TryGetJsonUInt32(root, "commandTimeoutMs", commandTimeoutMinMs, 600'000u, nextSettings.commandTimeoutMs));
        static_cast<void>(TryGetJsonUInt32(root, "enumerationPageSize", 1u, 256u, nextSettings.enumerationPageSize));
        if (yyjson_val* extra = yyjson_obj_get(root, "extra"); extra && yyjson_is_obj(extra))
        {
            static_cast<void>(TryGetJsonUInt32(extra, "commandTimeoutMs", commandTimeoutMinMs, 600'000u, nextSettings.commandTimeoutMs));
            static_cast<void>(TryGetJsonUInt32(extra, "enumerationPageSize", 1u, 256u, nextSettings.enumerationPageSize));
        }
    }

    {
        std::lock_guard lock(_stateMutex);
        _settings            = std::move(nextSettings);
        _configurationSource = std::move(source);
    }
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemMtp::GetConfiguration(const char** configurationJsonUtf8) noexcept
{
    if (! configurationJsonUtf8)
    {
        return E_POINTER;
    }

    std::lock_guard lock(_jsonMutex);
    *configurationJsonUtf8 = StoreJson(_configurationJson, BuildConfigurationJson());
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemMtp::SomethingToSave(BOOL* pSomethingToSave) noexcept
{
    if (! pSomethingToSave)
    {
        return E_POINTER;
    }

    std::lock_guard lock(_stateMutex);
    *pSomethingToSave = (! _configurationSource.empty() && _configurationSource != "{}") ? TRUE : FALSE;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemMtp::Initialize(const wchar_t* rootPath, const char* optionsJsonUtf8) noexcept
{
    if (optionsJsonUtf8 && optionsJsonUtf8[0] != '\0')
    {
        const HRESULT hr = SetConfiguration(optionsJsonUtf8);
        if (FAILED(hr))
        {
            return hr;
        }
    }

    std::lock_guard lock(_stateMutex);
    const std::wstring normalizedRoot = NormalizeMtpPath(rootPath ? rootPath : L"/");
    if (IsConnectionRootPath(normalizedRoot) && ! _settings.connectionHost.empty())
    {
        const std::wstring deviceRoot = DeviceRootFromConnectionSettings(_settings.connectionHost, _settings.connectionFriendlyName);
        _rootPath                     = JoinRootAndChildPath(
            deviceRoot, _settings.connectionInitialPath.empty() ? std::wstring_view(L"/") : std::wstring_view(_settings.connectionInitialPath));
    }
    else
    {
        _rootPath = normalizedRoot;
    }
    if (! _backendWorker && _backend)
    {
        const HRESULT workerHr = CreateBackendWorkerLocked();
        if (FAILED(workerHr))
        {
            return workerHr;
        }
    }
    _initialized  = true;
    _disconnected = false;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemMtp::ReadDirectoryInfo(const wchar_t* path, IFilesInformation** ppFilesInformation) noexcept
{
    if (! ppFilesInformation)
    {
        return E_POINTER;
    }
    *ppFilesInformation = nullptr;

    std::wstring normalized;
    HRESULT hr = NormalizeInputPath(path, normalized);
    if (FAILED(hr))
    {
        return hr;
    }
    hr = CheckConnected();
    if (FAILED(hr))
    {
        return hr;
    }

    Debug::Perf::Scope perf(L"mtp.enum.directory_us");

    auto itemsResult = std::make_shared<std::vector<MtpItem>>();
    {
        const std::wstring commandPath = normalized;
        hr = RunBackendCommand([commandPath, itemsResult](IMtpBackend& backend) noexcept { return backend.EnumerateDirectory(commandPath, *itemsResult); },
                               OverwriteJournalIdentityForPath(commandPath));
    }
    if (FAILED(hr))
    {
        return hr;
    }
    std::vector<MtpItem> items = std::move(*itemsResult);

    std::vector<FilesInformationMtp::Entry> entries;
    entries.reserve(items.size());
    unsigned long index = 0;
    for (const MtpItem& item : items)
    {
        entries.push_back(FilesInformationMtp::Entry{
            .name           = item.name,
            .fileIndex      = index,
            .attributes     = item.attributes,
            .sizeBytes      = item.sizeBytes,
            .creationTime   = item.creationTime,
            .lastAccessTime = item.lastAccessTime,
            .lastWriteTime  = item.lastWriteTime,
            .changeTime     = item.changeTime,
        });
        ++index;
    }

    std::unique_ptr<FilesInformationMtp> info(new (std::nothrow) FilesInformationMtp());
    if (! info)
    {
        return E_OUTOFMEMORY;
    }

    hr = info->BuildFromEntries(std::move(entries));
    if (FAILED(hr))
    {
        return hr;
    }

    *ppFilesInformation = info.release();
    Debug::Perf::EmitValue(L"mtp.enum.objects", static_cast<uint64_t>(items.size()), S_OK);
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemMtp::GetCapabilities(const char** jsonUtf8) noexcept
{
    if (! jsonUtf8)
    {
        return E_POINTER;
    }

    std::lock_guard lock(_jsonMutex);
    *jsonUtf8 = StoreJson(_capabilitiesJson, BuildCapabilitiesJson());
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemMtp::GetTransferHints(const wchar_t* path,
                                                          FileSystemOperation operationType,
                                                          FileSystemTransferEndpoint endpoint,
                                                          FileSystemTransferHints* hints) noexcept
{
    static_cast<void>(path);
    static_cast<void>(operationType);
    static_cast<void>(endpoint);

    if (! hints)
    {
        return E_POINTER;
    }
    if (hints->sizeBytes != sizeof(FileSystemTransferHints))
    {
        return E_INVALIDARG;
    }

    hints->latencyClass = FILESYSTEM_TRANSFER_LATENCY_CLOUD;
    hints->flags =
        FILESYSTEM_TRANSFER_HINT_PREFERS_LARGE_BUFFERS | FILESYSTEM_TRANSFER_HINT_PREFERS_SEQUENTIAL_IO | FILESYSTEM_TRANSFER_HINT_HIGH_METADATA_COST;
    hints->preferredBufferBytes      = 1024u * 1024u;
    hints->preferredProgressPeriodMs = 250u;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemMtp::GetStorageCharacteristics(const wchar_t* path, FileSystemStorageCharacteristics* characteristics) noexcept
{
    static_cast<void>(path);

    if (! characteristics)
    {
        return E_POINTER;
    }
    if (characteristics->sizeBytes != sizeof(FileSystemStorageCharacteristics))
    {
        return E_INVALIDARG;
    }

    characteristics->storageKind                  = FILESYSTEM_STORAGE_VIRTUAL;
    characteristics->flags                        = FILESYSTEM_STORAGE_FLAG_HIGH_LATENCY | FILESYSTEM_STORAGE_FLAG_PREFERS_SEQUENTIAL_IO;
    characteristics->queueDepthHint               = 1;
    characteristics->preferredCopyMoveConcurrency = 1;
    characteristics->preferredDeleteConcurrency   = 1;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemMtp::GetAttributes(const wchar_t* path, unsigned long* fileAttributes) noexcept
{
    if (! fileAttributes)
    {
        return E_POINTER;
    }

    std::wstring normalized;
    HRESULT hr = NormalizeInputPath(path, normalized);
    if (FAILED(hr))
    {
        return hr;
    }
    hr = CheckConnected();
    if (FAILED(hr))
    {
        return hr;
    }

    auto attributesResult          = std::make_shared<unsigned long>(0);
    const std::wstring commandPath = normalized;
    hr = RunBackendCommand([commandPath, attributesResult](IMtpBackend& backend) noexcept { return backend.GetAttributes(commandPath, *attributesResult); },
                           OverwriteJournalIdentityForPath(commandPath));
    if (FAILED(hr))
    {
        return hr;
    }

    *fileAttributes = *attributesResult;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemMtp::CreateFileReader(const wchar_t* path, IFileReader** reader) noexcept
{
    if (! reader)
    {
        return E_POINTER;
    }
    *reader = nullptr;

    std::wstring normalized;
    HRESULT hr = NormalizeInputPath(path, normalized);
    if (FAILED(hr))
    {
        return hr;
    }
    hr = CheckConnected();
    if (FAILED(hr))
    {
        return hr;
    }

    uint64_t backendGeneration = 0u;
    {
        std::lock_guard lock(_stateMutex);
        backendGeneration = _backendGeneration;
    }
    auto backendReaderResult = std::make_shared<std::shared_ptr<IMtpBackendFileReader>>();
    {
        const std::wstring commandPath = normalized;
        hr                             = RunBackendCommand([commandPath, backendReaderResult](IMtpBackend& backend) noexcept {
            return backend.CreateFileReader(commandPath, *backendReaderResult);
        }, OverwriteJournalIdentityForPath(commandPath));
    }
    if (FAILED(hr))
    {
        return hr;
    }
    if (! *backendReaderResult)
    {
        return E_FAIL;
    }

    auto* instance = new (std::nothrow) MtpBackendReader(this, *backendReaderResult, backendGeneration);
    if (! instance)
    {
        return E_OUTOFMEMORY;
    }

    *reader = instance;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemMtp::CreateFileWriter(const wchar_t* path, FileSystemFlags flags, IFileWriter** writer) noexcept
{
    if (! writer)
    {
        return E_POINTER;
    }
    *writer = nullptr;

    HRESULT hr = CheckMutationAllowed();
    if (FAILED(hr))
    {
        return hr;
    }

    std::wstring normalized;
    hr = NormalizeInputPath(path, normalized);
    if (FAILED(hr))
    {
        return hr;
    }

    auto* instance = new (std::nothrow) MtpBufferedWriter(this, std::move(normalized), flags);
    if (! instance)
    {
        return E_OUTOFMEMORY;
    }

    *writer = instance;
    return S_OK;
}

HRESULT FileSystemMtp::CommitFileWriter(std::wstring_view normalizedPath, FileSystemFlags flags, std::span<const std::byte> bytes) noexcept
{
    HRESULT hr = CheckMutationAllowed();
    if (FAILED(hr))
    {
        return hr;
    }

    Settings settings;
    {
        std::lock_guard lock(_stateMutex);
        settings = _settings;
    }

    const bool allowOverwrite = (flags & FILESYSTEM_FLAG_ALLOW_OVERWRITE) != 0;
    Debug::Perf::EmitValue(L"mtp.writer.stage_bytes", static_cast<uint64_t>(bytes.size()), S_OK);

    const std::wstring commandPath(normalizedPath);
    std::vector<std::byte> commandBytes(bytes.begin(), bytes.end());
    const std::string verifyLevel           = settings.byteVerifyOnOverwrite;
    const bool createdObjectPuidUnsupported = CreatedObjectPuidUnsupported();
    const OverwriteJournalContext journalContext{
        .failWrites     = settings.failOverwriteJournalWrites,
        .deviceIdentity = OverwriteJournalDeviceIdentity(settings.connectionDevicePuid, settings.connectionHost, commandPath),
    };
    auto tempPuidMissing = std::make_shared<std::atomic_bool>(false);
    auto tempPuidPresent = std::make_shared<std::atomic_bool>(false);
    hr                   = RunBackendCommand(
        [commandPath,
         commandBytes = std::move(commandBytes),
         allowOverwrite,
         verifyLevel,
         createdObjectPuidUnsupported,
         journalContext,
         tempPuidMissing,
         tempPuidPresent](IMtpBackend& backend) noexcept
    {
        if (! allowOverwrite)
        {
            return backend.WriteFile(commandPath, commandBytes, false);
        }

        unsigned long existingAttributes = 0;
        HRESULT commandHr                = backend.GetAttributes(commandPath, existingAttributes);
        if (commandHr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) || commandHr == HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND))
        {
            commandHr = backend.WriteFile(commandPath, commandBytes, true);
            if (FAILED(commandHr))
            {
                return commandHr;
            }

            return VerifyCommittedOverwrite(backend, commandPath, verifyLevel, commandBytes);
        }
        if (FAILED(commandHr))
        {
            return commandHr;
        }
        if ((existingAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
        {
            return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        }
        if (createdObjectPuidUnsupported)
        {
            constexpr HRESULT unsupportedHr = HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
            Debug::Perf::EmitValue(L"mtp.overwrite.temp_puid_policy_blocked", 1u, unsupportedHr);
            return unsupportedHr;
        }

        return CommitWriterOverwriteWithTempSwap(backend, commandPath, commandBytes, verifyLevel, journalContext, tempPuidMissing, tempPuidPresent);
    },
        journalContext.deviceIdentity,
        MtpBackendCommandKind::Mutating);
    if (tempPuidPresent->load(std::memory_order_acquire))
    {
        RecordCreatedObjectPuidProbe(true);
    }
    if (tempPuidMissing->load(std::memory_order_acquire))
    {
        RecordCreatedObjectPuidProbe(false);
    }
    return hr;
}

HRESULT STDMETHODCALLTYPE FileSystemMtp::GetFileBasicInformation(const wchar_t* path, FileSystemBasicInformation* info) noexcept
{
    if (! info)
    {
        return E_POINTER;
    }
    if (info->sizeBytes != sizeof(FileSystemBasicInformation))
    {
        return E_INVALIDARG;
    }

    std::wstring normalized;
    HRESULT hr = NormalizeInputPath(path, normalized);
    if (FAILED(hr))
    {
        return hr;
    }
    hr = CheckConnected();
    if (FAILED(hr))
    {
        return hr;
    }

    auto infoResult                = std::make_shared<FileSystemBasicInformation>();
    *infoResult                    = *info;
    const std::wstring commandPath = normalized;
    hr = RunBackendCommand([commandPath, infoResult](IMtpBackend& backend) noexcept { return backend.GetBasicInformation(commandPath, *infoResult); },
                           OverwriteJournalIdentityForPath(commandPath));
    if (FAILED(hr))
    {
        return hr;
    }

    *info = *infoResult;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemMtp::SetFileBasicInformation(const wchar_t* path, const FileSystemBasicInformation* info) noexcept
{
    static_cast<void>(path);
    if (! info)
    {
        return E_POINTER;
    }
    if (info->sizeBytes != sizeof(FileSystemBasicInformation))
    {
        return E_INVALIDARG;
    }

    const HRESULT hr = CheckMutationAllowed();
    if (FAILED(hr))
    {
        return hr;
    }

    return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
}

HRESULT STDMETHODCALLTYPE FileSystemMtp::GetItemProperties(const wchar_t* path, const char** jsonUtf8) noexcept
{
    if (! jsonUtf8)
    {
        return E_POINTER;
    }
    *jsonUtf8 = nullptr;

    std::wstring normalized;
    HRESULT hr = NormalizeInputPath(path, normalized);
    if (FAILED(hr))
    {
        return hr;
    }
    hr = CheckConnected();
    if (FAILED(hr))
    {
        return hr;
    }

    auto jsonResult = std::make_shared<std::string>();
    {
        const std::wstring commandPath = normalized;
        hr = RunBackendCommand([commandPath, jsonResult](IMtpBackend& backend) noexcept { return backend.GetItemProperties(commandPath, *jsonResult); },
                               OverwriteJournalIdentityForPath(commandPath));
    }
    if (FAILED(hr))
    {
        return hr;
    }

    std::lock_guard lock(_jsonMutex);
    *jsonUtf8 = StoreJson(_itemPropertiesJson, std::move(*jsonResult));
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemMtp::CreateDirectory(const wchar_t* path) noexcept
{
    HRESULT hr = CheckMutationAllowed();
    if (FAILED(hr))
    {
        return hr;
    }

    std::wstring normalized;
    hr = NormalizeInputPath(path, normalized);
    if (FAILED(hr))
    {
        return hr;
    }

    const std::wstring commandPath = normalized;
    return RunBackendCommand([commandPath](IMtpBackend& backend) noexcept {
        return backend.CreateDirectory(commandPath);
    }, OverwriteJournalIdentityForPath(commandPath), MtpBackendCommandKind::Mutating);
}

HRESULT STDMETHODCALLTYPE FileSystemMtp::GetDirectorySize(
    const wchar_t* path, FileSystemFlags flags, IFileSystemDirectorySizeCallback* callback, void* cookie, FileSystemDirectorySizeResult* result) noexcept
{
    if (! result)
    {
        return E_POINTER;
    }
    if (result->sizeBytes != sizeof(FileSystemDirectorySizeResult))
    {
        return E_INVALIDARG;
    }

    result->totalBytes     = 0;
    result->fileCount      = 0;
    result->directoryCount = 0;
    result->status         = S_OK;

    std::wstring normalized;
    HRESULT hr = NormalizeInputPath(path, normalized);
    if (FAILED(hr))
    {
        result->status = hr;
        return hr;
    }
    hr = CheckConnected();
    if (FAILED(hr))
    {
        result->status = hr;
        return hr;
    }

    uint64_t scannedEntries = 0;
    hr                      = AccumulateDirectorySize(normalized, (flags & FILESYSTEM_FLAG_RECURSIVE) != 0, callback, cookie, *result, scannedEntries);
    if (SUCCEEDED(hr) && callback)
    {
        hr = callback->DirectorySizeProgress(scannedEntries, result->totalBytes, result->fileCount, result->directoryCount, nullptr, cookie);
    }
    result->status = hr;
    return hr;
}

HRESULT STDMETHODCALLTYPE FileSystemMtp::CopyItem(const wchar_t* sourcePath,
                                                  const wchar_t* destinationPath,
                                                  FileSystemFlags flags,
                                                  const FileSystemOptions* options,
                                                  IFileSystemCallback* callback,
                                                  void* cookie) noexcept
{
    HRESULT hr = CheckMutationAllowed();
    Settings settings;
    {
        std::lock_guard lock(_stateMutex);
        settings = _settings;
    }

    if (SUCCEEDED(hr))
    {
        std::wstring source;
        std::wstring dest;
        const bool allowOverwrite = (flags & FILESYSTEM_FLAG_ALLOW_OVERWRITE) != 0;
        hr                        = NormalizeInputPath(sourcePath, source);
        if (SUCCEEDED(hr))
        {
            hr = NormalizeInputPath(destinationPath, dest);
        }
        if (SUCCEEDED(hr))
        {
            hr = ReportSingleItemStartAndCheckCancel(FILESYSTEM_COPY, sourcePath, destinationPath, options, callback, cookie);
        }
        if (SUCCEEDED(hr))
        {
            const std::wstring commandSource        = source;
            const std::wstring commandDest          = dest;
            const std::string verifyLevel           = settings.byteVerifyOnOverwrite;
            const bool createdObjectPuidUnsupported = CreatedObjectPuidUnsupported();
            const OverwriteJournalContext journalContext{
                .failWrites     = settings.failOverwriteJournalWrites,
                .deviceIdentity = OverwriteJournalDeviceIdentity(settings.connectionDevicePuid, settings.connectionHost, commandDest),
            };
            auto tempPuidMissing = std::make_shared<std::atomic_bool>(false);
            auto tempPuidPresent = std::make_shared<std::atomic_bool>(false);
            hr                   = RunBackendCommand(
                [commandSource, commandDest, allowOverwrite, verifyLevel, createdObjectPuidUnsupported, journalContext, tempPuidMissing, tempPuidPresent](
                    IMtpBackend& backend) mutable noexcept
            {
                return ExecuteDeviceSourceOperation(backend,
                                                    DeviceSourceOperation::Copy,
                                                    commandSource,
                                                    commandDest,
                                                    allowOverwrite,
                                                    verifyLevel,
                                                    createdObjectPuidUnsupported,
                                                    journalContext,
                                                    tempPuidMissing,
                                                    tempPuidPresent);
            },
                journalContext.deviceIdentity,
                MtpBackendCommandKind::Mutating);
            if (tempPuidPresent->load(std::memory_order_acquire))
            {
                RecordCreatedObjectPuidProbe(true);
            }
            if (tempPuidMissing->load(std::memory_order_acquire))
            {
                RecordCreatedObjectPuidProbe(false);
            }
        }
    }

    static_cast<void>(CompleteSingleItem(FILESYSTEM_COPY, 0, sourcePath, destinationPath, hr, options, callback, cookie));
    return hr;
}

HRESULT STDMETHODCALLTYPE FileSystemMtp::MoveItem(const wchar_t* sourcePath,
                                                  const wchar_t* destinationPath,
                                                  FileSystemFlags flags,
                                                  const FileSystemOptions* options,
                                                  IFileSystemCallback* callback,
                                                  void* cookie) noexcept
{
    HRESULT hr = CheckMutationAllowed();
    Settings settings;
    {
        std::lock_guard lock(_stateMutex);
        settings = _settings;
    }

    if (SUCCEEDED(hr))
    {
        std::wstring source;
        std::wstring dest;
        const bool allowOverwrite = (flags & FILESYSTEM_FLAG_ALLOW_OVERWRITE) != 0;
        hr                        = NormalizeInputPath(sourcePath, source);
        if (SUCCEEDED(hr))
        {
            hr = NormalizeInputPath(destinationPath, dest);
        }
        if (SUCCEEDED(hr))
        {
            hr = ReportSingleItemStartAndCheckCancel(FILESYSTEM_MOVE, sourcePath, destinationPath, options, callback, cookie);
        }
        if (SUCCEEDED(hr))
        {
            const std::wstring commandSource        = source;
            const std::wstring commandDest          = dest;
            const std::string verifyLevel           = settings.byteVerifyOnOverwrite;
            const bool createdObjectPuidUnsupported = CreatedObjectPuidUnsupported();
            const OverwriteJournalContext journalContext{
                .failWrites     = settings.failOverwriteJournalWrites,
                .deviceIdentity = OverwriteJournalDeviceIdentity(settings.connectionDevicePuid, settings.connectionHost, commandDest),
            };
            auto tempPuidMissing = std::make_shared<std::atomic_bool>(false);
            auto tempPuidPresent = std::make_shared<std::atomic_bool>(false);
            hr                   = RunBackendCommand(
                [commandSource, commandDest, allowOverwrite, verifyLevel, createdObjectPuidUnsupported, journalContext, tempPuidMissing, tempPuidPresent](
                    IMtpBackend& backend) mutable noexcept
            {
                return ExecuteDeviceSourceOperation(backend,
                                                    DeviceSourceOperation::Move,
                                                    commandSource,
                                                    commandDest,
                                                    allowOverwrite,
                                                    verifyLevel,
                                                    createdObjectPuidUnsupported,
                                                    journalContext,
                                                    tempPuidMissing,
                                                    tempPuidPresent);
            },
                journalContext.deviceIdentity,
                MtpBackendCommandKind::Mutating);
            if (tempPuidPresent->load(std::memory_order_acquire))
            {
                RecordCreatedObjectPuidProbe(true);
            }
            if (tempPuidMissing->load(std::memory_order_acquire))
            {
                RecordCreatedObjectPuidProbe(false);
            }
        }
    }

    static_cast<void>(CompleteSingleItem(FILESYSTEM_MOVE, 0, sourcePath, destinationPath, hr, options, callback, cookie));
    return hr;
}

HRESULT STDMETHODCALLTYPE
FileSystemMtp::DeleteItem(const wchar_t* path, FileSystemFlags flags, const FileSystemOptions* options, IFileSystemCallback* callback, void* cookie) noexcept
{
    HRESULT hr = CheckMutationAllowed();
    if (SUCCEEDED(hr))
    {
        std::wstring normalized;
        hr = NormalizeInputPath(path, normalized);
        if (SUCCEEDED(hr))
        {
            if ((flags & FILESYSTEM_FLAG_USE_RECYCLE_BIN) != 0)
            {
                hr = HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
            }
            else
            {
                const std::wstring commandPath = normalized;
                const bool recursive           = (flags & FILESYSTEM_FLAG_RECURSIVE) != 0;
                hr                             = RunBackendCommand([commandPath, recursive](IMtpBackend& backend) noexcept {
                    return backend.DeleteItem(commandPath, recursive);
                }, OverwriteJournalIdentityForPath(commandPath), MtpBackendCommandKind::Mutating);
            }
        }
    }

    static_cast<void>(CompleteSingleItem(FILESYSTEM_DELETE, 0, path, nullptr, hr, options, callback, cookie));
    return hr;
}

HRESULT STDMETHODCALLTYPE FileSystemMtp::RenameItem(const wchar_t* sourcePath,
                                                    const wchar_t* destinationPath,
                                                    FileSystemFlags flags,
                                                    const FileSystemOptions* options,
                                                    IFileSystemCallback* callback,
                                                    void* cookie) noexcept
{
    HRESULT hr = CheckMutationAllowed();
    Settings settings;
    {
        std::lock_guard lock(_stateMutex);
        settings = _settings;
    }

    if (SUCCEEDED(hr))
    {
        std::wstring source;
        std::wstring dest;
        const bool allowOverwrite = (flags & FILESYSTEM_FLAG_ALLOW_OVERWRITE) != 0;
        hr                        = NormalizeInputPath(sourcePath, source);
        if (SUCCEEDED(hr))
        {
            hr = NormalizeInputPath(destinationPath, dest);
        }
        if (SUCCEEDED(hr))
        {
            const std::wstring commandSource        = source;
            const std::wstring commandDest          = dest;
            const std::string verifyLevel           = settings.byteVerifyOnOverwrite;
            const bool createdObjectPuidUnsupported = CreatedObjectPuidUnsupported();
            const OverwriteJournalContext journalContext{
                .failWrites     = settings.failOverwriteJournalWrites,
                .deviceIdentity = OverwriteJournalDeviceIdentity(settings.connectionDevicePuid, settings.connectionHost, commandDest),
            };
            auto tempPuidMissing = std::make_shared<std::atomic_bool>(false);
            auto tempPuidPresent = std::make_shared<std::atomic_bool>(false);
            hr                   = RunBackendCommand(
                [commandSource, commandDest, allowOverwrite, verifyLevel, createdObjectPuidUnsupported, journalContext, tempPuidMissing, tempPuidPresent](
                    IMtpBackend& backend) mutable noexcept
            {
                return ExecuteDeviceSourceOperation(backend,
                                                    DeviceSourceOperation::Rename,
                                                    commandSource,
                                                    commandDest,
                                                    allowOverwrite,
                                                    verifyLevel,
                                                    createdObjectPuidUnsupported,
                                                    journalContext,
                                                    tempPuidMissing,
                                                    tempPuidPresent);
            },
                journalContext.deviceIdentity,
                MtpBackendCommandKind::Mutating);
            if (tempPuidPresent->load(std::memory_order_acquire))
            {
                RecordCreatedObjectPuidProbe(true);
            }
            if (tempPuidMissing->load(std::memory_order_acquire))
            {
                RecordCreatedObjectPuidProbe(false);
            }
        }
    }

    static_cast<void>(CompleteSingleItem(FILESYSTEM_RENAME, 0, sourcePath, destinationPath, hr, options, callback, cookie));
    return hr;
}

HRESULT STDMETHODCALLTYPE FileSystemMtp::CopyItems(const wchar_t* const* sourcePaths,
                                                   unsigned long count,
                                                   const wchar_t* destinationFolder,
                                                   FileSystemFlags flags,
                                                   const FileSystemOptions* options,
                                                   IFileSystemCallback* callback,
                                                   void* cookie) noexcept
{
    return CopyOrMoveItems(false, sourcePaths, count, destinationFolder, flags, options, callback, cookie);
}

HRESULT STDMETHODCALLTYPE FileSystemMtp::MoveItems(const wchar_t* const* sourcePaths,
                                                   unsigned long count,
                                                   const wchar_t* destinationFolder,
                                                   FileSystemFlags flags,
                                                   const FileSystemOptions* options,
                                                   IFileSystemCallback* callback,
                                                   void* cookie) noexcept
{
    return CopyOrMoveItems(true, sourcePaths, count, destinationFolder, flags, options, callback, cookie);
}

HRESULT STDMETHODCALLTYPE FileSystemMtp::DeleteItems(const wchar_t* const* paths,
                                                     unsigned long count,
                                                     FileSystemFlags flags,
                                                     const FileSystemOptions* options,
                                                     IFileSystemCallback* callback,
                                                     void* cookie) noexcept
{
    if (! paths && count > 0)
    {
        return E_POINTER;
    }

    HRESULT finalHr = S_OK;
    for (unsigned long index = 0; index < count; ++index)
    {
        SingleItemIndexCallback indexedCallback(callback, index);
        IFileSystemCallback* itemCallback = callback ? &indexedCallback : nullptr;
        const HRESULT hr                  = DeleteItem(paths[index], flags, options, itemCallback, cookie);
        if (FAILED(hr) && SUCCEEDED(finalHr))
        {
            finalHr = hr;
        }
    }
    return finalHr;
}

HRESULT STDMETHODCALLTYPE FileSystemMtp::RenameItems(const FileSystemRenamePair* items,
                                                     unsigned long count,
                                                     FileSystemFlags flags,
                                                     const FileSystemOptions* options,
                                                     IFileSystemCallback* callback,
                                                     void* cookie) noexcept
{
    if (! items && count > 0)
    {
        return E_POINTER;
    }

    HRESULT finalHr = S_OK;
    for (unsigned long index = 0; index < count; ++index)
    {
        if (items[index].sizeBytes != sizeof(FileSystemRenamePair))
        {
            finalHr = E_INVALIDARG;
            static_cast<void>(CompleteSingleItem(FILESYSTEM_RENAME, index, items[index].sourcePath, items[index].newName, finalHr, options, callback, cookie));
            continue;
        }

        std::wstring source;
        HRESULT hr = NormalizeInputPath(items[index].sourcePath, source);
        std::wstring dest;
        if (SUCCEEDED(hr))
        {
            dest = JoinPath(ParentPath(source), items[index].newName ? std::wstring_view(items[index].newName) : std::wstring_view());
            SingleItemIndexCallback indexedCallback(callback, index);
            IFileSystemCallback* itemCallback = callback ? &indexedCallback : nullptr;
            hr                                = RenameItem(source.c_str(), dest.c_str(), flags, options, itemCallback, cookie);
        }
        if (FAILED(hr) && SUCCEEDED(finalHr))
        {
            finalHr = hr;
        }
    }
    return finalHr;
}

HRESULT STDMETHODCALLTYPE FileSystemMtp::GetMenuItems(const NavigationMenuItem** items, unsigned int* count) noexcept
{
    if (! items || ! count)
    {
        return E_POINTER;
    }

    *items = nullptr;
    *count = 0;

    std::vector<MtpItem> rootItems;
    auto rootItemsResult = std::make_shared<std::vector<MtpItem>>();
    HRESULT hr           = RunBackendCommand([rootItemsResult](IMtpBackend& backend) noexcept { return backend.EnumerateDirectory(L"/", *rootItemsResult); },
                                             OverwriteJournalIdentityForPath(L"/"));
    if (SUCCEEDED(hr))
    {
        rootItems = std::move(*rootItemsResult);
    }

    std::lock_guard lock(_stateMutex);
    _menuEntries.clear();
    _menuEntryView.clear();

    if (FAILED(hr) || rootItems.empty())
    {
        _menuEntries.push_back(MenuEntry{
            .label = LoadStringResource(g_hInstance, IDS_FILESYSTEMMTP_MENU_NO_DEVICES),
            .flags = NAV_MENU_ITEM_FLAG_DISABLED,
        });
    }
    else
    {
        unsigned int commandId = 1;
        for (const MtpItem& item : rootItems)
        {
            _menuEntries.push_back(MenuEntry{
                .label     = item.name,
                .path      = JoinPath(L"/", item.name),
                .iconPath  = JoinPath(L"/", item.name),
                .flags     = NAV_MENU_ITEM_FLAG_NONE,
                .commandId = commandId,
            });
            ++commandId;
        }
    }

    _menuEntryView.reserve(_menuEntries.size());
    for (const MenuEntry& entry : _menuEntries)
    {
        _menuEntryView.push_back(NavigationMenuItem{
            .flags     = entry.flags,
            .label     = entry.label.empty() ? nullptr : entry.label.c_str(),
            .path      = entry.path.empty() ? nullptr : entry.path.c_str(),
            .iconPath  = entry.iconPath.empty() ? nullptr : entry.iconPath.c_str(),
            .commandId = entry.commandId,
        });
    }

    *items = _menuEntryView.empty() ? nullptr : _menuEntryView.data();
    *count = static_cast<unsigned int>(_menuEntryView.size());
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemMtp::ExecuteMenuCommand(unsigned int commandId) noexcept
{
    INavigationMenuCallback* callback = nullptr;
    void* cookie                      = nullptr;
    std::wstring path;

    {
        std::lock_guard lock(_stateMutex);
        for (const MenuEntry& entry : _menuEntries)
        {
            if (entry.commandId == commandId && ! entry.path.empty())
            {
                callback = _navigationMenuCallback;
                cookie   = _navigationMenuCookie;
                path     = entry.path;
                if (callback)
                {
                    ++_navigationMenuActiveCallbacks;
                }
                break;
            }
        }
    }

    if (! callback || path.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
    }

    auto callbackComplete = wil::scope_exit([&]() noexcept
    {
        std::lock_guard lock(_stateMutex);
        if (_navigationMenuActiveCallbacks > 0u)
        {
            --_navigationMenuActiveCallbacks;
        }
        _navigationMenuCallbackCv.notify_all();
    });

    return callback->NavigationMenuRequestNavigate(path.c_str(), cookie);
}

HRESULT STDMETHODCALLTYPE FileSystemMtp::SetCallback(INavigationMenuCallback* callback, void* cookie) noexcept
{
    std::unique_lock lock(_stateMutex);
    _navigationMenuCallback = callback;
    _navigationMenuCookie   = callback ? cookie : nullptr;
    if (! callback)
    {
        _navigationMenuCallbackCv.wait(lock, [&]() noexcept { return _navigationMenuActiveCallbacks == 0u; });
    }
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemMtp::GetDriveInfo(const wchar_t* path, DriveInfo* info) noexcept
{
    if (! info)
    {
        return E_POINTER;
    }

    std::wstring normalized;
    HRESULT hr = NormalizeInputPath(path ? path : L"/", normalized);
    if (FAILED(hr))
    {
        return hr;
    }

    std::lock_guard lock(_stateMutex);
    const auto segments = SplitPathSegments(normalized);
    _driveDisplayName   = segments.empty() ? std::wstring(L"MTP") : std::wstring(segments.front());
    _driveVolumeLabel   = _driveDisplayName;

    _driveInfo = {
        .flags       = static_cast<DriveInfoFlags>(DRIVE_INFO_FLAG_HAS_DISPLAY_NAME | DRIVE_INFO_FLAG_HAS_VOLUME_LABEL | DRIVE_INFO_FLAG_HAS_FILE_SYSTEM),
        .displayName = _driveDisplayName.c_str(),
        .volumeLabel = _driveVolumeLabel.c_str(),
        .fileSystem  = _driveFileSystem.c_str(),
        .totalBytes  = 0,
        .freeBytes   = 0,
        .usedBytes   = 0,
    };

    *info = _driveInfo;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemMtp::GetDriveMenuItems(const wchar_t* path, const NavigationMenuItem** items, unsigned int* count) noexcept
{
    static_cast<void>(path);
    if (! items || ! count)
    {
        return E_POINTER;
    }

    std::lock_guard lock(_stateMutex);
    _driveMenuEntries.clear();
    _driveMenuEntryView.clear();
    _driveMenuEntries.push_back(MenuEntry{
        .label     = LoadStringResource(g_hInstance, IDS_FILESYSTEMMTP_MENU_PROPERTIES),
        .flags     = NAV_MENU_ITEM_FLAG_DISABLED,
        .commandId = DRIVE_INFO_COMMAND_PROPERTIES,
    });
    _driveMenuEntries.push_back(MenuEntry{
        .label     = LoadStringResource(g_hInstance, IDS_FILESYSTEMMTP_MENU_DISCONNECT),
        .flags     = NAV_MENU_ITEM_FLAG_NONE,
        .commandId = DRIVE_INFO_COMMAND_CLEANUP,
    });

    for (const MenuEntry& entry : _driveMenuEntries)
    {
        _driveMenuEntryView.push_back(NavigationMenuItem{
            .flags     = entry.flags,
            .label     = entry.label.c_str(),
            .path      = nullptr,
            .iconPath  = nullptr,
            .commandId = entry.commandId,
        });
    }

    *items = _driveMenuEntryView.data();
    *count = static_cast<unsigned int>(_driveMenuEntryView.size());
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemMtp::ExecuteDriveMenuCommand(unsigned int commandId, const wchar_t* path) noexcept
{
    static_cast<void>(path);
    if (commandId == DRIVE_INFO_COMMAND_CLEANUP)
    {
        std::shared_ptr<IMtpBackend> backend;
        MtpDetachedBackendWorker detachedWorker;
        {
            std::lock_guard lock(_stateMutex);
            _disconnected = true;
            backend       = _backend;
            if (_backendWorker)
            {
                detachedWorker = _backendWorker->DetachForQuarantine(HRESULT_FROM_WIN32(ERROR_DEVICE_NOT_CONNECTED));
                _backendWorker.reset();
            }
            AbandonBackendSessionLocked(backend);
        }

        QueueBackendCancel(backend);
        if (detachedWorker.worker.joinable())
        {
            std::lock_guard lock(MtpQuarantineMutex());
            MtpQuarantineCommands().emplace_back(std::move(detachedWorker.worker), std::move(detachedWorker.exitStatus));
        }
        return S_OK;
    }

    return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
}

const char* FileSystemMtp::StoreJson(JsonReturnBuffers& buffers, std::string jsonUtf8) noexcept
{
    buffers.active    = static_cast<uint8_t>((buffers.active + 1u) % buffers.slots.size());
    std::string& slot = buffers.slots[buffers.active];
    slot.clear();
    slot.append(jsonUtf8);
    return slot.c_str();
}

std::string FileSystemMtp::BuildConfigurationJson() const
{
    Settings settings;
    {
        std::lock_guard lock(_stateMutex);
        settings = _settings;
    }

    return std::format(R"json({{"version":1,"readOnly":{},"byteVerifyOnOverwrite":"{}","commandTimeoutMs":{},"enumerationPageSize":{}}})json",
                       BoolText(settings.readOnly),
                       settings.byteVerifyOnOverwrite,
                       settings.commandTimeoutMs,
                       settings.enumerationPageSize);
}

std::string FileSystemMtp::BuildCapabilitiesJson() const
{
    Settings settings;
    bool backendSupportsWrite = false;
    bool backendLiveWpd       = false;
    bool disconnected         = false;
    {
        std::lock_guard lock(_stateMutex);
        settings = _settings;
        if (_backend)
        {
            const FileSystemMtpInternal::MtpBackendInfo backendInfo = _backend->GetInfo();
            backendSupportsWrite                                    = backendInfo.supportsWrite;
            backendLiveWpd                                          = backendInfo.liveWpd;
        }
        disconnected = _disconnected;
    }

    const bool read  = ! disconnected;
    const bool write = read && ! settings.readOnly && backendSupportsWrite;
    return std::format(
        R"json(
{{
  "version": 1,
  "operations": {{
    "copy": {},
    "move": {},
    "delete": {},
    "rename": {},
    "properties": {},
    "read": {},
    "write": {},
    "recycleBin": false,
    "randomWrite": false,
    "watch": false,
    "search": false,
    "byteVerifyOnOverwrite": "{}"
  }},
  "concurrency": {{
    "copyMoveMax": 1,
    "deleteMax": 1,
    "deleteRecycleBinMax": 0
  }},
  "crossFileSystem": {{
    "export": {{ "copy": ["*"], "move": [] }},
    "import": {{ "copy": {}, "move": [] }}
  }},
  "pathIdentity": {{
    "version": 1,
    "pathTextStableIdentity": true,
    "componentComparison": "ordinalCaseSensitive",
    "normalization": "none",
    "preferredSeparator": "/",
    "acceptedSeparators": ["/"],
    "casePreserving": true,
    "caseOnlyRename": "supported"
  }},
  "mtp": {{
    "backend": "{}",
    "verifyDeviceStorageOnlyWith": "deviceReread"
  }}
}}
)json",
        BoolText(write),
        BoolText(write),
        BoolText(write),
        BoolText(write),
        BoolText(read),
        BoolText(read),
        BoolText(write),
        settings.byteVerifyOnOverwrite,
        write ? R"(["*"])" : "[]",
        backendLiveWpd ? "wpd" : "fake");
}

HRESULT FileSystemMtp::NormalizeInputPath(const wchar_t* path, std::wstring& normalized) const noexcept
{
    if (! path)
    {
        return E_POINTER;
    }

    const std::wstring input = NormalizeMtpPath(path);

    std::wstring rootPath;
    {
        std::lock_guard lock(_stateMutex);
        rootPath = _rootPath;
    }

    if (rootPath.empty() || rootPath == L"/")
    {
        normalized = input;
        return S_OK;
    }

    if (IsConnectionRootPath(input))
    {
        normalized = JoinRootAndChildPath(rootPath, ConnectionPathSuffix(input));
        return S_OK;
    }

    if (input == rootPath || input.rfind(rootPath + L"/", 0) == 0)
    {
        normalized = input;
        return S_OK;
    }

    normalized = JoinRootAndChildPath(rootPath, input);
    return S_OK;
}

bool FileSystemMtp::MutationsAllowed() const noexcept
{
    std::lock_guard lock(_stateMutex);
    return ! _disconnected && _backend && _backend->GetInfo().supportsWrite && ! _settings.readOnly;
}

HRESULT FileSystemMtp::CheckMutationAllowed() const noexcept
{
    std::lock_guard lock(_stateMutex);
    if (_disconnected)
    {
        return HRESULT_FROM_WIN32(ERROR_DEVICE_NOT_CONNECTED);
    }
    if (_backend && _backend->GetInfo().supportsWrite && ! _settings.readOnly)
    {
        return S_OK;
    }

    return HRESULT_FROM_WIN32(ERROR_WRITE_PROTECT);
}

HRESULT FileSystemMtp::CheckConnected() const noexcept
{
    std::lock_guard lock(_stateMutex);
    return _disconnected ? HRESULT_FROM_WIN32(ERROR_DEVICE_NOT_CONNECTED) : S_OK;
}

bool FileSystemMtp::CreatedObjectPuidUnsupported() const noexcept
{
    std::lock_guard lock(_stateMutex);
    return _createdObjectPuidPolicy == CreatedObjectPuidPolicy::Unsupported;
}

void FileSystemMtp::RecordCreatedObjectPuidProbe(bool supported) noexcept
{
    std::lock_guard lock(_stateMutex);
    _createdObjectPuidPolicy = supported ? CreatedObjectPuidPolicy::Supported : CreatedObjectPuidPolicy::Unsupported;
}

HRESULT FileSystemMtp::CompleteSingleItem(FileSystemOperation operationType,
                                          unsigned long itemIndex,
                                          const wchar_t* sourcePath,
                                          const wchar_t* destinationPath,
                                          HRESULT status,
                                          const FileSystemOptions* options,
                                          IFileSystemCallback* callback,
                                          void* cookie) noexcept
{
    if (! callback)
    {
        return S_OK;
    }

    FileSystemOptions mutableOptions{};
    FileSystemOptions* mutableOptionsPtr = nullptr;
    if (options)
    {
        mutableOptions    = *options;
        mutableOptionsPtr = &mutableOptions;
    }

    return callback->FileSystemItemCompleted(operationType, itemIndex, sourcePath, destinationPath, status, mutableOptionsPtr, cookie);
}

HRESULT FileSystemMtp::ReportSingleItemStartAndCheckCancel(FileSystemOperation operationType,
                                                           const wchar_t* sourcePath,
                                                           const wchar_t* destinationPath,
                                                           const FileSystemOptions* options,
                                                           IFileSystemCallback* callback,
                                                           void* cookie) noexcept
{
    if (! callback)
    {
        return S_OK;
    }

    FileSystemOptions mutableOptions{};
    FileSystemOptions* mutableOptionsPtr = nullptr;
    if (options)
    {
        mutableOptions    = *options;
        mutableOptionsPtr = &mutableOptions;
    }

    const HRESULT progressHr = callback->FileSystemProgress(operationType, 1, 0, 0, 0, sourcePath, destinationPath, 0, 0, mutableOptionsPtr, 0, cookie);
    if (FAILED(progressHr))
    {
        return progressHr;
    }

    BOOL cancel            = FALSE;
    const HRESULT cancelHr = callback->FileSystemShouldCancel(&cancel, cookie);
    if (FAILED(cancelHr))
    {
        return cancelHr;
    }

    return cancel != FALSE ? HRESULT_FROM_WIN32(ERROR_CANCELLED) : S_OK;
}

HRESULT FileSystemMtp::CopyOrMoveItems(bool move,
                                       const wchar_t* const* sourcePaths,
                                       unsigned long count,
                                       const wchar_t* destinationFolder,
                                       FileSystemFlags flags,
                                       const FileSystemOptions* options,
                                       IFileSystemCallback* callback,
                                       void* cookie) noexcept
{
    if (! sourcePaths && count > 0)
    {
        return E_POINTER;
    }
    if (! destinationFolder)
    {
        return E_POINTER;
    }

    std::wstring destFolder;
    HRESULT finalHr = NormalizeInputPath(destinationFolder, destFolder);
    if (FAILED(finalHr))
    {
        return finalHr;
    }

    for (unsigned long index = 0; index < count; ++index)
    {
        std::wstring source;
        HRESULT hr = NormalizeInputPath(sourcePaths[index], source);
        std::wstring dest;
        if (SUCCEEDED(hr))
        {
            dest = JoinPath(destFolder, LeafName(source));
            SingleItemIndexCallback indexedCallback(callback, index);
            IFileSystemCallback* itemCallback = callback ? &indexedCallback : nullptr;
            hr                                = move ? MoveItem(source.c_str(), dest.c_str(), flags, options, itemCallback, cookie)
                                                     : CopyItem(source.c_str(), dest.c_str(), flags, options, itemCallback, cookie);
        }

        if (FAILED(hr) && SUCCEEDED(finalHr))
        {
            finalHr = hr;
        }
    }

    return finalHr;
}

HRESULT FileSystemMtp::AccumulateDirectorySize(std::wstring_view path,
                                               bool recursive,
                                               IFileSystemDirectorySizeCallback* callback,
                                               void* cookie,
                                               FileSystemDirectorySizeResult& result,
                                               uint64_t& scannedEntries) noexcept
{
    HRESULT hr = CheckConnected();
    if (FAILED(hr))
    {
        return hr;
    }

    unsigned long attrs = 0;
    {
        auto attrsResult = std::make_shared<unsigned long>(0);
        const std::wstring commandPath(path);
        hr = RunBackendCommand([commandPath, attrsResult](IMtpBackend& backend) noexcept { return backend.GetAttributes(commandPath, *attrsResult); },
                               OverwriteJournalIdentityForPath(commandPath));
        if (SUCCEEDED(hr))
        {
            attrs = *attrsResult;
        }
    }
    if (FAILED(hr))
    {
        return hr;
    }

    if (callback)
    {
        const std::wstring currentPath(path);
        hr = callback->DirectorySizeProgress(scannedEntries, result.totalBytes, result.fileCount, result.directoryCount, currentPath.c_str(), cookie);
        if (FAILED(hr))
        {
            return hr;
        }

        BOOL cancel            = FALSE;
        const HRESULT cancelHr = callback->DirectorySizeShouldCancel(&cancel, cookie);
        if (FAILED(cancelHr))
        {
            return cancelHr;
        }
        if (cancel)
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
    }

    if ((attrs & FILE_ATTRIBUTE_DIRECTORY) == 0)
    {
        hr = CheckConnected();
        if (FAILED(hr))
        {
            return hr;
        }

        uint64_t fileSizeBytes = 0;
        {
            auto sizeResult = std::make_shared<uint64_t>(0);
            const std::wstring commandPath(path);
            hr = RunBackendCommand([commandPath, sizeResult](IMtpBackend& backend) noexcept { return backend.GetFileSize(commandPath, *sizeResult); },
                                   OverwriteJournalIdentityForPath(commandPath));
            if (SUCCEEDED(hr))
            {
                fileSizeBytes = *sizeResult;
            }
        }
        if (FAILED(hr))
        {
            return hr;
        }

        result.totalBytes += fileSizeBytes;
        ++result.fileCount;
        ++scannedEntries;
        return callback ? callback->DirectorySizeProgress(
                              scannedEntries, result.totalBytes, result.fileCount, result.directoryCount, std::wstring(path).c_str(), cookie)
                        : S_OK;
    }

    hr = CheckConnected();
    if (FAILED(hr))
    {
        return hr;
    }

    std::vector<MtpItem> children;
    {
        auto childrenResult = std::make_shared<std::vector<MtpItem>>();
        const std::wstring commandPath(path);
        hr = RunBackendCommand([commandPath, childrenResult](IMtpBackend& backend) noexcept {
            return backend.EnumerateDirectory(commandPath, *childrenResult);
        }, OverwriteJournalIdentityForPath(commandPath));
        if (SUCCEEDED(hr))
        {
            children = std::move(*childrenResult);
        }
    }
    if (FAILED(hr))
    {
        return hr;
    }

    for (const MtpItem& child : children)
    {
        const std::wstring childPath = JoinPath(path, child.name);
        ++scannedEntries;
        if ((child.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
        {
            ++result.directoryCount;
            if (recursive)
            {
                hr = AccumulateDirectorySize(childPath, recursive, callback, cookie, result, scannedEntries);
                if (FAILED(hr))
                {
                    return hr;
                }
            }
        }
        else
        {
            result.totalBytes += child.sizeBytes;
            ++result.fileCount;
        }

        if (callback)
        {
            hr = callback->DirectorySizeProgress(scannedEntries, result.totalBytes, result.fileCount, result.directoryCount, childPath.c_str(), cookie);
            if (FAILED(hr))
            {
                return hr;
            }
        }
    }

    return S_OK;
}

[[nodiscard]] const char* GetFileSystemMtpStaticConfigurationSchema() noexcept
{
    return FileSystemMtp::StaticConfigurationSchema();
}

void ShutdownFileSystemMtpModule() noexcept
{
    SweepMtpQuarantinedCommands();
}

bool CanUnloadFileSystemMtpModule() noexcept
{
    SweepMtpQuarantinedCommands();

    if (MtpPendingCancelRequests().load(std::memory_order_acquire) != 0u)
    {
        return false;
    }

    std::lock_guard lock(MtpQuarantineMutex());
    return MtpQuarantineCommands().empty();
}

bool RetainFileSystemMtpModuleUntilProcessExit() noexcept
{
    return ! CanUnloadFileSystemMtpModule();
}
