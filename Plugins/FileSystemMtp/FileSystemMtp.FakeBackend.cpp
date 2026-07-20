#include "FileSystemMtp.h"

#include <algorithm>
#include <array>
#include <chrono>
#include <condition_variable>
#include <cstdint>
#include <cstring>
#include <cwctype>
#include <format>
#include <limits>
#include <thread>
#include <unordered_map>
#include <utility>

#pragma warning(push)
// (C6297) Arithmetic overflow. Results might not be an expected value.
// (C28182) Dereferencing NULL pointer.
#pragma warning(disable : 6297 28182)
#include <yyjson.h>
#pragma warning(pop)

#include "Helpers.h"

namespace FileSystemMtpInternal
{
namespace
{
struct FakeNode
{
    unsigned long attributes = 0;
    uint64_t sizeBytes       = 0;
    __int64 creationTime     = 0;
    __int64 lastAccessTime   = 0;
    __int64 lastWriteTime    = 0;
    __int64 changeTime       = 0;
    std::wstring persistentId;
    std::wstring objectId;
    std::wstring displayName;
    std::vector<std::byte> content;
};

struct FakeMtpReaderStats
{
    FakeMtpReaderStats()  = default;
    ~FakeMtpReaderStats() = default;

    FakeMtpReaderStats(const FakeMtpReaderStats&)            = delete;
    FakeMtpReaderStats(FakeMtpReaderStats&&)                 = delete;
    FakeMtpReaderStats& operator=(const FakeMtpReaderStats&) = delete;
    FakeMtpReaderStats& operator=(FakeMtpReaderStats&&)      = delete;

    std::atomic_uint32_t readFileCalls{0};
    std::atomic_uint64_t lastReadBytes{0};
};

[[nodiscard]] bool IsDirectory(const FakeNode& node) noexcept
{
    return (node.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
}

[[nodiscard]] uint32_t ReadUInt32Option(std::string_view optionsJsonUtf8, const char* key, uint32_t defaultValue, uint32_t maxValue) noexcept
{
    if (optionsJsonUtf8.empty() || ! key)
    {
        return defaultValue;
    }

    yyjson_doc* doc = yyjson_read(optionsJsonUtf8.data(), optionsJsonUtf8.size(), YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM);
    if (! doc)
    {
        return defaultValue;
    }

    auto freeDoc     = wil::scope_exit([&]() noexcept { yyjson_doc_free(doc); });
    yyjson_val* root = yyjson_doc_get_root(doc);
    if (! root || ! yyjson_is_obj(root))
    {
        return defaultValue;
    }

    yyjson_val* value = yyjson_obj_get(root, key);
    if (! value || ! yyjson_is_uint(value))
    {
        return defaultValue;
    }

    return static_cast<uint32_t>(std::min<uint64_t>(yyjson_get_uint(value), maxValue));
}

[[nodiscard]] bool ReadBoolOption(std::string_view optionsJsonUtf8, const char* key, bool defaultValue) noexcept
{
    if (optionsJsonUtf8.empty() || ! key)
    {
        return defaultValue;
    }

    yyjson_doc* doc = yyjson_read(optionsJsonUtf8.data(), optionsJsonUtf8.size(), YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM);
    if (! doc)
    {
        return defaultValue;
    }

    auto freeDoc     = wil::scope_exit([&]() noexcept { yyjson_doc_free(doc); });
    yyjson_val* root = yyjson_doc_get_root(doc);
    if (! root || ! yyjson_is_obj(root))
    {
        return defaultValue;
    }

    yyjson_val* value = yyjson_obj_get(root, key);
    if (! value || ! yyjson_is_bool(value))
    {
        return defaultValue;
    }

    return yyjson_get_bool(value) != 0;
}

[[nodiscard]] std::wstring ReadStringOptionWide(std::string_view optionsJsonUtf8, const char* key) noexcept
{
    if (optionsJsonUtf8.empty() || ! key)
    {
        return {};
    }

    yyjson_doc* doc = yyjson_read(optionsJsonUtf8.data(), optionsJsonUtf8.size(), YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM);
    if (! doc)
    {
        return {};
    }

    auto freeDoc     = wil::scope_exit([&]() noexcept { yyjson_doc_free(doc); });
    yyjson_val* root = yyjson_doc_get_root(doc);
    if (! root || ! yyjson_is_obj(root))
    {
        return {};
    }

    yyjson_val* value = yyjson_obj_get(root, key);
    if (! value || ! yyjson_is_str(value))
    {
        return {};
    }

    const char* text = yyjson_get_str(value);
    const size_t len = yyjson_get_len(value);
    if (! text)
    {
        return {};
    }

    return Utf16FromUtf8(std::string_view(text, len));
}

void DisambiguateDuplicateNames(std::vector<MtpItem>& items)
{
    std::unordered_map<std::wstring, uint32_t> counts;
    counts.reserve(items.size());

    for (const MtpItem& item : items)
    {
        std::wstring key = item.name;
        std::transform(key.begin(), key.end(), key.begin(), [](wchar_t ch) { return static_cast<wchar_t>(::towlower(ch)); });
        ++counts[key];
    }

    uint64_t duplicateGroups = 0;
    uint64_t suffixedEntries = 0;
    for (const auto& [key, count] : counts)
    {
        static_cast<void>(key);
        if (count > 1u)
        {
            ++duplicateGroups;
            suffixedEntries += count;
        }
    }

    for (MtpItem& item : items)
    {
        std::wstring key = item.name;
        std::transform(key.begin(), key.end(), key.begin(), [](wchar_t ch) { return static_cast<wchar_t>(::towlower(ch)); });
        if (counts[key] > 1u)
        {
            item.name += MtpDuplicateObjectSuffix(item);
        }
    }

    Debug::Perf::EmitValue(L"mtp.path.duplicate_groups", duplicateGroups, S_OK);
    Debug::Perf::EmitValue(L"mtp.path.suffixed_entries", suffixedEntries, S_OK);
}

class FakeMtpBackend;

class FakeMtpActiveCallScope final
{
public:
    explicit FakeMtpActiveCallScope(FakeMtpBackend& owner) noexcept;

    FakeMtpActiveCallScope(const FakeMtpActiveCallScope&)            = delete;
    FakeMtpActiveCallScope(FakeMtpActiveCallScope&&)                 = delete;
    FakeMtpActiveCallScope& operator=(const FakeMtpActiveCallScope&) = delete;
    FakeMtpActiveCallScope& operator=(FakeMtpActiveCallScope&&)      = delete;

    ~FakeMtpActiveCallScope() noexcept;

private:
    FakeMtpBackend& _owner;
};

class FakeMtpBackend final : public IMtpBackend
{
public:
    explicit FakeMtpBackend(std::string_view optionsJsonUtf8)
    {
        _operationDelayMs                   = ReadUInt32Option(optionsJsonUtf8, "operationDelayMs", 0u, 5'000u);
        _readFileDelayMs                    = ReadUInt32Option(optionsJsonUtf8, "readFileDelayMs", 0u, 30'000u);
        _cancelUnblocksDelay                = ReadBoolOption(optionsJsonUtf8, "cancelUnblocksDelay", false);
        _moveFallbackDeleteSourceFails      = ReadBoolOption(optionsJsonUtf8, "moveFallbackDeleteSourceFails", false);
        _omitPersistentIdForCreatedFiles    = ReadBoolOption(optionsJsonUtf8, "omitPersistentIdForCreatedFiles", false);
        std::wstring deleteItemFailOncePath = ReadStringOptionWide(optionsJsonUtf8, "deleteItemFailOncePath");
        if (! deleteItemFailOncePath.empty())
        {
            _deleteItemFailOncePath = NormalizeMtpPath(deleteItemFailOncePath);
        }
        std::wstring renameItemFailOnceDestinationPath = ReadStringOptionWide(optionsJsonUtf8, "renameItemFailOnceDestinationPath");
        if (! renameItemFailOnceDestinationPath.empty())
        {
            _renameItemFailOnceDestinationPath = NormalizeMtpPath(renameItemFailOnceDestinationPath);
        }
        std::wstring renameItemFailDestinationPath = ReadStringOptionWide(optionsJsonUtf8, "renameItemFailDestinationPath");
        if (! renameItemFailDestinationPath.empty())
        {
            _renameItemFailDestinationPath = NormalizeMtpPath(renameItemFailDestinationPath);
        }
        _writeFileFailOncePathContains           = ReadStringOptionWide(optionsJsonUtf8, "writeFileFailOncePathContains");
        _copyItemFailOnceDestinationPathContains = ReadStringOptionWide(optionsJsonUtf8, "copyItemFailOnceDestinationPathContains");
        std::wstring disconnectEnumerateOncePath = ReadStringOptionWide(optionsJsonUtf8, "disconnectEnumerateOncePath");
        if (! disconnectEnumerateOncePath.empty())
        {
            _disconnectEnumerateOncePath = NormalizeMtpPath(disconnectEnumerateOncePath);
        }
        const __int64 now = NowFileTime64();
        AddDirectory(L"/", L"root", now);
        AddDirectory(L"/Fake Phone [devpuid:fake-device]", L"dev-fake-phone", now);
        AddDirectory(L"/Fake Phone [devpuid:fake-device]/Internal Storage", L"storage-internal", now);
        AddDirectory(L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM", L"folder-dcim", now);
        AddDirectory(L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera", L"folder-camera", now);

        const char sample[] = "RedSalamander deterministic MTP fixture\r\n";
        std::vector<std::byte> bytes(sizeof(sample) - 1u);
        std::memcpy(bytes.data(), sample, bytes.size());
        AddFile(L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera/photo001.txt", L"file-photo001", bytes, now);

        const char duplicate[] = "duplicate fixture\r\n";
        std::vector<std::byte> duplicateBytes(sizeof(duplicate) - 1u);
        std::memcpy(duplicateBytes.data(), duplicate, duplicateBytes.size());
        AddFile(L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera/name [puid:literal].txt", L"file-literal-suffix", duplicateBytes, now);

        const auto duplicatePathLeaf = [](std::wstring_view persistentId)
        {
            MtpItem item;
            item.persistentId = std::wstring(persistentId);
            return std::wstring(L"duplicate-sibling.txt") + MtpDuplicateObjectSuffix(item);
        };

        const char duplicateOne[] = "duplicate sibling one\r\n";
        std::vector<std::byte> duplicateOneBytes(sizeof(duplicateOne) - 1u);
        std::memcpy(duplicateOneBytes.data(), duplicateOne, duplicateOneBytes.size());
        AddFile(L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera/" + duplicatePathLeaf(L"file-duplicate-one"),
                L"file-duplicate-one",
                duplicateOneBytes,
                now,
                L"duplicate-sibling.txt");

        const char duplicateTwo[] = "duplicate sibling two\r\n";
        std::vector<std::byte> duplicateTwoBytes(sizeof(duplicateTwo) - 1u);
        std::memcpy(duplicateTwoBytes.data(), duplicateTwo, duplicateTwoBytes.size());
        AddFile(L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera/" + duplicatePathLeaf(L"file-duplicate-two"),
                L"file-duplicate-two",
                duplicateTwoBytes,
                now,
                L"duplicate-sibling.txt");

        const char special[] = "special name fixture\r\n";
        std::vector<std::byte> specialBytes(sizeof(special) - 1u);
        std::memcpy(specialBytes.data(), special, specialBytes.size());
        AddFile(L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera/percent %.txt", L"file-percent", specialBytes, now);
        AddFile(L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera/bracket ].txt", L"file-bracket", specialBytes, now);
        AddFile(L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera/trailing-space .txt", L"file-trailing-space", specialBytes, now);
        AddFile(L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera/caf\u00E9.txt", L"file-nonascii", specialBytes, now);
    }

    FakeMtpBackend(const FakeMtpBackend&)            = delete;
    FakeMtpBackend(FakeMtpBackend&&)                 = delete;
    FakeMtpBackend& operator=(const FakeMtpBackend&) = delete;
    FakeMtpBackend& operator=(FakeMtpBackend&&)      = delete;

    MtpBackendInfo GetInfo() const noexcept override
    {
        return {.readOnly = false, .supportsWrite = true, .liveWpd = false};
    }

    void RequestCancel() noexcept override
    {
        static_cast<void>(_cancelRequests.fetch_add(1u, std::memory_order_acq_rel));
        Debug::Perf::EmitValue(L"mtp.device.cancel_requests", 1u, S_OK);

        if (_cancelUnblocksDelay)
        {
            {
                std::lock_guard lock(_delayMutex);
                _cancelRequested = true;
            }
            _delayCv.notify_all();
        }
    }

    HRESULT EnumerateDirectory(std::wstring_view path, std::vector<MtpItem>& items) noexcept override
    {
        const FakeMtpActiveCallScope activeCall(*this);
        std::lock_guard lock(_mutex);
        items.clear();

        const std::wstring normalized = NormalizeMtpPath(path);
        if (! _disconnectEnumerateOncePath.empty() && ! _disconnectEnumerateOnceConsumed && normalized == _disconnectEnumerateOncePath)
        {
            _disconnectEnumerateOnceConsumed = true;
            constexpr HRESULT disconnectedHr = HRESULT_FROM_WIN32(ERROR_DEVICE_NOT_CONNECTED);
            Debug::Perf::EmitValue(L"mtp.device.disconnects", 1u, disconnectedHr);
            return disconnectedHr;
        }

        const auto it = _nodes.find(normalized);
        if (it == _nodes.end())
        {
            return HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND);
        }
        if (! IsDirectory(it->second))
        {
            return HRESULT_FROM_WIN32(ERROR_DIRECTORY);
        }

        for (const auto& [childPath, child] : _nodes)
        {
            if (childPath == normalized)
            {
                continue;
            }
            if (ParentPath(childPath) != normalized)
            {
                continue;
            }

            items.push_back(MtpItem{
                .name           = child.displayName.empty() ? LeafName(childPath) : child.displayName,
                .attributes     = child.attributes,
                .sizeBytes      = child.sizeBytes,
                .creationTime   = child.creationTime,
                .lastAccessTime = child.lastAccessTime,
                .lastWriteTime  = child.lastWriteTime,
                .changeTime     = child.changeTime,
                .persistentId   = child.persistentId,
                .objectId       = child.objectId,
                .streamable     = ! IsDirectory(child),
            });
        }

        DisambiguateDuplicateNames(items);

        static_cast<void>(_propertyBatchCalls.fetch_add(1u, std::memory_order_acq_rel));
        Debug::Perf::EmitValue(L"mtp.props.bulk_batches", 1u, S_OK);
        Debug::Perf::EmitValue(L"mtp.props.per_item_calls", 0u, S_OK);
        return S_OK;
    }

    HRESULT GetAttributes(std::wstring_view path, unsigned long& attributes) noexcept override
    {
        const FakeMtpActiveCallScope activeCall(*this);
        std::lock_guard lock(_mutex);
        const FakeNode* node = FindNodeLocked(path);
        if (! node)
        {
            return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
        }

        attributes = node->attributes;
        return S_OK;
    }

    HRESULT GetBasicInformation(std::wstring_view path, FileSystemBasicInformation& info) noexcept override
    {
        const FakeMtpActiveCallScope activeCall(*this);
        std::lock_guard lock(_mutex);
        const FakeNode* node = FindNodeLocked(path);
        if (! node)
        {
            return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
        }

        info.creationTime   = node->creationTime;
        info.lastAccessTime = node->lastAccessTime;
        info.lastWriteTime  = node->lastWriteTime;
        info.attributes     = node->attributes;
        return S_OK;
    }

    HRESULT ReadFile(std::wstring_view path, std::vector<std::byte>& bytes) noexcept override
    {
        const FakeMtpActiveCallScope activeCall(*this);
        static_cast<void>(_readerStats->readFileCalls.fetch_add(1u, std::memory_order_acq_rel));
        if (_readFileDelayMs != 0u)
        {
            std::this_thread::sleep_for(std::chrono::milliseconds(_readFileDelayMs));
        }

        std::lock_guard lock(_mutex);
        const FakeNode* node = FindNodeLocked(path);
        if (! node)
        {
            return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
        }
        if (IsDirectory(*node))
        {
            return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        }

        bytes = node->content;
        _readerStats->lastReadBytes.store(static_cast<uint64_t>(bytes.size()), std::memory_order_release);
        return S_OK;
    }

    HRESULT CreateFileReader(std::wstring_view path, std::shared_ptr<IMtpBackendFileReader>& reader) noexcept override
    {
        reader.reset();

        const FakeMtpActiveCallScope activeCall(*this);
        std::lock_guard lock(_mutex);
        const FakeNode* node = FindNodeLocked(path);
        if (! node)
        {
            return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
        }
        if (IsDirectory(*node))
        {
            return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        }

        reader = CreateMemoryBackendFileReader(node->content,
                                               _readFileDelayMs,
                                               _readerStats,
                                               [](void* context, uint64_t bytesRead) noexcept
        {
            auto* stats = static_cast<FakeMtpReaderStats*>(context);
            if (stats)
            {
                static_cast<void>(stats->readFileCalls.fetch_add(1u, std::memory_order_acq_rel));
                stats->lastReadBytes.store(bytesRead, std::memory_order_release);
            }
        });
        return S_OK;
    }

    HRESULT GetFileSize(std::wstring_view path, uint64_t& sizeBytes) noexcept override
    {
        const FakeMtpActiveCallScope activeCall(*this);
        static_cast<void>(_fileSizeCalls.fetch_add(1u, std::memory_order_acq_rel));
        std::lock_guard lock(_mutex);
        const FakeNode* node = FindNodeLocked(path);
        if (! node)
        {
            return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
        }
        if (IsDirectory(*node))
        {
            return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        }

        sizeBytes = node->sizeBytes;
        return S_OK;
    }

    HRESULT WriteFile(std::wstring_view path, std::span<const std::byte> bytes, bool allowOverwrite) noexcept override
    {
        const FakeMtpActiveCallScope activeCall(*this);
        static_cast<void>(_writeFileCalls.fetch_add(1u, std::memory_order_acq_rel));
        std::lock_guard lock(_mutex);
        const std::wstring normalized = NormalizeMtpPath(path);
        if (! _writeFileFailOncePathContains.empty() && ! _writeFileFailOnceConsumed && normalized.find(_writeFileFailOncePathContains) != std::wstring::npos)
        {
            _writeFileFailOnceConsumed = true;
            return HRESULT_FROM_WIN32(ERROR_WRITE_FAULT);
        }

        const std::wstring parent = ParentPath(normalized);
        FakeNode* parentNode      = FindNodeLocked(parent);
        if (! parentNode || ! IsDirectory(*parentNode))
        {
            return HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND);
        }

        FakeNode* existing = FindNodeLocked(normalized);
        if (existing && IsDirectory(*existing))
        {
            return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        }
        if (existing && ! allowOverwrite)
        {
            return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
        }

        const __int64 now = NowFileTime64();
        FakeNode node;
        if (existing)
        {
            node = *existing;
        }
        else
        {
            node.creationTime = now;
            node.persistentId = _omitPersistentIdForCreatedFiles ? std::wstring() : L"puid-" + std::to_wstring(_nextObjectId);
            node.objectId     = L"oid-" + std::to_wstring(_nextObjectId);
            node.displayName  = LeafName(normalized);
            ++_nextObjectId;
        }

        node.attributes     = FILE_ATTRIBUTE_NORMAL;
        node.lastAccessTime = now;
        node.lastWriteTime  = now;
        node.changeTime     = now;
        node.sizeBytes      = static_cast<uint64_t>(bytes.size());
        node.content.assign(bytes.begin(), bytes.end());
        _nodes[normalized] = std::move(node);
        return S_OK;
    }

    HRESULT CreateDirectory(std::wstring_view path) noexcept override
    {
        const FakeMtpActiveCallScope activeCall(*this);
        std::lock_guard lock(_mutex);
        const std::wstring normalized = NormalizeMtpPath(path);
        if (_nodes.contains(normalized))
        {
            return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
        }

        const std::wstring parent  = ParentPath(normalized);
        const FakeNode* parentNode = FindNodeLocked(parent);
        if (! parentNode || ! IsDirectory(*parentNode))
        {
            return HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND);
        }

        AddDirectoryLocked(normalized, L"puid-" + std::to_wstring(_nextObjectId), NowFileTime64());
        ++_nextObjectId;
        return S_OK;
    }

    HRESULT DeleteItem(std::wstring_view path, bool recursive) noexcept override
    {
        const FakeMtpActiveCallScope activeCall(*this);
        std::lock_guard lock(_mutex);
        const std::wstring normalized = NormalizeMtpPath(path);
        const auto it                 = _nodes.find(normalized);
        if (it == _nodes.end())
        {
            if (HasAmbiguousDisplayNameLocked(normalized))
            {
                Debug::Perf::EmitValue(L"mtp.path.ambiguous_resolve_failures", 1u, HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND));
            }
            return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
        }
        if (normalized == L"/")
        {
            return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        }
        if (! _deleteItemFailOncePath.empty() && ! _deleteItemFailOnceConsumed && normalized == _deleteItemFailOncePath)
        {
            _deleteItemFailOnceConsumed = true;
            return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        }
        if (IsDirectory(it->second) && ! recursive && HasChildrenLocked(normalized))
        {
            return HRESULT_FROM_WIN32(ERROR_DIR_NOT_EMPTY);
        }

        EraseTreeLocked(normalized);
        return S_OK;
    }

    HRESULT RenameItem(std::wstring_view sourcePath, std::wstring_view destinationPath, bool allowOverwrite) noexcept override
    {
        const FakeMtpActiveCallScope activeCall(*this);
        std::lock_guard lock(_mutex);
        return RenameItemLocked(sourcePath, destinationPath, allowOverwrite);
    }

    HRESULT CopyItem(std::wstring_view sourcePath, std::wstring_view destinationPath, bool allowOverwrite) noexcept override
    {
        const FakeMtpActiveCallScope activeCall(*this);
        static_cast<void>(_copyItemCalls.fetch_add(1u, std::memory_order_acq_rel));
        std::lock_guard lock(_mutex);
        uint64_t copiedBytes = 0;
        const HRESULT hr     = CopyItemLocked(sourcePath, destinationPath, allowOverwrite, &copiedBytes);
        if (SUCCEEDED(hr))
        {
            _lastCopyBytes.store(copiedBytes, std::memory_order_release);
            Debug::Perf::EmitValue(L"mtp.transfer.copy_bytes", copiedBytes, S_OK);
        }
        return hr;
    }

    HRESULT MoveItem(std::wstring_view sourcePath, std::wstring_view destinationPath, bool allowOverwrite) noexcept override
    {
        const FakeMtpActiveCallScope activeCall(*this);
        static_cast<void>(allowOverwrite);
        std::lock_guard lock(_mutex);
        const std::wstring source = NormalizeMtpPath(sourcePath);
        const std::wstring dest   = NormalizeMtpPath(destinationPath);
        if (source == L"/" || dest == L"/")
        {
            return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        }

        const auto sourceIt = _nodes.find(source);
        if (sourceIt == _nodes.end())
        {
            return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
        }

        const std::wstring parent  = ParentPath(dest);
        const FakeNode* parentNode = FindNodeLocked(parent);
        if (! parentNode || ! IsDirectory(*parentNode))
        {
            return HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND);
        }
        if (_nodes.contains(dest))
        {
            return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
        }

        const bool sourceDirectory    = IsDirectory(sourceIt->second);
        const std::wstring sourceLeaf = LeafName(source);
        const std::wstring destLeaf   = LeafName(dest);
        const bool preservesLeaf      = EqualsPathComponent(sourceLeaf, destLeaf);

        // Mirror WpdMtpBackend::MoveItem: directories have no transfer fallback when the native leaf-preserving path is unavailable.
        if (sourceDirectory && ! preservesLeaf)
        {
            return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        }
        if (sourceDirectory || preservesLeaf)
        {
            return RenameItemLocked(source, dest, false);
        }

        if (_moveFallbackDeleteSourceFails)
        {
            uint64_t copiedBytes = 0;
            const HRESULT copyHr = CopyItemLocked(source, dest, false, &copiedBytes);
            if (FAILED(copyHr))
            {
                return copyHr;
            }

            _lastMoveFallbackBytes.store(copiedBytes, std::memory_order_release);
            Debug::Perf::EmitValue(L"mtp.transfer.move_fallback_bytes", copiedBytes, HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED));
            return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        }

        uint64_t copiedBytes = 0;
        const HRESULT copyHr = CopyItemLocked(source, dest, false, &copiedBytes);
        if (FAILED(copyHr))
        {
            return copyHr;
        }

        EraseTreeLocked(source);
        _lastMoveFallbackBytes.store(copiedBytes, std::memory_order_release);
        Debug::Perf::EmitValue(L"mtp.transfer.move_fallback_bytes", copiedBytes, S_OK);
        return S_OK;
    }

    HRESULT GetItemProperties(std::wstring_view path, std::string& jsonUtf8) noexcept override
    {
        const FakeMtpActiveCallScope activeCall(*this);
        std::lock_guard lock(_mutex);
        const std::wstring normalized = NormalizeMtpPath(path);
        const FakeNode* node          = FindNodeLocked(normalized);
        if (! node)
        {
            return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
        }

        const std::string nameUtf8 = Utf8FromUtf16(LeafName(normalized));
        const std::string puidUtf8 = Utf8FromUtf16(node->persistentId);
        const auto [backendThreadIdsObserved, backendThreadIdsOverflow] = BackendThreadStats();
        jsonUtf8                   = std::format(
            R"json({{"version":1,"backend":"fake","name":"{}","persistentId":"{}","streamable":{},"sizeBytes":{},"instrumentation":{{"activeBackendCalls":{},"maxConcurrentBackendCalls":{},"backendThreadIdsObserved":{},"backendThreadIdsOverflow":{},"operationDelayMs":{},"cancelRequests":{},"writeFileCalls":{},"readFileCalls":{},"lastReadBytes":{},"fileSizeCalls":{},"copyItemCalls":{},"lastCopyBytes":{},"lastMoveFallbackBytes":{},"propertyBatchCalls":{},"propertyPerItemCalls":{}}}}})json",
            JsonEscapeUtf8(nameUtf8),
            JsonEscapeUtf8(puidUtf8),
            IsDirectory(*node) ? "false" : "true",
            node->sizeBytes,
            _activeBackendCalls.load(std::memory_order_acquire),
            _maxConcurrentBackendCalls.load(std::memory_order_acquire),
            backendThreadIdsObserved,
            backendThreadIdsOverflow ? "true" : "false",
            _operationDelayMs,
            _cancelRequests.load(std::memory_order_acquire),
            _writeFileCalls.load(std::memory_order_acquire),
            _readerStats->readFileCalls.load(std::memory_order_acquire),
            _readerStats->lastReadBytes.load(std::memory_order_acquire),
            _fileSizeCalls.load(std::memory_order_acquire),
            _copyItemCalls.load(std::memory_order_acquire),
            _lastCopyBytes.load(std::memory_order_acquire),
            _lastMoveFallbackBytes.load(std::memory_order_acquire),
            _propertyBatchCalls.load(std::memory_order_acquire),
            _propertyPerItemCalls.load(std::memory_order_acquire));
        return S_OK;
    }

private:
    friend class FakeMtpActiveCallScope;

    void RecordBackendThread() noexcept
    {
        const DWORD threadId = GetCurrentThreadId();
        std::lock_guard lock(_threadStatsMutex);
        for (uint32_t index = 0; index < _backendThreadIdsObserved; ++index)
        {
            if (_backendThreadIds[index] == threadId)
            {
                return;
            }
        }

        if (_backendThreadIdsObserved < _backendThreadIds.size())
        {
            _backendThreadIds[_backendThreadIdsObserved] = threadId;
            ++_backendThreadIdsObserved;
        }
        else
        {
            _backendThreadIdsOverflow = true;
        }
    }

    [[nodiscard]] std::pair<uint32_t, bool> BackendThreadStats() const noexcept
    {
        std::lock_guard lock(_threadStatsMutex);
        return {_backendThreadIdsObserved, _backendThreadIdsOverflow};
    }

    void DelayForOperation() noexcept
    {
        if (_operationDelayMs == 0u)
        {
            return;
        }

        if (! _cancelUnblocksDelay)
        {
            std::this_thread::sleep_for(std::chrono::milliseconds(_operationDelayMs));
            return;
        }

        std::unique_lock lock(_delayMutex);
        static_cast<void>(_delayCv.wait_for(lock, std::chrono::milliseconds(_operationDelayMs), [&]() noexcept { return _cancelRequested; }));
    }

    void AddDirectory(std::wstring path, std::wstring persistentId, __int64 now)
    {
        std::lock_guard lock(_mutex);
        AddDirectoryLocked(std::move(path), std::move(persistentId), now);
    }

    void AddDirectoryLocked(std::wstring path, std::wstring persistentId, __int64 now)
    {
        FakeNode node;
        node.attributes     = FILE_ATTRIBUTE_DIRECTORY;
        node.creationTime   = now;
        node.lastAccessTime = now;
        node.lastWriteTime  = now;
        node.changeTime     = now;
        node.persistentId   = std::move(persistentId);
        node.objectId       = L"oid-" + std::to_wstring(_nextObjectId);
        node.displayName    = LeafName(path);
        ++_nextObjectId;
        _nodes[NormalizeMtpPath(path)] = std::move(node);
    }

    void AddFile(std::wstring path, std::wstring persistentId, std::vector<std::byte> content, __int64 now, std::wstring displayName = {})
    {
        std::lock_guard lock(_mutex);
        FakeNode node;
        node.attributes     = FILE_ATTRIBUTE_NORMAL;
        node.sizeBytes      = static_cast<uint64_t>(content.size());
        node.creationTime   = now;
        node.lastAccessTime = now;
        node.lastWriteTime  = now;
        node.changeTime     = now;
        node.persistentId   = std::move(persistentId);
        node.objectId       = L"oid-" + std::to_wstring(_nextObjectId);
        node.displayName    = displayName.empty() ? LeafName(path) : std::move(displayName);
        node.content        = std::move(content);
        ++_nextObjectId;
        _nodes[NormalizeMtpPath(path)] = std::move(node);
    }

    [[nodiscard]] FakeNode* FindNodeLocked(std::wstring_view path) noexcept
    {
        const std::wstring normalized = NormalizeMtpPath(path);
        const auto it                 = _nodes.find(normalized);
        return it == _nodes.end() ? nullptr : &it->second;
    }

    [[nodiscard]] const FakeNode* FindNodeLocked(std::wstring_view path) const noexcept
    {
        const std::wstring normalized = NormalizeMtpPath(path);
        const auto it                 = _nodes.find(normalized);
        return it == _nodes.end() ? nullptr : &it->second;
    }

    [[nodiscard]] bool HasChildrenLocked(std::wstring_view path) const
    {
        const std::wstring normalized = NormalizeMtpPath(path);
        return std::any_of(_nodes.begin(), _nodes.end(), [&](const auto& entry) { return entry.first != normalized && ParentPath(entry.first) == normalized; });
    }

    [[nodiscard]] bool HasAmbiguousDisplayNameLocked(std::wstring_view path) const
    {
        const std::wstring normalized = NormalizeMtpPath(path);
        const std::wstring parent     = ParentPath(normalized);
        const std::wstring leaf       = LeafName(normalized);
        uint32_t matches              = 0;
        for (const auto& [childPath, child] : _nodes)
        {
            if (ParentPath(childPath) != parent)
            {
                continue;
            }

            const std::wstring displayName = child.displayName.empty() ? LeafName(childPath) : child.displayName;
            if (EqualsPathComponent(displayName, leaf))
            {
                ++matches;
            }
        }

        return matches > 1u;
    }

    void EraseTreeLocked(std::wstring_view path)
    {
        const std::wstring normalized = NormalizeMtpPath(path);
        std::erase_if(_nodes, [&](const auto& entry) {
            return entry.first == normalized || (entry.first.size() > normalized.size() && entry.first.rfind(normalized + L"/", 0) == 0);
        });
    }

    HRESULT CopyItemLocked(std::wstring_view sourcePath, std::wstring_view destinationPath, bool allowOverwrite, uint64_t* copiedBytes)
    {
        if (copiedBytes)
        {
            *copiedBytes = 0;
        }

        const std::wstring source = NormalizeMtpPath(sourcePath);
        const std::wstring dest   = NormalizeMtpPath(destinationPath);
        const auto sourceIt       = _nodes.find(source);
        if (sourceIt == _nodes.end())
        {
            return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
        }
        if (! _copyItemFailOnceDestinationPathContains.empty() && ! _copyItemFailOnceConsumed &&
            dest.find(_copyItemFailOnceDestinationPathContains) != std::wstring::npos)
        {
            _copyItemFailOnceConsumed = true;
            return HRESULT_FROM_WIN32(ERROR_WRITE_FAULT);
        }

        const std::wstring parent  = ParentPath(dest);
        const FakeNode* parentNode = FindNodeLocked(parent);
        if (! parentNode || ! IsDirectory(*parentNode))
        {
            return HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND);
        }

        if (_nodes.contains(dest) && ! allowOverwrite)
        {
            return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
        }

        std::vector<std::pair<std::wstring, FakeNode>> copies;
        copies.reserve(_nodes.size());
        uint64_t totalBytes = 0;
        for (const auto& [path, node] : _nodes)
        {
            if (path == source || (path.size() > source.size() && path.rfind(source + L"/", 0) == 0))
            {
                if (! IsDirectory(node))
                {
                    totalBytes += node.sizeBytes;
                }

                std::wstring rel = path.substr(source.size());
                copies.push_back({dest + rel, node});
            }
        }

        for (auto& [targetPath, node] : copies)
        {
            node.persistentId = L"puid-" + std::to_wstring(_nextObjectId);
            node.objectId     = L"oid-" + std::to_wstring(_nextObjectId);
            node.displayName  = LeafName(targetPath);
            ++_nextObjectId;
            _nodes[targetPath] = std::move(node);
        }

        if (copiedBytes)
        {
            *copiedBytes = totalBytes;
        }
        return S_OK;
    }

    HRESULT RenameItemLocked(std::wstring_view sourcePath, std::wstring_view destinationPath, bool allowOverwrite)
    {
        const std::wstring source = NormalizeMtpPath(sourcePath);
        const std::wstring dest   = NormalizeMtpPath(destinationPath);
        if (source == L"/" || dest == L"/")
        {
            return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        }
        if (! _nodes.contains(source))
        {
            return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
        }
        if (_nodes.contains(dest) && ! allowOverwrite)
        {
            return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
        }
        if (! _renameItemFailOnceDestinationPath.empty() && ! _renameItemFailOnceConsumed && dest == _renameItemFailOnceDestinationPath)
        {
            _renameItemFailOnceConsumed = true;
            return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        }
        if (! _renameItemFailDestinationPath.empty() && dest == _renameItemFailDestinationPath)
        {
            return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        }

        const std::wstring parent  = ParentPath(dest);
        const FakeNode* parentNode = FindNodeLocked(parent);
        if (! parentNode || ! IsDirectory(*parentNode))
        {
            return HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND);
        }

        if (_nodes.contains(dest))
        {
            EraseTreeLocked(dest);
        }

        std::vector<std::pair<std::wstring, FakeNode>> moved;
        moved.reserve(_nodes.size());
        for (const auto& [path, node] : _nodes)
        {
            if (path == source || (path.size() > source.size() && path.rfind(source + L"/", 0) == 0))
            {
                std::wstring rel = path.substr(source.size());
                moved.push_back({dest + rel, node});
            }
        }

        EraseTreeLocked(source);
        for (auto& [targetPath, node] : moved)
        {
            node.displayName   = LeafName(targetPath);
            _nodes[targetPath] = std::move(node);
        }
        return S_OK;
    }

    mutable std::mutex _mutex;
    std::unordered_map<std::wstring, FakeNode> _nodes;
    uint64_t _nextObjectId                = 1;
    uint32_t _operationDelayMs            = 0;
    uint32_t _readFileDelayMs             = 0;
    bool _cancelUnblocksDelay             = false;
    bool _moveFallbackDeleteSourceFails   = false;
    bool _omitPersistentIdForCreatedFiles = false;
    std::wstring _writeFileFailOncePathContains;
    bool _writeFileFailOnceConsumed = false;
    std::wstring _copyItemFailOnceDestinationPathContains;
    bool _copyItemFailOnceConsumed = false;
    std::wstring _deleteItemFailOncePath;
    bool _deleteItemFailOnceConsumed = false;
    std::wstring _renameItemFailOnceDestinationPath;
    bool _renameItemFailOnceConsumed = false;
    std::wstring _renameItemFailDestinationPath;
    std::wstring _disconnectEnumerateOncePath;
    bool _disconnectEnumerateOnceConsumed = false;
    std::mutex _delayMutex;
    std::condition_variable _delayCv;
    bool _cancelRequested = false;
    std::atomic_uint32_t _activeBackendCalls{0};
    std::atomic_uint32_t _maxConcurrentBackendCalls{0};
    mutable std::mutex _threadStatsMutex;
    std::array<DWORD, 64> _backendThreadIds{};
    uint32_t _backendThreadIdsObserved = 0;
    bool _backendThreadIdsOverflow     = false;
    std::atomic_uint32_t _cancelRequests{0};
    std::atomic_uint32_t _writeFileCalls{0};
    std::shared_ptr<FakeMtpReaderStats> _readerStats = std::make_shared<FakeMtpReaderStats>();
    std::atomic_uint32_t _fileSizeCalls{0};
    std::atomic_uint32_t _copyItemCalls{0};
    std::atomic_uint64_t _lastCopyBytes{0};
    std::atomic_uint64_t _lastMoveFallbackBytes{0};
    std::atomic_uint32_t _propertyBatchCalls{0};
    std::atomic_uint32_t _propertyPerItemCalls{0};
};

FakeMtpActiveCallScope::FakeMtpActiveCallScope(FakeMtpBackend& owner) noexcept : _owner(owner)
{
    _owner.RecordBackendThread();

    const uint32_t active = _owner._activeBackendCalls.fetch_add(1u, std::memory_order_acq_rel) + 1u;
    uint32_t observed     = _owner._maxConcurrentBackendCalls.load(std::memory_order_acquire);
    while (active > observed &&
           ! _owner._maxConcurrentBackendCalls.compare_exchange_weak(observed, active, std::memory_order_acq_rel, std::memory_order_acquire))
    {
    }

    if (_owner._operationDelayMs != 0u)
    {
        _owner.DelayForOperation();
    }
}

FakeMtpActiveCallScope::~FakeMtpActiveCallScope() noexcept
{
    static_cast<void>(_owner._activeBackendCalls.fetch_sub(1u, std::memory_order_acq_rel));
}
} // namespace

[[nodiscard]] std::unique_ptr<IMtpBackend> CreateFakeMtpBackend(std::string_view optionsJsonUtf8) noexcept
{
    return std::make_unique<FakeMtpBackend>(optionsJsonUtf8);
}
} // namespace FileSystemMtpInternal
