#include "Framework.h"

#include "FileOperationMoveBreadcrumb.h"

#include "AppDataPaths.h"
#include "FileOperationDurableStore.h"
#include "Helpers.h"
#include "StringConversion.h"
#include "YyjsonHelpers.h"

#include <algorithm>
#include <array>
#include <chrono>
#include <limits>
#include <mutex>
#include <optional>
#include <string_view>
#include <system_error>

namespace FileOperationMoveBreadcrumb
{
namespace
{
constexpr std::string_view kRecordKind = "move-task-breadcrumb";
constexpr size_t kMaximumTextCharacters = 32u * 1024u;
constexpr std::wstring_view kDefaultInstanceId = L"host/default";
constexpr std::wstring_view kContextInstanceIdPrefix = L"host/context/";

#ifdef ENABLE_TESTS
std::mutex g_rootMutex;
std::filesystem::path g_rootOverride;
#endif

[[nodiscard]] std::filesystem::path ResolveRoot() noexcept
{
#ifdef ENABLE_TESTS
    {
        std::scoped_lock lock(g_rootMutex);
        if (! g_rootOverride.empty())
        {
            return g_rootOverride;
        }
    }
#endif
    const std::filesystem::path localAppData = AppDataPaths::GetLocalAppDataPath();
    return localAppData.empty()
        ? std::filesystem::path{}
        : localAppData / L"RedSalamander" / L"State" / L"FileOperations" / L"MoveBreadcrumbs";
}

[[nodiscard]] uint64_t CurrentFileTime() noexcept
{
    FILETIME now{};
    GetSystemTimeAsFileTime(&now);
    ULARGE_INTEGER value{};
    value.LowPart = now.dwLowDateTime;
    value.HighPart = now.dwHighDateTime;
    return value.QuadPart;
}

[[nodiscard]] HRESULT MakeRecordPath(const std::filesystem::path& root,
                                     const uint64_t taskId,
                                     std::filesystem::path& out) noexcept
{
    std::array<std::byte, 16u> random{};
    const HRESULT randomHr = Common::Crypto::GenerateRandomBytes(random);
    if (FAILED(randomHr))
    {
        return randomHr;
    }
    std::wstring token;
    Common::Crypto::AppendHexToken(token, random);
    out = root / std::format(L"{0:016X}-{1}.move.json", taskId, token);
    return S_OK;
}

[[nodiscard]] bool IsDirectBreadcrumbChild(const std::filesystem::path& path) noexcept
{
    if (path.empty() || ! path.filename().native().ends_with(L".move.json"))
    {
        return false;
    }
    const std::filesystem::path root = ResolveRoot();
    if (root.empty())
    {
        return false;
    }
    std::error_code error;
    const std::filesystem::path absoluteRoot = std::filesystem::absolute(root, error).lexically_normal();
    if (error)
    {
        return false;
    }
    const std::filesystem::path absolutePath = std::filesystem::absolute(path, error).lexically_normal();
    return ! error &&
           OrdinalString::EqualsNoCase(absolutePath.parent_path().native(), absoluteRoot.native());
}

[[nodiscard]] std::string_view PhaseName(const DurablePhase phase) noexcept
{
    switch (phase)
    {
        case DurablePhase::Admitted: return "admitted";
        case DurablePhase::Executing: return "executing";
        case DurablePhase::Terminal: return "terminal";
    }
    return "invalid";
}

[[nodiscard]] std::optional<DurablePhase> ParsePhase(const std::string_view value) noexcept
{
    if (value == "admitted") return DurablePhase::Admitted;
    if (value == "executing") return DurablePhase::Executing;
    if (value == "terminal") return DurablePhase::Terminal;
    return std::nullopt;
}

[[nodiscard]] std::string_view PaneName(const PaneHint pane) noexcept
{
    switch (pane)
    {
        case PaneHint::Left: return "left";
        case PaneHint::Right: return "right";
    }
    return "invalid";
}

[[nodiscard]] std::optional<PaneHint> ParsePane(const std::string_view value) noexcept
{
    if (value == "left") return PaneHint::Left;
    if (value == "right") return PaneHint::Right;
    return std::nullopt;
}

[[nodiscard]] HRESULT AddUtf8String(yyjson_mut_doc* document,
                                    yyjson_mut_val* object,
                                    const char* key,
                                    const std::string_view value) noexcept
{
    return yyjson_mut_obj_add_strncpy(document, object, key, value.data(), value.size()) ? S_OK : E_OUTOFMEMORY;
}

[[nodiscard]] HRESULT AddWideString(yyjson_mut_doc* document,
                                    yyjson_mut_val* object,
                                    const char* key,
                                    const std::wstring_view value) noexcept
{
    if (value.size() > kMaximumTextCharacters)
    {
        return HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE);
    }
    const std::optional<std::string> utf8 = Common::Strings::TryUtf8FromUtf16Strict(value);
    return utf8.has_value() ? AddUtf8String(document, object, key, utf8.value())
                            : HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);
}

[[nodiscard]] HRESULT AddLocation(yyjson_mut_doc* document,
                                  yyjson_mut_val* parent,
                                  const char* key,
                                  const QualifiedLocation& location) noexcept
{
    yyjson_mut_val* object = yyjson_mut_obj(document);
    if (! object || ! yyjson_mut_obj_add_val(document, parent, key, object))
    {
        return E_OUTOFMEMORY;
    }
    HRESULT hr = AddWideString(document, object, "pluginId", location.pluginId);
    if (SUCCEEDED(hr)) hr = AddWideString(document, object, "pluginShortId", location.pluginShortId);
    if (SUCCEEDED(hr)) hr = AddWideString(document, object, "instanceId", location.instanceId);
    if (SUCCEEDED(hr)) hr = AddWideString(document, object, "profileId", location.profileId);
    if (SUCCEEDED(hr)) hr = AddWideString(document, object, "rootId", location.rootId);
    if (SUCCEEDED(hr)) hr = AddWideString(document, object, "representativePath", location.representativePath.native());
    return hr;
}

[[nodiscard]] bool IsValidQualifiedLocation(const QualifiedLocation& location) noexcept
{
    const bool validInstanceId = location.instanceId == kDefaultInstanceId ||
        (location.instanceId.starts_with(kContextInstanceIdPrefix) &&
         location.instanceId.size() > kContextInstanceIdPrefix.size());
    return ! location.pluginId.empty() && ! location.pluginShortId.empty() &&
           validInstanceId && ! location.profileId.empty() && ! location.rootId.empty() &&
           ! location.representativePath.empty();
}

[[nodiscard]] HRESULT Serialize(const Record& record, std::string& json) noexcept
{
    const bool haveAdmittedItem = record.admittedStrategies.nativeItems != 0u ||
                                  record.admittedStrategies.managedItems != 0u ||
                                  record.admittedStrategies.copyOnlyItems != 0u;
    const bool boundedRootsAreConsistent =
        (! record.sourceRootSamplesTruncated && record.sourceRootCount == record.sourceRootSamples.size()) ||
        (record.sourceRootSamplesTruncated &&
         record.sourceRootSamples.size() == kMaximumPersistedSourceRootSamples &&
         record.sourceRootCount > record.sourceRootSamples.size());
    if (record.schemaVersion != kSchemaVersion || record.taskId == 0u || record.createdFileTime == 0u ||
        ! haveAdmittedItem || record.sourceRootCount == 0u || record.sourceRootSamples.empty() ||
        record.sourceRootSamples.size() > kMaximumPersistedSourceRootSamples ||
        ! boundedRootsAreConsistent ||
        ! IsValidQualifiedLocation(record.destination) ||
        ! std::ranges::all_of(record.sourceRootSamples, IsValidQualifiedLocation))
    {
        return E_INVALIDARG;
    }

    Common::Json::UniqueMutableDocument document(yyjson_mut_doc_new(nullptr));
    if (! document)
    {
        return E_OUTOFMEMORY;
    }
    yyjson_mut_val* root = yyjson_mut_obj(document.get());
    if (! root)
    {
        return E_OUTOFMEMORY;
    }
    yyjson_mut_doc_set_root(document.get(), root);

    HRESULT hr = AddUtf8String(document.get(), root, "kind", kRecordKind);
    if (SUCCEEDED(hr) && ! yyjson_mut_obj_add_uint(document.get(), root, "schemaVersion", record.schemaVersion)) hr = E_OUTOFMEMORY;
    if (SUCCEEDED(hr) && ! yyjson_mut_obj_add_uint(document.get(), root, "taskId", record.taskId)) hr = E_OUTOFMEMORY;
    if (SUCCEEDED(hr) && ! yyjson_mut_obj_add_uint(document.get(), root, "createdFileTime", record.createdFileTime)) hr = E_OUTOFMEMORY;
    if (SUCCEEDED(hr)) hr = AddUtf8String(document.get(), root, "phase", PhaseName(record.phase));
    if (SUCCEEDED(hr)) hr = AddUtf8String(document.get(), root, "sourcePane", PaneName(record.sourcePane));
    if (SUCCEEDED(hr)) hr = AddUtf8String(document.get(), root, "destinationPane", PaneName(record.destinationPane));
    if (SUCCEEDED(hr) && ! yyjson_mut_obj_add_uint(document.get(), root, "nativeItems", record.admittedStrategies.nativeItems)) hr = E_OUTOFMEMORY;
    if (SUCCEEDED(hr) && ! yyjson_mut_obj_add_uint(document.get(), root, "managedItems", record.admittedStrategies.managedItems)) hr = E_OUTOFMEMORY;
    if (SUCCEEDED(hr) && ! yyjson_mut_obj_add_uint(document.get(), root, "copyOnlyItems", record.admittedStrategies.copyOnlyItems)) hr = E_OUTOFMEMORY;
    if (SUCCEEDED(hr) && ! yyjson_mut_obj_add_uint(document.get(), root, "sourceRootCount", record.sourceRootCount)) hr = E_OUTOFMEMORY;
    if (SUCCEEDED(hr) && ! yyjson_mut_obj_add_bool(document.get(), root, "sourceRootSamplesTruncated", record.sourceRootSamplesTruncated)) hr = E_OUTOFMEMORY;
    if (FAILED(hr))
    {
        return hr;
    }

    yyjson_mut_val* roots = yyjson_mut_arr(document.get());
    if (! roots || ! yyjson_mut_obj_add_val(document.get(), root, "sourceRootSamples", roots))
    {
        return E_OUTOFMEMORY;
    }
    for (const QualifiedLocation& location : record.sourceRootSamples)
    {
        yyjson_mut_val* item = yyjson_mut_obj(document.get());
        if (! item || ! yyjson_mut_arr_append(roots, item))
        {
            return E_OUTOFMEMORY;
        }
        hr = AddWideString(document.get(), item, "pluginId", location.pluginId);
        if (SUCCEEDED(hr)) hr = AddWideString(document.get(), item, "pluginShortId", location.pluginShortId);
        if (SUCCEEDED(hr)) hr = AddWideString(document.get(), item, "instanceId", location.instanceId);
        if (SUCCEEDED(hr)) hr = AddWideString(document.get(), item, "profileId", location.profileId);
        if (SUCCEEDED(hr)) hr = AddWideString(document.get(), item, "rootId", location.rootId);
        if (SUCCEEDED(hr)) hr = AddWideString(document.get(), item, "representativePath", location.representativePath.native());
        if (FAILED(hr)) return hr;
    }
    hr = AddLocation(document.get(), root, "destination", record.destination);
    if (FAILED(hr)) return hr;

    size_t length = 0u;
    yyjson_write_err writeError{};
    Common::Json::UniqueMallocString output(
        yyjson_mut_write_opts(document.get(), YYJSON_WRITE_NOFLAG, nullptr, &length, &writeError));
    if (! output)
    {
        return E_OUTOFMEMORY;
    }
    if (length == 0u || length > kMaximumRecordBytes)
    {
        return HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE);
    }
    json.assign(output.get(), length);
    return S_OK;
}

[[nodiscard]] std::optional<std::string_view> ReadString(yyjson_val* object, const char* key) noexcept
{
    yyjson_val* value = yyjson_obj_get(object, key);
    return value && yyjson_is_str(value)
        ? std::optional<std::string_view>(std::string_view(yyjson_get_str(value), yyjson_get_len(value)))
        : std::nullopt;
}

[[nodiscard]] std::optional<std::wstring> ReadWideString(yyjson_val* object, const char* key) noexcept
{
    const std::optional<std::string_view> value = ReadString(object, key);
    if (! value.has_value() || value->size() > kMaximumTextCharacters * 4u)
    {
        return std::nullopt;
    }
    return Common::Strings::TryUtf16FromUtf8Strict(value.value());
}

[[nodiscard]] bool ReadUint64(yyjson_val* object, const char* key, uint64_t& value) noexcept
{
    yyjson_val* member = yyjson_obj_get(object, key);
    if (! member || ! yyjson_is_uint(member)) return false;
    value = yyjson_get_uint(member);
    return true;
}

[[nodiscard]] bool ReadBool(yyjson_val* object, const char* key, bool& value) noexcept
{
    yyjson_val* member = yyjson_obj_get(object, key);
    if (! member || ! yyjson_is_bool(member)) return false;
    value = yyjson_get_bool(member);
    return true;
}

[[nodiscard]] HRESULT ReadLocation(yyjson_val* object, QualifiedLocation& out) noexcept
{
    if (! object || ! yyjson_is_obj(object)) return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    const auto pluginId = ReadWideString(object, "pluginId");
    const auto pluginShortId = ReadWideString(object, "pluginShortId");
    const auto instanceId = ReadWideString(object, "instanceId");
    const auto profileId = ReadWideString(object, "profileId");
    const auto rootId = ReadWideString(object, "rootId");
    const auto path = ReadWideString(object, "representativePath");
    if (! pluginId.has_value() || ! pluginShortId.has_value() || ! instanceId.has_value() ||
        ! profileId.has_value() || ! rootId.has_value() || ! path.has_value() ||
        pluginId->empty() || pluginShortId->empty() || profileId->empty() ||
        rootId->empty() || path->empty())
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }
    out.pluginId = pluginId.value();
    out.pluginShortId = pluginShortId.value();
    out.instanceId = instanceId.value();
    out.profileId = profileId.value();
    out.rootId = rootId.value();
    out.representativePath = std::filesystem::path(path.value());
    return IsValidQualifiedLocation(out) ? S_OK : HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
}

[[nodiscard]] HRESULT ReadFileBounded(const std::filesystem::path& path, std::string& bytes) noexcept
{
    return FileOperationDurableStore::ReadBoundedRegularFile(path, kMaximumRecordBytes, bytes);
}

[[nodiscard]] HRESULT Parse(std::string json, Record& out) noexcept
{
    Common::Json::UniqueDocument document(
        yyjson_read_opts(json.data(), json.size(), YYJSON_READ_ALLOW_BOM, nullptr, nullptr));
    if (! document) return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    yyjson_val* root = yyjson_doc_get_root(document.get());
    if (! root || ! yyjson_is_obj(root)) return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    const auto kind = ReadString(root, "kind");
    const auto phaseText = ReadString(root, "phase");
    const auto phase = phaseText.has_value() ? ParsePhase(phaseText.value()) : std::nullopt;
    const auto sourcePaneText = ReadString(root, "sourcePane");
    const auto sourcePane = sourcePaneText.has_value() ? ParsePane(sourcePaneText.value()) : std::nullopt;
    const auto destinationPaneText = ReadString(root, "destinationPane");
    const auto destinationPane = destinationPaneText.has_value() ? ParsePane(destinationPaneText.value()) : std::nullopt;
    uint64_t schema = 0u;
    Record parsed{};
    if (! kind.has_value() || kind.value() != kRecordKind || ! phase.has_value() ||
        ! sourcePane.has_value() || ! destinationPane.has_value() ||
        ! ReadUint64(root, "schemaVersion", schema) || schema != kSchemaVersion ||
        ! ReadUint64(root, "taskId", parsed.taskId) || parsed.taskId == 0u ||
        ! ReadUint64(root, "createdFileTime", parsed.createdFileTime) || parsed.createdFileTime == 0u ||
        ! ReadUint64(root, "nativeItems", parsed.admittedStrategies.nativeItems) ||
        ! ReadUint64(root, "managedItems", parsed.admittedStrategies.managedItems) ||
        ! ReadUint64(root, "copyOnlyItems", parsed.admittedStrategies.copyOnlyItems) ||
        ! ReadUint64(root, "sourceRootCount", parsed.sourceRootCount) || parsed.sourceRootCount == 0u ||
        ! ReadBool(root, "sourceRootSamplesTruncated", parsed.sourceRootSamplesTruncated))
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }
    parsed.schemaVersion = static_cast<uint32_t>(schema);
    parsed.phase = phase.value();
    parsed.sourcePane = sourcePane.value();
    parsed.destinationPane = destinationPane.value();
    yyjson_val* roots = yyjson_obj_get(root, "sourceRootSamples");
    if (! roots || ! yyjson_is_arr(roots) || yyjson_arr_size(roots) == 0u ||
        yyjson_arr_size(roots) > kMaximumPersistedSourceRootSamples)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }
    size_t index = 0u;
    size_t maximum = 0u;
    yyjson_val* value = nullptr;
    yyjson_arr_foreach(roots, index, maximum, value)
    {
        QualifiedLocation location;
        const HRESULT locationHr = ReadLocation(value, location);
        if (FAILED(locationHr)) return locationHr;
        parsed.sourceRootSamples.push_back(std::move(location));
    }
    const HRESULT destinationHr = ReadLocation(yyjson_obj_get(root, "destination"), parsed.destination);
    if (FAILED(destinationHr)) return destinationHr;
    const bool haveAdmittedItem = parsed.admittedStrategies.nativeItems != 0u ||
                                  parsed.admittedStrategies.managedItems != 0u ||
                                  parsed.admittedStrategies.copyOnlyItems != 0u;
    const bool boundedRootsAreConsistent =
        (! parsed.sourceRootSamplesTruncated && parsed.sourceRootCount == parsed.sourceRootSamples.size()) ||
        (parsed.sourceRootSamplesTruncated &&
         parsed.sourceRootSamples.size() == kMaximumPersistedSourceRootSamples &&
         parsed.sourceRootCount > parsed.sourceRootSamples.size());
    if (! haveAdmittedItem || ! boundedRootsAreConsistent)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }
    out = std::move(parsed);
    return S_OK;
}
} // namespace

std::optional<std::wstring> TryDecodeNavigationInstanceContext(
    const std::wstring_view canonicalInstanceId)
{
    if (canonicalInstanceId == kDefaultInstanceId)
    {
        return std::wstring{};
    }
    if (canonicalInstanceId.starts_with(kContextInstanceIdPrefix) &&
        canonicalInstanceId.size() > kContextInstanceIdPrefix.size())
    {
        return std::wstring(canonicalInstanceId.substr(kContextInstanceIdPrefix.size()));
    }
    return std::nullopt;
}

HRESULT Breadcrumb::Create(Record record, Breadcrumb& out) noexcept
{
    const std::filesystem::path root = ResolveRoot();
    if (root.empty()) return HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND);
    record.schemaVersion = kSchemaVersion;
    record.phase = DurablePhase::Admitted;
    if (record.createdFileTime == 0u) record.createdFileTime = CurrentFileTime();
    std::filesystem::path path;
    const HRESULT pathHr = MakeRecordPath(root, record.taskId, path);
    if (FAILED(pathHr)) return pathHr;
    Breadcrumb created;
    created._record = std::move(record);
    created._path = std::move(path);
    const HRESULT hr = created.Persist(L"create", true);
    if (FAILED(hr)) return hr;
    out = std::move(created);
    return S_OK;
}

HRESULT Breadcrumb::Load(const std::filesystem::path& path, Record& out) noexcept
{
    std::string bytes;
    const HRESULT readHr = ReadFileBounded(path, bytes);
    return FAILED(readHr) ? readHr : Parse(std::move(bytes), out);
}

HRESULT Breadcrumb::LoadInterrupted(std::vector<LoadedRecord>& out, LoadStats* const stats) noexcept
{
    const auto startedAt = std::chrono::steady_clock::now();
    out.clear();
    LoadStats collected{};
    const std::filesystem::path root = ResolveRoot();
    if (root.empty()) return HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND);
    std::vector<std::filesystem::path> files;
    std::vector<FileOperationDurableStore::RejectedDirectChild> rejectedChildren;
    const HRESULT scanHr = FileOperationDurableStore::EnumerateDirectRegularFiles(
        root, kMaximumStoreEntries, files, &rejectedChildren);
    if (scanHr == S_FALSE)
    {
        if (stats) *stats = collected;
        return S_FALSE;
    }
    if (FAILED(scanHr))
    {
        if (scanHr == HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE))
        {
            Debug::Warning(L"File Operations breadcrumb store '{}' holds more than {} direct children; interrupted-Move notices are withheld until it is cleaned.",
                           root.native(),
                           kMaximumStoreEntries);
        }
        if (stats) *stats = collected;
        return scanHr;
    }
    for (const FileOperationDurableStore::RejectedDirectChild& child : rejectedChildren)
    {
        if (child.path.extension() == L".json" && child.path.filename().native().ends_with(L".move.json"))
        {
            ++collected.recordsScanned;
            ++collected.rejectedRecords;
        }
    }
    for (const std::filesystem::path& path : files)
    {
        if (path.extension() != L".json" || ! path.filename().native().ends_with(L".move.json"))
        {
            continue;
        }
        ++collected.recordsScanned;
        Record record;
        const HRESULT loadHr = Load(path, record);
        if (FAILED(loadHr))
        {
            ++collected.rejectedRecords;
            continue;
        }
        if (record.phase == DurablePhase::Terminal)
        {
            if (SUCCEEDED(Acknowledge(path))) ++collected.terminalRecordsRemoved;
            continue;
        }
        out.push_back(LoadedRecord{path, std::move(record)});
        ++collected.interruptedRecords;
    }
    if (stats) *stats = collected;
    const auto elapsed = std::chrono::duration_cast<std::chrono::microseconds>(
        std::chrono::steady_clock::now() - startedAt).count();
    Debug::Perf::Emit(L"fileops.move_breadcrumb.load.us",
                      collected.rejectedRecords == 0u ? L"loaded" : L"partial",
                      static_cast<uint64_t>((std::max)(elapsed, int64_t{0})),
                      collected.interruptedRecords,
                      collected.recordsScanned,
                      S_OK);
    return collected.rejectedRecords == 0u ? S_OK : S_FALSE;
}

HRESULT Breadcrumb::Persist(const std::wstring_view detail, const bool failIfExists) noexcept
{
    const auto startedAt = std::chrono::steady_clock::now();
    std::string json;
    HRESULT hr = Serialize(_record, json);
    if (SUCCEEDED(hr))
    {
        hr = FileOperationDurableStore::PersistDirectChild(
            ResolveRoot(),
            _path,
            json,
            kMaximumRecordBytes,
            failIfExists ? Common::Files::ExistingTargetPolicy::FailIfExists
                         : Common::Files::ExistingTargetPolicy::Replace);
    }
    const auto elapsed = std::chrono::duration_cast<std::chrono::microseconds>(
        std::chrono::steady_clock::now() - startedAt).count();
    Debug::Perf::Emit(L"fileops.move_breadcrumb.persist.us",
                      detail,
                      static_cast<uint64_t>((std::max)(elapsed, int64_t{0})),
                      static_cast<uint64_t>(_record.phase),
                      static_cast<uint64_t>(json.size()),
                      hr);
    return hr;
}

HRESULT Breadcrumb::Advance(const DurablePhase phase) noexcept
{
    if (phase == DurablePhase::Terminal) return E_INVALIDARG;
    if (static_cast<uint8_t>(phase) <= static_cast<uint8_t>(_record.phase)) return S_FALSE;
    const DurablePhase previous = _record.phase;
    _record.phase = phase;
    const HRESULT hr = Persist(L"advance");
    if (FAILED(hr)) _record.phase = previous;
    return hr;
}

HRESULT Breadcrumb::Finalize() noexcept
{
    const DurablePhase previous = _record.phase;
    _record.phase = DurablePhase::Terminal;
    const HRESULT persistHr = Persist(L"terminal");
    if (FAILED(persistHr))
    {
        _record.phase = previous;
        return persistHr;
    }
    return Acknowledge(_path);
}

HRESULT Acknowledge(const std::filesystem::path& path) noexcept
{
    if (! IsDirectBreadcrumbChild(path)) return E_INVALIDARG;
    return FileOperationDurableStore::RemoveDirectRegularFile(ResolveRoot(), path);
}

#ifdef ENABLE_TESTS
void SetRootForSelfTest(const std::filesystem::path& root) noexcept
{
    std::scoped_lock lock(g_rootMutex);
    g_rootOverride = root;
}

void ClearRootForSelfTest() noexcept
{
    std::scoped_lock lock(g_rootMutex);
    g_rootOverride.clear();
}
#endif
} // namespace FileOperationMoveBreadcrumb
