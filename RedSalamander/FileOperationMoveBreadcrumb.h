#pragma once

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <cstdint>
#include <filesystem>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

namespace FileOperationMoveBreadcrumb
{
inline constexpr uint32_t kSchemaVersion                   = 2u;
inline constexpr size_t kMaximumPersistedSourceRootSamples = 16u;
inline constexpr uint64_t kMaximumRecordBytes              = 128u * 1024u;
inline constexpr size_t kMaximumStoreEntries               = 4096u;

enum class DurablePhase : uint8_t
{
    Admitted,
    Executing,
    Terminal,
};

enum class PaneHint : uint8_t
{
    Left,
    Right,
};

struct StrategyCounts final
{
    uint64_t nativeItems   = 0u;
    uint64_t managedItems  = 0u;
    uint64_t copyOnlyItems = 0u;
};

struct QualifiedLocation final
{
    std::wstring pluginId;
    std::wstring pluginShortId;
    // Comparison-only File Operations instance identity (host/default or
    // host/context/<opaque-context>), not a NavigationLocation instance context.
    std::wstring instanceId;
    std::wstring profileId;
    std::wstring rootId;
    std::filesystem::path representativePath;
};

[[nodiscard]] std::optional<std::wstring> TryDecodeNavigationInstanceContext(std::wstring_view canonicalInstanceId);

struct Record final
{
    uint32_t schemaVersion         = kSchemaVersion;
    uint64_t taskId                = 0u;
    uint64_t createdFileTime       = 0u;
    uint32_t writerProcessId       = 0u;
    uint64_t writerStartedFileTime = 0u;
    std::wstring singleItemName;
    DurablePhase phase = DurablePhase::Admitted;
    StrategyCounts admittedStrategies;
    PaneHint sourcePane             = PaneHint::Left;
    PaneHint destinationPane        = PaneHint::Right;
    uint64_t sourceRootCount        = 0u;
    bool sourceRootSamplesTruncated = false;
    std::vector<QualifiedLocation> sourceRootSamples;
    QualifiedLocation destination;
};

struct LoadedRecord final
{
    std::filesystem::path path;
    Record record;
};

struct LoadStats final
{
    uint64_t recordsScanned          = 0u;
    uint64_t interruptedRecords      = 0u;
    uint64_t terminalRecordsRemoved  = 0u;
    uint64_t rejectedRecords         = 0u;
    uint64_t activeWriterRecords     = 0u;
    uint64_t unverifiedWriterRecords = 0u;
    uint64_t legacyRecordsWithheld   = 0u;
    uint64_t legacyPeerObservations  = 0u;
};

// A Move breadcrumb is deliberately not a recovery journal. It contains no per-item mapping,
// identity receipt, or mutation authority and therefore exposes no Resume/Delete operation.
class Breadcrumb final
{
public:
    Breadcrumb() noexcept                        = default;
    Breadcrumb(const Breadcrumb&)                = delete;
    Breadcrumb& operator=(const Breadcrumb&)     = delete;
    Breadcrumb(Breadcrumb&&) noexcept            = default;
    Breadcrumb& operator=(Breadcrumb&&) noexcept = default;

    [[nodiscard]] static HRESULT Create(Record record, Breadcrumb& out) noexcept;
    [[nodiscard]] static HRESULT Load(const std::filesystem::path& path, Record& out) noexcept;
    [[nodiscard]] static HRESULT LoadInterrupted(std::vector<LoadedRecord>& out, LoadStats* stats = nullptr) noexcept;

    [[nodiscard]] HRESULT Advance(DurablePhase phase) noexcept;
    [[nodiscard]] HRESULT Finalize() noexcept;
#ifdef ENABLE_TESTS
    // Simulates a stale process identity only inside an explicitly selected selftest store.
    [[nodiscard]] HRESULT SimulateExitedWriterForSelfTest() noexcept;
#endif

    [[nodiscard]] const Record& GetRecord() const noexcept
    {
        return _record;
    }
    [[nodiscard]] const std::filesystem::path& GetPath() const noexcept
    {
        return _path;
    }

private:
    [[nodiscard]] HRESULT Persist(std::wstring_view detail, bool failIfExists = false) noexcept;

    Record _record;
    std::filesystem::path _path;
};

[[nodiscard]] HRESULT Acknowledge(const std::filesystem::path& path) noexcept;
[[nodiscard]] uint64_t AdmittedItemCount(const StrategyCounts& counts) noexcept;

#ifdef ENABLE_TESTS
enum class WriterObservationForSelfTest : uint8_t
{
    UseSystem,
    Active,
    Exited,
    Unverified,
};

enum class LegacyPeerObservationForSelfTest : uint8_t
{
    UseSystem,
    NoOtherProcess,
    OtherProcess,
    Uncertain,
};

void SetWriterObservationForSelfTest(WriterObservationForSelfTest observation) noexcept;
void SetLegacyPeerObservationForSelfTest(LegacyPeerObservationForSelfTest observation) noexcept;
void SetRootForSelfTest(const std::filesystem::path& root) noexcept;
void ClearRootForSelfTest() noexcept;
#endif
} // namespace FileOperationMoveBreadcrumb
