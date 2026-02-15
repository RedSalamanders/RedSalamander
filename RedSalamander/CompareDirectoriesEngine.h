#pragma once

#include <atomic>
#include <cstdint>
#include <cwctype>
#include <filesystem>
#include <functional>
#include <limits>
#include <memory>
#include <mutex>
#include <optional>
#include <string>
#include <string_view>
#include <unordered_map>

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027)
#include <wil/com.h>
#pragma warning(pop)

#include "PlugInterfaces/FileSystem.h"
#include "PlugInterfaces/Informations.h"
#include "SettingsStore.h"

enum class ComparePane : uint8_t
{
    Left,
    Right,
};

struct WStringViewNoCaseHash
{
    using is_transparent = void;

    size_t operator()(std::wstring_view value) const noexcept
    {
        // Case-insensitive FNV-1a for UTF-16 (ordinal).
        uint64_t hash = 14695981039346656037ull;
        for (const wchar_t ch : value)
        {
            const wchar_t lower = static_cast<wchar_t>(std::towlower(static_cast<wint_t>(ch)));
            hash ^= static_cast<uint64_t>(lower);
            hash *= 1099511628211ull;
        }
        return static_cast<size_t>(hash);
    }

    size_t operator()(const std::wstring& value) const noexcept
    {
        return (*this)(std::wstring_view{value});
    }
};

struct WStringViewNoCaseEq
{
    using is_transparent = void;

    bool operator()(std::wstring_view left, std::wstring_view right) const noexcept
    {
        if (left.size() > static_cast<size_t>(std::numeric_limits<int>::max()) || right.size() > static_cast<size_t>(std::numeric_limits<int>::max()))
        {
            return left == right;
        }
        const int leftLen  = static_cast<int>(left.size());
        const int rightLen = static_cast<int>(right.size());
        return CompareStringOrdinal(left.data(), leftLen, right.data(), rightLen, TRUE) == CSTR_EQUAL;
    }
};

enum class CompareDirectoriesDiffBit : uint32_t
{
    OnlyInLeft  = 0x01u,
    OnlyInRight = 0x02u,

    TypeMismatch = 0x04u,

    Size       = 0x08u,
    DateTime   = 0x10u,
    Attributes = 0x20u,
    Content    = 0x40u,

    SubdirAttributes = 0x80u,
    SubdirContent    = 0x100u,
};

[[nodiscard]] inline constexpr uint32_t operator|(CompareDirectoriesDiffBit a, CompareDirectoriesDiffBit b) noexcept
{
    return static_cast<uint32_t>(a) | static_cast<uint32_t>(b);
}

[[nodiscard]] inline constexpr uint32_t operator|(uint32_t a, CompareDirectoriesDiffBit b) noexcept
{
    return a | static_cast<uint32_t>(b);
}

[[nodiscard]] inline constexpr bool HasFlag(uint32_t mask, CompareDirectoriesDiffBit bit) noexcept
{
    return (mask & static_cast<uint32_t>(bit)) != 0u;
}

struct CompareDirectoriesItemDecision
{
    bool isDirectory = false;
    bool existsLeft  = false;
    bool existsRight = false;

    bool isDifferent  = false;
    bool selectLeft   = false;
    bool selectRight  = false;

    uint32_t differenceMask = 0;

    uint64_t leftSizeBytes    = 0;
    int64_t leftLastWriteTime = 0;
    DWORD leftFileAttributes  = 0;

    uint64_t rightSizeBytes    = 0;
    int64_t rightLastWriteTime = 0;
    DWORD rightFileAttributes  = 0;
};

struct CompareDirectoriesFolderDecision
{
    uint64_t version = 0;
    HRESULT hr       = S_OK;
    std::unordered_map<std::wstring, CompareDirectoriesItemDecision, WStringViewNoCaseHash, WStringViewNoCaseEq> items;
};

class CompareDirectoriesSession final : public std::enable_shared_from_this<CompareDirectoriesSession>
{
public:
    using ScanProgressCallback = std::function<void(const std::filesystem::path& relativeFolder,
                                                    std::wstring_view currentEntryName,
                                                    uint64_t scannedFolders,
                                                    uint64_t scannedEntries,
                                                    uint32_t activeScans)>;

    CompareDirectoriesSession(wil::com_ptr<IFileSystem> baseFileSystem,
                              std::filesystem::path leftRoot,
                              std::filesystem::path rightRoot,
                              Common::Settings::CompareDirectoriesSettings settings);

    CompareDirectoriesSession(const CompareDirectoriesSession&)            = delete;
    CompareDirectoriesSession& operator=(const CompareDirectoriesSession&) = delete;
    CompareDirectoriesSession(CompareDirectoriesSession&&)                 = delete;
    CompareDirectoriesSession& operator=(CompareDirectoriesSession&&)      = delete;

    void SetRoots(std::filesystem::path leftRoot, std::filesystem::path rightRoot);
    void SetSettings(Common::Settings::CompareDirectoriesSettings settings);
    void Invalidate() noexcept;
    void InvalidateForAbsolutePath(const std::filesystem::path& absolutePath, bool includeSubtree) noexcept;
    void SetScanProgressCallback(ScanProgressCallback callback) noexcept;

    [[nodiscard]] Common::Settings::CompareDirectoriesSettings GetSettings() const;
    [[nodiscard]] std::filesystem::path GetRoot(ComparePane pane) const;
    [[nodiscard]] uint64_t GetVersion() const noexcept;
    [[nodiscard]] uint64_t GetUiVersion() const noexcept;

    [[nodiscard]] wil::com_ptr<IFileSystem> GetBaseFileSystem() const noexcept;
    [[nodiscard]] wil::com_ptr<IInformations> GetBaseInformations() const noexcept;
    [[nodiscard]] wil::com_ptr<IFileSystemIO> GetBaseFileSystemIO() const noexcept;

    [[nodiscard]] std::optional<std::filesystem::path> TryMakeRelative(ComparePane pane, const std::filesystem::path& absoluteFolder) const;
    [[nodiscard]] std::filesystem::path ResolveAbsolute(ComparePane pane, const std::filesystem::path& relativeFolder) const;

    [[nodiscard]] std::shared_ptr<const CompareDirectoriesFolderDecision> GetOrComputeDecision(const std::filesystem::path& relativeFolder);

private:
    std::wstring MakeCacheKey(const std::filesystem::path& relativeFolder) const;
    void InvalidateForRelativePathLocked(const std::filesystem::path& relativePath, bool includeSubtree) noexcept;
    void NotifyScanProgress(const std::filesystem::path& relativeFolder, std::wstring_view currentEntryName, bool force) noexcept;

    wil::com_ptr<IFileSystem> _baseFileSystem;
    wil::com_ptr<IInformations> _baseInformations;
    wil::com_ptr<IFileSystemIO> _baseFileSystemIo;

    mutable std::mutex _mutex;
    std::filesystem::path _leftRoot;
    std::filesystem::path _rightRoot;
    Common::Settings::CompareDirectoriesSettings _settings;
    uint64_t _version = 1;
    uint64_t _uiVersion = 1;

    std::unordered_map<std::wstring, std::shared_ptr<const CompareDirectoriesFolderDecision>, WStringViewNoCaseHash, WStringViewNoCaseEq> _cache;

    std::atomic_uint32_t _scanActiveScans{0};
    std::atomic_uint64_t _scanFoldersScanned{0};
    std::atomic_uint64_t _scanEntriesScanned{0};
    std::atomic_uint64_t _scanLastNotifyTickMs{0};
    std::atomic<std::shared_ptr<const ScanProgressCallback>> _scanProgressCallback;
};

[[nodiscard]] wil::com_ptr<IFileSystem> CreateCompareDirectoriesFileSystem(ComparePane pane, std::shared_ptr<CompareDirectoriesSession> session) noexcept;
