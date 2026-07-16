#include "pch.h"

#include "FolderView.h"
#include "IconCache.h"
#include <array>
#include <cstdint>
#include <filesystem>
#include <fstream>
#include <string>
#include <string_view>
#include <unordered_map>
#include <vector>
using namespace Microsoft::VisualStudio::CppUnitTestFramework;

namespace
{
struct SimulatedFolderItem
{
    std::filesystem::path fullPath;
    bool isDirectory = false;

    [[nodiscard]] std::wstring GetExtension() const
    {
        if (isDirectory)
        {
            return L"<directory>";
        }

        return fullPath.extension().wstring();
    }
};

struct DuplicatePathEnumerationResult
{
    uint64_t checksum             = 0;
    size_t perFileItemCount       = 0;
    size_t uniquePerFilePathCount = 0;
};

[[nodiscard]] std::filesystem::path GetCurrentModulePath()
{
    std::wstring path(MAX_PATH, L'\0');
    for (;;)
    {
        wchar_t* const buffer = path.empty() ? nullptr : &path[0];
        const DWORD length    = GetModuleFileNameW(nullptr, buffer, static_cast<DWORD>(path.size()));
        if (length == 0)
        {
            return {};
        }

        if (length < path.size() - 1u)
        {
            path.resize(static_cast<size_t>(length));
            return std::filesystem::path(path);
        }

        path.resize(path.size() * 2u);
    }
}

void WriteSmallTextFile(const std::filesystem::path& path, std::string_view content)
{
    std::ofstream stream(path, std::ios::binary | std::ios::trunc);
    Assert::IsTrue(stream.is_open(), L"Failed to create test file.");
    stream.write(content.data(), static_cast<std::streamsize>(content.size()));
    Assert::IsTrue(stream.good(), L"Failed to write test file.");
}

void AppendDuplicateItems(const std::filesystem::path& path, const bool isDirectory, const int duplicateCount, std::vector<SimulatedFolderItem>& items)
{
    for (int i = 0; i < duplicateCount; ++i)
    {
        items.push_back({path, isDirectory});
    }
}

void BuildDuplicatePathDataSet(const std::filesystem::path& root, std::vector<SimulatedFolderItem>& items)
{
    items.clear();

    static constexpr std::array<const wchar_t*, 4> kCachedExtensions{{
        L".txt",
        L".cpp",
        L".json",
        L".log",
    }};

    for (int i = 0; i < 1200; ++i)
    {
        const wchar_t* const extension       = kCachedExtensions[static_cast<size_t>(i) % kCachedExtensions.size()];
        const std::filesystem::path filePath = root / (std::wstring(L"cached_") + std::to_wstring(i) + extension);
        WriteSmallTextFile(filePath, "payload");
        items.push_back({filePath, false});
    }

    const std::filesystem::path modulePath = GetCurrentModulePath();
    Assert::IsTrue(! modulePath.empty(), L"Failed to resolve current module path.");

    constexpr int kUniqueExeCount = 24;
    constexpr int kUniqueUrlCount = 24;
    constexpr int kDuplicateCount = 8;

    for (int i = 0; i < kUniqueExeCount; ++i)
    {
        const std::filesystem::path exePath = root / (std::wstring(L"dup_tool_") + std::to_wstring(i) + L".exe");
        std::filesystem::copy_file(modulePath, exePath, std::filesystem::copy_options::overwrite_existing);
        AppendDuplicateItems(exePath, false, kDuplicateCount, items);
    }

    for (int i = 0; i < kUniqueUrlCount; ++i)
    {
        const std::filesystem::path urlPath = root / (std::wstring(L"dup_shortcut_") + std::to_wstring(i) + L".url");
        const std::string content           = "[InternetShortcut]\r\nURL=https://example.invalid/duplicate/" + std::to_string(i) + "\r\n";
        WriteSmallTextFile(urlPath, content);
        AppendDuplicateItems(urlPath, false, kDuplicateCount, items);
    }
}

[[nodiscard]] DuplicatePathEnumerationResult EnumerateLargeFolderIconsWithDuplicatePaths(const std::vector<SimulatedFolderItem>& items)
{
    auto& cache = IconCache::GetInstance();

    std::unordered_map<std::wstring, int> extensionResults;
    std::vector<int> iconIndices(items.size(), -1);
    std::unordered_map<std::wstring, DWORD> uniqueExtensions;
    std::unordered_map<std::wstring, std::vector<size_t>> perFileItemIndicesByPath;

    for (size_t i = 0; i < items.size(); ++i)
    {
        const auto& item             = items[i];
        const std::wstring extension = item.GetExtension();
        const DWORD fileAttributes   = item.isDirectory ? FILE_ATTRIBUTE_DIRECTORY : FILE_ATTRIBUTE_NORMAL;

        const auto cachedIndex = cache.GetIconIndexByExtension(extension);
        if (cachedIndex.has_value())
        {
            iconIndices[i] = cachedIndex.value();
            continue;
        }

        if (cache.RequiresPerFileLookup(extension))
        {
            perFileItemIndicesByPath[item.fullPath.wstring()].push_back(i);
            continue;
        }

        uniqueExtensions.try_emplace(extension, fileAttributes);
    }

    for (const auto& entry : uniqueExtensions)
    {
        const auto iconIndex = cache.GetOrQueryIconIndexByExtension(entry.first, entry.second);
        if (iconIndex.has_value())
        {
            extensionResults.emplace(entry.first, iconIndex.value());
        }
    }

    for (size_t i = 0; i < items.size(); ++i)
    {
        if (iconIndices[i] >= 0)
        {
            continue;
        }

        const std::wstring extension = items[i].GetExtension();
        const auto it                = extensionResults.find(extension);
        if (it != extensionResults.end())
        {
            iconIndices[i] = it->second;
        }
    }

    for (const auto& entry : perFileItemIndicesByPath)
    {
        const auto iconIndex = cache.QuerySysIconIndexForPath(entry.first.c_str(), 0, false).value_or(-1);
        for (const size_t itemIndex : entry.second)
        {
            iconIndices[itemIndex] = iconIndex;
        }
    }

    DuplicatePathEnumerationResult result{};
    result.perFileItemCount = 0;
    for (const auto& entry : perFileItemIndicesByPath)
    {
        result.perFileItemCount += entry.second.size();
    }
    result.uniquePerFilePathCount = perFileItemIndicesByPath.size();

    for (const int iconIndex : iconIndices)
    {
        result.checksum += static_cast<uint64_t>(iconIndex + 1);
    }

    return result;
}
} // namespace

namespace RedSalamanderTests
{
TEST_CLASS(FolderIconEnumerationDuplicatePathPerfTest)
{
public:
    static std::filesystem::path s_root;
    static std::vector<SimulatedFolderItem> s_items;

#pragma warning(push)
#pragma warning(disable : 5246) // CppUnitTest TEST_CLASS_* macros expand to framework-owned registration initializers.
    TEST_CLASS_INITIALIZE(ClassInitialize)
    {
        std::error_code ec;
        s_root = PerformanceTests2::AcquirePerformanceTestSandbox(L"folder_icon_enumeration_duplicate_path_perf", ec);
        Assert::IsFalse(static_cast<bool>(ec), L"Failed to create PerformanceTests2 duplicate-path icon TestSandbox root.");
        Assert::IsFalse(s_root.empty(), L"PerformanceTests2 duplicate-path icon TestSandbox root is empty.");
        BuildDuplicatePathDataSet(s_root, s_items);
    }

    TEST_CLASS_CLEANUP(ClassCleanup)
    {
        std::error_code ec;
        std::filesystem::remove_all(s_root, ec);
        s_items.clear();
        s_root.clear();
    }
#pragma warning(pop)

    TEST_METHOD(LargeFolderIconEnumeration_DuplicatePaths)
    {
        Assert::IsTrue(! s_items.empty(), L"Test data set was not created.");

        DuplicatePathEnumerationResult result{};
        for (int iteration = 0; iteration < 5; ++iteration)
        {
            result = EnumerateLargeFolderIconsWithDuplicatePaths(s_items);
        }

        Assert::IsTrue(result.checksum != 0, L"Icon enumeration produced an unexpected checksum.");
        Assert::IsTrue(result.uniquePerFilePathCount < result.perFileItemCount, L"Duplicate-path test data did not contain coalescible per-file requests.");
    }
};

std::filesystem::path FolderIconEnumerationDuplicatePathPerfTest::s_root;
std::vector<SimulatedFolderItem> FolderIconEnumerationDuplicatePathPerfTest::s_items;
} // namespace RedSalamanderTests
