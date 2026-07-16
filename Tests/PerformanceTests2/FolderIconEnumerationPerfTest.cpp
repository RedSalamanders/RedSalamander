#include "pch.h"

#include "IconCache.h"
#include "WSLDistro.h"

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

struct ExtensionQuery
{
    std::wstring extension;
    DWORD fileAttributes = 0;
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

void BuildLargeFolderDataSet(const std::filesystem::path& root, std::vector<SimulatedFolderItem>& items)
{
    items.clear();

    static constexpr std::array<const wchar_t*, 5> kCachedExtensions{{
        L".txt",
        L".cpp",
        L".json",
        L".log",
        L".md",
    }};

    for (int i = 0; i < 1500; ++i)
    {
        const wchar_t* const extension       = kCachedExtensions[static_cast<size_t>(i) % kCachedExtensions.size()];
        const std::filesystem::path filePath = root / (std::wstring(L"file_") + std::to_wstring(i) + extension);
        WriteSmallTextFile(filePath, "payload");
        items.push_back({filePath, false});
    }

    for (int i = 0; i < 150; ++i)
    {
        const std::filesystem::path dirPath = root / (std::wstring(L"dir_") + std::to_wstring(i));
        std::filesystem::create_directory(dirPath);
        items.push_back({dirPath, true});
    }

    const std::filesystem::path modulePath = GetCurrentModulePath();
    Assert::IsTrue(! modulePath.empty(), L"Failed to resolve current module path.");

    for (int i = 0; i < 96; ++i)
    {
        const std::filesystem::path exePath = root / (std::wstring(L"tool_") + std::to_wstring(i) + L".exe");
        std::filesystem::copy_file(modulePath, exePath, std::filesystem::copy_options::overwrite_existing);
        items.push_back({exePath, false});
    }

    for (int i = 0; i < 96; ++i)
    {
        const std::filesystem::path urlPath = root / (std::wstring(L"shortcut_") + std::to_wstring(i) + L".url");
        const std::string content           = "[InternetShortcut]\r\nURL=https://example.invalid/" + std::to_string(i) + "\r\n";
        WriteSmallTextFile(urlPath, content);
        items.push_back({urlPath, false});
    }
}

[[nodiscard]] uint64_t EnumerateLargeFolderIcons(const std::vector<SimulatedFolderItem>& items)
{
    auto& cache = IconCache::GetInstance();

    std::unordered_map<std::wstring, int> extensionResults;
    std::vector<size_t> perFileIconIndices;
    std::vector<int> iconIndices(items.size(), -1);
    std::unordered_map<std::wstring, ExtensionQuery> uniqueExtensions;

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
            perFileIconIndices.push_back(i);
            continue;
        }

        uniqueExtensions.try_emplace(extension, ExtensionQuery{extension, fileAttributes});
    }

    for (const auto& entry : uniqueExtensions)
    {
        const std::wstring& extension = entry.first;
        const ExtensionQuery& query   = entry.second;
        const auto iconIndex          = cache.GetOrQueryIconIndexByExtension(query.extension, query.fileAttributes);
        if (iconIndex.has_value())
        {
            extensionResults.emplace(extension, iconIndex.value());
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

    for (const size_t index : perFileIconIndices)
    {
        const auto iconIndex = cache.QuerySysIconIndexForPath(items[index].fullPath.c_str(), 0, false);
        iconIndices[index]   = iconIndex.value_or(-1);
    }

    uint64_t checksum = 0;
    for (const int iconIndex : iconIndices)
    {
        checksum += static_cast<uint64_t>(iconIndex + 1);
    }
    return checksum;
}
} // namespace

wil::unique_hicon WSLDistro::LoadDistributionIcon(const std::wstring&, int) noexcept
{
    return {};
}

namespace RedSalamanderTests
{
TEST_CLASS(FolderIconEnumerationPerfTest)
{
public:
    static std::filesystem::path s_root;
    static std::vector<SimulatedFolderItem> s_items;

#pragma warning(push)
#pragma warning(disable : 5246) // CppUnitTest TEST_CLASS_* macros expand to framework-owned registration initializers.
    TEST_CLASS_INITIALIZE(ClassInitialize)
    {
        std::error_code ec;
        s_root = PerformanceTests2::AcquirePerformanceTestSandbox(L"folder_icon_enumeration_perf", ec);
        Assert::IsFalse(static_cast<bool>(ec), L"Failed to create PerformanceTests2 icon-enumeration TestSandbox root.");
        Assert::IsFalse(s_root.empty(), L"PerformanceTests2 icon-enumeration TestSandbox root is empty.");
        BuildLargeFolderDataSet(s_root, s_items);
    }

    TEST_CLASS_CLEANUP(ClassCleanup)
    {
        std::error_code ec;
        std::filesystem::remove_all(s_root, ec);
        s_items.clear();
        s_root.clear();
    }
#pragma warning(pop)

    TEST_METHOD(LargeFolderIconEnumeration_MixedItems)
    {
        Assert::IsTrue(! s_items.empty(), L"Test data set was not created.");

        uint64_t checksum = 0;
        for (int iteration = 0; iteration < 5; ++iteration)
        {
            checksum ^= EnumerateLargeFolderIcons(s_items);
        }

        Assert::IsTrue(checksum != 0, L"Icon enumeration produced an unexpected checksum.");
    }
};

std::filesystem::path FolderIconEnumerationPerfTest::s_root;
std::vector<SimulatedFolderItem> FolderIconEnumerationPerfTest::s_items;
} // namespace RedSalamanderTests
