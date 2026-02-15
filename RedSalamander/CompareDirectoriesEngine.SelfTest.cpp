#include "CompareDirectoriesEngine.SelfTest.h"

#ifdef _DEBUG

#include "Framework.h"

#include <algorithm>
#include <array>
#include <cstddef>
#include <cstdint>
#include <filesystem>
#include <limits>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027)
#include <wil/com.h>
#include <wil/resource.h>
#pragma warning(pop)

#include "CompareDirectoriesEngine.h"
#include "FileSystemPluginManager.h"
#include "Helpers.h"

namespace
{
[[nodiscard]] std::filesystem::path GetTempDirectory() noexcept
{
    std::array<wchar_t, MAX_PATH + 2> buffer{};
    const DWORD len = ::GetTempPathW(static_cast<DWORD>(buffer.size()), buffer.data());
    if (len == 0 || len >= buffer.size())
    {
        return {};
    }

    return std::filesystem::path(buffer.data());
}

[[nodiscard]] std::wstring MakeGuidText() noexcept
{
    GUID guid{};
    if (FAILED(::CoCreateGuid(&guid)))
    {
        return {};
    }

    std::array<wchar_t, 64> buffer{};
    if (::StringFromGUID2(guid, buffer.data(), static_cast<int>(buffer.size())) <= 0)
    {
        return {};
    }

    std::wstring text(buffer.data());
    text.erase(std::remove(text.begin(), text.end(), L'{'), text.end());
    text.erase(std::remove(text.begin(), text.end(), L'}'), text.end());
    return text;
}

[[nodiscard]] bool EnsureDirectoryExists(const std::filesystem::path& path) noexcept
{
    std::error_code ec;
    std::filesystem::create_directories(path, ec);
    if (ec)
    {
        return false;
    }
    return std::filesystem::exists(path, ec) && std::filesystem::is_directory(path, ec) && ! ec;
}

[[nodiscard]] bool WriteFileBytes(const std::filesystem::path& path, const void* data, size_t sizeBytes) noexcept
{
    if (sizeBytes > static_cast<size_t>(std::numeric_limits<DWORD>::max()))
    {
        return false;
    }

    wil::unique_handle file(::CreateFileW(path.c_str(),
                                          GENERIC_WRITE,
                                          FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                          nullptr,
                                          CREATE_ALWAYS,
                                          FILE_ATTRIBUTE_NORMAL,
                                          nullptr));
    if (! file)
    {
        return false;
    }

    DWORD written = 0;
    if (::WriteFile(file.get(), data, static_cast<DWORD>(sizeBytes), &written, nullptr) == 0)
    {
        return false;
    }

    return written == static_cast<DWORD>(sizeBytes);
}

[[nodiscard]] bool WriteFileText(const std::filesystem::path& path, std::string_view text) noexcept
{
    return WriteFileBytes(path, text.data(), text.size());
}

[[nodiscard]] bool WriteFileFill(const std::filesystem::path& path, char ch, size_t sizeBytes) noexcept
{
    std::string text(sizeBytes, ch);
    return WriteFileText(path, text);
}

[[nodiscard]] bool SetFileLastWriteTime(const std::filesystem::path& path, const FILETIME& lastWriteTime) noexcept
{
    wil::unique_handle file(::CreateFileW(path.c_str(),
                                          FILE_WRITE_ATTRIBUTES,
                                          FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                          nullptr,
                                          OPEN_EXISTING,
                                          FILE_ATTRIBUTE_NORMAL,
                                          nullptr));
    if (! file)
    {
        return false;
    }

    return ::SetFileTime(file.get(), nullptr, nullptr, &lastWriteTime) != 0;
}

[[nodiscard]] wil::com_ptr<IFileSystem> GetLocalFileSystem() noexcept
{
    for (const FileSystemPluginManager::PluginEntry& entry : FileSystemPluginManager::GetInstance().GetPlugins())
    {
        if (CompareStringOrdinal(entry.id.c_str(), -1, L"builtin/file-system", -1, TRUE) == CSTR_EQUAL && entry.fileSystem)
        {
            return entry.fileSystem;
        }
    }
    return {};
}

struct CaseFolders
{
    std::filesystem::path left;
    std::filesystem::path right;
};

[[nodiscard]] std::optional<CaseFolders> CreateCaseFolders(const std::filesystem::path& base, std::wstring_view caseName) noexcept
{
    std::filesystem::path caseRoot = base / std::filesystem::path(caseName);
    std::filesystem::path left     = caseRoot / L"left";
    std::filesystem::path right    = caseRoot / L"right";

    if (! EnsureDirectoryExists(left) || ! EnsureDirectoryExists(right))
    {
        return std::nullopt;
    }

    return CaseFolders{std::move(left), std::move(right)};
}

struct TestState
{
    bool failed = false;

    void Require(bool condition, std::wstring_view message) noexcept
    {
        if (condition)
        {
            return;
        }

        failed = true;
        Debug::Error(L"CompareSelfTest: {0}", message);
    }
};

[[nodiscard]] std::vector<std::wstring> EnumerateDirectoryNames(const wil::com_ptr<IFileSystem>& fs,
                                                                const std::filesystem::path& folder,
                                                                TestState& state) noexcept
{
    if (! fs)
    {
        state.Require(false, L"EnumerateDirectoryNames: file system is null.");
        return {};
    }

    wil::com_ptr<IFilesInformation> info;
    const HRESULT hr = fs->ReadDirectoryInfo(folder.c_str(), info.put());
    state.Require(SUCCEEDED(hr), L"EnumerateDirectoryNames: ReadDirectoryInfo failed.");
    if (FAILED(hr) || ! info)
    {
        return {};
    }

    FileInfo* head = nullptr;
    const HRESULT hrBuffer = info->GetBuffer(&head);
    state.Require(SUCCEEDED(hrBuffer), L"EnumerateDirectoryNames: GetBuffer failed.");
    if (FAILED(hrBuffer) || head == nullptr)
    {
        return {};
    }

    std::vector<std::wstring> result;
    for (FileInfo* entry = head; entry != nullptr;)
    {
        const size_t nameChars = static_cast<size_t>(entry->FileNameSize) / sizeof(wchar_t);
        result.emplace_back(entry->FileName, nameChars);

        if (entry->NextEntryOffset == 0)
        {
            break;
        }
        entry = reinterpret_cast<FileInfo*>(reinterpret_cast<unsigned char*>(entry) + entry->NextEntryOffset);
    }

    return result;
}

[[nodiscard]] bool ContainsName(const std::vector<std::wstring>& names, std::wstring_view name) noexcept
{
    return std::any_of(names.begin(), names.end(), [&](const std::wstring& value) noexcept { return value == name; });
}

[[nodiscard]] std::shared_ptr<const CompareDirectoriesFolderDecision> ComputeRootDecision(wil::com_ptr<IFileSystem> baseFs,
                                                                                         const CaseFolders& folders,
                                                                                         Common::Settings::CompareDirectoriesSettings settings,
                                                                                         TestState& state) noexcept
{
    if (! baseFs)
    {
        state.Require(false, L"Base file system is null.");
        return {};
    }

    auto session = std::make_shared<CompareDirectoriesSession>(std::move(baseFs), folders.left, folders.right, std::move(settings));
    auto decision = session->GetOrComputeDecision(std::filesystem::path{});
    state.Require(static_cast<bool>(decision), L"GetOrComputeDecision returned null.");
    if (! decision)
    {
        return {};
    }

    state.Require(SUCCEEDED(decision->hr), L"Decision hr is failure.");
    return decision;
}

[[nodiscard]] const CompareDirectoriesItemDecision* FindItem(const CompareDirectoriesFolderDecision& decision, std::wstring_view name) noexcept
{
    const auto it = decision.items.find(std::wstring(name));
    if (it == decision.items.end())
    {
        return nullptr;
    }
    return &it->second;
}
} // namespace

bool CompareDirectoriesSelfTest::Run() noexcept
{
    Debug::Info(L"CompareSelfTest: begin");

    wil::com_ptr<IFileSystem> baseFs = GetLocalFileSystem();
    if (! baseFs)
    {
        Debug::Error(L"CompareSelfTest: local file system plugin not available.");
        return false;
    }

    const std::filesystem::path temp = GetTempDirectory();
    if (temp.empty())
    {
        Debug::Error(L"CompareSelfTest: temp directory not available.");
        return false;
    }

    const std::wstring guid = MakeGuidText();
    if (guid.empty())
    {
        Debug::Error(L"CompareSelfTest: failed to generate GUID.");
        return false;
    }

    const std::filesystem::path root = temp / (L"RedSalamander_CompareSelfTest_" + guid);
    if (! EnsureDirectoryExists(root))
    {
        Debug::Error(L"CompareSelfTest: failed to create root folder.");
        return false;
    }

    const auto cleanup = wil::scope_exit(
        [&]
        {
            std::error_code ec;
            std::filesystem::remove_all(root, ec);
        });

    TestState state{};

    // Case: Unique files/dirs selected; identical excluded by default.
    if (const auto foldersOpt = CreateCaseFolders(root, L"unique"))
    {
        const auto& folders = *foldersOpt;
        state.Require(WriteFileText(folders.left / L"only_left.txt", "L"), L"Failed to create only_left.txt (left).");
        state.Require(WriteFileText(folders.right / L"only_right.txt", "R"), L"Failed to create only_right.txt (right).");
        state.Require(EnsureDirectoryExists(folders.left / L"only_left_dir"), L"Failed to create only_left_dir (left).");
        state.Require(WriteFileText(folders.left / L"same.txt", "S"), L"Failed to create same.txt (left).");
        state.Require(WriteFileText(folders.right / L"same.txt", "S"), L"Failed to create same.txt (right).");

        auto decision = ComputeRootDecision(baseFs, folders, Common::Settings::CompareDirectoriesSettings{}, state);
        if (decision)
        {
            {
                const auto* item = FindItem(*decision, L"only_left.txt");
                state.Require(item != nullptr, L"only_left.txt missing from decision.");
                if (item)
                {
                    state.Require(item->isDifferent, L"only_left.txt expected isDifferent.");
                    state.Require(item->selectLeft && ! item->selectRight, L"only_left.txt expected selectLeft only.");
                    state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::OnlyInLeft),
                                  L"only_left.txt expected differenceMask=OnlyInLeft.");
                }
            }
            {
                const auto* item = FindItem(*decision, L"only_right.txt");
                state.Require(item != nullptr, L"only_right.txt missing from decision.");
                if (item)
                {
                    state.Require(item->isDifferent, L"only_right.txt expected isDifferent.");
                    state.Require(! item->selectLeft && item->selectRight, L"only_right.txt expected selectRight only.");
                    state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::OnlyInRight),
                                  L"only_right.txt expected differenceMask=OnlyInRight.");
                }
            }
            {
                const auto* item = FindItem(*decision, L"only_left_dir");
                state.Require(item != nullptr, L"only_left_dir missing from decision.");
                if (item)
                {
                    state.Require(item->isDirectory, L"only_left_dir expected isDirectory.");
                    state.Require(item->isDifferent, L"only_left_dir expected isDifferent.");
                    state.Require(item->selectLeft && ! item->selectRight, L"only_left_dir expected selectLeft only.");
                    state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::OnlyInLeft),
                                  L"only_left_dir expected differenceMask=OnlyInLeft.");
                }
            }
            {
                const auto* item = FindItem(*decision, L"same.txt");
                state.Require(item != nullptr, L"same.txt missing from decision.");
                if (item)
                {
                    state.Require(! item->isDifferent, L"same.txt expected identical.");
                    state.Require(item->differenceMask == 0u, L"same.txt expected differenceMask=0.");
                }
            }

            auto session = std::make_shared<CompareDirectoriesSession>(baseFs, folders.left, folders.right, Common::Settings::CompareDirectoriesSettings{});
            const auto fsLeft  = CreateCompareDirectoriesFileSystem(ComparePane::Left, session);
            const auto fsRight = CreateCompareDirectoriesFileSystem(ComparePane::Right, session);

            const auto leftNames  = EnumerateDirectoryNames(fsLeft, folders.left, state);
            const auto rightNames = EnumerateDirectoryNames(fsRight, folders.right, state);

            state.Require(ContainsName(leftNames, L"only_left.txt"), L"only_left.txt expected in left enumeration.");
            state.Require(! ContainsName(leftNames, L"only_right.txt"), L"only_right.txt unexpected in left enumeration.");
            state.Require(! ContainsName(leftNames, L"same.txt"), L"same.txt expected excluded in left enumeration.");

            state.Require(ContainsName(rightNames, L"only_right.txt"), L"only_right.txt expected in right enumeration.");
            state.Require(! ContainsName(rightNames, L"only_left.txt"), L"only_left.txt unexpected in right enumeration.");
            state.Require(! ContainsName(rightNames, L"same.txt"), L"same.txt expected excluded in right enumeration.");
        }
    }
    else
    {
        state.Require(false, L"Failed to create case folders: unique.");
    }

    // Case: File vs directory mismatch selects both sides.
    if (const auto foldersOpt = CreateCaseFolders(root, L"typemismatch"))
    {
        const auto& folders = *foldersOpt;
        state.Require(WriteFileText(folders.left / L"mix", "F"), L"Failed to create mix file (left).");
        state.Require(EnsureDirectoryExists(folders.right / L"mix"), L"Failed to create mix directory (right).");

        auto decision = ComputeRootDecision(baseFs, folders, Common::Settings::CompareDirectoriesSettings{}, state);
        if (decision)
        {
            const auto* item = FindItem(*decision, L"mix");
            state.Require(item != nullptr, L"mix missing from decision.");
            if (item)
            {
                state.Require(item->isDifferent, L"mix expected isDifferent on type mismatch.");
                state.Require(item->selectLeft && item->selectRight, L"mix expected select both on type mismatch.");
                state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::TypeMismatch),
                              L"mix expected differenceMask=TypeMismatch.");
            }
        }
    }
    else
    {
        state.Require(false, L"Failed to create case folders: typemismatch.");
    }

    // Case: Size compare selects bigger file.
    if (const auto foldersOpt = CreateCaseFolders(root, L"size"))
    {
        const auto& folders = *foldersOpt;
        state.Require(WriteFileFill(folders.left / L"a.bin", 'A', 200), L"Failed to create a.bin (left).");
        state.Require(WriteFileFill(folders.right / L"a.bin", 'B', 100), L"Failed to create a.bin (right).");

        Common::Settings::CompareDirectoriesSettings settings{};
        settings.compareSize = true;

        auto decision = ComputeRootDecision(baseFs, folders, settings, state);
        if (decision)
        {
            const auto* item = FindItem(*decision, L"a.bin");
            state.Require(item != nullptr, L"a.bin missing from decision.");
            if (item)
            {
                state.Require(item->isDifferent, L"a.bin expected isDifferent with compareSize.");
                state.Require(item->selectLeft && ! item->selectRight, L"a.bin expected selectLeft only when left is bigger.");
                state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::Size), L"a.bin expected differenceMask=Size.");
            }
        }
    }
    else
    {
        state.Require(false, L"Failed to create case folders: size.");
    }

    // Case: Date/time compare selects newer file.
    if (const auto foldersOpt = CreateCaseFolders(root, L"time"))
    {
        const auto& folders = *foldersOpt;
        state.Require(WriteFileText(folders.left / L"a.txt", "T"), L"Failed to create a.txt (left).");
        state.Require(WriteFileText(folders.right / L"a.txt", "T"), L"Failed to create a.txt (right).");

        FILETIME now{};
        ::GetSystemTimeAsFileTime(&now);
        ULARGE_INTEGER newer{};
        newer.LowPart  = now.dwLowDateTime;
        newer.HighPart = now.dwHighDateTime;
        newer.QuadPart += 60ull * 10'000'000ull;

        FILETIME leftFt{};
        leftFt.dwLowDateTime  = newer.LowPart;
        leftFt.dwHighDateTime = newer.HighPart;

        state.Require(SetFileLastWriteTime(folders.left / L"a.txt", leftFt), L"Failed to set a.txt last write time (left).");
        state.Require(SetFileLastWriteTime(folders.right / L"a.txt", now), L"Failed to set a.txt last write time (right).");

        Common::Settings::CompareDirectoriesSettings settings{};
        settings.compareDateTime = true;

        auto decision = ComputeRootDecision(baseFs, folders, settings, state);
        if (decision)
        {
            const auto* item = FindItem(*decision, L"a.txt");
            state.Require(item != nullptr, L"a.txt missing from decision.");
            if (item)
            {
                state.Require(item->isDifferent, L"a.txt expected isDifferent with compareDateTime.");
                state.Require(item->selectLeft && ! item->selectRight, L"a.txt expected selectLeft only when left is newer.");
                state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::DateTime), L"a.txt expected differenceMask=DateTime.");
            }
        }
    }
    else
    {
        state.Require(false, L"Failed to create case folders: time.");
    }

    // Case: Attribute compare selects both sides.
    if (const auto foldersOpt = CreateCaseFolders(root, L"attributes"))
    {
        const auto& folders = *foldersOpt;
        state.Require(WriteFileText(folders.left / L"a.txt", "A"), L"Failed to create a.txt (left).");
        state.Require(WriteFileText(folders.right / L"a.txt", "A"), L"Failed to create a.txt (right).");

        const std::filesystem::path leftPath = folders.left / L"a.txt";
        const DWORD leftAttrs               = ::GetFileAttributesW(leftPath.c_str());
        state.Require(leftAttrs != INVALID_FILE_ATTRIBUTES, L"GetFileAttributesW failed for a.txt (left).");
        if (leftAttrs != INVALID_FILE_ATTRIBUTES)
        {
            state.Require(::SetFileAttributesW(leftPath.c_str(), leftAttrs | FILE_ATTRIBUTE_HIDDEN) != 0, L"SetFileAttributesW failed for a.txt (left).");
        }

        Common::Settings::CompareDirectoriesSettings settings{};
        settings.compareAttributes = true;

        auto decision = ComputeRootDecision(baseFs, folders, settings, state);
        if (decision)
        {
            const auto* item = FindItem(*decision, L"a.txt");
            state.Require(item != nullptr, L"a.txt missing from decision.");
            if (item)
            {
                state.Require(item->isDifferent, L"a.txt expected isDifferent with compareAttributes.");
                state.Require(item->selectLeft && item->selectRight, L"a.txt expected select both when attributes differ.");
                state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::Attributes), L"a.txt expected differenceMask=Attributes.");
            }
        }
    }
    else
    {
        state.Require(false, L"Failed to create case folders: attributes.");
    }

    // Case: Content compare selects both sides.
    if (const auto foldersOpt = CreateCaseFolders(root, L"content"))
    {
        const auto& folders = *foldersOpt;
        state.Require(WriteFileFill(folders.left / L"a.bin", 'X', 64), L"Failed to create a.bin (left).");
        state.Require(WriteFileFill(folders.right / L"a.bin", 'Y', 64), L"Failed to create a.bin (right).");

        Common::Settings::CompareDirectoriesSettings settings{};
        settings.compareContent = true;

        auto decision = ComputeRootDecision(baseFs, folders, settings, state);
        if (decision)
        {
            const auto* item = FindItem(*decision, L"a.bin");
            state.Require(item != nullptr, L"a.bin missing from decision.");
            if (item)
            {
                state.Require(item->isDifferent, L"a.bin expected isDifferent with compareContent.");
                state.Require(item->selectLeft && item->selectRight, L"a.bin expected select both when content differs.");
                state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::Content), L"a.bin expected differenceMask=Content.");
            }
        }
    }
    else
    {
        state.Require(false, L"Failed to create case folders: content.");
    }

    // Case: Subdirectory content compare selects both directories.
    if (const auto foldersOpt = CreateCaseFolders(root, L"subdirs"))
    {
        const auto& folders = *foldersOpt;
        state.Require(EnsureDirectoryExists(folders.left / L"sub"), L"Failed to create sub (left).");
        state.Require(EnsureDirectoryExists(folders.right / L"sub"), L"Failed to create sub (right).");
        state.Require(WriteFileText(folders.left / L"sub" / L"child.txt", "C"), L"Failed to create sub\\child.txt (left).");

        Common::Settings::CompareDirectoriesSettings settings{};
        settings.compareSubdirectories = true;

        auto decision = ComputeRootDecision(baseFs, folders, settings, state);
        if (decision)
        {
            const auto* item = FindItem(*decision, L"sub");
            state.Require(item != nullptr, L"sub missing from decision.");
            if (item)
            {
                state.Require(item->isDirectory, L"sub expected isDirectory.");
                state.Require(item->isDifferent, L"sub expected isDifferent with compareSubdirectories.");
                state.Require(item->selectLeft && item->selectRight, L"sub expected select both when content differs.");
                state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::SubdirContent),
                              L"sub expected differenceMask=SubdirContent.");
            }
        }
    }
    else
    {
        state.Require(false, L"Failed to create case folders: subdirs.");
    }

    // Case: Compare attributes of subdirectories selects both.
    if (const auto foldersOpt = CreateCaseFolders(root, L"subdirattrs"))
    {
        const auto& folders = *foldersOpt;
        state.Require(EnsureDirectoryExists(folders.left / L"sub"), L"Failed to create sub (left).");
        state.Require(EnsureDirectoryExists(folders.right / L"sub"), L"Failed to create sub (right).");

        const std::filesystem::path leftDir = folders.left / L"sub";
        const DWORD leftAttrs              = ::GetFileAttributesW(leftDir.c_str());
        state.Require(leftAttrs != INVALID_FILE_ATTRIBUTES, L"GetFileAttributesW failed for sub (left).");
        if (leftAttrs != INVALID_FILE_ATTRIBUTES)
        {
            state.Require(::SetFileAttributesW(leftDir.c_str(), leftAttrs | FILE_ATTRIBUTE_HIDDEN) != 0, L"SetFileAttributesW failed for sub (left).");
        }

        Common::Settings::CompareDirectoriesSettings settings{};
        settings.compareSubdirectoryAttributes = true;

        auto decision = ComputeRootDecision(baseFs, folders, settings, state);
        if (decision)
        {
            const auto* item = FindItem(*decision, L"sub");
            state.Require(item != nullptr, L"sub missing from decision.");
            if (item)
            {
                state.Require(item->isDirectory, L"sub expected isDirectory.");
                state.Require(item->isDifferent, L"sub expected isDifferent with compareSubdirectoryAttributes.");
                state.Require(item->selectLeft && item->selectRight, L"sub expected select both when attributes differ.");
                state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::SubdirAttributes),
                              L"sub expected differenceMask=SubdirAttributes.");
            }
        }
    }
    else
    {
        state.Require(false, L"Failed to create case folders: subdirattrs.");
    }

    // Case: Ignore patterns exclude files/directories.
    if (const auto foldersOpt = CreateCaseFolders(root, L"ignore"))
    {
        const auto& folders = *foldersOpt;
        state.Require(WriteFileText(folders.left / L"ignore.log", "I"), L"Failed to create ignore.log (left).");
        state.Require(WriteFileText(folders.left / L"keep.txt", "K"), L"Failed to create keep.txt (left).");
        state.Require(EnsureDirectoryExists(folders.left / L"ignore_dir"), L"Failed to create ignore_dir (left).");
        state.Require(EnsureDirectoryExists(folders.left / L"keep_dir"), L"Failed to create keep_dir (left).");

        Common::Settings::CompareDirectoriesSettings settings{};
        settings.ignoreFiles               = true;
        settings.ignoreFilesPatterns       = L"*.log";
        settings.ignoreDirectories         = true;
        settings.ignoreDirectoriesPatterns = L"ignore*";

        auto decision = ComputeRootDecision(baseFs, folders, settings, state);
        if (decision)
        {
            state.Require(FindItem(*decision, L"keep.txt") != nullptr, L"keep.txt expected in decision.");
            state.Require(FindItem(*decision, L"ignore.log") == nullptr, L"ignore.log expected to be ignored.");
            state.Require(FindItem(*decision, L"keep_dir") != nullptr, L"keep_dir expected in decision.");
            state.Require(FindItem(*decision, L"ignore_dir") == nullptr, L"ignore_dir expected to be ignored.");
        }
    }
    else
    {
        state.Require(false, L"Failed to create case folders: ignore.");
    }

    // Case: showIdenticalItems includes identical files.
    if (const auto foldersOpt = CreateCaseFolders(root, L"identical"))
    {
        const auto& folders = *foldersOpt;
        state.Require(WriteFileText(folders.left / L"same.txt", "SAME"), L"Failed to create same.txt (left).");
        state.Require(WriteFileText(folders.right / L"same.txt", "SAME"), L"Failed to create same.txt (right).");

        Common::Settings::CompareDirectoriesSettings settings{};
        auto session = std::make_shared<CompareDirectoriesSession>(baseFs, folders.left, folders.right, settings);
        const auto fsLeft  = CreateCompareDirectoriesFileSystem(ComparePane::Left, session);
        const auto fsRight = CreateCompareDirectoriesFileSystem(ComparePane::Right, session);

        const uint64_t versionBefore = session->GetVersion();
        const auto decisionBefore    = session->GetOrComputeDecision(std::filesystem::path{});
        state.Require(static_cast<bool>(decisionBefore), L"Decision missing (before showIdentical).");
        if (decisionBefore)
        {
            const auto* item = FindItem(*decisionBefore, L"same.txt");
            state.Require(item != nullptr, L"same.txt missing from decision (before showIdentical).");
            if (item)
            {
                state.Require(! item->isDifferent, L"same.txt expected identical (before showIdentical).");
                state.Require(item->differenceMask == 0u, L"same.txt expected differenceMask=0 (before showIdentical).");
            }
        }

        state.Require(! ContainsName(EnumerateDirectoryNames(fsLeft, folders.left, state), L"same.txt"),
                      L"same.txt expected excluded from left enumeration (before showIdentical).");
        state.Require(! ContainsName(EnumerateDirectoryNames(fsRight, folders.right, state), L"same.txt"),
                      L"same.txt expected excluded from right enumeration (before showIdentical).");

        settings.showIdenticalItems = true;
        session->SetSettings(settings);

        const uint64_t versionAfter = session->GetVersion();
        state.Require(versionAfter == versionBefore, L"SetSettings(showIdenticalItems) should not invalidate decisions.");

        const auto decisionAfter = session->GetOrComputeDecision(std::filesystem::path{});
        state.Require(decisionAfter == decisionBefore, L"Decision should remain cached across showIdenticalItems toggle.");

        state.Require(ContainsName(EnumerateDirectoryNames(fsLeft, folders.left, state), L"same.txt"),
                      L"same.txt expected included in left enumeration (after showIdentical).");
        state.Require(ContainsName(EnumerateDirectoryNames(fsRight, folders.right, state), L"same.txt"),
                      L"same.txt expected included in right enumeration (after showIdentical).");
    }
    else
    {
        state.Require(false, L"Failed to create case folders: identical.");
    }

    if (state.failed)
    {
        Debug::Error(L"CompareSelfTest: failed.");
        return false;
    }

    Debug::Info(L"CompareSelfTest: passed.");
    return true;
}

#endif // _DEBUG
