#include "FileActionLauncher.h"

#include <algorithm>
#include <array>
#include <atomic>
#include <chrono>
#include <cstddef>
#include <format>
#include <limits>
#include <memory>
#include <new>
#include <shellapi.h>
#include <span>
#include <vector>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 28182)
#include <wil/filesystem.h>
#include <wil/resource.h>
#pragma warning(pop)

#include "AppDataPaths.h"
#include "Helpers.h"
#include "HandleIo.h"
#include "PathUtils.h"
#include "ProcessCommandLine.h"

namespace
{
[[nodiscard]] bool IsSupportedMacroName(const std::wstring_view name) noexcept
{
    return name == L"Path" || name == L"FullPath" || name == L"PathAndFilename" || name == L"Filename" || name == L"SelectedPathsFile" ||
           name == L"OppositePanePath" || name == L"ComputerName";
}

[[nodiscard]] HRESULT MissingMacroValue() noexcept
{
    return HRESULT_FROM_WIN32(ERROR_BAD_ARGUMENTS);
}

[[nodiscard]] HRESULT InvalidMacroSyntax() noexcept
{
    return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
}

// This content-only operation has deliberately different semantics from quoting
// one complete argv item: the template already supplies the surrounding quotes.
void AppendWindowsQuotedArgumentContent(std::wstring& out, std::wstring_view value)
{
    // These different semantics are required only when the template owns both quotes.
    size_t backslashes = 0u;
    for (const wchar_t ch : value)
    {
        if (ch == L'\\')
        {
            ++backslashes;
            continue;
        }

        if (ch == L'"')
        {
            out.append((backslashes * 2u) + 1u, L'\\');
            out.push_back(L'"');
            backslashes = 0u;
            continue;
        }

        out.append(backslashes, L'\\');
        backslashes = 0u;
        out.push_back(ch);
    }

    out.append(backslashes * 2u, L'\\');
}

[[nodiscard]] bool IsMacroWrappedByTemplateQuotes(std::wstring_view templateText, size_t open, size_t close) noexcept
{
    return open > 0u && close + 1u < templateText.size() && templateText[open - 1u] == L'"' && templateText[close + 1u] == L'"';
}

[[nodiscard]] std::filesystem::path CurrentDirectoryForContext(const FileActionLauncher::MacroContext& context)
{
    if (! context.currentDirectory.empty())
    {
        return context.currentDirectory;
    }
    if (! context.itemPath.empty())
    {
        return context.itemPath.parent_path();
    }
    return {};
}

[[nodiscard]] HRESULT ResolveMacro(std::wstring_view name, const FileActionLauncher::MacroContext& context, std::wstring& value) noexcept
{
    value.clear();

    if (name == L"Path")
    {
        const std::filesystem::path path = CurrentDirectoryForContext(context);
        if (path.empty())
        {
            return MissingMacroValue();
        }
        value = path.wstring();
        return S_OK;
    }

    if (name == L"FullPath" || name == L"PathAndFilename")
    {
        if (context.itemPath.empty())
        {
            return MissingMacroValue();
        }
        value = context.itemPath.wstring();
        return S_OK;
    }

    if (name == L"Filename")
    {
        if (context.itemPath.empty())
        {
            return MissingMacroValue();
        }
        value = context.itemPath.filename().wstring();
        if (value.empty())
        {
            return MissingMacroValue();
        }
        return S_OK;
    }

    if (name == L"SelectedPathsFile")
    {
        if (context.selectedPathsFile.empty())
        {
            return MissingMacroValue();
        }
        value = context.selectedPathsFile.wstring();
        return S_OK;
    }

    if (name == L"OppositePanePath")
    {
        if (context.oppositePanePath.empty())
        {
            return MissingMacroValue();
        }
        value = context.oppositePanePath.wstring();
        return S_OK;
    }

    if (name == L"ComputerName")
    {
        if (context.computerName.empty())
        {
            return MissingMacroValue();
        }
        value = context.computerName;
        return S_OK;
    }

    return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
}

[[nodiscard]] HRESULT TemplateReferencesMacro(std::wstring_view templateText, std::wstring_view macroName, bool& references) noexcept
{
    references = false;

    for (size_t index = 0; index < templateText.size();)
    {
        const wchar_t ch = templateText[index];
        if (ch == L'{')
        {
            if (index + 1u < templateText.size() && templateText[index + 1u] == L'{')
            {
                index += 2u;
                continue;
            }

            const size_t close = templateText.find(L'}', index + 1u);
            if (close == std::wstring_view::npos)
            {
                return InvalidMacroSyntax();
            }

            const std::wstring_view name = templateText.substr(index + 1u, close - index - 1u);
            if (name.empty())
            {
                return InvalidMacroSyntax();
            }
            if (name == macroName)
            {
                references = true;
                return S_OK;
            }

            index = close + 1u;
            continue;
        }

        if (ch == L'}')
        {
            if (index + 1u < templateText.size() && templateText[index + 1u] == L'}')
            {
                index += 2u;
                continue;
            }
            return InvalidMacroSyntax();
        }

        ++index;
    }

    return S_OK;
}

[[nodiscard]] HRESULT ActionReferencesSelectedPathsFile(const Common::Settings::FileActionDefinition& action, bool& references) noexcept
{
    references = false;

    bool currentReferences = false;
    if (const HRESULT hr = TemplateReferencesMacro(action.executablePath, L"SelectedPathsFile", currentReferences); FAILED(hr))
    {
        return hr;
    }
    references = references || currentReferences;

    if (const HRESULT hr = TemplateReferencesMacro(action.arguments, L"SelectedPathsFile", currentReferences); FAILED(hr))
    {
        return hr;
    }
    references = references || currentReferences;

    if (const HRESULT hr = TemplateReferencesMacro(action.workingDirectory, L"SelectedPathsFile", currentReferences); FAILED(hr))
    {
        return hr;
    }
    references = references || currentReferences;

    return S_OK;
}

[[nodiscard]] HRESULT ActionPathReferences(const Common::Settings::FileActionDefinition& action,
                                           FileActionLauncher::ExternalActionPathReferences& references) noexcept
{
    references = {};

    const std::array templates{std::wstring_view(action.executablePath),
                               std::wstring_view(action.arguments),
                               std::wstring_view(action.workingDirectory)};
    const auto referencesAny = [&](const std::span<const std::wstring_view> macroNames, bool& result) noexcept -> HRESULT
    {
        result = false;
        for (const std::wstring_view templateText : templates)
        {
            for (const std::wstring_view macroName : macroNames)
            {
                bool current = false;
                if (const HRESULT hr = TemplateReferencesMacro(templateText, macroName, current); FAILED(hr))
                {
                    return hr;
                }
                result = result || current;
            }
        }
        return S_OK;
    };

    constexpr std::array itemMacros{std::wstring_view(L"FullPath"), std::wstring_view(L"PathAndFilename"), std::wstring_view(L"Filename")};
    constexpr std::array currentDirectoryMacros{std::wstring_view(L"Path")};
    constexpr std::array oppositePaneMacros{std::wstring_view(L"OppositePanePath")};
    constexpr std::array selectedPathsMacros{std::wstring_view(L"SelectedPathsFile")};
    if (const HRESULT hr = referencesAny(itemMacros, references.itemPath); FAILED(hr))
    {
        return hr;
    }
    if (const HRESULT hr = referencesAny(currentDirectoryMacros, references.currentDirectory); FAILED(hr))
    {
        return hr;
    }
    if (const HRESULT hr = referencesAny(oppositePaneMacros, references.oppositePanePath); FAILED(hr))
    {
        return hr;
    }
    return referencesAny(selectedPathsMacros, references.selectedPathsFile);
}

[[nodiscard]] HRESULT ExpandMacrosInternal(std::wstring_view templateText,
                                           const FileActionLauncher::MacroContext& context,
                                           bool quoteArgumentMacros,
                                           std::wstring& out) noexcept
{
    out.clear();
    out.reserve(templateText.size());

    for (size_t index = 0; index < templateText.size();)
    {
        const wchar_t ch = templateText[index];
        if (ch == L'{')
        {
            if (index + 1u < templateText.size() && templateText[index + 1u] == L'{')
            {
                out.push_back(L'{');
                index += 2u;
                continue;
            }

            const size_t close = templateText.find(L'}', index + 1u);
            if (close == std::wstring_view::npos)
            {
                return InvalidMacroSyntax();
            }

            const std::wstring_view name = templateText.substr(index + 1u, close - index - 1u);
            if (name.empty())
            {
                return InvalidMacroSyntax();
            }

            std::wstring value;
            const HRESULT hr = ResolveMacro(name, context, value);
            if (FAILED(hr))
            {
                return hr;
            }

            if (! quoteArgumentMacros)
            {
                out.append(value);
            }
            else if (IsMacroWrappedByTemplateQuotes(templateText, index, close))
            {
                AppendWindowsQuotedArgumentContent(out, value);
            }
            else
            {
                out.append(Common::Process::QuoteWindowsCommandLineArgument(value));
            }

            index = close + 1u;
            continue;
        }

        if (ch == L'}')
        {
            if (index + 1u < templateText.size() && templateText[index + 1u] == L'}')
            {
                out.push_back(L'}');
                index += 2u;
                continue;
            }

            return InvalidMacroSyntax();
        }

        out.push_back(ch);
        ++index;
    }

    return S_OK;
}

[[nodiscard]] HRESULT ExpandArgumentMacros(std::wstring_view templateText, const FileActionLauncher::MacroContext& context, std::wstring& out) noexcept
{
    return ExpandMacrosInternal(templateText, context, true, out);
}

constexpr std::wstring_view kSelectedPathsCompanyDirectory{L"RedSalamander"};
constexpr std::wstring_view kSelectedPathsManifestDirectory{L"SelectedPaths"};
constexpr std::wstring_view kSelectedPathsManifestPrefix{L"selected-paths-"};
constexpr std::wstring_view kSelectedPathsManifestSuffix{L".txt"};

[[nodiscard]] bool EqualFileIdentity(const FILE_ID_INFO& left, const FILE_ID_INFO& right) noexcept
{
    return left.VolumeSerialNumber == right.VolumeSerialNumber &&
           std::equal(std::begin(left.FileId.Identifier), std::end(left.FileId.Identifier), std::begin(right.FileId.Identifier));
}

[[nodiscard]] HRESULT QueryFileIdentity(HANDLE file, FILE_ID_INFO& identity) noexcept
{
    identity = {};
    if (file == nullptr || file == INVALID_HANDLE_VALUE)
    {
        return E_INVALIDARG;
    }
    if (GetFileInformationByHandleEx(file, FileIdInfo, &identity, sizeof(identity)) == FALSE)
    {
        const DWORD error = GetLastError();
        return HRESULT_FROM_WIN32(error == ERROR_SUCCESS ? ERROR_INVALID_DATA : error);
    }
    return S_OK;
}

[[nodiscard]] HRESULT QueryFileAttributes(HANDLE file, FILE_ATTRIBUTE_TAG_INFO& attributes) noexcept
{
    attributes = {};
    if (GetFileInformationByHandleEx(file, FileAttributeTagInfo, &attributes, sizeof(attributes)) == FALSE)
    {
        const DWORD error = GetLastError();
        return HRESULT_FROM_WIN32(error == ERROR_SUCCESS ? ERROR_INVALID_DATA : error);
    }
    return S_OK;
}

[[nodiscard]] HRESULT OpenValidatedSelectedPathsDirectory(std::wstring_view path,
                                                          DWORD shareMode,
                                                          wil::unique_hfile& directory,
                                                          FILE_ID_INFO& identity) noexcept
{
    directory.reset();
    identity = {};
    if (path.empty())
    {
        return E_INVALIDARG;
    }

    wil::unique_hfile candidate(CreateFileW(std::wstring(path).c_str(),
                                            FILE_READ_ATTRIBUTES,
                                            shareMode,
                                            nullptr,
                                            OPEN_EXISTING,
                                            FILE_FLAG_BACKUP_SEMANTICS | FILE_FLAG_OPEN_REPARSE_POINT,
                                            nullptr));
    if (! candidate)
    {
        const DWORD error = GetLastError();
        return HRESULT_FROM_WIN32(error == ERROR_SUCCESS ? ERROR_PATH_NOT_FOUND : error);
    }

    FILE_ATTRIBUTE_TAG_INFO attributes{};
    if (const HRESULT hr = QueryFileAttributes(candidate.get(), attributes); FAILED(hr))
    {
        return hr;
    }
    if ((attributes.FileAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0u)
    {
        return HRESULT_FROM_WIN32(ERROR_DIRECTORY);
    }
    if ((attributes.FileAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0u)
    {
        return HRESULT_FROM_WIN32(ERROR_CANT_ACCESS_FILE);
    }
    if (const HRESULT hr = QueryFileIdentity(candidate.get(), identity); FAILED(hr))
    {
        return hr;
    }

    directory = std::move(candidate);
    return S_OK;
}

[[nodiscard]] HRESULT MarkFileForDeletion(HANDLE file) noexcept
{
    FILE_DISPOSITION_INFO disposition{};
    disposition.DeleteFile = TRUE;
    if (SetFileInformationByHandle(file, FileDispositionInfo, &disposition, sizeof(disposition)) == FALSE)
    {
        const DWORD error = GetLastError();
        return HRESULT_FROM_WIN32(error == ERROR_SUCCESS ? ERROR_ACCESS_DENIED : error);
    }
    return S_OK;
}

[[nodiscard]] bool EqualPathTextNoCase(const std::filesystem::path& left, const std::filesystem::path& right) noexcept
{
    const std::wstring& leftText  = left.native();
    const std::wstring& rightText = right.native();
    if (leftText.size() != rightText.size() || leftText.size() > static_cast<size_t>((std::numeric_limits<int>::max)()))
    {
        return false;
    }
    return CompareStringOrdinal(leftText.data(),
                                static_cast<int>(leftText.size()),
                                rightText.data(),
                                static_cast<int>(rightText.size()),
                                TRUE) == CSTR_EQUAL;
}

enum class SelectedPathsCleanupOutcome : uint64_t
{
    Deleted = 0u,
    MissingOrBusy,
    RootChanged,
    FileChanged,
    ReparseOrDirectory,
    UnexpectedError,
};

void CleanupSelectedPathsFile(const std::filesystem::path& file,
                              const std::filesystem::path& rootPath,
                              const FILE_ID_INFO& expectedFileIdentity,
                              const FILE_ID_INFO& expectedRootIdentity) noexcept
{
    if (file.empty() || rootPath.empty())
    {
        return;
    }

    Debug::Perf::Scope perf(L"fileaction.selected_paths_file.cleanup_us");
    perf.SetValue0(1u);
    const auto finish = [&](SelectedPathsCleanupOutcome outcome, HRESULT hr = S_OK) noexcept
    {
        perf.SetValue1(static_cast<uint64_t>(outcome));
        if (FAILED(hr))
        {
            perf.SetHr(hr);
        }
    };

    if (! EqualPathTextNoCase(file.parent_path(), rootPath))
    {
        finish(SelectedPathsCleanupOutcome::RootChanged);
        return;
    }

    wil::unique_hfile root;
    FILE_ID_INFO currentRootIdentity{};
    const HRESULT rootHr = OpenValidatedSelectedPathsDirectory(
        rootPath.native(), FILE_SHARE_READ | FILE_SHARE_WRITE, root, currentRootIdentity);
    if (FAILED(rootHr) || ! EqualFileIdentity(currentRootIdentity, expectedRootIdentity))
    {
        finish(SelectedPathsCleanupOutcome::RootChanged);
        return;
    }

    wil::unique_hfile candidate(CreateFileW(file.c_str(),
                                            DELETE | FILE_READ_ATTRIBUTES,
                                            FILE_SHARE_READ | FILE_SHARE_WRITE,
                                            nullptr,
                                            OPEN_EXISTING,
                                            FILE_FLAG_OPEN_REPARSE_POINT,
                                            nullptr));
    if (! candidate)
    {
        const DWORD error = GetLastError();
        if (error == ERROR_FILE_NOT_FOUND || error == ERROR_PATH_NOT_FOUND || error == ERROR_SHARING_VIOLATION || error == ERROR_ACCESS_DENIED)
        {
            finish(SelectedPathsCleanupOutcome::MissingOrBusy);
            return;
        }
        const HRESULT hr = HRESULT_FROM_WIN32(error == ERROR_SUCCESS ? ERROR_OPEN_FAILED : error);
        finish(SelectedPathsCleanupOutcome::UnexpectedError, hr);
        Debug::Warning(L"FileActionLauncher: exact selected-path cleanup could not open the candidate (error={}).", error);
        return;
    }

    FILE_ATTRIBUTE_TAG_INFO attributes{};
    const HRESULT attributesHr = QueryFileAttributes(candidate.get(), attributes);
    if (FAILED(attributesHr))
    {
        finish(SelectedPathsCleanupOutcome::UnexpectedError, attributesHr);
        return;
    }
    if ((attributes.FileAttributes & (FILE_ATTRIBUTE_DIRECTORY | FILE_ATTRIBUTE_REPARSE_POINT)) != 0u)
    {
        finish(SelectedPathsCleanupOutcome::ReparseOrDirectory);
        return;
    }

    FILE_ID_INFO currentFileIdentity{};
    const HRESULT identityHr = QueryFileIdentity(candidate.get(), currentFileIdentity);
    if (FAILED(identityHr))
    {
        finish(SelectedPathsCleanupOutcome::UnexpectedError, identityHr);
        return;
    }
    if (! EqualFileIdentity(currentFileIdentity, expectedFileIdentity))
    {
        finish(SelectedPathsCleanupOutcome::FileChanged);
        return;
    }

    const HRESULT deleteHr = MarkFileForDeletion(candidate.get());
    if (FAILED(deleteHr))
    {
        finish(SelectedPathsCleanupOutcome::UnexpectedError, deleteHr);
        return;
    }
    finish(SelectedPathsCleanupOutcome::Deleted);
}

[[nodiscard]] bool IsSelectedPathsManifestName(std::wstring_view fileName) noexcept
{
    if (fileName.size() <= kSelectedPathsManifestPrefix.size() + kSelectedPathsManifestSuffix.size())
    {
        return false;
    }
    return CompareStringOrdinal(fileName.data(),
                                static_cast<int>(kSelectedPathsManifestPrefix.size()),
                                kSelectedPathsManifestPrefix.data(),
                                static_cast<int>(kSelectedPathsManifestPrefix.size()),
                                TRUE) == CSTR_EQUAL &&
           CompareStringOrdinal(fileName.data() + fileName.size() - kSelectedPathsManifestSuffix.size(),
                                static_cast<int>(kSelectedPathsManifestSuffix.size()),
                                kSelectedPathsManifestSuffix.data(),
                                static_cast<int>(kSelectedPathsManifestSuffix.size()),
                                TRUE) == CSTR_EQUAL;
}

struct SelectedPathsRecoveryLimits final
{
    uint64_t minimumAgeMs  = 0u;
    size_t maximumEntries  = 0u;
    size_t maximumDeletes  = 0u;
    uint64_t maximumElapsedMs = 0u;
};

struct SelectedPathsRecoveryObservation final
{
    uint64_t inspected = 0u;
    uint64_t deleted   = 0u;
    uint64_t skipped   = 0u;
    uint64_t errors    = 0u;
    uint64_t elapsedUs = 0u;
    bool stoppedByEntryBound  = false;
    bool stoppedByDeleteBound = false;
    bool stoppedByTimeBound   = false;
    bool resetMissingCursor   = false;
};

[[nodiscard]] uint64_t FileTimeTicks(const FILETIME& value) noexcept
{
    ULARGE_INTEGER ticks{};
    ticks.LowPart  = value.dwLowDateTime;
    ticks.HighPart = value.dwHighDateTime;
    return ticks.QuadPart;
}

[[nodiscard]] bool EqualLeafNoCase(std::wstring_view left, std::wstring_view right) noexcept
{
    return left.size() == right.size() && left.size() <= static_cast<size_t>((std::numeric_limits<int>::max)()) &&
           CompareStringOrdinal(left.data(), static_cast<int>(left.size()), right.data(), static_cast<int>(right.size()), TRUE) == CSTR_EQUAL;
}

[[nodiscard]] std::wstring BuildSelectedPathsChildPath(std::wstring_view root, std::wstring_view leaf)
{
    std::wstring path(root);
    if (! path.empty() && ! Common::Paths::IsSeparator(path.back()))
    {
        path.push_back(L'\\');
    }
    path.append(leaf);
    return path;
}

[[nodiscard]] HRESULT RunSelectedPathsRecoveryPass(std::wstring_view rootPath,
                                                   const SelectedPathsRecoveryLimits& limits,
                                                   std::wstring& cursor,
                                                   SelectedPathsRecoveryObservation& observation) noexcept
{
    observation = {};
    const auto startedAt = std::chrono::steady_clock::now();
    const auto finish = wil::scope_exit([&]() noexcept { observation.elapsedUs = Debug::Perf::ElapsedUs(startedAt); });
    if (rootPath.empty())
    {
        return E_INVALIDARG;
    }

    wil::unique_hfile root;
    FILE_ID_INFO rootIdentity{};
    if (const HRESULT hr = OpenValidatedSelectedPathsDirectory(
            rootPath, FILE_SHARE_READ | FILE_SHARE_WRITE, root, rootIdentity);
        FAILED(hr))
    {
        return hr;
    }

    const std::wstring searchPath = BuildSelectedPathsChildPath(
        rootPath, std::wstring(kSelectedPathsManifestPrefix) + L"*" + std::wstring(kSelectedPathsManifestSuffix));
    WIN32_FIND_DATAW data{};
    wil::unique_hfind find(FindFirstFileExW(searchPath.c_str(), FindExInfoBasic, &data, FindExSearchNameMatch, nullptr, 0u));
    if (! find)
    {
        const DWORD error = GetLastError();
        if (error == ERROR_FILE_NOT_FOUND)
        {
            cursor.clear();
            return S_OK;
        }
        return HRESULT_FROM_WIN32(error == ERROR_SUCCESS ? ERROR_GEN_FAILURE : error);
    }

    FILETIME now{};
    GetSystemTimeAsFileTime(&now);
    const uint64_t nowTicks = FileTimeTicks(now);
    const uint64_t minimumAgeTicks = limits.minimumAgeMs > (std::numeric_limits<uint64_t>::max)() / 10'000u
                                          ? (std::numeric_limits<uint64_t>::max)()
                                          : limits.minimumAgeMs * 10'000u;
    const auto deadline = startedAt + std::chrono::milliseconds{limits.maximumElapsedMs};
    bool seekingCursor  = ! cursor.empty();
    bool cursorMatched  = ! seekingCursor;
    bool stoppedByBound = false;

    for (;;)
    {
        const std::wstring_view leaf(data.cFileName);
        if (seekingCursor)
        {
            if (! EqualLeafNoCase(leaf, cursor))
            {
                if (FindNextFileW(find.get(), &data) != FALSE)
                {
                    continue;
                }
                break;
            }
            seekingCursor = false;
            cursorMatched = true;
        }

        if (observation.inspected >= limits.maximumEntries)
        {
            cursor = leaf;
            observation.stoppedByEntryBound = true;
            stoppedByBound = true;
            break;
        }
        if (std::chrono::steady_clock::now() >= deadline)
        {
            cursor = leaf;
            observation.stoppedByTimeBound = true;
            stoppedByBound = true;
            break;
        }
        ++observation.inspected;

        if (! IsSelectedPathsManifestName(leaf) ||
            (data.dwFileAttributes & (FILE_ATTRIBUTE_DIRECTORY | FILE_ATTRIBUTE_REPARSE_POINT)) != 0u)
        {
            ++observation.skipped;
        }
        else
        {
            const std::wstring candidatePath = BuildSelectedPathsChildPath(rootPath, leaf);
            wil::unique_hfile candidate(CreateFileW(candidatePath.c_str(),
                                                    DELETE | FILE_READ_ATTRIBUTES,
                                                    FILE_SHARE_READ | FILE_SHARE_WRITE,
                                                    nullptr,
                                                    OPEN_EXISTING,
                                                    FILE_FLAG_OPEN_REPARSE_POINT,
                                                    nullptr));
            if (! candidate)
            {
                const DWORD error = GetLastError();
                if (error == ERROR_FILE_NOT_FOUND || error == ERROR_PATH_NOT_FOUND || error == ERROR_SHARING_VIOLATION || error == ERROR_ACCESS_DENIED)
                {
                    ++observation.skipped;
                }
                else
                {
                    ++observation.errors;
                }
            }
            else
            {
                FILE_ATTRIBUTE_TAG_INFO attributes{};
                FILE_BASIC_INFO basic{};
                const HRESULT attributesHr = QueryFileAttributes(candidate.get(), attributes);
                if (FAILED(attributesHr) ||
                    GetFileInformationByHandleEx(candidate.get(), FileBasicInfo, &basic, sizeof(basic)) == FALSE)
                {
                    ++observation.errors;
                }
                else if ((attributes.FileAttributes & (FILE_ATTRIBUTE_DIRECTORY | FILE_ATTRIBUTE_REPARSE_POINT)) != 0u ||
                         basic.LastWriteTime.QuadPart < 0 || static_cast<uint64_t>(basic.LastWriteTime.QuadPart) > nowTicks ||
                         nowTicks - static_cast<uint64_t>(basic.LastWriteTime.QuadPart) < minimumAgeTicks)
                {
                    ++observation.skipped;
                }
                else if (observation.deleted >= limits.maximumDeletes)
                {
                    cursor = leaf;
                    observation.stoppedByDeleteBound = true;
                    stoppedByBound = true;
                    break;
                }
                else if (SUCCEEDED(MarkFileForDeletion(candidate.get())))
                {
                    ++observation.deleted;
                }
                else
                {
                    ++observation.errors;
                }
            }
        }

        if (FindNextFileW(find.get(), &data) == FALSE)
        {
            break;
        }
    }

    if (stoppedByBound)
    {
        return S_OK;
    }
    const DWORD finalError = GetLastError();
    if (! cursorMatched)
    {
        cursor.clear();
        observation.resetMissingCursor = true;
        observation.stoppedByEntryBound = true;
        return S_OK;
    }
    cursor.clear();
    return finalError == ERROR_NO_MORE_FILES ? S_OK : HRESULT_FROM_WIN32(finalError == ERROR_SUCCESS ? ERROR_GEN_FAILURE : finalError);
}

struct SelectedPathsRecoveryContext final
{
    std::wstring rootPath;
};

std::atomic_bool g_selectedPathsRecoveryScheduled{false};

void CALLBACK SelectedPathsRecoveryCallback(PTP_CALLBACK_INSTANCE /*instance*/, void* rawContext) noexcept
{
    std::unique_ptr<SelectedPathsRecoveryContext> context(static_cast<SelectedPathsRecoveryContext*>(rawContext));
    if (! context)
    {
        return;
    }

    constexpr size_t kMaximumPasses = 32u;
    constexpr SelectedPathsRecoveryLimits kLimits{24ull * 60ull * 60ull * 1000ull, 128u, 16u, 20u};
    std::wstring cursor;
    uint64_t inspected = 0u;
    uint64_t deleted   = 0u;
    uint64_t skipped   = 0u;
    uint64_t errors    = 0u;
    uint64_t boundHits = 0u;
    size_t passes      = 0u;
    HRESULT finalHr    = S_OK;
    const auto startedAt = std::chrono::steady_clock::now();
    for (; passes < kMaximumPasses; ++passes)
    {
        SelectedPathsRecoveryObservation current{};
        finalHr = RunSelectedPathsRecoveryPass(context->rootPath, kLimits, cursor, current);
        inspected += current.inspected;
        deleted += current.deleted;
        skipped += current.skipped;
        errors += current.errors;
        const bool hitBound = current.stoppedByEntryBound || current.stoppedByDeleteBound || current.stoppedByTimeBound;
        boundHits += hitBound ? 1u : 0u;
        if (FAILED(finalHr) || ! hitBound)
        {
            ++passes;
            break;
        }
    }

    Debug::Perf::EmitDurationUs(
        L"fileaction.selected_paths_recovery.elapsed_us", Debug::Perf::ElapsedUs(startedAt), inspected, deleted, finalHr);
    Debug::Perf::EmitValue(L"fileaction.selected_paths_recovery.inspected", inspected);
    Debug::Perf::EmitValue(L"fileaction.selected_paths_recovery.deleted", deleted);
    Debug::Perf::EmitValue(L"fileaction.selected_paths_recovery.skipped", skipped);
    Debug::Perf::EmitValue(L"fileaction.selected_paths_recovery.errors", errors);
    Debug::Perf::EmitValue(L"fileaction.selected_paths_recovery.bound_hits", boundHits);
    Debug::Perf::EmitValue(L"fileaction.selected_paths_recovery.passes", static_cast<uint64_t>(passes));
}

void ScheduleSelectedPathsRecovery(std::wstring_view rootPath) noexcept
{
    if (rootPath.empty() || g_selectedPathsRecoveryScheduled.exchange(true, std::memory_order_acq_rel))
    {
        return;
    }

    std::unique_ptr<SelectedPathsRecoveryContext> context(new (std::nothrow) SelectedPathsRecoveryContext());
    if (! context)
    {
        g_selectedPathsRecoveryScheduled.store(false, std::memory_order_release);
        return;
    }
    context->rootPath = rootPath;
    if (TrySubmitThreadpoolCallback(SelectedPathsRecoveryCallback, context.get(), nullptr) == FALSE)
    {
        g_selectedPathsRecoveryScheduled.store(false, std::memory_order_release);
        return;
    }
    static_cast<void>(context.release());
}

struct SelectedPathsValidationSummary final
{
    uint64_t recordCount = 0u;
    uint64_t pathBytes   = 0u;
    uint64_t totalBytes  = 2u;
};

[[nodiscard]] HRESULT AddSelectedPathsByteCount(uint64_t amount, uint64_t& total) noexcept
{
    if (amount > (std::numeric_limits<uint64_t>::max)() - total)
    {
        return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
    }
    total += amount;
    return S_OK;
}

[[nodiscard]] HRESULT ValidateSelectedPathRecord(const std::filesystem::path& selectedPath,
                                                 SelectedPathsValidationSummary& summary) noexcept
{
    if (selectedPath.empty())
    {
        return S_FALSE;
    }

    const std::wstring& record = selectedPath.native();
    if (record.find_first_of(std::wstring_view(L"\0\r\n", 3u)) != std::wstring::npos)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }
    if (record.size() > (std::numeric_limits<uint64_t>::max)() / sizeof(wchar_t))
    {
        return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
    }

    const uint64_t recordBytes = static_cast<uint64_t>(record.size()) * sizeof(wchar_t);
    if (const HRESULT hr = AddSelectedPathsByteCount(recordBytes, summary.pathBytes); FAILED(hr))
    {
        return hr;
    }
    if (const HRESULT hr = AddSelectedPathsByteCount(recordBytes, summary.totalBytes); FAILED(hr))
    {
        return hr;
    }
    if (const HRESULT hr = AddSelectedPathsByteCount(2u * sizeof(wchar_t), summary.totalBytes); FAILED(hr))
    {
        return hr;
    }
    ++summary.recordCount;
    return S_OK;
}

[[nodiscard]] HRESULT ValidateSelectedPaths(std::span<const std::filesystem::path> selectedPaths,
                                            SelectedPathsValidationSummary& summary) noexcept
{
    summary = {};
    for (const std::filesystem::path& selectedPath : selectedPaths)
    {
        const HRESULT hr = ValidateSelectedPathRecord(selectedPath, summary);
        if (FAILED(hr))
        {
            return hr;
        }
    }
    return summary.recordCount == 0u ? MissingMacroValue() : S_OK;
}

[[nodiscard]] HRESULT BuildSelectedPathsIoPath(const std::filesystem::path& path, size_t leafBudget, std::wstring& ioPath) noexcept
{
    ioPath.clear();
    if (path.empty())
    {
        return E_INVALIDARG;
    }

    const std::filesystem::path normalized = path.lexically_normal();
    const std::wstring& native             = normalized.native();
    if (! Common::Paths::IsFullyAbsoluteWindowsPath(native))
    {
        return HRESULT_FROM_WIN32(ERROR_BAD_PATHNAME);
    }
    if (native.size() > (std::numeric_limits<size_t>::max)() - leafBudget)
    {
        return HRESULT_FROM_WIN32(ERROR_FILENAME_EXCED_RANGE);
    }

    ioPath = Common::Paths::IsExtendedWindowsPath(native) || native.size() + leafBudget >= MAX_PATH
                 ? Common::Paths::ToExtendedWin32Path(native)
                 : native;
    return S_OK;
}

[[nodiscard]] HRESULT EnsureSelectedPathsDirectory(const std::filesystem::path& directory,
                                                    size_t leafBudget,
                                                    std::wstring& ioPath,
                                                    wil::unique_hfile* retainedDirectory = nullptr,
                                                    FILE_ID_INFO* retainedIdentity = nullptr) noexcept
{
    if ((retainedDirectory == nullptr) != (retainedIdentity == nullptr))
    {
        return E_INVALIDARG;
    }
    if (retainedDirectory != nullptr)
    {
        retainedDirectory->reset();
        *retainedIdentity = {};
    }
    if (const HRESULT hr = BuildSelectedPathsIoPath(directory, leafBudget, ioPath); FAILED(hr))
    {
        return hr;
    }
    if (const HRESULT hr = wil::CreateDirectoryDeepNoThrow(ioPath.c_str()); FAILED(hr))
    {
        return hr;
    }

    wil::unique_hfile directoryHandle;
    FILE_ID_INFO directoryIdentity{};
    const DWORD shareMode = retainedDirectory == nullptr ? FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE
                                                          : FILE_SHARE_READ | FILE_SHARE_WRITE;
    if (const HRESULT hr = OpenValidatedSelectedPathsDirectory(ioPath, shareMode, directoryHandle, directoryIdentity); FAILED(hr))
    {
        return hr;
    }
    if (retainedDirectory != nullptr)
    {
        *retainedIdentity  = directoryIdentity;
        *retainedDirectory = std::move(directoryHandle);
    }
    return S_OK;
}

[[nodiscard]] HRESULT ResolveSelectedPathsManifestRoot(std::wstring& rootForIo,
                                                       wil::unique_hfile& rootHandle,
                                                       FILE_ID_INFO& rootIdentity
#if defined(ENABLE_TESTS)
                                                       ,
                                                       const FileActionLauncher::Testing::SelectedPathsFileOptions* testOptions
#endif
                                                       ) noexcept
{
    constexpr size_t kManifestLeafBudget = 96u;
    rootHandle.reset();
    rootIdentity = {};
#if defined(ENABLE_TESTS)
    if (testOptions != nullptr && ! testOptions->rootOverride.empty())
    {
        return EnsureSelectedPathsDirectory(
            testOptions->rootOverride, kManifestLeafBudget, rootForIo, std::addressof(rootHandle), &rootIdentity);
    }
#endif

    const std::filesystem::path localAppData = AppDataPaths::GetLocalAppDataPath();
    if (localAppData.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND);
    }
    const std::filesystem::path companyRoot = localAppData / kSelectedPathsCompanyDirectory;
    std::wstring companyRootForIo;
    if (const HRESULT hr = EnsureSelectedPathsDirectory(companyRoot, kManifestLeafBudget, companyRootForIo); FAILED(hr))
    {
        return hr;
    }
    const HRESULT manifestRootHr =
        EnsureSelectedPathsDirectory(
            companyRoot / kSelectedPathsManifestDirectory,
            kManifestLeafBudget,
            rootForIo,
            std::addressof(rootHandle),
            &rootIdentity);
    if (SUCCEEDED(manifestRootHr))
    {
        ScheduleSelectedPathsRecovery(rootForIo);
    }
    return manifestRootHr;
}

#if defined(ENABLE_TESTS)
[[nodiscard]] HRESULT CreateSelectedPathsCollisionFixture(std::wstring_view rootForIo,
                                                          std::wstring_view candidateToken,
                                                          std::filesystem::path& pathOut) noexcept
{
    std::wstring candidate(rootForIo);
    if (! Common::Paths::IsSeparator(candidate.back()))
    {
        candidate.push_back(L'\\');
    }
    candidate.append(kSelectedPathsManifestPrefix);
    candidate.append(candidateToken);
    candidate.append(kSelectedPathsManifestSuffix);
    pathOut = std::filesystem::path(candidate);

    wil::unique_hfile collision(CreateFileW(candidate.c_str(),
                                            GENERIC_WRITE,
                                            FILE_SHARE_READ | FILE_SHARE_WRITE,
                                            nullptr,
                                            CREATE_NEW,
                                            FILE_ATTRIBUTE_TEMPORARY | FILE_ATTRIBUTE_NOT_CONTENT_INDEXED,
                                            nullptr));
    if (! collision)
    {
        const DWORD error = GetLastError();
        return error == ERROR_FILE_EXISTS || error == ERROR_ALREADY_EXISTS ? S_OK : HRESULT_FROM_WIN32(error);
    }

    constexpr std::string_view kCollisionPayload{"selected-path collision sentinel"};
    return Common::HandleIo::WriteAll(collision.get(), kCollisionPayload.data(), kCollisionPayload.size());
}
#endif

[[nodiscard]] HRESULT CreateSelectedPathsManifestHandle(std::wstring_view rootForIo,
                                                        std::wstring& filePath,
                                                        wil::unique_hfile& file
#if defined(ENABLE_TESTS)
                                                        ,
                                                        const FileActionLauncher::Testing::SelectedPathsFileOptions* testOptions,
                                                        FileActionLauncher::Testing::SelectedPathsFileObservation* observation
#endif
                                                        ) noexcept
{
    Common::Paths::UniqueSiblingFileOptions createOptions{};
    createOptions.prefix             = kSelectedPathsManifestPrefix;
    createOptions.suffix             = kSelectedPathsManifestSuffix;
    createOptions.desiredAccess      = GENERIC_WRITE | FILE_READ_ATTRIBUTES | DELETE;
    createOptions.shareMode          = 0u;
    createOptions.flagsAndAttributes = FILE_ATTRIBUTE_TEMPORARY | FILE_ATTRIBUTE_NOT_CONTENT_INDEXED | FILE_FLAG_SEQUENTIAL_SCAN;

#if defined(ENABLE_TESTS)
    if (testOptions != nullptr)
    {
        createOptions.maximumAttempts = testOptions->maximumCreateAttempts;
        const auto makeCandidateToken = [&](size_t attempt, std::wstring& candidateToken) noexcept
        {
            if (observation != nullptr)
            {
                observation->createAttempts = attempt + 1u;
            }
            candidateToken = std::format(L"test-{:04}", attempt);
            if (attempt >= testOptions->injectedCreateCollisions)
            {
                return S_OK;
            }

            std::filesystem::path collisionPath;
            const HRESULT collisionHr = CreateSelectedPathsCollisionFixture(rootForIo, candidateToken, collisionPath);
            if (SUCCEEDED(collisionHr) && observation != nullptr && observation->collidedPath.empty())
            {
                observation->collidedPath = std::move(collisionPath);
            }
            return collisionHr;
        };
        return Common::Paths::CreateUniqueFileInDirectory(rootForIo, createOptions, makeCandidateToken, filePath, file);
    }
#endif

    return Common::Paths::CreateUniqueFileInDirectory(rootForIo, createOptions, filePath, file);
}

[[nodiscard]] HRESULT CreateSelectedPathsFile(std::span<const std::filesystem::path> selectedPaths,
                                               FileActionLauncher::LaunchPlan::SelectedPathsFileLease& outLease
#if defined(ENABLE_TESTS)
                                               ,
                                               const FileActionLauncher::Testing::SelectedPathsFileOptions* testOptions
#endif
                                               ) noexcept
{
    outLease.Reset();

#if defined(ENABLE_TESTS)
    FileActionLauncher::Testing::SelectedPathsFileObservation* observation = testOptions != nullptr ? testOptions->observation : nullptr;
    if (observation != nullptr)
    {
        *observation = {};
    }
#endif

    Debug::Perf::Scope buildPerf(L"fileaction.selected_paths_file.build_us");
    buildPerf.SetValue0(static_cast<uint64_t>(selectedPaths.size()));

    SelectedPathsValidationSummary summary{};
    const auto validationStartedAt = std::chrono::steady_clock::now();
    const HRESULT validationHr     = ValidateSelectedPaths(selectedPaths, summary);
    Debug::Perf::EmitDurationUs(
        L"fileaction.selected_paths_file.validate_us", Debug::Perf::ElapsedUs(validationStartedAt), summary.recordCount, summary.pathBytes, validationHr);
#if defined(ENABLE_TESTS)
    if (observation != nullptr)
    {
        observation->recordCount = summary.recordCount;
        observation->pathBytes   = summary.pathBytes;
    }
#endif
    if (FAILED(validationHr))
    {
        buildPerf.SetHr(validationHr);
        return validationHr;
    }
    buildPerf.SetValue1(summary.totalBytes);

    std::wstring rootForIo;
    wil::unique_hfile rootHandle;
    FILE_ID_INFO rootIdentity{};
    if (const HRESULT hr = ResolveSelectedPathsManifestRoot(rootForIo,
                                                            rootHandle,
                                                            rootIdentity
#if defined(ENABLE_TESTS)
                                                            ,
                                                            testOptions
#endif
                                                            );
        FAILED(hr))
    {
        buildPerf.SetHr(hr);
        return hr;
    }

    std::wstring filePath;
    wil::unique_hfile file;
    const auto createStartedAt = std::chrono::steady_clock::now();
    const HRESULT createHr = CreateSelectedPathsManifestHandle(rootForIo,
                                                               filePath,
                                                               file
#if defined(ENABLE_TESTS)
                                                               ,
                                                               testOptions,
                                                               observation
#endif
    );
    Debug::Perf::EmitDurationUs(L"fileaction.selected_paths_file.create_us",
                                Debug::Perf::ElapsedUs(createStartedAt),
#if defined(ENABLE_TESTS)
                                observation != nullptr ? observation->createAttempts : 0u,
#else
                                0u,
#endif
                                0u,
                                createHr);
    if (FAILED(createHr))
    {
        buildPerf.SetHr(createHr);
        return createHr;
    }

    bool keepCompletedFile = false;
    const auto deleteIncompleteFile = wil::scope_exit([&]() noexcept
    {
        if (file && ! keepCompletedFile)
        {
            static_cast<void>(MarkFileForDeletion(file.get()));
        }
    });
    Debug::Perf::Scope writePerf(L"fileaction.selected_paths_file.write_us");
    writePerf.SetValue0(summary.recordCount);
    writePerf.SetValue1(summary.totalBytes);

    constexpr uint64_t kMaximumSelectedPathsWriteBufferBytes = 64u * 1024u;
    const size_t writeBufferSize = static_cast<size_t>((std::min)(summary.totalBytes, kMaximumSelectedPathsWriteBufferBytes));
    std::vector<std::byte> writeBuffer(writeBufferSize);
    size_t bufferedBytes = 0u;
    uint64_t writeCallCount = 0u;

    const auto flushBufferedBytes = [&]() noexcept -> HRESULT
    {
        if (bufferedBytes == 0u)
        {
            return S_OK;
        }
        const HRESULT hr = Common::HandleIo::WriteAll(file.get(), writeBuffer.data(), bufferedBytes);
        if (SUCCEEDED(hr))
        {
            ++writeCallCount;
            bufferedBytes = 0u;
        }
        return hr;
    };

    const auto appendBytes = [&](std::span<const std::byte> bytes) noexcept -> HRESULT
    {
        while (! bytes.empty())
        {
            if (bufferedBytes == writeBuffer.size())
            {
                if (const HRESULT hr = flushBufferedBytes(); FAILED(hr))
                {
                    return hr;
                }
            }

            const size_t copyBytes = (std::min)(bytes.size(), writeBuffer.size() - bufferedBytes);
            std::copy_n(bytes.begin(), copyBytes, writeBuffer.begin() + static_cast<std::ptrdiff_t>(bufferedBytes));
            bufferedBytes += copyBytes;
            bytes = bytes.subspan(copyBytes);
        }
        return S_OK;
    };

    constexpr std::array<std::byte, 2u> kUtf16LeBom{{std::byte{0xFFu}, std::byte{0xFEu}}};
    constexpr std::array<wchar_t, 2u> kRecordTerminator{{L'\r', L'\n'}};
    if (const HRESULT hr = appendBytes(kUtf16LeBom); FAILED(hr))
    {
        writePerf.SetHr(hr);
        buildPerf.SetHr(hr);
        return hr;
    }

    for (const std::filesystem::path& selectedPath : selectedPaths)
    {
        if (selectedPath.empty())
        {
            continue;
        }
        const std::wstring& record = selectedPath.native();
        if (const HRESULT hr = appendBytes(std::as_bytes(std::span(record.data(), record.size()))); FAILED(hr))
        {
            writePerf.SetHr(hr);
            buildPerf.SetHr(hr);
            return hr;
        }
        if (const HRESULT hr = appendBytes(std::as_bytes(std::span{kRecordTerminator})); FAILED(hr))
        {
            writePerf.SetHr(hr);
            buildPerf.SetHr(hr);
            return hr;
        }
    }

    if (const HRESULT hr = flushBufferedBytes(); FAILED(hr))
    {
        writePerf.SetHr(hr);
        buildPerf.SetHr(hr);
        return hr;
    }

    if (FlushFileBuffers(file.get()) == FALSE)
    {
        const DWORD error = GetLastError();
        const HRESULT hr  = HRESULT_FROM_WIN32(error == ERROR_SUCCESS ? ERROR_WRITE_FAULT : error);
        writePerf.SetHr(hr);
        buildPerf.SetHr(hr);
        return hr;
    }

    FILE_ID_INFO fileIdentity{};
    if (const HRESULT hr = QueryFileIdentity(file.get(), fileIdentity); FAILED(hr))
    {
        writePerf.SetHr(hr);
        buildPerf.SetHr(hr);
        return hr;
    }

    const uint64_t peakSerializationBytes = static_cast<uint64_t>(writeBuffer.size());
    Debug::Perf::EmitValue(L"fileaction.selected_paths_file.path_bytes", summary.pathBytes);
    Debug::Perf::EmitValue(L"fileaction.selected_paths_file.write_calls", writeCallCount);
    Debug::Perf::EmitValue(L"fileaction.selected_paths_file.peak_additional_bytes", peakSerializationBytes);
    Debug::Perf::EmitValue(L"fileaction.selected_paths_file.selection_copy_bytes", 0u);
#if defined(ENABLE_TESTS)
    if (observation != nullptr)
    {
        observation->writeCallCount      = writeCallCount;
        observation->peakAdditionalBytes = peakSerializationBytes;
        observation->selectionCopyBytes  = 0u;
        observation->usedAggregatePayload = false;
    }
#endif

    FileActionLauncher::LaunchPlan::SelectedPathsFileLease lease(
        std::filesystem::path(filePath), std::filesystem::path(rootForIo), fileIdentity, rootIdentity);
    keepCompletedFile = true;
    file.reset();
    rootHandle.reset();
    outLease = std::move(lease);
    return S_OK;
}

struct DeferredCleanupContext
{
    DeferredCleanupContext()                                         = default;
    DeferredCleanupContext(const DeferredCleanupContext&)            = delete;
    DeferredCleanupContext& operator=(const DeferredCleanupContext&) = delete;
    DeferredCleanupContext(DeferredCleanupContext&&)                 = delete;
    DeferredCleanupContext& operator=(DeferredCleanupContext&&)      = delete;

    wil::unique_threadpool_wait_nowait wait;
    wil::unique_handle process;
    wil::unique_handle fallbackEvent;
    FileActionLauncher::LaunchPlan::SelectedPathsFileLease lease;
};

void CALLBACK DeferredCleanupCallback(PTP_CALLBACK_INSTANCE instance, void* rawContext, PTP_WAIT /*wait*/, TP_WAIT_RESULT /*waitResult*/) noexcept
{
    std::unique_ptr<DeferredCleanupContext> context(static_cast<DeferredCleanupContext*>(rawContext));
    if (! context)
    {
        return;
    }

    DisassociateCurrentThreadFromCallback(instance);
    context->wait.reset();
    context->process.reset();
    context->fallbackEvent.reset();
}

enum class InjectedLaunchFault : uint8_t
{
    None,
    DeferredCleanupAllocation,
    DeferredCleanupWaitCreation,
    ShellExecute,
    NullProcessHandle,
    WaitFailure,
    ExitCodeQuery,
};

#if defined(ENABLE_TESTS)
std::atomic<InjectedLaunchFault> g_nextInjectedLaunchFault{InjectedLaunchFault::None};
#endif

[[nodiscard]] InjectedLaunchFault TakeInjectedLaunchFault() noexcept
{
#if defined(ENABLE_TESTS)
    return g_nextInjectedLaunchFault.exchange(InjectedLaunchFault::None, std::memory_order_relaxed);
#else
    return InjectedLaunchFault::None;
#endif
}

[[nodiscard]] HRESULT PrepareDeferredCleanupContext(FileActionLauncher::LaunchPlan::SelectedPathsFileLease& lease,
                                                     InjectedLaunchFault fault,
                                                     std::unique_ptr<DeferredCleanupContext>& context) noexcept
{
    context.reset();
    if (! lease.HasValue())
    {
        return S_OK;
    }
    if (fault == InjectedLaunchFault::DeferredCleanupAllocation)
    {
        return E_OUTOFMEMORY;
    }

    std::unique_ptr<DeferredCleanupContext> prepared(new (std::nothrow) DeferredCleanupContext());
    if (! prepared)
    {
        return E_OUTOFMEMORY;
    }
    if (fault != InjectedLaunchFault::DeferredCleanupWaitCreation)
    {
        prepared->wait.reset(CreateThreadpoolWait(DeferredCleanupCallback, prepared.get(), nullptr));
    }
    if (! prepared->wait)
    {
        if (fault == InjectedLaunchFault::DeferredCleanupWaitCreation)
        {
            return HRESULT_FROM_WIN32(ERROR_NOT_ENOUGH_MEMORY);
        }
        const DWORD error = GetLastError();
        return HRESULT_FROM_WIN32(error == ERROR_SUCCESS ? ERROR_NOT_ENOUGH_MEMORY : error);
    }

    prepared->fallbackEvent.reset(CreateEventW(nullptr, TRUE, FALSE, nullptr));
    if (! prepared->fallbackEvent)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    prepared->lease = std::move(lease);
    context         = std::move(prepared);
    return S_OK;
}

void ScheduleDeferredCleanup(std::unique_ptr<DeferredCleanupContext>& context, wil::unique_handle process, DWORD requestedRetentionMs) noexcept
{
    if (! context)
    {
        return;
    }

    constexpr DWORD kMaximumRetentionMs = 10u * 60u * 1000u;
    const DWORD retentionMs = std::clamp<DWORD>(requestedRetentionMs, DWORD{1u}, kMaximumRetentionMs);
    context->process        = std::move(process);
    const HANDLE target     = context->process ? context->process.get() : context->fallbackEvent.get();

    LARGE_INTEGER relativeTimeout{};
    relativeTimeout.QuadPart = -static_cast<LONGLONG>(retentionMs) * 10'000ll;
    FILETIME timeout{};
    timeout.dwLowDateTime  = relativeTimeout.LowPart;
    timeout.dwHighDateTime = static_cast<DWORD>(relativeTimeout.HighPart);
    SetThreadpoolWait(context->wait.get(), target, &timeout);
    static_cast<void>(context.release());
}

} // namespace

namespace FileActionLauncher
{
LaunchPlan::SelectedPathsFileLease::SelectedPathsFileLease(std::filesystem::path path,
                                                          std::filesystem::path rootPath,
                                                          const FILE_ID_INFO& fileIdentity,
                                                          const FILE_ID_INFO& rootIdentity) noexcept
    : _path(std::move(path)),
      _rootPath(std::move(rootPath)),
      _fileIdentity(fileIdentity),
      _rootIdentity(rootIdentity),
      _hasIdentity(true)
{
}

LaunchPlan::SelectedPathsFileLease::~SelectedPathsFileLease() noexcept
{
    Reset();
}

LaunchPlan::SelectedPathsFileLease::SelectedPathsFileLease(SelectedPathsFileLease&& other) noexcept
    : _path(std::move(other._path)),
      _rootPath(std::move(other._rootPath)),
      _fileIdentity(other._fileIdentity),
      _rootIdentity(other._rootIdentity),
      _hasIdentity(other._hasIdentity)
{
    other._path.clear();
    other._rootPath.clear();
    other._fileIdentity = {};
    other._rootIdentity = {};
    other._hasIdentity  = false;
}

LaunchPlan::SelectedPathsFileLease& LaunchPlan::SelectedPathsFileLease::operator=(SelectedPathsFileLease&& other) noexcept
{
    if (this != &other)
    {
        Reset();
        _path         = std::move(other._path);
        _rootPath     = std::move(other._rootPath);
        _fileIdentity = other._fileIdentity;
        _rootIdentity = other._rootIdentity;
        _hasIdentity  = other._hasIdentity;
        other._path.clear();
        other._rootPath.clear();
        other._fileIdentity = {};
        other._rootIdentity = {};
        other._hasIdentity  = false;
    }
    return *this;
}

const std::filesystem::path& LaunchPlan::SelectedPathsFileLease::Path() const noexcept
{
    return _path;
}

bool LaunchPlan::SelectedPathsFileLease::HasValue() const noexcept
{
    return _hasIdentity && ! _path.empty() && ! _rootPath.empty();
}

void LaunchPlan::SelectedPathsFileLease::Reset() noexcept
{
    if (! HasValue())
    {
        _path.clear();
        _rootPath.clear();
        _fileIdentity = {};
        _rootIdentity = {};
        _hasIdentity  = false;
        return;
    }
    CleanupSelectedPathsFile(_path, _rootPath, _fileIdentity, _rootIdentity);
    _path.clear();
    _rootPath.clear();
    _fileIdentity = {};
    _rootIdentity = {};
    _hasIdentity  = false;
}

#if defined(ENABLE_TESTS)
namespace Testing
{
void SetNextExternalLaunchFault(ExternalLaunchFault fault) noexcept
{
    g_nextInjectedLaunchFault.store(static_cast<InjectedLaunchFault>(fault), std::memory_order_relaxed);
}

HRESULT RunSelectedPathsRecoveryPassForTest(const std::filesystem::path& root,
                                            uint64_t minimumAgeMs,
                                            size_t maximumEntries,
                                            size_t maximumDeletes,
                                            uint64_t maximumElapsedMs,
                                            std::wstring& cursor,
                                            SelectedPathsRecoveryObservation& observation) noexcept
{
    observation = {};
    std::wstring rootForIo;
    if (const HRESULT hr = BuildSelectedPathsIoPath(root, 96u, rootForIo); FAILED(hr))
    {
        return hr;
    }
    const SelectedPathsRecoveryLimits limits{minimumAgeMs, maximumEntries, maximumDeletes, maximumElapsedMs};
    ::SelectedPathsRecoveryObservation internal{};
    const HRESULT hr = RunSelectedPathsRecoveryPass(rootForIo, limits, cursor, internal);
    observation.inspected            = internal.inspected;
    observation.deleted              = internal.deleted;
    observation.skipped              = internal.skipped;
    observation.errors               = internal.errors;
    observation.elapsedUs            = internal.elapsedUs;
    observation.stoppedByEntryBound  = internal.stoppedByEntryBound;
    observation.stoppedByDeleteBound = internal.stoppedByDeleteBound;
    observation.stoppedByTimeBound   = internal.stoppedByTimeBound;
    observation.resetMissingCursor   = internal.resetMissingCursor;
    return hr;
}
} // namespace Testing
#endif

HRESULT ExpandMacros(std::wstring_view templateText, const MacroContext& context, std::wstring& out) noexcept
{
    return ExpandMacrosInternal(templateText, context, false, out);
}

bool TemplateContainsSupportedMacro(const std::wstring_view templateText) noexcept
{
    for (size_t index = 0; index < templateText.size();)
    {
        const wchar_t ch = templateText[index];
        if (ch == L'{')
        {
            if (index + 1u < templateText.size() && templateText[index + 1u] == L'{')
            {
                index += 2u;
                continue;
            }

            const size_t close = templateText.find(L'}', index + 1u);
            if (close == std::wstring_view::npos)
            {
                ++index;
                continue;
            }

            if (IsSupportedMacroName(templateText.substr(index + 1u, close - index - 1u)))
            {
                return true;
            }

            index = close + 1u;
            continue;
        }

        if (ch == L'}' && index + 1u < templateText.size() && templateText[index + 1u] == L'}')
        {
            index += 2u;
            continue;
        }

        ++index;
    }

    return false;
}

HRESULT GetExternalActionPathReferences(const Common::Settings::FileActionDefinition& action,
                                        ExternalActionPathReferences& references) noexcept
{
    return ActionPathReferences(action, references);
}

HRESULT ExternalActionUsesSelectedPathsFile(const Common::Settings::FileActionDefinition& action,
                                            bool& usesSelectedPathsFile) noexcept
{
    return ActionReferencesSelectedPathsFile(action, usesSelectedPathsFile);
}

HRESULT BuildExternalLaunchPlan(const Common::Settings::FileActionDefinition& action, const MacroContext& context, LaunchPlan& out) noexcept
{
    out = LaunchPlan{};

    if (action.kind != Common::Settings::FileActionKind::ExternalProgram || action.executablePath.empty())
    {
        return E_INVALIDARG;
    }

    MacroContext effectiveContext{};
    effectiveContext.itemPath          = context.itemPath;
    effectiveContext.currentDirectory  = context.currentDirectory;
    effectiveContext.oppositePanePath  = context.oppositePanePath;
    effectiveContext.selectedPathsFile = context.selectedPathsFile;
    effectiveContext.computerName      = context.computerName;
#if defined(ENABLE_TESTS)
    effectiveContext.selectedPathsFileOptions = context.selectedPathsFileOptions;
#endif

    bool referencesSelectedPathsFile = false;
    if (const HRESULT hr = ActionReferencesSelectedPathsFile(action, referencesSelectedPathsFile); FAILED(hr))
    {
        return hr;
    }
    if (referencesSelectedPathsFile && effectiveContext.selectedPathsFile.empty())
    {
        std::array<std::filesystem::path, 1u> focusedPathFallback{};
        std::span<const std::filesystem::path> selectedPaths = context.selectedPaths;
        if (selectedPaths.empty() && ! context.itemPath.empty())
        {
            focusedPathFallback[0] = context.itemPath;
            selectedPaths          = focusedPathFallback;
        }

        if (const HRESULT hr = CreateSelectedPathsFile(selectedPaths,
                                                       out.selectedPathsFileLease
#if defined(ENABLE_TESTS)
                                                       ,
                                                       effectiveContext.selectedPathsFileOptions
#endif
                                                       );
            FAILED(hr))
        {
            return hr;
        }

        effectiveContext.selectedPathsFile = out.selectedPathsFileLease.Path();
    }

    if (const HRESULT hr = ExpandMacros(action.executablePath, effectiveContext, out.executablePath); FAILED(hr))
    {
        out = LaunchPlan{};
        return hr;
    }
    if (out.executablePath.empty())
    {
        out = LaunchPlan{};
        return E_INVALIDARG;
    }
    if (! Common::Paths::IsExplicitAbsoluteExecutablePath(out.executablePath))
    {
        out = LaunchPlan{};
        return HRESULT_FROM_WIN32(ERROR_BAD_PATHNAME);
    }

    if (const HRESULT hr = ExpandArgumentMacros(action.arguments, effectiveContext, out.arguments); FAILED(hr))
    {
        out = LaunchPlan{};
        return hr;
    }

    if (! action.workingDirectory.empty())
    {
        if (const HRESULT hr = ExpandMacros(action.workingDirectory, effectiveContext, out.workingDirectory); FAILED(hr))
        {
            out = LaunchPlan{};
            return hr;
        }
    }
    else
    {
        out.workingDirectory = CurrentDirectoryForContext(effectiveContext).wstring();
    }

    return S_OK;
}

HRESULT LaunchExternalPlan(LaunchPlan plan, const LaunchOptions& options, LaunchResult* result) noexcept
{
    Debug::Perf::Scope perf(L"fileaction.external.launch_us");
    perf.SetDetail(options.waitForExit ? L"wait" : L"start");
    perf.SetValue0(static_cast<uint64_t>(plan.arguments.size()));
    perf.SetValue1(options.waitForExit ? options.waitTimeoutMs : 0u);

    if (result != nullptr)
    {
        *result = LaunchResult{};
    }

    if (! Common::Paths::IsExplicitAbsoluteExecutablePath(plan.executablePath))
    {
        const HRESULT hr = HRESULT_FROM_WIN32(ERROR_BAD_PATHNAME);
        perf.SetHr(hr);
        return hr;
    }

    const InjectedLaunchFault injectedFault = TakeInjectedLaunchFault();
    std::unique_ptr<DeferredCleanupContext> cleanupContext;
    if (const HRESULT prepareHr = PrepareDeferredCleanupContext(plan.selectedPathsFileLease, injectedFault, cleanupContext); FAILED(prepareHr))
    {
        perf.SetHr(prepareHr);
        return prepareHr;
    }

    SHELLEXECUTEINFOW executeInfo{};
    executeInfo.cbSize       = sizeof(executeInfo);
    executeInfo.fMask        = SEE_MASK_NOCLOSEPROCESS | SEE_MASK_FLAG_NO_UI;
    executeInfo.hwnd         = options.ownerWindow;
    executeInfo.lpVerb       = L"open";
    executeInfo.lpFile       = plan.executablePath.c_str();
    executeInfo.lpParameters = plan.arguments.empty() ? nullptr : plan.arguments.c_str();
    executeInfo.lpDirectory  = plan.workingDirectory.empty() ? nullptr : plan.workingDirectory.c_str();
    executeInfo.nShow        = options.showCommand;

    if (injectedFault == InjectedLaunchFault::ShellExecute || ShellExecuteExW(&executeInfo) == FALSE)
    {
        const DWORD error = injectedFault == InjectedLaunchFault::ShellExecute
                                ? ERROR_FILE_NOT_FOUND
                                : Debug::ErrorWithLastError(L"FileActionLauncher: ShellExecuteExW failed for external action '{}'", plan.executablePath);
        const HRESULT hr  = HRESULT_FROM_WIN32(error == ERROR_SUCCESS ? ERROR_FILE_NOT_FOUND : error);
        perf.SetHr(hr);
        return hr;
    }

    wil::unique_handle processHandle(executeInfo.hProcess);
    if (injectedFault == InjectedLaunchFault::NullProcessHandle)
    {
        processHandle.reset();
    }
    if (processHandle)
    {
        if (result != nullptr)
        {
            result->processId = GetProcessId(processHandle.get());
        }
        if (result != nullptr && options.captureProcessHandle)
        {
            HANDLE duplicatedHandle = nullptr;
            if (DuplicateHandle(GetCurrentProcess(),
                                processHandle.get(),
                                GetCurrentProcess(),
                                &duplicatedHandle,
                                SYNCHRONIZE | PROCESS_QUERY_LIMITED_INFORMATION,
                                FALSE,
                                0) != FALSE)
            {
                result->processHandle.reset(duplicatedHandle);
            }
            else
            {
                Debug::Warning(L"FileActionLauncher: failed to duplicate external process handle for tracking.");
            }
        }
    }
    if (! options.waitForExit)
    {
        ScheduleDeferredCleanup(cleanupContext, std::move(processHandle), options.selectedPathsMaximumRetentionMs);
        return S_OK;
    }
    if (! processHandle)
    {
        const HRESULT hr = HRESULT_FROM_WIN32(ERROR_INVALID_HANDLE);
        perf.SetHr(hr);
        ScheduleDeferredCleanup(cleanupContext, {}, options.selectedPathsMaximumRetentionMs);
        return hr;
    }

    const DWORD waitResult = injectedFault == InjectedLaunchFault::WaitFailure ? WAIT_FAILED : WaitForSingleObject(processHandle.get(), options.waitTimeoutMs);
    if (waitResult == WAIT_TIMEOUT)
    {
        const HRESULT hr = HRESULT_FROM_WIN32(ERROR_TIMEOUT);
        perf.SetHr(hr);
        ScheduleDeferredCleanup(cleanupContext, std::move(processHandle), options.selectedPathsMaximumRetentionMs);
        return hr;
    }
    if (waitResult == WAIT_FAILED)
    {
        const HRESULT hr = HRESULT_FROM_WIN32(injectedFault == InjectedLaunchFault::WaitFailure ? ERROR_GEN_FAILURE : GetLastError());
        perf.SetHr(hr);
        ScheduleDeferredCleanup(cleanupContext, std::move(processHandle), options.selectedPathsMaximumRetentionMs);
        return hr;
    }
    if (waitResult != WAIT_OBJECT_0)
    {
        perf.SetHr(E_FAIL);
        ScheduleDeferredCleanup(cleanupContext, std::move(processHandle), options.selectedPathsMaximumRetentionMs);
        return E_FAIL;
    }

    DWORD exitCode = 0;
    if (injectedFault == InjectedLaunchFault::ExitCodeQuery || GetExitCodeProcess(processHandle.get(), &exitCode) == FALSE)
    {
        const HRESULT hr = HRESULT_FROM_WIN32(injectedFault == InjectedLaunchFault::ExitCodeQuery ? ERROR_GEN_FAILURE : GetLastError());
        perf.SetHr(hr);
        return hr;
    }

    if (result != nullptr)
    {
        result->exitCodeAvailable = true;
        result->exitCode          = exitCode;
    }

    return S_OK;
}
} // namespace FileActionLauncher
