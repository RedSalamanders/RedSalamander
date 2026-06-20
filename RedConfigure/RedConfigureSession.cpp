#include "RedConfigureSession.h"

#include "Helpers.h"
#include "Localization/RcParser.h"
#include "Localization/RcWriter.h"
#include "ThemeDefinitionIo.h"

#include <algorithm>
#include <cstdint>
#include <cstring>
#include <format>
#include <fstream>
#include <future>
#include <iterator>
#include <limits>
#include <optional>
#include <set>
#include <span>
#include <string>
#include <system_error>
#include <thread>
#include <unordered_map>
#include <vector>

namespace
{
namespace fs = std::filesystem;

[[nodiscard]] wchar_t ToLowerAscii(wchar_t ch) noexcept
{
    if (ch >= L'A' && ch <= L'Z')
    {
        return static_cast<wchar_t>(ch - L'A' + L'a');
    }
    return ch;
}

[[nodiscard]] bool ContainsIgnoreCase(std::wstring_view text, std::wstring_view needle) noexcept
{
    if (needle.empty())
    {
        return true;
    }
    if (needle.size() > text.size())
    {
        return false;
    }

    for (size_t index = 0u; index + needle.size() <= text.size(); ++index)
    {
        bool matched = true;
        for (size_t offset = 0u; offset < needle.size(); ++offset)
        {
            if (ToLowerAscii(text[index + offset]) != ToLowerAscii(needle[offset]))
            {
                matched = false;
                break;
            }
        }
        if (matched)
        {
            return true;
        }
    }
    return false;
}

[[nodiscard]] std::wstring_view TranslationColumnText(const RedConfigure::TranslationEntry& entry, RedConfigure::LocalizationViewColumn column) noexcept
{
    switch (column)
    {
        case RedConfigure::LocalizationViewColumn::Owner: return {};
        case RedConfigure::LocalizationViewColumn::Id: return entry.id;
        case RedConfigure::LocalizationViewColumn::Source: return entry.sourceText;
        case RedConfigure::LocalizationViewColumn::Target: return entry.targetText;
        case RedConfigure::LocalizationViewColumn::Status: return {};
        default: return {};
    }
}

[[nodiscard]] int CompareTextIgnoreCase(std::wstring_view lhs, std::wstring_view rhs) noexcept
{
    const int result = ::CompareStringOrdinal(lhs.data(),
                                              static_cast<int>((std::min)(lhs.size(), static_cast<size_t>((std::numeric_limits<int>::max)()))),
                                              rhs.data(),
                                              static_cast<int>((std::min)(rhs.size(), static_cast<size_t>((std::numeric_limits<int>::max)()))),
                                              TRUE);
    if (result == CSTR_LESS_THAN)
    {
        return -1;
    }
    if (result == CSTR_GREATER_THAN)
    {
        return 1;
    }
    return 0;
}

[[nodiscard]] bool TranslationMatchesViewFilters(const RedConfigure::TranslationEntry& entry, const RedConfigure::LocalizationViewOptions& options) noexcept
{
    if (! options.idFilterText.empty() && ! ContainsIgnoreCase(entry.id, options.idFilterText))
    {
        return false;
    }

    const bool okStatus = entry.validation.status == RedConfigure::Localization::PlaceholderStatus::Ok;
    if (options.statusFilter == RedConfigure::LocalizationStatusFilter::Ok && ! okStatus)
    {
        return false;
    }
    if (options.statusFilter == RedConfigure::LocalizationStatusFilter::Problems && okStatus)
    {
        return false;
    }

    return options.searchText.empty() || ContainsIgnoreCase(entry.id, options.searchText) || ContainsIgnoreCase(entry.sourceText, options.searchText) ||
           ContainsIgnoreCase(entry.targetText, options.searchText);
}

[[nodiscard]] HRESULT ReadBinaryFile(const fs::path& path, std::vector<uint8_t>& outBytes)
{
    outBytes.clear();
    std::ifstream input(path, std::ios::binary);
    if (! input)
    {
        return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
    }

    outBytes.assign(std::istreambuf_iterator<char>(input), std::istreambuf_iterator<char>());
    return input.bad() ? HRESULT_FROM_WIN32(ERROR_READ_FAULT) : S_OK;
}

[[nodiscard]] bool DecodeBytes(UINT codePage, const uint8_t* bytes, size_t byteCount, std::wstring& outText) noexcept
{
    outText.clear();
    if (byteCount == 0u)
    {
        return true;
    }
    if (byteCount > static_cast<size_t>(std::numeric_limits<int>::max()))
    {
        return false;
    }

    const DWORD flags  = (codePage == CP_UTF8) ? MB_ERR_INVALID_CHARS : 0u;
    const int required = ::MultiByteToWideChar(codePage, flags, reinterpret_cast<const char*>(bytes), static_cast<int>(byteCount), nullptr, 0);
    if (required <= 0)
    {
        return false;
    }

    std::wstring result(static_cast<size_t>(required), L'\0');
    const int written = ::MultiByteToWideChar(codePage, flags, reinterpret_cast<const char*>(bytes), static_cast<int>(byteCount), result.data(), required);
    if (written != required)
    {
        return false;
    }

    outText = std::move(result);
    return true;
}

[[nodiscard]] bool DecodeUtf16LeBytes(const uint8_t* bytes, size_t byteCount, std::wstring& outText) noexcept
{
    static_assert(sizeof(wchar_t) == 2u);
    outText.clear();
    if ((byteCount % sizeof(wchar_t)) != 0u)
    {
        return false;
    }

    outText.resize(byteCount / sizeof(wchar_t));
    if (! outText.empty())
    {
        std::memcpy(outText.data(), bytes, byteCount);
    }
    return true;
}

[[nodiscard]] bool DecodeUtf16BeBytes(const uint8_t* bytes, size_t byteCount, std::wstring& outText) noexcept
{
    static_assert(sizeof(wchar_t) == 2u);
    outText.clear();
    if ((byteCount % sizeof(wchar_t)) != 0u)
    {
        return false;
    }

    outText.reserve(byteCount / sizeof(wchar_t));
    for (size_t index = 0u; index + 1u < byteCount; index += 2u)
    {
        const auto codeUnit = static_cast<wchar_t>((static_cast<uint16_t>(bytes[index]) << 8u) | static_cast<uint16_t>(bytes[index + 1u]));
        outText.push_back(codeUnit);
    }
    return true;
}

enum class BomlessUtf16Guess : uint8_t
{
    None,
    LittleEndian,
    BigEndian,
};

[[nodiscard]] BomlessUtf16Guess GuessBomlessUtf16Encoding(std::span<const uint8_t> bytes) noexcept
{
    if (bytes.size() < 8u || (bytes.size() % 2u) != 0u)
    {
        return BomlessUtf16Guess::None;
    }

    const size_t sampleBytes = (std::min)(bytes.size(), static_cast<size_t>(4096u));
    size_t evenZeroCount     = 0u;
    size_t oddZeroCount      = 0u;
    for (size_t index = 0u; index + 1u < sampleBytes; index += 2u)
    {
        if (bytes[index] == 0u)
        {
            ++evenZeroCount;
        }
        if (bytes[index + 1u] == 0u)
        {
            ++oddZeroCount;
        }
    }

    const size_t sampledPairs = sampleBytes / 2u;
    if (oddZeroCount >= sampledPairs / 4u && oddZeroCount >= (evenZeroCount * 4u))
    {
        return BomlessUtf16Guess::LittleEndian;
    }
    if (evenZeroCount >= sampledPairs / 4u && evenZeroCount >= (oddZeroCount * 4u))
    {
        return BomlessUtf16Guess::BigEndian;
    }
    return BomlessUtf16Guess::None;
}

[[nodiscard]] HRESULT ReadTextFile(const fs::path& path, std::wstring& outText)
{
    outText.clear();

    std::vector<uint8_t> bytes;
    if (const HRESULT hr = ReadBinaryFile(path, bytes); FAILED(hr))
    {
        return hr;
    }

    if (bytes.size() >= 2u && bytes[0] == 0xFFu && bytes[1] == 0xFEu)
    {
        if (! DecodeUtf16LeBytes(bytes.data() + 2u, bytes.size() - 2u, outText))
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }
        return S_OK;
    }
    if (bytes.size() >= 2u && bytes[0] == 0xFEu && bytes[1] == 0xFFu)
    {
        if (! DecodeUtf16BeBytes(bytes.data() + 2u, bytes.size() - 2u, outText))
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }
        return S_OK;
    }

    switch (GuessBomlessUtf16Encoding(bytes))
    {
        case BomlessUtf16Guess::LittleEndian:
            if (DecodeUtf16LeBytes(bytes.data(), bytes.size(), outText))
            {
                return S_OK;
            }
            break;
        case BomlessUtf16Guess::BigEndian:
            if (DecodeUtf16BeBytes(bytes.data(), bytes.size(), outText))
            {
                return S_OK;
            }
            break;
        case BomlessUtf16Guess::None:
        default: break;
    }

    const uint8_t* textBytes = bytes.data();
    size_t textSize          = bytes.size();
    if (bytes.size() >= 3u && bytes[0] == 0xEFu && bytes[1] == 0xBBu && bytes[2] == 0xBFu)
    {
        textBytes += 3u;
        textSize -= 3u;
    }

    if (DecodeBytes(CP_UTF8, textBytes, textSize, outText))
    {
        return S_OK;
    }
    if (DecodeBytes(CP_ACP, textBytes, textSize, outText))
    {
        Debug::Warning(L"RedConfigure: '{}' is not valid UTF-8; decoded using the active Windows code page fallback.", path.wstring());
        return S_OK;
    }

    return HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);
}

[[nodiscard]] HRESULT EnsureParentDirectory(const fs::path& path)
{
    const fs::path parent = path.parent_path();
    if (parent.empty())
    {
        return S_OK;
    }

    std::error_code ec;
    fs::create_directories(parent, ec);
    return ec ? HRESULT_FROM_WIN32(static_cast<DWORD>(ec.value())) : S_OK;
}

[[nodiscard]] HRESULT WriteUtf16LeTextFile(const fs::path& path, std::wstring_view text)
{
    if (const HRESULT hr = EnsureParentDirectory(path); FAILED(hr))
    {
        return hr;
    }

    std::ofstream output(path, std::ios::binary);
    if (! output)
    {
        return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
    }

    constexpr uint8_t bom[] = {0xFFu, 0xFEu};
    output.write(reinterpret_cast<const char*>(bom), sizeof(bom));
    if (! text.empty())
    {
        output.write(reinterpret_cast<const char*>(text.data()), static_cast<std::streamsize>(text.size() * sizeof(wchar_t)));
    }
    return output.good() ? S_OK : HRESULT_FROM_WIN32(ERROR_WRITE_FAULT);
}

[[nodiscard]] HRESULT WriteBinaryFile(const fs::path& path, std::string_view bytes)
{
    if (const HRESULT hr = EnsureParentDirectory(path); FAILED(hr))
    {
        return hr;
    }

    std::ofstream output(path, std::ios::binary);
    if (! output)
    {
        return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
    }

    if (! bytes.empty())
    {
        output.write(bytes.data(), static_cast<std::streamsize>(bytes.size()));
    }
    return output.good() ? S_OK : HRESULT_FROM_WIN32(ERROR_WRITE_FAULT);
}

[[nodiscard]] std::optional<fs::path> FindSatelliteForCulture(const RedConfigure::Workspace::ResourceOwner& owner, std::wstring_view cultureName)
{
    for (const fs::path& path : owner.satelliteResourcePaths)
    {
        const fs::path cultureFolder = path.parent_path();
        if (! cultureFolder.empty() && CompareTextIgnoreCase(cultureFolder.filename().wstring(), cultureName) == 0)
        {
            return path;
        }
    }
    return std::nullopt;
}

[[nodiscard]] std::optional<std::wstring> SatelliteCultureName(const fs::path& path)
{
    const fs::path cultureFolder = path.parent_path();
    if (cultureFolder.empty() || cultureFolder.parent_path().filename() != L"Lang")
    {
        return std::nullopt;
    }
    return cultureFolder.filename().wstring();
}

[[nodiscard]] std::vector<std::wstring> DiscoverLocalizationReviewCultures(const RedConfigure::Workspace::WorkspaceScanResult& workspace)
{
    std::set<std::wstring> cultures;
    for (const RedConfigure::Workspace::ResourceOwner& owner : workspace.resourceOwners)
    {
        for (const fs::path& path : owner.satelliteResourcePaths)
        {
            if (const std::optional<std::wstring> culture = SatelliteCultureName(path))
            {
                // English is the embedded source language; never offer en-US as a review
                // target even when a Lang\en-US satellite exists on disk.
                if (CompareTextIgnoreCase(culture.value(), L"en-US") != 0)
                {
                    cultures.insert(culture.value());
                }
            }
        }
    }
    return std::vector<std::wstring>(cultures.begin(), cultures.end());
}

using StringTableById = std::unordered_map<std::wstring, std::wstring>;

[[nodiscard]] StringTableById BuildStringTableById(std::span<const RedConfigure::Localization::RcStringEntry> strings)
{
    StringTableById byId;
    byId.reserve(strings.size());
    for (const RedConfigure::Localization::RcStringEntry& entry : strings)
    {
        byId.try_emplace(entry.id, entry.text);
    }
    return byId;
}

[[nodiscard]] bool ContainsNameIgnoreCase(std::span<const std::wstring> names, std::wstring_view name) noexcept
{
    for (const std::wstring& candidate : names)
    {
        if (CompareTextIgnoreCase(candidate, name) == 0)
        {
            return true;
        }
    }
    return false;
}

void AddWorkspaceError(std::vector<std::wstring>& errors, std::wstring error)
{
    if (std::find(errors.begin(), errors.end(), error) == errors.end())
    {
        errors.push_back(std::move(error));
    }
}

void AddWorkspaceFileError(std::vector<std::wstring>& errors, const fs::path& path, std::wstring_view operation, HRESULT hr)
{
    AddWorkspaceError(errors,
                      std::format(L"{} failed for '{}'. hr=0x{:08X}",
                                  operation,
                                  path.wstring(),
                                  static_cast<unsigned long>(hr)));
}

void AddWorkspaceParseErrors(std::vector<std::wstring>& errors, const fs::path& path, std::span<const std::wstring> parseErrors)
{
    for (const std::wstring& error : parseErrors)
    {
        AddWorkspaceError(errors, std::format(L"{}: {}", path.wstring(), error));
    }
}

[[nodiscard]] bool ReadAndParseRcStringTables(const fs::path& path,
                                              std::wstring_view readOperation,
                                              std::wstring_view parseOperation,
                                              RedConfigure::Localization::RcParseResult& outResult,
                                              std::vector<std::wstring>& errors)
{
    outResult = {};

    std::wstring text;
    if (const HRESULT hr = ReadTextFile(path, text); FAILED(hr))
    {
        AddWorkspaceFileError(errors, path, readOperation, hr);
        return false;
    }

    if (const HRESULT hr = RedConfigure::Localization::ParseRcStringTables(text, outResult); FAILED(hr))
    {
        AddWorkspaceFileError(errors, path, parseOperation, hr);
        outResult = {};
        return false;
    }

    AddWorkspaceParseErrors(errors, path, outResult.errors);
    return true;
}

struct LocalizationReviewOwnerLoadResult
{
    size_t ownerIndex = 0u;
    std::vector<RedConfigure::LocalizationReviewRow> rows;
    std::vector<std::wstring> errors;
};

[[nodiscard]] LocalizationReviewOwnerLoadResult LoadLocalizationReviewOwner(const RedConfigure::Workspace::ResourceOwner& owner,
                                                                            std::span<const std::wstring> cultures,
                                                                            size_t ownerIndex)
{
    Debug::Perf::Scope ownerPerf(L"redconfigure.localization_review.owner_load_us");
    ownerPerf.SetDetail(owner.name);
    ownerPerf.SetValue1(static_cast<uint64_t>(cultures.size()));

    LocalizationReviewOwnerLoadResult result;
    result.ownerIndex = ownerIndex;

    RedConfigure::Localization::RcParseResult source;
    if (! ReadAndParseRcStringTables(owner.embeddedResourcePath,
                                     L"Read localization source",
                                     L"Parse localization source",
                                     source,
                                     result.errors))
    {
        ownerPerf.SetHr(HRESULT_FROM_WIN32(ERROR_INVALID_DATA));
        return result;
    }

    std::vector<StringTableById> targetStringsByCulture;
    targetStringsByCulture.reserve(cultures.size());
    for (const std::wstring& culture : cultures)
    {
        StringTableById targetStrings;
        if (const std::optional<fs::path> satellitePath = FindSatelliteForCulture(owner, culture))
        {
            RedConfigure::Localization::RcParseResult target;
            if (ReadAndParseRcStringTables(satellitePath.value(),
                                           L"Read localization target",
                                           L"Parse localization target",
                                           target,
                                           result.errors))
            {
                targetStrings = BuildStringTableById(target.strings);
            }
        }
        targetStringsByCulture.push_back(std::move(targetStrings));
    }

    result.rows.reserve(source.strings.size());
    for (const RedConfigure::Localization::RcStringEntry& sourceString : source.strings)
    {
        RedConfigure::LocalizationReviewRow row;
        row.ownerName  = owner.name;
        row.id         = sourceString.id;
        row.sourceText = sourceString.text;
        row.targets.reserve(cultures.size());

        for (size_t cultureIndex = 0u; cultureIndex < cultures.size(); ++cultureIndex)
        {
            const StringTableById& targetStrings = targetStringsByCulture[cultureIndex];
            const auto targetIt                  = targetStrings.find(sourceString.id);
            RedConfigure::LocalizationTargetCell cell;
            cell.cultureName            = cultures[cultureIndex];
            cell.hasExistingTranslation = targetIt != targetStrings.end();
            cell.targetText             = cell.hasExistingTranslation ? targetIt->second : sourceString.text;
            cell.validation             = RedConfigure::Localization::ValidatePlaceholders(sourceString.text, cell.targetText);
            row.targets.push_back(std::move(cell));
        }

        result.rows.push_back(std::move(row));
    }

    ownerPerf.SetValue0(static_cast<uint64_t>(result.rows.size()));
    return result;
}

[[nodiscard]] std::vector<LocalizationReviewOwnerLoadResult> LoadLocalizationReviewOwnerRange(
    std::span<const RedConfigure::Workspace::ResourceOwner> owners,
    std::span<const std::wstring> cultures,
    size_t beginIndex,
    size_t endIndex)
{
    std::vector<LocalizationReviewOwnerLoadResult> results;
    results.reserve(endIndex - beginIndex);
    for (size_t ownerIndex = beginIndex; ownerIndex < endIndex; ++ownerIndex)
    {
        results.push_back(LoadLocalizationReviewOwner(owners[ownerIndex], cultures, ownerIndex));
    }
    return results;
}

void AppendLocalizationReviewLoadResult(LocalizationReviewOwnerLoadResult& result,
                                        std::vector<RedConfigure::LocalizationReviewRow>& rows,
                                        std::vector<std::wstring>& errors)
{
    for (std::wstring& error : result.errors)
    {
        AddWorkspaceError(errors, std::move(error));
    }

    rows.insert(rows.end(), std::make_move_iterator(result.rows.begin()), std::make_move_iterator(result.rows.end()));
}

[[nodiscard]] size_t LocalizationReviewWorkerCount(size_t ownerCount) noexcept
{
    if (ownerCount <= 1u)
    {
        return 1u;
    }

    const unsigned hardwareCount = std::thread::hardware_concurrency();
    const size_t desiredCount    = hardwareCount == 0u ? 2u : static_cast<size_t>(hardwareCount);
    return (std::min)({ownerCount, desiredCount, static_cast<size_t>(8u)});
}

[[nodiscard]] const RedConfigure::LocalizationTargetCell* FindReviewTargetCell(const RedConfigure::LocalizationReviewRow& row,
                                                                               std::wstring_view cultureName) noexcept
{
    for (const RedConfigure::LocalizationTargetCell& cell : row.targets)
    {
        if (CompareTextIgnoreCase(cell.cultureName, cultureName) == 0)
        {
            return &cell;
        }
    }
    return nullptr;
}

[[nodiscard]] bool ReviewOwnerCultureHasDirtyCell(std::span<const RedConfigure::LocalizationReviewRow> rows,
                                                  std::wstring_view ownerName,
                                                  std::wstring_view cultureName) noexcept
{
    for (const RedConfigure::LocalizationReviewRow& row : rows)
    {
        if (CompareTextIgnoreCase(row.ownerName, ownerName) != 0)
        {
            continue;
        }

        const RedConfigure::LocalizationTargetCell* const cell = FindReviewTargetCell(row, cultureName);
        if (cell && cell->dirty)
        {
            return true;
        }
    }
    return false;
}

[[nodiscard]] RedConfigure::LocalizationTargetCell* FindReviewTargetCell(RedConfigure::LocalizationReviewRow& row, std::wstring_view cultureName) noexcept
{
    for (RedConfigure::LocalizationTargetCell& cell : row.targets)
    {
        if (CompareTextIgnoreCase(cell.cultureName, cultureName) == 0)
        {
            return &cell;
        }
    }
    return nullptr;
}

[[nodiscard]] bool IsReviewCultureVisible(std::span<const std::wstring> visibleCultureNames, std::wstring_view cultureName) noexcept
{
    return ContainsNameIgnoreCase(visibleCultureNames, cultureName);
}

[[nodiscard]] bool ReviewRowHasVisibleProblem(const RedConfigure::LocalizationReviewRow& row, std::span<const std::wstring> visibleCultureNames) noexcept
{
    for (const RedConfigure::LocalizationTargetCell& cell : row.targets)
    {
        if (! IsReviewCultureVisible(visibleCultureNames, cell.cultureName))
        {
            continue;
        }
        if (cell.validation.status != RedConfigure::Localization::PlaceholderStatus::Ok)
        {
            return true;
        }
        // Untranslated cells fall back to the English source text; surface them as
        // problems so the review Problems filter does not hide missing translations.
        if (! cell.hasExistingTranslation && ! cell.dirty)
        {
            return true;
        }
    }
    return false;
}

[[nodiscard]] bool ReviewRowMatchesSearch(const RedConfigure::LocalizationReviewRow& row,
                                          std::span<const std::wstring> visibleCultureNames,
                                          std::wstring_view searchText) noexcept
{
    if (searchText.empty() || ContainsIgnoreCase(row.ownerName, searchText) || ContainsIgnoreCase(row.id, searchText) ||
        ContainsIgnoreCase(row.sourceText, searchText))
    {
        return true;
    }

    for (const RedConfigure::LocalizationTargetCell& cell : row.targets)
    {
        if (IsReviewCultureVisible(visibleCultureNames, cell.cultureName) && ContainsIgnoreCase(cell.targetText, searchText))
        {
            return true;
        }
    }
    return false;
}

[[nodiscard]] std::wstring_view ReviewTargetSortText(const RedConfigure::LocalizationReviewRow& row, std::wstring_view cultureName) noexcept
{
    if (const RedConfigure::LocalizationTargetCell* cell = FindReviewTargetCell(row, cultureName))
    {
        return cell->targetText;
    }
    return {};
}

[[nodiscard]] std::wstring_view ReviewColumnSortText(const RedConfigure::LocalizationReviewRow& row,
                                                     const RedConfigure::LocalizationReviewViewOptions& options) noexcept
{
    switch (options.sortColumn)
    {
        case RedConfigure::LocalizationViewColumn::Owner: return row.ownerName;
        case RedConfigure::LocalizationViewColumn::Id: return row.id;
        case RedConfigure::LocalizationViewColumn::Source: return row.sourceText;
        case RedConfigure::LocalizationViewColumn::Target: return ReviewTargetSortText(row, options.sortCultureName);
        case RedConfigure::LocalizationViewColumn::Status: return {};
        default: return row.id;
    }
}

[[nodiscard]] Common::Settings::ThemeDefinition MakeFallbackTheme()
{
    Common::Settings::ThemeDefinition theme;
    theme.id          = L"user/redconfigure";
    theme.name        = L"RedConfigure";
    theme.baseThemeId = L"builtin/dark";
    theme.colors.emplace(L"app.accent", 0xFF0F6CBDu);
    theme.colors.emplace(L"window.background", 0xFFF8F8F8u);
    theme.colors.emplace(L"menu.background", 0xFFFFFFFFu);
    theme.colors.emplace(L"folderView.background", 0xFFFFFFFFu);
    theme.colors.emplace(L"folderView.itemBackgroundSelected", 0xFF0F6CBDu);
    return theme;
}

[[nodiscard]] std::wstring SanitizePathPart(std::wstring text)
{
    for (wchar_t& ch : text)
    {
        switch (ch)
        {
            case L'\\':
            case L'/':
            case L':':
            case L'*':
            case L'?':
            case L'"':
            case L'<':
            case L'>':
            case L'|': ch = L'-'; break;
            default: break;
        }
    }
    return text.empty() ? std::wstring(L"RedConfigure") : text;
}
} // namespace

namespace RedConfigure
{
std::vector<size_t> BuildTranslationView(std::span<const TranslationEntry> translations, const LocalizationViewOptions& options)
{
    std::vector<size_t> result;
    result.reserve(translations.size());
    for (size_t index = 0u; index < translations.size(); ++index)
    {
        if (TranslationMatchesViewFilters(translations[index], options))
        {
            result.push_back(index);
        }
    }

    if (options.sortDirection == LocalizationSortDirection::None)
    {
        return result;
    }

    std::stable_sort(result.begin(),
                     result.end(),
                     [&](size_t lhsIndex, size_t rhsIndex) noexcept
    {
        const TranslationEntry& lhs = translations[lhsIndex];
        const TranslationEntry& rhs = translations[rhsIndex];

        int comparison = 0;
        if (options.sortColumn == LocalizationViewColumn::Status)
        {
            comparison = static_cast<int>(lhs.validation.status) - static_cast<int>(rhs.validation.status);
        }
        else
        {
            comparison = CompareTextIgnoreCase(TranslationColumnText(lhs, options.sortColumn), TranslationColumnText(rhs, options.sortColumn));
        }

        if (comparison == 0)
        {
            comparison = CompareTextIgnoreCase(lhs.id, rhs.id);
        }
        if (comparison == 0)
        {
            return lhsIndex < rhsIndex;
        }

        return options.sortDirection == LocalizationSortDirection::Ascending ? comparison < 0 : comparison > 0;
    });

    return result;
}

std::vector<size_t> BuildLocalizationReviewView(std::span<const LocalizationReviewRow> rows, const LocalizationReviewViewOptions& options)
{
    std::vector<size_t> result;
    result.reserve(rows.size());
    for (size_t index = 0u; index < rows.size(); ++index)
    {
        const LocalizationReviewRow& row = rows[index];
        if (! ContainsNameIgnoreCase(options.visibleOwnerNames, row.ownerName))
        {
            continue;
        }
        if (! options.idFilterText.empty() && ! ContainsIgnoreCase(row.id, options.idFilterText))
        {
            continue;
        }

        const bool hasProblem = ReviewRowHasVisibleProblem(row, options.visibleCultureNames);
        if (options.statusFilter == LocalizationStatusFilter::Ok && hasProblem)
        {
            continue;
        }
        if (options.statusFilter == LocalizationStatusFilter::Problems && ! hasProblem)
        {
            continue;
        }
        if (! ReviewRowMatchesSearch(row, options.visibleCultureNames, options.searchText))
        {
            continue;
        }

        result.push_back(index);
    }

    if (options.sortDirection == LocalizationSortDirection::None)
    {
        return result;
    }

    std::stable_sort(result.begin(),
                     result.end(),
                     [&](size_t lhsIndex, size_t rhsIndex) noexcept
    {
        const LocalizationReviewRow& lhs = rows[lhsIndex];
        const LocalizationReviewRow& rhs = rows[rhsIndex];

        int comparison = 0;
        if (options.sortColumn == LocalizationViewColumn::Status)
        {
            comparison = static_cast<int>(ReviewRowHasVisibleProblem(lhs, options.visibleCultureNames)) -
                         static_cast<int>(ReviewRowHasVisibleProblem(rhs, options.visibleCultureNames));
        }
        else
        {
            comparison = CompareTextIgnoreCase(ReviewColumnSortText(lhs, options), ReviewColumnSortText(rhs, options));
        }

        if (comparison == 0)
        {
            comparison = CompareTextIgnoreCase(lhs.ownerName, rhs.ownerName);
        }
        if (comparison == 0)
        {
            comparison = CompareTextIgnoreCase(lhs.id, rhs.id);
        }
        if (comparison == 0)
        {
            return lhsIndex < rhsIndex;
        }

        return options.sortDirection == LocalizationSortDirection::Ascending ? comparison < 0 : comparison > 0;
    });

    return result;
}

HRESULT RedConfigureSession::LoadWorkspace(const std::filesystem::path& root, std::wstring cultureName)
{
    _workspace                = {};
    _themeCatalog             = {};
    _themePreview             = {};
    _translations             = {};
    _localizationReviewRows   = {};
    _localizationReviewCultures = {};
    _inventoryEntries         = {};
    _activeResourceOwnerName  = {};
    _activeResourceOwnerIndex = 0u;
    _activeThemeIndex         = 0u;
    _cultureName              = cultureName.empty() ? std::wstring(L"en-US") : std::move(cultureName);

    {
        Debug::Perf::Scope discoverPerf(L"redconfigure.workspace.discover_us");
        if (const HRESULT hr = Workspace::DiscoverWorkspace(root, _workspace); FAILED(hr))
        {
            discoverPerf.SetHr(hr);
            return hr;
        }
        discoverPerf.SetValue0(static_cast<uint64_t>(_workspace.resourceOwners.size()));
        discoverPerf.SetValue1(static_cast<uint64_t>(_workspace.themeFiles.size()));
    }

    {
        Debug::Perf::Scope themePerf(L"redconfigure.theme_catalog.load_us");
        if (const HRESULT hr = Themes::LoadThemeCatalog(_workspace.themeFiles, _themeCatalog); FAILED(hr))
        {
            themePerf.SetHr(hr);
            return hr;
        }
        themePerf.SetValue0(static_cast<uint64_t>(_themeCatalog.themes.size()));
    }

    if (_themeCatalog.themes.empty())
    {
        _themePreview.SetTheme(MakeFallbackTheme());
    }
    else
    {
        _themePreview.SetTheme(_themeCatalog.themes.front().definition);
    }

    {
        Debug::Perf::Scope reviewPerf(L"redconfigure.localization_review.load_us");
        if (const HRESULT hr = LoadLocalizationReview(); FAILED(hr))
        {
            reviewPerf.SetHr(hr);
            return hr;
        }
        reviewPerf.SetValue0(static_cast<uint64_t>(_localizationReviewRows.size()));
        reviewPerf.SetValue1(static_cast<uint64_t>(_localizationReviewCultures.size()));
    }

    Debug::Perf::Scope activeOwnerPerf(L"redconfigure.localization_active_owner.load_us");
    const HRESULT activeOwnerHr = LoadLocalizationForActiveOwner();
    activeOwnerPerf.SetValue0(static_cast<uint64_t>(_translations.size()));
    activeOwnerPerf.SetValue1(static_cast<uint64_t>(_inventoryEntries.size()));
    activeOwnerPerf.SetHr(activeOwnerHr);
    return activeOwnerHr;
}

const Workspace::WorkspaceScanResult& RedConfigureSession::GetWorkspace() const noexcept
{
    return _workspace;
}

const Themes::ThemeCatalog& RedConfigureSession::GetThemeCatalog() const noexcept
{
    return _themeCatalog;
}

const Themes::ThemePreviewModel& RedConfigureSession::GetThemePreviewModel() const noexcept
{
    return _themePreview;
}

Themes::ThemePreviewModel& RedConfigureSession::GetThemePreviewModel() noexcept
{
    return _themePreview;
}

std::span<const TranslationEntry> RedConfigureSession::GetTranslations() const noexcept
{
    return _translations;
}

std::span<const LocalizationReviewRow> RedConfigureSession::GetLocalizationReviewRows() const noexcept
{
    return _localizationReviewRows;
}

std::span<const std::wstring> RedConfigureSession::GetLocalizationReviewCultures() const noexcept
{
    return _localizationReviewCultures;
}

std::span<const InventoryEntry> RedConfigureSession::GetInventoryEntries() const noexcept
{
    return _inventoryEntries;
}

std::wstring_view RedConfigureSession::GetCultureName() const noexcept
{
    return _cultureName;
}

std::wstring_view RedConfigureSession::GetActiveResourceOwnerName() const noexcept
{
    return _activeResourceOwnerName;
}

size_t RedConfigureSession::GetActiveResourceOwnerIndex() const noexcept
{
    return _activeResourceOwnerIndex;
}

size_t RedConfigureSession::GetActiveThemeIndex() const noexcept
{
    return _activeThemeIndex;
}

HRESULT RedConfigureSession::SetActiveResourceOwner(size_t ownerIndex)
{
    if (ownerIndex >= _workspace.resourceOwners.size())
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_INDEX);
    }

    _activeResourceOwnerIndex = ownerIndex;
    return LoadLocalizationForActiveOwner();
}

bool RedConfigureSession::SetActiveTheme(size_t themeIndex)
{
    if (themeIndex >= _themeCatalog.themes.size())
    {
        return false;
    }

    _activeThemeIndex = themeIndex;
    _themePreview.SetTheme(_themeCatalog.themes[themeIndex].definition);
    return true;
}

bool RedConfigureSession::UpdateTranslation(size_t rowIndex, std::wstring_view targetText)
{
    if (rowIndex >= _translations.size())
    {
        return false;
    }

    TranslationEntry& entry = _translations[rowIndex];
    const auto validation   = Localization::ValidatePlaceholders(entry.sourceText, targetText);
    if (validation.status != Localization::PlaceholderStatus::Ok)
    {
        return false;
    }

    entry.targetText.assign(targetText);
    entry.validation = validation;
    return true;
}

bool RedConfigureSession::EnsureLocalizationReviewCulture(std::wstring_view cultureName)
{
    if (cultureName.empty() || CompareTextIgnoreCase(cultureName, L"en-US") == 0)
    {
        return false;
    }

    for (const std::wstring& existingCulture : _localizationReviewCultures)
    {
        if (CompareTextIgnoreCase(existingCulture, cultureName) == 0)
        {
            return true;
        }
    }

    const std::wstring culture(cultureName);
    _localizationReviewCultures.push_back(culture);
    for (LocalizationReviewRow& row : _localizationReviewRows)
    {
        LocalizationTargetCell cell;
        cell.cultureName            = culture;
        cell.targetText             = row.sourceText;
        cell.validation             = Localization::ValidatePlaceholders(row.sourceText, cell.targetText);
        cell.hasExistingTranslation = false;
        cell.dirty                  = false;
        row.targets.push_back(std::move(cell));
    }
    return true;
}

bool RedConfigureSession::UpdateLocalizationReviewTarget(size_t rowIndex, std::wstring_view cultureName, std::wstring_view targetText)
{
    if (rowIndex >= _localizationReviewRows.size() || cultureName.empty())
    {
        return false;
    }

    if (! FindReviewTargetCell(_localizationReviewRows[rowIndex], cultureName) && ! EnsureLocalizationReviewCulture(cultureName))
    {
        return false;
    }

    LocalizationReviewRow& row          = _localizationReviewRows[rowIndex];
    LocalizationTargetCell* const cell = FindReviewTargetCell(row, cultureName);
    if (! cell)
    {
        return false;
    }

    const auto validation = Localization::ValidatePlaceholders(row.sourceText, targetText);
    if (validation.status != Localization::PlaceholderStatus::Ok)
    {
        return false;
    }

    cell->targetText.assign(targetText);
    cell->validation             = validation;
    cell->hasExistingTranslation = true;
    cell->dirty                  = true;
    return true;
}

bool RedConfigureSession::UpdateThemeColor(std::wstring_view colorKey, std::wstring_view colorText)
{
    return _themePreview.TryEditOverride(colorKey, colorText);
}

std::filesystem::path RedConfigureSession::GetDefaultLocalizationExportPath() const
{
    const std::wstring owner = _activeResourceOwnerName.empty() ? std::wstring(L"RedConfigure") : _activeResourceOwnerName;
    return _workspace.root / L"RedConfigureOutput" / (SanitizePathPart(owner + L"-" + _cultureName) + L".rc");
}

std::filesystem::path RedConfigureSession::GetDefaultThemeExportPath() const
{
    std::wstring themePart                 = _themePreview.GetTheme().id;
    constexpr std::wstring_view userPrefix = L"user/";
    if (themePart.rfind(userPrefix, 0u) == 0u)
    {
        themePart.erase(0u, userPrefix.size());
    }
    return _workspace.root / L"RedConfigureOutput" / (SanitizePathPart(std::move(themePart)) + L".theme.json5");
}

HRESULT RedConfigureSession::ExportLocalization(const std::filesystem::path& path) const
{
    std::wstring rcText;
    if (const HRESULT hr = BuildLocalizationExportText(rcText); FAILED(hr))
    {
        return hr;
    }

    return WriteUtf16LeTextFile(path, rcText);
}

HRESULT RedConfigureSession::ExportLocalizationReview(const std::filesystem::path& outputRoot, size_t* exportedFileCount) const
{
    if (exportedFileCount)
    {
        *exportedFileCount = 0u;
    }

    std::vector<LocalizationExportPreview> previews;
    if (const HRESULT hr = BuildLocalizationReviewExportPreviews(previews); FAILED(hr))
    {
        return hr;
    }

    for (const LocalizationExportPreview& preview : previews)
    {
        const fs::path path = outputRoot.empty() ? preview.path : outputRoot / preview.path.filename();
        if (const HRESULT hr = WriteUtf16LeTextFile(path, preview.text); FAILED(hr))
        {
            return hr;
        }
    }

    if (exportedFileCount)
    {
        *exportedFileCount = previews.size();
    }
    return S_OK;
}

HRESULT RedConfigureSession::BuildLocalizationExportText(std::wstring& outText) const
{
    outText.clear();

    std::vector<Localization::MergedStringEntry> merged;
    merged.reserve(_translations.size());
    for (const TranslationEntry& entry : _translations)
    {
        merged.push_back(Localization::MergedStringEntry{.id = entry.id, .sourceText = entry.sourceText, .targetText = entry.targetText});
    }

    outText = Localization::BuildSatelliteRcStringTable(L"resource.h", _cultureName, merged);
    return S_OK;
}

HRESULT RedConfigureSession::BuildLocalizationReviewExportPreviews(std::vector<LocalizationExportPreview>& outPreviews) const
{
    outPreviews.clear();

    for (const Workspace::ResourceOwner& owner : _workspace.resourceOwners)
    {
        for (const std::wstring& culture : _localizationReviewCultures)
        {
            if (! ReviewOwnerCultureHasDirtyCell(_localizationReviewRows, owner.name, culture))
            {
                continue;
            }

            std::vector<Localization::MergedStringEntry> merged;
            for (const LocalizationReviewRow& row : _localizationReviewRows)
            {
                if (CompareTextIgnoreCase(row.ownerName, owner.name) != 0)
                {
                    continue;
                }

                const LocalizationTargetCell* const cell = FindReviewTargetCell(row, culture);
                if (! cell)
                {
                    continue;
                }

                merged.push_back(Localization::MergedStringEntry{.id = row.id, .sourceText = row.sourceText, .targetText = cell->targetText});
            }

            if (merged.empty())
            {
                continue;
            }

            LocalizationExportPreview preview;
            preview.ownerName   = owner.name;
            preview.cultureName = culture;
            preview.path        = _workspace.root / L"RedConfigureOutput" / (SanitizePathPart(owner.name + L"-" + culture) + L".rc");
            preview.text        = Localization::BuildSatelliteRcStringTable(L"resource.h", culture, merged);
            outPreviews.push_back(std::move(preview));
        }
    }

    return S_OK;
}

HRESULT RedConfigureSession::BuildThemeExportText(std::string& outJson) const
{
    outJson.clear();
    return Common::Settings::BuildThemeDefinitionJson5(_themePreview.BuildFlattenedTheme(), outJson);
}

HRESULT RedConfigureSession::ExportTheme(const std::filesystem::path& path) const
{
    std::string jsonText;
    if (const HRESULT hr = BuildThemeExportText(jsonText); FAILED(hr))
    {
        return hr;
    }

    return WriteBinaryFile(path, jsonText);
}

HRESULT RedConfigureSession::LoadLocalizationReview()
{
    _localizationReviewRows.clear();
    _localizationReviewCultures = DiscoverLocalizationReviewCultures(_workspace);

    const size_t ownerCount  = _workspace.resourceOwners.size();
    const size_t workerCount = LocalizationReviewWorkerCount(ownerCount);
    Debug::Perf::EmitValue(L"redconfigure.localization_review.worker_count", static_cast<uint64_t>(workerCount));
    if (ownerCount == 0u)
    {
        return S_OK;
    }

    if (workerCount == 1u)
    {
        std::vector<LocalizationReviewOwnerLoadResult> results =
            LoadLocalizationReviewOwnerRange(_workspace.resourceOwners, _localizationReviewCultures, 0u, ownerCount);
        for (LocalizationReviewOwnerLoadResult& result : results)
        {
            AppendLocalizationReviewLoadResult(result, _localizationReviewRows, _workspace.errors);
        }
        return S_OK;
    }

    const size_t chunkSize = (ownerCount + workerCount - 1u) / workerCount;
    std::vector<std::future<std::vector<LocalizationReviewOwnerLoadResult>>> futures;
    futures.reserve(workerCount);

    size_t nextOwnerIndex = 0u;
    for (; nextOwnerIndex < ownerCount; nextOwnerIndex += chunkSize)
    {
        const size_t beginIndex = nextOwnerIndex;
        const size_t endIndex   = (std::min)(ownerCount, beginIndex + chunkSize);
        try
        {
            futures.push_back(std::async(std::launch::async,
                                         [owners = std::span<const Workspace::ResourceOwner>(_workspace.resourceOwners),
                                          cultures = std::span<const std::wstring>(_localizationReviewCultures),
                                          beginIndex,
                                          endIndex]
            {
                return LoadLocalizationReviewOwnerRange(owners, cultures, beginIndex, endIndex);
            }));
        }
        catch (const std::system_error&)
        {
            // Thread creation is opportunistic here; already scheduled workers are joined below and the remaining owners load serially.
            break;
        }
    }

    for (std::future<std::vector<LocalizationReviewOwnerLoadResult>>& future : futures)
    {
        std::vector<LocalizationReviewOwnerLoadResult> results = future.get();
        for (LocalizationReviewOwnerLoadResult& result : results)
        {
            AppendLocalizationReviewLoadResult(result, _localizationReviewRows, _workspace.errors);
        }
    }

    if (nextOwnerIndex < ownerCount)
    {
        std::vector<LocalizationReviewOwnerLoadResult> results =
            LoadLocalizationReviewOwnerRange(_workspace.resourceOwners, _localizationReviewCultures, nextOwnerIndex, ownerCount);
        for (LocalizationReviewOwnerLoadResult& result : results)
        {
            AppendLocalizationReviewLoadResult(result, _localizationReviewRows, _workspace.errors);
        }
    }

    return S_OK;
}

HRESULT RedConfigureSession::LoadLocalizationForActiveOwner()
{
    _translations.clear();
    _inventoryEntries.clear();
    _activeResourceOwnerName.clear();
    if (_workspace.resourceOwners.empty())
    {
        return S_OK;
    }

    if (_activeResourceOwnerIndex >= _workspace.resourceOwners.size())
    {
        _activeResourceOwnerIndex = 0u;
    }

    const Workspace::ResourceOwner& owner = _workspace.resourceOwners[_activeResourceOwnerIndex];
    _activeResourceOwnerName              = owner.name;

    Localization::RcParseResult source;
    if (! ReadAndParseRcStringTables(owner.embeddedResourcePath,
                                     L"Read active localization source",
                                     L"Parse active localization source",
                                     source,
                                     _workspace.errors))
    {
        return S_OK;
    }

    _inventoryEntries.reserve(source.localizableEntries.size());
    for (const Localization::RcLocalizableEntry& entry : source.localizableEntries)
    {
        _inventoryEntries.push_back(InventoryEntry{
            .kind       = entry.kind,
            .ownerName  = _activeResourceOwnerName,
            .resourceId = entry.ownerId.empty() ? _activeResourceOwnerName : entry.ownerId,
            .itemId     = entry.id,
            .sourceText = entry.text,
            .sourceLine = entry.sourceLine,
        });
    }

    Localization::RcParseResult target;
    if (const auto satellitePath = FindSatelliteForCulture(owner, _cultureName))
    {
        static_cast<void>(ReadAndParseRcStringTables(satellitePath.value(),
                                                     L"Read active localization target",
                                                     L"Parse active localization target",
                                                     target,
                                                     _workspace.errors));
    }

    const std::vector<Localization::MergedStringEntry> merged = Localization::MergeStringTables(source.strings, target.strings);
    _translations.reserve(merged.size());
    for (const Localization::MergedStringEntry& entry : merged)
    {
        TranslationEntry translation;
        translation.id         = entry.id;
        translation.sourceText = entry.sourceText;
        translation.targetText = entry.targetText.empty() ? entry.sourceText : entry.targetText;
        translation.validation = Localization::ValidatePlaceholders(translation.sourceText, translation.targetText);
        _translations.push_back(std::move(translation));
    }

    return S_OK;
}
} // namespace RedConfigure
