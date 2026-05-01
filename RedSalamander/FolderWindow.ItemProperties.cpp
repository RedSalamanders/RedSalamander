#include "FolderWindow.h"

#include "AppTheme.h"
#include "DxUi/DxUi.h"
#include "DxUiThemePalette.h"
#include "Resource.h"
#include "WindowMessages.h"
#include "WindowMaximizeBehavior.h"
#include "WindowPlacementPersistence.h"
#include "WindowSizing.h"

#include "Helpers.h"
#include "PlugInterfaces/Informations.h"
#include "SettingsHotReload.h"

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4514 28182) // WIL headers: deleted copy/move and unused inline Helpers
#include <wil/resource.h>
#pragma warning(pop)

#pragma warning(push)
#pragma warning(disable : 6297 28182) // yyjson warnings
#include <yyjson.h>
#pragma warning(pop)

#include <algorithm>
#include <array>
#include <atomic>
#include <cstdint>
#include <cmath>
#include <format>
#include <functional>
#include <iterator>
#include <limits>
#include <memory>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

namespace
{
struct ItemPropertiesField
{
    std::wstring key;
    std::wstring value;
};

struct ItemPropertiesSection
{
    std::wstring title;
    std::vector<ItemPropertiesField> fields;
};

struct ItemPropertiesStream
{
    std::wstring name;
    uint64_t sizeBytes = 0u;
    std::wstring displaySize;
    bool canRemove = false;
};

struct ItemPropertiesDocument
{
    std::wstring title;
    std::vector<ItemPropertiesSection> sections;
    std::vector<ItemPropertiesStream> streams;
};

using ItemPropertiesOpenStreamCallback = std::function<HRESULT(std::wstring_view streamName)>;

struct ItemPropertiesLoadResult
{
    uint64_t generation = 0u;
    HRESULT hr = E_FAIL;
    std::string jsonUtf8;
};

class ItemPropertiesWindow;

struct ItemPropertiesLoadWork
{
    HWND hwnd = nullptr;
    ItemPropertiesWindow* window = nullptr;
    uint64_t windowToken = 0u;
    uint64_t generation = 0u;
    std::filesystem::path itemPath;
    wil::com_ptr<IFileSystemIO> itemIo;
#ifdef ENABLE_TESTS
    uint32_t delayMs = 0u;
#endif
};

inline std::atomic<uint64_t> g_nextItemPropertiesWindowToken{1u};

[[nodiscard]] std::wstring Utf16FromUtf8(std::string_view text) noexcept
{
    if (text.empty())
    {
        return {};
    }

    if (text.size() > static_cast<size_t>(std::numeric_limits<int>::max()))
    {
        return {};
    }

    const int required = ::MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), nullptr, 0);
    if (required <= 0)
    {
        return {};
    }

    std::wstring result(static_cast<size_t>(required), L'\0');
    const int written = ::MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), result.data(), required);
    if (written != required)
    {
        return {};
    }

    return result;
}

[[nodiscard]] std::wstring NormalizeItemPropertiesTitle(std::wstring_view title);
[[nodiscard]] std::wstring NormalizeItemPropertiesSectionTitle(std::wstring_view title);
[[nodiscard]] std::wstring NormalizeItemPropertiesFieldKey(std::wstring_view key);
[[nodiscard]] std::wstring NormalizeItemPropertiesFieldValue(std::wstring_view normalizedKey, std::wstring_view value);
void NormalizeItemPropertiesSectionOrder(ItemPropertiesDocument& doc);

[[nodiscard]] std::optional<ItemPropertiesDocument> TryParseItemPropertiesJson(std::string_view jsonUtf8) noexcept
{
    if (jsonUtf8.empty())
    {
        return std::nullopt;
    }

    // yyjson may modify the input buffer; it requires a mutable char*.
    std::string jsonCopy(jsonUtf8);
    yyjson_read_err err{};
    yyjson_doc* doc = yyjson_read_opts(jsonCopy.data(), jsonCopy.size(), YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM, nullptr, &err);
    if (! doc)
    {
        return std::nullopt;
    }
    auto freeDoc = wil::scope_exit([&] { yyjson_doc_free(doc); });

    yyjson_val* root = yyjson_doc_get_root(doc);
    if (! root || ! yyjson_is_obj(root))
    {
        return std::nullopt;
    }

    yyjson_val* versionVal = yyjson_obj_get(root, "version");
    if (! versionVal || ! yyjson_is_int(versionVal) || yyjson_get_int(versionVal) != 1)
    {
        return std::nullopt;
    }

    ItemPropertiesDocument out{};

    if (yyjson_val* titleVal = yyjson_obj_get(root, "title"); titleVal && yyjson_is_str(titleVal))
    {
        if (const char* titleUtf8 = yyjson_get_str(titleVal); titleUtf8 && titleUtf8[0] != '\0')
        {
            out.title = NormalizeItemPropertiesTitle(Utf16FromUtf8(titleUtf8));
        }
    }

    yyjson_val* sectionsVal = yyjson_obj_get(root, "sections");
    if (! sectionsVal || ! yyjson_is_arr(sectionsVal))
    {
        return out;
    }

    const size_t sectionCount = yyjson_arr_size(sectionsVal);
    out.sections.reserve(sectionCount);

    for (size_t i = 0; i < sectionCount; ++i)
    {
        yyjson_val* sectionVal = yyjson_arr_get(sectionsVal, i);
        if (! sectionVal || ! yyjson_is_obj(sectionVal))
        {
            continue;
        }

        ItemPropertiesSection section{};
        if (yyjson_val* sectionTitleVal = yyjson_obj_get(sectionVal, "title"); sectionTitleVal && yyjson_is_str(sectionTitleVal))
        {
            if (const char* titleUtf8 = yyjson_get_str(sectionTitleVal); titleUtf8 && titleUtf8[0] != '\0')
            {
                section.title = NormalizeItemPropertiesSectionTitle(Utf16FromUtf8(titleUtf8));
            }
        }

        if (yyjson_val* fieldsVal = yyjson_obj_get(sectionVal, "fields"); fieldsVal && yyjson_is_arr(fieldsVal))
        {
            const size_t fieldCount = yyjson_arr_size(fieldsVal);
            section.fields.reserve(fieldCount);

            for (size_t f = 0; f < fieldCount; ++f)
            {
                yyjson_val* fieldVal = yyjson_arr_get(fieldsVal, f);
                if (! fieldVal || ! yyjson_is_obj(fieldVal))
                {
                    continue;
                }

                yyjson_val* keyVal   = yyjson_obj_get(fieldVal, "key");
                yyjson_val* valueVal = yyjson_obj_get(fieldVal, "value");
                if (! keyVal || ! valueVal || ! yyjson_is_str(keyVal) || ! yyjson_is_str(valueVal))
                {
                    continue;
                }

                const char* keyUtf8   = yyjson_get_str(keyVal);
                const char* valueUtf8 = yyjson_get_str(valueVal);
                if (! keyUtf8 || keyUtf8[0] == '\0' || ! valueUtf8)
                {
                    continue;
                }

                ItemPropertiesField field{};
                field.key   = NormalizeItemPropertiesFieldKey(Utf16FromUtf8(keyUtf8));
                field.value = NormalizeItemPropertiesFieldValue(field.key, Utf16FromUtf8(valueUtf8));
                if (! field.key.empty())
                {
                    section.fields.emplace_back(std::move(field));
                }
            }
        }

        out.sections.emplace_back(std::move(section));
    }

    NormalizeItemPropertiesSectionOrder(out);

    if (yyjson_val* streamsVal = yyjson_obj_get(root, "streams"); streamsVal && yyjson_is_arr(streamsVal))
    {
        const size_t streamCount = yyjson_arr_size(streamsVal);
        out.streams.reserve(streamCount);

        for (size_t i = 0; i < streamCount; ++i)
        {
            yyjson_val* streamVal = yyjson_arr_get(streamsVal, i);
            if (! streamVal || ! yyjson_is_obj(streamVal))
            {
                continue;
            }

            yyjson_val* nameVal = yyjson_obj_get(streamVal, "name");
            if (! nameVal || ! yyjson_is_str(nameVal))
            {
                continue;
            }

            const char* nameUtf8 = yyjson_get_str(nameVal);
            if (! nameUtf8 || nameUtf8[0] == '\0')
            {
                continue;
            }

            ItemPropertiesStream stream{};
            stream.name = Utf16FromUtf8(nameUtf8);
            if (stream.name.empty())
            {
                continue;
            }

            if (yyjson_val* sizeVal = yyjson_obj_get(streamVal, "sizeBytes"); sizeVal && yyjson_is_uint(sizeVal))
            {
                stream.sizeBytes = yyjson_get_uint(sizeVal);
            }
            else if (sizeVal && yyjson_is_sint(sizeVal) && yyjson_get_sint(sizeVal) >= 0)
            {
                stream.sizeBytes = static_cast<uint64_t>(yyjson_get_sint(sizeVal));
            }

            if (yyjson_val* displaySizeVal = yyjson_obj_get(streamVal, "displaySize"); displaySizeVal && yyjson_is_str(displaySizeVal))
            {
                if (const char* displaySizeUtf8 = yyjson_get_str(displaySizeVal); displaySizeUtf8 && displaySizeUtf8[0] != '\0')
                {
                    stream.displaySize = Utf16FromUtf8(displaySizeUtf8);
                }
            }
            if (stream.displaySize.empty())
            {
                stream.displaySize = std::format(L"{} bytes", stream.sizeBytes);
            }

            if (yyjson_val* canRemoveVal = yyjson_obj_get(streamVal, "canRemove"); canRemoveVal && yyjson_is_bool(canRemoveVal))
            {
                stream.canRemove = yyjson_get_bool(canRemoveVal) != 0;
            }

            out.streams.emplace_back(std::move(stream));
        }
    }

    return out;
}

constexpr wchar_t kItemPropertiesWindowClass[] = L"RedSalamander.ItemPropertiesWindow";
constexpr wchar_t kItemPropertiesWindowId[]    = L"ItemPropertiesWindow";
constexpr wchar_t kSettingsAppId[]             = L"RedSalamander";
constexpr UINT_PTR kItemPropertiesLoadingTimerId  = 1u;

#ifdef ENABLE_TESTS
std::atomic_uint32_t g_nextItemPropertiesLoadDelayMs{0u};
#endif

[[nodiscard]] bool IsCompactItemPropertiesSection(std::wstring_view title) noexcept
{
    return title == L"Timestamps" || title == L"Attributes";
}

[[nodiscard]] bool TryParseUnsignedDecimal(std::wstring_view text, uint64_t& valueOut) noexcept
{
    if (text.empty())
    {
        return false;
    }

    uint64_t value = 0u;
    for (const wchar_t ch : text)
    {
        if (ch < L'0' || ch > L'9')
        {
            return false;
        }

        constexpr uint64_t kMaxBeforeMultiply = (std::numeric_limits<uint64_t>::max)() / 10u;
        if (value > kMaxBeforeMultiply)
        {
            return false;
        }

        const uint64_t digit = static_cast<uint64_t>(ch - L'0');
        value *= 10u;
        if (value > (std::numeric_limits<uint64_t>::max)() - digit)
        {
            return false;
        }

        value += digit;
    }

    valueOut = value;
    return true;
}

[[nodiscard]] std::optional<std::wstring> TryFormatFileTimeLocal(uint64_t value) noexcept
{
    if (value == 0u)
    {
        return std::nullopt;
    }

    ULARGE_INTEGER ull{};
    ull.QuadPart = value;

    FILETIME fileTime{};
    fileTime.dwLowDateTime  = ull.LowPart;
    fileTime.dwHighDateTime = ull.HighPart;

    FILETIME localFileTime{};
    if (::FileTimeToLocalFileTime(&fileTime, &localFileTime) == 0)
    {
        return std::nullopt;
    }

    SYSTEMTIME localSystemTime{};
    if (::FileTimeToSystemTime(&localFileTime, &localSystemTime) == 0)
    {
        return std::nullopt;
    }

    return std::format(L"{:04}-{:02}-{:02} {:02}:{:02}:{:02}",
                       localSystemTime.wYear,
                       localSystemTime.wMonth,
                       localSystemTime.wDay,
                       localSystemTime.wHour,
                       localSystemTime.wMinute,
                       localSystemTime.wSecond);
}

[[nodiscard]] std::wstring FormatItemPropertiesSizeValue(uint64_t sizeBytes)
{
    const std::wstring exactBytes = std::format(L"{} bytes", sizeBytes);
    if (sizeBytes < 1024ull)
    {
        return exactBytes;
    }

    return std::format(L"{} ({})", FormatBytesCompact(sizeBytes), exactBytes);
}

[[nodiscard]] std::wstring NormalizeItemPropertiesTitle(std::wstring_view title)
{
    if (OrdinalString::EqualsNoCase(title, L"properties"))
    {
        return LoadStringResource(nullptr, IDS_CAPTION_PROPERTIES);
    }

    return std::wstring(title);
}

void NormalizeItemPropertiesSectionOrder(ItemPropertiesDocument& doc)
{
    auto generalIt = std::find_if(doc.sections.begin(), doc.sections.end(), [](const ItemPropertiesSection& section) {
        return section.title == L"General";
    });
    if (generalIt == doc.sections.end())
    {
        return;
    }

    auto timestampsIt = std::find_if(doc.sections.begin(), doc.sections.end(), [](const ItemPropertiesSection& section) {
        return section.title == L"Timestamps";
    });
    if (timestampsIt == doc.sections.end())
    {
        return;
    }

    auto targetIt = std::next(generalIt);
    if (timestampsIt == targetIt)
    {
        return;
    }

    if (timestampsIt < targetIt)
    {
        std::rotate(timestampsIt, std::next(timestampsIt), targetIt);
        return;
    }

    std::rotate(targetIt, timestampsIt, std::next(timestampsIt));
}

[[nodiscard]] std::wstring NormalizeItemPropertiesSectionTitle(std::wstring_view title)
{
    if (OrdinalString::EqualsNoCase(title, L"general"))
    {
        return L"General";
    }
    if (OrdinalString::EqualsNoCase(title, L"timestamps"))
    {
        return L"Timestamps";
    }
    if (OrdinalString::EqualsNoCase(title, L"attributes"))
    {
        return L"Attributes";
    }
    if (OrdinalString::EqualsNoCase(title, L"remote"))
    {
        return L"Remote";
    }
    if (OrdinalString::EqualsNoCase(title, L"connection"))
    {
        return L"Connection";
    }
    if (OrdinalString::EqualsNoCase(title, L"drive"))
    {
        return L"Drive";
    }
    if (OrdinalString::EqualsNoCase(title, L"imap"))
    {
        return L"IMAP";
    }
    if (OrdinalString::EqualsNoCase(title, L"s3"))
    {
        return L"S3";
    }
    if (OrdinalString::EqualsNoCase(title, L"s3table"))
    {
        return L"S3 Table";
    }
    if (OrdinalString::EqualsNoCase(title, L"graph"))
    {
        return L"Graph metadata";
    }

    return std::wstring(title);
}

[[nodiscard]] std::wstring NormalizeItemPropertiesFieldKey(std::wstring_view key)
{
    if (OrdinalString::EqualsNoCase(key, L"name"))
    {
        return L"Name";
    }
    if (OrdinalString::EqualsNoCase(key, L"path"))
    {
        return L"Path";
    }
    if (OrdinalString::EqualsNoCase(key, L"type"))
    {
        return L"Type";
    }
    if (OrdinalString::EqualsNoCase(key, L"size") || OrdinalString::EqualsNoCase(key, L"sizeBytes"))
    {
        return L"Size";
    }
    if (OrdinalString::EqualsNoCase(key, L"attributes"))
    {
        return L"Attributes";
    }
    if (OrdinalString::EqualsNoCase(key, L"creationTime") || OrdinalString::EqualsNoCase(key, L"createdDateTime"))
    {
        return L"Created";
    }
    if (OrdinalString::EqualsNoCase(key, L"lastWriteTime") || OrdinalString::EqualsNoCase(key, L"lastModifiedDateTime"))
    {
        return L"Modified";
    }
    if (OrdinalString::EqualsNoCase(key, L"lastAccessTime"))
    {
        return L"Accessed";
    }
    if (OrdinalString::EqualsNoCase(key, L"changeTime"))
    {
        return L"Changed";
    }
    if (OrdinalString::EqualsNoCase(key, L"remotePath"))
    {
        return L"Remote path";
    }
    if (OrdinalString::EqualsNoCase(key, L"displayPath"))
    {
        return L"Display path";
    }
    if (OrdinalString::EqualsNoCase(key, L"basePath"))
    {
        return L"Base path";
    }
    if (OrdinalString::EqualsNoCase(key, L"connectionName"))
    {
        return L"Connection name";
    }
    if (OrdinalString::EqualsNoCase(key, L"connectionId"))
    {
        return L"Connection ID";
    }
    if (OrdinalString::EqualsNoCase(key, L"connectionAuthMode"))
    {
        return L"Authentication";
    }
    if (OrdinalString::EqualsNoCase(key, L"connectionSavePassword"))
    {
        return L"Save password";
    }
    if (OrdinalString::EqualsNoCase(key, L"connectionRequireHello"))
    {
        return L"Require hello";
    }
    if (OrdinalString::EqualsNoCase(key, L"connectTimeoutMs"))
    {
        return L"Connect timeout (ms)";
    }
    if (OrdinalString::EqualsNoCase(key, L"operationTimeoutMs"))
    {
        return L"Operation timeout (ms)";
    }
    if (OrdinalString::EqualsNoCase(key, L"ignoreSslTrust"))
    {
        return L"Ignore TLS trust";
    }
    if (OrdinalString::EqualsNoCase(key, L"ftpUseEpsv"))
    {
        return L"Use EPSV";
    }
    if (OrdinalString::EqualsNoCase(key, L"fromConnectionManagerProfile"))
    {
        return L"From connection profile";
    }
    if (OrdinalString::EqualsNoCase(key, L"hasPassword"))
    {
        return L"Has password";
    }
    if (OrdinalString::EqualsNoCase(key, L"hasSshPrivateKey"))
    {
        return L"Has SSH private key";
    }
    if (OrdinalString::EqualsNoCase(key, L"hasSshPublicKey"))
    {
        return L"Has SSH public key";
    }
    if (OrdinalString::EqualsNoCase(key, L"hasSshKnownHosts"))
    {
        return L"Has SSH known hosts";
    }
    if (OrdinalString::EqualsNoCase(key, L"fullPath"))
    {
        return L"Full path";
    }
    if (OrdinalString::EqualsNoCase(key, L"sentTime"))
    {
        return L"Sent";
    }
    if (OrdinalString::EqualsNoCase(key, L"recvTime"))
    {
        return L"Received";
    }
    if (OrdinalString::EqualsNoCase(key, L"uid"))
    {
        return L"UID";
    }
    if (OrdinalString::EqualsNoCase(key, L"archiveItemIndex"))
    {
        return L"Archive item index";
    }
    if (OrdinalString::EqualsNoCase(key, L"archivePath"))
    {
        return L"Archive path";
    }

    return std::wstring(key);
}

[[nodiscard]] bool IsItemPropertiesTimeKey(std::wstring_view normalizedKey) noexcept
{
    return normalizedKey == L"Created" || normalizedKey == L"Modified" || normalizedKey == L"Accessed" || normalizedKey == L"Changed" ||
           normalizedKey == L"Sent" || normalizedKey == L"Received";
}

[[nodiscard]] std::wstring NormalizeItemPropertiesFieldValue(std::wstring_view normalizedKey, std::wstring_view value)
{
    if (normalizedKey == L"Type")
    {
        if (OrdinalString::EqualsNoCase(value, L"file"))
        {
            return L"File";
        }
        if (OrdinalString::EqualsNoCase(value, L"directory") || OrdinalString::EqualsNoCase(value, L"folder"))
        {
            return L"Directory";
        }
    }

    uint64_t numericValue = 0u;
    if (normalizedKey == L"Size" && TryParseUnsignedDecimal(value, numericValue))
    {
        return FormatItemPropertiesSizeValue(numericValue);
    }

    if (IsItemPropertiesTimeKey(normalizedKey) && TryParseUnsignedDecimal(value, numericValue))
    {
        if (const std::optional<std::wstring> formatted = TryFormatFileTimeLocal(numericValue); formatted.has_value())
        {
            return formatted.value();
        }
    }

    return std::wstring(value);
}

[[nodiscard]] bool IsItemPropertiesWrapOpportunity(wchar_t ch) noexcept
{
    return ch == L'\\' || ch == L'/' || ch == L'-' || ch == L'_' || ch == L'.';
}

[[nodiscard]] std::wstring MakeItemPropertiesWrapFriendlyText(std::wstring_view text)
{
    constexpr wchar_t kZeroWidthBreak = static_cast<wchar_t>(0x200B);
    constexpr size_t kMinWrapLength   = 32u;
    constexpr size_t kMaxRunLength    = 24u;

    if (text.size() < kMinWrapLength)
    {
        return std::wstring(text);
    }

    std::wstring wrapped;
    wrapped.reserve(text.size() + (text.size() / kMaxRunLength) + 4u);

    size_t runLength = 0u;
    for (wchar_t ch : text)
    {
        wrapped.push_back(ch);
        ++runLength;

        if (IsItemPropertiesWrapOpportunity(ch) || runLength >= kMaxRunLength)
        {
            wrapped.push_back(kZeroWidthBreak);
            runLength = 0u;
        }
    }

    return wrapped;
}

[[nodiscard]] float MeasureItemPropertiesWrappedTextHeightDip(const RedSalamander::DxUi::WindowHost& host,
                                                              std::wstring_view text,
                                                              RedSalamander::DxUi::FontRole role,
                                                              float widthDip) noexcept
{
    constexpr float kFallbackHeightDip = 20.0f;
    if (widthDip <= 1.0f || text.empty() || text.size() > static_cast<size_t>((std::numeric_limits<UINT32>::max)()))
    {
        return kFallbackHeightDip;
    }

    auto* factory = host.GetWriteFactory();
    auto* format  = host.GetTextFormat(role, DWRITE_TEXT_ALIGNMENT_LEADING, DWRITE_PARAGRAPH_ALIGNMENT_NEAR, true);
    if (! factory || ! format)
    {
        return kFallbackHeightDip;
    }

    wil::com_ptr<IDWriteTextLayout> layout;
    if (FAILED(factory->CreateTextLayout(text.data(),
                                         static_cast<UINT32>(text.size()),
                                         format,
                                         (std::max)(1.0f, widthDip),
                                         4096.0f,
                                         layout.addressof())) ||
        ! layout)
    {
        return kFallbackHeightDip;
    }

    DWRITE_TEXT_METRICS metrics{};
    if (FAILED(layout->GetMetrics(&metrics)))
    {
        return kFallbackHeightDip;
    }

    return (std::max)(kFallbackHeightDip, std::ceil(metrics.height));
}

[[nodiscard]] std::wstring BuildItemPropertiesText(const ItemPropertiesDocument& doc) noexcept
{
    std::wstring text;
    if (! doc.title.empty())
    {
        text.append(doc.title);
        text.append(L"\r\n\r\n");
    }

    for (size_t sectionIndex = 0; sectionIndex < doc.sections.size(); ++sectionIndex)
    {
        const auto& section = doc.sections[sectionIndex];
        if (! section.title.empty())
        {
            text.append(section.title);
            text.append(L"\r\n");
        }

        for (const auto& field : section.fields)
        {
            text.append(field.key);
            text.append(L": ");
            text.append(field.value);
            text.append(L"\r\n");
        }

        if (sectionIndex + 1u < doc.sections.size())
        {
            text.append(L"\r\n");
        }
    }

    if (doc.streams.empty())
    {
        return text;
    }

    if (! doc.sections.empty())
    {
        text.append(L"\r\n");
    }

    text.append(LoadStringResource(nullptr, IDS_PROPERTIES_STREAMS_TITLE));
    text.append(L"\r\n");
    for (const auto& stream : doc.streams)
    {
        text.append(stream.name);
        text.append(L": ");
        text.append(stream.displaySize);
        text.append(L"\r\n");
    }

    return text;
}

[[nodiscard]] size_t CountItemPropertiesFields(const ItemPropertiesDocument& doc) noexcept
{
    size_t count = 0u;
    for (const auto& section : doc.sections)
    {
        count += section.fields.size();
    }
    return count;
}

[[nodiscard]] size_t CountRemovableItemPropertiesStreams(const ItemPropertiesDocument& doc, bool streamRemovalAvailable) noexcept
{
    if (! streamRemovalAvailable)
    {
        return 0u;
    }

    return static_cast<size_t>(std::ranges::count_if(doc.streams, [](const ItemPropertiesStream& stream) noexcept { return stream.canRemove; }));
}

[[nodiscard]] size_t CountViewableItemPropertiesStreams(const ItemPropertiesDocument& doc, bool streamViewingAvailable) noexcept
{
    return streamViewingAvailable ? doc.streams.size() : 0u;
}

[[nodiscard]] std::wstring BuildItemPropertiesLoadingText()
{
    std::wstring text = LoadStringResource(nullptr, IDS_CAPTION_PROPERTIES);
    text.append(L"\r\n\r\n");
    text.append(LoadStringResource(nullptr, IDS_PROPERTIES_LOADING));
    return text;
}

#ifdef ENABLE_TESTS
[[nodiscard]] size_t CountVisibleItemPropertiesChildWindows(HWND hwnd) noexcept
{
    if (! hwnd || ::IsWindow(hwnd) == FALSE)
    {
        return 0u;
    }

    size_t count = 0u;
    ::EnumChildWindows(hwnd,
                       [](HWND child, LPARAM lParam) noexcept -> BOOL
    {
        auto& countRef = *reinterpret_cast<size_t*>(lParam);
        if (::IsWindowVisible(child) != FALSE)
        {
            ++countRef;
        }
        return TRUE;
    },
                       reinterpret_cast<LPARAM>(&count));
    return count;
}
#endif

class ItemPropertiesRootPanel final : public RedSalamander::DxUi::Panel
{
public:
    explicit ItemPropertiesRootPanel(const std::wstring* contentText) noexcept : _contentText(contentText)
    {
        SetFocusable(true);
    }

    bool OnKeyDown(RedSalamander::DxUi::WindowHost& host, UINT virtualKey, UINT modifiers) override
    {
        if ((modifiers & MK_CONTROL) != 0u)
        {
            if (virtualKey == L'A')
            {
                return true;
            }
            if (virtualKey == L'C')
            {
                return OnCopy(host);
            }
        }

        return RedSalamander::DxUi::Panel::OnKeyDown(host, virtualKey, modifiers);
    }

    bool OnCopy(RedSalamander::DxUi::WindowHost& host) override
    {
        return _contentText != nullptr && host.CopyTextToClipboard(*_contentText);
    }

    bool OnSelectAll(RedSalamander::DxUi::WindowHost& /*host*/) override
    {
        return true;
    }

private:
    const std::wstring* _contentText = nullptr;
};

struct ItemPropertiesFieldRowControls
{
    RedSalamander::DxUi::Label* key = nullptr;
    RedSalamander::DxUi::Label* value = nullptr;
};

struct ItemPropertiesSectionControls
{
    RedSalamander::DxUi::CardPanel* card = nullptr;
    RedSalamander::DxUi::Label* title = nullptr;
    bool compact = false;
    std::vector<ItemPropertiesFieldRowControls> fields;
};

struct ItemPropertiesStreamRowControls
{
    RedSalamander::DxUi::Label* name = nullptr;
    RedSalamander::DxUi::Label* size = nullptr;
    RedSalamander::DxUi::Button* view = nullptr;
    RedSalamander::DxUi::Button* remove = nullptr;
};

[[nodiscard]] bool EnsureItemPropertiesWindowClassRegistered() noexcept;

class ItemPropertiesWindow final
{
public:
    ItemPropertiesWindow(Common::Settings::Settings* settings,
                         const AppTheme& theme,
                         std::filesystem::path itemPath,
                         wil::com_ptr<IFileSystemIO> itemIo,
                         wil::com_ptr<IFileSystemItemStreams> streamOps,
                         ItemPropertiesOpenStreamCallback openStream) noexcept
        : _settings(settings),
          _theme(theme),
          _itemPath(std::move(itemPath)),
          _itemIo(std::move(itemIo)),
          _streamOps(std::move(streamOps)),
          _openStream(std::move(openStream)),
          _windowToken(g_nextItemPropertiesWindowToken.fetch_add(1u, std::memory_order_relaxed)),
          _contentText(BuildItemPropertiesLoadingText())
    {
    }

    ItemPropertiesWindow(const ItemPropertiesWindow&)            = delete;
    ItemPropertiesWindow& operator=(const ItemPropertiesWindow&) = delete;
    ItemPropertiesWindow(ItemPropertiesWindow&&)                 = delete;
    ItemPropertiesWindow& operator=(ItemPropertiesWindow&&)      = delete;

    [[nodiscard]] HRESULT CreateAndShow(HWND owner) noexcept
    {
        if (! EnsureItemPropertiesWindowClassRegistered())
        {
            return HRESULT_FROM_WIN32(::GetLastError());
        }

        _ownerWindow        = (owner && ::IsWindow(owner) != FALSE) ? owner : nullptr;
        _restoreFocusWindow = nullptr;
        if (_ownerWindow)
        {
            const HWND focused = ::GetFocus();
            if (focused && ::IsWindow(focused) != FALSE && (focused == _ownerWindow || ::IsChild(_ownerWindow, focused) != FALSE))
            {
                _restoreFocusWindow = focused;
            }
        }

        const std::wstring caption = LoadStringResource(nullptr, IDS_CAPTION_PROPERTIES);
        const UINT dpi             = owner ? ::GetDpiForWindow(owner) : USER_DEFAULT_SCREEN_DPI;
        const int w                = UiMetrics::ScaleDip(dpi, 720);
        const int h                = UiMetrics::ScaleDip(dpi, 520);

        RECT ownerRc{};
        if (owner && ::GetWindowRect(owner, &ownerRc) == 0)
        {
            owner = nullptr;
        }

        int x = CW_USEDEFAULT;
        int y = CW_USEDEFAULT;
        if (owner)
        {
            const int ownerW = std::max(0l, ownerRc.right - ownerRc.left);
            const int ownerH = std::max(0l, ownerRc.bottom - ownerRc.top);
            x                = ownerRc.left + std::max(0, (ownerW - w) / 2);
            y                = ownerRc.top + std::max(0, (ownerH - h) / 2);
        }

        const HWND hwnd = ::CreateWindowExW(0,
                                            kItemPropertiesWindowClass,
                                            caption.c_str(),
                                            WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
                                            x,
                                            y,
                                            w,
                                            h,
                                            nullptr,
                                            nullptr,
                                            ::GetModuleHandleW(nullptr),
                                            this);
        if (! hwnd)
        {
            return HRESULT_FROM_WIN32(::GetLastError());
        }

        const int showCmd = _settings ? WindowPlacementPersistence::Restore(*_settings, kItemPropertiesWindowId, hwnd) : SW_SHOWNORMAL;
        ::ShowWindow(hwnd, showCmd);
        ::SetForegroundWindow(hwnd);
        return S_OK;
    }

    [[nodiscard]] HWND GetHwnd() const noexcept
    {
        return _hWnd.get();
    }

    static LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) noexcept
    {
        ItemPropertiesWindow* self = reinterpret_cast<ItemPropertiesWindow*>(::GetWindowLongPtrW(hwnd, GWLP_USERDATA));
        if (msg == WM_NCCREATE)
        {
            const auto* create = reinterpret_cast<const CREATESTRUCTW*>(lParam);
            self               = create ? static_cast<ItemPropertiesWindow*>(create->lpCreateParams) : nullptr;
            ::SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
            if (self)
            {
                self->_hWnd.reset(hwnd);
                InitPostedPayloadWindow(hwnd);
            }
        }

        if (! self)
        {
            return ::DefWindowProcW(hwnd, msg, wParam, lParam);
        }

        ++self->_dispatchDepth;
        const auto finishDispatch = wil::scope_exit([self]() noexcept
        {
            if (self->_dispatchDepth > 0u)
            {
                --self->_dispatchDepth;
            }
            if (self->_dispatchDepth == 0u && self->_deletePending)
            {
                delete self;
            }
        });

        return self->WindowProc(hwnd, msg, wParam, lParam);
    }

#ifdef ENABLE_TESTS
    [[nodiscard]] bool DebugGetSnapshot(ItemPropertiesWindowDebugSnapshot& out) const noexcept
    {
        const HWND hwnd             = _hWnd.get();
        out.usesDxUiHost            = hwnd && ::IsWindow(hwnd) != FALSE;
        out.visibleChildWindowCount = hwnd ? CountVisibleItemPropertiesChildWindows(hwnd) : 0u;
        out.sectionCount            = _doc.sections.size();
        out.fieldCount              = _fieldCount;
        out.streamCount             = _doc.streams.size();
        out.removableStreamCount    = _removableStreamCount;
        out.viewableStreamCount     = _viewableStreamCount;
        out.loading                 = _loading;
        out.loadFailed              = _loadFailed;
        if (_contentScroll)
        {
            const D2D1_RECT_F bounds    = _contentScroll->GetBounds();
            out.bodyFirstVisibleLine    = static_cast<size_t>((std::max)(0.0f, _contentScroll->GetScrollOffset()));
            out.bodyVisibleLineCount    = static_cast<size_t>((std::max)(0.0f, bounds.bottom - bounds.top));
            out.bodyTotalLineCount      = static_cast<size_t>((std::max)(0.0f, _contentScroll->GetContentHeight()));
            out.bodyCanScrollVertically = _contentScroll->NeedsScrollbar();
            out.layoutOverflowRightDip  = ComputeLayoutOverflowRightDip();
        }
        out.renderCount        = _dxHost.DebugGetRenderCount();
        out.resizeCount        = _dxHost.DebugGetResizeCount();
        out.resizeFailureCount = _dxHost.DebugGetResizeFailureCount();
        out.contentText        = _contentText;
        return out.usesDxUiHost;
    }

    [[nodiscard]] bool DebugScrollByWheelDetents(int detents) noexcept
    {
        const HWND hwnd = _hWnd.get();
        if (! hwnd || ::IsWindow(hwnd) == FALSE || ! _contentScroll)
        {
            return false;
        }

        if (::IsIconic(hwnd))
        {
            ::ShowWindow(hwnd, SW_RESTORE);
        }

        const float wheelDelta = detents > 0 ? static_cast<float>(WHEEL_DELTA) : -static_cast<float>(WHEEL_DELTA);
        const int stepCount    = detents > 0 ? detents : -detents;
        for (int remaining = stepCount; remaining > 0; --remaining)
        {
            if (! _contentScroll->OnMouseWheel(_dxHost, D2D1::Point2F(0.0f, 0.0f), wheelDelta, 0u))
            {
                return false;
            }
        }
        return true;
    }

    [[nodiscard]] HRESULT DebugRemoveStream(std::wstring_view streamName) noexcept
    {
        return RemoveStreamByName(streamName, false);
    }

    [[nodiscard]] HRESULT DebugOpenStream(std::wstring_view streamName) noexcept
    {
        return OpenStreamByName(streamName, false);
    }
#endif

private:
    [[nodiscard]] LRESULT WindowProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) noexcept
    {
        if (msg == WM_KEYDOWN && wParam == VK_ESCAPE)
        {
            RequestClose(hwnd);
            return 0;
        }

        if (msg == WM_KEYDOWN && (wParam == VK_CONTROL || wParam == VK_LCONTROL || wParam == VK_RCONTROL))
        {
            _ctrlKeyDown = true;
        }
        else if (msg == WM_KEYUP && (wParam == VK_CONTROL || wParam == VK_LCONTROL || wParam == VK_RCONTROL))
        {
            _ctrlKeyDown = false;
        }
        else if (msg == WM_KILLFOCUS)
        {
            _ctrlKeyDown = false;
        }

        const bool ctrlDown = _ctrlKeyDown || (::GetKeyState(VK_CONTROL) & 0x8000) != 0;
        if (msg == WM_KEYDOWN && ctrlDown)
        {
            if (wParam == L'A')
            {
                return 0;
            }
            if (wParam == L'C')
            {
                static_cast<void>(_dxHost.CopyTextToClipboard(_contentText));
                return 0;
            }
        }

        bool handled           = false;
        const LRESULT dxResult = _dxHost.HandleMessage(hwnd, msg, wParam, lParam, handled);
        if (msg == WM_NCDESTROY)
        {
            OnNcDestroy(hwnd);
            if (handled)
            {
                return dxResult;
            }
        }
        if (handled)
        {
            if (msg == WM_SIZE || msg == WM_DPICHANGED || msg == WM_DPICHANGED_AFTERPARENT)
            {
                LayoutControls();
            }
            return dxResult;
        }

        switch (msg)
        {
            case WM_CREATE: return OnCreate(hwnd);
            case WndMsg::kItemPropertiesLoadComplete: return OnLoadComplete(lParam);
            case WndMsg::kItemPropertiesRemoveStream: return OnRemoveStreamMessage();
            case WM_TIMER: return OnTimer(wParam);
            case WM_SIZE: return 0;
            case WM_WINDOWPOSCHANGED:
            {
                const auto* windowPos = reinterpret_cast<const WINDOWPOS*>(lParam);
                if (windowPos && (windowPos->flags & SWP_NOSIZE) == 0)
                {
                    LayoutControls();
                }
                return ::DefWindowProcW(hwnd, msg, wParam, lParam);
            }
            case WM_DPICHANGED: return OnDpiChanged(hwnd, wParam, lParam);
            case WM_ACTIVATE: ApplyTitleBarTheme(hwnd, _theme, wParam != FALSE); return 0;
            case WM_GETMINMAXINFO:
            {
                if (auto* info = reinterpret_cast<MINMAXINFO*>(lParam))
                {
                    Common::WindowSizing::ApplyMinimumClientTrackSizeForDips(hwnd, *info, 520, 320);
                    static_cast<void>(WindowMaximizeBehavior::ApplyVerticalMaximize(hwnd, *info));
                }
                return 0;
            }
            case WM_ERASEBKGND:
            {
                RECT rc{};
                if (::GetClientRect(hwnd, &rc) != 0)
                {
                    ::FillRect(reinterpret_cast<HDC>(wParam), &rc, _backgroundBrush.get());
                    return TRUE;
                }
                return FALSE;
            }
            case WM_CLOSE: RequestClose(hwnd); return 0;
        }

        return ::DefWindowProcW(hwnd, msg, wParam, lParam);
    }

    void RequestClose(HWND hwnd) noexcept
    {
        if (! hwnd)
        {
            return;
        }

        const bool releaseOwnedHwnd = _hWnd.get() == hwnd;
        if (releaseOwnedHwnd)
        {
            static_cast<void>(_hWnd.release());
        }

        if (::DestroyWindow(hwnd) != FALSE)
        {
            return;
        }

        const DWORD lastError = ::GetLastError();
        if (releaseOwnedHwnd && ::IsWindow(hwnd) != FALSE)
        {
            _hWnd.reset(hwnd);
        }
        Debug::Error(L"ItemProperties: DestroyWindow failed after HWND ownership release error={}", lastError);
    }

    [[nodiscard]] LRESULT OnCreate(HWND hwnd) noexcept
    {
        _dpi = ::GetDpiForWindow(hwnd);
        _backgroundBrush.reset(::CreateSolidBrush(_theme.windowBackground));
        if (! _dxHost.Attach(hwnd))
        {
            Debug::Error(L"ItemProperties: failed to attach DxUi host.");
            return -1;
        }
        _dxHost.SetTheme(MakeAppThemeDxPalette(_theme, _theme.windowBackground));
        SyncDxHostClientSize(hwnd);
        BuildUi();
        LayoutControls();
        ApplyWindowChromeTheme(hwnd, _theme, WindowBackdropTarget::Tool, ::GetActiveWindow() == hwnd);
        StartLoadingAnimation();
        StartLoadPropertiesAsync();
        return 0;
    }

    void SyncDxHostClientSize(HWND hwnd) noexcept
    {
        RECT rc{};
        if (! hwnd || ::GetClientRect(hwnd, &rc) == 0)
        {
            return;
        }

        const int widthPx  = (std::max)(0, static_cast<int>(rc.right - rc.left));
        const int heightPx = (std::max)(0, static_cast<int>(rc.bottom - rc.top));
        const UINT width   = static_cast<UINT>((std::min)(widthPx, static_cast<int>((std::numeric_limits<WORD>::max)())));
        const UINT height  = static_cast<UINT>((std::min)(heightPx, static_cast<int>((std::numeric_limits<WORD>::max)())));
        bool handled       = false;
        static_cast<void>(_dxHost.HandleMessage(hwnd, WM_SIZE, 0, MAKELPARAM(width, height), handled));
    }

    [[nodiscard]] LRESULT OnDpiChanged(HWND hwnd, WPARAM wParam, LPARAM lParam) noexcept
    {
        _dpi = static_cast<UINT>(wParam);
        LayoutControls();
        if (const auto* rc = reinterpret_cast<const RECT*>(lParam))
        {
            ::SetWindowPos(
                hwnd, nullptr, rc->left, rc->top, std::max(0l, rc->right - rc->left), std::max(0l, rc->bottom - rc->top), SWP_NOZORDER | SWP_NOACTIVATE);
        }
        return 0;
    }

    void BuildUi()
    {
        auto root = std::make_unique<ItemPropertiesRootPanel>(&_contentText);
        _root     = root.get();

        _titleLabel = root->AddChild<RedSalamander::DxUi::Label>(LoadStringResource(nullptr, IDS_CAPTION_PROPERTIES));
        _titleLabel->SetFontRole(RedSalamander::DxUi::FontRole::Title);

        _contentScroll = root->AddChild<RedSalamander::DxUi::ScrollPanel>();
        _contentScroll->SetScrollStepDip(48.0f);
        RebuildCards();

        _copyHintLabel = root->AddChild<RedSalamander::DxUi::Label>(LoadStringResource(nullptr, IDS_PROPERTIES_COPY_HINT));
        _copyHintLabel->SetFontRole(RedSalamander::DxUi::FontRole::Small);
        _copyHintLabel->SetTextColor(_dxHost.GetTheme().subduedText);

        _closeButton = root->AddChild<RedSalamander::DxUi::Button>(LoadStringResource(nullptr, IDS_PROPERTIES_BTN_CLOSE));
        _closeButton->SetPrimary(true);
        _closeButton->SetMnemonic(L'o');
        _closeButton->SetOnClick([this]()
        {
            if (const HWND hwnd = _hWnd.get(); hwnd && ::IsWindow(hwnd) != FALSE)
            {
                ::PostMessageW(hwnd, WM_CLOSE, 0, 0);
            }
        });

        _dxHost.SetRoot(std::move(root));
        _dxHost.SetDefaultButton(_closeButton);
        _dxHost.SetCancelButton(_closeButton);
        _dxHost.SetFocusControl(_root);
    }

    void LayoutControls() noexcept
    {
        if (! _root || ! _titleLabel || ! _contentScroll || ! _copyHintLabel || ! _closeButton)
        {
            return;
        }

        const D2D1_RECT_F client  = _dxHost.GetClientBoundsDip();
        constexpr float kMargin   = 16.0f;
        constexpr float kGap      = 10.0f;
        constexpr float kTitleH   = 32.0f;
        constexpr float kButtonH  = 28.0f;
        constexpr float kButtonW  = 82.0f;
        constexpr float kHintH    = 18.0f;

        _root->SetBounds(client);

        const float titleTop  = client.top + kMargin;
        const float titleLeft = client.left + kMargin;
        _titleLabel->SetBounds(D2D1::RectF(titleLeft, titleTop, client.right - kMargin, titleTop + kTitleH));

        const float buttonTop = (std::max)(client.top + kMargin + kTitleH + kGap, client.bottom - kMargin - kButtonH);
        _closeButton->SetBounds(D2D1::RectF(client.right - kMargin - kButtonW, buttonTop, client.right - kMargin, buttonTop + kButtonH));

        const float hintRight = client.right - kMargin - kButtonW - kGap;
        const bool showHint   = hintRight >= client.left + kMargin + 160.0f;
        _copyHintLabel->SetVisible(showHint);
        if (showHint)
        {
            const float hintTop = buttonTop + std::floor((kButtonH - kHintH) * 0.5f);
            _copyHintLabel->SetBounds(D2D1::RectF(client.left + kMargin, hintTop, hintRight, hintTop + kHintH));
        }

        const float scrollTop    = titleTop + kTitleH + kGap;
        const float scrollBottom = (std::max)(scrollTop, buttonTop - kGap);
        _contentScroll->SetBounds(D2D1::RectF(client.left + kMargin, scrollTop, client.right - kMargin, scrollBottom));
        LayoutCards();
        _dxHost.Invalidate();
    }

    void UpdateDerivedDocumentState()
    {
        _contentText          = BuildItemPropertiesText(_doc);
        _fieldCount           = CountItemPropertiesFields(_doc);
        _removableStreamCount = CountRemovableItemPropertiesStreams(_doc, _streamOps != nullptr);
        _viewableStreamCount  = CountViewableItemPropertiesStreams(_doc, static_cast<bool>(_openStream));
    }

    void RebuildCards()
    {
        if (! _contentScroll)
        {
            return;
        }

        _contentScroll->ClearChildren();
        _sectionControls.clear();
        _streamRows.clear();
        _loadingCard         = nullptr;
        _loadingSpinnerLabel = nullptr;
        _loadingMessageLabel = nullptr;
        _streamCard          = nullptr;
        _streamTitle         = nullptr;

        if (_loading || _loadFailed)
        {
            _loadingCard = _contentScroll->AddChild<RedSalamander::DxUi::CardPanel>();
            _loadingCard->SetCornerRadius(6.0f);
            _loadingSpinnerLabel = _loadingCard->AddChild<RedSalamander::DxUi::Label>(_loading ? CurrentLoadingSpinnerText() : L"!");
            _loadingSpinnerLabel->SetFontRole(RedSalamander::DxUi::FontRole::BodyLarge);
            _loadingSpinnerLabel->SetAlignment(DWRITE_TEXT_ALIGNMENT_CENTER);
            _loadingMessageLabel = _loadingCard->AddChild<RedSalamander::DxUi::Label>(_loadingMessageText);
            _loadingMessageLabel->SetMultiline(true);
            if (_loadFailed)
            {
                _loadingMessageLabel->SetTextColor(_dxHost.GetTheme().errorText);
            }
            return;
        }

        for (const ItemPropertiesSection& section : _doc.sections)
        {
            ItemPropertiesSectionControls controls{};
            controls.compact = IsCompactItemPropertiesSection(section.title);
            controls.title   = _contentScroll->AddChild<RedSalamander::DxUi::Label>(section.title);
            controls.title->SetFontRole(RedSalamander::DxUi::FontRole::BodyLarge);

            controls.card = _contentScroll->AddChild<RedSalamander::DxUi::CardPanel>();
            controls.card->SetCornerRadius(6.0f);

            controls.fields.reserve(section.fields.size());
            for (const ItemPropertiesField& field : section.fields)
            {
                ItemPropertiesFieldRowControls row{};
                row.key = controls.card->AddChild<RedSalamander::DxUi::Label>(field.key);
                row.key->SetFontRole(RedSalamander::DxUi::FontRole::BodyStrong);
                row.value = controls.card->AddChild<RedSalamander::DxUi::Label>(MakeItemPropertiesWrapFriendlyText(field.value));
                row.value->SetAccessibleName(field.value);
                row.value->SetMultiline(true);
                controls.fields.emplace_back(row);
            }

            _sectionControls.emplace_back(std::move(controls));
        }

        if (_doc.streams.empty())
        {
            return;
        }

        _streamTitle = _contentScroll->AddChild<RedSalamander::DxUi::Label>(LoadStringResource(nullptr, IDS_PROPERTIES_STREAMS_TITLE));
        _streamTitle->SetFontRole(RedSalamander::DxUi::FontRole::BodyLarge);

        _streamCard = _contentScroll->AddChild<RedSalamander::DxUi::CardPanel>();
        _streamCard->SetCornerRadius(6.0f);

        _streamRows.reserve(_doc.streams.size());
        const std::wstring viewText   = LoadStringResource(nullptr, IDS_PROPERTIES_STREAM_VIEW);
        const std::wstring removeText = LoadStringResource(nullptr, IDS_PROPERTIES_STREAM_REMOVE);
        for (size_t index = 0u; index < _doc.streams.size(); ++index)
        {
            const ItemPropertiesStream& stream = _doc.streams[index];
            ItemPropertiesStreamRowControls row{};
            row.name = _streamCard->AddChild<RedSalamander::DxUi::Label>(MakeItemPropertiesWrapFriendlyText(stream.name));
            row.name->SetAccessibleName(stream.name);
            row.name->SetFontRole(RedSalamander::DxUi::FontRole::BodyStrong);
            row.name->SetMultiline(true);
            row.size   = _streamCard->AddChild<RedSalamander::DxUi::Label>(stream.displaySize);
            row.view   = _streamCard->AddChild<RedSalamander::DxUi::Button>(viewText);
            row.view->SetEnabled(static_cast<bool>(_openStream));
            row.remove = _streamCard->AddChild<RedSalamander::DxUi::Button>(removeText);
            row.remove->SetEnabled(_streamOps != nullptr && stream.canRemove);

            std::wstring viewStreamName = stream.name;
            row.view->SetOnClick([this, streamName = std::move(viewStreamName)]() noexcept
            {
                static_cast<void>(OpenStreamByName(streamName, true));
            });

            std::wstring removeStreamName = stream.name;
            row.remove->SetOnClick([this, streamName = std::move(removeStreamName)]() noexcept
            {
                QueueRemoveStreamByName(streamName);
            });

            _streamRows.emplace_back(row);
        }
    }

    void LayoutCards() noexcept
    {
        if (! _contentScroll)
        {
            return;
        }

        const D2D1_RECT_F bounds = _contentScroll->GetBounds();
        const float viewportH    = (std::max)(0.0f, bounds.bottom - bounds.top);

        auto applyLayout = [&](float width) noexcept -> float
        {
            constexpr float kPairGap       = 12.0f;
            constexpr float kPadX          = 12.0f;
            constexpr float kPadTop        = 10.0f;
            constexpr float kPadBottom     = 12.0f;
            constexpr float kSectionTitleH = 26.0f;
            constexpr float kTitleGap      = 6.0f;
            constexpr float kRowH          = 28.0f;
            constexpr float kButtonW       = 82.0f;
            constexpr float kButtonH       = 24.0f;
            constexpr float kStreamGap     = 10.0f;
            constexpr float kTwoColumnMinW = 460.0f;

            const float x = bounds.left;
            float y       = bounds.top;

            const float cardW = (std::max)(0.0f, width);

            auto addInterGroupGap = [&]() noexcept
            {
                if (y > bounds.top)
                {
                    y += 12.0f;
                }
            };

            if (_loadingCard)
            {
                const float cardH = 72.0f;
                const D2D1_RECT_F cardRect = D2D1::RectF(x, y, x + cardW, y + cardH);
                _loadingCard->SetBounds(cardRect);
                const float spinnerW = 38.0f;
                if (_loadingSpinnerLabel)
                {
                    _loadingSpinnerLabel->SetBounds(D2D1::RectF(cardRect.left + kPadX, cardRect.top + kPadTop, cardRect.left + kPadX + spinnerW, cardRect.bottom - kPadBottom));
                }
                if (_loadingMessageLabel)
                {
                    _loadingMessageLabel->SetBounds(D2D1::RectF(cardRect.left + kPadX + spinnerW + kStreamGap,
                                                                 cardRect.top + kPadTop,
                                                                 cardRect.right - kPadX,
                                                                 cardRect.bottom - kPadBottom));
                }
                return cardH;
            }

            auto sectionCardHeight = [&](const ItemPropertiesSectionControls& controls, float sectionW) noexcept
            {
                if (controls.fields.empty())
                {
                    return kPadTop + kRowH + kPadBottom;
                }

                constexpr float kFieldColumnGap = 12.0f;
                const float innerW = (std::max)(1.0f, sectionW - (kPadX * 2.0f));
                const bool stacked = innerW < 260.0f;
                const float keyW = stacked ? innerW : (controls.compact ? (std::clamp)(innerW * 0.34f, 72.0f, 104.0f)
                                                                         : (std::clamp)(innerW * 0.30f, 120.0f, 210.0f));
                const float valueW = stacked ? innerW : (std::max)(1.0f, innerW - keyW - kFieldColumnGap);

                float height = kPadTop + kPadBottom;
                for (const ItemPropertiesFieldRowControls& row : controls.fields)
                {
                    const float valueH = row.value ? MeasureItemPropertiesWrappedTextHeightDip(_dxHost,
                                                                                                row.value->GetText(),
                                                                                                RedSalamander::DxUi::FontRole::Body,
                                                                                                valueW)
                                                   : kRowH;
                    height += stacked ? (kRowH + (std::max)(kRowH, valueH) + 4.0f) : (std::max)(kRowH, valueH + 4.0f);
                }
                return height;
            };

            auto layoutSectionFields = [&](ItemPropertiesSectionControls& controls, D2D1_RECT_F cardRect) noexcept
            {
                constexpr float kFieldColumnGap = 12.0f;
                const float innerW = (std::max)(1.0f, cardRect.right - cardRect.left - (kPadX * 2.0f));
                const bool stacked = innerW < 260.0f;
                const float keyW = stacked ? innerW : (controls.compact ? (std::clamp)(innerW * 0.34f, 72.0f, 104.0f)
                                                                         : (std::clamp)(innerW * 0.30f, 120.0f, 210.0f));
                const float valueW = stacked ? innerW : (std::max)(1.0f, innerW - keyW - kFieldColumnGap);
                controls.card->SetBounds(cardRect);

                float rowY = cardRect.top + kPadTop;
                for (ItemPropertiesFieldRowControls& row : controls.fields)
                {
                    const float valueH = row.value ? MeasureItemPropertiesWrappedTextHeightDip(_dxHost,
                                                                                                row.value->GetText(),
                                                                                                RedSalamander::DxUi::FontRole::Body,
                                                                                                valueW)
                                                   : kRowH;
                    const float rowH = stacked ? (kRowH + (std::max)(kRowH, valueH) + 4.0f) : (std::max)(kRowH, valueH + 4.0f);
                    if (stacked)
                    {
                        row.key->SetBounds(D2D1::RectF(cardRect.left + kPadX, rowY, cardRect.right - kPadX, rowY + kRowH));
                        const float valueTop = rowY + kRowH;
                        row.value->SetBounds(D2D1::RectF(cardRect.left + kPadX, valueTop, cardRect.right - kPadX, valueTop + (std::max)(kRowH, valueH)));
                    }
                    else
                    {
                        const float valueLeft = cardRect.left + kPadX + keyW + kFieldColumnGap;
                        row.key->SetBounds(D2D1::RectF(cardRect.left + kPadX, rowY, cardRect.left + kPadX + keyW, rowY + rowH));
                        row.value->SetBounds(D2D1::RectF(valueLeft, rowY, valueLeft + valueW, rowY + rowH));
                    }
                    rowY += rowH;
                }
            };

            size_t sectionIndex = 0u;
            while (sectionIndex < _sectionControls.size())
            {
                ItemPropertiesSectionControls& controls = _sectionControls[sectionIndex];
                const bool pairWithNext = sectionIndex + 1u < _sectionControls.size() && controls.compact &&
                                          _sectionControls[sectionIndex + 1u].compact && cardW >= kTwoColumnMinW;

                addInterGroupGap();
                if (pairWithNext)
                {
                    ItemPropertiesSectionControls& nextControls = _sectionControls[sectionIndex + 1u];
                    const float columnW                         = std::floor((std::max)(0.0f, cardW - kPairGap) * 0.5f);
                    const float rightX                          = x + columnW + kPairGap;
                    const float cardTop                         = y + kSectionTitleH + kTitleGap;
                    const float pairCardH                       = (std::max)(sectionCardHeight(controls, columnW), sectionCardHeight(nextControls, columnW));

                    controls.title->SetBounds(D2D1::RectF(x, y, x + columnW, y + kSectionTitleH));
                    nextControls.title->SetBounds(D2D1::RectF(rightX, y, rightX + columnW, y + kSectionTitleH));
                    layoutSectionFields(controls, D2D1::RectF(x, cardTop, x + columnW, cardTop + pairCardH));
                    layoutSectionFields(nextControls, D2D1::RectF(rightX, cardTop, rightX + columnW, cardTop + pairCardH));

                    y += kSectionTitleH + kTitleGap + pairCardH;
                    sectionIndex += 2u;
                    continue;
                }

                const float cardTop    = y + kSectionTitleH + kTitleGap;
                const float cardH      = sectionCardHeight(controls, cardW);
                const D2D1_RECT_F cardRect = D2D1::RectF(x, cardTop, x + cardW, cardTop + cardH);
                controls.title->SetBounds(D2D1::RectF(x, y, x + cardW, y + kSectionTitleH));
                layoutSectionFields(controls, cardRect);

                y += kSectionTitleH + kTitleGap + cardH;
                ++sectionIndex;
            }

            if (_streamCard)
            {
                addInterGroupGap();
                const float streamInnerW    = (std::max)(1.0f, cardW - (kPadX * 2.0f));
                const bool stackStreams     = streamInnerW < 320.0f;
                constexpr float kActionGap  = 8.0f;
                const float streamButtonW   = (std::min)(kButtonW, (std::max)(1.0f, streamInnerW - kActionGap) * 0.5f);
                const float streamActionsW  = (streamButtonW * 2.0f) + kActionGap;
                const float streamSizeW     = stackStreams ? (std::max)(1.0f, streamInnerW - streamActionsW - kStreamGap)
                                                           : (std::min)(150.0f, (std::max)(90.0f, cardW * 0.22f));
                const float streamNameW     = stackStreams ? streamInnerW
                                                           : (std::max)(1.0f, streamInnerW - streamSizeW - streamActionsW - (kStreamGap * 2.0f));
                float streamRowsHeight      = 0.0f;
                for (const ItemPropertiesStreamRowControls& row : _streamRows)
                {
                    const float nameH = row.name ? MeasureItemPropertiesWrappedTextHeightDip(_dxHost,
                                                                                              row.name->GetText(),
                                                                                              RedSalamander::DxUi::FontRole::BodyStrong,
                                                                                              streamNameW)
                                                 : kRowH;
                    streamRowsHeight += stackStreams ? ((std::max)(kRowH, nameH + 4.0f) + kRowH + 4.0f) : (std::max)(kRowH, nameH + 4.0f);
                }

                const float cardTop         = y + kSectionTitleH + kTitleGap;
                const float cardH           = kPadTop + streamRowsHeight + kPadBottom;
                const D2D1_RECT_F cardRect  = D2D1::RectF(x, cardTop, x + cardW, cardTop + cardH);
                const D2D1_RECT_F titleRect = D2D1::RectF(x, y, x + cardW, y + kSectionTitleH);
                if (_streamTitle)
                {
                    _streamTitle->SetBounds(titleRect);
                }
                _streamCard->SetBounds(cardRect);

                float rowY = cardRect.top + kPadTop;
                for (ItemPropertiesStreamRowControls& row : _streamRows)
                {
                    const float nameH = row.name ? MeasureItemPropertiesWrappedTextHeightDip(_dxHost,
                                                                                              row.name->GetText(),
                                                                                              RedSalamander::DxUi::FontRole::BodyStrong,
                                                                                              streamNameW)
                                                 : kRowH;
                    if (stackStreams)
                    {
                        const float nameRowH = (std::max)(kRowH, nameH + 4.0f);
                        row.name->SetBounds(D2D1::RectF(cardRect.left + kPadX, rowY, cardRect.right - kPadX, rowY + nameRowH));
                        const float commandTop = rowY + nameRowH;
                        const float actionsLeft = cardRect.right - kPadX - streamActionsW;
                        row.size->SetBounds(
                            D2D1::RectF(cardRect.left + kPadX, commandTop, (std::max)(cardRect.left + kPadX, actionsLeft - kStreamGap), commandTop + kRowH));
                        const float buttonTop = commandTop + std::floor((kRowH - kButtonH) * 0.5f);
                        row.view->SetBounds(D2D1::RectF(actionsLeft, buttonTop, actionsLeft + streamButtonW, buttonTop + kButtonH));
                        row.remove->SetBounds(D2D1::RectF(actionsLeft + streamButtonW + kActionGap,
                                                           buttonTop,
                                                           actionsLeft + streamButtonW + kActionGap + streamButtonW,
                                                           buttonTop + kButtonH));
                        rowY += nameRowH + kRowH + 4.0f;
                    }
                    else
                    {
                        const float rowH       = (std::max)(kRowH, nameH + 4.0f);
                        const float actionsLeft = cardRect.right - kPadX - streamActionsW;
                        const float sizeRight  = actionsLeft - kStreamGap;
                        const float sizeLeft   = sizeRight - streamSizeW;
                        row.name->SetBounds(D2D1::RectF(cardRect.left + kPadX, rowY, cardRect.left + kPadX + streamNameW, rowY + rowH));
                        row.size->SetBounds(D2D1::RectF(sizeLeft, rowY, sizeRight, rowY + rowH));
                        const float buttonTop = rowY + std::floor((rowH - kButtonH) * 0.5f);
                        row.view->SetBounds(D2D1::RectF(actionsLeft, buttonTop, actionsLeft + streamButtonW, buttonTop + kButtonH));
                        row.remove->SetBounds(D2D1::RectF(actionsLeft + streamButtonW + kActionGap,
                                                           buttonTop,
                                                           actionsLeft + streamButtonW + kActionGap + streamButtonW,
                                                           buttonTop + kButtonH));
                        rowY += rowH;
                    }
                }

                y += kSectionTitleH + kTitleGap + cardH;
            }

            return (std::max)(0.0f, y - bounds.top);
        };

        float availableW = (std::max)(0.0f, bounds.right - bounds.left);
        float contentH   = applyLayout(availableW);
        if (contentH > viewportH && availableW > _contentScroll->GetScrollbarThickness())
        {
            availableW -= _contentScroll->GetScrollbarThickness();
            contentH = applyLayout(availableW);
        }

        _contentScroll->SetContentHeight(contentH);
    }

    [[nodiscard]] std::wstring CurrentLoadingSpinnerText() const
    {
        constexpr std::array<std::wstring_view, 4u> kFrames{{L"|", L"/", L"-", L"\\"}};
        return std::wstring(kFrames[_loadingSpinnerStep % kFrames.size()]);
    }

    void StartLoadingAnimation() noexcept
    {
        const HWND hwnd = _hWnd.get();
        if (_loading && hwnd && ::IsWindow(hwnd) != FALSE)
        {
            static_cast<void>(::SetTimer(hwnd, kItemPropertiesLoadingTimerId, 120u, nullptr));
        }
    }

    void StopLoadingAnimation() noexcept
    {
        const HWND hwnd = _hWnd.get();
        if (hwnd && ::IsWindow(hwnd) != FALSE)
        {
            static_cast<void>(::KillTimer(hwnd, kItemPropertiesLoadingTimerId));
        }
    }

    [[nodiscard]] LRESULT OnTimer(WPARAM timerId) noexcept
    {
        if (timerId != kItemPropertiesLoadingTimerId || ! _loading)
        {
            return 0;
        }

        ++_loadingSpinnerStep;
        if (_loadingSpinnerLabel)
        {
            _loadingSpinnerLabel->SetText(CurrentLoadingSpinnerText());
        }
        _dxHost.Invalidate();
        return 0;
    }

    void StartLoadPropertiesAsync() noexcept
    {
        if (! _itemIo)
        {
            ApplyPropertiesLoadFailure(E_POINTER);
            return;
        }

        _loading            = true;
        _loadFailed         = false;
        _loadGeneration     += 1u;
        _loadingMessageText = LoadStringResource(nullptr, IDS_PROPERTIES_LOADING);
        _contentText        = BuildItemPropertiesLoadingText();

        auto work = std::unique_ptr<ItemPropertiesLoadWork>(new (std::nothrow) ItemPropertiesLoadWork{});
        if (! work)
        {
            ApplyPropertiesLoadFailure(E_OUTOFMEMORY);
            return;
        }

        work->hwnd        = _hWnd.get();
        work->window      = this;
        work->windowToken = _windowToken;
        work->generation  = _loadGeneration;
        work->itemPath    = _itemPath;
        work->itemIo      = _itemIo;
#ifdef ENABLE_TESTS
        work->delayMs = g_nextItemPropertiesLoadDelayMs.exchange(0u, std::memory_order_relaxed);
#endif

        const BOOL queued = ::TrySubmitThreadpoolCallback(
            [](PTP_CALLBACK_INSTANCE /*instance*/, void* context) noexcept
            {
                std::unique_ptr<ItemPropertiesLoadWork> workItem(static_cast<ItemPropertiesLoadWork*>(context));
                if (! workItem)
                {
                    return;
                }

#ifdef ENABLE_TESTS
                if (workItem->delayMs > 0u)
                {
                    ::Sleep(workItem->delayMs);
                }
#endif

                auto result = std::unique_ptr<ItemPropertiesLoadResult>(new (std::nothrow) ItemPropertiesLoadResult{});
                if (! result)
                {
                    return;
                }

                result->generation = workItem->generation;
                result->hr         = E_FAIL;

                const char* jsonUtf8 = nullptr;
                result->hr           = workItem->itemIo ? workItem->itemIo->GetItemProperties(workItem->itemPath.c_str(), &jsonUtf8) : E_POINTER;
                if (SUCCEEDED(result->hr))
                {
                    if (jsonUtf8 && jsonUtf8[0] != '\0')
                    {
                        result->jsonUtf8 = jsonUtf8;
                    }
                    else
                    {
                        result->hr = HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
                    }
                }

                auto* window = ! workItem->hwnd ? nullptr
                                                : reinterpret_cast<ItemPropertiesWindow*>(::GetWindowLongPtrW(workItem->hwnd, GWLP_USERDATA));
                if (window != workItem->window || window->_windowToken != workItem->windowToken)
                {
                    return;
                }

                static_cast<void>(PostMessagePayload(workItem->hwnd, WndMsg::kItemPropertiesLoadComplete, 0, std::move(result)));
            },
            work.get(),
            nullptr);
        if (queued == 0)
        {
            const DWORD lastError = ::GetLastError();
            ApplyPropertiesLoadFailure(HRESULT_FROM_WIN32(lastError != 0u ? lastError : ERROR_GEN_FAILURE));
            return;
        }

        static_cast<void>(work.release());
    }

    [[nodiscard]] LRESULT OnLoadComplete(LPARAM lParam) noexcept
    {
        auto result = TakeMessagePayload<ItemPropertiesLoadResult>(lParam);
        if (! result || result->generation != _loadGeneration)
        {
            return 0;
        }

        StopLoadingAnimation();

        if (FAILED(result->hr))
        {
            ApplyPropertiesLoadFailure(result->hr);
            return 0;
        }

        const std::optional<ItemPropertiesDocument> doc = TryParseItemPropertiesJson(std::string_view(result->jsonUtf8));
        if (! doc.has_value())
        {
            ApplyPropertiesLoadFailure(HRESULT_FROM_WIN32(ERROR_INVALID_DATA));
            return 0;
        }

        _doc                = doc.value();
        _loading            = false;
        _loadFailed         = false;
        _loadingMessageText.clear();
        UpdateDerivedDocumentState();
        _dxHost.SetFocusControl(_root);
        RebuildCards();
        LayoutControls();
        return 0;
    }

    void ApplyPropertiesLoadFailure(HRESULT hr) noexcept
    {
        StopLoadingAnimation();
        _doc                = {};
        _loading            = false;
        _loadFailed         = true;
        _loadingMessageText = FormatStringResource(nullptr, IDS_FMT_PROPERTIES_LOAD_FAILED, static_cast<unsigned long>(hr));
        _contentText        = LoadStringResource(nullptr, IDS_CAPTION_PROPERTIES);
        _contentText.append(L"\r\n\r\n");
        _contentText.append(_loadingMessageText);
        _fieldCount           = 0u;
        _removableStreamCount = 0u;
        _viewableStreamCount  = 0u;
        RebuildCards();
        LayoutControls();
    }

#ifdef ENABLE_TESTS
    [[nodiscard]] float ComputeLayoutOverflowRightDip() const noexcept
    {
        if (! _contentScroll)
        {
            return 0.0f;
        }

        const float viewportRight = _contentScroll->GetBounds().right;
        float maxRight            = viewportRight;
        const auto includeControl = [&](const RedSalamander::DxUi::Control* control) noexcept
        {
            if (! control)
            {
                return;
            }

            const D2D1_RECT_F bounds = control->GetBounds();
            maxRight                 = (std::max)(maxRight, bounds.right);
        };

        for (const ItemPropertiesSectionControls& controls : _sectionControls)
        {
            includeControl(controls.title);
            includeControl(controls.card);
            for (const ItemPropertiesFieldRowControls& row : controls.fields)
            {
                includeControl(row.key);
                includeControl(row.value);
            }
        }

        includeControl(_loadingCard);
        includeControl(_loadingSpinnerLabel);
        includeControl(_loadingMessageLabel);
        includeControl(_streamTitle);
        includeControl(_streamCard);
        for (const ItemPropertiesStreamRowControls& row : _streamRows)
        {
            includeControl(row.name);
            includeControl(row.size);
            includeControl(row.view);
            includeControl(row.remove);
        }

        return (std::max)(0.0f, maxRight - viewportRight);
    }
#endif

    void QueueRemoveStreamByName(std::wstring streamName) noexcept
    {
        if (streamName.empty() || _pendingRemoveStreamName.has_value())
        {
            return;
        }

        const HWND hwnd = _hWnd.get();
        if (! hwnd || ::IsWindow(hwnd) == FALSE)
        {
            return;
        }

        _pendingRemoveStreamName = std::move(streamName);
        if (::PostMessageW(hwnd, WndMsg::kItemPropertiesRemoveStream, 0, 0) == 0)
        {
            Debug::Warning(L"ItemProperties: failed to queue stream removal.");
            _pendingRemoveStreamName.reset();
        }
    }

    [[nodiscard]] LRESULT OnRemoveStreamMessage() noexcept
    {
        std::optional<std::wstring> streamName = std::move(_pendingRemoveStreamName);
        _pendingRemoveStreamName.reset();
        if (! streamName.has_value() || streamName.value().empty())
        {
            return 0;
        }

        static_cast<void>(RemoveStreamByName(streamName.value(), true));
        return 0;
    }

    [[nodiscard]] HRESULT RefreshDocumentFromIo() noexcept
    {
        if (! _itemIo)
        {
            return E_POINTER;
        }

        const char* jsonUtf8  = nullptr;
        const HRESULT hrProps = _itemIo->GetItemProperties(_itemPath.c_str(), &jsonUtf8);
        if (FAILED(hrProps))
        {
            return hrProps;
        }
        if (! jsonUtf8 || jsonUtf8[0] == '\0')
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        const std::optional<ItemPropertiesDocument> doc = TryParseItemPropertiesJson(std::string_view(jsonUtf8));
        if (! doc.has_value())
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        _doc = doc.value();
        UpdateDerivedDocumentState();
        _dxHost.SetFocusControl(_root);
        RebuildCards();
        LayoutControls();
        return S_OK;
    }

    void ShowRemoveStreamError(std::wstring_view streamName, HRESULT hr) noexcept
    {
        const std::wstring message =
            FormatStringResource(nullptr, IDS_FMT_PROPERTIES_STREAM_REMOVE_FAILED, std::wstring(streamName), static_cast<unsigned long>(hr));
        const std::wstring caption = LoadStringResource(nullptr, IDS_CAPTION_ERROR);
        ::MessageBoxW(_hWnd.get(), message.c_str(), caption.c_str(), MB_OK | MB_ICONERROR);
    }

    void ShowOpenStreamError(std::wstring_view streamName, HRESULT hr) noexcept
    {
        const std::wstring message =
            FormatStringResource(nullptr, IDS_FMT_PROPERTIES_STREAM_VIEW_FAILED, std::wstring(streamName), static_cast<unsigned long>(hr));
        const std::wstring caption = LoadStringResource(nullptr, IDS_CAPTION_ERROR);
        ::MessageBoxW(_hWnd.get(), message.c_str(), caption.c_str(), MB_OK | MB_ICONERROR);
    }

    [[nodiscard]] HRESULT OpenStreamByName(std::wstring_view streamName, bool showErrors) noexcept
    {
        if (streamName.empty())
        {
            return E_INVALIDARG;
        }
        if (! _openStream)
        {
            return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        }

        const HRESULT hr = _openStream(streamName);
        if (FAILED(hr) && showErrors)
        {
            ShowOpenStreamError(streamName, hr);
        }
        return hr;
    }

    [[nodiscard]] HRESULT RemoveStreamByName(std::wstring_view streamName, bool showErrors) noexcept
    {
        if (streamName.empty())
        {
            return E_INVALIDARG;
        }
        if (! _streamOps)
        {
            return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        }

        const std::wstring streamNameText(streamName);
        const HRESULT hrRemove = _streamOps->DeleteItemStream(_itemPath.c_str(), streamNameText.c_str());
        if (FAILED(hrRemove))
        {
            if (showErrors)
            {
                ShowRemoveStreamError(streamName, hrRemove);
            }
            return hrRemove;
        }

        const HRESULT hrRefresh = RefreshDocumentFromIo();
        if (FAILED(hrRefresh))
        {
            std::erase_if(_doc.streams, [&](const ItemPropertiesStream& stream) noexcept { return stream.name == streamNameText; });
            UpdateDerivedDocumentState();
            _dxHost.SetFocusControl(_root);
            RebuildCards();
            LayoutControls();
            Debug::Warning(L"ItemProperties: stream '{}' was removed but properties refresh failed (hr=0x{:08X}).",
                           streamNameText,
                           static_cast<unsigned long>(hrRefresh));
        }

        return hrRemove;
    }

    void OnNcDestroy(HWND hwnd) noexcept
    {
        ::SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
        _loadGeneration += 1u;
        StopLoadingAnimation();
        static_cast<void>(DrainPostedPayloadsForWindow(hwnd));

        if (_settings)
        {
            WindowPlacementPersistence::Save(*_settings, kItemPropertiesWindowId, hwnd);
            const HRESULT saveHr = SettingsHotReload::SaveSettingsAndSchema(kSettingsAppId, *_settings);
            if (FAILED(saveHr))
            {
                const std::filesystem::path settingsPath = Common::Settings::GetSettingsPath(kSettingsAppId);
                Debug::Error(L"SaveSettings failed (hr=0x{:08X}) path={}", static_cast<unsigned long>(saveHr), settingsPath.wstring());
            }
        }

        if (_ownerWindow && ::IsWindow(_ownerWindow) != FALSE)
        {
            static_cast<void>(::SetActiveWindow(_ownerWindow));

            const HWND restoreFocus = (_restoreFocusWindow && ::IsWindow(_restoreFocusWindow) != FALSE &&
                                       (_restoreFocusWindow == _ownerWindow || ::IsChild(_ownerWindow, _restoreFocusWindow) != FALSE))
                                          ? _restoreFocusWindow
                                          : _ownerWindow;
            static_cast<void>(::SetFocus(restoreFocus));
        }

        _hWnd.release();
        _dxHost.Detach();
        _deletePending = true;
        if (_dispatchDepth == 0u)
        {
            delete this;
        }
    }

    wil::unique_hwnd _hWnd;
    HWND _ownerWindow                     = nullptr;
    HWND _restoreFocusWindow              = nullptr;
    Common::Settings::Settings* _settings = nullptr;
    AppTheme _theme{};
    std::filesystem::path _itemPath;
    wil::com_ptr<IFileSystemIO> _itemIo;
    wil::com_ptr<IFileSystemItemStreams> _streamOps;
    ItemPropertiesOpenStreamCallback _openStream;
    uint64_t _windowToken = 0u;
    ItemPropertiesDocument _doc;
    std::wstring _contentText;
    std::wstring _loadingMessageText = LoadStringResource(nullptr, IDS_PROPERTIES_LOADING);
    std::optional<std::wstring> _pendingRemoveStreamName;
    size_t _fieldCount           = 0u;
    size_t _removableStreamCount = 0u;
    size_t _viewableStreamCount  = 0u;
    size_t _dispatchDepth        = 0u;
    uint64_t _loadGeneration     = 0u;
    size_t _loadingSpinnerStep   = 0u;
    bool _deletePending          = false;
    bool _ctrlKeyDown            = false;
    bool _loading                = true;
    bool _loadFailed             = false;
    UINT _dpi                    = USER_DEFAULT_SCREEN_DPI;
    wil::unique_hbrush _backgroundBrush;
    RedSalamander::DxUi::WindowHost _dxHost;
    RedSalamander::DxUi::Panel* _root               = nullptr;
    RedSalamander::DxUi::Label* _titleLabel         = nullptr;
    RedSalamander::DxUi::ScrollPanel* _contentScroll = nullptr;
    RedSalamander::DxUi::Label* _copyHintLabel      = nullptr;
    RedSalamander::DxUi::Button* _closeButton       = nullptr;
    RedSalamander::DxUi::CardPanel* _loadingCard    = nullptr;
    RedSalamander::DxUi::Label* _loadingSpinnerLabel = nullptr;
    RedSalamander::DxUi::Label* _loadingMessageLabel = nullptr;
    std::vector<ItemPropertiesSectionControls> _sectionControls;
    RedSalamander::DxUi::CardPanel* _streamCard = nullptr;
    RedSalamander::DxUi::Label* _streamTitle = nullptr;
    std::vector<ItemPropertiesStreamRowControls> _streamRows;
};

[[nodiscard]] bool EnsureItemPropertiesWindowClassRegistered() noexcept
{
    static const bool registered = []
    {
        WNDCLASSEXW wc{};
        wc.cbSize        = sizeof(wc);
        wc.lpfnWndProc   = ItemPropertiesWindow::WndProc;
        wc.hInstance     = ::GetModuleHandleW(nullptr);
        wc.hCursor       = ::LoadCursorW(nullptr, IDC_ARROW);
        wc.hIcon         = ::LoadIconW(wc.hInstance, MAKEINTRESOURCEW(IDI_REDSALAMANDER));
        wc.hIconSm       = ::LoadIconW(wc.hInstance, MAKEINTRESOURCEW(IDI_SMALL));
        wc.lpszClassName = kItemPropertiesWindowClass;

        return ::RegisterClassExW(&wc) != 0;
    }();

    return registered;
}

HRESULT ShowItemPropertiesWindow(HWND owner,
                                 Common::Settings::Settings* settings,
                                 const AppTheme& theme,
                                 std::filesystem::path itemPath,
                                 wil::com_ptr<IFileSystemIO> itemIo,
                                 wil::com_ptr<IFileSystemItemStreams> streamOps,
                                 ItemPropertiesOpenStreamCallback openStream) noexcept
{
    auto* window =
        new (std::nothrow) ItemPropertiesWindow(settings, theme, std::move(itemPath), std::move(itemIo), std::move(streamOps), std::move(openStream));
    if (! window)
    {
        return E_OUTOFMEMORY;
    }

    const HRESULT hr = window->CreateAndShow(owner);
    if (FAILED(hr))
    {
        delete window;
    }
    return hr;
}
} // namespace

HRESULT FolderWindow::ShowItemPropertiesFromFolderView(Pane pane, std::filesystem::path path) noexcept
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (! state.fileSystem)
    {
        return E_POINTER;
    }

    if (path.empty())
    {
        return E_INVALIDARG;
    }

    wil::com_ptr<IFileSystemIO> io;
    const HRESULT hrQI = state.fileSystem->QueryInterface(IID_PPV_ARGS(io.addressof()));
    if (FAILED(hrQI) || ! io)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    wil::com_ptr<IFileSystemItemStreams> streamOps;
    static_cast<void>(state.fileSystem->QueryInterface(IID_PPV_ARGS(streamOps.addressof())));

    wil::com_ptr<IFileSystem> fileSystem = state.fileSystem;
    std::wstring fileSystemName;
    if (fileSystem)
    {
        wil::com_ptr<IInformations> fileSystemInfo;
        if (SUCCEEDED(fileSystem->QueryInterface(__uuidof(IInformations), fileSystemInfo.put_void())) && fileSystemInfo)
        {
            const PluginMetaData* meta = nullptr;
            if (SUCCEEDED(fileSystemInfo->GetMetaData(&meta)) && meta && meta->name && meta->name[0] != L'\0')
            {
                fileSystemName = meta->name;
            }
        }
    }
    if (fileSystemName.empty())
    {
        if (! state.pluginShortId.empty())
        {
            fileSystemName = state.pluginShortId;
        }
        else if (! state.pluginId.empty())
        {
            fileSystemName = state.pluginId;
        }
    }

    ItemPropertiesOpenStreamCallback openStream = [this,
                                                   fileSystem,
                                                   fileSystemName = std::move(fileSystemName),
                                                   itemPath = path](std::wstring_view streamName) noexcept -> HRESULT
    {
        if (! fileSystem)
        {
            return E_POINTER;
        }
        if (streamName.empty() || streamName.find(L':') != std::wstring_view::npos || streamName.find(L'\\') != std::wstring_view::npos ||
            streamName.find(L'/') != std::wstring_view::npos || streamName.find(L'\0') != std::wstring_view::npos)
        {
            return E_INVALIDARG;
        }

        std::wstring streamPath = itemPath.wstring();
        streamPath.push_back(L':');
        streamPath.append(streamName);

        HWND ownerWindow = _hWnd ? GetAncestor(_hWnd.get(), GA_ROOT) : nullptr;
        if (! ownerWindow)
        {
            ownerWindow = _hWnd.get();
        }

        ViewerOpenContext context{};
        context.ownerWindow           = ownerWindow;
        context.fileSystem            = fileSystem.get();
        context.fileSystemName        = fileSystemName.empty() ? nullptr : fileSystemName.c_str();
        context.focusedPath           = streamPath.c_str();
        context.selectionPaths        = nullptr;
        context.selectionCount        = 0;
        context.otherFiles            = nullptr;
        context.otherFileCount        = 0;
        context.focusedOtherFileIndex = 0;
        context.flags                 = VIEWER_OPEN_FLAG_NONE;

        return OpenViewerWithPlugin(L"builtin/viewer-text", context);
    };

    return ShowItemPropertiesWindow(_hWnd.get(), _settings, _theme, std::move(path), std::move(io), std::move(streamOps), std::move(openStream));
}

#ifdef ENABLE_TESTS
HWND GetItemPropertiesWindowHandle() noexcept
{
    const HWND hwnd = ::FindWindowW(kItemPropertiesWindowClass, nullptr);
    return (hwnd && ::IsWindow(hwnd) != FALSE) ? hwnd : nullptr;
}

bool DebugGetItemPropertiesWindowSnapshot(ItemPropertiesWindowDebugSnapshot& out) noexcept
{
    const HWND hwnd = GetItemPropertiesWindowHandle();
    if (! hwnd)
    {
        return false;
    }

    auto* window = reinterpret_cast<ItemPropertiesWindow*>(::GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    return window ? window->DebugGetSnapshot(out) : false;
}

bool DebugScrollItemPropertiesWindowByWheelDetents(int detents) noexcept
{
    const HWND hwnd = GetItemPropertiesWindowHandle();
    if (! hwnd)
    {
        return false;
    }

    auto* window = reinterpret_cast<ItemPropertiesWindow*>(::GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    return window ? window->DebugScrollByWheelDetents(detents) : false;
}

HRESULT DebugRemoveItemPropertiesStream(std::wstring_view streamName) noexcept
{
    const HWND hwnd = GetItemPropertiesWindowHandle();
    if (! hwnd)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
    }

    auto* window = reinterpret_cast<ItemPropertiesWindow*>(::GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    return window ? window->DebugRemoveStream(streamName) : E_POINTER;
}

HRESULT DebugOpenItemPropertiesStream(std::wstring_view streamName) noexcept
{
    const HWND hwnd = GetItemPropertiesWindowHandle();
    if (! hwnd)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
    }

    auto* window = reinterpret_cast<ItemPropertiesWindow*>(::GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    return window ? window->DebugOpenStream(streamName) : E_POINTER;
}

std::wstring DebugBuildItemPropertiesContentTextFromJson(std::string_view jsonUtf8) noexcept
{
    const std::optional<ItemPropertiesDocument> doc = TryParseItemPropertiesJson(jsonUtf8);
    if (! doc.has_value())
    {
        return {};
    }

    return BuildItemPropertiesText(doc.value());
}

void DebugSetNextItemPropertiesLoadDelayMs(uint32_t delayMs) noexcept
{
    g_nextItemPropertiesLoadDelayMs.store(delayMs, std::memory_order_relaxed);
}
#endif
