#pragma once

// Internal helpers for NavigationView implementation split across multiple .cpp files.
// Keep this header private to the NavigationView translation units.

#ifdef min
#undef min
#endif
#ifdef max
#undef max
#endif

#include "NavigationLocation.h"
#include "NavigationView.h"
#include "PlugInterfaces/FileSystem.h"
#include "WindowMessages.h"

#include <algorithm>
#include <array>
#include <cmath>
#include <cstddef>
#include <cwctype>
#include <format>
#include <iterator>
#include <limits>
#include <new>
#include <string_view>

#pragma warning(push)
// C5245 : unreferenced function with internal linkage has been removed
#pragma warning(disable : 5245)

namespace
{
constexpr UINT_PTR EDIT_SUBCLASS_ID          = 1;
constexpr float kIntrinsicTextLayoutMaxWidth = 4096.0f;
constexpr float kFocusRingCornerRadiusDip    = 2.0f;

// Path layout constants in DIPs (96 DPI).
constexpr float kPathPaddingDip                 = 8.0f;
constexpr float kPathSpacingDip                 = 4.0f;
constexpr float kPathSeparatorWidthDip          = 32.0f;
constexpr float kPathTextInsetDip               = kPathSpacingDip * 0.5f;
constexpr float kBreadcrumbHoverInsetDip        = 1.0f;
constexpr float kBreadcrumbHoverCornerRadiusDip = 2.0f;
constexpr int kEditCloseButtonWidthDip          = 24;
constexpr float kEditCloseIconHalfDip           = 5.0f;
constexpr float kEditCloseIconStrokeDip         = 1.5f;
constexpr int kEditTextPaddingXDip              = 6;
constexpr int kEditTextPaddingYDip              = 0;
constexpr int kEditUnderlineHeightDip           = 2;

constexpr size_t kEditSuggestMaxItems      = 11u;
constexpr size_t kEditSuggestMaxCandidates = 256u;

constexpr std::wstring_view kEllipsisText = L"...";
constexpr wchar_t kSeparatorText[]        = L"›";
constexpr wchar_t kHistoryText[]          = L"⩔";

[[nodiscard]] D2D1::ColorF BlendColorF(const D2D1::ColorF& base, const D2D1::ColorF& overlay, float overlayWeight) noexcept
{
    const float t = std::clamp(overlayWeight, 0.0f, 1.0f);
    const float s = 1.0f - t;
    return D2D1::ColorF(base.r * s + overlay.r * t, base.g * s + overlay.g * t, base.b * s + overlay.b * t, 1.0f);
}

[[nodiscard]] COLORREF BlendColorRef(COLORREF base, COLORREF overlay, float overlayWeight) noexcept
{
    const float t = std::clamp(overlayWeight, 0.0f, 1.0f);
    const float s = 1.0f - t;

    const int r = static_cast<int>(std::lround(static_cast<float>(GetRValue(base)) * s + static_cast<float>(GetRValue(overlay)) * t));
    const int g = static_cast<int>(std::lround(static_cast<float>(GetGValue(base)) * s + static_cast<float>(GetGValue(overlay)) * t));
    const int b = static_cast<int>(std::lround(static_cast<float>(GetBValue(base)) * s + static_cast<float>(GetBValue(overlay)) * t));

    return RGB(static_cast<BYTE>(std::clamp(r, 0, 255)), static_cast<BYTE>(std::clamp(g, 0, 255)), static_cast<BYTE>(std::clamp(b, 0, 255)));
}

[[nodiscard]] float DipsToPixels(float dips, UINT dpi) noexcept
{
    return dips * static_cast<float>(dpi) / static_cast<float>(USER_DEFAULT_SCREEN_DPI);
}

[[nodiscard]] int DipsToPixelsInt(int dips, UINT dpi) noexcept
{
    return std::max(0, MulDiv(dips, static_cast<int>(dpi), USER_DEFAULT_SCREEN_DPI));
}

struct EditChromeRects
{
    RECT editRect{};
    RECT closeRect{};
    RECT underlineRect{};
};

[[nodiscard]] EditChromeRects ComputeEditChromeRects(const RECT& pathRect, UINT dpi) noexcept
{
    const int closeWidth      = std::max(1, DipsToPixelsInt(kEditCloseButtonWidthDip, dpi));
    const int underlineHeight = std::max(1, DipsToPixelsInt(kEditUnderlineHeightDip, dpi));

    EditChromeRects result{};
    result.editRect        = pathRect;
    result.editRect.right  = std::max(result.editRect.left, result.editRect.right - closeWidth);
    result.editRect.bottom = std::max(result.editRect.top, result.editRect.bottom - underlineHeight);

    result.closeRect        = pathRect;
    result.closeRect.left   = std::max(result.closeRect.left, result.closeRect.right - closeWidth);
    result.closeRect.bottom = result.editRect.bottom;

    result.underlineRect       = pathRect;
    result.underlineRect.left  = result.editRect.left;
    result.underlineRect.right = result.editRect.right;
    result.underlineRect.top   = std::max(result.underlineRect.top, result.underlineRect.bottom - underlineHeight);

    return result;
}

[[nodiscard]] std::wstring TrimWhitespace(std::wstring_view text)
{
    size_t start = 0;
    size_t end   = text.size();

    while (start < end && std::iswspace(text[start]) != 0)
    {
        ++start;
    }

    while (end > start && std::iswspace(text[end - 1]) != 0)
    {
        --end;
    }

    return std::wstring(text.substr(start, end - start));
}

[[nodiscard]] bool ContainsInsensitive(std::wstring_view text, std::wstring_view needle) noexcept
{
    if (needle.empty())
    {
        return true;
    }

    if (needle.size() > text.size())
    {
        return false;
    }

    const size_t sourceSize = std::min(text.size(), static_cast<size_t>(std::numeric_limits<int>::max()));
    const size_t valueSize  = std::min(needle.size(), static_cast<size_t>(std::numeric_limits<int>::max()));

    const int sourceLen = static_cast<int>(sourceSize);
    const int valueLen  = static_cast<int>(valueSize);
    const int foundAt   = FindStringOrdinal(0, text.data(), sourceLen, needle.data(), valueLen, TRUE);
    return foundAt >= 0;
}

[[nodiscard]] bool EqualsNoCase(std::wstring_view a, std::wstring_view b) noexcept
{
    if (a.size() != b.size())
    {
        return false;
    }

    if (a.size() > static_cast<size_t>(std::numeric_limits<int>::max()))
    {
        return false;
    }

    const int len = static_cast<int>(a.size());
    return CompareStringOrdinal(a.data(), len, b.data(), len, TRUE) == CSTR_EQUAL;
}

[[nodiscard]] bool LooksLikeWindowsDrivePath(std::wstring_view text) noexcept
{
    if (text.size() < 2u)
    {
        return false;
    }

    const wchar_t drive = text[0];
    if ((drive < L'A' || drive > L'Z') && (drive < L'a' || drive > L'z'))
    {
        return false;
    }

    if (text[1] != L':')
    {
        return false;
    }

    if (text.size() < 3u)
    {
        return true; // "C:" (drive-relative)
    }

    const wchar_t slash = text[2];
    return slash == L'\\' || slash == L'/';
}

[[nodiscard]] bool LooksLikeUncPath(std::wstring_view text) noexcept
{
    return text.size() >= 2u && text[0] == L'\\' && text[1] == L'\\';
}

[[nodiscard]] bool LooksLikeExtendedPath(std::wstring_view text) noexcept
{
    return text.rfind(L"\\\\?\\", 0) == 0 || text.rfind(L"\\\\.\\", 0) == 0;
}

[[nodiscard]] bool LooksLikeWindowsAbsolutePath(std::wstring_view text) noexcept
{
    if (text.empty())
    {
        return false;
    }

    if (LooksLikeExtendedPath(text))
    {
        return true;
    }

    if (LooksLikeUncPath(text))
    {
        return true;
    }

    if (! LooksLikeWindowsDrivePath(text))
    {
        return false;
    }

    if (text.size() < 3u)
    {
        return false;
    }

    const wchar_t slash = text[2];
    return slash == L'\\' || slash == L'/';
}

[[nodiscard]] bool IsValidPluginShortIdPrefix(std::wstring_view prefix) noexcept
{
    if (prefix.empty())
    {
        return false;
    }

    for (wchar_t ch : prefix)
    {
        if (std::iswalnum(static_cast<wint_t>(ch)) == 0)
        {
            return false;
        }
    }

    return true;
}

[[nodiscard]] bool TryParsePluginPrefix(std::wstring_view text, std::wstring_view& outPrefix, std::wstring_view& outRemainder) noexcept
{
    outPrefix    = {};
    outRemainder = {};

    if (text.empty())
    {
        return false;
    }

    const size_t colon = text.find(L':');
    if (colon == std::wstring_view::npos || colon < 1)
    {
        return false;
    }

    if (colon == 1u && std::iswalpha(static_cast<wint_t>(text[0])) != 0)
    {
        // Avoid treating Windows drive-letter paths ("C:\\...") as plugin prefixes.
        return false;
    }

    const size_t sep = text.find_first_of(L"\\/");
    if (sep != std::wstring_view::npos && sep < colon)
    {
        return false;
    }

    const std::wstring_view prefix = text.substr(0, colon);
    if (! IsValidPluginShortIdPrefix(prefix))
    {
        return false;
    }

    outPrefix    = prefix;
    outRemainder = text.substr(colon + 1u);
    return true;
}

[[nodiscard]] std::optional<std::filesystem::path> TryGetDriveInfoPath(std::wstring_view pluginShortId, const std::filesystem::path& displayPath)
{
    const std::wstring_view displayText(displayPath.native());

    if (pluginShortId.empty())
    {
        return displayPath;
    }

    if (EqualsNoCase(pluginShortId, L"file"))
    {
        if (! LooksLikeWindowsAbsolutePath(displayText))
        {
            return std::nullopt;
        }
        return displayPath;
    }

    if (LooksLikeWindowsAbsolutePath(displayText))
    {
        return std::nullopt;
    }

    std::wstring_view prefix;
    std::wstring_view remainder;
    if (TryParsePluginPrefix(displayText, prefix, remainder))
    {
        if (! EqualsNoCase(prefix, pluginShortId))
        {
            return std::nullopt;
        }

        const std::wstring pluginPathText = NavigationLocation::NormalizePluginPathText(remainder);
        return std::filesystem::path(pluginPathText);
    }

    return displayPath;
}

struct EditSuggestParseResult
{
    std::wstring enumerationShortId;
    std::wstring instanceContext;
    bool instanceContextSpecified = false;
    std::filesystem::path displayFolder;
    std::filesystem::path pluginFolder;
    std::wstring filter;
    wchar_t directorySeparator = L'\\';
};

[[nodiscard]] bool TryParseEditSuggestQuery(std::wstring_view rawInput,
                                            std::wstring_view currentPluginShortId,
                                            const std::optional<std::filesystem::path>& currentPath,
                                            EditSuggestParseResult& result)
{
    result = {};

    std::wstring text = TrimWhitespace(rawInput);
    if (text.size() >= 2u && text.front() == L'"' && text.back() == L'"')
    {
        text = text.substr(1, text.size() - 2u);
        text = TrimWhitespace(text);
    }

    if (text.empty())
    {
        return false;
    }

    const auto isFileShortId = [](std::wstring_view shortId) noexcept { return shortId.empty() || EqualsNoCase(shortId, L"file"); };

    const bool currentIsFile = isFileShortId(currentPluginShortId);

    const auto parseWindowsPath = [&](std::wstring input, std::filesystem::path& folder, std::wstring& filter) -> bool
    {
        std::replace(input.begin(), input.end(), L'/', L'\\');

        folder.clear();
        filter.clear();

        const bool hasSlash = input.find(L'\\') != std::wstring::npos;
        const auto isAlpha  = [](wchar_t ch) noexcept { return (ch >= L'A' && ch <= L'Z') || (ch >= L'a' && ch <= L'z'); };

        if (! hasSlash && input.size() >= 2u && input[1] == L':' && isAlpha(input[0]) != 0)
        {
            folder = std::filesystem::path(std::wstring(input.substr(0, 2)) + L"\\");
            filter = input.substr(2);
            return true;
        }

        const size_t lastSlash = input.find_last_of(L'\\');
        if (lastSlash == std::wstring::npos)
        {
            if (! currentPath || ! LooksLikeWindowsAbsolutePath(currentPath.value().native()))
            {
                return false;
            }
            folder = currentPath.value();
            filter = std::move(input);
            return true;
        }

        folder = std::filesystem::path(input.substr(0, lastSlash + 1u));
        filter = input.substr(lastSlash + 1u);
        return true;
    };

    std::wstring_view typedPrefix;
    std::wstring_view typedRemainder;
    const bool hasTypedPrefix = TryParsePluginPrefix(text, typedPrefix, typedRemainder);

    if (hasTypedPrefix)
    {
        if (EqualsNoCase(typedPrefix, L"file"))
        {
            std::filesystem::path folder;
            std::wstring filter;
            if (! parseWindowsPath(std::wstring(typedRemainder), folder, filter))
            {
                return false;
            }

            result.enumerationShortId = L"file";
            result.directorySeparator = L'\\';
            result.pluginFolder       = folder;
            result.filter             = std::move(filter);

            std::wstring displayText;
            displayText.reserve(typedPrefix.size() + 1u + folder.wstring().size());
            displayText.append(typedPrefix);
            displayText.push_back(L':');
            displayText.append(folder.wstring());
            result.displayFolder = std::filesystem::path(std::move(displayText));
            return true;
        }

        if (EqualsNoCase(typedPrefix, L"7z") && typedRemainder.find(L'|') == std::wstring_view::npos && ! typedRemainder.empty() &&
            typedRemainder.front() != L'/' && typedRemainder.front() != L'\\')
        {
            std::filesystem::path folder;
            std::wstring filter;
            if (! parseWindowsPath(std::wstring(typedRemainder), folder, filter))
            {
                return false;
            }

            result.enumerationShortId = L"file";
            result.directorySeparator = L'\\';
            result.pluginFolder       = folder;
            result.filter             = std::move(filter);

            std::wstring displayText;
            displayText.reserve(typedPrefix.size() + 1u + folder.wstring().size());
            displayText.append(typedPrefix);
            displayText.push_back(L':');
            displayText.append(folder.wstring());
            result.displayFolder = std::filesystem::path(std::move(displayText));
            return true;
        }

        result.enumerationShortId.assign(typedPrefix);
        result.directorySeparator = L'/';

        std::wstring_view mountPart;
        std::wstring_view pluginPathPart = typedRemainder;
        const size_t bar                 = typedRemainder.find(L'|');
        if (bar != std::wstring_view::npos)
        {
            result.instanceContextSpecified = true;
            result.instanceContext.assign(TrimWhitespace(typedRemainder.substr(0, bar)));

            mountPart      = typedRemainder.substr(0, bar + 1u);
            pluginPathPart = typedRemainder.substr(bar + 1u);
        }

        std::filesystem::path folderPart;
        std::wstring filter;
        if (! NavigationLocation::TrySplitPluginPathIntoFolderAndLeaf(pluginPathPart, folderPart, filter))
        {
            return false;
        }

        result.filter = std::move(filter);

        const std::wstring folderPartText = folderPart.wstring();
        std::wstring displayText;
        displayText.reserve(typedPrefix.size() + 1u + mountPart.size() + folderPartText.size());
        displayText.append(typedPrefix);
        displayText.push_back(L':');
        displayText.append(mountPart);
        displayText.append(folderPartText);

        result.displayFolder = std::filesystem::path(std::move(displayText));
        result.pluginFolder  = std::move(folderPart);
        return true;
    }

    if (! currentIsFile)
    {
        if (! text.empty() && (text.front() == L'/' || text.front() == L'\\'))
        {
            std::filesystem::path folderPart;
            std::wstring filter;
            if (! NavigationLocation::TrySplitPluginPathIntoFolderAndLeaf(text, folderPart, filter))
            {
                return false;
            }

            result.filter = std::move(filter);

            const std::wstring folderPartText = folderPart.wstring();

            std::wstring displayText;
            displayText.reserve(currentPluginShortId.size() + 1u + folderPartText.size());
            displayText.append(currentPluginShortId);
            displayText.push_back(L':');
            displayText.append(folderPartText);

            result.enumerationShortId.assign(currentPluginShortId);
            result.directorySeparator = L'/';
            result.displayFolder      = std::filesystem::path(std::move(displayText));
            result.pluginFolder       = std::move(folderPart);
            return true;
        }

        const bool hasSeparator = text.find_first_of(L"\\/") != std::wstring::npos;
        const bool hasColon     = text.find(L':') != std::wstring::npos;
        if (! hasSeparator && ! hasColon && currentPath)
        {
            std::wstring_view currentPrefix;
            std::wstring_view currentRemainder;
            if (TryParsePluginPrefix(currentPath.value().native(), currentPrefix, currentRemainder) && EqualsNoCase(currentPrefix, currentPluginShortId))
            {
                std::wstring folderPart = NavigationLocation::NormalizePluginPathText(currentRemainder,
                                                                                      NavigationLocation::EmptyPathPolicy::Root,
                                                                                      NavigationLocation::LeadingSlashPolicy::Ensure,
                                                                                      NavigationLocation::TrailingSlashPolicy::Ensure);

                std::wstring displayText;
                displayText.reserve(currentPluginShortId.size() + 1u + folderPart.size());
                displayText.append(currentPluginShortId);
                displayText.push_back(L':');
                displayText.append(folderPart);

                result.enumerationShortId.assign(currentPluginShortId);
                result.directorySeparator = L'/';
                result.displayFolder      = std::filesystem::path(std::move(displayText));
                result.pluginFolder       = std::filesystem::path(std::move(folderPart));
                result.filter             = text;
                return true;
            }
        }

        return false;
    }

    std::filesystem::path folder;
    std::wstring filter;
    if (! parseWindowsPath(text, folder, filter))
    {
        return false;
    }

    result.enumerationShortId = L"file";
    result.directorySeparator = L'\\';
    result.displayFolder      = folder;
    result.pluginFolder       = folder;
    result.filter             = std::move(filter);
    return true;
}

void AppendMatchingDirectoryNamesFromFilesInformation(IFilesInformation* info, std::wstring_view filter, std::vector<std::wstring>& names) noexcept
{
    if (! info)
    {
        return;
    }

    FileInfo* entry = nullptr;
    if (FAILED(info->GetBuffer(&entry)) || entry == nullptr)
    {
        return;
    }

    while (entry != nullptr)
    {
        if ((entry->FileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
        {
            const size_t nameChars = static_cast<size_t>(entry->FileNameSize) / sizeof(wchar_t);
            const std::wstring_view name(entry->FileName, nameChars);
            if (name != L"." && name != L".." && ContainsInsensitive(name, filter))
            {
                names.emplace_back(name);
                if (names.size() >= kEditSuggestMaxCandidates)
                {
                    return;
                }
            }
        }

        if (entry->NextEntryOffset == 0)
        {
            return;
        }

        entry = reinterpret_cast<FileInfo*>(reinterpret_cast<std::byte*>(entry) + entry->NextEntryOffset);
    }
}

[[nodiscard]] bool SortAndTrimEditSuggestNames(std::vector<std::wstring>& names)
{
    std::sort(names.begin(), names.end(), [](const std::wstring& a, const std::wstring& b) { return _wcsicmp(a.c_str(), b.c_str()) < 0; });
    const bool hasMore     = names.size() > kEditSuggestMaxItems;
    const size_t keepCount = (hasMore && kEditSuggestMaxItems > 0u) ? (kEditSuggestMaxItems - 1u) : kEditSuggestMaxItems;
    if (names.size() > keepCount)
    {
        names.resize(keepCount);
    }
    return hasMore;
}

void BuildEditSuggestLists(const std::filesystem::path& displayFolder,
                           const std::vector<std::wstring>& names,
                           wchar_t directorySeparator,
                           std::vector<std::wstring>& displayItems,
                           std::vector<std::wstring>& insertItems)
{
    displayItems.clear();
    insertItems.clear();

    displayItems.reserve(names.size());
    insertItems.reserve(names.size());

    std::wstring base = displayFolder.wstring();
    if (! base.empty())
    {
        const wchar_t last = base.back();
        if (last != L'\\' && last != L'/')
        {
            base.push_back(directorySeparator);
        }
    }

    for (const auto& name : names)
    {
        displayItems.push_back(name);
        std::wstring insert = base;
        insert.append(name);
        insertItems.push_back(std::move(insert));
    }
}

[[nodiscard]] D2D1_RECT_F InsetRectF(D2D1_RECT_F rect, float insetX, float insetY) noexcept
{
    rect.left += insetX;
    rect.right -= insetX;
    rect.top += insetY;
    rect.bottom -= insetY;

    if (rect.right < rect.left)
    {
        const float mid = (rect.left + rect.right) * 0.5f;
        rect.left       = mid;
        rect.right      = mid;
    }

    if (rect.bottom < rect.top)
    {
        const float mid = (rect.top + rect.bottom) * 0.5f;
        rect.top        = mid;
        rect.bottom     = mid;
    }

    return rect;
}

[[nodiscard]] D2D1_ROUNDED_RECT RoundedRect(D2D1_RECT_F rect, float radius) noexcept
{
    const float width           = std::max(0.0f, rect.right - rect.left);
    const float height          = std::max(0.0f, rect.bottom - rect.top);
    const float maxCornerRadius = std::min(width, height) * 0.5f;
    const float cornerRadius    = std::min(std::max(0.0f, radius), maxCornerRadius);
    return D2D1::RoundedRect(rect, cornerRadius, cornerRadius);
}

[[nodiscard]] bool IsWin32MenuWindow(HWND hwnd) noexcept
{
    if (! hwnd)
    {
        return false;
    }

    wchar_t className[16]{};
    if (GetClassNameW(hwnd, className, static_cast<int>(std::size(className))) == 0)
    {
        return false;
    }

    return wcscmp(className, L"#32768") == 0;
}

void LayoutSingleLineEditInRect(HWND edit, const RECT& containerRect) noexcept
{
    if (! edit)
    {
        return;
    }

    const LONG containerWidth  = std::max(0l, containerRect.right - containerRect.left);
    const LONG containerHeight = std::max(0l, containerRect.bottom - containerRect.top);

    SetWindowPos(edit, nullptr, containerRect.left, containerRect.top, containerWidth, containerHeight, SWP_NOZORDER | SWP_NOACTIVATE);

    RECT clientRect{};
    GetClientRect(edit, &clientRect);

    RECT formatRect     = clientRect;
    const UINT dpi      = GetDpiForWindow(edit);
    const LONG paddingX = static_cast<LONG>(DipsToPixelsInt(kEditTextPaddingXDip, dpi));
    const LONG paddingY = static_cast<LONG>(DipsToPixelsInt(kEditTextPaddingYDip, dpi));
    formatRect.left     = std::min(formatRect.right, formatRect.left + paddingX);
    formatRect.right    = std::max(formatRect.left, formatRect.right - paddingX);
    formatRect.top      = std::min(formatRect.bottom, formatRect.top + paddingY);
    formatRect.bottom   = std::max(formatRect.top, formatRect.bottom - paddingY);

    HFONT font = reinterpret_cast<HFONT>(SendMessageW(edit, WM_GETFONT, 0, 0));
    auto hdc   = wil::GetDC(edit);
    if (hdc)
    {
        HFONT fontToUse = font ? font : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
        auto oldFont    = wil::SelectObject(hdc.get(), fontToUse);

        TEXTMETRICW tm{};
        if (GetTextMetricsW(hdc.get(), &tm) != 0)
        {
            const LONG lineHeight      = std::max(1l, static_cast<LONG>(tm.tmHeight));
            const LONG availableHeight = std::max(0l, formatRect.bottom - formatRect.top);

            if (availableHeight > lineHeight)
            {
                formatRect.top = formatRect.top + (availableHeight - lineHeight) / 2;
            }
        }
    }

    SendMessageW(edit, EM_SETRECTNP, 0, reinterpret_cast<LPARAM>(&formatRect));
    InvalidateRect(edit, nullptr, FALSE);
}

std::filesystem::path NormalizeDirectoryPath(std::filesystem::path path)
{
    path = path.lexically_normal();
    while (! path.empty() && ! path.has_filename() && path != path.root_path())
    {
        path = path.parent_path();
    }
    return path;
}

std::wstring FilenameOrPath(const std::filesystem::path& path)
{
    const std::filesystem::path filename = path.filename();
    if (! filename.empty())
    {
        return filename.wstring();
    }
    return path.wstring();
}

void CreateTextLayoutAndWidth(IDWriteFactory* factory,
                              IDWriteTextFormat* format,
                              std::wstring_view text,
                              float maxWidth,
                              float height,
                              wil::com_ptr<IDWriteTextLayout>& layout,
                              float& width)
{
    layout.reset();
    width = 0.0f;

    if (! factory || ! format || text.empty())
    {
        return;
    }

    wil::com_ptr<IDWriteTextLayout> tempLayout;
    const HRESULT hr = factory->CreateTextLayout(text.data(), static_cast<UINT32>(text.size()), format, maxWidth, height, tempLayout.addressof());
    if (FAILED(hr) || ! tempLayout)
    {
        return;
    }

    DWRITE_TEXT_METRICS metrics{};
    if (FAILED(tempLayout->GetMetrics(&metrics)))
    {
        return;
    }

    if (metrics.width > 0.0f)
    {
        tempLayout->SetMaxWidth(metrics.width);
    }

    width  = metrics.width;
    layout = std::move(tempLayout);
}

float MeasureTextWidth(IDWriteFactory* factory, IDWriteTextFormat* format, std::wstring_view text, float maxWidth, float height)
{
    wil::com_ptr<IDWriteTextLayout> layout;
    float width = 0.0f;
    CreateTextLayoutAndWidth(factory, format, text, maxWidth, height, layout, width);
    return width;
}

std::wstring
TruncateTextToWidth(IDWriteFactory* factory, IDWriteTextFormat* format, std::wstring_view text, float maxWidth, float height, std::wstring_view ellipsisText)
{
    const float currentWidth = MeasureTextWidth(factory, format, text, kIntrinsicTextLayoutMaxWidth, height);
    if (currentWidth <= maxWidth)
    {
        return std::wstring(text);
    }

    const float dotsWidth = MeasureTextWidth(factory, format, ellipsisText, kIntrinsicTextLayoutMaxWidth, height);
    if (dotsWidth <= 0.0f || maxWidth <= dotsWidth)
    {
        return std::wstring(ellipsisText);
    }

    size_t low  = 0;
    size_t high = text.size();
    while (low < high)
    {
        const size_t mid = (low + high + 1) / 2;
        std::wstring candidate;
        candidate.reserve(mid + ellipsisText.size());
        candidate.append(text.substr(0, mid));
        candidate.append(ellipsisText);

        const float candidateWidth = MeasureTextWidth(factory, format, candidate, kIntrinsicTextLayoutMaxWidth, height);
        if (candidateWidth <= maxWidth)
        {
            low = mid;
        }
        else
        {
            high = mid - 1;
        }
    }

    std::wstring result;
    result.reserve(low + ellipsisText.size());
    result.append(text.substr(0, low));
    result.append(ellipsisText);
    return result;
}

// AlphaBlend replacement (software, premultiplied or straight alpha)
[[nodiscard]] BOOL BlitAlphaBlend(HDC hdcDest,
                                  int xoriginDest,
                                  int yoriginDest,
                                  int wDest,
                                  int hDest,
                                  HDC hdcSrc,
                                  int xoriginSrc,
                                  int yoriginSrc,
                                  int wSrc,
                                  int hSrc,
                                  BLENDFUNCTION ftn) noexcept
{
    if (! hdcDest || ! hdcSrc || wDest <= 0 || hDest <= 0 || wSrc <= 0 || hSrc <= 0)
    {
        return TRUE;
    }
    if (ftn.BlendOp != AC_SRC_OVER)
    {
        return FALSE;
    }

    const bool useSrcAlpha     = (ftn.AlphaFormat & AC_SRC_ALPHA) != 0;
    const uint32_t globalAlpha = ftn.SourceConstantAlpha;

    BITMAPINFO bmi{};
    bmi.bmiHeader.biSize        = sizeof(BITMAPINFOHEADER);
    bmi.bmiHeader.biWidth       = wDest;
    bmi.bmiHeader.biHeight      = -hDest; // top-down
    bmi.bmiHeader.biPlanes      = 1;
    bmi.bmiHeader.biBitCount    = 32;
    bmi.bmiHeader.biCompression = BI_RGB;

    void* destBits = nullptr;
    wil::unique_hbitmap destDib(CreateDIBSection(hdcDest, &bmi, DIB_RGB_COLORS, &destBits, nullptr, 0));
    if (! destDib || ! destBits)
    {
        return FALSE;
    }

    wil::unique_hdc destMem(CreateCompatibleDC(hdcDest));
    if (! destMem)
    {
        return FALSE;
    }
    auto oldDestBmp = wil::SelectObject(destMem.get(), destDib.get());
    if (! BitBlt(destMem.get(), 0, 0, wDest, hDest, hdcDest, xoriginDest, yoriginDest, SRCCOPY))
    {
        return FALSE;
    }

    void* srcBits = nullptr;
    wil::unique_hbitmap srcDib(CreateDIBSection(hdcDest, &bmi, DIB_RGB_COLORS, &srcBits, nullptr, 0));
    if (! srcDib || ! srcBits)
    {
        return FALSE;
    }

    wil::unique_hdc srcMem(CreateCompatibleDC(hdcDest));
    if (! srcMem)
    {
        return FALSE;
    }
    auto oldSrcBmp = wil::SelectObject(srcMem.get(), srcDib.get());

    const int prevStretch = SetStretchBltMode(srcMem.get(), HALFTONE);
    const bool copyOk     = (wSrc == wDest && hSrc == hDest) ? BitBlt(srcMem.get(), 0, 0, wDest, hDest, hdcSrc, xoriginSrc, yoriginSrc, SRCCOPY)
                                                             : StretchBlt(srcMem.get(), 0, 0, wDest, hDest, hdcSrc, xoriginSrc, yoriginSrc, wSrc, hSrc, SRCCOPY);
    SetStretchBltMode(srcMem.get(), prevStretch);
    if (! copyOk)
    {
        return FALSE;
    }

    auto* dst = static_cast<uint32_t*>(destBits);
    auto* src = static_cast<uint32_t*>(srcBits);

    for (int y = 0; y < hDest; ++y)
    {
        const auto rowOffset = static_cast<size_t>(y) * static_cast<size_t>(wDest);
        for (int x = 0; x < wDest; ++x)
        {
            const uint32_t s   = src[rowOffset + static_cast<size_t>(x)];
            const uint8_t srcA = useSrcAlpha ? static_cast<uint8_t>(s >> 24) : 255u;

            const uint32_t alpha = (static_cast<uint32_t>(srcA) * globalAlpha + 127u) / 255u;
            if (alpha == 0)
            {
                continue;
            }

            const uint32_t invA = 255u - alpha;

            const uint8_t srcB = static_cast<uint8_t>(s);
            const uint8_t srcG = static_cast<uint8_t>(s >> 8);
            const uint8_t srcR = static_cast<uint8_t>(s >> 16);

            const uint32_t d   = dst[rowOffset + static_cast<size_t>(x)];
            const uint8_t dstB = static_cast<uint8_t>(d);
            const uint8_t dstG = static_cast<uint8_t>(d >> 8);
            const uint8_t dstR = static_cast<uint8_t>(d >> 16);

            const uint8_t outB = static_cast<uint8_t>((static_cast<uint32_t>(srcB) * alpha + static_cast<uint32_t>(dstB) * invA + 127u) / 255u);
            const uint8_t outG = static_cast<uint8_t>((static_cast<uint32_t>(srcG) * alpha + static_cast<uint32_t>(dstG) * invA + 127u) / 255u);
            const uint8_t outR = static_cast<uint8_t>((static_cast<uint32_t>(srcR) * alpha + static_cast<uint32_t>(dstR) * invA + 127u) / 255u);

            dst[rowOffset + static_cast<size_t>(x)] = (static_cast<uint32_t>(outR) << 16) | (static_cast<uint32_t>(outG) << 8) | outB | 0xFF000000u;
        }
    }

    return BitBlt(hdcDest, xoriginDest, yoriginDest, wDest, hDest, destMem.get(), 0, 0, SRCCOPY) != 0;
}

} // namespace
#pragma warning(pop)
