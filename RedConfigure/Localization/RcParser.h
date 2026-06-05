#pragma once

#ifndef NOMINMAX
#define NOMINMAX
#endif
#define WIN32_LEAN_AND_MEAN
#include <Windows.h>

#include <cstddef>
#include <string>
#include <string_view>
#include <vector>

namespace RedConfigure::Localization
{
struct RcStringEntry
{
    std::wstring id;
    std::wstring text;
    size_t sourceLine = 0u;
    bool duplicate    = false;
};

enum class RcLocalizableKind
{
    StringTable,
    MenuPopup,
    MenuItem,
    DialogCaption,
    DialogControl,
};

struct RcLocalizableEntry
{
    RcLocalizableKind kind = RcLocalizableKind::StringTable;
    std::wstring ownerId;
    std::wstring id;
    std::wstring text;
    size_t sourceLine = 0u;
    bool duplicate    = false;
};

struct RcParseResult
{
    std::vector<RcStringEntry> strings;
    std::vector<RcLocalizableEntry> localizableEntries;
    std::vector<std::wstring> errors;
};

HRESULT ParseRcStringTables(std::wstring_view text, RcParseResult& outResult);
} // namespace RedConfigure::Localization
