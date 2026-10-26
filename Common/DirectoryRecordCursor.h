#pragma once

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <cstddef>
#include <cstdint>
#include <span>
#include <string_view>

namespace Common::DirectoryRecords
{
enum class NameKind
{
    Leaf,
    RelativePath,
};

template <typename T> struct RecordView final
{
    const T* information = nullptr;
    std::wstring_view name;
    size_t nextOffset = 0u; // Absolute buffer offset; zero marks the last record.
};

// Windows directory/notification layouts share these three fields, but keep their own
// alignment and header size. Capacity is the bound when the Win32 query returns no byte count.
// Validate the complete record before publishing any borrowed fields to a consumer.
template <typename T>
[[nodiscard]] HRESULT Read(std::span<const std::byte> buffer, size_t offset, RecordView<T>& output, NameKind kind = NameKind::Leaf) noexcept
{
    output                    = {};
    constexpr size_t header   = offsetof(T, FileName);
    constexpr HRESULT invalid = HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    if (offset > buffer.size() || buffer.size() - offset < header || (reinterpret_cast<uintptr_t>(buffer.data() + offset) % alignof(T)) != 0u)
    {
        return invalid;
    }
    const auto* information  = reinterpret_cast<const T*>(buffer.data() + offset);
    const size_t remaining   = buffer.size() - offset;
    const size_t advance     = information->NextEntryOffset;
    const size_t recordBytes = advance != 0u ? advance : remaining;
    const size_t nameBytes   = information->FileNameLength;
    if (recordBytes < header || recordBytes > remaining || nameBytes == 0u || nameBytes % sizeof(wchar_t) != 0u || nameBytes > recordBytes - header)
    {
        return invalid;
    }
    if (advance != 0u && (advance % alignof(T) != 0u || remaining - advance < header + sizeof(wchar_t)))
    {
        return invalid;
    }
    const std::wstring_view name(information->FileName, nameBytes / sizeof(wchar_t));
    if (name.find_first_of(std::wstring_view(L"\0:/", 3u)) != std::wstring_view::npos ||
        (kind == NameKind::Leaf && name.find(L'\\') != std::wstring_view::npos))
    {
        return invalid;
    }
    if (kind == NameKind::RelativePath)
    {
        size_t start = 0u;
        for (;;)
        {
            const size_t end                  = name.find(L'\\', start);
            const std::wstring_view component = name.substr(start, end == std::wstring_view::npos ? end : end - start);
            if (component.empty() || component == L"." || component == L"..")
                return invalid;
            if (end == std::wstring_view::npos)
                break;
            start = end + 1u;
        }
    }
    output.information = information;
    output.name        = name;
    output.nextOffset  = advance != 0u ? offset + advance : 0u;
    return S_OK;
}
} // namespace Common::DirectoryRecords
