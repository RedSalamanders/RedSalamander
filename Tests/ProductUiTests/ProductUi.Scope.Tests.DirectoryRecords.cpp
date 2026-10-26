#include "ProductUi.Scope.Tests.DirectoryRecords.h"

#include "DirectoryRecordCursor.h"

#include <array>
#include <cstring>
#include <limits>

namespace
{
template <typename T> [[nodiscard]] bool CheckLayout() noexcept
{
    using namespace Common::DirectoryRecords;
    alignas(T) std::array<std::byte, 512u> buffer{};
    constexpr size_t header = offsetof(T, FileName);
    const auto put          = [&](size_t offset, std::wstring_view name, ULONG next) noexcept
    {
        auto* entry            = reinterpret_cast<T*>(buffer.data() + offset);
        entry->NextEntryOffset = next;
        entry->FileNameLength  = static_cast<ULONG>(name.size() * sizeof(wchar_t));
        std::memcpy(entry->FileName, name.data(), name.size() * sizeof(wchar_t));
    };
    auto* first = reinterpret_cast<T*>(buffer.data());
    RecordView<T> view{};
    put(0u, L"one", 256u);
    put(256u, L"two", 0u);
    if (Read<T>(buffer, 0u, view) != S_OK || view.name != L"one" || view.nextOffset != 256u || Read<T>(buffer, view.nextOffset, view) != S_OK ||
        view.name != L"two" || view.nextOffset != 0u)
        return false;
    for (const size_t size : {size_t{0u}, header - 1u})
        if (Read<T>(std::span(buffer).first(size), 0u, view) != HRESULT_FROM_WIN32(ERROR_INVALID_DATA) || view.information != nullptr)
            return false;
    if (Read<T>(buffer, buffer.size(), view) != HRESULT_FROM_WIN32(ERROR_INVALID_DATA) ||
        Read<T>(buffer, std::numeric_limits<size_t>::max(), view) != HRESULT_FROM_WIN32(ERROR_INVALID_DATA) ||
        Read<T>(buffer, 1u, view) != HRESULT_FROM_WIN32(ERROR_INVALID_DATA))
        return false;
    for (const ULONG next : {ULONG{1u}, static_cast<ULONG>(header), ULONG{255u}, ULONG{512u}, std::numeric_limits<ULONG>::max()})
    {
        put(0u, L"one", next);
        if (Read<T>(buffer, 0u, view) != HRESULT_FROM_WIN32(ERROR_INVALID_DATA))
            return false;
    }
    put(0u, L"one", 256u);
    for (const ULONG length : {ULONG{0u}, ULONG{1u}, ULONG{257u}, ULONG{256u}, std::numeric_limits<ULONG>::max()})
    {
        first->FileNameLength = length;
        if (Read<T>(buffer, 0u, view) != HRESULT_FROM_WIN32(ERROR_INVALID_DATA))
            return false;
    }
    for (const std::wstring_view name : std::array<std::wstring_view, 4u>{{L"a\\b", L"a/b", L"a:b", std::wstring_view(L"a\0b", 3u)}})
    {
        put(0u, name, 0u);
        if (Read<T>(buffer, 0u, view) != HRESULT_FROM_WIN32(ERROR_INVALID_DATA))
            return false;
    }
    for (const std::wstring_view name : {L".", L"..", L"ordinary", L"\u00E9\U0001F600", L"trailing. "})
    {
        put(0u, name, 0u);
        if (Read<T>(buffer, 0u, view) != S_OK || view.name != name)
            return false;
    }
    put(0u, L"folder\\file", 0u);
    if (Read<T>(buffer, 0u, view, NameKind::RelativePath) != S_OK)
        return false;
    for (const std::wstring_view name : {L"\\root", L"folder\\", L"folder\\..\\file", L"folder\\\\file", L".\\file"})
    {
        put(0u, name, 0u);
        if (Read<T>(buffer, 0u, view, NameKind::RelativePath) != HRESULT_FROM_WIN32(ERROR_INVALID_DATA))
            return false;
    }
    return true;
}
} // namespace

bool TestCheckedDirectoryRecords() noexcept
{
    return CheckLayout<FILE_FULL_DIR_INFO>() && CheckLayout<FILE_ID_BOTH_DIR_INFO>() && CheckLayout<FILE_NOTIFY_INFORMATION>();
}
